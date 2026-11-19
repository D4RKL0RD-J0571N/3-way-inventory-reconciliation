using Acr.Core.Models;

namespace Acr.Infrastructure.Safety;

/// <summary>
/// Metadata summarizing a completed atomic file transaction.
/// </summary>
/// <param name="TargetPath">Canonicalized target file path committed.</param>
/// <param name="BackupPath">Backup (<c>.bak</c>) path created when replacing an existing file, otherwise null.</param>
/// <param name="TempPath">Temporary staging path used prior to atomic replacement.</param>
/// <param name="BackupCreated">True when an existing target file was backed up via <c>File.Replace</c>.</param>
/// <param name="Verified">True when post-commit verification succeeded.</param>
public sealed record AtomicTransactionResult(
    string TargetPath,
    string? BackupPath,
    string TempPath,
    bool BackupCreated,
    bool Verified);

/// <summary>
/// Performs crash-safe, OneDrive/M365 lock-aware atomic file updates using staging <c>.tmp</c> files,
/// pre-commit validation, <c>.bak</c> backup creation via <see cref="File.Replace(string, string, string, bool)"/>,
/// and post-commit verification.
/// </summary>
public static class AtomicFileTransaction
{
    /// <summary>
    /// Returns true when the target path itself is an Office lockfile (<c>~$</c>) or a sibling
    /// OneDrive/Microsoft 365 lockfile (<c>~$&lt;filename&gt;</c>) exists in the same directory.
    /// </summary>
    public static bool HasOfficeLockFile(string targetPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(targetPath);
        string fullPath = Path.GetFullPath(targetPath);
        string fileName = Path.GetFileName(fullPath);
        if (fileName.StartsWith("~$", StringComparison.OrdinalIgnoreCase))
            return true;

        string? directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(directory))
            return false;

        string lockFilePath = Path.Combine(directory, $"~${fileName}");
        return File.Exists(lockFilePath);
    }

    /// <summary>
    /// Verifies exclusive read/write access on an existing file using <see cref="FileShare.None"/>,
    /// throwing <see cref="IOException"/> if another process (such as Excel or OneDrive) holds a handle.
    /// </summary>
    public static void EnsureExclusiveFileAccess(string targetPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(targetPath);
        string fullPath = Path.GetFullPath(targetPath);
        if (!File.Exists(fullPath))
            return;

        try
        {
            using var stream = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new IOException($"Target file is locked or read-only: '{fullPath}'.", ex);
        }
        catch (IOException ex)
        {
            throw new IOException($"Target file is currently in use by another process: '{fullPath}'.", ex);
        }
    }

    /// <summary>
    /// Executes an atomic write transaction against <paramref name="targetPath"/> after enforcing
    /// <paramref name="guard"/>, checking for Office lockfiles and exclusive handles, validating the
    /// temporary file before commit, committing atomically, and running optional post-commit verification.
    /// </summary>
    public static AtomicTransactionResult Execute(
        string targetPath,
        OperationalSafetyGuard guard,
        OperationalActionKind action,
        Action<string> writeTempFile,
        Func<string, bool>? validateTempFile = null,
        Func<string, bool>? verifyCommittedFile = null,
        OperationalTargetKind? explicitKind = null)
    {
        ArgumentNullException.ThrowIfNull(guard);
        ArgumentNullException.ThrowIfNull(writeTempFile);

        string canonicalTarget = guard.EnsureCanWrite(targetPath, action, explicitKind);
        EnsureTargetUnlocked(canonicalTarget);

        (string directory, string tempPath, string backupPath) = PrepareStagingPaths(canonicalTarget);
        Directory.CreateDirectory(directory);

        try
        {
            StageAndValidateTempFile(tempPath, writeTempFile, validateTempFile);
            bool backupCreated = CommitTempFile(tempPath, canonicalTarget, backupPath);
            VerifyCommittedTarget(canonicalTarget, verifyCommittedFile);
            return new AtomicTransactionResult(
                canonicalTarget,
                backupCreated ? backupPath : null,
                tempPath,
                backupCreated,
                Verified: true);
        }
        finally
        {
            TryDeleteFile(tempPath);
        }
    }

    private static void EnsureTargetUnlocked(string canonicalTarget)
    {
        if (HasOfficeLockFile(canonicalTarget))
        {
            throw new IOException(
                $"Detected active Office/OneDrive lockfile (~$) for '{canonicalTarget}'. Close the workbook before saving.");
        }

        EnsureExclusiveFileAccess(canonicalTarget);
    }

    private static (string Directory, string TempPath, string BackupPath) PrepareStagingPaths(string canonicalTarget)
    {
        string? directory = Path.GetDirectoryName(canonicalTarget);
        string fileName = Path.GetFileName(canonicalTarget);
        if (string.IsNullOrEmpty(directory) || string.IsNullOrEmpty(fileName))
            throw new ArgumentException($"Target path must include a directory and filename: '{canonicalTarget}'.");

        string tempPath = Path.Combine(directory, $".{fileName}.{Guid.NewGuid():N}.tmp");
        string backupPath = $"{canonicalTarget}.bak";
        return (directory, tempPath, backupPath);
    }

    private static void StageAndValidateTempFile(
        string tempPath,
        Action<string> writeTempFile,
        Func<string, bool>? validateTempFile)
    {
        writeTempFile(tempPath);
        if (!File.Exists(tempPath) || new FileInfo(tempPath).Length == 0)
            throw new InvalidOperationException($"Staging temp file was not written or is empty: '{tempPath}'.");

        if (validateTempFile is not null && !validateTempFile(tempPath))
            throw new InvalidOperationException($"Pre-commit validation failed for '{tempPath}'; target file untouched.");
    }

    private static bool CommitTempFile(string tempPath, string canonicalTarget, string backupPath)
    {
        if (File.Exists(canonicalTarget))
        {
            File.Replace(tempPath, canonicalTarget, backupPath, ignoreMetadataErrors: true);
            return true;
        }

        File.Move(tempPath, canonicalTarget, overwrite: false);
        return false;
    }

    private static void VerifyCommittedTarget(string canonicalTarget, Func<string, bool>? verifyCommittedFile)
    {
        if (!File.Exists(canonicalTarget) || new FileInfo(canonicalTarget).Length == 0)
            throw new IOException($"Post-commit verification failed: '{canonicalTarget}' is missing or empty.");

        if (verifyCommittedFile is not null && !verifyCommittedFile(canonicalTarget))
            throw new InvalidOperationException($"Post-commit verification callback failed for '{canonicalTarget}'.");
    }

    private static void TryDeleteFile(string path)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
