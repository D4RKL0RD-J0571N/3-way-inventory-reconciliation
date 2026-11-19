using Acr.Core.Models;
using Acr.Core.Services;

namespace Acr.Infrastructure.Safety;

/// <summary>
/// Exception thrown when an operational safety policy prohibits a requested filesystem read or write.
/// </summary>
public sealed class OperationalSafetyException : InvalidOperationException
{
    /// <summary>Policy decision that caused the rejection.</summary>
    public SafetyPolicyDecision Decision { get; }

    /// <summary>Canonicalized target path evaluated by <see cref="OperationalSafetyGuard"/>.</summary>
    public string CanonicalTargetPath { get; }

    /// <summary>Initializes a new <see cref="OperationalSafetyException"/> for a denied policy decision.</summary>
    public OperationalSafetyException(SafetyPolicyDecision decision, string canonicalTargetPath)
        : base($"{decision.ReasonCode}: {decision.Message} (Target: '{canonicalTargetPath}')")
    {
        Decision = decision ?? throw new ArgumentNullException(nameof(decision));
        CanonicalTargetPath = canonicalTargetPath ?? "";
    }
}

/// <summary>
/// Enforces runtime operational data safety modes (<see cref="OperationalSafetyMode"/>) with real filesystem
/// path canonicalization, NTFS Alternate Data Stream rejection, and symlink/junction resolution before
/// delegating policy checks to <see cref="OperationalSafetyPolicy"/>.
/// </summary>
public sealed class OperationalSafetyGuard
{
    private const int MaxReparseHops = 16;

    /// <summary>Active runtime safety mode.</summary>
    public OperationalSafetyMode Mode { get; }

    /// <summary>Canonical root directory containing read-only production source workbooks.</summary>
    public string ProductionDataRoot { get; }

    /// <summary>Canonical path to the operational Daily Log workbook.</summary>
    public string OperationalDailyLogPath { get; }

    /// <summary>Canonical root directory for isolated temporary writes.</summary>
    public string IsolatedTempRoot { get; }

    /// <summary>Initializes a new safety guard for the given mode and protected filesystem paths.</summary>
    public OperationalSafetyGuard(
        OperationalSafetyMode mode,
        string productionDataRoot,
        string operationalDailyLogPath,
        string? isolatedTempRoot = null)
    {
        Mode = mode;
        ProductionDataRoot = CanonicalizePath(productionDataRoot);
        OperationalDailyLogPath = CanonicalizePath(operationalDailyLogPath);
        string resolvedTemp = string.IsNullOrWhiteSpace(isolatedTempRoot)
            ? Path.Combine(Path.GetTempPath(), "AcrWorkstation")
            : isolatedTempRoot;
        IsolatedTempRoot = CanonicalizePath(resolvedTemp);

        if (OperationalSafetyPolicy.IsSameProtectedFile(IsolatedTempRoot, OperationalDailyLogPath) ||
            OperationalSafetyPolicy.IsUnderProtectedRoot(IsolatedTempRoot, ProductionDataRoot))
        {
            throw new ArgumentException(
                "IsolatedTempRoot must not point to the operational Daily Log or reside inside ProductionDataRoot.",
                nameof(isolatedTempRoot));
        }
    }

    /// <summary>
    /// Canonicalizes a filesystem path by rejecting NTFS Alternate Data Streams, normalizing extended prefixes,
    /// calling <see cref="Path.GetFullPath(string)"/>, and resolving existing ancestor junctions or symlinks.
    /// </summary>
    public string CanonicalizePath(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        string cleaned = StripExtendedPrefixAndValidateAds(path, nameof(path));
        string fullPath = Path.GetFullPath(cleaned);
        ValidateNoAdsColon(fullPath, nameof(path));
        ValidateNoTrailingSpaceOrDotInDirectorySegments(fullPath, nameof(path));
        string resolved = ResolveDeepestExistingLinkTarget(fullPath, nameof(path));
        ValidateNoTrailingSpaceOrDotInDirectorySegments(resolved, nameof(path));
        return TrimTrailingDirectorySeparators(resolved);
    }

    /// <summary>Classifies a target path against protected production roots and isolated temp directories.</summary>
    public OperationalTargetKind ClassifyTarget(string path, OperationalTargetKind? explicitKind = null)
    {
        string canonical = CanonicalizePath(path);
        if (OperationalSafetyPolicy.IsSameProtectedFile(canonical, OperationalDailyLogPath))
            return OperationalTargetKind.OperationalDailyLog;

        if (OperationalSafetyPolicy.IsUnderProtectedRoot(canonical, ProductionDataRoot))
            return OperationalTargetKind.ProductionSourceWorkbook;

        if (OperationalSafetyPolicy.IsUnderProtectedRoot(canonical, IsolatedTempRoot))
            return OperationalTargetKind.IsolatedTempCopy;

        return explicitKind ?? OperationalTargetKind.ClosureExportOutput;
    }

    /// <summary>Evaluates read access policy for <paramref name="path"/> without throwing.</summary>
    public SafetyPolicyDecision EvaluateRead(string path)
    {
        OperationalTargetKind kind = ClassifyTarget(path);
        return OperationalSafetyPolicy.EvaluateAccess(Mode, kind, OperationalActionKind.Read);
    }

    /// <summary>Validates read access for <paramref name="path"/> and returns its canonicalized path.</summary>
    public string EnsureCanRead(string path)
    {
        string canonical = CanonicalizePath(path);
        OperationalTargetKind kind = ClassifyTarget(canonical);
        SafetyPolicyDecision decision = OperationalSafetyPolicy.EvaluateAccess(Mode, kind, OperationalActionKind.Read);
        if (!decision.IsAllowed)
            throw new OperationalSafetyException(decision, canonical);

        return canonical;
    }

    /// <summary>Evaluates write access policy for <paramref name="path"/> without throwing.</summary>
    public SafetyPolicyDecision EvaluateWrite(
        string path,
        OperationalActionKind action,
        OperationalTargetKind? explicitKind = null)
    {
        OperationalTargetKind kind = ClassifyTarget(path, explicitKind);
        return OperationalSafetyPolicy.EvaluateAccess(Mode, kind, action);
    }

    /// <summary>
    /// Enforces write access policy for <paramref name="path"/>, throwing <see cref="OperationalSafetyException"/>
    /// when denied or when <see cref="OperationalSafetyMode.Testing"/> requires temp-only isolation.
    /// </summary>
    public string EnsureCanWrite(
        string path,
        OperationalActionKind action,
        OperationalTargetKind? explicitKind = null)
    {
        string canonical = CanonicalizePath(path);
        OperationalTargetKind kind = ClassifyTarget(canonical, explicitKind);
        SafetyPolicyDecision decision = OperationalSafetyPolicy.EvaluateAccess(Mode, kind, action);

        if (!decision.IsAllowed)
            throw new OperationalSafetyException(decision, canonical);

        if (decision.MustRedirectToTemp && kind != OperationalTargetKind.IsolatedTempCopy)
        {
            var redirectedDeny = new SafetyPolicyDecision(
                false,
                true,
                decision.ReasonCode,
                $"Testing mode prohibits writing outside IsolatedTempRoot ('{IsolatedTempRoot}').");
            throw new OperationalSafetyException(redirectedDeny, canonical);
        }

        return canonical;
    }

    /// <summary>
    /// Resolves a safe write destination, redirecting permitted export/artifact writes into
    /// <see cref="IsolatedTempRoot"/> when running in <see cref="OperationalSafetyMode.Testing"/>.
    /// </summary>
    public string ResolveSafeWritePath(
        string requestedPath,
        OperationalActionKind action,
        OperationalTargetKind? explicitKind = null)
    {
        string canonical = CanonicalizePath(requestedPath);
        OperationalTargetKind kind = ClassifyTarget(canonical, explicitKind);
        SafetyPolicyDecision decision = OperationalSafetyPolicy.EvaluateAccess(Mode, kind, action);

        if (!decision.IsAllowed)
            throw new OperationalSafetyException(decision, canonical);

        if (decision.MustRedirectToTemp && kind != OperationalTargetKind.IsolatedTempCopy)
        {
            string fileName = Path.GetFileName(canonical);
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("Target path must include a filename.", nameof(requestedPath));

            return CanonicalizePath(Path.Combine(IsolatedTempRoot, fileName));
        }

        return canonical;
    }

    private static string StripExtendedPrefixAndValidateAds(string path, string paramName)
    {
        string trimmed = path.Trim();
        if (trimmed.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            trimmed = @"\\" + trimmed[8..];
        else if (trimmed.StartsWith(@"\\?\", StringComparison.OrdinalIgnoreCase))
            trimmed = trimmed[4..];

        ValidateNoAdsColon(trimmed, paramName);
        return trimmed;
    }

    private static void ValidateNoAdsColon(string candidate, string paramName)
    {
        if (candidate.StartsWith(@"\\", StringComparison.Ordinal))
        {
            if (candidate.IndexOf(':') >= 0)
                throw new ArgumentException("NTFS Alternate Data Streams are not permitted in UNC paths.", paramName);
            return;
        }

        if (candidate.Length >= 3 && candidate.IndexOf(':', 2) >= 0)
            throw new ArgumentException("NTFS Alternate Data Streams are not permitted.", paramName);
    }

    private static void ValidateNoTrailingSpaceOrDotInDirectorySegments(string fullPath, string paramName)
    {
        string? root = Path.GetPathRoot(fullPath);
        string? dir = Path.GetDirectoryName(fullPath);
        while (!string.IsNullOrEmpty(dir) &&
               !string.Equals(dir, root, StringComparison.OrdinalIgnoreCase))
        {
            string segment = Path.GetFileName(dir);
            if (segment.EndsWith(' ') || segment.EndsWith('.'))
            {
                throw new ArgumentException(
                    $"Invalid directory segment '{segment}': trailing spaces or dots are not permitted.",
                    paramName);
            }

            string? parent = Path.GetDirectoryName(dir);
            if (string.IsNullOrEmpty(parent) || string.Equals(parent, dir, StringComparison.Ordinal))
                break;
            dir = parent;
        }
    }

    private static string ResolveDeepestExistingLinkTarget(string fullPath, string paramName)
    {
        string current = fullPath;
        for (int hop = 0; hop < MaxReparseHops; hop++)
        {
            if (!TryResolveSingleReparseStep(current, out string rewritten))
                return current;

            string cleaned = StripExtendedPrefixAndValidateAds(rewritten, paramName);
            current = Path.GetFullPath(cleaned);
            ValidateNoAdsColon(current, paramName);
        }
        return current;
    }

    private static bool TryResolveSingleReparseStep(string fullPath, out string rewritten)
    {
        var tail = new Stack<string>();
        string? probe = fullPath;

        while (!string.IsNullOrEmpty(probe))
        {
            string? parent = Path.GetDirectoryName(probe);
            if (string.IsNullOrEmpty(parent) || string.Equals(parent, probe, StringComparison.Ordinal))
                break;

            string? target = TryGetReparseTarget(probe);
            if (!string.IsNullOrEmpty(target))
            {
                rewritten = CombineResolvedTail(target, tail);
                return !string.Equals(
                    TrimTrailingDirectorySeparators(rewritten),
                    TrimTrailingDirectorySeparators(fullPath),
                    StringComparison.OrdinalIgnoreCase);
            }

            string leaf = Path.GetFileName(probe);
            if (!string.IsNullOrEmpty(leaf))
                tail.Push(leaf);
            probe = parent;
        }

        rewritten = fullPath;
        return false;
    }

    private static string? TryGetReparseTarget(string probe)
    {
        try
        {
            if (File.Exists(probe))
                return File.ResolveLinkTarget(probe, returnFinalTarget: true)?.FullName;

            if (Directory.Exists(probe))
                return Directory.ResolveLinkTarget(probe, returnFinalTarget: true)?.FullName;
        }
        catch (Exception ex) when (ex is DirectoryNotFoundException or FileNotFoundException)
        {
            return null;
        }

        return null;
    }

    private static string CombineResolvedTail(string baseTarget, Stack<string> tail)
    {
        string combined = baseTarget;
        while (tail.Count > 0)
            combined = Path.Combine(combined, tail.Pop());
        return combined;
    }

    private static string TrimTrailingDirectorySeparators(string path)
    {
        if (path.Length <= 3 && path.Length >= 2 && path[1] == ':')
            return path;

        return path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }
}
