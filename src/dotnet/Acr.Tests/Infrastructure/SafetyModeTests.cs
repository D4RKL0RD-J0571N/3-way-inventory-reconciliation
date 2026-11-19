using System.Diagnostics;
using System.Security.Cryptography;
using Acr.Core.Models;
using Acr.Infrastructure.Safety;

namespace Acr.Tests.Infrastructure;

public sealed class SafetyModeTests : IDisposable
{
    private readonly string _sandboxRoot;
    private readonly string _prodDataRoot;
    private readonly string _dailyLogPath;
    private readonly string _isolatedTempRoot;

    public SafetyModeTests()
    {
        _sandboxRoot = Path.Combine(Path.GetTempPath(), $"AcrSafetyTests_{Guid.NewGuid():N}");
        _prodDataRoot = Path.Combine(_sandboxRoot, "prod_data");
        _dailyLogPath = Path.Combine(_sandboxRoot, "Control_Diario_Operaciones_Jostin.xlsx");
        _isolatedTempRoot = Path.Combine(_sandboxRoot, "isolated_temp");

        Directory.CreateDirectory(_prodDataRoot);
        Directory.CreateDirectory(_isolatedTempRoot);
        File.WriteAllText(Path.Combine(_prodDataRoot, "IMP-100.xlsm"), "prod-source");
        File.WriteAllText(_dailyLogPath, "daily-log-v1");
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_sandboxRoot))
                Directory.Delete(_sandboxRoot, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    [Fact]
    public void DevelopmentMode_BlocksWritesToProductionDataAndDailyLog_ViaTraversalPaths()
    {
        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Development, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);
        string traversalProd = Path.Combine(_isolatedTempRoot, "..", "prod_data", "IMP-100.xlsm");
        string traversalDaily = Path.Combine(_prodDataRoot, "..", "Control_Diario_Operaciones_Jostin.xlsx");

        string readProd = guard.EnsureCanRead(traversalProd);
        string readDaily = guard.EnsureCanRead(traversalDaily);
        Assert.Equal(guard.CanonicalizePath(Path.Combine(_prodDataRoot, "IMP-100.xlsm")), readProd);
        Assert.Equal(guard.CanonicalizePath(_dailyLogPath), readDaily);

        OperationalSafetyException prodEx = Assert.Throws<OperationalSafetyException>(
            () => guard.EnsureCanWrite(traversalProd, OperationalActionKind.WriteExplicitUserSaveOrSync));
        Assert.Equal("PRODUCTION_SOURCE_READ_ONLY", prodEx.Decision.ReasonCode);

        OperationalSafetyException dailyEx = Assert.Throws<OperationalSafetyException>(
            () => guard.EnsureCanWrite(traversalDaily, OperationalActionKind.WriteExplicitUserSaveOrSync));
        Assert.Equal("DEV_DAILY_LOG_READ_ONLY", dailyEx.Decision.ReasonCode);
    }

    [Fact]
    public void CanonicalizePath_HandlesCaseInsensitivityExtendedPrefixAndRejectsNtfsAds()
    {
        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Development, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);
        string mixedCase = Path.Combine(_prodDataRoot.ToUpperInvariant(), "imp-100.xlsm").Replace('\\', '/');
        string extended = @"\\?\" + Path.Combine(_prodDataRoot, "IMP-100.xlsm") + Path.DirectorySeparatorChar;

        Assert.Equal(OperationalTargetKind.ProductionSourceWorkbook, guard.ClassifyTarget(mixedCase));
        Assert.Equal(OperationalTargetKind.ProductionSourceWorkbook, guard.ClassifyTarget(extended));
        Assert.Equal(guard.CanonicalizePath(mixedCase), guard.CanonicalizePath(extended), StringComparer.OrdinalIgnoreCase);

        string adsPath = Path.Combine(_isolatedTempRoot, "test.xlsx:Zone.Identifier");
        Assert.Throws<ArgumentException>(() => guard.CanonicalizePath(adsPath));
    }

    [Fact]
    public void JunctionIntoProtectedRoot_IsBlockedForWrites()
    {
        string junctionPath = Path.Combine(_sandboxRoot, "junction_to_prod");
        if (!TryCreateNtfsJunction(junctionPath, _prodDataRoot))
            return;

        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Production, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);
        string writeThroughJunction = Path.Combine(junctionPath, "new_file.xlsm");

        Assert.Equal(OperationalTargetKind.ProductionSourceWorkbook, guard.ClassifyTarget(writeThroughJunction));
        OperationalSafetyException ex = Assert.Throws<OperationalSafetyException>(
            () => guard.EnsureCanWrite(writeThroughJunction, OperationalActionKind.WriteToNewExportPath));
        Assert.Equal("PRODUCTION_SOURCE_READ_ONLY", ex.Decision.ReasonCode);
    }

    [Fact]
    public void TrailingSpaceAndDotSpaceIntermediateDirectories_FailClosedWithoutDirectoryNotFound()
    {
        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Production, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);
        string[] invalidPaths =
        [
            Path.Combine(_sandboxRoot, "prod_data ", "IMP-100.xlsm"),
            Path.Combine(_sandboxRoot, "prod_data. ", "IMP-100.xlsm"),
            @"\\?\" + Path.Combine(_sandboxRoot, "prod_data ", "IMP-100.xlsm")
        ];

        foreach (string invalid in invalidPaths)
        {
            Assert.Throws<ArgumentException>(() => guard.CanonicalizePath(invalid));
            Assert.Throws<ArgumentException>(() => guard.ClassifyTarget(invalid));
            Assert.Throws<ArgumentException>(() => guard.EnsureCanRead(invalid));
            Assert.Throws<ArgumentException>(() => guard.EnsureCanWrite(invalid, OperationalActionKind.WriteToNewExportPath));
            Assert.Throws<ArgumentException>(() => guard.ResolveSafeWritePath(invalid, OperationalActionKind.WriteToNewExportPath));
        }

        string trailingDotDir = Path.Combine(_sandboxRoot, "prod_data.", "IMP-100.xlsm");
        string trailingSpaceFile = Path.Combine(_prodDataRoot, "IMP-100.xlsm ");
        Assert.Equal(OperationalTargetKind.ProductionSourceWorkbook, guard.ClassifyTarget(trailingDotDir));
        Assert.Equal(OperationalTargetKind.ProductionSourceWorkbook, guard.ClassifyTarget(trailingSpaceFile));
        Assert.Throws<OperationalSafetyException>(() => guard.EnsureCanWrite(trailingDotDir, OperationalActionKind.WriteToNewExportPath));
        Assert.Throws<OperationalSafetyException>(() => guard.EnsureCanWrite(trailingSpaceFile, OperationalActionKind.WriteToNewExportPath));
    }

    [Fact]
    public void TestingMode_AllowsIsolatedTempWrites_BlocksDailyLog_AndRedirectsExports()
    {
        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Testing, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);
        string tempFile = Path.Combine(_isolatedTempRoot, "copy.xlsx");
        string outsideExport = Path.Combine(_sandboxRoot, "outside_export.xlsm");

        Assert.Equal(guard.CanonicalizePath(tempFile), guard.EnsureCanWrite(tempFile, OperationalActionKind.WriteExplicitUserSaveOrSync));

        OperationalSafetyException dailyEx = Assert.Throws<OperationalSafetyException>(
            () => guard.EnsureCanWrite(_dailyLogPath, OperationalActionKind.WriteExplicitUserSaveOrSync));
        Assert.Equal("TESTING_REDIRECT_TO_TEMP", dailyEx.Decision.ReasonCode);

        Assert.Throws<OperationalSafetyException>(
            () => guard.EnsureCanWrite(outsideExport, OperationalActionKind.WriteToNewExportPath, OperationalTargetKind.ClosureExportOutput));

        string redirected = guard.ResolveSafeWritePath(
            outsideExport,
            OperationalActionKind.WriteToNewExportPath,
            OperationalTargetKind.ClosureExportOutput);
        Assert.StartsWith(guard.IsolatedTempRoot, redirected, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("outside_export.xlsm", Path.GetFileName(redirected));
    }

    [Fact]
    public void ProductionMode_RequiresExplicitUserActionForDailyLog_AndProtectsSourceData()
    {
        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Production, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);

        OperationalSafetyException bgEx = Assert.Throws<OperationalSafetyException>(
            () => guard.EnsureCanWrite(_dailyLogPath, OperationalActionKind.WriteAutomatedOrBackground));
        Assert.Equal("PROD_REQUIRES_EXPLICIT_USER_ACTION", bgEx.Decision.ReasonCode);

        string allowed = guard.EnsureCanWrite(_dailyLogPath, OperationalActionKind.WriteExplicitUserSaveOrSync);
        Assert.Equal(guard.CanonicalizePath(_dailyLogPath), allowed);

        Assert.Throws<OperationalSafetyException>(
            () => guard.EnsureCanWrite(Path.Combine(_prodDataRoot, "IMP-100.xlsm"), OperationalActionKind.WriteExplicitUserSaveOrSync));
    }

    [Fact]
    public void AtomicFileTransaction_RejectsOfficeLockFile_AndLeavesTargetUntouched()
    {
        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Testing, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);
        string target = Path.Combine(_isolatedTempRoot, "DailyCopy.xlsx");
        File.WriteAllText(target, "original-content");
        string originalHash = Sha256Hex(target);
        File.WriteAllText(Path.Combine(_isolatedTempRoot, "~$DailyCopy.xlsx"), "lock");

        Assert.True(AtomicFileTransaction.HasOfficeLockFile(target));
        Assert.Throws<IOException>(() => AtomicFileTransaction.Execute(
            target,
            guard,
            OperationalActionKind.WriteExplicitUserSaveOrSync,
            tmp => File.WriteAllText(tmp, "updated")));

        Assert.Equal(originalHash, Sha256Hex(target));
        Assert.False(File.Exists(target + ".bak"));
        Assert.Empty(Directory.GetFiles(_isolatedTempRoot, "*.tmp"));
    }

    [Fact]
    public void AtomicFileTransaction_RejectsLockedTargetViaExclusiveHandleCheck()
    {
        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Testing, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);
        string target = Path.Combine(_isolatedTempRoot, "LockedCopy.xlsx");
        File.WriteAllText(target, "original-content");

        using var lockStream = new FileStream(target, FileMode.Open, FileAccess.Read, FileShare.Read);
        Assert.Throws<IOException>(() => AtomicFileTransaction.Execute(
            target,
            guard,
            OperationalActionKind.WriteExplicitUserSaveOrSync,
            tmp => File.WriteAllText(tmp, "updated")));

        Assert.False(File.Exists(target + ".bak"));
        Assert.Empty(Directory.GetFiles(_isolatedTempRoot, "*.tmp"));
    }

    [Fact]
    public void AtomicFileTransaction_PreCommitValidationFailure_RollsBackTempWithoutBackupOrOverwrite()
    {
        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Testing, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);
        string target = Path.Combine(_isolatedTempRoot, "ValidatedCopy.xlsx");
        File.WriteAllText(target, "original-v1");
        string originalHash = Sha256Hex(target);

        Assert.Throws<InvalidOperationException>(() => AtomicFileTransaction.Execute(
            target,
            guard,
            OperationalActionKind.WriteExplicitUserSaveOrSync,
            tmp => File.WriteAllText(tmp, "corrupt-v2"),
            validateTempFile: _ => false));

        Assert.Equal(originalHash, Sha256Hex(target));
        Assert.False(File.Exists(target + ".bak"));
        Assert.Empty(Directory.GetFiles(_isolatedTempRoot, "*.tmp"));
    }

    [Fact]
    public void AtomicFileTransaction_SuccessfulCommit_CreatesBackupUpdatesTargetAndCleansTemp()
    {
        var guard = new OperationalSafetyGuard(OperationalSafetyMode.Testing, _prodDataRoot, _dailyLogPath, _isolatedTempRoot);
        string target = Path.Combine(_isolatedTempRoot, "AtomicTarget.xlsx");
        File.WriteAllText(target, "version-1");
        string v1Hash = Sha256Hex(target);
        bool postVerified = false;

        AtomicTransactionResult result = AtomicFileTransaction.Execute(
            target,
            guard,
            OperationalActionKind.WriteExplicitUserSaveOrSync,
            tmp => File.WriteAllText(tmp, "version-2"),
            validateTempFile: tmp => File.ReadAllText(tmp) == "version-2",
            verifyCommittedFile: committed =>
            {
                postVerified = File.ReadAllText(committed) == "version-2";
                return postVerified;
            });

        Assert.True(result.BackupCreated);
        Assert.True(result.Verified);
        Assert.True(postVerified);
        Assert.NotNull(result.BackupPath);
        Assert.Equal(v1Hash, Sha256Hex(result.BackupPath!));
        Assert.Equal("version-2", File.ReadAllText(target));
        Assert.Empty(Directory.GetFiles(_isolatedTempRoot, "*.tmp"));
    }

    private static string Sha256Hex(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static bool TryCreateNtfsJunction(string junctionPath, string targetPath)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c mklink /J \"{junctionPath}\" \"{targetPath}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            });
            process?.WaitForExit(5000);
            return process?.ExitCode == 0 && Directory.Exists(junctionPath);
        }
        catch (Exception)
        {
            return false;
        }
    }
}
