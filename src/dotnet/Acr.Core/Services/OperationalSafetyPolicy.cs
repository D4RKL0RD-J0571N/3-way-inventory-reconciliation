using Acr.Core.Models;

namespace Acr.Core.Services;

/// <summary>
/// Pure domain safety policy evaluator enforcing read/write rules across <see cref="OperationalSafetyMode"/> values
/// without performing any filesystem I/O.
/// </summary>
public static class OperationalSafetyPolicy
{
    /// <summary>
    /// Returns false in all modes because source QC and closure workbooks in production data are strictly read-only.
    /// </summary>
    public static bool CanWriteProductionSource(OperationalSafetyMode mode) => false;

    /// <summary>
    /// Returns true only in <see cref="OperationalSafetyMode.Production"/> when triggered by an explicit user action.
    /// </summary>
    public static bool CanWriteOperationalDailyLog(OperationalSafetyMode mode, bool isExplicitUserAction) =>
        mode == OperationalSafetyMode.Production && isExplicitUserAction;

    /// <summary>
    /// Returns true when the runtime mode requires all write operations to target isolated temporary directories.
    /// </summary>
    public static bool MustRedirectWritesToTemp(OperationalSafetyMode mode) =>
        mode == OperationalSafetyMode.Testing;

    /// <summary>
    /// Evaluates whether a given <paramref name="action"/> is permitted on <paramref name="target"/> under <paramref name="mode"/>.
    /// </summary>
    public static SafetyPolicyDecision EvaluateAccess(
        OperationalSafetyMode mode,
        OperationalTargetKind target,
        OperationalActionKind action)
    {
        if (action == OperationalActionKind.Read)
            return new SafetyPolicyDecision(true, false, "READ_ALLOWED", "Read access is permitted across all modes.");

        return target switch
        {
            OperationalTargetKind.ProductionSourceWorkbook => new SafetyPolicyDecision(
                false,
                mode == OperationalSafetyMode.Testing,
                "PRODUCTION_SOURCE_READ_ONLY",
                "Source QC and template workbooks are strictly read-only in all operational modes."),
            OperationalTargetKind.OperationalDailyLog => EvaluateDailyLogWrite(mode, action),
            OperationalTargetKind.IsolatedTempCopy => new SafetyPolicyDecision(
                true, false, "TEMP_COPY_WRITE_ALLOWED", "Writes to isolated temporary copies are allowed."),
            OperationalTargetKind.ClosureExportOutput or OperationalTargetKind.LocalReportArtifact =>
                EvaluateExportOrArtifactWrite(mode, action),
            _ => new SafetyPolicyDecision(false, false, "UNKNOWN_TARGET", "Unrecognized operational target.")
        };
    }

    /// <summary>
    /// Normalizes a path string without filesystem calls by collapsing slash runs and lexically resolving
    /// <c>.</c> and <c>..</c> segments while preserving drive and UNC roots.
    /// </summary>
    public static string NormalizeComparablePath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return "";

        string trimmed = path.Trim().Replace('/', '\\');
        if (trimmed.StartsWith(@"\\", StringComparison.Ordinal))
            return NormalizeUncPath(trimmed);

        if (trimmed.Length >= 2 && char.IsAsciiLetter(trimmed[0]) && trimmed[1] == ':')
            return NormalizeDrivePath(trimmed);

        if (trimmed.StartsWith('\\'))
        {
            List<string> resolved = ResolveSegments(trimmed[1..], clampAtRoot: true);
            return resolved.Count == 0 ? @"\" : @"\" + string.Join(@"\", resolved).ToUpperInvariant();
        }

        List<string> relative = ResolveSegments(trimmed, clampAtRoot: false);
        return relative.Count == 0 ? "." : string.Join(@"\", relative).ToUpperInvariant();
    }

    /// <summary>
    /// Determines whether <paramref name="candidatePath"/> resides at or under <paramref name="protectedRootPath"/>.
    /// </summary>
    public static bool IsUnderProtectedRoot(string? candidatePath, string? protectedRootPath)
    {
        string candidate = NormalizeComparablePath(candidatePath);
        string root = NormalizeComparablePath(protectedRootPath);
        if (candidate.Length == 0 || root.Length == 0)
            return false;

        string prefix = root.EndsWith('\\') ? root : root + @"\";
        return string.Equals(candidate, root, StringComparison.Ordinal) ||
               candidate.StartsWith(prefix, StringComparison.Ordinal);
    }

    /// <summary>
    /// Determines whether two paths refer to the same protected file after pure string normalization.
    /// </summary>
    public static bool IsSameProtectedFile(string? candidatePath, string? protectedFilePath)
    {
        string candidate = NormalizeComparablePath(candidatePath);
        string protectedFile = NormalizeComparablePath(protectedFilePath);
        return candidate.Length > 0 && string.Equals(candidate, protectedFile, StringComparison.Ordinal);
    }

    private static string NormalizeUncPath(string trimmed)
    {
        string[] parts = trimmed.Split('\\', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0)
            return "";
        if (parts.Length == 1)
            return $@"\\{parts[0]}".ToUpperInvariant();

        string root = $@"\\{parts[0]}\{parts[1]}";
        string remainder = string.Join(@"\", parts.Skip(2));
        List<string> resolved = ResolveSegments(remainder, clampAtRoot: true);
        string full = resolved.Count == 0 ? root : $@"{root}\{string.Join(@"\", resolved)}";
        return full.ToUpperInvariant();
    }

    private static string NormalizeDrivePath(string trimmed)
    {
        string drive = char.ToUpperInvariant(trimmed[0]) + ":";
        string remainder = trimmed.Length > 2 ? trimmed[2..] : "";
        bool isRooted = remainder.StartsWith('\\');
        List<string> resolved = ResolveSegments(remainder, clampAtRoot: isRooted);

        if (resolved.Count == 0)
            return isRooted ? drive + @"\" : drive;

        string joined = string.Join(@"\", resolved).ToUpperInvariant();
        return isRooted ? $@"{drive}\{joined}" : $@"{drive}{joined}";
    }

    private static List<string> ResolveSegments(string segmentText, bool clampAtRoot)
    {
        var stack = new List<string>();
        foreach (string raw in segmentText.Split('\\', StringSplitOptions.RemoveEmptyEntries))
        {
            if (raw == ".")
                continue;

            if (raw == "..")
            {
                if (stack.Count > 0 && stack[^1] != "..")
                    stack.RemoveAt(stack.Count - 1);
                else if (!clampAtRoot)
                    stack.Add("..");
                continue;
            }

            stack.Add(raw);
        }

        return stack;
    }

    private static SafetyPolicyDecision EvaluateDailyLogWrite(
        OperationalSafetyMode mode,
        OperationalActionKind action)
    {
        return mode switch
        {
            OperationalSafetyMode.Development => new SafetyPolicyDecision(
                false, false, "DEV_DAILY_LOG_READ_ONLY",
                "Operational Daily Log is strictly read-only in Development mode."),
            OperationalSafetyMode.Testing => new SafetyPolicyDecision(
                false, true, "TESTING_REDIRECT_TO_TEMP",
                "Testing mode prohibits writing the live Daily Log; redirect writes to an isolated temporary copy."),
            OperationalSafetyMode.Production when action == OperationalActionKind.WriteExplicitUserSaveOrSync =>
                new SafetyPolicyDecision(true, false, "PROD_EXPLICIT_SAVE_ALLOWED",
                    "Production mode allows Daily Log updates via explicit user save/sync action."),
            OperationalSafetyMode.Production => new SafetyPolicyDecision(
                false, false, "PROD_REQUIRES_EXPLICIT_USER_ACTION",
                "Production mode requires explicit user confirmation before writing the Daily Log."),
            _ => new SafetyPolicyDecision(false, false, "UNSUPPORTED_MODE", "Unsupported safety mode.")
        };
    }

    private static SafetyPolicyDecision EvaluateExportOrArtifactWrite(
        OperationalSafetyMode mode,
        OperationalActionKind action)
    {
        if (mode == OperationalSafetyMode.Testing)
        {
            return new SafetyPolicyDecision(
                true, true, "TESTING_EXPORT_ISOLATED_TEMP",
                "Testing mode redirects closure exports and local artifacts to an isolated temporary directory.");
        }

        bool allowedAction = action is OperationalActionKind.WriteToNewExportPath
            or OperationalActionKind.WriteExplicitUserSaveOrSync;
        return allowedAction
            ? new SafetyPolicyDecision(true, false, "EXPORT_ALLOWED", "Export or artifact write to new destination is permitted.")
            : new SafetyPolicyDecision(false, false, "BACKGROUND_EXPORT_BLOCKED", "Automated background export without explicit action is blocked.");
    }
}
