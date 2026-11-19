using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Acr.Core.Models;
using Acr.Infrastructure.Safety;

namespace Acr.Infrastructure.Reporting;

/// <summary>
/// Summary of workbook formula and package-preservation checks attached to a <see cref="ProcessingReport"/>.
/// </summary>
public sealed record ClosureValidationSummary
{
    /// <summary>True when validation passed without blocking or exporter-introduced errors.</summary>
    public bool IsValid { get; init; }

    /// <summary>True when <c>xl/vbaProject.bin</c> SHA-256 matches the source template.</summary>
    public bool VbaProjectSha256Preserved { get; init; } = true;

    /// <summary>Source template <c>xl/vbaProject.bin</c> SHA-256 digest.</summary>
    public string? TemplateVbaSha256 { get; init; }

    /// <summary>Exported workbook <c>xl/vbaProject.bin</c> SHA-256 digest.</summary>
    public string? ExportedVbaSha256 { get; init; }

    /// <summary>Count of pre-existing source errors retained from the template.</summary>
    public int PreExistingErrorCount { get; init; }

    /// <summary>Count of errors newly introduced by the exporter (zero expected).</summary>
    public int ExporterIntroducedErrorCount { get; init; }

    /// <summary>Count of guarded optional <c>VLOOKUP</c> cells classified as legitimate.</summary>
    public int LegitimateOptionalNaCount { get; init; }

    /// <summary>Non-blocking validation warnings.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();

    /// <summary>Blocking validation errors.</summary>
    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();
}

/// <summary>
/// Structured local JSON processing report generated for each receiving reconciliation run.
/// </summary>
public sealed record ProcessingReport
{
    /// <summary>Unique identifier for this processing report.</summary>
    public string ReportId { get; init; } = Guid.NewGuid().ToString("N");

    /// <summary>UTC timestamp when the reconciliation completed.</summary>
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Container identifier resolved from template or QC header.</summary>
    public string ContainerId { get; init; } = "";

    /// <summary>Canonicalized QC source workbook path.</summary>
    public string QcSourcePath { get; init; } = "";

    /// <summary>Canonicalized closure template workbook path.</summary>
    public string ClosureTemplatePath { get; init; } = "";

    /// <summary>Destination path of exported closure workbook, when exported.</summary>
    public string? ExportDestinationPath { get; init; }

    /// <summary>Total distinct materials compared.</summary>
    public int MaterialsProcessedCount { get; init; }

    /// <summary>Count of materials classified as exact match.</summary>
    public int CleanMaterialsCount { get; init; }

    /// <summary>Count of materials presenting any discrepancy.</summary>
    public int DiscrepantMaterialsCount { get; init; }

    /// <summary>Total declared units from <c>CIERRE RECEP</c>.</summary>
    public int TotalDeclaredUnits { get; init; }

    /// <summary>Total physical units from QC inspection.</summary>
    public int TotalPhysicalUnits { get; init; }

    /// <summary>Total declared cartons from <c>CIERRE RECEP</c> column <c>AW</c>.</summary>
    public int TotalDeclaredCartons { get; init; }

    /// <summary>Total physical cartons from QC inspection.</summary>
    public int TotalPhysicalCartons { get; init; }

    /// <summary>Total shortage units across discrepant materials.</summary>
    public int TotalShortageUnits { get; init; }

    /// <summary>Total overage units across discrepant materials.</summary>
    public int TotalOverageUnits { get; init; }

    /// <summary>Total damaged units recorded.</summary>
    public int TotalDamagedUnits { get; init; }

    /// <summary>Count of candidate revuelto pairs detected.</summary>
    public int RevueltoPairsCount { get; init; }

    /// <summary>Detailed auditable revuelto pairs detected.</summary>
    public IReadOnlyList<RevueltoMatch> RevueltoMatches { get; init; } = Array.Empty<RevueltoMatch>();

    /// <summary>Pre-print closure validation summary when export occurred.</summary>
    public ClosureValidationSummary? ValidationResult { get; init; }

    /// <summary>Pipeline warnings accumulated during reading, comparison, or validation.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
}

/// <summary>
/// Local JSON persistence store for <see cref="ProcessingReport"/> artifacts using
/// <see cref="OperationalSafetyGuard"/> and <see cref="AtomicFileTransaction"/>.
/// </summary>
public static class JsonProcessingReportStore
{
    internal static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>
    /// Atomically saves <paramref name="report"/> as an indented JSON file inside <paramref name="reportsDirectory"/>.
    /// </summary>
    public static string Save(
        ProcessingReport report,
        string reportsDirectory,
        OperationalSafetyGuard? guard = null)
    {
        ArgumentNullException.ThrowIfNull(report);
        ArgumentException.ThrowIfNullOrWhiteSpace(reportsDirectory);

        OperationalSafetyGuard effectiveGuard = guard ?? CreateDefaultPermissiveGuard();
        string resolvedDir = Path.GetFullPath(reportsDirectory);
        string fileName = BuildReportFileName(report);
        string requestedPath = Path.GetFullPath(Path.Combine(resolvedDir, fileName));
        string dirPrefix = resolvedDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!requestedPath.StartsWith(dirPrefix, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"Report path '{requestedPath}' escapes directory '{resolvedDir}'.");

        string safeTarget = effectiveGuard.ResolveSafeWritePath(
            requestedPath,
            OperationalActionKind.WriteToNewExportPath,
            OperationalTargetKind.LocalReportArtifact);

        AtomicTransactionResult tx = AtomicFileTransaction.Execute(
            safeTarget,
            effectiveGuard,
            OperationalActionKind.WriteToNewExportPath,
            tempPath => WriteJsonAtomic(tempPath, report),
            ValidateReportTempFile,
            committedPath => File.Exists(committedPath) && new FileInfo(committedPath).Length > 0,
            OperationalTargetKind.LocalReportArtifact);

        return tx.TargetPath;
    }

    /// <summary>
    /// Loads a single <see cref="ProcessingReport"/> from a local JSON file.
    /// </summary>
    public static ProcessingReport Load(string reportFilePath, OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(reportFilePath);
        string resolved = guard is not null ? guard.EnsureCanRead(reportFilePath) : Path.GetFullPath(reportFilePath);
        if (!File.Exists(resolved))
            throw new FileNotFoundException($"Processing report not found: '{resolved}'.", resolved);

        string json = File.ReadAllText(resolved, Encoding.UTF8);
        return JsonSerializer.Deserialize<ProcessingReport>(json, JsonOptions)
            ?? throw new InvalidDataException($"Deserialized null ProcessingReport from '{resolved}'.");
    }

    /// <summary>
    /// Loads all <see cref="ProcessingReport"/> JSON files in <paramref name="reportsDirectory"/>,
    /// ordered descending by <see cref="ProcessingReport.TimestampUtc"/>.
    /// </summary>
    public static IReadOnlyList<ProcessingReport> LoadAll(
        string reportsDirectory,
        OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(reportsDirectory);
        string resolvedDir = guard is not null ? guard.EnsureCanRead(reportsDirectory) : Path.GetFullPath(reportsDirectory);
        if (!Directory.Exists(resolvedDir))
            return Array.Empty<ProcessingReport>();

        var reports = new List<ProcessingReport>();
        foreach (string file in Directory.EnumerateFiles(resolvedDir, "report_*.json", SearchOption.TopDirectoryOnly))
        {
            if (file.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase) ||
                file.EndsWith(".bak", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
            reports.Add(Load(file, guard));
        }

        return reports.OrderByDescending(r => r.TimestampUtc).ToList();
    }

    private static OperationalSafetyGuard CreateDefaultPermissiveGuard()
    {
        string sentinelRoot = Path.Combine(Path.GetTempPath(), $"AcrNoProdReport_{Guid.NewGuid():N}");
        return new OperationalSafetyGuard(
            OperationalSafetyMode.Production,
            sentinelRoot,
            Path.Combine(sentinelRoot, "NoDailyLog.xlsx"));
    }

    private static string BuildReportFileName(ProcessingReport report)
    {
        string safeContainer = SanitizeFileToken(report.ContainerId, maxLength: 40);
        string shortId = SanitizeFileToken(report.ReportId, maxLength: 16);
        return $"report_{report.TimestampUtc:yyyyMMdd_HHmmss}_{safeContainer}_{shortId}.json";
    }

    private static string SanitizeFileToken(string? raw, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return "UNKNOWN";

        var sb = new StringBuilder(Math.Min(raw.Length, maxLength));
        foreach (char ch in raw.Trim())
        {
            if (sb.Length >= maxLength)
                break;
            sb.Append(char.IsLetterOrDigit(ch) || ch == '-' || ch == '_' ? ch : '_');
        }
        return sb.Length == 0 ? "UNKNOWN" : sb.ToString();
    }

    private static void WriteJsonAtomic(string tempPath, ProcessingReport report)
    {
        string json = JsonSerializer.Serialize(report, JsonOptions);
        File.WriteAllText(tempPath, json, Encoding.UTF8);
    }

    private static bool ValidateReportTempFile(string tempPath)
    {
        try
        {
            string json = File.ReadAllText(tempPath, Encoding.UTF8);
            ProcessingReport? parsed = JsonSerializer.Deserialize<ProcessingReport>(json, JsonOptions);
            return parsed is not null && !string.IsNullOrWhiteSpace(parsed.ReportId);
        }
        catch (JsonException)
        {
            return false;
        }
    }
}
