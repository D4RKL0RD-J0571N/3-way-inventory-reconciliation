using Acr.Core.Models;
using Acr.Core.Services;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Reporting;
using Acr.Infrastructure.Safety;

namespace Acr.Infrastructure.Receiving;

/// <summary>
/// Configuration options controlling a single end-to-end receiving reconciliation execution.
/// </summary>
public sealed record ReceivingPipelineOptions
{
    /// <summary>Path to the Quality Control (<c>CONTROL DATA</c>) workbook.</summary>
    public required string QcFilePath { get; init; }

    /// <summary>Path to the closure template (<c>CIERRE RECEP</c>) workbook.</summary>
    public required string ClosureTemplatePath { get; init; }

    /// <summary>Optional destination path to export a populated macro-safe closure workbook.</summary>
    public string? ExportOutputPath { get; init; }

    /// <summary>Optional directory where a local JSON <see cref="ProcessingReport"/> is persisted.</summary>
    public string? ReportsDirectory { get; init; }

    /// <summary>Optional local JSON operation history file path to append an entry into.</summary>
    public string? HistoryFilePath { get; init; }

    /// <summary>Optional operational safety guard enforcing Development, Testing, or Production access rules.</summary>
    public OperationalSafetyGuard? SafetyGuard { get; init; }
}

/// <summary>
/// Complete structured output produced by <see cref="ReceivingPipeline.Execute"/>.
/// </summary>
public sealed record ReceivingPipelineResult
{
    /// <summary>Resolved container identifier.</summary>
    public required string ContainerId { get; init; }

    /// <summary>Parsed QC inspection report.</summary>
    public required QcInspectionReport QcReport { get; init; }

    /// <summary>Declared items extracted from <c>CIERRE RECEP</c>.</summary>
    public required IReadOnlyList<DeclaredItem> DeclaredItems { get; init; }

    /// <summary>Primary (non-child) physical items mapped from <see cref="QcReport"/>.</summary>
    public required IReadOnlyList<PhysicalItem> PhysicalItems { get; init; }

    /// <summary>Deterministic 3-way discrepancy comparison summary.</summary>
    public required ComparisonSummary Comparison { get; init; }

    /// <summary>Closure export result when <see cref="ReceivingPipelineOptions.ExportOutputPath"/> was specified.</summary>
    public ClosureExportResult? ExportResult { get; init; }

    /// <summary>Structured processing report generated for this run.</summary>
    public ProcessingReport? Report { get; init; }

    /// <summary>Saved JSON report file path when <see cref="ReceivingPipelineOptions.ReportsDirectory"/> was specified.</summary>
    public string? SavedReportPath { get; init; }

    /// <summary>Saved operation history JSON path when <see cref="ReceivingPipelineOptions.HistoryFilePath"/> was specified.</summary>
    public string? SavedHistoryPath { get; init; }

    /// <summary>Combined non-blocking warnings emitted during template reading or validation.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
}

/// <summary>
/// End-to-end receiving orchestrator bridging <see cref="QcReader"/>, <see cref="ClosureTemplateReader"/>,
/// <see cref="DiscrepancyClassifier"/>, <see cref="ClosureExporter"/>, <see cref="JsonProcessingReportStore"/>,
/// and <see cref="OperationHistoryStore"/>.
/// </summary>
public static class ReceivingPipeline
{
    /// <summary>
    /// Executes the full receiving reconciliation workflow and optional export/reporting steps.
    /// </summary>
    public static ReceivingPipelineResult Execute(ReceivingPipelineOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentException.ThrowIfNullOrWhiteSpace(options.QcFilePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(options.ClosureTemplatePath);

        QcInspectionReport qcReport = QcReader.Read(options.QcFilePath, options.SafetyGuard);
        ClosureTemplateData templateData = ClosureTemplateReader.ReadDeclaredItems(options.ClosureTemplatePath, options.SafetyGuard);
        IReadOnlyList<PhysicalItem> physicalItems = MapPhysicalItems(qcReport);
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(templateData.DeclaredItems, physicalItems);
        string containerId = ResolveContainerId(templateData, qcReport);

        ClosureExportResult? exportResult = ExportIfRequested(options, summary, qcReport);
        List<string> warnings = CollectWarnings(templateData, exportResult);
        ProcessingReport report = BuildProcessingReport(containerId, qcReport, templateData, summary, exportResult, warnings);

        string? savedReportPath = SaveReportIfRequested(options, report);
        string? savedHistoryPath = AppendHistoryIfRequested(options, containerId, summary, exportResult, savedReportPath);

        return new ReceivingPipelineResult
        {
            ContainerId = containerId,
            QcReport = qcReport,
            DeclaredItems = templateData.DeclaredItems,
            PhysicalItems = physicalItems,
            Comparison = summary,
            ExportResult = exportResult,
            Report = report,
            SavedReportPath = savedReportPath,
            SavedHistoryPath = savedHistoryPath,
            Warnings = warnings
        };
    }

    private static IReadOnlyList<PhysicalItem> MapPhysicalItems(QcInspectionReport qcReport)
    {
        var list = new List<PhysicalItem>();
        foreach (QcInspectionItem item in qcReport.Items)
        {
            if (item.IsChildSku)
                continue;

            list.Add(new PhysicalItem
            {
                MaterialCode = OpenXmlCellHelper.NormalizeMaterialCode(item.MaterialCode),
                VendorMaterial = string.IsNullOrWhiteSpace(item.VendorMaterial) ? null : item.VendorMaterial,
                Description = string.IsNullOrWhiteSpace(item.Description) ? null : item.Description,
                Category = string.IsNullOrWhiteSpace(item.Category) ? null : item.Category,
                SourceRow = item.RowIndex,
                PhysicalUnits = item.CalculationResult?.CalculatedUnits ?? 0,
                PhysicalCartons = item.CalculationResult?.CalculatedCartons ?? 0,
                DamagedUnits = 0,
                Uom = item.Uom,
                Flags = item.CalculationResult?.Flags ?? Array.Empty<string>()
            });
        }
        return list;
    }

    private static string ResolveContainerId(ClosureTemplateData templateData, QcInspectionReport qcReport)
    {
        if (!string.IsNullOrWhiteSpace(templateData.ContainerId))
            return templateData.ContainerId.Trim();
        if (!string.IsNullOrWhiteSpace(qcReport.ContainerId))
            return qcReport.ContainerId.Trim();
        return Path.GetFileNameWithoutExtension(templateData.FilePath);
    }

    private static ClosureExportResult? ExportIfRequested(
        ReceivingPipelineOptions options,
        ComparisonSummary summary,
        QcInspectionReport qcReport)
    {
        if (string.IsNullOrWhiteSpace(options.ExportOutputPath))
            return null;

        string safeOutput = options.SafetyGuard is not null
            ? options.SafetyGuard.ResolveSafeWritePath(
                options.ExportOutputPath,
                OperationalActionKind.WriteToNewExportPath,
                OperationalTargetKind.ClosureExportOutput)
            : Path.GetFullPath(options.ExportOutputPath);

        return ClosureExporter.Export(options.ClosureTemplatePath, safeOutput, summary, qcReport, options.SafetyGuard);
    }

    private static List<string> CollectWarnings(ClosureTemplateData templateData, ClosureExportResult? exportResult)
    {
        var warnings = new List<string>(templateData.Warnings);
        if (exportResult is not null)
            warnings.AddRange(exportResult.Validation.Warnings);
        return warnings;
    }

    private static ProcessingReport BuildProcessingReport(
        string containerId,
        QcInspectionReport qcReport,
        ClosureTemplateData templateData,
        ComparisonSummary summary,
        ClosureExportResult? exportResult,
        IReadOnlyList<string> warnings)
    {
        ClosureValidationSummary? validationSummary = exportResult is null
            ? null
            : BuildValidationSummary(exportResult.Validation);

        return new ProcessingReport
        {
            ContainerId = containerId,
            QcSourcePath = qcReport.FilePath,
            ClosureTemplatePath = templateData.FilePath,
            ExportDestinationPath = exportResult?.OutputPath,
            MaterialsProcessedCount = summary.Items.Count,
            CleanMaterialsCount = summary.CleanMaterialsCount,
            DiscrepantMaterialsCount = summary.DiscrepantMaterialsCount,
            TotalDeclaredUnits = summary.TotalDeclaredUnits,
            TotalPhysicalUnits = summary.TotalPhysicalUnits,
            TotalDeclaredCartons = summary.TotalDeclaredCartons,
            TotalPhysicalCartons = summary.TotalPhysicalCartons,
            TotalShortageUnits = summary.TotalShortageUnits,
            TotalOverageUnits = summary.TotalOverageUnits,
            TotalDamagedUnits = summary.TotalDamagedUnits,
            RevueltoPairsCount = summary.RevueltoPairsCount,
            RevueltoMatches = summary.RevueltoMatches,
            ValidationResult = validationSummary,
            Warnings = warnings
        };
    }

    private static ClosureValidationSummary BuildValidationSummary(PrePrintValidationResult val) =>
        new()
        {
            IsValid = val.IsValid,
            VbaProjectSha256Preserved = val.VbaProjectHashPreserved,
            TemplateVbaSha256 = val.SourceHashes?.VbaProjectSha256,
            ExportedVbaSha256 = val.OutputHashes?.VbaProjectSha256,
            PreExistingErrorCount = val.PreExistingSourceErrors.Count,
            ExporterIntroducedErrorCount = val.ExporterIntroducedErrors.Count,
            LegitimateOptionalNaCount = val.LegitimateOptionalNas.Count,
            Warnings = val.Warnings,
            Errors = val.Errors
        };

    private static string? SaveReportIfRequested(ReceivingPipelineOptions options, ProcessingReport report)
    {
        if (string.IsNullOrWhiteSpace(options.ReportsDirectory))
            return null;

        return JsonProcessingReportStore.Save(report, options.ReportsDirectory, options.SafetyGuard);
    }

    private static string? AppendHistoryIfRequested(
        ReceivingPipelineOptions options,
        string containerId,
        ComparisonSummary summary,
        ClosureExportResult? exportResult,
        string? savedReportPath)
    {
        if (string.IsNullOrWhiteSpace(options.HistoryFilePath))
            return null;

        var entry = new OperationHistoryEntry
        {
            Kind = exportResult is null ? OperationHistoryKind.Reconciliation : OperationHistoryKind.ClosureExport,
            ContainerId = containerId,
            QcSourcePath = options.QcFilePath,
            ClosureTemplatePath = options.ClosureTemplatePath,
            OutputPath = exportResult?.OutputPath,
            ReportPath = savedReportPath,
            Succeeded = exportResult?.Validation.IsValid ?? true,
            SummaryText = $"{containerId}: {summary.Items.Count} materiales, {summary.DiscrepantMaterialsCount} discrepancias, {summary.RevueltoPairsCount} revueltos.",
            DiscrepantMaterialsCount = summary.DiscrepantMaterialsCount,
            RevueltoPairsCount = summary.RevueltoPairsCount,
            ErrorMessage = exportResult?.Validation.Errors.FirstOrDefault()
        };

        return OperationHistoryStore.Append(options.HistoryFilePath, entry, options.SafetyGuard);
    }
}
