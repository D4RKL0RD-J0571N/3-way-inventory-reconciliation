using System.Text.RegularExpressions;
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

    /// <summary>Indica si se deben omitir materiales en cero (sin movimiento) de las tiras y hoja Bultos al pegar en Excel.</summary>
    public bool ExcludeZeroTotalFromStripsAndBultos { get; init; } = true;
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

    /// <summary>Proyección oficial para la hoja 'Bultos' y análisis de empaque.</summary>
    public IReadOnlyList<BultosPresentationItem> BultosItems { get; init; } = Array.Empty<BultosPresentationItem>();

    /// <summary>Tira 1 para pegar en Excel columnas AS:AT (Bultos y Unidades de Muelle).</summary>
    public string Tira1AsAtTsv { get; init; } = "";

    /// <summary>Tira 2 para pegar en Excel columnas AW:AY (Packing List y Empaque).</summary>
    public string Tira2AwAyTsv { get; init; } = "";
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

        string qcPath = options.QcFilePath;
        string closurePath = options.ClosureTemplatePath;

        // Auto-detect swapped workbooks or hybrid workbooks using internal sheet signatures
        try
        {
            WorkbookSignature sigQc = WorkbookSignatureScanner.Inspect(qcPath, options.SafetyGuard);
            WorkbookSignature sigClosure = WorkbookSignatureScanner.Inspect(closurePath, options.SafetyGuard);

            if (sigQc.Kind == WorkbookKind.ReceivingClosure && sigClosure.Kind == WorkbookKind.QcInspection)
            {
                qcPath = options.ClosureTemplatePath;
                closurePath = options.QcFilePath;
            }
            else if (sigQc.Kind == WorkbookKind.HybridQcAndClosure && !File.Exists(closurePath))
            {
                closurePath = qcPath;
            }
            else if (sigClosure.Kind == WorkbookKind.HybridQcAndClosure && !File.Exists(qcPath))
            {
                qcPath = closurePath;
            }
        }
        catch
        {
            // Fallback to configured paths on any inspection anomaly
        }

        QcInspectionReport qcReport = QcReader.Read(qcPath, options.SafetyGuard);
        ClosureTemplateData templateData = ClosureTemplateReader.ReadDeclaredItems(closurePath, options.SafetyGuard);
        IReadOnlyList<PhysicalItem> physicalItems = MapPhysicalItems(qcReport);
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(templateData.DeclaredItems, physicalItems);

        // Enrich discrepancy items with BOM proportional child SKU distribution, canonical packaging, and curves
        var declMap = templateData.DeclaredItems.ToDictionary(d => OpenXmlCellHelper.NormalizeMaterialCode(d.MaterialCode), StringComparer.OrdinalIgnoreCase);
        var qcMap = qcReport.Items.Where(q => !q.IsChildSku).ToDictionary(q => OpenXmlCellHelper.NormalizeMaterialCode(q.MaterialCode), StringComparer.OrdinalIgnoreCase);
        var enrichedItems = new List<DiscrepancyItem>(summary.Items.Count);

        foreach (var item in summary.Items)
        {
            declMap.TryGetValue(item.MaterialCode, out DeclaredItem? decl);
            qcMap.TryGetValue(item.MaterialCode, out QcInspectionItem? qc);

            int sourceRow = decl?.SourceRow ?? qc?.RowIndex ?? 0;
            decimal unitCost = decl?.UnitCost ?? 0m;
            int curve = templateData.RowCurves.TryGetValue(sourceRow, out int rc) && rc > 0
                ? rc
                : (qc?.CurveSystem is > 0 ? qc.CurveSystem.Value : (qc?.CurvePhysical ?? 12));

            int fullCartons = qc?.FullCartons ?? (qc?.CalculationResult?.FullCartonsUsed ?? (curve > 0 ? item.PhysicalUnits / curve : item.PhysicalCartons));
            int unitsPerFull = qc?.UnitsPerFullCarton ?? (qc?.CalculationResult?.UnitsPerFullCartonUsed ?? curve);
            int compCartons = qc?.ComplementCartons ?? (qc?.CalculationResult?.ComplementCartonsUsed ?? 0);
            int unitsInComp = qc?.UnitsInComplement ?? (qc?.CalculationResult?.ComplementUnits ?? 0);

            string canonicalPack = QuantityEngine.FormatPackagingExpression(
                fullCartons,
                unitsPerFull,
                compCartons,
                unitsInComp,
                curve,
                totalUnits: item.PhysicalUnits != 0 ? item.PhysicalUnits : item.DeclaredUnits,
                totalCartons: item.PhysicalCartons != 0 ? item.PhysicalCartons : item.DeclaredCartons);

            IReadOnlyList<ChildSkuItem> children = Array.Empty<ChildSkuItem>();
            if (qcReport.BomsByParent.TryGetValue(item.MaterialCode, out var bomDef) && bomDef.Count > 0)
            {
                int ped41Prk = curve > 0 ? item.DeclaredUnits / curve : 0;
                children = BomAllocationService.AllocateChildSkus(
                    bomDef,
                    item.DeclaredUnits,
                    item.PhysicalUnits,
                    ped41Prk,
                    curve,
                    item.Description,
                    unitCost,
                    qcReport.ChildDescriptions);
            }

            enrichedItems.Add(item with
            {
                SourceRow = sourceRow,
                UnitCost = unitCost,
                Curve = curve,
                ContainerPart = qcReport.DetectedParts.Count > 0 ? qcReport.DetectedParts[0] : null,
                CanonicalPackaging = canonicalPack,
                Children = children
            });
        }

        summary = summary with { Items = enrichedItems };
        string containerId = ResolveContainerId(templateData, qcReport);

        ClosureExportResult? exportResult = ExportIfRequested(options, summary, qcReport);
        List<string> warnings = CollectWarnings(templateData, exportResult);
        warnings.AddRange(qcReport.Warnings);

        (IReadOnlyList<BultosPresentationItem> bultosItems, string tira1Tsv, string tira2Tsv) =
            BuildBultosAndPasteStrips(templateData, qcReport, summary, options.ExcludeZeroTotalFromStripsAndBultos);
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
            Warnings = warnings,
            BultosItems = bultosItems,
            Tira1AsAtTsv = tira1Tsv,
            Tira2AwAyTsv = tira2Tsv
        };
    }

    private static (IReadOnlyList<BultosPresentationItem> Items, string Tira1Tsv, string Tira2Tsv) BuildBultosAndPasteStrips(
        ClosureTemplateData templateData,
        QcInspectionReport qcReport,
        ComparisonSummary summary,
        bool excludeZeroTotal = true)
    {
        var bultosList = new List<BultosPresentationItem>();
        var tira1Sb = new System.Text.StringBuilder();
        var tira2Sb = new System.Text.StringBuilder();

        var qcMap = new Dictionary<string, QcInspectionItem>(StringComparer.OrdinalIgnoreCase);
        foreach (var q in qcReport.Items)
        {
            if (!q.IsChildSku)
            {
                string norm = OpenXmlCellHelper.NormalizeMaterialCode(q.MaterialCode);
                if (norm.Length > 0)
                    qcMap.TryAdd(norm, q);
            }
        }

        var discMap = new Dictionary<string, DiscrepancyItem>(StringComparer.OrdinalIgnoreCase);
        foreach (var d in summary.Items)
        {
            string norm = OpenXmlCellHelper.NormalizeMaterialCode(d.MaterialCode);
            if (norm.Length > 0)
                discMap.TryAdd(norm, d);
        }

        foreach (DeclaredItem decl in templateData.DeclaredItems)
        {
            string norm = OpenXmlCellHelper.NormalizeMaterialCode(decl.MaterialCode);
            qcMap.TryGetValue(norm, out QcInspectionItem? qcItem);
            discMap.TryGetValue(norm, out DiscrepancyItem? discItem);

            int physUnits = discItem?.PhysicalUnits ?? (qcItem?.CalculationResult?.CalculatedUnits ?? 0);
            int physCartons = discItem?.PhysicalCartons ?? (qcItem?.CalculationResult?.CalculatedCartons ?? 0);
            int declCartons = decl.DeclaredCartons;
            int declUnits = decl.DeclaredUnits;

            bool isZeroTotal = physUnits == 0 && declUnits == 0 && physCartons == 0 && declCartons == 0;
            if (isZeroTotal && excludeZeroTotal)
            {
                // Al pegar en Excel se ocultan o eliminan materiales que tienen 0: no se incluyen en las tiras ni en bultos
                continue;
            }

            int curve = templateData.RowCurves.TryGetValue(decl.SourceRow ?? 0, out int rc) && rc > 0
                ? rc
                : (qcItem?.CurveSystem is > 0 ? qcItem.CurveSystem.Value : (qcItem?.CurvePhysical ?? 1));

            int fullCartons = qcItem?.FullCartons ?? (qcItem?.CalculationResult?.FullCartonsUsed ?? (curve > 0 ? physUnits / curve : physCartons));
            int unitsPerFull = qcItem?.UnitsPerFullCarton ?? (qcItem?.CalculationResult?.UnitsPerFullCartonUsed ?? curve);
            int compCartons = qcItem?.ComplementCartons ?? (qcItem?.CalculationResult?.ComplementCartonsUsed ?? 0);
            int unitsInComp = qcItem?.UnitsInComplement ?? (qcItem?.CalculationResult?.ComplementUnits ?? 0);

            string empaque = QuantityEngine.FormatPackagingExpression(
                fullCartons,
                unitsPerFull,
                compCartons,
                unitsInComp,
                curve,
                totalUnits: physUnits != 0 ? physUnits : declUnits,
                totalCartons: physCartons != 0 ? physCartons : declCartons);

            var bultoItem = new BultosPresentationItem
            {
                CierreRow = decl.SourceRow ?? 0,
                Category = decl.Category ?? qcItem?.Category ?? "",
                MaterialCode = decl.MaterialCode,
                Description = decl.Description ?? qcItem?.Description ?? "",
                PhysicalCartons = physCartons,
                PhysicalUnits = physUnits,
                ReceivedUnits = physUnits,
                Curve = curve,
                Uom = decl.Uom ?? qcItem?.Uom ?? "UN",
                PackagingExpression = empaque,
                FullCartons = fullCartons,
                UnitsPerFullCarton = unitsPerFull,
                ComplementCartons = compCartons,
                UnitsInComplement = unitsInComp,
                DeclaredCartons = declCartons,
                DeclaredUnits = declUnits,
                Observation = qcItem?.Observation ?? ""
            };
            bultosList.Add(bultoItem);

            // Tira 1: AS:AT (Bultos Físicos \t Unidades Físicas)
            tira1Sb.AppendLine($"{physCartons}\t{physUnits}");

            // Tira 2: AW:AY (Bultos PL \t Unidades PL \t Empaque)
            tira2Sb.AppendLine($"{declCartons}\t{declUnits}\t{empaque}");
        }

        return (bultosList, tira1Sb.ToString(), tira2Sb.ToString());
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
        if (!string.IsNullOrWhiteSpace(templateData.ContainerId) && !IsPureNumeric(templateData.ContainerId))
            return templateData.ContainerId.Trim();
        if (!string.IsNullOrWhiteSpace(qcReport.ContainerId) && !IsPureNumeric(qcReport.ContainerId))
            return qcReport.ContainerId.Trim();

        // Check fallback from template filename
        string fileName = Path.GetFileName(templateData.FilePath);
        Match m = Regex.Match(fileName, @"\b(IMP)[\s\-_]*([0-9]+[A-Z0-9\-]*)(?=[_\s.\-]|$|\b)", RegexOptions.IgnoreCase);
        if (m.Success)
            return $"IMP-{m.Groups[2].Value.ToUpperInvariant()}";

        // Check fallback from QC filename
        string qcName = Path.GetFileName(qcReport.FilePath);
        Match mQc = Regex.Match(qcName, @"\b(IMP)[\s\-_]*([0-9]+[A-Z0-9\-]*)(?=[_\s.\-]|$|\b)", RegexOptions.IgnoreCase);
        if (mQc.Success)
            return $"IMP-{mQc.Groups[2].Value.ToUpperInvariant()}";

        if (!string.IsNullOrWhiteSpace(templateData.ContainerId))
            return templateData.ContainerId.Trim();
        if (!string.IsNullOrWhiteSpace(qcReport.ContainerId))
            return qcReport.ContainerId.Trim();

        return Path.GetFileNameWithoutExtension(templateData.FilePath);
    }

    private static bool IsPureNumeric(string? val) =>
        !string.IsNullOrWhiteSpace(val) && val.Trim().All(char.IsDigit);

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
