using System.Globalization;
using Acr.Core.Models;
using Acr.Infrastructure.Safety;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Infrastructure.Excel;

/// <summary>
/// Complete result returned by <see cref="ClosureExporter.Export"/> including export metrics,
/// 3-way formula audit validation, and atomic transaction metadata.
/// </summary>
/// <param name="TemplatePath">Canonicalized source closure template path.</param>
/// <param name="OutputPath">Canonicalized destination closure workbook path.</param>
/// <param name="MatchedRowsCount">Number of <c>CIERRE RECEP</c> rows matched and populated.</param>
/// <param name="RepairedFormulaCount">Number of missing or broken formulas repaired in <c>H</c>, <c>AU</c>, or <c>AZ</c>.</param>
/// <param name="BultosSyncedRowsCount">Number of <c>Bultos</c> rows synchronized with <c>CIERRE RECEP</c>.</param>
/// <param name="BultosTrailingClearedCount">Number of trailing <c>Bultos</c> rows whose stale literals/caches were cleared.</param>
/// <param name="Validation">Full 3-way formula and package-preservation audit result.</param>
/// <param name="Transaction">Atomic file transaction metadata.</param>
public sealed record ClosureExportResult(
    string TemplatePath,
    string OutputPath,
    int MatchedRowsCount,
    int RepairedFormulaCount,
    int BultosSyncedRowsCount,
    int BultosTrailingClearedCount,
    PrePrintValidationResult Validation,
    AtomicTransactionResult Transaction);

/// <summary>
/// Macro-safe direct OpenXML exporter for corporate receiving closure (<c>.xlsm</c>) workbooks.
/// Preserves <c>xl/vbaProject.bin</c>, <c>styles.xml</c>, <c>definedNames</c>, workbook relationships,
/// and Excel Tables/PivotTables while populating <c>CIERRE RECEP</c> and <c>Bultos</c>.
/// </summary>
public static class ClosureExporter
{
    private static readonly (int BultosCol, string CierreCol)[] BultosCrossSheetMap =
    [
        (1, "C"),
        (2, "AS"),
        (3, "AT"),
        (4, "E"),
        (5, "H"),
        (6, "BB"),
        (7, "BC"),
        (8, "AD"),
        (9, "AY")
    ];

    private sealed class ExportMetrics
    {
        public int MatchedRowsCount { get; set; }
        public int RepairedFormulaCount { get; set; }
        public int BultosSyncedRowsCount { get; set; }
        public int BultosTrailingClearedCount { get; set; }
        public List<string> Warnings { get; } = [];
        public PrePrintValidationResult? Validation { get; set; }
    }

    /// <summary>
    /// Exports a reconciled closure workbook from <paramref name="templatePath"/> to <paramref name="outputPath"/>
    /// inside an <see cref="AtomicFileTransaction"/>, enforcing <see cref="OperationalSafetyGuard"/> and
    /// pre-commit <see cref="FormulaIntegrityAuditor"/> verification.
    /// </summary>
    public static ClosureExportResult Export(
        string templatePath,
        string outputPath,
        ComparisonSummary summary,
        QcInspectionReport? qcReport = null,
        OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(summary);

        OperationalSafetyGuard effectiveGuard = guard ?? CreateDefaultPermissiveGuard();
        string resolvedTemplate = effectiveGuard.EnsureCanRead(templatePath);
        EnsureValidClosureTemplate(resolvedTemplate, effectiveGuard);

        var metrics = new ExportMetrics();
        AtomicTransactionResult txResult = AtomicFileTransaction.Execute(
            outputPath,
            effectiveGuard,
            OperationalActionKind.WriteToNewExportPath,
            tempPath => WriteClosureToTempFile(resolvedTemplate, tempPath, summary, qcReport, metrics),
            tempPath => ValidateTempClosure(tempPath, resolvedTemplate, summary, effectiveGuard, metrics),
            committedPath => WorkbookSignatureScanner.Inspect(committedPath, effectiveGuard).Kind
                is WorkbookKind.ReceivingClosure or WorkbookKind.HybridQcAndClosure,
            OperationalTargetKind.ClosureExportOutput);

        PrePrintValidationResult validation = metrics.Validation
            ?? FormulaIntegrityAuditor.ValidateClosureDocument(txResult.TargetPath, resolvedTemplate, summary, effectiveGuard);

        return new ClosureExportResult(
            resolvedTemplate,
            txResult.TargetPath,
            metrics.MatchedRowsCount,
            metrics.RepairedFormulaCount,
            metrics.BultosSyncedRowsCount,
            metrics.BultosTrailingClearedCount,
            validation,
            txResult);
    }

    private static OperationalSafetyGuard CreateDefaultPermissiveGuard()
    {
        string sentinelRoot = Path.Combine(Path.GetTempPath(), $"AcrNoProd_{Guid.NewGuid():N}");
        return new OperationalSafetyGuard(
            OperationalSafetyMode.Production,
            sentinelRoot,
            Path.Combine(sentinelRoot, "NoDailyLog.xlsx"));
    }

    private static void EnsureValidClosureTemplate(string resolvedTemplate, OperationalSafetyGuard guard)
    {
        if (!File.Exists(resolvedTemplate))
            throw new FileNotFoundException($"Closure template not found: '{resolvedTemplate}'.", resolvedTemplate);

        WorkbookSignature sig = WorkbookSignatureScanner.Inspect(resolvedTemplate, guard);
        if (!sig.SheetNames.Any(s => s.Contains("CIERRE RECEP", StringComparison.OrdinalIgnoreCase)))
            throw new InvalidDataException($"Template '{resolvedTemplate}' does not contain required 'CIERRE RECEP' worksheet.");
    }

    private static void WriteClosureToTempFile(
        string resolvedTemplate,
        string tempPath,
        ComparisonSummary summary,
        QcInspectionReport? qcReport,
        ExportMetrics metrics)
    {
        File.Copy(resolvedTemplate, tempPath, overwrite: true);
        OpenXmlRelsSanitizer.SanitizePackageFile(tempPath);
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(tempPath, isEditable: true);
        WorkbookPart wbPart = doc.WorkbookPart ?? throw new InvalidDataException($"Missing WorkbookPart in '{tempPath}'.");

        WorksheetPart cierrePart = OpenXmlCellHelper.FindWorksheetPart(wbPart, "CIERRE RECEP", allowContains: true)
            ?? throw new InvalidDataException("Missing 'CIERRE RECEP' worksheet.");
        Dictionary<string, int> sstCache = OpenXmlCellHelper.BuildSharedStringCache(wbPart);
        List<string> sharedStrings = OpenXmlCellHelper.ReadSharedStrings(wbPart);

        (List<int> matchedRows, HashSet<int> activeRows) =
            PopulateCierreRecepSheet(wbPart, cierrePart, summary, qcReport, sharedStrings, sstCache, metrics);

        UpdateBultosSheetIfPresent(wbPart, matchedRows, activeRows, metrics);
        OpenXmlCellHelper.PrepareWorkbookForFullRecalc(wbPart);

        cierrePart.Worksheet?.Save();
        wbPart.SharedStringTablePart?.SharedStringTable?.Save();
        wbPart.Workbook?.Save();
    }

    private static (List<int> MatchedRows, HashSet<int> ActiveRows) PopulateCierreRecepSheet(
        WorkbookPart wbPart,
        WorksheetPart cierrePart,
        ComparisonSummary summary,
        QcInspectionReport? qcReport,
        IReadOnlyList<string> sharedStrings,
        Dictionary<string, int> sstCache,
        ExportMetrics metrics)
    {
        Worksheet ws = cierrePart.Worksheet ?? throw new InvalidDataException("'CIERRE RECEP' worksheet root is missing.");
        SheetData sheetData = ws.GetFirstChild<SheetData>() ?? ws.AppendChild(new SheetData());

        WriteHeaderSummaryRows(sheetData, summary);

        Dictionary<string, DiscrepancyItem> itemMap = BuildItemMap(summary);
        Dictionary<string, int> curveMap = BuildQcCurveMap(qcReport);
        var sharedMasters = new Dictionary<uint, string>();
        var matchedRows = new List<int>();
        var activeRows = new HashSet<int>();

        foreach (Row row in sheetData.Elements<Row>())
        {
            int rowIdx = checked((int)(row.RowIndex?.Value ?? 0));
            if (rowIdx < 13)
                continue;

            ProcessCierreRow(
                wbPart, sheetData, row, rowIdx, itemMap, curveMap, sharedStrings, sstCache,
                sharedMasters, matchedRows, activeRows, metrics);
        }

        return (matchedRows, activeRows);
    }

    private static void WriteHeaderSummaryRows(SheetData sheetData, ComparisonSummary summary)
    {
        Row row5 = OpenXmlCellHelper.FindOrCreateRow(sheetData, 5);
        OpenXmlCellHelper.SetIntegerCell(row5, 1, 5, summary.TotalDeclaredCartons);

        Row row6 = OpenXmlCellHelper.FindOrCreateRow(sheetData, 6);
        OpenXmlCellHelper.SetIntegerCell(row6, 1, 6, summary.TotalPhysicalCartons);

        Row row7 = OpenXmlCellHelper.FindOrCreateRow(sheetData, 7);
        OpenXmlCellHelper.SetFormulaCell(row7, 1, 7, "+B6-B5");
    }

    private static Dictionary<string, DiscrepancyItem> BuildItemMap(ComparisonSummary summary)
    {
        var map = new Dictionary<string, DiscrepancyItem>(StringComparer.OrdinalIgnoreCase);
        foreach (DiscrepancyItem item in summary.Items)
        {
            string norm = OpenXmlCellHelper.NormalizeMaterialCode(item.MaterialCode);
            if (norm.Length > 0)
                map.TryAdd(norm, item);
        }
        return map;
    }

    private static Dictionary<string, int> BuildQcCurveMap(QcInspectionReport? qcReport)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        if (qcReport is null)
            return map;

        foreach (QcInspectionItem item in qcReport.Items)
        {
            if (item.IsChildSku)
                continue;
            string norm = OpenXmlCellHelper.NormalizeMaterialCode(item.MaterialCode);
            int curve = item.CurveSystem is > 0 ? item.CurveSystem.Value : (item.CurvePhysical ?? 0);
            if (norm.Length > 0 && curve > 0)
                map.TryAdd(norm, curve);
        }
        return map;
    }

    private static void ProcessCierreRow(
        WorkbookPart wbPart,
        SheetData sheetData,
        Row row,
        int rowIdx,
        Dictionary<string, DiscrepancyItem> itemMap,
        Dictionary<string, int> curveMap,
        IReadOnlyList<string> sharedStrings,
        Dictionary<string, int> sstCache,
        Dictionary<uint, string> sharedMasters,
        List<int> matchedRows,
        HashSet<int> activeRows,
        ExportMetrics metrics)
    {
        TrackSharedMastersInRow(row, sharedMasters);
        ClearSumFormulaCachesInRow(row);

        Cell? matCell = OpenXmlCellHelper.TryFindCell(row, 4);
        string matCode = OpenXmlCellHelper.NormalizeMaterialCode(OpenXmlCellHelper.DecodeCellText(matCell, sharedStrings));
        if (matCode.Length == 0 ||
            matCode.Equals("MATERIAL", StringComparison.OrdinalIgnoreCase) ||
            matCode.Contains("TOTAL", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        activeRows.Add(rowIdx);
        if (!itemMap.TryGetValue(matCode, out DiscrepancyItem? item))
            return;

        matchedRows.Add(rowIdx);
        metrics.MatchedRowsCount++;
        PopulateMatchedCierreRow(wbPart, sheetData, row, rowIdx, item, curveMap, sstCache, sharedMasters, metrics);
    }

    private static void TrackSharedMastersInRow(Row row, Dictionary<uint, string> sharedMasters)
    {
        foreach (Cell cell in row.Elements<Cell>())
        {
            CellFormula? f = cell.CellFormula;
            if (f is not null &&
                f.FormulaType?.Value == CellFormulaValues.Shared &&
                f.SharedIndex?.Value is uint si &&
                !string.IsNullOrWhiteSpace(f.Text))
            {
                sharedMasters[si] = f.Text;
            }
        }
    }

    private static void ClearSumFormulaCachesInRow(Row row)
    {
        foreach (Cell cell in row.Elements<Cell>())
        {
            string? text = cell.CellFormula?.Text;
            if (!string.IsNullOrWhiteSpace(text) &&
                (text.Contains("SUM(", StringComparison.OrdinalIgnoreCase) ||
                 text.Contains("SUBTOTAL(", StringComparison.OrdinalIgnoreCase)))
            {
                OpenXmlCellHelper.ClearFormulaCacheOnly(cell);
            }
        }
    }

    private static void PopulateMatchedCierreRow(
        WorkbookPart wbPart,
        SheetData sheetData,
        Row row,
        int rowIdx,
        DiscrepancyItem item,
        Dictionary<string, int> curveMap,
        Dictionary<string, int> sstCache,
        Dictionary<uint, string> sharedMasters,
        ExportMetrics metrics)
    {
        string r = rowIdx.ToString(CultureInfo.InvariantCulture);
        EnsureOrRepairFormula(sheetData, row, 7, rowIdx, $"+G{r}-J{r}-N{r}-O{r}+L{r}+P{r}+Q{r}", sharedMasters, metrics);

        WriteRevueltoColumns(wbPart, row, rowIdx, item, sstCache);
        WritePhysicalAndDeclaredColumns(wbPart, row, rowIdx, item, curveMap, sstCache);

        EnsureOrRepairFormula(sheetData, row, 46, rowIdx, $"+AT{r}-H{r}", sharedMasters, metrics);
        EnsureOrRepairFormula(sheetData, row, 51, rowIdx, $"+AX{r}-H{r}", sharedMasters, metrics);
    }

    private static void EnsureOrRepairFormula(
        SheetData sheetData,
        Row row,
        int colIdx,
        int rowIdx,
        string fallbackFormula,
        Dictionary<uint, string> sharedMasters,
        ExportMetrics metrics)
    {
        Cell? existing = OpenXmlCellHelper.TryFindCell(row, colIdx);
        if (existing?.CellFormula is not null && IsValidExistingFormula(existing.CellFormula, sharedMasters))
        {
            OpenXmlCellHelper.ClearFormulaCacheOnly(existing);
            return;
        }

        if (existing?.CellFormula is not null)
            MaterializeSharedFollowersIfMaster(sheetData, existing, rowIdx, colIdx);

        OpenXmlCellHelper.SetFormulaCell(row, colIdx, rowIdx, fallbackFormula);
        metrics.RepairedFormulaCount++;
    }

    private static void MaterializeSharedFollowersIfMaster(
        SheetData sheetData,
        Cell masterCell,
        int masterRowIdx,
        int masterColIdx)
    {
        CellFormula? f = masterCell.CellFormula;
        if (f is null ||
            f.FormulaType?.Value != CellFormulaValues.Shared ||
            f.SharedIndex?.Value is not uint si ||
            string.IsNullOrWhiteSpace(f.Text))
        {
            return;
        }

        int currentRow = 0;
        foreach (Row candidateRow in sheetData.Elements<Row>())
        {
            int targetRowIdx = checked((int)(candidateRow.RowIndex?.Value ?? (uint)(currentRow + 1)));
            currentRow = targetRowIdx;
            MaterializeFollowersInRow(candidateRow, masterCell, si, f.Text, masterRowIdx, masterColIdx, targetRowIdx);
        }
    }

    private static void MaterializeFollowersInRow(
        Row candidateRow,
        Cell masterCell,
        uint sharedIndex,
        string masterFormula,
        int masterRowIdx,
        int masterColIdx,
        int targetRowIdx)
    {
        int fallbackCol = 0;
        foreach (Cell candidateCell in candidateRow.Elements<Cell>())
        {
            int targetColIdx = QcReader.ParseColumnIndex(candidateCell.CellReference?.Value, ref fallbackCol);
            if (ReferenceEquals(candidateCell, masterCell) ||
                candidateCell.CellFormula?.FormulaType?.Value != CellFormulaValues.Shared ||
                candidateCell.CellFormula?.SharedIndex?.Value != sharedIndex)
            {
                continue;
            }

            string shifted = OpenXmlCellHelper.ShiftFormulaReferences(
                masterFormula,
                targetColIdx - masterColIdx,
                targetRowIdx - masterRowIdx);
            candidateCell.CellFormula = new CellFormula(shifted);
        }
    }

    private static bool IsValidExistingFormula(CellFormula formula, Dictionary<uint, string> sharedMasters)
    {
        string? text = formula.Text;
        if (string.IsNullOrWhiteSpace(text) &&
            formula.FormulaType?.Value == CellFormulaValues.Shared &&
            formula.SharedIndex?.Value is uint si)
        {
            sharedMasters.TryGetValue(si, out text);
        }

        if (string.IsNullOrWhiteSpace(text))
            return false;

        return !text.Contains("#REF!", StringComparison.OrdinalIgnoreCase) &&
               !text.Contains("#NAME?", StringComparison.OrdinalIgnoreCase);
    }

    private static void WriteRevueltoColumns(
        WorkbookPart wbPart,
        Row row,
        int rowIdx,
        DiscrepancyItem item,
        Dictionary<string, int> sstCache)
    {
        if (item.DiscrepancyType != DiscrepancyType.RevueltoCandidate)
            return;

        string partner = item.RevueltoPartner ?? "";
        if (item.DifferenceUnits < 0)
        {
            OpenXmlCellHelper.SetIntegerCell(row, 9, rowIdx, Math.Abs(item.DifferenceUnits));
            OpenXmlCellHelper.SetSharedStringCell(row, 10, rowIdx, partner, wbPart, sstCache);
        }
        else if (item.DifferenceUnits > 0)
        {
            OpenXmlCellHelper.SetIntegerCell(row, 11, rowIdx, item.DifferenceUnits);
            OpenXmlCellHelper.SetSharedStringCell(row, 12, rowIdx, partner, wbPart, sstCache);
        }
    }

    private static void WritePhysicalAndDeclaredColumns(
        WorkbookPart wbPart,
        Row row,
        int rowIdx,
        DiscrepancyItem item,
        Dictionary<string, int> curveMap,
        Dictionary<string, int> sstCache)
    {
        int physUnits = item.PhysicalUnits == 0 ? 0 : item.PhysicalUnits;
        int physCartons = item.PhysicalUnits == 0 ? 0 : item.PhysicalCartons;
        OpenXmlCellHelper.SetIntegerCell(row, 44, rowIdx, physCartons);
        OpenXmlCellHelper.SetIntegerCell(row, 45, rowIdx, physUnits);

        if (item.PhysicalUnits == 0 && item.DeclaredUnits == 0)
        {
            OpenXmlCellHelper.SetIntegerCell(row, 48, rowIdx, 0);
            OpenXmlCellHelper.SetIntegerCell(row, 49, rowIdx, 0);
            OpenXmlCellHelper.SetSharedStringCell(row, 50, rowIdx, "0", wbPart, sstCache);
            return;
        }

        OpenXmlCellHelper.SetIntegerCell(row, 48, rowIdx, item.DeclaredCartons);
        OpenXmlCellHelper.SetIntegerCell(row, 49, rowIdx, item.DeclaredUnits);

        string normCode = OpenXmlCellHelper.NormalizeMaterialCode(item.MaterialCode);
        string empaque = curveMap.TryGetValue(normCode, out int curve) && curve > 1
            ? curve.ToString(CultureInfo.InvariantCulture)
            : "1";
        OpenXmlCellHelper.SetSharedStringCell(row, 50, rowIdx, empaque, wbPart, sstCache);
    }

    private static void UpdateBultosSheetIfPresent(
        WorkbookPart wbPart,
        List<int> matchedCierreRows,
        HashSet<int> activeCierreRows,
        ExportMetrics metrics)
    {
        WorksheetPart? bultosPart = OpenXmlCellHelper.FindWorksheetPart(wbPart, "Bultos", allowContains: false);
        Worksheet? bultosWs = bultosPart?.Worksheet;
        SheetData? sheetData = bultosWs?.GetFirstChild<SheetData>();
        if (bultosWs is null || sheetData is null)
            return;

        int totalsRow = FindBultosTotalsRow(sheetData);
        int lastMatchedBultosRow = SyncMatchedBultosRows(sheetData, matchedCierreRows, totalsRow, metrics);
        ClearTrailingBultosRows(sheetData, lastMatchedBultosRow, totalsRow, activeCierreRows, metrics);
        if (totalsRow > 0)
            ClearSumFormulaCachesInRow(OpenXmlCellHelper.FindOrCreateRow(sheetData, totalsRow));
        bultosWs.Save();
    }

    private static int FindBultosTotalsRow(SheetData sheetData)
    {
        foreach (Row row in sheetData.Elements<Row>())
        {
            int rowIdx = checked((int)(row.RowIndex?.Value ?? 0));
            if (rowIdx < 3)
                continue;

            foreach (Cell cell in row.Elements<Cell>())
            {
                string? formula = cell.CellFormula?.Text;
                if (!string.IsNullOrWhiteSpace(formula) &&
                    (formula.Contains("SUM(", StringComparison.OrdinalIgnoreCase) ||
                     formula.Contains("SUBTOTAL(", StringComparison.OrdinalIgnoreCase)))
                {
                    return rowIdx;
                }
            }
        }
        return -1;
    }

    private static int SyncMatchedBultosRows(
        SheetData sheetData,
        List<int> matchedCierreRows,
        int totalsRow,
        ExportMetrics metrics)
    {
        int lastBultosRow = 2;
        foreach (int cierreRow in matchedCierreRows)
        {
            int bultosRow = cierreRow - 10;
            if (bultosRow < 3)
            {
                metrics.Warnings.Add(
                    $"BULTOS_SYNC_OUT_OF_RANGE: CIERRE RECEP row {cierreRow} maps to Bultos row {bultosRow} (< 3); row sync skipped.");
                continue;
            }
            if (totalsRow > 0 && bultosRow >= totalsRow)
            {
                metrics.Warnings.Add(
                    $"BULTOS_SYNC_TRUNCATED: CIERRE RECEP row {cierreRow} maps to Bultos row {bultosRow}, which is at or past totals row {totalsRow}; row sync skipped.");
                continue;
            }

            Row row = OpenXmlCellHelper.FindOrCreateRow(sheetData, bultosRow);
            SyncSingleBultosRow(row, bultosRow, cierreRow);
            metrics.BultosSyncedRowsCount++;
            if (bultosRow > lastBultosRow)
                lastBultosRow = bultosRow;
        }
        return lastBultosRow;
    }

    private static void SyncSingleBultosRow(Row row, int bultosRow, int cierreRow)
    {
        string rStr = cierreRow.ToString(CultureInfo.InvariantCulture);
        foreach ((int bultosCol, string cierreCol) in BultosCrossSheetMap)
        {
            string expectedFormula = $"'CIERRE RECEP'!{cierreCol}{rStr}";
            Cell? existing = OpenXmlCellHelper.TryFindCell(row, bultosCol);
            if (existing?.CellFormula is not null &&
                !string.IsNullOrWhiteSpace(existing.CellFormula.Text) &&
                existing.CellFormula.Text.Contains("CIERRE RECEP", StringComparison.OrdinalIgnoreCase) &&
                !existing.CellFormula.Text.Contains("#REF!", StringComparison.OrdinalIgnoreCase))
            {
                OpenXmlCellHelper.ClearFormulaCacheOnly(existing);
            }
            else
            {
                OpenXmlCellHelper.SetFormulaCell(row, bultosCol, bultosRow, expectedFormula);
            }
        }
    }

    private static void ClearTrailingBultosRows(
        SheetData sheetData,
        int lastMatchedBultosRow,
        int totalsRow,
        HashSet<int> activeCierreRows,
        ExportMetrics metrics)
    {
        foreach (Row row in sheetData.Elements<Row>())
        {
            int bultosRow = checked((int)(row.RowIndex?.Value ?? 0));
            if (bultosRow < 3 || bultosRow <= lastMatchedBultosRow)
                continue;
            if (totalsRow > 0 && bultosRow >= totalsRow)
                continue;
            if (activeCierreRows.Contains(bultosRow + 10))
                continue;

            if (ClearTrailingRowCells(row))
                metrics.BultosTrailingClearedCount++;
        }
    }

    private static bool ClearTrailingRowCells(Row row)
    {
        bool modified = false;
        for (int colIdx = 1; colIdx <= 9; colIdx++)
        {
            Cell? cell = OpenXmlCellHelper.TryFindCell(row, colIdx);
            if (cell is null)
                continue;

            if (cell.CellFormula is null)
            {
                if (cell.CellValue is not null || cell.InlineString is not null || cell.DataType is not null)
                {
                    cell.CellValue = null;
                    cell.InlineString = null;
                    cell.DataType = null;
                    modified = true;
                }
            }
            else if (cell.CellValue is not null)
            {
                OpenXmlCellHelper.ClearFormulaCacheOnly(cell);
                modified = true;
            }
        }
        return modified;
    }

    private static bool ValidateTempClosure(
        string tempPath,
        string resolvedTemplate,
        ComparisonSummary summary,
        OperationalSafetyGuard guard,
        ExportMetrics metrics)
    {
        PrePrintValidationResult validation =
            FormulaIntegrityAuditor.ValidateClosureDocument(tempPath, resolvedTemplate, summary, guard);
        if (metrics.Warnings.Count > 0)
        {
            validation = validation with
            {
                Warnings = validation.Warnings.Concat(metrics.Warnings).ToList()
            };
        }
        metrics.Validation = validation;
        return validation.VbaProjectHashPreserved &&
               validation.StylesPreserved &&
               validation.DefinedNamesPreserved &&
               validation.RelationshipsPreserved &&
               validation.Errors.Count == 0 &&
               validation.ExporterIntroducedErrors.Count == 0;
    }
}
