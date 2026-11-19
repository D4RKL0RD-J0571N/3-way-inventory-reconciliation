using System.Globalization;
using System.Text.RegularExpressions;
using Acr.Core.Models;
using Acr.Core.Services;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Safety;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Infrastructure.Receiving;

/// <summary>
/// Declared merchandise lines and template metadata extracted from <c>CIERRE RECEP</c>.
/// </summary>
/// <param name="FilePath">Canonicalized path of the closure template workbook.</param>
/// <param name="ContainerId">Container identifier discovered in header rows <c>1..9</c>, if present.</param>
/// <param name="Pedido41">Nationalization Pedido 41 discovered in header rows <c>1..9</c>, if present.</param>
/// <param name="DeclaredItems">Parsed declared items from rows <c>13+</c>.</param>
/// <param name="RowCurves">Prepack curve values keyed by 1-based worksheet row index.</param>
/// <param name="Warnings">Non-blocking extraction warnings (such as missing column <c>AW</c> cartons).</param>
public sealed record ClosureTemplateData(
    string FilePath,
    string? ContainerId,
    string? Pedido41,
    IReadOnlyList<DeclaredItem> DeclaredItems,
    IReadOnlyDictionary<int, int> RowCurves,
    IReadOnlyList<string> Warnings);

/// <summary>
/// Forward-only streaming OpenXML reader for corporate <c>CIERRE RECEP</c> templates,
/// extracting declared units from column <c>G</c> (fallback <c>BE * BB</c>, never <c>F</c>)
/// and declared cartons from column <c>AW</c> (never <c>BD</c>/<c>BE</c>).
/// </summary>
public static class ClosureTemplateReader
{
    private static readonly string[] ImportIdExcludeTokens =
    [
        "IMPORTADORA", "FALTANTE", "DAÑAD", "DANAD", "SOBRANT", "REVUELT",
        "TOTAL", "BULTO", "VALOR", "COSTO", "UNIDAD", "FECHA", "DIFERENCIA",
        "NCR", "RECLAMO", "AVERI", "MONEDA", "MONTO", "CANTIDAD", "ESTADO"
    ];

    private static readonly Dictionary<string, int> DefaultColumns = new(StringComparer.Ordinal)
    {
        ["po43"] = 0,
        ["category"] = 2,
        ["vendor_material"] = 3,
        ["material"] = 4,
        ["units_ped41"] = 6,
        ["description"] = 29,
        ["cost_ae"] = 30,
        ["cost_af"] = 31,
        ["cartons_aw"] = 48,
        ["curve_bb"] = 53,
        ["uom_bc"] = 54,
        ["qty_be"] = 56
    };

    /// <summary>
    /// Streams <c>CIERRE RECEP</c> rows <c>13+</c> (no 299-row cap, 20-empty-row early exit)
    /// and returns structured <see cref="DeclaredItem"/> records plus template metadata.
    /// </summary>
    public static ClosureTemplateData ReadDeclaredItems(
        string templatePath,
        OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(templatePath);
        string resolved = guard is not null ? guard.EnsureCanRead(templatePath) : Path.GetFullPath(templatePath);
        if (!File.Exists(resolved))
            throw new FileNotFoundException($"Closure template not found: '{resolved}'.", resolved);

        using WorkbookReadSession session = OpenXmlRelsSanitizer.OpenReadOnly(resolved);
        WorkbookPart wbPart = session.Document.WorkbookPart ?? throw new InvalidDataException($"Missing WorkbookPart in '{resolved}'.");
        WorksheetPart wsPart = OpenXmlCellHelper.FindWorksheetPart(wbPart, "CIERRE RECEP", allowContains: true)
            ?? throw new InvalidDataException($"Workbook '{resolved}' does not contain required 'CIERRE RECEP' worksheet.");

        List<string> sharedStrings = OpenXmlCellHelper.ReadSharedStrings(wbPart);

        string? containerId = null;
        string? pedido41 = null;
        WorksheetPart? dgPart = OpenXmlCellHelper.FindWorksheetPart(wbPart, "DATOS GENERALES", allowContains: true);
        if (dgPart is not null)
        {
            ExtractDatosGeneralesMetadata(dgPart, sharedStrings, ref containerId, ref pedido41);
        }

        return StreamCierreRecep(resolved, wsPart, sharedStrings, containerId, pedido41);
    }

    private static void ExtractDatosGeneralesMetadata(
        WorksheetPart dgPart,
        IReadOnlyList<string> sharedStrings,
        ref string? containerId,
        ref string? pedido41)
    {
        using OpenXmlReader reader = OpenXmlReader.Create(dgPart);
        while (reader.Read())
        {
            if (reader.ElementType != typeof(Row) || reader.LoadCurrentElement() is not Row rowEl)
                continue;

            int rowIdx = (int)(rowEl.RowIndex?.Value ?? 0);
            if (rowIdx > 35)
                break;

            Dictionary<int, string> cells = ExtractRowCells(rowEl, sharedStrings);
            ExtractHeaderMetadata(cells, ref containerId, ref pedido41);
        }
    }

    private static ClosureTemplateData StreamCierreRecep(
        string resolvedPath,
        WorksheetPart wsPart,
        IReadOnlyList<string> sharedStrings,
        string? initialContainerId,
        string? initialPedido41)
    {
        var colMap = new Dictionary<string, int>(DefaultColumns, StringComparer.Ordinal);
        var items = new List<DeclaredItem>();
        var rowCurves = new Dictionary<int, int>();
        var warnings = new List<string>();
        string? containerId = initialContainerId;
        string? pedido41 = initialPedido41;
        int lastRow = 0;
        int emptyRows = 0;

        using OpenXmlReader reader = OpenXmlReader.Create(wsPart);
        while (reader.Read())
        {
            if (reader.ElementType != typeof(Row) || reader.LoadCurrentElement() is not Row rowEl)
                continue;

            int rowIdx = checked((int)(rowEl.RowIndex?.Value ?? (uint)(lastRow + 1)));
            if (ShouldExitForRowGap(rowIdx, ref lastRow, ref emptyRows))
                break;

            Dictionary<int, string> cells = ExtractRowCells(rowEl, sharedStrings);
            if (ProcessStreamedRow(rowIdx, cells, colMap, ref containerId, ref pedido41, items, rowCurves, warnings))
                emptyRows = 0;
            else if (rowIdx >= 13 && ++emptyRows >= 20)
                break;
        }

        // Fallback to filename if containerId is null or suspiciously pure-numeric without IMP prefix (IMP-207 -> 288 fix)
        string fileName = Path.GetFileName(resolvedPath);
        Match m = Regex.Match(fileName, @"\b(IMP)[\s\-_]*([0-9]+[A-Z0-9\-]*)(?=[_\s.\-]|$|\b)", RegexOptions.IgnoreCase);
        if (m.Success)
        {
            string fileImp = $"IMP-{m.Groups[2].Value.ToUpperInvariant()}";
            if (string.IsNullOrWhiteSpace(containerId) || IsPureNumeric(containerId) || !containerId.StartsWith("IMP-", StringComparison.OrdinalIgnoreCase))
            {
                containerId = fileImp;
            }
        }
        else if (IsPureNumeric(containerId) && containerId != "0")
        {
            containerId = $"IMP-{containerId}";
        }

        return new ClosureTemplateData(resolvedPath, containerId, pedido41, items, rowCurves, warnings);
    }

    private static bool ShouldExitForRowGap(int rowIdx, ref int lastRow, ref int emptyRows)
    {
        int gap = rowIdx - lastRow - 1;
        lastRow = rowIdx;
        if (rowIdx >= 13 && gap > 0)
        {
            emptyRows += gap;
            return emptyRows >= 20;
        }
        return false;
    }

    private static bool ProcessStreamedRow(
        int rowIdx,
        Dictionary<int, string> cells,
        Dictionary<string, int> colMap,
        ref string? containerId,
        ref string? pedido41,
        List<DeclaredItem> items,
        Dictionary<int, int> rowCurves,
        List<string> warnings)
    {
        if (cells.Count == 0)
            return false;

        if (rowIdx < 12)
        {
            ExtractHeaderMetadata(cells, ref containerId, ref pedido41);
            return true;
        }

        if (rowIdx == 12)
        {
            RefineColumnMap(cells, colMap);
            return true;
        }

        DeclaredItem? item = ParseDeclaredRow(rowIdx, cells, colMap, rowCurves, warnings);
        if (item is not null)
            items.Add(item);

        return item is not null || cells.Values.Any(v => !string.IsNullOrWhiteSpace(v));
    }

    private static Dictionary<int, string> ExtractRowCells(Row rowEl, IReadOnlyList<string> sharedStrings)
    {
        var cells = new Dictionary<int, string>();
        int fallbackCol = 0;
        foreach (Cell cell in rowEl.Elements<Cell>())
        {
            int colIdx = QcReader.ParseColumnIndex(cell.CellReference?.Value, ref fallbackCol);
            string text = OpenXmlCellHelper.DecodeCellText(cell, sharedStrings);
            if (text.Length > 0)
                cells[colIdx] = text;
        }
        return cells;
    }

    private static void ExtractHeaderMetadata(
        Dictionary<int, string> cells,
        ref string? containerId,
        ref string? pedido41)
    {
        foreach (int colIdx in cells.Keys.OrderBy(k => k))
        {
            string label = cells[colIdx].Trim().ToUpperInvariant();
            string nextVal = cells.TryGetValue(colIdx + 1, out string? next) ? next.Trim() : "";
            if (nextVal.Length == 0)
                continue;

            // 1. Strictly exclude Discrepancy/Summary rows like "Faltante Importación" = 288 (IMP-207 -> 288 fix)
            bool isImportLabel =
                (label.Contains("CONTENEDOR", StringComparison.Ordinal) ||
                 label.Contains("IMPORTACI", StringComparison.Ordinal) ||
                 label == "IMP" ||
                 label.StartsWith("NO. IMP", StringComparison.Ordinal) ||
                 label.StartsWith("# IMP", StringComparison.Ordinal)) &&
                !ImportIdExcludeTokens.Any(tok => label.Contains(tok, StringComparison.Ordinal));

            if (isImportLabel && nextVal.Length > 0 && nextVal != "0")
            {
                string normalized = NormalizeImportIdValue(nextVal);
                bool candidateHasImp = normalized.StartsWith("IMP-", StringComparison.OrdinalIgnoreCase);

                if (string.IsNullOrWhiteSpace(containerId) ||
                    (candidateHasImp && !containerId.StartsWith("IMP-", StringComparison.OrdinalIgnoreCase)))
                {
                    containerId = normalized;
                }
            }
            else if (pedido41 is null &&
                (label.Contains("PEDIDO 41", StringComparison.Ordinal) ||
                 label.Contains("PED/NAC", StringComparison.Ordinal) ||
                 label.Contains("NACIONALIZACI", StringComparison.Ordinal)))
            {
                pedido41 = nextVal;
            }
        }
    }

    private static string NormalizeImportIdValue(string raw)
    {
        string trimmed = raw.Trim();
        Match m = Regex.Match(trimmed, @"^IMP[\s\-_]*(\d+[A-Z0-9\-]*)$", RegexOptions.IgnoreCase);
        return m.Success ? $"IMP-{m.Groups[1].Value.ToUpperInvariant()}" : trimmed;
    }

    private static bool IsPureNumeric(string? val)
    {
        if (string.IsNullOrWhiteSpace(val)) return false;
        string t = val.Trim();
        return t.All(char.IsDigit);
    }

    private static void RefineColumnMap(Dictionary<int, string> row12Cells, Dictionary<string, int> colMap)
    {
        foreach (int colIdx in row12Cells.Keys.OrderBy(k => k))
        {
            string h = row12Cells[colIdx].Replace("\r\n", " ", StringComparison.Ordinal)
                .Replace('\n', ' ')
                .Trim()
                .ToUpperInvariant();

            if (h.Equals("MATERIAL", StringComparison.Ordinal))
                colMap["material"] = colIdx;
            else if (colIdx < 20 && h.Contains("CANTIDAD PEDIDO 41", StringComparison.Ordinal))
                colMap["units_ped41"] = colIdx;
            else if (colIdx >= 50 && h.Contains("CANTIDAD PEDIDO 41", StringComparison.Ordinal))
                colMap["qty_be"] = colIdx;
            else if (h.Contains("BULTOS SEGÚN PAKING LIST", StringComparison.Ordinal) ||
                     h.Contains("BULTOS SEGUN PAKING LIST", StringComparison.Ordinal) ||
                     h.Contains("BULTOS SEGÚN PACKING LIST", StringComparison.Ordinal) ||
                     h.Contains("BULTOS SEGUN PACKING LIST", StringComparison.Ordinal))
            {
                colMap["cartons_aw"] = colIdx;
            }
        }
    }

    private static DeclaredItem? ParseDeclaredRow(
        int rowIdx,
        Dictionary<int, string> cells,
        Dictionary<string, int> colMap,
        Dictionary<int, int> rowCurves,
        List<string> warnings)
    {
        string matCode = OpenXmlCellHelper.NormalizeMaterialCode(GetCell(cells, colMap, "material"));
        if (matCode.Length == 0 ||
            matCode.Equals("MATERIAL", StringComparison.OrdinalIgnoreCase) ||
            matCode.Contains("TOTAL", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        int? curve = QcReader.ToInt(GetCell(cells, colMap, "curve_bb"));
        if (curve is > 0)
            rowCurves[rowIdx] = curve.Value;

        int declaredUnits = ResolveDeclaredUnits(cells, colMap, curve);
        int declaredCartons = ParseCartonCount(GetCell(cells, colMap, "cartons_aw"), rowIdx, matCode, warnings);
        string rawUom = QcReader.ToCleanString(GetCell(cells, colMap, "uom_bc")).ToUpperInvariant();

        return new DeclaredItem
        {
            MaterialCode = matCode,
            OrderPo43 = NullIfEmpty(QcReader.ToCleanString(GetCell(cells, colMap, "po43"))),
            Category = NullIfEmpty(QcReader.ToCleanString(GetCell(cells, colMap, "category"))),
            VendorMaterial = NullIfEmpty(QcReader.ToCleanString(GetCell(cells, colMap, "vendor_material"))),
            Description = NullIfEmpty(QcReader.ToCleanString(GetCell(cells, colMap, "description"))),
            SourceRow = rowIdx,
            DeclaredUnits = declaredUnits,
            DeclaredCartons = declaredCartons,
            Uom = rawUom.Length > 0 ? rawUom : "UN",
            UnitCost = ResolveUnitCost(cells, colMap)
        };
    }

    private static int ResolveDeclaredUnits(Dictionary<int, string> cells, Dictionary<string, int> colMap, int? curve)
    {
        if (QcReader.ToInt(GetCell(cells, colMap, "units_ped41")) is int ped41)
            return ped41;

        if (QcReader.ToInt(GetCell(cells, colMap, "qty_be")) is int qtyBe)
        {
            int mult = curve is > 0 ? curve.Value : 1;
            return checked(qtyBe * mult);
        }

        return 0;
    }

    private static int ParseCartonCount(
        string? rawAw,
        int rowIdx,
        string matCode,
        List<string> warnings)
    {
        string cleaned = QcReader.ToCleanString(rawAw);
        if (cleaned.Length == 0 || cleaned == "-")
        {
            warnings.Add($"AW_MISSING_IN_TEMPLATE: Row {rowIdx} ({matCode}) Col AW is blank; defaulted DeclaredCartons to 0.");
            return 0;
        }

        if (double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out double scalar) &&
            !double.IsNaN(scalar) && !double.IsInfinity(scalar))
        {
            return checked((int)Math.Round(scalar, MidpointRounding.ToEven));
        }

        if (cleaned.IndexOfAny(['*', 'x', 'X', '×', '+']) >= 0)
            return QuantityEngine.ParsePackagingExpression(cleaned).TotalCartons;

        warnings.Add($"AW_UNPARSEABLE_IN_TEMPLATE: Row {rowIdx} ({matCode}) Col AW='{cleaned}'; defaulted DeclaredCartons to 0.");
        return 0;
    }

    private static decimal ResolveUnitCost(Dictionary<int, string> cells, Dictionary<string, int> colMap)
    {
        string? afRaw = GetCell(cells, colMap, "cost_af");
        if (decimal.TryParse(afRaw?.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out decimal af) && af > 0m)
            return af;

        string? aeRaw = GetCell(cells, colMap, "cost_ae");
        if (decimal.TryParse(aeRaw?.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out decimal ae) && ae > 0m)
            return ae;

        return 0m;
    }

    private static string? GetCell(Dictionary<int, string> cells, Dictionary<string, int> colMap, string key) =>
        colMap.TryGetValue(key, out int colIdx) && cells.TryGetValue(colIdx, out string? value)
            ? value
            : null;

    private static string? NullIfEmpty(string value) =>
        value.Length == 0 ? null : value;
}
