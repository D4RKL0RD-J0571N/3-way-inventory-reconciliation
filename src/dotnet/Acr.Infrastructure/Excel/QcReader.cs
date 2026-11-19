using System.Globalization;
using System.Text;
using Acr.Core.Models;
using Acr.Core.Services;
using Acr.Infrastructure.Safety;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Infrastructure.Excel;

/// <summary>
/// Structured item line extracted from a QC inspection worksheet (<c>CONTROL DATA</c>).
/// </summary>
public sealed record QcInspectionItem
{
    /// <summary>1-based Excel row index.</summary>
    public int RowIndex { get; init; }

    /// <summary>Pallet / Tarima identifier or label (e.g., <c>VARIANTE SUELTA</c>).</summary>
    public string? PalletNumber { get; init; }

    /// <summary>Product category or UEN.</summary>
    public string Category { get; init; } = "";

    /// <summary>Vendor material reference code.</summary>
    public string VendorMaterial { get; init; } = "";

    /// <summary>Internal material / SKU code.</summary>
    public string MaterialCode { get; init; } = "";

    /// <summary>Material description.</summary>
    public string Description { get; init; } = "";

    /// <summary>Unit of measure resolved with v1 Python fallback rules.</summary>
    public string Uom { get; init; } = "UN";

    /// <summary>Declared quantity from QC sheet.</summary>
    public int? DeclaredUnits { get; init; }

    /// <summary>System prepack curve from QC sheet.</summary>
    public int? CurveSystem { get; init; }

    /// <summary>Physical prepack curve from QC sheet.</summary>
    public int? CurvePhysical { get; init; }

    /// <summary>Full carton count.</summary>
    public int? FullCartons { get; init; }

    /// <summary>Units per full carton.</summary>
    public int? UnitsPerFullCarton { get; init; }

    /// <summary>Complement carton count.</summary>
    public int? ComplementCartons { get; init; }

    /// <summary>Units in complement cartons or loose units.</summary>
    public int? UnitsInComplement { get; init; }

    /// <summary>Total cartons reported in the QC sheet.</summary>
    public int? ReportedTotalCartons { get; init; }

    /// <summary>Total units reported in the QC sheet.</summary>
    public int? ReportedTotalUnits { get; init; }

    /// <summary>Total prepacks reported in the QC sheet.</summary>
    public int? ReportedTotalPrk { get; init; }

    /// <summary>Reported difference column value.</summary>
    public int? ReportedDiff { get; init; }

    /// <summary>Operator observation text.</summary>
    public string? Observation { get; init; }

    /// <summary>True when row is marked as <c>VARIANTE SUELTA</c>.</summary>
    public bool IsVarianteSuelta { get; init; }

    /// <summary>True when row represents a child size SKU under a primary prepack material.</summary>
    public bool IsChildSku { get; init; }

    /// <summary>Parent material code when <see cref="IsChildSku"/> is true.</summary>
    public string? ParentMaterial { get; init; }

    /// <summary>Deterministic calculation result from <see cref="QuantityEngine"/>.</summary>
    public CartonCalculationResult? CalculationResult { get; init; }
}

/// <summary>
/// Complete QC inspection report parsed from a container QC workbook.
/// </summary>
public sealed record QcInspectionReport
{
    /// <summary>Container identifier from row 1-5 metadata or filename fallback.</summary>
    public string ContainerId { get; init; } = "";

    /// <summary>Canonicalized path of the QC workbook.</summary>
    public string FilePath { get; init; } = "";

    /// <summary>Extracted header metadata (<c>container_id</c>, <c>responsible</c>, <c>nationalization_po</c>, <c>date</c>).</summary>
    public IReadOnlyDictionary<string, string> Metadata { get; init; } = new Dictionary<string, string>();

    /// <summary>All parsed merchandise and child SKU records.</summary>
    public IReadOnlyList<QcInspectionItem> Items { get; init; } = Array.Empty<QcInspectionItem>();

    /// <summary>Sum of calculated units across primary (non-child) items.</summary>
    public int TotalCalculatedUnits { get; init; }

    /// <summary>Sum of calculated cartons across primary (non-child) items.</summary>
    public int TotalCalculatedCartons { get; init; }

    /// <summary>Sum of reported units across primary (non-child) items.</summary>
    public int TotalReportedUnits { get; init; }

    /// <summary>Sum of reported cartons across primary (non-child) items.</summary>
    public int TotalReportedCartons { get; init; }

    /// <summary>Subset of <see cref="Items"/> that emitted calculation audit flags.</summary>
    public IReadOnlyList<QcInspectionItem> FlaggedItems { get; init; } = Array.Empty<QcInspectionItem>();
}

/// <summary>
/// Forward-only streaming OpenXML reader for Quality Control workbooks (<c>CONTROL DATA</c>),
/// achieving exact v1 behavioral parity with <c>src/qc_reader.py</c> while exiting early after
/// 15 consecutive blank rows to avoid scanning 70-83 MB trailing empty-row bloat.
/// </summary>
public static class QcReader
{
    private readonly record struct QcCellValue(string? Text, bool IsTruthy);

    private static readonly KeyValuePair<string, string[]>[] FieldKeywords =
    [
        new("pallet", ["TARIMA", "NUMERO DE TARIMA"]),
        new("category", ["CATEGORIA", "CATEGORÍA", "UEN"]),
        new("vendor_material", ["MATERIAL PROVEEDOR", "MAT. PROVEEDOR"]),
        new("material", ["MATERIAL"]),
        new("declared_units", ["CANTIDAD DE UNIDADES", "CANTIDAD"]),
        new("curve_system", ["CURVA SISTEMA"]),
        new("curve_physical", ["CURVA FISICA", "CURVA FÍSICA"]),
        new("uom", ["UM", "UNIDAD DE MEDIDA"]),
        new("description", ["DESCRIPCIÓN", "DESCRIPCION"]),
        new("full_cartons", ["BULTOS ENTEROS"]),
        new("units_per_full", ["UNIDADES POR BULTOS COMPLETOS", "UNIDADES POR BULTO"]),
        new("comp_cartons", ["BULTOS COMPLEMENTOS", "BULTOS COMPLEMENTO"]),
        new("units_comp", ["UNIDADES BULTOS COMPLEMENTOS", "UNIDADES EN BULTOS COMPLEMENTOS"]),
        new("total_cartons", ["TOTAL DE BULTOS", "TOTAL BULTOS"]),
        new("total_units", ["TOTAL DE UNIDADES", "TOTAL UNIDADES"]),
        new("total_prk", ["TOTAL PRK"]),
        new("diff_units", ["DIFERENCIA EN UNIDADES", "DIFERENCIA"]),
        new("observation", ["OBSERVACIONES", "OBSERVACION"])
    ];

    private static readonly string[] FooterKeywords = ["TOTAL", "RESUMEN", "FIRMA"];

    /// <summary>
    /// Reads and validates a QC inspection workbook in a single forward-only OpenXML pass.
    /// </summary>
    public static QcInspectionReport Read(string filePath, OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        string resolvedPath = guard is not null ? guard.EnsureCanRead(filePath) : Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
            throw new FileNotFoundException($"QC file not found at: {resolvedPath}", resolvedPath);

        using WorkbookReadSession session = OpenXmlRelsSanitizer.OpenReadOnly(resolvedPath);
        WorkbookPart wbPart = session.Document.WorkbookPart ?? throw new InvalidDataException($"Missing WorkbookPart in '{resolvedPath}'.");

        WorksheetPart wsPart = ResolveQcWorksheetPart(wbPart);
        IReadOnlyList<string> sharedStrings = ReadSharedStrings(wbPart);

        var metadata = new Dictionary<string, string>(StringComparer.Ordinal);
        var items = new List<QcInspectionItem>();
        var flagged = new List<QcInspectionItem>();
        StreamQcSheet(wsPart, sharedStrings, metadata, items, flagged);

        string containerId = metadata.TryGetValue("container_id", out string? metaContainer) && !string.IsNullOrEmpty(metaContainer)
            ? metaContainer
            : Path.GetFileName(resolvedPath);

        (int calcU, int calcC, int repU, int repC) = SumReportTotals(items);
        return new QcInspectionReport
        {
            ContainerId = containerId,
            FilePath = resolvedPath,
            Metadata = metadata,
            Items = items,
            TotalCalculatedUnits = calcU,
            TotalCalculatedCartons = calcC,
            TotalReportedUnits = repU,
            TotalReportedCartons = repC,
            FlaggedItems = flagged
        };
    }

    internal static int FindQcSheetIndex(IReadOnlyList<string> sheetNames)
    {
        for (int i = 0; i < sheetNames.Count; i++)
        {
            if (sheetNames[i].Contains("CONTROL DATA", StringComparison.OrdinalIgnoreCase))
                return i;
        }

        for (int i = 0; i < sheetNames.Count; i++)
        {
            if (sheetNames[i].Contains("CONTROL", StringComparison.OrdinalIgnoreCase) ||
                sheetNames[i].Contains("CALIDAD", StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return 0;
    }

    internal static Dictionary<string, int> BuildColumnMapping(IReadOnlyDictionary<int, string?> headerCells)
    {
        var mapping = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (int colIdx in headerCells.Keys.OrderBy(k => k))
        {
            string? raw = headerCells[colIdx];
            if (string.IsNullOrEmpty(raw))
                continue;

            string cleaned = raw.Replace("\r\n", " ", StringComparison.Ordinal)
                .Replace('\n', ' ')
                .Trim()
                .ToUpperInvariant();
            MapSingleHeaderCell(colIdx, cleaned, mapping);
        }

        return mapping;
    }

    internal static int ParseColumnIndex(string? cellRef, ref int fallbackIndex)
    {
        if (string.IsNullOrEmpty(cellRef))
            return fallbackIndex++;

        int col = 0;
        bool hasLetter = false;
        foreach (char ch in cellRef)
        {
            if (ch is >= 'A' and <= 'Z')
            {
                col = (col * 26) + (ch - 'A' + 1);
                hasLetter = true;
            }
            else if (ch is >= 'a' and <= 'z')
            {
                col = (col * 26) + (ch - 'a' + 1);
                hasLetter = true;
            }
            else
            {
                break;
            }
        }

        int resolved = hasLetter ? col - 1 : fallbackIndex;
        fallbackIndex = resolved + 1;
        return resolved;
    }

    internal static int? ToInt(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;

        string s = raw.Trim();
        if (s.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            s.Equals("nan", StringComparison.OrdinalIgnoreCase) ||
            s.Equals("null", StringComparison.OrdinalIgnoreCase) ||
            s == "-" ||
            s.Equals("c", StringComparison.OrdinalIgnoreCase) ||
            s.Equals("x", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        if (s.Equals("true", StringComparison.OrdinalIgnoreCase))
            return 1;
        if (s.Equals("false", StringComparison.OrdinalIgnoreCase))
            return 0;

        if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
            double.IsNaN(value) || double.IsInfinity(value))
        {
            return null;
        }

        return checked((int)Math.Round(value, MidpointRounding.ToEven));
    }

    internal static string ToCleanString(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return "";

        string s = raw.Trim();
        if (s.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            s.Equals("nan", StringComparison.OrdinalIgnoreCase) ||
            s.Equals("null", StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        return s;
    }

    private static WorksheetPart ResolveQcWorksheetPart(WorkbookPart wbPart)
    {
        List<Sheet> sheets = wbPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
        if (sheets.Count == 0)
            throw new InvalidDataException("Workbook contains no worksheets.");

        List<string> names = sheets.Select(s => s.Name?.Value ?? "").ToList();
        Sheet target = sheets[FindQcSheetIndex(names)];
        string relId = target.Id?.Value ?? throw new InvalidDataException("Target QC sheet is missing relationship ID.");
        return (WorksheetPart)wbPart.GetPartById(relId);
    }

    private static IReadOnlyList<string> ReadSharedStrings(WorkbookPart wbPart)
    {
        SharedStringTablePart? sstPart = wbPart.SharedStringTablePart;
        if (sstPart is null)
            return Array.Empty<string>();

        var strings = new List<string>();
        using OpenXmlReader reader = OpenXmlReader.Create(sstPart);
        while (reader.Read())
        {
            if (reader.ElementType == typeof(SharedStringItem) && reader.LoadCurrentElement() is SharedStringItem si)
                strings.Add(ExtractCompositeText(si));
        }
        return strings;
    }

    private static string ExtractCompositeText(OpenXmlCompositeElement container)
    {
        var sb = new StringBuilder();
        foreach (OpenXmlElement child in container.ChildElements)
        {
            if (child is Text text)
                sb.Append(text.Text);
            else if (child is Run run && run.Text is not null)
                sb.Append(run.Text.Text);
        }
        return sb.ToString();
    }

    private static void StreamQcSheet(
        WorksheetPart wsPart,
        IReadOnlyList<string> sharedStrings,
        Dictionary<string, string> metadata,
        List<QcInspectionItem> items,
        List<QcInspectionItem> flagged)
    {
        Dictionary<string, int>? mapping = null;
        int? headerRow = null;
        string? lastParent = null;
        int lastRowIndex = 0;
        int emptyCount = 0;

        using OpenXmlReader reader = OpenXmlReader.Create(wsPart);
        while (reader.Read())
        {
            if (reader.ElementType != typeof(Row) || reader.LoadCurrentElement() is not Row rowEl)
                continue;

            int rowIndex = checked((int)(rowEl.RowIndex?.Value ?? (uint)(lastRowIndex + 1)));
            if (ShouldExitForRowGap(rowIndex, ref lastRowIndex, headerRow.HasValue, ref emptyCount))
                break;

            Dictionary<int, QcCellValue> cells = ExtractRowCells(rowEl, sharedStrings);
            if (cells.Count == 0)
            {
                if (headerRow.HasValue && ++emptyCount >= 15)
                    break;
                continue;
            }

            emptyCount = 0;
            ProcessNonEmptyRow(rowIndex, cells, metadata, ref headerRow, ref mapping, ref lastParent, items, flagged);
        }
    }

    private static bool ShouldExitForRowGap(int rowIndex, ref int lastRowIndex, bool hasHeader, ref int emptyCount)
    {
        int gap = rowIndex - lastRowIndex - 1;
        lastRowIndex = rowIndex;
        if (gap > 0 && hasHeader)
        {
            emptyCount += gap;
            if (emptyCount >= 15)
                return true;
        }
        return false;
    }

    private static void ProcessNonEmptyRow(
        int rowIndex,
        Dictionary<int, QcCellValue> cells,
        Dictionary<string, string> metadata,
        ref int? headerRow,
        ref Dictionary<string, int>? mapping,
        ref string? lastParent,
        List<QcInspectionItem> items,
        List<QcInspectionItem> flagged)
    {
        if (rowIndex <= 5)
        {
            ExtractRowMetadata(cells, metadata);
            return;
        }

        if (!headerRow.HasValue)
        {
            if (IsHeaderRow(cells))
            {
                headerRow = rowIndex;
                mapping = BuildColumnMapping(cells.ToDictionary(kv => kv.Key, kv => kv.Value.Text));
            }
            return;
        }

        QcInspectionItem? record = ParseRowToRecord(rowIndex, cells, mapping!, lastParent);
        if (record is null)
            return;

        if (!record.IsChildSku)
            lastParent = record.MaterialCode;
        items.Add(record);
        if (record.CalculationResult is { Flags.Count: > 0 })
            flagged.Add(record);
    }

    private static Dictionary<int, QcCellValue> ExtractRowCells(Row rowEl, IReadOnlyList<string> sharedStrings)
    {
        var cells = new Dictionary<int, QcCellValue>();
        int nextCol = 0;

        foreach (Cell cell in rowEl.Elements<Cell>())
        {
            int colIdx = ParseColumnIndex(cell.CellReference?.Value, ref nextCol);
            if (ReadCellValue(cell, sharedStrings) is QcCellValue val)
                cells[colIdx] = val;
        }

        return cells;
    }

    private static QcCellValue? ReadCellValue(Cell cell, IReadOnlyList<string> sharedStrings)
    {
        CellValues? type = cell.DataType?.Value;
        if (type == CellValues.SharedString)
            return ReadSharedStringCell(cell, sharedStrings);
        if (type == CellValues.InlineString)
            return cell.InlineString is null ? null : CreateStringCellValue(ExtractCompositeText(cell.InlineString));
        if (type == CellValues.String || type == CellValues.Error)
            return CreateStringCellValue(cell.CellValue?.Text);
        if (type == CellValues.Boolean)
            return ReadBooleanCell(cell.CellValue?.Text);

        return ReadNumericOrDefaultCell(cell.CellValue?.Text);
    }

    private static QcCellValue? ReadSharedStringCell(Cell cell, IReadOnlyList<string> sharedStrings)
    {
        if (!int.TryParse(cell.CellValue?.Text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx) ||
            idx < 0 || idx >= sharedStrings.Count)
        {
            return null;
        }

        string text = sharedStrings[idx];
        return new QcCellValue(text, text.Length > 0);
    }

    private static QcCellValue? CreateStringCellValue(string? text) =>
        text is null ? null : new QcCellValue(text, text.Length > 0);

    private static QcCellValue? ReadBooleanCell(string? raw)
    {
        if (raw is null)
            return null;
        bool isTrue = raw == "1" || raw.Equals("true", StringComparison.OrdinalIgnoreCase);
        return new QcCellValue(isTrue ? "True" : "False", isTrue);
    }

    private static QcCellValue? ReadNumericOrDefaultCell(string? raw)
    {
        if (string.IsNullOrEmpty(raw))
            return null;
        bool truthy = double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double num)
            ? num != 0.0
            : raw.Length > 0;
        return new QcCellValue(raw, truthy);
    }

    private static void ExtractRowMetadata(
        IReadOnlyDictionary<int, QcCellValue> cells,
        Dictionary<string, string> metadata)
    {
        foreach (int colIdx in cells.Keys.OrderBy(k => k))
        {
            string tu = (cells[colIdx].Text ?? "").Trim().ToUpperInvariant();
            string nextVal = cells.TryGetValue(colIdx + 1, out QcCellValue nextCell)
                ? (nextCell.Text ?? "").Trim()
                : "";
            if (nextVal.Length == 0)
                continue;

            if (tu.Contains("CONTENEDOR", StringComparison.Ordinal))
                metadata["container_id"] = nextVal;
            else if (tu.Contains("RESPONSABLE", StringComparison.Ordinal))
                metadata["responsible"] = nextVal;
            else if (tu.Contains("NACIONALIZACIÓN", StringComparison.Ordinal) || tu.Contains("NACIONALIZACION", StringComparison.Ordinal))
                metadata["nationalization_po"] = nextVal;
            else if (tu.Contains("FECHA", StringComparison.Ordinal))
                metadata["date"] = nextVal;
        }
    }

    private static bool IsHeaderRow(IReadOnlyDictionary<int, QcCellValue> cells)
    {
        string rowStr = string.Join(" ", cells.OrderBy(kv => kv.Key)
            .Select(kv => (kv.Value.Text ?? "").ToUpperInvariant()));
        return rowStr.Contains("MATERIAL", StringComparison.Ordinal) &&
               (rowStr.Contains("BULTOS", StringComparison.Ordinal) || rowStr.Contains("UNIDADES", StringComparison.Ordinal));
    }

    private static void MapSingleHeaderCell(int colIdx, string cleaned, Dictionary<string, int> mapping)
    {
        foreach ((string fieldName, string[] keywords) in FieldKeywords)
        {
            if (mapping.ContainsKey(fieldName))
                continue;

            foreach (string kw in keywords)
            {
                if (!cleaned.Contains(kw, StringComparison.Ordinal))
                    continue;
                if (fieldName == "material" && cleaned.Contains("PROVEEDOR", StringComparison.Ordinal))
                    continue;
                if (fieldName == "comp_cartons" && cleaned.Contains("UNIDADES", StringComparison.Ordinal))
                    continue;

                mapping[fieldName] = colIdx;
                break;
            }
        }
    }

    private static bool IsFooterOrTotalRow(string matStr, string descStr)
    {
        string matUpper = matStr.ToUpperInvariant();
        string descUpper = descStr.ToUpperInvariant();

        if (matUpper.Contains("TOTAL", StringComparison.Ordinal) ||
            matUpper.Contains("RESUMEN", StringComparison.Ordinal) ||
            matUpper.Contains("FIRMA", StringComparison.Ordinal) ||
            matUpper.StartsWith("SUMA", StringComparison.Ordinal) ||
            descUpper.StartsWith("TOTAL", StringComparison.Ordinal) ||
            descUpper.StartsWith("RESUMEN", StringComparison.Ordinal))
        {
            return true;
        }

        return !matStr.Any(char.IsDigit) && descStr.Length == 0;
    }

    private static QcInspectionItem? ParseRowToRecord(
        int rowIndex,
        IReadOnlyDictionary<int, QcCellValue> cells,
        IReadOnlyDictionary<string, int> mapping,
        string? lastParent)
    {
        string matRaw = ToCleanString(GetCellText(cells, mapping, "material"));
        string desc = ToCleanString(GetCellText(cells, mapping, "description"));
        if (matRaw.Length == 0 || IsFooterOrTotalRow(matRaw, desc))
            return null;

        string tarima = ToCleanString(GetCellText(cells, mapping, "pallet"));
        bool isVs = tarima.Contains("VARIANTE SUELTA", StringComparison.OrdinalIgnoreCase);
        bool isChild = matRaw.Length >= 10 && tarima.Length == 0 && !isVs;
        string rawUom = ToCleanString(GetCellText(cells, mapping, "uom"));
        string uom = rawUom.Length > 0 ? rawUom : (IsCellTruthy(cells, mapping, "total_prk") ? "PRK" : "UN");

        QcInspectionItem item = BuildItemRecord(rowIndex, cells, mapping, matRaw, desc, tarima, uom, isVs, isChild, lastParent);
        return item with { CalculationResult = QuantityEngine.CalculateCartonQuantities(BuildCalculationInput(item)) };
    }

    private static QcInspectionItem BuildItemRecord(
        int rowIndex,
        IReadOnlyDictionary<int, QcCellValue> cells,
        IReadOnlyDictionary<string, int> mapping,
        string matRaw,
        string desc,
        string tarima,
        string uom,
        bool isVs,
        bool isChild,
        string? lastParent)
    {
        string obs = ToCleanString(GetCellText(cells, mapping, "observation"));
        return new QcInspectionItem
        {
            RowIndex = rowIndex,
            PalletNumber = tarima.Length > 0 ? tarima : null,
            Category = ToCleanString(GetCellText(cells, mapping, "category")),
            VendorMaterial = ToCleanString(GetCellText(cells, mapping, "vendor_material")),
            MaterialCode = matRaw,
            Description = desc,
            Uom = uom,
            DeclaredUnits = ToInt(GetCellText(cells, mapping, "declared_units")),
            CurveSystem = ToInt(GetCellText(cells, mapping, "curve_system")),
            CurvePhysical = ToInt(GetCellText(cells, mapping, "curve_physical")),
            FullCartons = ToInt(GetCellText(cells, mapping, "full_cartons")),
            UnitsPerFullCarton = ToInt(GetCellText(cells, mapping, "units_per_full")),
            ComplementCartons = ToInt(GetCellText(cells, mapping, "comp_cartons")),
            UnitsInComplement = ToInt(GetCellText(cells, mapping, "units_comp")),
            ReportedTotalCartons = ToInt(GetCellText(cells, mapping, "total_cartons")),
            ReportedTotalUnits = ToInt(GetCellText(cells, mapping, "total_units")),
            ReportedTotalPrk = ToInt(GetCellText(cells, mapping, "total_prk")),
            ReportedDiff = ToInt(GetCellText(cells, mapping, "diff_units")),
            Observation = obs.Length > 0 ? obs : null,
            IsVarianteSuelta = isVs,
            IsChildSku = isChild,
            ParentMaterial = isChild ? lastParent : null
        };
    }

    private static CartonCalculationInput BuildCalculationInput(QcInspectionItem item) =>
        new()
        {
            FullCartons = item.FullCartons,
            UnitsPerFullCarton = item.UnitsPerFullCarton,
            ComplementCartons = item.ComplementCartons,
            UnitsInComplement = item.UnitsInComplement,
            ReportedTotalCartons = item.ReportedTotalCartons,
            ReportedTotalUnits = item.ReportedTotalUnits,
            ReportedDiff = item.ReportedDiff,
            Curve = item.CurveSystem is > 0 ? item.CurveSystem : item.CurvePhysical,
            Uom = item.Uom,
            Observation = item.Observation,
            IsVarianteSuelta = item.IsVarianteSuelta
        };

    private static string? GetCellText(
        IReadOnlyDictionary<int, QcCellValue> cells,
        IReadOnlyDictionary<string, int> mapping,
        string key) =>
        mapping.TryGetValue(key, out int colIdx) && cells.TryGetValue(colIdx, out QcCellValue val)
            ? val.Text
            : null;

    private static bool IsCellTruthy(
        IReadOnlyDictionary<int, QcCellValue> cells,
        IReadOnlyDictionary<string, int> mapping,
        string key) =>
        mapping.TryGetValue(key, out int colIdx) && cells.TryGetValue(colIdx, out QcCellValue val) && val.IsTruthy;

    private static (int CalcU, int CalcC, int RepU, int RepC) SumReportTotals(IReadOnlyList<QcInspectionItem> items)
    {
        int calcU = 0;
        int calcC = 0;
        int repU = 0;
        int repC = 0;

        checked
        {
            foreach (QcInspectionItem item in items)
            {
                if (item.IsChildSku)
                    continue;
                if (item.CalculationResult is { } calc)
                {
                    calcU += calc.CalculatedUnits;
                    calcC += calc.CalculatedCartons;
                }
                repU += item.ReportedTotalUnits ?? 0;
                repC += item.ReportedTotalCartons ?? 0;
            }
        }

        return (calcU, calcC, repU, repC);
    }
}
