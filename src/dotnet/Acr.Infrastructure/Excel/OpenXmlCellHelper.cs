using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Infrastructure.Excel;

/// <summary>
/// Shared internal OpenXML cell, row, shared-string, and normalization helpers for
/// <see cref="ClosureExporter"/>, <see cref="FormulaIntegrityAuditor"/>, and <see cref="DailyLogRepository"/>.
/// </summary>
internal static class OpenXmlCellHelper
{
    public static string GetColumnName(int colIdxZeroBased)
    {
        int colNum = checked(colIdxZeroBased + 1);
        var sb = new StringBuilder(3);
        while (colNum > 0)
        {
            int rem = (colNum - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            colNum = (colNum - 1) / 26;
        }
        return sb.ToString();
    }

    public static string BuildCellRef(int colIdxZeroBased, int rowIndex) =>
        $"{GetColumnName(colIdxZeroBased)}{rowIndex.ToString(CultureInfo.InvariantCulture)}";

    public static WorksheetPart? FindWorksheetPart(
        WorkbookPart wbPart,
        string sheetName,
        bool allowContains = false)
    {
        IEnumerable<Sheet> sheets = wbPart.Workbook?.Sheets?.Elements<Sheet>() ?? [];
        foreach (Sheet sheet in sheets)
        {
            string? name = sheet.Name?.Value?.Trim();
            if (string.IsNullOrEmpty(name))
                continue;

            bool match = allowContains
                ? name.Contains(sheetName, StringComparison.OrdinalIgnoreCase)
                : name.Equals(sheetName, StringComparison.OrdinalIgnoreCase);

            if (match && sheet.Id?.Value is string relId)
                return wbPart.GetPartById(relId) as WorksheetPart;
        }

        return null;
    }

    public static List<string> ReadSharedStrings(WorkbookPart wbPart)
    {
        var result = new List<string>();
        SharedStringTable? table = wbPart.SharedStringTablePart?.SharedStringTable;
        if (table is null)
            return result;

        foreach (SharedStringItem si in table.Elements<SharedStringItem>())
            result.Add(ExtractCompositeText(si));
        return result;
    }

    public static Dictionary<string, int> BuildSharedStringCache(WorkbookPart wbPart)
    {
        var cache = new Dictionary<string, int>(StringComparer.Ordinal);
        SharedStringTablePart sstPart = EnsureSharedStringTablePart(wbPart);
        SharedStringTable table = sstPart.SharedStringTable!;
        int idx = 0;
        foreach (SharedStringItem si in table.Elements<SharedStringItem>())
        {
            string text = ExtractCompositeText(si);
            cache.TryAdd(text, idx);
            idx++;
        }
        return cache;
    }

    public static SharedStringTablePart EnsureSharedStringTablePart(WorkbookPart wbPart)
    {
        SharedStringTablePart sstPart = wbPart.SharedStringTablePart ?? wbPart.AddNewPart<SharedStringTablePart>();
        sstPart.SharedStringTable ??= new SharedStringTable();
        return sstPart;
    }

    public static int GetOrAddSharedStringIndex(
        WorkbookPart wbPart,
        Dictionary<string, int> cache,
        string text)
    {
        if (cache.TryGetValue(text, out int existing))
            return existing;

        SharedStringTablePart sstPart = EnsureSharedStringTablePart(wbPart);
        SharedStringTable table = sstPart.SharedStringTable!;
        int newIdx = table.Elements<SharedStringItem>().Count();

        var textEl = new Text(text);
        if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1])))
            textEl.Space = SpaceProcessingModeValues.Preserve;

        table.AppendChild(new SharedStringItem(textEl));
        if (table.UniqueCount is not null)
            table.UniqueCount = checked((uint)(newIdx + 1));
        cache[text] = newIdx;
        return newIdx;
    }

    public static string ExtractCompositeText(OpenXmlCompositeElement container)
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

    public static Row FindOrCreateRow(SheetData sheetData, int rowIndex)
    {
        uint targetRow = checked((uint)rowIndex);
        Row? insertBefore = null;
        foreach (Row row in sheetData.Elements<Row>())
        {
            uint current = row.RowIndex?.Value ?? 0;
            if (current == targetRow)
                return row;
            if (current > targetRow)
            {
                insertBefore = row;
                break;
            }
        }

        var created = new Row { RowIndex = targetRow };
        return insertBefore is not null
            ? sheetData.InsertBefore(created, insertBefore)
            : sheetData.AppendChild(created);
    }

    public static Cell FindOrCreateCell(Row row, int colIdxZeroBased, int rowIndex)
    {
        string targetRef = BuildCellRef(colIdxZeroBased, rowIndex);
        Cell? insertBefore = null;
        int fallbackCol = 0;

        foreach (Cell cell in row.Elements<Cell>())
        {
            int currentCol = QcReader.ParseColumnIndex(cell.CellReference?.Value, ref fallbackCol);
            if (currentCol == colIdxZeroBased)
            {
                cell.CellReference ??= targetRef;
                return cell;
            }
            if (currentCol > colIdxZeroBased)
            {
                insertBefore = cell;
                break;
            }
        }

        var created = new Cell { CellReference = targetRef };
        return insertBefore is not null
            ? row.InsertBefore(created, insertBefore)
            : row.AppendChild(created);
    }

    public static Cell? TryFindCell(Row row, int colIdxZeroBased)
    {
        int fallbackCol = 0;
        foreach (Cell cell in row.Elements<Cell>())
        {
            int currentCol = QcReader.ParseColumnIndex(cell.CellReference?.Value, ref fallbackCol);
            if (currentCol == colIdxZeroBased)
                return cell;
            if (currentCol > colIdxZeroBased)
                return null;
        }
        return null;
    }

    public static void SetNumericCell(
        Row row,
        int colIdxZeroBased,
        int rowIndex,
        double value,
        uint? fallbackStyle = null)
    {
        Cell cell = FindOrCreateCell(row, colIdxZeroBased, rowIndex);
        cell.CellFormula = null;
        cell.InlineString = null;
        cell.DataType = null;
        if (cell.StyleIndex is null && fallbackStyle.HasValue)
            cell.StyleIndex = fallbackStyle.Value;
        cell.CellValue = new CellValue(value.ToString("G17", CultureInfo.InvariantCulture));
    }

    public static void SetIntegerCell(
        Row row,
        int colIdxZeroBased,
        int rowIndex,
        int value,
        uint? fallbackStyle = null)
    {
        Cell cell = FindOrCreateCell(row, colIdxZeroBased, rowIndex);
        cell.CellFormula = null;
        cell.InlineString = null;
        cell.DataType = null;
        if (cell.StyleIndex is null && fallbackStyle.HasValue)
            cell.StyleIndex = fallbackStyle.Value;
        cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
    }

    public static void SetSharedStringCell(
        Row row,
        int colIdxZeroBased,
        int rowIndex,
        string text,
        WorkbookPart wbPart,
        Dictionary<string, int> sstCache,
        uint? fallbackStyle = null)
    {
        int sstIdx = GetOrAddSharedStringIndex(wbPart, sstCache, text);
        Cell cell = FindOrCreateCell(row, colIdxZeroBased, rowIndex);
        cell.CellFormula = null;
        cell.InlineString = null;
        cell.DataType = CellValues.SharedString;
        if (cell.StyleIndex is null && fallbackStyle.HasValue)
            cell.StyleIndex = fallbackStyle.Value;
        cell.CellValue = new CellValue(sstIdx.ToString(CultureInfo.InvariantCulture));
    }

    public static void SetFormulaCell(
        Row row,
        int colIdxZeroBased,
        int rowIndex,
        string formulaText,
        string? cachedValue = null,
        CellValues? cachedDataType = null,
        uint? fallbackStyle = null)
    {
        Cell cell = FindOrCreateCell(row, colIdxZeroBased, rowIndex);
        string cleanFormula = formulaText.StartsWith('=') ? formulaText[1..] : formulaText;
        cell.InlineString = null;
        cell.CellFormula = new CellFormula(cleanFormula);
        cell.DataType = cachedDataType;
        if (cell.StyleIndex is null && fallbackStyle.HasValue)
            cell.StyleIndex = fallbackStyle.Value;
        cell.CellValue = cachedValue is null ? null : new CellValue(cachedValue);
    }

    public static void ClearFormulaCacheOnly(Cell cell)
    {
        cell.CellValue = null;
        if (cell.DataType?.Value == CellValues.Error)
            cell.DataType = null;
    }

    public static string DecodeCellText(Cell? cell, IReadOnlyList<string> sharedStrings)
    {
        if (cell is null)
            return "";

        CellValues? type = cell.DataType?.Value;
        if (type == CellValues.SharedString)
            return DecodeSharedString(cell.CellValue?.Text, sharedStrings);
        if (type == CellValues.InlineString)
            return cell.InlineString is null ? "" : ExtractCompositeText(cell.InlineString);
        if (type == CellValues.Boolean)
            return cell.CellValue?.Text == "1" ? "True" : "False";

        return cell.CellValue?.Text ?? "";
    }

    private static string DecodeSharedString(string? rawIndex, IReadOnlyList<string> sharedStrings)
    {
        if (!int.TryParse(rawIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx) ||
            idx < 0 || idx >= sharedStrings.Count)
        {
            return "";
        }
        return sharedStrings[idx];
    }

    public static string NormalizeMaterialCode(string? raw)
    {
        string cleaned = QcReader.ToCleanString(raw);
        if (cleaned.Length == 0)
            return "";

        if (cleaned.EndsWith(".0", StringComparison.Ordinal) &&
            double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out double dotVal) &&
            !double.IsNaN(dotVal) && !double.IsInfinity(dotVal))
        {
            double rounded = Math.Round(dotVal, MidpointRounding.ToEven);
            if (Math.Abs(dotVal - rounded) < 1e-6 && rounded >= long.MinValue && rounded < 9223372036854775808.0)
                return ((long)rounded).ToString(CultureInfo.InvariantCulture);
        }

        if ((cleaned.Contains('E', StringComparison.OrdinalIgnoreCase)) &&
            double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out double sciVal) &&
            !double.IsNaN(sciVal) && !double.IsInfinity(sciVal))
        {
            double rounded = Math.Round(sciVal, MidpointRounding.ToEven);
            if (Math.Abs(sciVal - rounded) < 1e-6 && rounded >= long.MinValue && rounded < 9223372036854775808.0)
                return ((long)rounded).ToString(CultureInfo.InvariantCulture);
        }

        return cleaned;
    }

    private static readonly Regex FormulaTokenRegex = new(
        @"""[^""]*""|'[^']*'|(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]*)(?![A-Za-z0-9_(])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static string ShiftFormulaReferences(string formula, int colDelta, int rowDelta)
    {
        if (string.IsNullOrWhiteSpace(formula) || (colDelta == 0 && rowDelta == 0))
            return formula;

        return FormulaTokenRegex.Replace(formula, match =>
        {
            if (match.Value[0] is '"' or '\'')
                return match.Value;

            string colLock = match.Groups[1].Value;
            string colText = match.Groups[2].Value;
            string rowLock = match.Groups[3].Value;
            string rowText = match.Groups[4].Value;

            int newCol = colLock.Length > 0 ? ParseColLetters(colText) : ParseColLetters(colText) + colDelta;
            int rowNum = int.Parse(rowText, NumberStyles.Integer, CultureInfo.InvariantCulture);
            int newRow = rowLock.Length > 0 ? rowNum : rowNum + rowDelta;
            if (newCol < 0 || newRow < 1)
                return match.Value;

            return $"{colLock}{GetColumnName(newCol)}{rowLock}{newRow.ToString(CultureInfo.InvariantCulture)}";
        });
    }

    private static int ParseColLetters(string colText)
    {
        int col = 0;
        foreach (char ch in colText)
            col = (col * 26) + (ch - 'A' + 1);
        return col - 1;
    }

    public static void PrepareWorkbookForFullRecalc(WorkbookPart wbPart)
    {
        if (wbPart.CalculationChainPart is not null)
            wbPart.DeletePart(wbPart.CalculationChainPart);

        Workbook workbook = wbPart.Workbook ?? throw new InvalidDataException("Workbook XML root is missing.");
        workbook.CalculationProperties ??= new CalculationProperties();
        workbook.CalculationProperties.FullCalculationOnLoad = true;
    }
}
