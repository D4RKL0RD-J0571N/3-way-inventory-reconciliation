using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Acr.Core.Services;

/// <summary>
/// Parsed and sanitized SAP clipboard table representation.
/// </summary>
public sealed record SanitizedSapTable
{
    /// <summary>Detected delimiter label in Spanish.</summary>
    public string DelimiterLabel { get; init; } = "Tabulación";

    /// <summary>Sanitized column headers.</summary>
    public IReadOnlyList<string> Headers { get; init; } = Array.Empty<string>();

    /// <summary>Sanitized matrix rows.</summary>
    public IReadOnlyList<IReadOnlyList<string>> Rows { get; init; } = Array.Empty<IReadOnlyList<string>>();

    /// <summary>Clean TSV representation suitable for pasting into Excel or ERP.</summary>
    public string CleanTsv { get; init; } = "";

    /// <summary>Total valid data rows sanitized.</summary>
    public int RowCount => Rows.Count;
}

/// <summary>
/// High-performance data sanitization and defensive conversion utilities for SAP clipboard grids.
/// Accurately preserves negative signs (e.g. '$-50.00' and '(48)'), decodes floating-point scientific notation
/// (e.g. '1.455102001E+09') without corrupting alphanumeric SKUs like '1E2', and strips 18-character SAP zero padding.
/// </summary>
public static class SapClipboardSanitizer
{
    private static readonly Regex TrueScientificRegex =
        new(@"^(?:\d+\.\d+[eE][+-]?\d+|\d+[eE][+-]\d+)$", RegexOptions.Compiled);

    private static readonly Regex TrailingZeroRegex =
        new(@"\.\d*?0+$", RegexOptions.Compiled);

    private static readonly Regex CurrencySignPrefixRegex =
        new(@"^[€$₡£\s]*-", RegexOptions.Compiled);

    /// <summary>
    /// Cleans and trims an arbitrary cell string, filtering formula errors and null placeholders.
    /// </summary>
    public static string CleanString(string? val)
    {
        if (string.IsNullOrWhiteSpace(val))
            return "";

        string s = val.Replace('\u00a0', ' ').Trim();
        string low = s.ToLowerInvariant();
        if (low is "nan" or "none" or "null" or "#n/a" or "#value!" or "#ref!" or "#div/0!" or "#name?")
            return "";

        // Collapse multiple whitespace
        var sb = new StringBuilder(s.Length);
        bool inSpace = false;
        foreach (char c in s)
        {
            if (char.IsWhiteSpace(c))
            {
                if (!inSpace)
                {
                    sb.Append(' ');
                    inSpace = true;
                }
            }
            else
            {
                sb.Append(c);
                inSpace = false;
            }
        }
        return sb.ToString().Trim();
    }

    /// <summary>
    /// Cleans and normalizes an SAP material SKU code:
    /// - Decodes scientific notation (<c>1.455102001E+09</c> -> <c>1455102001</c>) without touching <c>1E2</c>
    /// - Strips trailing decimal zeros (<c>1455102003.0</c> -> <c>1455102003</c>)
    /// - Strips SAP 18-digit leading zeros when code is numeric and longer than 10 digits
    /// </summary>
    public static string CleanMaterialCode(string? val)
    {
        if (string.IsNullOrWhiteSpace(val))
            return "";

        string s = val.Trim().Replace("\u00a0", "", StringComparison.Ordinal);
        string low = s.ToLowerInvariant();
        if (low is "nan" or "none" or "null" or "#n/a" or "#value!" or "#ref!")
            return "";

        // 1. Strictly decode true numeric scientific notation (e.g. 1.455102001E+09), never alphanumeric like '1E2'
        if (TrueScientificRegex.IsMatch(s))
        {
            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double d) &&
                !double.IsNaN(d) && !double.IsInfinity(d))
            {
                s = Math.Round(d).ToString("F0", CultureInfo.InvariantCulture);
            }
        }

        // 2. Strip trailing .0 from numeric Excel values
        if (s.IndexOf('.') >= 0 && Regex.IsMatch(s, @"^\d+\.0+$"))
        {
            s = Regex.Replace(s, @"\.0+$", "");
        }

        // 3. Keep alphanumeric, underscore, and hyphen
        var filtered = new StringBuilder(s.Length);
        foreach (char c in s)
        {
            if (char.IsLetterOrDigit(c) || c == '_' || c == '-')
                filtered.Append(c);
        }
        s = filtered.ToString();

        // 4. Strip SAP 18-char leading zero padding (e.g. 000000001455102001 -> 1455102001)
        if (s.Length > 10 && s.StartsWith("000", StringComparison.Ordinal) && s.All(char.IsDigit))
        {
            string unpadded = s.TrimStart('0');
            s = unpadded.Length > 0 ? unpadded : "0";
        }

        return s;
    }

    /// <summary>
    /// Defensively parses numerical strings, preserving accounting parentheses <c>(48)</c> as negative
    /// and currency values with negative signs like <c>$-50.00</c> or <c>-$50.00</c>.
    /// </summary>
    public static decimal CleanNumber(string? val, decimal defaultVal = 0m)
    {
        if (string.IsNullOrWhiteSpace(val))
            return defaultVal;

        string s = val.Trim().Replace("\u00a0", "", StringComparison.Ordinal);
        string low = s.ToLowerInvariant();
        if (low is "nan" or "none" or "null" or "-" or "c" or "x" or "#n/a" or "#value!" or "#ref!" or "#div/0!")
            return defaultVal;

        // Detect negative before stripping currency symbols
        bool isNegative = false;
        if (s.StartsWith('(') && s.EndsWith(')'))
        {
            isNegative = true;
            s = s[1..^1].Trim();
        }
        else if (s.Contains('-'))
        {
            if (CurrencySignPrefixRegex.IsMatch(s) || s.StartsWith('-'))
                isNegative = true;
        }

        // Strip non-digit / non-separators
        var cleanSb = new StringBuilder(s.Length);
        foreach (char c in s)
        {
            if (char.IsDigit(c) || c == '.' || c == ',')
                cleanSb.Append(c);
        }
        s = cleanSb.ToString();

        if (s.Length == 0)
            return defaultVal;

        // Format commas vs dots
        if (s.Contains(',') && s.Contains('.'))
        {
            if (s.LastIndexOf(',') > s.LastIndexOf('.'))
            {
                // European format 1.234,56
                s = s.Replace(".", "", StringComparison.Ordinal).Replace(',', '.');
            }
            else
            {
                // US format 1,234.56
                s = s.Replace(",", "", StringComparison.Ordinal);
            }
        }
        else if (s.Contains(','))
        {
            string[] parts = s.Split(',');
            if (parts.Length == 2 && (parts[1].Length == 1 || parts[1].Length == 2))
                s = s.Replace(',', '.');
            else
                s = s.Replace(",", "", StringComparison.Ordinal);
        }

        if (decimal.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal num))
            return isNegative ? -num : num;

        return defaultVal;
    }

    /// <summary>
    /// Parses and sanitizes a raw clipboard dump from SAP or Excel.
    /// Supports tab-delimited, pipe-delimited, semicolon-delimited, and comma-delimited layouts.
    /// </summary>
    public static SanitizedSapTable SanitizeSapTable(string rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText))
            return new SanitizedSapTable();

        string[] lines = rawText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                                .Where(l => !string.IsNullOrWhiteSpace(l))
                                .ToArray();

        if (lines.Length == 0)
            return new SanitizedSapTable();

        string firstLine = lines[0];
        char sep = '\t';
        string sepLabel = "Tabulación";

        if (firstLine.Contains('\t'))
        {
            sep = '\t';
            sepLabel = "Tabulación";
        }
        else if (firstLine.Contains('|'))
        {
            sep = '|';
            sepLabel = "Tubo (|)";
        }
        else if (firstLine.Contains(';'))
        {
            sep = ';';
            sepLabel = "Punto y coma (;)";
        }
        else if (firstLine.Contains(','))
        {
            sep = ',';
            sepLabel = "Coma (,)";
        }

        string[] rawHeaderParts = firstLine.Split(sep);
        var headers = new List<string>();

        for (int i = 0; i < rawHeaderParts.Length; i++)
        {
            string h = CleanString(rawHeaderParts[i]);
            // Ignore blank outer cells caused by leading/trailing pipes (e.g. | col1 | col2 |)
            if (sep == '|' && h.Length == 0 && (i == 0 || i == rawHeaderParts.Length - 1))
                continue;

            headers.Add(h.Length > 0 ? h : $"Columna {headers.Count + 1}");
        }

        var cleanedRows = new List<IReadOnlyList<string>>();
        var tsvSb = new StringBuilder();
        tsvSb.AppendLine(string.Join("\t", headers));

        for (int lineIdx = 1; lineIdx < lines.Length; lineIdx++)
        {
            string line = lines[lineIdx];
            string[] rawCells = line.Split(sep);
            var rowCells = new List<string>(headers.Count);

            int cellIdx = 0;
            for (int i = 0; i < rawCells.Length && cellIdx < headers.Count; i++)
            {
                if (sep == '|' && (i == 0 || i == rawCells.Length - 1) && string.IsNullOrWhiteSpace(rawCells[i]))
                    continue;

                string val = rawCells[i];
                string hdrLow = headers[cellIdx].ToLowerInvariant();

                if (hdrLow.Contains("material") || hdrLow.Contains("producto") || hdrLow.Contains("sku") || hdrLow.Contains("código") || hdrLow.Contains("codigo"))
                {
                    rowCells.Add(CleanMaterialCode(val));
                }
                else if (hdrLow.Contains("cant") || hdrLow.Contains("unidades") || hdrLow.Contains("qty") || hdrLow.Contains("recibido") || hdrLow.Contains("bulto"))
                {
                    decimal d = CleanNumber(val);
                    rowCells.Add(d == Math.Floor(d) ? ((long)d).ToString(CultureInfo.InvariantCulture) : d.ToString("F2", CultureInfo.InvariantCulture));
                }
                else
                {
                    rowCells.Add(CleanString(val));
                }
                cellIdx++;
            }

            // Fill trailing columns if line had fewer cells
            while (rowCells.Count < headers.Count)
                rowCells.Add("");

            cleanedRows.Add(rowCells);
            tsvSb.AppendLine(string.Join("\t", rowCells));
        }

        return new SanitizedSapTable
        {
            DelimiterLabel = sepLabel,
            Headers = headers,
            Rows = cleanedRows,
            CleanTsv = tsvSb.ToString().TrimEnd()
        };
    }
}
