using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;
using ExcelDataReader;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace Acr.Infrastructure.Receiving;

/// <summary>A normalized merchandise row extracted from a supplier packing list.</summary>
public sealed record SupplierPackingListRow(
    string MaterialCode,
    string Description,
    int? Units,
    int? Cartons,
    string SourceLocation,
    string RawText,
    double Confidence = 1.0);

/// <summary>Field-specific declared values explicitly accepted from the supplier document.</summary>
public sealed record SupplierPackingListOverride(
    int? Units,
    int? Cartons,
    string? SourceFilePath = null,
    string? UnitsSourceLocation = null,
    string? CartonsSourceLocation = null);

/// <summary>Reads tabular supplier packing lists from Excel and text-based PDFs.</summary>
public static class SupplierPackingListReader
{
    private static readonly string[] MaterialHeaders = ["MATERIAL", "MATERIAL CODE", "SKU", "ITEM", "ITEM NO", "STYLE", "ARTICULO", "ARTICULO NO", "CODIGO", "REF"];
    private static readonly string[] DescriptionHeaders = ["DESCRIPTION", "PRODUCT DESCRIPTION", "DESCRIPCION", "PRODUCT", "PRODUCT NAME"];
    private static readonly string[] UnitHeaders = ["QUANTITY", "QTY", "UNITS", "PCS", "PIECES", "CANTIDAD", "UNIDADES", "TOTAL QTY"];
    private static readonly string[] CartonHeaders = ["CARTONS", "CTNS", "CARTON", "CARTON NOS", "NO OF CARTONS", "PACKAGES", "CASES", "BULTOS", "CAJAS"];

    /// <summary>Reads supported supplier document formats and returns normalized rows.</summary>
    public static IReadOnlyList<SupplierPackingListRow> Read(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
        if (!File.Exists(filePath)) throw new FileNotFoundException("No se encontró la lista de empaque.", filePath);
        try
        {
            return Path.GetExtension(filePath).ToLowerInvariant() switch
            {
                ".xlsx" or ".xlsm" or ".xls" => ReadExcel(filePath),
                ".pdf" => ReadPdf(filePath),
                ".png" or ".jpg" or ".jpeg" or ".tif" or ".tiff" => ReadImage(filePath),
                _ => throw new NotSupportedException("Formato no compatible. Use Excel (.xls, .xlsx, .xlsm), PDF, PNG, JPG o TIFF.")
            };
        }
        catch (Exception ex) when (IsProtectedDocumentError(ex))
        {
            throw new InvalidDataException("El archivo está cifrado o protegido con contraseña. Proporcione una copia sin protección para importar el Packing List.", ex);
        }
    }

    private static IReadOnlyList<SupplierPackingListRow> ReadExcel(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
        var sheetCandidates = new List<(int Score, IReadOnlyList<SupplierPackingListRow> Rows)>();
        do
        {
            var sheetRows = new List<(int Number, string[] Values)>();
            int rowNumber = 0;
            while (reader.Read())
            {
                rowNumber++;
                var values = new string[reader.FieldCount];
                for (int c = 0; c < reader.FieldCount; c++) values[c] = ToText(reader.GetValue(c));
                sheetRows.Add((rowNumber, values));
            }

            (int Row, int Material, int Description, int Units, int Cartons)? header = FindHeader(sheetRows);
            if (header is null) continue;
            string sheetName = reader.Name ?? "";
            int score = (sheetName.Contains("PACKING", StringComparison.OrdinalIgnoreCase) ? 10 : 0)
                        + (header.Value.Cartons >= 0 ? 10 : 0)
                        + (header.Value.Units >= 0 ? 3 : 0);
            var parsedSheetRows = new List<SupplierPackingListRow>();
            foreach ((int number, string[] values) in sheetRows.Where(x => x.Number > header.Value.Row))
            {
                string material = At(values, header.Value.Material).Trim();
                if (!LooksLikeMaterial(material)) continue;
                string description = At(values, header.Value.Description).Trim();
                int? units = ParseQuantity(At(values, header.Value.Units));
                int? cartons = ParseQuantity(At(values, header.Value.Cartons));
                if (units is null && cartons is null) continue;
                parsedSheetRows.Add(new SupplierPackingListRow(material, description, units, cartons,
                    $"{sheetName}!fila {number}", string.Join(" | ", values.Where(v => v.Length > 0))));
            }
            if (parsedSheetRows.Count > 0) sheetCandidates.Add((score, parsedSheetRows));
        } while (reader.NextResult());

        if (sheetCandidates.Count == 0) throw new InvalidDataException("No se encontró una tabla de packing list. Revise los encabezados de material, unidades y bultos.");
        return sheetCandidates.OrderByDescending(x => x.Score).First().Rows;
    }

    private static (int Row, int Material, int Description, int Units, int Cartons)? FindHeader(
        IReadOnlyList<(int Number, string[] Values)> rows)
    {
        foreach (var row in rows.Take(30))
        {
            int material = -1, description = -1, units = -1, cartons = -1;
            for (int c = 0; c < row.Values.Length; c++)
            {
                string value = NormalizeHeader(row.Values[c]);
                if (material < 0 && MaterialHeaders.Any(h => value == NormalizeHeader(h) || value.Contains(NormalizeHeader(h)))) material = c;
                if (description < 0 && DescriptionHeaders.Any(h => value == NormalizeHeader(h) || value.Contains(NormalizeHeader(h)))) description = c;
                if (units < 0 && UnitHeaders.Any(h => value == NormalizeHeader(h) || value.Contains(NormalizeHeader(h)))) units = c;
                if (cartons < 0 && CartonHeaders.Any(h => value == NormalizeHeader(h) || value.Contains(NormalizeHeader(h)))) cartons = c;
            }
            if (material >= 0 && (units >= 0 || cartons >= 0)) return (row.Number, material, description, units, cartons);
        }
        return null;
    }

    private static IReadOnlyList<SupplierPackingListRow> ReadPdf(string path)
    {
        using PdfDocument document = PdfDocument.Open(path);
        var rows = new List<SupplierPackingListRow>();
        Regex itemPattern = GetPdfItemPattern();

        foreach (var page in document.GetPages())
        {
            string text = ContentOrderTextExtractor.GetText(page);
            foreach (Match match in itemPattern.Matches(text))
            {
                string normalized = Regex.Replace(match.Value, @"\s+", " ").Trim();
                int? units = ParseQuantity(match.Groups["units"].Value);
                int? cartons = ParseQuantity(match.Groups["cartons"].Value);
                if (units is null && cartons is null) continue;
                int lineNumber = text.Take(match.Index).Count(c => c == '\n') + 1;
                rows.Add(new SupplierPackingListRow(match.Groups["code"].Value,
                    match.Groups["description"].Value.Trim(), units, cartons,
                    $"página {page.Number}, línea {lineNumber}", normalized));
            }
        }

        if (rows.Count == 0)
        {
            string? rasterizer = FindExecutable("pdftoppm");
            if (rasterizer is null)
                throw new InvalidDataException("El PDF no contiene texto seleccionable. Para leerlo como escaneo instale pdftoppm y Tesseract OCR, o use una versión digital del packing list.");
            string tempDir = Path.Combine(Path.GetTempPath(), $"acr-pl-{Guid.NewGuid():N}");
            Directory.CreateDirectory(tempDir);
            try
            {
                string prefix = Path.Combine(tempDir, "page");
                RunProcess(rasterizer, ["-png", "-r", "300", path, prefix]);
                foreach (string image in Directory.EnumerateFiles(tempDir, "page-*.png").OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                    rows.AddRange(ReadOcrImage(image, Path.GetFileName(path)));
            }
            finally
            {
                try { Directory.Delete(tempDir, recursive: true); } catch { }
            }
        }
        if (rows.Count == 0)
            throw new InvalidDataException("No se detectaron filas de artículos en el PDF. Revise la calidad del documento o use una versión con texto seleccionable.");
        return rows;
    }

    private static IReadOnlyList<SupplierPackingListRow> ReadImage(string path) => ReadOcrImage(path, Path.GetFileName(path));

    private static IReadOnlyList<SupplierPackingListRow> ReadOcrImage(string imagePath, string sourceName)
    {
        string? tesseract = FindExecutable("tesseract");
        if (tesseract is null)
            throw new NotSupportedException("No se encontró Tesseract OCR. Instale Tesseract localmente para leer imágenes o escaneos, o use Excel / un PDF con texto seleccionable.");

        string tsv;
        try { tsv = RunProcess(tesseract, [imagePath, "stdout", "-l", "eng+spa", "tsv"]); }
        catch (InvalidDataException) { tsv = RunProcess(tesseract, [imagePath, "stdout", "-l", "eng", "tsv"]); }
        var lines = new List<(int Page, int Block, int Paragraph, int Line, double Confidence, string Text)>();
        foreach (string row in tsv.Split('\n', StringSplitOptions.RemoveEmptyEntries).Skip(1))
        {
            string[] fields = row.Split('\t');
            if (fields.Length < 12 || !int.TryParse(fields[0], out int level) || level != 5) continue;
            if (!int.TryParse(fields[1], out int page) || !int.TryParse(fields[2], out int block) ||
                !int.TryParse(fields[3], out int paragraph) || !int.TryParse(fields[4], out int line)) continue;
            double.TryParse(fields[10], NumberStyles.Float, CultureInfo.InvariantCulture, out double confidence);
            string text = fields[11].Trim();
            if (text.Length > 0) lines.Add((page, block, paragraph, line, Math.Max(0, confidence) / 100d, text));
        }

        var parsed = new List<SupplierPackingListRow>();
        var pattern = GetPdfItemPattern();
        foreach (var group in lines.GroupBy(x => (x.Page, x.Block, x.Paragraph, x.Line)))
        {
            string text = string.Join(" ", group.Select(x => x.Text));
            Match match = pattern.Match(text);
            if (!match.Success) continue;
            double confidence = group.Average(x => x.Confidence);
            parsed.Add(new SupplierPackingListRow(
                match.Groups["code"].Value,
                match.Groups["description"].Value.Trim(),
                ParseQuantity(match.Groups["units"].Value),
                ParseQuantity(match.Groups["cartons"].Value),
                $"{sourceName}, OCR línea {group.Key.Line}",
                text,
                confidence));
        }
        return parsed;
    }

    private static Regex GetPdfItemPattern() => new(
        @"^\s*(?<code>[A-Z0-9][A-Z0-9._/-]{4,})\s+(?<description>.+?)\s+(?<units>[\d,]+)\s*(?<uom>SET|PCS?|UN|PAIRS?)\s+(?<cartons>[\d,]+)\s+(?<gross>[\d,.]+)\s+(?<net>[\d,.]+)\s+(?<cbm>[\d,.]+)(?=\s|$)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Singleline | RegexOptions.Multiline);

    private static string? FindExecutable(string name)
    {
        string? path = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(path)) return null;
        foreach (string folder in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            string candidate = Path.Combine(folder.Trim('"'), name + ".exe");
            if (File.Exists(candidate)) return candidate;
            candidate = Path.Combine(folder.Trim('"'), name);
            if (File.Exists(candidate)) return candidate;
        }
        return null;
    }

    private static bool IsProtectedDocumentError(Exception error)
    {
        for (Exception? current = error; current is not null; current = current.InnerException)
        {
            string message = current.Message;
            if (message.Contains("password", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("encrypted", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("decrypt", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("protected", StringComparison.OrdinalIgnoreCase))
                return true;
            if (current.GetType().Name.Contains("Encrypted", StringComparison.OrdinalIgnoreCase) ||
                current.GetType().Name.Contains("Password", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private static string RunProcess(string executable, IReadOnlyList<string> arguments)
    {
        var start = new ProcessStartInfo(executable)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        using Process process = Process.Start(start) ?? throw new InvalidOperationException($"No se pudo iniciar {Path.GetFileName(executable)}.");
        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0) throw new InvalidDataException($"Falló la lectura OCR: {error.Trim()}");
        return output;
    }

    private static string ToText(object? value) => value switch
    {
        null => "",
        DateTime date => date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        double number => number.ToString("0.########", CultureInfo.InvariantCulture),
        _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""
    };

    private static string At(string[] row, int column) => column >= 0 && column < row.Length ? row[column] : "";

    private static int? ParseQuantity(string value)
    {
        string cleaned = value.Trim().Replace(",", "").Replace(" ", "");
        if (cleaned.Length == 0) return null;
        return decimal.TryParse(cleaned, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal parsed)
            ? (int)Math.Round(parsed, MidpointRounding.AwayFromZero)
            : null;
    }

    private static bool LooksLikeMaterial(string value) => value.Length >= 4 && value.Any(char.IsLetterOrDigit) && value.Any(char.IsDigit);

    private static string NormalizeHeader(string value)
    {
        string decomposed = value.Normalize(NormalizationForm.FormD);
        var output = new StringBuilder(decomposed.Length);
        foreach (char c in decomposed)
            if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                output.Append(char.IsLetterOrDigit(c) ? char.ToUpperInvariant(c) : ' ');
        return Regex.Replace(output.ToString(), @"\s+", " ").Trim();
    }
}
