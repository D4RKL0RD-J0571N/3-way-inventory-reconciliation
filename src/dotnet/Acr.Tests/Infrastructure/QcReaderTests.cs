using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using Acr.Core.Models;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Safety;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Tests.Infrastructure;

public sealed class QcReaderTests : IDisposable
{
    private const string ProductionDataDir = @"D:\Development\Software Development\3 Way Inventory Reconciliation\data";
    private const string ProductionDailyLogPath = @"D:\Development\Software Development\3 Way Inventory Reconciliation\Control_Diario_Operaciones_Jostin.xlsx";

    private readonly string _tempRoot;
    private readonly OperationalSafetyGuard _testingGuard;

    public QcReaderTests()
    {
        _tempRoot = Path.Combine(Path.GetTempPath(), $"AcrQcTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempRoot);
        _testingGuard = new OperationalSafetyGuard(
            OperationalSafetyMode.Testing,
            ProductionDataDir,
            ProductionDailyLogPath,
            _tempRoot);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempRoot))
                Directory.Delete(_tempRoot, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    [Fact]
    public void Read_SyntheticQcWorkbook_MatchesPythonReferenceTotalsAndFlags()
    {
        string path = CreateSyntheticQcWorkbook("synthetic_qc.xlsx", includeTrailingBloat: false);

        QcInspectionReport report = QcReader.Read(path, _testingGuard);

        Assert.Equal("TEST-999", report.ContainerId);
        Assert.Equal("Analista Pruebas", report.Metadata["responsible"]);
        Assert.Equal("4100009999", report.Metadata["nationalization_po"]);
        Assert.Equal(5, report.Items.Count);
        Assert.Equal(17, report.TotalCalculatedCartons);
        Assert.Equal(2410, report.TotalCalculatedUnits);
        Assert.Contains(QuantityEngineFlags.FlagQcComplementWithoutCartonCount, report.Items[1].CalculationResult!.Flags);
        Assert.Contains(QuantityEngineFlags.FlagSharedCartonPrimary, report.Items[2].CalculationResult!.Flags);
        Assert.Contains(QuantityEngineFlags.FlagSharedCartonSecondary, report.Items[3].CalculationResult!.Flags);
        Assert.Contains(QuantityEngineFlags.FlagVarianteSueltaRow, report.Items[4].CalculationResult!.Flags);
    }

    [Fact]
    public void Read_MissingFile_ThrowsFileNotFoundException()
    {
        string missing = Path.Combine(_tempRoot, "does_not_exist.xlsm");
        Assert.Throws<FileNotFoundException>(() => QcReader.Read(missing, _testingGuard));
    }

    [Fact]
    public void Read_TrailingEmptyStyledRowsAndRowGaps_ExitsAt15BlankRows()
    {
        string path = CreateSyntheticQcWorkbook("trailing_bloat.xlsx", includeTrailingBloat: true);

        QcInspectionReport report = QcReader.Read(path, _testingGuard);

        Assert.Equal(5, report.Items.Count);
        Assert.Equal(17, report.TotalCalculatedCartons);
        Assert.Equal(2410, report.TotalCalculatedUnits);
    }

    [Fact]
    public void Read_SparseColumnsRichTextSharedStringsErrorsAndZeroPrk_HandlesOpenXmlEdgeCases()
    {
        string path = CreateEdgeCaseQcWorkbook("edge_cases.xlsx");

        QcInspectionReport report = QcReader.Read(path, _testingGuard);

        Assert.Equal(2, report.Items.Count);
        QcInspectionItem first = report.Items[0];
        Assert.Equal("MAT-1001", first.MaterialCode);
        Assert.Equal("Rich Text Desc", first.Description);
        Assert.Equal("UN", first.Uom);
        Assert.Null(first.ReportedDiff);
        Assert.Equal(2, first.CalculationResult!.FullCartonsUsed);
        Assert.Equal(12, first.CalculationResult!.UnitsPerFullCartonUsed);
        Assert.Equal(24, first.CalculationResult!.CalculatedUnits);

        QcInspectionItem second = report.Items[1];
        Assert.Equal("PRK", second.Uom);
        Assert.Equal(4, second.FullCartons);
        Assert.Null(second.UnitsInComplement);
    }

    [Fact]
    public void WorkbookSignatureScanner_SkipsLockfilesAndDistinguishesControlDataVsCierreRecep()
    {
        string qcPath = CreateSimpleWorkbook("qc_only.xlsm", ["CONTROL DATA"], hasVba: true);
        string closureWithControlName = CreateSimpleWorkbook("closure.xlsm", ["STOCK CONTROL DE INVENTARIO", "CIERRE RECEP"], hasVba: true);
        string hybridPath = CreateSimpleWorkbook("hybrid.xlsm", ["CONTROL DATA", "CIERRE RECEP"], hasVba: false);
        string lockPath = Path.Combine(_tempRoot, "~$qc_only.xlsm");
        File.WriteAllText(lockPath, "office-lock-bytes");

        Assert.True(WorkbookSignatureScanner.IsLockOrTempFile(lockPath));
        Assert.Equal(WorkbookKind.Unknown, WorkbookSignatureScanner.Inspect(lockPath, _testingGuard).Kind);
        Assert.Equal(WorkbookKind.QcInspection, WorkbookSignatureScanner.Inspect(qcPath, _testingGuard).Kind);
        Assert.Equal(WorkbookKind.ReceivingClosure, WorkbookSignatureScanner.Inspect(closureWithControlName, _testingGuard).Kind);
        Assert.Equal(WorkbookKind.HybridQcAndClosure, WorkbookSignatureScanner.Inspect(hybridPath, _testingGuard).Kind);

        IReadOnlyList<WorkbookSignature> scanned = WorkbookSignatureScanner.ScanDirectory(_tempRoot, recursive: false, _testingGuard);
        Assert.Equal(3, scanned.Count);
        Assert.DoesNotContain(scanned, s => Path.GetFileName(s.FilePath).StartsWith("~$", StringComparison.Ordinal));
    }

    [Fact]
    public void WorkbookSignatureScanner_NtfsJunctionCycle_CompletesWithoutPathTooLongException()
    {
        string loopRoot = Path.Combine(_tempRoot, "loop_scan");
        string subDir = Path.Combine(loopRoot, "sub");
        Directory.CreateDirectory(subDir);
        string validBook = Path.Combine(subDir, "valid_in_loop.xlsx");
        File.Copy(CreateSimpleWorkbook("temp_src.xlsx", ["CONTROL DATA"], hasVba: false), validBook);
        File.Delete(Path.Combine(_tempRoot, "temp_src.xlsx"));

        string junctionPath = Path.Combine(subDir, "back_to_parent");
        if (!TryCreateNtfsJunction(junctionPath, loopRoot))
            return;

        IReadOnlyList<WorkbookSignature> withGuard = WorkbookSignatureScanner.ScanDirectory(loopRoot, recursive: true, _testingGuard);
        IReadOnlyList<WorkbookSignature> withoutGuard = WorkbookSignatureScanner.ScanDirectory(loopRoot, recursive: true);

        WorkbookSignature sig = Assert.Single(withGuard);
        Assert.Equal(WorkbookKind.QcInspection, sig.Kind);
        Assert.Single(withoutGuard);
    }

    [Fact]
    public void WorkbookSignatureScanner_CaseInsensitiveWorkbookXmlEntry_DetectsSheets()
    {
        string path = Path.Combine(_tempRoot, "uppercase_entry.xlsx");
        using (ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            ZipArchiveEntry entry = zip.CreateEntry("XL/WORKBOOK.XML");
            using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
            writer.Write(
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                "<sheets><sheet name=\"CIERRE RECEP\" sheetId=\"1\" r:id=\"rId1\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/></sheets>" +
                "</workbook>");
        }

        WorkbookSignature sig = WorkbookSignatureScanner.Inspect(path, _testingGuard);
        Assert.Equal(WorkbookKind.ReceivingClosure, sig.Kind);
    }

    [Fact]
    public void QcReader_SumReportTotalsOverflow_ThrowsOverflowException()
    {
        string path = Path.Combine(_tempRoot, "sum_overflow.xlsx");
        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
        {
            SheetData sheetData = InitializeSingleSheetWorkbook(doc, "CONTROL DATA");
            AppendSyntheticMetadataAndHeaders(sheetData);
            sheetData.AppendChild(BuildRow(7,
                ("A7", "1"), ("C7", "2001"), ("D7", "OVERFLOW A"),
                ("G7", "1500000"), ("H7", "1000")));
            sheetData.AppendChild(BuildRow(8,
                ("A8", "2"), ("C8", "2002"), ("D8", "OVERFLOW B"),
                ("G8", "1500000"), ("H8", "1000")));
            doc.WorkbookPart!.Workbook!.Save();
        }

        Assert.Throws<OverflowException>(() => QcReader.Read(path, _testingGuard));
    }

    [Theory]
    [InlineData(@"IMP-103\Informacion General\imp 103 control calidad corregirdo.xlsm", "103", 267, 1038, 102854, 1038, 102711, 41)]
    [InlineData(@"IMP-100\Informacion General\IMP 100 Control Calidad.xlsm", "100", 18, 990, 990, 990, 990, 0)]
    [InlineData(@"IMP-218\IMP-218.xlsm", "218", 8, 781, 51740, 781, 51740, 0)]
    [InlineData(@"IMP-298\Informacion General\imp298 control calidad ver2.xlsm", "298", 163, 1036, 67746, 1032, 67314, 11)]
    [InlineData(@"IMP-327\Informacion General\cierre imp 327 Control de calidad.xlsm", "327", 6, 669, 4014, 669, 4014, 0)]
    [InlineData(@"IMP-348\Informacion General\imp 348.xlsm", "348", 1, 620, 620, 620, 620, 0)]
    [InlineData(@"IMP-406\INFORMACION GENERAL\IMP 406.xlsm", "IMP 406.xlsm", 191, 855, 41040, 855, 41040, 0)]
    [InlineData(@"IMP-84\INFORMACION GENERAL\imp 84.xlsm", "84", 32, 2184, 58324, 2184, 29162, 32)]
    [InlineData(@"IMP-244\Información general\IMP244.xlsm", "244", 6, 57, 2795, 57, 2655, 1)]
    [InlineData(@"IMP-188-189\INFORMACIÓN GENERAL\IMP 188-189Control de calidad.xlsm", "188", 16, 8672, 186636, 8672, 156492, 2)]
    [InlineData(@"IMP-88\INFORMACIÓN GENERAL\IMP-88 Control Calidad.xlsm", "88", 8, 839, 5034, 839, 5034, 0)]
    public void Read_IsolatedRealDataCopies_MatchesPythonReferenceParity(
        string relativePath,
        string expectedContainerId,
        int expectedItems,
        int expectedCalcCartons,
        int expectedCalcUnits,
        int expectedRepCartons,
        int expectedRepUnits,
        int expectedFlagged)
    {
        string source = Path.Combine(ProductionDataDir, relativePath);
        if (!File.Exists(source))
            return;

        string isolatedCopy = Path.Combine(_tempRoot, Path.GetFileName(source));
        File.Copy(source, isolatedCopy, overwrite: true);

        QcInspectionReport report = QcReader.Read(isolatedCopy, _testingGuard);

        Assert.Equal(expectedContainerId, report.ContainerId);
        Assert.Equal(expectedItems, report.Items.Count);
        Assert.Equal(expectedCalcCartons, report.TotalCalculatedCartons);
        Assert.Equal(expectedCalcUnits, report.TotalCalculatedUnits);
        Assert.Equal(expectedRepCartons, report.TotalReportedCartons);
        Assert.Equal(expectedRepUnits, report.TotalReportedUnits);
        Assert.Equal(expectedFlagged, report.FlaggedItems.Count);
    }

    private string CreateSyntheticQcWorkbook(string filename, bool includeTrailingBloat)
    {
        string path = Path.Combine(_tempRoot, filename);
        using SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        SheetData sheetData = InitializeSingleSheetWorkbook(doc, "CONTROL DATA");

        AppendSyntheticMetadataAndHeaders(sheetData);
        AppendSyntheticItemRows(sheetData);

        if (includeTrailingBloat)
        {
            AppendTrailingBloatAndUnreachableRow(sheetData);
        }

        doc.WorkbookPart!.Workbook!.Save();
        return path;
    }

    private static void AppendSyntheticMetadataAndHeaders(SheetData sheetData)
    {
        sheetData.AppendChild(BuildRow(1, ("A1", "CONTENEDOR"), ("B1", "TEST-999")));
        sheetData.AppendChild(BuildRow(2, ("A2", "RESPONSABLE:"), ("B2", "Analista Pruebas")));
        sheetData.AppendChild(BuildRow(3, ("A3", "NACIONALIZACIÓN:"), ("B3", "4100009999")));
        sheetData.AppendChild(BuildRow(6,
            ("A6", "TARIMA"), ("B6", "CATEGORIA"), ("C6", "MATERIAL"),
            ("D6", "DESCRIPCION"), ("E6", "UM"), ("F6", "CANTIDAD"),
            ("G6", "BULTOS ENTEROS"), ("H6", "UNIDADES POR BULTO"),
            ("I6", "BULTOS COMPLEMENTOS"), ("J6", "UNIDADES BULTOS COMPLEMENTOS"),
            ("K6", "TOTAL DE BULTOS"), ("L6", "TOTAL DE UNIDADES"), ("M6", "OBSERVACIONES")));
    }

    private static void AppendSyntheticItemRows(SheetData sheetData)
    {
        sheetData.AppendChild(BuildRow(7,
            ("A7", "1"), ("B7", "ROPA"), ("C7", "1001"), ("D7", "CAMISA"),
            ("E7", "UN"), ("F7", "600"), ("G7", "10"), ("H7", "60"),
            ("K7", "10"), ("L7", "600")));
        sheetData.AppendChild(BuildRow(8,
            ("A8", "2"), ("B8", "ROPA"), ("C8", "1002"), ("D8", "PANTALON"),
            ("E8", "UN"), ("F8", "1200"), ("G8", "6"), ("H8", "192"),
            ("J8", "48"), ("K8", "6"), ("L8", "1200")));
        sheetData.AppendChild(BuildRow(9,
            ("A9", "3"), ("B9", "CALZADO"), ("C9", "1003"), ("D9", "ZAPATO A"),
            ("E9", "UN"), ("F9", "300"), ("G9", "1"), ("H9", "300"),
            ("K9", "1"), ("L9", "300"), ("M9", "DOS MATERIALES POR BULTO")));
        sheetData.AppendChild(BuildRow(10,
            ("A10", "3"), ("B10", "CALZADO"), ("C10", "1004"), ("D10", "ZAPATO B"),
            ("E10", "UN"), ("F10", "300"), ("H10", "300"),
            ("K10", "0"), ("L10", "300"), ("M10", "DOS MATERIALES POR BULTO")));
        sheetData.AppendChild(BuildRow(11,
            ("A11", "VARIANTE SUELTA"), ("B11", "ROPA"), ("C11", "1005"), ("D11", "SUELTA"),
            ("E11", "UN"), ("F11", "10"), ("L11", "10")));
    }

    private static void AppendTrailingBloatAndUnreachableRow(SheetData sheetData)
    {
        for (uint r = 12; r <= 1050; r++)
        {
            var emptyRow = new Row { RowIndex = r };
            emptyRow.AppendChild(new Cell { CellReference = $"A{r}", StyleIndex = 1 });
            sheetData.AppendChild(emptyRow);
        }

        sheetData.AppendChild(BuildRow(1060,
            ("A1060", "99"), ("C1060", "9999"), ("D1060", "SHOULD NOT BE REACHED"),
            ("G1060", "50"), ("H1060", "10")));
    }

    private string CreateEdgeCaseQcWorkbook(string filename)
    {
        string path = Path.Combine(_tempRoot, filename);
        using SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        SheetData sheetData = InitializeSingleSheetWorkbook(doc, "CONTROL DATA");
        SharedStringTablePart sstPart = doc.WorkbookPart!.AddNewPart<SharedStringTablePart>();
        sstPart.SharedStringTable = BuildEdgeCaseSharedStrings();

        sheetData.AppendChild(BuildRow(1, ("A1", "CONTENEDOR"), ("B1", "EDGE-1")));
        sheetData.AppendChild(BuildRow(6,
            ("A6", "TARIMA"), ("B6", "MATERIAL PROVEEDOR"), ("C6", "MATERIAL"),
            ("E6", "DESCRIPCION"), ("G6", "BULTOS ENTEROS"), ("H6", "UNIDADES POR BULTO"),
            ("J6", "UNIDADES BULTOS COMPLEMENTOS"), ("L6", "BULTOS COMPLEMENTOS"),
            ("N6", "TOTAL PRK"), ("P6", "DIFERENCIA")));

        sheetData.AppendChild(BuildEdgeRow7());
        sheetData.AppendChild(BuildEdgeRow8());
        doc.WorkbookPart.Workbook!.Save();
        return path;
    }

    private static SharedStringTable BuildEdgeCaseSharedStrings()
    {
        var sst = new SharedStringTable();
        var richSi = new SharedStringItem();
        richSi.AppendChild(new Run(new Text("Rich ")));
        richSi.AppendChild(new Run(new Text("Text Desc")));
        richSi.AppendChild(new PhoneticRun(new Text("IGNORED_PHONETIC")));
        sst.AppendChild(richSi);
        return sst;
    }

    private static Row BuildEdgeRow7()
    {
        var row = new Row { RowIndex = 7 };
        row.AppendChild(CreateCell("A7", "1", CellValues.String));
        row.AppendChild(CreateCell("B7", "PROV-SKIP", CellValues.String));
        row.AppendChild(CreateCell("C7", "MAT-1001", CellValues.String));
        row.AppendChild(new Cell { CellReference = "E7", DataType = CellValues.SharedString, CellValue = new CellValue("0") });
        row.AppendChild(CreateCell("G7", "2.5", CellValues.Number));
        row.AppendChild(CreateCell("H7", "12", CellValues.Number));
        row.AppendChild(CreateCell("N7", "0", CellValues.Number));
        row.AppendChild(CreateCell("P7", "#VALUE!", CellValues.Error));
        return row;
    }

    private static Row BuildEdgeRow8()
    {
        var row = new Row { RowIndex = 8 };
        row.AppendChild(CreateCell("A8", "2", CellValues.String));
        row.AppendChild(CreateCell("C8", "MAT-1002", CellValues.String));
        row.AppendChild(new Cell { CellReference = "E8", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("Inline Desc")) });
        row.AppendChild(CreateCell("G8", "3.5", CellValues.Number));
        row.AppendChild(CreateCell("H8", "10", CellValues.Number));
        row.AppendChild(CreateCell("J8", "-", CellValues.String));
        row.AppendChild(CreateCell("N8", "5", CellValues.Number));
        return row;
    }

    private string CreateSimpleWorkbook(string filename, string[] sheetNames, bool hasVba)
    {
        string path = Path.Combine(_tempRoot, filename);
        SpreadsheetDocumentType type = Path.GetExtension(filename).Equals(".xlsm", StringComparison.OrdinalIgnoreCase)
            ? SpreadsheetDocumentType.MacroEnabledWorkbook
            : SpreadsheetDocumentType.Workbook;

        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(path, type))
        {
            WorkbookPart wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
            uint id = 1;
            foreach (string name in sheetNames)
            {
                WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
                wsPart.Worksheet = new Worksheet(new SheetData());
                sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = id++, Name = name });
            }
            wbPart.Workbook.Save();
        }

        if (hasVba)
        {
            using ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Update);
            ZipArchiveEntry entry = zip.CreateEntry("xl/vbaProject.bin");
            using StreamWriter writer = new(entry.Open());
            writer.Write("dummy-vba");
        }

        return path;
    }

    private static SheetData InitializeSingleSheetWorkbook(SpreadsheetDocument doc, string sheetName)
    {
        WorkbookPart wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();
        WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        wsPart.Worksheet = new Worksheet(sheetData);
        Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = sheetName });
        return sheetData;
    }

    private static Row BuildRow(uint rowIndex, params (string Ref, string Value)[] cells)
    {
        var row = new Row { RowIndex = rowIndex };
        foreach ((string cellRef, string value) in cells)
        {
            bool isNum = double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
            row.AppendChild(CreateCell(cellRef, value, isNum ? CellValues.Number : CellValues.String));
        }
        return row;
    }

    private static Cell CreateCell(string cellRef, string value, CellValues type) =>
        new() { CellReference = cellRef, DataType = type, CellValue = new CellValue(value) };

    private static bool TryCreateNtfsJunction(string junctionPath, string targetPath)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c mklink /J \"{junctionPath}\" \"{targetPath}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            });
            process?.WaitForExit(5000);
            return process?.ExitCode == 0 && Directory.Exists(junctionPath);
        }
        catch (Exception)
        {
            return false;
        }
    }
}
