using System.IO.Compression;
using System.Text;
using Acr.Core.Models;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Receiving;
using Acr.Infrastructure.Reporting;
using Acr.Infrastructure.Safety;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Tests.Infrastructure;

public sealed class ReceivingPipelineTests : IDisposable
{
    private const string ProductionDataDir = @"D:\Development\Software Development\3 Way Inventory Reconciliation\data";
    private const string ProductionDailyLogPath = @"D:\Development\Software Development\3 Way Inventory Reconciliation\Control_Diario_Operaciones_Jostin.xlsx";

    private readonly string _tempRoot;
    private readonly OperationalSafetyGuard _testingGuard;

    public ReceivingPipelineTests()
    {
        _tempRoot = Path.Combine(Path.GetTempPath(), $"AcrPipelineTests_{Guid.NewGuid():N}");
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
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    [Fact]
    public void Execute_SyntheticPipelineWithVbaAndRevuelto_ProducesExportJsonReportAndHistoryWithBackup()
    {
        string qcPath = CreateSyntheticQcWorkbook("syn_qc.xlsm");
        string tplPath = CreateSyntheticClosureTemplate("syn_tpl.xlsm");
        string outPath = Path.Combine(_tempRoot, "out", "syn_export.xlsm");
        string reportsDir = Path.Combine(_tempRoot, "reports");
        string historyPath = Path.Combine(_tempRoot, "history", "operation_history.json");

        ReceivingPipelineResult result = ReceivingPipeline.Execute(new ReceivingPipelineOptions
        {
            QcFilePath = qcPath,
            ClosureTemplatePath = tplPath,
            ExportOutputPath = outPath,
            ReportsDirectory = reportsDir,
            HistoryFilePath = historyPath,
            SafetyGuard = _testingGuard
        });

        Assert.Equal("IMP-777", result.ContainerId);
        Assert.Equal(1, result.Comparison.RevueltoPairsCount);
        Assert.Equal(12, result.Comparison.TotalShortageUnits);
        Assert.Equal(12, result.Comparison.TotalOverageUnits);
        Assert.NotNull(result.ExportResult);
        Assert.True(result.ExportResult!.Validation.IsValid);
        Assert.True(result.ExportResult.Validation.VbaProjectHashPreserved);
        Assert.Empty(result.ExportResult.Validation.ExporterIntroducedErrors);

        ProcessingReport reloaded = JsonProcessingReportStore.Load(result.SavedReportPath!, _testingGuard);
        Assert.Equal(result.Report!.ReportId, reloaded.ReportId);
        Assert.Equal(1, reloaded.RevueltoPairsCount);
        Assert.True(reloaded.ValidationResult!.VbaProjectSha256Preserved);

        OperationHistoryStore.Append(historyPath, new OperationHistoryEntry
        {
            Kind = OperationHistoryKind.Reconciliation,
            ContainerId = "IMP-778",
            Succeeded = true,
            SummaryText = "Second run"
        }, _testingGuard);

        Assert.True(File.Exists(historyPath + ".bak"));
        IReadOnlyList<OperationHistoryEntry> recent = OperationHistoryStore.GetRecent(historyPath, 2, _testingGuard);
        Assert.Equal(2, recent.Count);
        Assert.Equal("IMP-778", recent[0].ContainerId);
    }

    [Fact]
    public void SafetyModes_TestingRedirectsOutsideArtifactsToIsolatedTemp_AndDevelopmentBlocksProductionWrites()
    {
        string qcPath = CreateSyntheticQcWorkbook("safe_qc.xlsm");
        string tplPath = CreateSyntheticClosureTemplate("safe_tpl.xlsm");
        string outsideRoot = Path.Combine(Path.GetTempPath(), $"AcrOutsideTarget_{Guid.NewGuid():N}");
        string outsideExport = Path.Combine(outsideRoot, "outside_export.xlsm");
        string outsideReports = Path.Combine(outsideRoot, "reports");
        string outsideHistory = Path.Combine(outsideRoot, "operation_history.json");

        ReceivingPipelineResult result = ReceivingPipeline.Execute(new ReceivingPipelineOptions
        {
            QcFilePath = qcPath,
            ClosureTemplatePath = tplPath,
            ExportOutputPath = outsideExport,
            ReportsDirectory = outsideReports,
            HistoryFilePath = outsideHistory,
            SafetyGuard = _testingGuard
        });

        Assert.StartsWith(_testingGuard.IsolatedTempRoot, result.ExportResult!.OutputPath, StringComparison.OrdinalIgnoreCase);
        Assert.StartsWith(_testingGuard.IsolatedTempRoot, result.SavedReportPath!, StringComparison.OrdinalIgnoreCase);
        Assert.StartsWith(_testingGuard.IsolatedTempRoot, result.SavedHistoryPath!, StringComparison.OrdinalIgnoreCase);
        Assert.False(Directory.Exists(outsideRoot));

        var devGuard = new OperationalSafetyGuard(OperationalSafetyMode.Development, ProductionDataDir, ProductionDailyLogPath, _tempRoot);
        Assert.Throws<OperationalSafetyException>(
            () => devGuard.EnsureCanWrite(Path.Combine(ProductionDataDir, "blocked.xlsm"), OperationalActionKind.WriteToNewExportPath));
    }

    [Fact]
    public void Execute_RealPair_Imp100AndImp327_ExactCleanParityAndVbaPreserved()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        RunAndAssertRealPair(
            @"IMP-100\Informacion General\IMP 100 Control Calidad.xlsm",
            @"IMP-100\IMP-100.xlsm",
            expectedDeclaredItems: 18,
            expectedDeclaredUnits: 990,
            expectedPhysicalUnits: 990,
            expectedDeclaredCartons: 990,
            expectedPhysicalCartons: 990,
            expectedDiscrepant: 0);

        RunAndAssertRealPair(
            @"IMP-327\Informacion General\cierre imp 327 Control de calidad.xlsm",
            @"IMP-327\IMP-327.xlsm",
            expectedDeclaredItems: 6,
            expectedDeclaredUnits: 4014,
            expectedPhysicalUnits: 4014,
            expectedDeclaredCartons: 669,
            expectedPhysicalCartons: 669,
            expectedDiscrepant: 0);
    }

    [Fact]
    public void Execute_RealPair_Imp103_FiltersChildSkusDetectsRevueltoAndUsesColG()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        ReceivingPipelineResult result = RunAndAssertRealPair(
            @"IMP-103\Informacion General\imp 103 control calidad corregirdo.xlsm",
            @"IMP-103\IMP-103.xlsm",
            expectedDeclaredItems: 89,
            expectedDeclaredUnits: 102711,
            expectedPhysicalUnits: 102854,
            expectedDeclaredCartons: 1038,
            expectedPhysicalCartons: 1038,
            expectedDiscrepant: 16);

        Assert.Equal(267, result.QcReport.Items.Count);
        Assert.Equal(87, result.PhysicalItems.Count);
        Assert.Equal(12, result.Comparison.TotalShortageUnits);
        Assert.Equal(155, result.Comparison.TotalOverageUnits);
        Assert.Equal(1, result.Comparison.RevueltoPairsCount);
    }

    [Fact]
    public void Execute_RealPair_Imp207_EnforcesCardinalRuleAndDetectsUnexpectedItem()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        ReceivingPipelineResult result = RunAndAssertRealPair(
            @"IMP-207 (NCR)\Informacion General\imp 207 control calidad.xlsm",
            @"IMP-207 (NCR)\IMP-207 NCR.xlsm",
            expectedDeclaredItems: 35,
            expectedDeclaredUnits: 10642,
            expectedPhysicalUnits: 10382,
            expectedDeclaredCartons: 789,
            expectedPhysicalCartons: 783,
            expectedDiscrepant: 8);

        Assert.Equal(14, result.PhysicalItems.Count);
        Assert.Equal(433, result.Comparison.TotalShortageUnits);
        Assert.Equal(173, result.Comparison.TotalOverageUnits);
        Assert.Contains(result.Comparison.Items, i => i.MaterialCode == "5033283" && i.DiscrepancyType == DiscrepancyType.UnexpectedItem);
    }

    [Fact]
    public void Execute_RealPair_Imp187_ProcessesRowsBeyond299AndPreservesVba()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        ReceivingPipelineResult result = RunAndAssertRealPair(
            @"IMP-187\Informacion General\IMP 187 control calidad.xlsm",
            @"IMP-187\IMP-187.xlsm",
            expectedDeclaredItems: 231,
            expectedDeclaredUnits: 60877,
            expectedPhysicalUnits: 61188,
            expectedDeclaredCartons: 986,
            expectedPhysicalCartons: 986,
            expectedDiscrepant: null);

        Assert.Equal(163, result.PhysicalItems.Count);
        Assert.Equal(231, result.ExportResult!.MatchedRowsCount);
    }

    [Fact]
    public void ReportStore_UnsafeReportIdAndLongContainerId_SanitizesBoundedWithinReportsDirectory()
    {
        string reportsDir = Path.Combine(_tempRoot, "reports_sanitized");
        Directory.CreateDirectory(reportsDir);
        string[] maliciousReportIds = [@"\..\..\x", "/../../y", "a/b/c/d", "evil:ads"];

        foreach (string badId in maliciousReportIds)
        {
            var report = new ProcessingReport
            {
                ReportId = badId,
                ContainerId = new string('X', 300)
            };

            string savedPath = JsonProcessingReportStore.Save(report, reportsDir, _testingGuard);

            Assert.True(File.Exists(savedPath));
            Assert.Equal(Path.GetFullPath(reportsDir), Path.GetDirectoryName(Path.GetFullPath(savedPath)));
            Assert.True(Path.GetFileName(savedPath).Length < 120);
        }
    }

    [Fact]
    public void ReportStore_LoadAll_IgnoresCoLocatedOperationHistoryJson()
    {
        string reportsDir = Path.Combine(_tempRoot, "reports_colocated");
        Directory.CreateDirectory(reportsDir);
        var report = new ProcessingReport { ContainerId = "IMP-555" };
        JsonProcessingReportStore.Save(report, reportsDir, _testingGuard);

        string historyPath = Path.Combine(reportsDir, "operation_history.json");
        File.WriteAllText(historyPath, "[]", Encoding.UTF8);
        File.WriteAllText(Path.Combine(reportsDir, "unrelated.json"), "{\"foo\":1}", Encoding.UTF8);

        IReadOnlyList<ProcessingReport> loaded = JsonProcessingReportStore.LoadAll(reportsDir, _testingGuard);

        Assert.Single(loaded);
        Assert.Equal("IMP-555", loaded[0].ContainerId);
    }

    [Fact]
    public void OperationHistoryStore_ParallelAppend_TenThreadsPersistAllEntriesWithoutExceptions()
    {
        string historyPath = Path.Combine(_tempRoot, "parallel_history", "operation_history.json");
        var errors = new System.Collections.Concurrent.ConcurrentBag<Exception>();

        Parallel.For(0, 10, i =>
        {
            try
            {
                OperationHistoryStore.Append(historyPath, new OperationHistoryEntry
                {
                    ContainerId = $"IMP-PAR-{i:D2}",
                    Succeeded = true,
                    SummaryText = $"Thread {i}"
                }, _testingGuard);
            }
            catch (Exception ex)
            {
                errors.Add(ex);
            }
        });

        Assert.Empty(errors);
        IReadOnlyList<OperationHistoryEntry> all = OperationHistoryStore.LoadAll(historyPath, _testingGuard);
        Assert.Equal(10, all.Count);
    }

    [Fact]
    public void Execute_CompoundPackagingAndZeroTotal_FormatsCorrectlyInBultosAndTiras()
    {
        string qcPath = Path.Combine(_tempRoot, "compound_qc.xlsm");
        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(qcPath, SpreadsheetDocumentType.MacroEnabledWorkbook))
        {
            WorkbookPart wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
            WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "CONTROL DATA" });

            SheetData sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;
            WriteInlineRow(sd, 2, [(0, "CONTENEDOR"), (1, "IMP-888")]);
            WriteInlineRow(sd, 6, [
                (0, "TARIMA"), (1, "CATEGORIA"), (2, "MATERIAL PROVEEDOR"), (3, "MATERIAL"),
                (4, "UM"), (5, "BULTOS ENTEROS"), (6, "UNIDADES POR BULTO"),
                (7, "BULTOS COMPLEMENTOS"), (8, "UNIDADES BULTOS COMPLEMENTOS"),
                (9, "TOTAL DE BULTOS"), (10, "TOTAL DE UNIDADES")
            ]);
            // Item 1: 8 enteros de 48 + 1 complemento de 36 = 9 bultos, 420 unidades
            WriteInlineRow(sd, 7, [
                (0, "1"), (1, "NAVIDAD"), (2, "V888"), (3, "5033280"),
                (4, "UN"), (5, "8"), (6, "48"), (7, "1"), (8, "36"), (9, "9"), (10, "420")
            ]);
            // Item 2: Material en cero en QC (0 unidades, 0 bultos)
            WriteInlineRow(sd, 8, [
                (0, "2"), (1, "NAVIDAD"), (2, "V888"), (3, "5033282"),
                (4, "UN"), (5, "0"), (6, "0"), (7, "0"), (8, "0"), (9, "0"), (10, "0")
            ]);
            // Item 3: 8 enteros de 48 + 0 bultos complementos y 36 unidades sueltas = 8 bultos, 420 unidades ("8*48 +36")
            WriteInlineRow(sd, 9, [
                (0, "3"), (1, "NAVIDAD"), (2, "V888"), (3, "5033285"),
                (4, "UN"), (5, "8"), (6, "48"), (7, "0"), (8, "36"), (9, "8"), (10, "420")
            ]);
            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
        }

        string tplPath = Path.Combine(_tempRoot, "compound_tpl.xlsm");
        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(tplPath, SpreadsheetDocumentType.MacroEnabledWorkbook))
        {
            WorkbookPart wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
            WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "CIERRE RECEP" });

            SheetData sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;
            WriteInlineRow(sd, 1, [(0, "CONTENEDOR"), (1, "IMP-888")]);
            WriteInlineRow(sd, 12, [(2, "Categoría"), (3, "MAT. PROVEEDOR"), (4, "MATERIAL"), (6, "CANTIDAD PEDIDO 41"), (48, "Bultos según Paking List")]);
            WriteInlineRow(sd, 13, [(2, "NAVIDAD"), (3, "V888"), (4, "5033280"), (6, "420"), (48, "9")]);
            WriteInlineRow(sd, 14, [(2, "NAVIDAD"), (3, "V888"), (4, "5033282"), (6, "0"), (48, "0")]);
            WriteInlineRow(sd, 15, [(2, "NAVIDAD"), (3, "V888"), (4, "5033285"), (6, "420"), (48, "8")]);

            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
        }

        ReceivingPipelineResult result = ReceivingPipeline.Execute(new ReceivingPipelineOptions
        {
            QcFilePath = qcPath,
            ClosureTemplatePath = tplPath,
            SafetyGuard = _testingGuard
        });

        // Con ExcludeZeroTotalFromStripsAndBultos activo por defecto, el material en cero no aparece en Bultos ni en las Tiras (quedan 2 activos)
        Assert.Equal(2, result.BultosItems.Count);

        // Item 1: 5033280 (1 bulto complemento -> "8*48 1*36")
        BultosPresentationItem bulto1 = result.BultosItems[0];
        Assert.Equal("5033280", bulto1.MaterialCode);
        Assert.Equal("8*48 1*36", bulto1.PackagingExpression);
        Assert.Equal(8, bulto1.FullCartons);
        Assert.Equal(48, bulto1.UnitsPerFullCarton);
        Assert.Equal(1, bulto1.ComplementCartons);
        Assert.Equal(36, bulto1.UnitsInComplement);
        Assert.False(bulto1.IsZeroTotal);

        // Item 3: 5033285 (sin bultos complementos pero con 36 unidades sueltas -> "8*48 +36")
        BultosPresentationItem bulto3 = result.BultosItems[1];
        Assert.Equal("5033285", bulto3.MaterialCode);
        Assert.Equal("8*48 +36", bulto3.PackagingExpression);
        Assert.Equal(8, bulto3.FullCartons);
        Assert.Equal(48, bulto3.UnitsPerFullCarton);
        Assert.Equal(0, bulto3.ComplementCartons);
        Assert.Equal(36, bulto3.UnitsInComplement);
        Assert.False(bulto3.IsZeroTotal);

        // Verificación de Tira 2 (AW:AY) y Tira 1 (AS:AT):
        Assert.Contains("9\t420\t8*48 1*36", result.Tira2AwAyTsv);
        Assert.Contains("8\t420\t8*48 +36", result.Tira2AwAyTsv);
        Assert.Contains("9\t420", result.Tira1AsAtTsv);
        Assert.Contains("8\t420", result.Tira1AsAtTsv);
        Assert.DoesNotContain("0\t0\t0", result.Tira2AwAyTsv);
        Assert.DoesNotContain("0\t0", result.Tira1AsAtTsv);

        // Si se desactiva explícitamente ExcludeZeroTotalFromStripsAndBultos, se incluyen los 3
        ReceivingPipelineResult resultWithZeros = ReceivingPipeline.Execute(new ReceivingPipelineOptions
        {
            QcFilePath = qcPath,
            ClosureTemplatePath = tplPath,
            ExcludeZeroTotalFromStripsAndBultos = false,
            SafetyGuard = _testingGuard
        });
        Assert.Equal(3, resultWithZeros.BultosItems.Count);
        Assert.Contains("0\t0\t0", resultWithZeros.Tira2AwAyTsv);
    }

    private ReceivingPipelineResult RunAndAssertRealPair(
        string qcRelativePath,
        string closureRelativePath,
        int expectedDeclaredItems,
        int expectedDeclaredUnits,
        int expectedPhysicalUnits,
        int expectedDeclaredCartons,
        int expectedPhysicalCartons,
        int? expectedDiscrepant)
    {
        string qcSource = Path.Combine(ProductionDataDir, qcRelativePath);
        string closureSource = Path.Combine(ProductionDataDir, closureRelativePath);
        Assert.True(File.Exists(qcSource), $"Missing fixture: {qcSource}");
        Assert.True(File.Exists(closureSource), $"Missing fixture: {closureSource}");

        string qcCopy = CopyToTemp(qcSource);
        string tplCopy = CopyToTemp(closureSource);
        string exportPath = Path.Combine(_tempRoot, $"export_{Path.GetFileName(closureSource)}");

        ReceivingPipelineResult result = ReceivingPipeline.Execute(new ReceivingPipelineOptions
        {
            QcFilePath = qcCopy,
            ClosureTemplatePath = tplCopy,
            ExportOutputPath = exportPath,
            ReportsDirectory = Path.Combine(_tempRoot, "reports"),
            HistoryFilePath = Path.Combine(_tempRoot, "history.json"),
            SafetyGuard = _testingGuard
        });

        Assert.Equal(expectedDeclaredItems, result.DeclaredItems.Count);
        Assert.Equal(expectedDeclaredUnits, result.Comparison.TotalDeclaredUnits);
        Assert.Equal(expectedPhysicalUnits, result.Comparison.TotalPhysicalUnits);
        Assert.Equal(expectedDeclaredCartons, result.Comparison.TotalDeclaredCartons);
        Assert.Equal(expectedPhysicalCartons, result.Comparison.TotalPhysicalCartons);
        if (expectedDiscrepant.HasValue)
            Assert.Equal(expectedDiscrepant.Value, result.Comparison.DiscrepantMaterialsCount);

        Assert.True(result.ExportResult!.Validation.VbaProjectHashPreserved);
        Assert.Empty(result.ExportResult.Validation.ExporterIntroducedErrors);
        Assert.True(result.ExportResult.Validation.IsValid);
        return result;
    }

    private string CopyToTemp(string sourcePath)
    {
        string dest = Path.Combine(_tempRoot, $"{Guid.NewGuid():N}_{Path.GetFileName(sourcePath)}");
        File.Copy(sourcePath, dest, overwrite: true);
        return dest;
    }

    private string CreateSyntheticQcWorkbook(string fileName)
    {
        string path = Path.Combine(_tempRoot, fileName);
        using SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.MacroEnabledWorkbook);
        WorkbookPart wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();
        Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
        WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = new Worksheet(new SheetData());
        sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "CONTROL DATA" });

        SheetData sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;
        WriteSyntheticQcRows(sd);
        wsPart.Worksheet.Save();
        wbPart.Workbook.Save();
        return path;
    }

    private static void WriteSyntheticQcRows(SheetData sd)
    {
        WriteInlineRow(sd, 2, [(0, "CONTENEDOR"), (1, "IMP-777")]);
        WriteInlineRow(sd, 6, [(0, "TARIMA"), (1, "CATEGORIA"), (2, "MATERIAL PROVEEDOR"), (3, "MATERIAL"), (4, "UM"), (5, "BULTOS ENTEROS"), (6, "UNIDADES POR BULTO"), (7, "TOTAL DE BULTOS"), (8, "TOTAL DE UNIDADES")]);
        WriteInlineRow(sd, 7, [(0, "1"), (1, "CABALLERO"), (2, "V777"), (3, "S1"), (4, "UN"), (5, "4"), (6, "22"), (7, "4"), (8, "88")]);
        WriteInlineRow(sd, 8, [(0, "2"), (1, "CABALLERO"), (2, "V777"), (3, "O1"), (4, "UN"), (5, "3"), (6, "20"), (7, "3"), (8, "60")]);
    }

    private string CreateSyntheticClosureTemplate(string fileName)
    {
        string path = Path.Combine(_tempRoot, fileName);
        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.MacroEnabledWorkbook))
        {
            WorkbookPart wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
            WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet(new SheetData());
            sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "CIERRE RECEP" });

            SheetData sd = wsPart.Worksheet.GetFirstChild<SheetData>()!;
            WriteInlineRow(sd, 1, [(0, "CONTENEDOR"), (1, "IMP-777")]);
            WriteInlineRow(sd, 12, [(2, "Categoría"), (3, "MAT. PROVEEDOR"), (4, "MATERIAL"), (6, "CANTIDAD PEDIDO 41"), (48, "Bultos según Paking List")]);
            WriteInlineRow(sd, 13, [(2, "CABALLERO"), (3, "V777"), (4, "S1"), (6, "100"), (48, "5")]);
            WriteInlineRow(sd, 14, [(2, "CABALLERO"), (3, "V777"), (4, "O1"), (6, "48"), (48, "2")]);

            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
        }

        using var archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.CreateEntry("xl/vbaProject.bin");
        using Stream s = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes("SYNTHETIC_PIPELINE_VBA");
        s.Write(bytes, 0, bytes.Length);
        return path;
    }

    private static void WriteInlineRow(SheetData sd, int rowIdx, (int ColIdx, string Val)[] values)
    {
        var row = new Row { RowIndex = checked((uint)rowIdx) };
        sd.AppendChild(row);
        foreach ((int colIdx, string val) in values.OrderBy(v => v.ColIdx))
        {
            var cell = new Cell { CellReference = BuildCellRef(colIdx, rowIdx) };
            if (int.TryParse(val, out _))
                cell.CellValue = new CellValue(val);
            else
            {
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(new Text(val));
            }
            row.AppendChild(cell);
        }
    }

    private static string BuildCellRef(int colIdx, int rowIdx)
    {
        int colNum = colIdx + 1;
        var sb = new StringBuilder(3);
        while (colNum > 0)
        {
            int rem = (colNum - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            colNum = (colNum - 1) / 26;
        }
        return $"{sb}{rowIdx}";
    }
}
