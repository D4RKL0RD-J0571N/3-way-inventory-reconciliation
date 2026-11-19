using System.IO.Compression;
using System.Text;
using Acr.Core.Models;
using Acr.Core.Services;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Safety;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Tests.Infrastructure;

public sealed class ClosureExporterMacroAndFormulaTests : IDisposable
{
    private const string ProductionDataDir = @"D:\Development\Software Development\3 Way Inventory Reconciliation\data";
    private const string ProductionDailyLogPath = @"D:\Development\Software Development\3 Way Inventory Reconciliation\Control_Diario_Operaciones_Jostin.xlsx";

    private readonly string _tempRoot;
    private readonly OperationalSafetyGuard _testingGuard;

    public ClosureExporterMacroAndFormulaTests()
    {
        _tempRoot = Path.Combine(Path.GetTempPath(), $"AcrClosureTests_{Guid.NewGuid():N}");
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
    public void Validate_MissingFile_ReturnsInvalidResult()
    {
        string missingPath = Path.Combine(_tempRoot, "missing_closure.xlsm");
        PrePrintValidationResult result = FormulaIntegrityAuditor.ValidateClosureDocument(missingPath, guard: _testingGuard);

        Assert.False(result.IsValid);
        Assert.NotEmpty(result.Errors);
    }

    [Fact]
    public void Validate_CorruptFormulaAndLiteralErrorTokens_DetectsAllForms()
    {
        string path = CreateSyntheticClosureWorkbook("corrupt_forms.xlsm", includeVba: true, includeBultos: true, populateErrors: true);

        IReadOnlyList<FormulaCellAudit> audits = FormulaIntegrityAuditor.ScanWorkbook(path, _testingGuard);
        PrePrintValidationResult result = FormulaIntegrityAuditor.ValidateClosureDocument(path, guard: _testingGuard);

        Assert.False(result.IsValid);
        Assert.True(result.Errors.Count >= 5);
        Assert.Contains(audits, a => a.CellReference == "D5" && a.IsError && a.ErrorToken == "#REF!");
        Assert.Contains(audits, a => a.CellReference == "B7" && a.IsError && a.ErrorToken == "#REF!");
        Assert.Contains(audits, a => a.CellReference == "AU14" && a.IsError && a.ErrorToken == "#REF!");
        Assert.Contains(audits, a => a.CellReference == "AU15" && a.IsError && a.ErrorToken == "#REF!");
        Assert.Contains(audits, a => a.CellReference == "AY29" && !a.HasFormula && a.IsError && a.ErrorToken == "#N/A");
        Assert.Contains(audits, a => a.CellReference == "V14" && a.IsError && a.ErrorToken == "#VALUE!");
    }

    [Fact]
    public void Validate_GuardedOptionalVlookup_ClassifiedAsLegitimateOptionalNa()
    {
        string path = CreateVlookupClassificationWorkbook("vlookup_audit.xlsx");

        PrePrintValidationResult result = FormulaIntegrityAuditor.ValidateClosureDocument(
            path,
            sourceTemplatePath: path,
            guard: _testingGuard);

        Assert.True(result.IsValid);
        Assert.Empty(result.ExporterIntroducedErrors);
        Assert.Contains(result.LegitimateOptionalNas, a => a.CellReference == "B13");
        Assert.Contains(result.LegitimateOptionalNas, a => a.CellReference == "C13");
        Assert.Contains(result.PreExistingSourceErrors, a => a.CellReference == "AT16" && a.ErrorToken == "#N/A");
    }

    [Fact]
    public void Export_SyntheticTemplate_PopulatesAllColumnsRevueltosAndPreservesXmlOrder()
    {
        string template = CreateSyntheticClosureWorkbook("template_rev.xlsm", includeVba: true, includeBultos: true, populateErrors: false);
        string output = Path.Combine(_tempRoot, "exported_rev.xlsm");
        ComparisonSummary summary = BuildSyntheticRevueltoSummary();

        ClosureExportResult result = ClosureExporter.Export(template, output, summary, guard: _testingGuard);

        Assert.True(result.Validation.IsValid);
        Assert.True(result.Validation.VbaProjectHashPreserved);
        Assert.Equal(result.Validation.SourceHashes!.VbaProjectSha256, result.Validation.OutputHashes!.VbaProjectSha256);
        Assert.Empty(result.Validation.ExporterIntroducedErrors);
        Assert.Equal(2, result.MatchedRowsCount);
        AssertExportedRevueltoCellsAndXmlOrder(output);
    }

    [Fact]
    public void Export_RowsBeyond299And612_PopulatesDynamicRowsWithoutTruncation()
    {
        string template = CreateLargeRowClosureWorkbook("large_620.xlsm", lastRow: 620);
        string output = Path.Combine(_tempRoot, "large_620_out.xlsm");
        var declared = new List<DeclaredItem>
        {
            new() { MaterialCode = "100615", DeclaredUnits = 120, DeclaredCartons = 5, SourceRow = 615 }
        };
        var physical = new List<PhysicalItem>
        {
            new() { MaterialCode = "100615", PhysicalUnits = 120, PhysicalCartons = 5, SourceRow = 615 }
        };
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(declared, physical);

        ClosureExportResult result = ClosureExporter.Export(template, output, summary, guard: _testingGuard);

        Assert.True(result.Validation.IsValid);
        Assert.Equal(1, result.MatchedRowsCount);
        Assert.Equal(120, result.Validation.TotalPhysicalUnits);
        Assert.Equal(5, result.Validation.TotalPhysicalCartons);
    }

    [Fact]
    public void Export_BultoCero_ZeroPhysicalAndZeroDeclared_WritesZeroNeverPhantomOne()
    {
        string template = CreateSyntheticClosureWorkbook("bulto_cero.xlsm", includeVba: true, includeBultos: false, populateErrors: false);
        string output = Path.Combine(_tempRoot, "bulto_cero_out.xlsm");
        var declared = new List<DeclaredItem>
        {
            new() { MaterialCode = "1419038", DeclaredUnits = 0, DeclaredCartons = 0, SourceRow = 13 }
        };
        var physical = new List<PhysicalItem>();
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(declared, physical);

        ClosureExportResult result = ClosureExporter.Export(template, output, summary, guard: _testingGuard);
        Assert.True(result.Validation.IsValid);

        Dictionary<string, string> row13 = ReadRowCells(output, "CIERRE RECEP", 13);
        Assert.Equal("0", row13["AS13"]);
        Assert.Equal("0", row13["AT13"]);
        Assert.Equal("0", row13["AW13"]);
        Assert.Equal("0", row13["AX13"]);
        Assert.Equal("0", row13["AY13"]);
    }

    [Fact]
    public void Export_BultosSyncAndTrailingCleanup_ClearsStaleTrailingCellsPreservingTableAndPivot()
    {
        string template = CreateSyntheticClosureWorkbook("bultos_cleanup.xlsm", includeVba: true, includeBultos: true, populateErrors: false);
        string output = Path.Combine(_tempRoot, "bultos_cleanup_out.xlsm");
        ComparisonSummary summary = BuildSyntheticRevueltoSummary();

        ClosureExportResult result = ClosureExporter.Export(template, output, summary, guard: _testingGuard);

        Assert.True(result.Validation.IsValid);
        Assert.Equal(2, result.BultosSyncedRowsCount);
        Assert.True(result.BultosTrailingClearedCount >= 1);

        Dictionary<string, string> staleTrailing = ReadRowCells(output, "Bultos", 5);
        Assert.Equal("", staleTrailing.GetValueOrDefault("C5", ""));
        Assert.Equal("", staleTrailing.GetValueOrDefault("D5", ""));
    }

    [Fact]
    public void Export_IsolatedRealMacroTemplate_PreservesVbaSha256StylesDefinedNamesRelsAndZeroExporterErrors()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        string[] candidates =
        [
            Path.Combine(ProductionDataDir, "Macro para cierre hannia.xlsm"),
            Path.Combine(ProductionDataDir, "IMP-187", "IMP-187.xlsm"),
            Path.Combine(ProductionDataDir, "IMP-207 (NCR)", "IMP-207 NCR.xlsm")
        ];

        foreach (string source in candidates.Where(File.Exists))
        {
            string copy = CopyToTemp(source);
            string output = Path.Combine(_tempRoot, $"exported_{Path.GetFileName(source)}");
            ComparisonSummary summary = BuildSummaryForTemplate(copy);

            ClosureExportResult result = ClosureExporter.Export(copy, output, summary, guard: _testingGuard);

            Assert.True(result.Validation.VbaProjectHashPreserved);
            Assert.True(result.Validation.StylesPreserved);
            Assert.True(result.Validation.DefinedNamesPreserved);
            Assert.True(result.Validation.RelationshipsPreserved);
            Assert.Empty(result.Validation.ExporterIntroducedErrors);
            Assert.True(result.Validation.IsValid);
            AssertNoCalcChainAndFullCalcEnabled(output);
        }
    }

    [Fact]
    public void Audit_IsolatedHistoricalErrorClosures_DistinguishesPreExistingErrors()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        (string RelPath, string Sheet, string CellRef, string Token)[] expectedErrors =
        [
            (@"IMP-241\IMP-241.xlsm", "CIERRE RECEP", "D5", "#REF!"),
            (@"IMP-244\IMP-244.xlsm", "CIERRE RECEP", "AS13", "#REF!"),
            (@"IMP-256\IMP-256.xlsm", "CIERRE RECEP", "AY29", "#N/A"),
            (@"IMP-360\IMP-360.xlsm", "CIERRE RECEP", "V14", "#VALUE!"),
            (@"IMP-90\IMP-90.xlsm", "CIERRE RECEP", "AT16", "#N/A"),
            (@"IMP-225\IMP-225.xlsm", "Bultos", "G1", "#REF!")
        ];

        foreach ((string relPath, string sheet, string cellRef, string token) in expectedErrors)
        {
            string source = Path.Combine(ProductionDataDir, relPath);
            if (!File.Exists(source))
                continue;

            string copy = CopyToTemp(source);
            PrePrintValidationResult result = FormulaIntegrityAuditor.ValidateClosureDocument(
                copy,
                sourceTemplatePath: copy,
                guard: _testingGuard);

            Assert.Contains(
                result.PreExistingSourceErrors,
                e => e.SheetName == sheet && e.CellReference == cellRef && e.ErrorToken == token);
            Assert.Empty(result.ExporterIntroducedErrors);
        }
    }

    [Fact]
    public void Export_ProductionDataTarget_BlockedBySafetyGuard()
    {
        string template = CreateSyntheticClosureWorkbook("safe_tpl.xlsm", includeVba: false, includeBultos: false, populateErrors: false);
        string forbiddenTarget = Path.Combine(ProductionDataDir, "SHOULD_NEVER_WRITE.xlsm");
        ComparisonSummary summary = BuildSyntheticRevueltoSummary();

        Assert.Throws<OperationalSafetyException>(
            () => ClosureExporter.Export(template, forbiddenTarget, summary, guard: _testingGuard));
    }

    [Fact]
    public void Export_RealMalformedRelsWorkbooks_Imp81AndImp182_SucceedsCleanly()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        string[] relPaths = [@"IMP-81\IMP-81.xlsm", @"IMP-182\IMP-182 inprogress.xlsm"];
        foreach (string relPath in relPaths)
        {
            string source = Path.Combine(ProductionDataDir, relPath);
            if (!File.Exists(source))
                continue;

            string copy = CopyToTemp(source);
            string output = Path.Combine(_tempRoot, $"sanitized_out_{Path.GetFileName(source)}");
            Acr.Infrastructure.Receiving.ClosureTemplateData tplData =
                Acr.Infrastructure.Receiving.ClosureTemplateReader.ReadDeclaredItems(copy, _testingGuard);
            Assert.NotEmpty(tplData.DeclaredItems);

            ComparisonSummary summary = BuildSummaryForTemplate(copy);
            ClosureExportResult result = ClosureExporter.Export(copy, output, summary, guard: _testingGuard);

            Assert.True(result.Validation.IsValid);
            Assert.True(result.Validation.VbaProjectHashPreserved);
            Assert.True(result.Validation.StylesPreserved);
            Assert.True(result.Validation.DefinedNamesPreserved);
            Assert.True(result.Validation.RelationshipsPreserved);
            Assert.Empty(result.Validation.ExporterIntroducedErrors);
        }
    }

    [Fact]
    public void Validate_DroppedRequiredFormulaInH13OrAu13_FailsAsExporterIntroducedError()
    {
        string template = CreateSyntheticClosureWorkbook("base_tpl.xlsm", includeVba: true, includeBultos: false, populateErrors: false);
        string missingOutput = Path.Combine(_tempRoot, "missing_formula_out.xlsm");
        File.Copy(template, missingOutput, overwrite: true);

        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(missingOutput, true))
        {
            WorksheetPart wsPart = doc.WorkbookPart!.WorksheetParts.First();
            SheetData sd = wsPart.Worksheet!.GetFirstChild<SheetData>()!;
            Row row14 = sd.Elements<Row>().First(r => r.RowIndex?.Value == 14);
            Cell h14 = row14.Elements<Cell>().First(c => c.CellReference?.Value == "H14");
            h14.CellFormula = null;
            h14.CellValue = new CellValue("50");
            wsPart.Worksheet.Save();
        }

        PrePrintValidationResult result = FormulaIntegrityAuditor.ValidateClosureDocument(
            missingOutput,
            sourceTemplatePath: template,
            guard: _testingGuard);

        Assert.False(result.IsValid);
        Assert.Contains(result.ExporterIntroducedErrors, e => e.CellReference == "H14" && e.ErrorToken == "#MISSING_FORMULA!");
    }

    [Fact]
    public void Validate_GuardedIsnaVlookupWithCachedRefError_FailsValidation()
    {
        string path = CreateVlookupClassificationWorkbook("vlookup_cached_ref.xlsx");
        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, true))
        {
            WorksheetPart wsPart = doc.WorkbookPart!.WorksheetParts.First();
            SheetData sd = wsPart.Worksheet!.GetFirstChild<SheetData>()!;
            SetInlineCell(sd, 14, 1, "IF(ISNA(VLOOKUP(AM14,A$2:B$10,2,0))=TRUE,\"\",VLOOKUP(AM14,A$2:B$10,2,0))", isFormula: true, cached: "#REF!", isError: true);
            wsPart.Worksheet.Save();
        }

        PrePrintValidationResult result = FormulaIntegrityAuditor.ValidateClosureDocument(path, guard: _testingGuard);

        Assert.False(result.IsValid);
        Assert.DoesNotContain(result.LegitimateOptionalNas, a => a.CellReference == "B14");
        Assert.Contains(result.PreExistingSourceErrors, a => a.CellReference == "B14" && a.ErrorToken == "#REF!");
    }

    [Fact]
    public void Export_BrokenSharedFormulaMaster_MaterializesFollowersWithoutOrphans()
    {
        string template = CreateBrokenSharedMasterWorkbook("shared_master_tpl.xlsm");
        string output = Path.Combine(_tempRoot, "shared_master_out.xlsm");
        var declared = new List<DeclaredItem> { new() { MaterialCode = "1419038", DeclaredUnits = 100, DeclaredCartons = 5, SourceRow = 13 } };
        var physical = new List<PhysicalItem> { new() { MaterialCode = "1419038", PhysicalUnits = 100, PhysicalCartons = 5, SourceRow = 13 } };
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(declared, physical);

        ClosureExportResult result = ClosureExporter.Export(template, output, summary, guard: _testingGuard);

        Assert.Equal(1, result.MatchedRowsCount);
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(output, false);
        WorksheetPart wsPart = doc.WorkbookPart!.WorksheetParts.First();
        Row row14 = wsPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().First(r => r.RowIndex?.Value == 14);
        Cell h14 = row14.Elements<Cell>().First(c => c.CellReference?.Value == "H14");
        Assert.NotNull(h14.CellFormula);
        Assert.Null(h14.CellFormula!.FormulaType);
        Assert.Null(h14.CellFormula.SharedIndex);
        Assert.Equal("+G14-#REF!", h14.CellFormula.Text);
    }

    [Fact]
    public void Export_BultosRowAtOrPastTotalsRow_EmitsExplicitTruncationWarning()
    {
        string template = CreateSyntheticClosureWorkbook("bultos_trunc.xlsm", includeVba: true, includeBultos: true, populateErrors: false);
        string output = Path.Combine(_tempRoot, "bultos_trunc_out.xlsm");
        var declared = new List<DeclaredItem>
        {
            new() { MaterialCode = "1419038", DeclaredUnits = 100, DeclaredCartons = 5, SourceRow = 13 },
            new() { MaterialCode = "1419040", DeclaredUnits = 50, DeclaredCartons = 2, SourceRow = 14 },
            new() { MaterialCode = "1419042", DeclaredUnits = 30, DeclaredCartons = 1, SourceRow = 16 }
        };
        var physical = new List<PhysicalItem>
        {
            new() { MaterialCode = "1419038", PhysicalUnits = 100, PhysicalCartons = 5, SourceRow = 13 },
            new() { MaterialCode = "1419040", PhysicalUnits = 50, PhysicalCartons = 2, SourceRow = 14 },
            new() { MaterialCode = "1419042", PhysicalUnits = 30, PhysicalCartons = 1, SourceRow = 16 }
        };
        using (SpreadsheetDocument doc = SpreadsheetDocument.Open(template, true))
        {
            WorkbookPart wbPart = doc.WorkbookPart!;
            Sheet sheet = wbPart.Workbook!.Sheets!.Elements<Sheet>()
                .First(s => s.Name!.Value!.Contains("CIERRE RECEP", StringComparison.OrdinalIgnoreCase));
            var cierrePart = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
            SheetData sd = cierrePart.Worksheet!.GetFirstChild<SheetData>()!;
            SetInlineCell(sd, 16, 4, "1419042", isFormula: false);
            SetInlineCell(sd, 16, 6, "30", isFormula: false);
            cierrePart.Worksheet.Save();
        }

        ClosureExportResult result = ClosureExporter.Export(template, output, DiscrepancyClassifier.CompareMerchandise(declared, physical), guard: _testingGuard);

        Assert.Contains(result.Validation.Warnings, w => w.Contains("BULTOS_SYNC_TRUNCATED", StringComparison.Ordinal));
    }

    [Fact]
    public void Validate_SubtotalTotalsRowAndSstNumericAsAt_DecodesAccuratelyWithoutDoubleCount()
    {
        string path = CreateSubtotalAndSstWorkbook("subtotal_sst.xlsx");

        PrePrintValidationResult result = FormulaIntegrityAuditor.ValidateClosureDocument(path, guard: _testingGuard);

        Assert.True(result.IsValid);
        Assert.Equal(5, result.TotalPhysicalCartons);
        Assert.Equal(100, result.TotalPhysicalUnits);
    }

    private string CopyToTemp(string sourcePath)
    {
        string dest = Path.Combine(_tempRoot, $"{Guid.NewGuid():N}_{Path.GetFileName(sourcePath)}");
        File.Copy(sourcePath, dest, overwrite: true);
        return dest;
    }

    private static ComparisonSummary BuildSummaryForTemplate(string templatePath)
    {
        Acr.Infrastructure.Receiving.ClosureTemplateData data =
            Acr.Infrastructure.Receiving.ClosureTemplateReader.ReadDeclaredItems(templatePath);
        List<PhysicalItem> physical = data.DeclaredItems
            .Select(d => new PhysicalItem
            {
                MaterialCode = d.MaterialCode,
                PhysicalUnits = d.DeclaredUnits,
                PhysicalCartons = d.DeclaredCartons,
                SourceRow = d.SourceRow
            })
            .ToList();
        return DiscrepancyClassifier.CompareMerchandise(data.DeclaredItems, physical);
    }

    private static ComparisonSummary BuildSyntheticRevueltoSummary()
    {
        var declared = new List<DeclaredItem>
        {
            new() { MaterialCode = "1419038", VendorMaterial = "V1", Category = "CAT", DeclaredUnits = 100, DeclaredCartons = 5, SourceRow = 13 },
            new() { MaterialCode = "1419040", VendorMaterial = "V1", Category = "CAT", DeclaredUnits = 50, DeclaredCartons = 2, SourceRow = 14 }
        };
        var physical = new List<PhysicalItem>
        {
            new() { MaterialCode = "1419038", VendorMaterial = "V1", Category = "CAT", PhysicalUnits = 88, PhysicalCartons = 4, SourceRow = 13 },
            new() { MaterialCode = "1419040", VendorMaterial = "V1", Category = "CAT", PhysicalUnits = 62, PhysicalCartons = 3, SourceRow = 14 }
        };
        return DiscrepancyClassifier.CompareMerchandise(declared, physical);
    }

    private string CreateSyntheticClosureWorkbook(string fileName, bool includeVba, bool includeBultos, bool populateErrors)
    {
        string path = Path.Combine(_tempRoot, fileName);
        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.MacroEnabledWorkbook))
        {
            WorkbookPart wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());

            WorksheetPart cierrePart = AddSheetWithPart(wbPart, sheets, "CIERRE RECEP", 1);
            BuildSyntheticCierreSheet(cierrePart, populateErrors);

            if (includeBultos)
            {
                WorksheetPart bultosPart = AddSheetWithPart(wbPart, sheets, "Bultos", 2);
                BuildSyntheticBultosSheet(bultosPart);
            }

            wbPart.Workbook.Save();
        }

        if (includeVba)
            InjectDummyVbaPart(path);

        return path;
    }

    private static WorksheetPart AddSheetWithPart(WorkbookPart wbPart, Sheets sheets, string name, uint id)
    {
        WorksheetPart part = wbPart.AddNewPart<WorksheetPart>();
        part.Worksheet = new Worksheet(new SheetData());
        sheets.AppendChild(new Sheet
        {
            Id = wbPart.GetIdOfPart(part),
            SheetId = id,
            Name = name
        });
        return part;
    }

    private static void BuildSyntheticCierreSheet(WorksheetPart part, bool populateErrors)
    {
        SheetData sd = part.Worksheet!.GetFirstChild<SheetData>()!;
        SetInlineCell(sd, 5, 1, "7", isFormula: false);
        SetInlineCell(sd, 6, 1, "7", isFormula: false);
        SetInlineCell(sd, 7, 1, "+B6-B5", isFormula: true, cached: populateErrors ? "#REF!" : null, isError: populateErrors);
        SetInlineCell(sd, 13, 4, "1419038.0", isFormula: false);
        SetInlineCell(sd, 13, 6, "100", isFormula: false);
        SetInlineCell(sd, 14, 4, "1419040", isFormula: false);
        SetInlineCell(sd, 14, 6, "50", isFormula: false);
        SetInlineCell(sd, 14, 7, "+G14-J14-N14-O14+L14+P14+Q14", isFormula: true);

        if (populateErrors)
            AppendCorruptTaxonomyCells(sd);

        part.Worksheet.Save();
    }

    private static void AppendCorruptTaxonomyCells(SheetData sd)
    {
        SetInlineCell(sd, 5, 3, "IFERROR(#REF!/B5,\"\")", isFormula: true);
        Row r14 = GetOrCreateRow(sd, 14);
        Cell au14 = GetOrCreateCell(r14, 46, 14);
        au14.CellFormula = new CellFormula("+AT14-H14") { FormulaType = CellFormulaValues.Shared, Reference = "AU14:AU15", SharedIndex = 17 };
        au14.DataType = CellValues.Error;
        au14.CellValue = new CellValue("#REF!");

        Row r15 = GetOrCreateRow(sd, 15);
        Cell au15 = GetOrCreateCell(r15, 46, 15);
        au15.CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 17 };
        au15.DataType = CellValues.Error;
        au15.CellValue = new CellValue("#REF!");

        SetInlineCell(sd, 14, 21, "+U14-I14", isFormula: true, cached: "#VALUE!", isError: true);
        Row r29 = GetOrCreateRow(sd, 29);
        Cell ay29 = GetOrCreateCell(r29, 50, 29);
        ay29.DataType = CellValues.Error;
        ay29.CellValue = new CellValue("#N/A");
    }

    private static void BuildSyntheticBultosSheet(WorksheetPart part)
    {
        SheetData sd = part.Worksheet!.GetFirstChild<SheetData>()!;
        SetInlineCell(sd, 3, 2, "'CIERRE RECEP'!AS13", isFormula: true);
        SetInlineCell(sd, 5, 2, "999", isFormula: false);
        SetInlineCell(sd, 5, 3, "9999", isFormula: false);
        SetInlineCell(sd, 6, 2, "SUM(C3:C5)", isFormula: true);
        part.Worksheet.Save();
    }

    private string CreateVlookupClassificationWorkbook(string fileName)
    {
        string path = Path.Combine(_tempRoot, fileName);
        using SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        WorkbookPart wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();
        Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
        WorksheetPart wsPart = AddSheetWithPart(wbPart, sheets, "CIERRE RECEP", 1);
        SheetData sd = wsPart.Worksheet!.GetFirstChild<SheetData>()!;

        SetInlineCell(sd, 13, 1, "IF(ISNA(VLOOKUP(AM13,A$2:B$10,2,0))=TRUE,\"\",VLOOKUP(AM13,A$2:B$10,2,0))", isFormula: true, cached: "");
        SetInlineCell(sd, 13, 2, "IFNA(VLOOKUP(AO13,F$2:G$10,2,0),\"\")", isFormula: true, cached: "#N/A", isError: true);
        SetInlineCell(sd, 16, 45, "VLOOKUP(AR16,[2]Hoja1!A$1:C$5,3,0)", isFormula: true, cached: "#N/A", isError: true);
        wsPart.Worksheet.Save();
        wbPart.Workbook.Save();
        return path;
    }

    private string CreateLargeRowClosureWorkbook(string fileName, int lastRow)
    {
        string path = Path.Combine(_tempRoot, fileName);
        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.MacroEnabledWorkbook))
        {
            WorkbookPart wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
            WorksheetPart wsPart = AddSheetWithPart(wbPart, sheets, "CIERRE RECEP", 1);
            SheetData sd = wsPart.Worksheet!.GetFirstChild<SheetData>()!;

            SetInlineCell(sd, 615, 4, "100615", isFormula: false);
            SetInlineCell(sd, 615, 6, "120", isFormula: false);
            SetInlineCell(sd, lastRow, 4, "TOTAL", isFormula: false);
            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
        }
        InjectDummyVbaPart(path);
        return path;
    }

    private static void SetInlineCell(
        SheetData sd,
        int rowIdx,
        int colIdx,
        string content,
        bool isFormula,
        string? cached = null,
        bool isError = false)
    {
        Row row = GetOrCreateRow(sd, rowIdx);
        Cell cell = GetOrCreateCell(row, colIdx, rowIdx);
        if (isFormula)
        {
            cell.CellFormula = new CellFormula(content);
            if (isError)
                cell.DataType = CellValues.Error;
            if (cached is not null)
                cell.CellValue = new CellValue(cached);
        }
        else
        {
            cell.CellValue = new CellValue(content);
        }
    }

    private static Row GetOrCreateRow(SheetData sd, int rowIdx)
    {
        uint target = checked((uint)rowIdx);
        Row? existing = sd.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == target);
        if (existing is not null)
            return existing;
        Row? before = sd.Elements<Row>().FirstOrDefault(r => (r.RowIndex?.Value ?? 0) > target);
        var created = new Row { RowIndex = target };
        return before is not null ? sd.InsertBefore(created, before) : sd.AppendChild(created);
    }

    private static Cell GetOrCreateCell(Row row, int colIdx, int rowIdx)
    {
        string cellRef = BuildCellRef(colIdx, rowIdx);
        Cell? existing = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value == cellRef);
        if (existing is not null)
            return existing;
        Cell? before = row.Elements<Cell>().FirstOrDefault(c => ParseColIdx(c.CellReference?.Value) > colIdx);
        var created = new Cell { CellReference = cellRef };
        return before is not null ? row.InsertBefore(created, before) : row.AppendChild(created);
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

    private static int ParseColIdx(string? cellRef)
    {
        int col = 0;
        if (string.IsNullOrEmpty(cellRef))
            return 0;
        foreach (char ch in cellRef)
        {
            if (ch is >= 'A' and <= 'Z')
                col = (col * 26) + (ch - 'A' + 1);
            else
                break;
        }
        return col - 1;
    }

    private static void InjectDummyVbaPart(string workbookPath)
    {
        using var archive = ZipFile.Open(workbookPath, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.CreateEntry("xl/vbaProject.bin");
        using Stream s = entry.Open();
        byte[] bytes = Encoding.UTF8.GetBytes("SYNTHETIC_VBA_MACRO_BINARY_FOR_HASH_TEST");
        s.Write(bytes, 0, bytes.Length);
    }

    private static void AssertExportedRevueltoCellsAndXmlOrder(string output)
    {
        Dictionary<string, string> r13 = ReadRowCells(output, "CIERRE RECEP", 13);
        Dictionary<string, string> r14 = ReadRowCells(output, "CIERRE RECEP", 14);
        Assert.Equal("12", r13["J13"]);
        Assert.Equal("1419040", r13["K13"]);
        Assert.Equal("12", r14["L14"]);
        Assert.Equal("1419038", r14["M14"]);
        Assert.Equal("4", r13["AS13"]);
        Assert.Equal("88", r13["AT13"]);
        Assert.Equal("5", r13["AW13"]);
        Assert.Equal("100", r13["AX13"]);

        using SpreadsheetDocument doc = SpreadsheetDocument.Open(output, false);
        WorkbookPart wbPart = doc.WorkbookPart!;
        Assert.True(wbPart.Workbook!.CalculationProperties?.FullCalculationOnLoad?.Value);
        Sheet sheet = wbPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name!.Value!.Contains("CIERRE RECEP", StringComparison.OrdinalIgnoreCase));
        var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
        Row row13 = wsPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().First(r => r.RowIndex?.Value == 13);
        int prevCol = -1;
        foreach (Cell cell in row13.Elements<Cell>())
        {
            int col = ParseColIdx(cell.CellReference?.Value);
            Assert.True(col > prevCol);
            prevCol = col;
        }
    }

    private static void AssertNoCalcChainAndFullCalcEnabled(string output)
    {
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(output, false);
        Assert.Null(doc.WorkbookPart!.CalculationChainPart);
        Assert.True(doc.WorkbookPart.Workbook!.CalculationProperties?.FullCalculationOnLoad?.Value);
    }

    private string CreateBrokenSharedMasterWorkbook(string fileName)
    {
        string path = Path.Combine(_tempRoot, fileName);
        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.MacroEnabledWorkbook))
        {
            WorkbookPart wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
            WorksheetPart wsPart = AddSheetWithPart(wbPart, sheets, "CIERRE RECEP", 1);
            SheetData sd = wsPart.Worksheet!.GetFirstChild<SheetData>()!;

            SetInlineCell(sd, 13, 4, "1419038", isFormula: false);
            SetInlineCell(sd, 13, 6, "100", isFormula: false);
            Row r13 = GetOrCreateRow(sd, 13);
            Cell h13 = GetOrCreateCell(r13, 7, 13);
            h13.CellFormula = new CellFormula("+G13-#REF!") { FormulaType = CellFormulaValues.Shared, Reference = "H13:H14", SharedIndex = 0 };

            SetInlineCell(sd, 14, 4, "1419040", isFormula: false);
            SetInlineCell(sd, 14, 6, "50", isFormula: false);
            Row r14 = GetOrCreateRow(sd, 14);
            Cell h14 = GetOrCreateCell(r14, 7, 14);
            h14.CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 };

            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
        }
        InjectDummyVbaPart(path);
        return path;
    }

    private string CreateSubtotalAndSstWorkbook(string fileName)
    {
        string path = Path.Combine(_tempRoot, fileName);
        using SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        WorkbookPart wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();
        Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());
        WorksheetPart wsPart = AddSheetWithPart(wbPart, sheets, "CIERRE RECEP", 1);
        SharedStringTablePart sstPart = wbPart.AddNewPart<SharedStringTablePart>();
        sstPart.SharedStringTable = new SharedStringTable(new SharedStringItem(new Text("5")), new SharedStringItem(new Text("100")));
        SheetData sd = wsPart.Worksheet!.GetFirstChild<SheetData>()!;

        Row r13 = GetOrCreateRow(sd, 13);
        Cell as13 = GetOrCreateCell(r13, 44, 13);
        as13.DataType = CellValues.SharedString;
        as13.CellValue = new CellValue("0");
        Cell at13 = GetOrCreateCell(r13, 45, 13);
        at13.DataType = CellValues.SharedString;
        at13.CellValue = new CellValue("1");

        SetInlineCell(sd, 14, 44, "SUBTOTAL(9,AS13:AS13)", isFormula: true, cached: "5");
        SetInlineCell(sd, 14, 45, "SUBTOTAL(9,AT13:AT13)", isFormula: true, cached: "100");
        wsPart.Worksheet.Save();
        sstPart.SharedStringTable.Save();
        wbPart.Workbook.Save();
        return path;
    }

    private static Dictionary<string, string> ReadRowCells(string path, string sheetName, int rowIdx)
    {
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false);
        WorkbookPart wbPart = doc.WorkbookPart!;
        string[] sst = wbPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>()
            .Select(si => si.InnerText).ToArray() ?? [];
        Sheet sheet = wbPart.Workbook!.Sheets!.Elements<Sheet>()
            .First(s => s.Name!.Value!.Contains(sheetName, StringComparison.OrdinalIgnoreCase));
        var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!.Value!);
        Row? row = wsPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIdx);
        var dict = new Dictionary<string, string>(StringComparer.Ordinal);
        if (row is null)
            return dict;

        foreach (Cell cell in row.Elements<Cell>())
        {
            if (cell.CellReference?.Value is not string cellRef)
                continue;
            string val = cell.DataType?.Value == CellValues.SharedString && int.TryParse(cell.CellValue?.Text, out int idx)
                ? sst[idx]
                : cell.CellValue?.Text ?? "";
            dict[cellRef] = val;
        }
        return dict;
    }
}
