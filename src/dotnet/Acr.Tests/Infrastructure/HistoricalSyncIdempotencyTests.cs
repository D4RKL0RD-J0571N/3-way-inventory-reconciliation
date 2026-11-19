using Acr.Core.Models;
using Acr.Core.Services;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Safety;
using Acr.Infrastructure.Sync;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Tests.Infrastructure;

public sealed class HistoricalSyncIdempotencyTests : IDisposable
{
    private const string ProductionDataDir = @"D:\Development\Software Development\3 Way Inventory Reconciliation\data";
    private const string ProductionDailyLogPath = @"D:\Development\Software Development\3 Way Inventory Reconciliation\Control_Diario_Operaciones_Jostin.xlsx";

    private readonly string _tempRoot;
    private readonly OperationalSafetyGuard _testingGuard;

    public HistoricalSyncIdempotencyTests()
    {
        _tempRoot = Path.Combine(Path.GetTempPath(), $"AcrSyncTests_{Guid.NewGuid():N}");
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
    public void WorkbookSignatureScanner_SkipsLockfileAndClassifiesRealWorkbooks()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        string lockCopy = CopyFixtureToTemp(@"IMP-349 PEND CONTROL\~$IMP-349 (BORRADOR)  AQUA FASHION S.A..xlsm");
        string closureCopy = CopyFixtureToTemp(@"IMP-104\IMP-104.xlsm");
        string hybridCopy = CopyFixtureToTemp(@"IMP-218\IMP-218.xlsm");
        string qcCopy = CopyFixtureToTemp(@"IMP-100\Informacion General\IMP 100 Control Calidad.xlsm");

        Assert.True(WorkbookSignatureScanner.IsLockOrTempFile(lockCopy));
        Assert.Equal(WorkbookKind.Unknown, WorkbookSignatureScanner.Inspect(lockCopy, _testingGuard).Kind);
        Assert.Equal(WorkbookKind.ReceivingClosure, WorkbookSignatureScanner.Inspect(closureCopy, _testingGuard).Kind);
        Assert.Equal(WorkbookKind.HybridQcAndClosure, WorkbookSignatureScanner.Inspect(hybridCopy, _testingGuard).Kind);
        Assert.Equal(WorkbookKind.QcInspection, WorkbookSignatureScanner.Inspect(qcCopy, _testingGuard).Kind);

        IReadOnlyList<WorkbookSignature> scanned = WorkbookSignatureScanner.ScanDirectory(_tempRoot, recursive: true, _testingGuard);
        Assert.Equal(3, scanned.Count);
        Assert.DoesNotContain(scanned, s => s.FilePath.Contains("~$", StringComparison.Ordinal));
    }

    [Fact]
    public void Imp209_SharedStringsTrap_DoesNotMatchJostinFromSharedStrings()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        CopyFixtureToTemp(@"IMP-209\IMP-209.xlsm");
        HistoricalSyncPreview preview = HistoricalClosureScanner.Scan(_tempRoot, Array.Empty<DailyLogEntry>(), _testingGuard);

        Assert.Equal(1, preview.TotalWorkbooksScanned);
        Assert.Equal(0, preview.TargetAuthorWorkbooksMatched);
        Assert.Empty(preview.RawRecords);
    }

    [Fact]
    public void Imp326_AuthorTypoSobalbaro_NormalizedAndMatched()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        CopyFixtureToTemp(@"IMP-326\IMP-326.xlsm");
        HistoricalSyncPreview preview = HistoricalClosureScanner.Scan(_tempRoot, Array.Empty<DailyLogEntry>(), _testingGuard);

        Assert.Equal(1, preview.TargetAuthorWorkbooksMatched);
        HistoricalSyncRecord record = Assert.Single(preview.RawRecords);
        Assert.Equal("IMP-326", record.Identity.NormalizedContainerId);
        Assert.Equal("4100002302", record.Identity.Pedido41);
        Assert.Equal(HistoricalIdentityMatcher.CanonicalTargetAuthor, record.Provenance.NormalizedAuthor);
        Assert.Equal(["4300007309", "4300008144"], record.Identity.Po43Numbers);
        Assert.Equal(2, record.CierreRecepDistinctPoCount);
        Assert.False(record.RequiresHumanReview);
    }

    [Fact]
    public void Imp182_DuplicateDraftsAndScientificNotation_CollapseToFinalWithoutDuplicates()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        CopyFixtureToTemp(@"IMP-182\IMP-182.xlsm");
        CopyFixtureToTemp(@"IMP-182\IMP-182 inprogress.xlsm");
        CopyFixtureToTemp(@"Copia de IMP-182 inprogress.xlsm");

        HistoricalSyncPreview preview = HistoricalClosureScanner.Scan(_tempRoot, Array.Empty<DailyLogEntry>(), _testingGuard);

        Assert.True(preview.SkippedOrErrorFiles.Count == 0, string.Join(" || ", preview.SkippedOrErrorFiles));
        Assert.Equal(3, preview.TargetAuthorWorkbooksMatched);
        CollapsedHistoricalCandidate candidate = Assert.Single(preview.CollapsedCandidates);
        Assert.Equal("IMP-182.xlsm", candidate.PrimaryRecord.Provenance.SourceFileName);
        Assert.Equal(3, candidate.AllProvenances.Count);
        Assert.Contains(HumanReviewReason.MultipleFileVersionsForIdentity, candidate.CombinedReviewReasons);

        HistoricalSyncRecord draft = preview.RawRecords.First(r => r.Provenance.SourceFileName == "IMP-182 inprogress.xlsm");
        Assert.Equal(12, draft.CierreRecepDistinctPoCount);
    }

    [Fact]
    public void MovedAndCopiedClosureFile_ProducesZeroDuplicatesOnResync()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        string orig = CopyFixtureToTemp(@"IMP-104\IMP-104.xlsm");
        string movedDir = Path.Combine(_tempRoot, "Archive");
        Directory.CreateDirectory(movedDir);
        File.Copy(orig, Path.Combine(movedDir, "RenamedCopy_IMP104.xlsm"), overwrite: true);

        DailyLogEntry[] existingLog =
        [
            new DailyLogEntry
            {
                ContainerId = "IMP-104",
                NormalizedContainerId = "IMP-104",
                Pedido41 = "4100002248"
            }
        ];

        HistoricalSyncPreview preview = HistoricalClosureScanner.Scan(_tempRoot, existingLog, _testingGuard);

        Assert.Equal(2, preview.TargetAuthorWorkbooksMatched);
        CollapsedHistoricalCandidate candidate = Assert.Single(preview.CollapsedCandidates);
        Assert.Equal("IMP-104", candidate.PrimaryRecord.Identity.NormalizedContainerId);
        Assert.Equal("4100002248", candidate.PrimaryRecord.Identity.Pedido41);

        HistoricalMatchResult match = Assert.Single(preview.MatchResults);
        Assert.Equal(HistoricalMatchKind.AlreadySyncedExact, match.MatchKind);
        Assert.False(match.IsIdempotentNoOp);
        Assert.Contains(HumanReviewReason.MultipleFileVersionsForIdentity, match.MatchReviewReasons);

        File.Delete(Path.Combine(movedDir, "RenamedCopy_IMP104.xlsm"));
        HistoricalSyncPreview singlePreview = HistoricalClosureScanner.Scan(_tempRoot, existingLog, _testingGuard);
        HistoricalMatchResult singleMatch = Assert.Single(singlePreview.MatchResults);
        Assert.True(singleMatch.IsIdempotentNoOp);
    }

    [Fact]
    public void Imp224_225_226_SharedPo4300007562_ProduceThreeDistinctRecords_AndFlags225RefError()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        CopyFixtureToTemp(@"IMP-224\IMP-224.xlsm");
        CopyFixtureToTemp(@"IMP-225\IMP-225.xlsm");
        CopyFixtureToTemp(@"IMP-226\IMP-226.xlsm");

        HistoricalSyncPreview preview = HistoricalClosureScanner.Scan(_tempRoot, Array.Empty<DailyLogEntry>(), _testingGuard);

        Assert.Equal(3, preview.CollapsedCandidates.Count);
        CollapsedHistoricalCandidate c225 = preview.CollapsedCandidates.Single(c => c.PrimaryRecord.Identity.NormalizedContainerId == "IMP-225");
        Assert.Contains("#REF!", c225.PrimaryRecord.DetectedErrorTokens);
        Assert.Contains(HumanReviewReason.FormulaOrLiteralErrorInSummary, c225.CombinedReviewReasons);
    }

    [Fact]
    public void HumanReviewTriggers_RealDatasetAnomalies()
    {
        if (!Directory.Exists(ProductionDataDir))
            return;

        CopyFixtureToTemp(@"IMP-195 SIN CONTROL\IMP-195.xlsm");
        CopyFixtureToTemp(@"IMP-241 Sin Control Calidad\IMP-241.xlsm");
        CopyFixtureToTemp(@"IMP-242  PEND CONTROL\IMP-242.xlsm");
        CopyFixtureToTemp(@"IMP-325 PEND CONTROL\IMP-325.xlsm");
        CopyFixtureToTemp(@"IMP-360\IMP-360.xlsm");

        HistoricalSyncPreview preview = HistoricalClosureScanner.Scan(_tempRoot, Array.Empty<DailyLogEntry>(), _testingGuard);

        Assert.Equal(5, preview.CollapsedCandidates.Count);
        AssertReviewReason(preview, "IMP-195", HumanReviewReason.ZeroCartonsWithPositiveLines);
        AssertReviewReason(preview, "IMP-242", HumanReviewReason.ZeroCartonsWithPositiveLines);
        AssertReviewReason(preview, "IMP-241", HumanReviewReason.FormulaOrLiteralErrorInSummary);
        AssertReviewReason(preview, "IMP-241", HumanReviewReason.LineCountMismatchDgVsCierreRecep);
        AssertReviewReason(preview, "IMP-325", HumanReviewReason.FolderStatusVsCartonMismatch);
        AssertReviewReason(preview, "IMP-360", HumanReviewReason.LineCountMismatchDgVsCierreRecep);
        AssertReviewReason(preview, "IMP-360", HumanReviewReason.FormulaOrLiteralErrorInSummary);
    }

    [Fact]
    public void SyntheticXmlEdgeCases_InlineStringAnd1904DateAndMissingPedido41()
    {
        CreateSynthetic1904ClosureWorkbook("synthetic_1904.xlsx");
        DailyLogEntry[] existingLog =
        [
            new DailyLogEntry
            {
                ContainerId = "IMP-901",
                NormalizedContainerId = "IMP-901",
                Pedido41 = "4100009001"
            }
        ];

        HistoricalSyncPreview preview = HistoricalClosureScanner.Scan(_tempRoot, existingLog, _testingGuard);

        HistoricalSyncRecord record = Assert.Single(preview.RawRecords);
        Assert.Equal("IMP-901", record.Identity.NormalizedContainerId);
        Assert.Null(record.Identity.Pedido41);
        Assert.Equal(new DateOnly(1904, 1, 2), record.Identity.RevisionDate);

        HistoricalMatchResult match = Assert.Single(preview.MatchResults);
        Assert.Equal(HistoricalMatchKind.MatchedByContainerFallbackRequiresReview, match.MatchKind);
        Assert.True(match.RequiresHumanReview);
        Assert.Contains(HumanReviewReason.MissingPedido41, match.MatchReviewReasons);
    }

    [Fact]
    public void Scan_MalformedOaDateMinus657435AndInt32Overflow_DoesNotAbortScan()
    {
        CreateSyntheticEdgeClosureWorkbook(
            "bad_oa_and_overflow.xlsx",
            "IMP-902",
            c9Value: "9223372036854775808",
            c13Value: "9999999999999",
            c30Value: "-657435",
            omitCierreRowAttributes: false);
        CreateSynthetic1904ClosureWorkbook("valid_neighbor.xlsx");

        HistoricalSyncPreview preview = HistoricalClosureScanner.Scan(_tempRoot, Array.Empty<DailyLogEntry>(), _testingGuard);

        Assert.Empty(preview.SkippedOrErrorFiles);
        Assert.Equal(2, preview.TargetAuthorWorkbooksMatched);
        HistoricalSyncRecord edgeRecord = preview.RawRecords.Single(r => r.Identity.NormalizedContainerId == "IMP-902");
        Assert.Null(edgeRecord.Identity.RevisionDate);
        Assert.Null(edgeRecord.TheoreticalCartons);
        Assert.Equal("9223372036854775808", edgeRecord.Identity.Pedido41);
        Assert.Contains(preview.RawRecords, r => r.Identity.NormalizedContainerId == "IMP-901");
    }

    [Fact]
    public void Scan_CierreRecepRowsWithoutRowAttribute_UsesSequentialFallbackRowIndex()
    {
        CreateSyntheticEdgeClosureWorkbook(
            "no_row_attr.xlsx",
            "IMP-903",
            c9Value: "4100009003",
            c13Value: "10",
            c30Value: "45000",
            omitCierreRowAttributes: true);

        HistoricalSyncPreview preview = HistoricalClosureScanner.Scan(_tempRoot, Array.Empty<DailyLogEntry>(), _testingGuard);

        Assert.Empty(preview.SkippedOrErrorFiles);
        HistoricalSyncRecord record = Assert.Single(preview.RawRecords);
        Assert.Equal("IMP-903", record.Identity.NormalizedContainerId);
        Assert.Equal(1, record.ActualCierreRecepItemsCount);
        Assert.Equal(1, record.CierreRecepDistinctPoCount);
    }

    private string CopyFixtureToTemp(string relativePath)
    {
        string source = Path.Combine(ProductionDataDir, relativePath);
        string dest = Path.Combine(_tempRoot, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(dest)!);
        File.Copy(source, dest, overwrite: true);
        return dest;
    }

    private static void AssertReviewReason(HistoricalSyncPreview preview, string containerId, HumanReviewReason expected)
    {
        CollapsedHistoricalCandidate candidate = preview.CollapsedCandidates
            .Single(c => c.PrimaryRecord.Identity.NormalizedContainerId == containerId);
        Assert.Contains(expected, candidate.CombinedReviewReasons);
    }

    private void CreateSynthetic1904ClosureWorkbook(string filename)
    {
        string path = Path.Combine(_tempRoot, filename);
        using SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        WorkbookPart wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook(new WorkbookProperties { Date1904 = true });
        Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());

        SheetData dgData = AddWorksheet(wbPart, sheets, "DATOS GENERALES", 1);
        PopulateSyntheticDgSheet(dgData);

        SheetData crData = AddWorksheet(wbPart, sheets, "CIERRE RECEP", 2);
        crData.AppendChild(CreateRow(13, ("A13", "4300009001", CellValues.String), ("E13", "1002003001", CellValues.String)));
        wbPart.Workbook.Save();
    }

    private void CreateSyntheticEdgeClosureWorkbook(
        string filename,
        string containerId,
        string c9Value,
        string c13Value,
        string c30Value,
        bool omitCierreRowAttributes)
    {
        string path = Path.Combine(_tempRoot, filename);
        using SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        WorkbookPart wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();
        Sheets sheets = wbPart.Workbook.AppendChild(new Sheets());

        SheetData dgData = AddWorksheet(wbPart, sheets, "DATOS GENERALES", 1);
        PopulateSyntheticDgSheet(dgData);
        dgData.Elements<Row>().First(r => r.RowIndex?.Value == 7).Elements<Cell>().First().CellValue = new CellValue(containerId);
        dgData.Elements<Row>().First(r => r.RowIndex?.Value == 13).Elements<Cell>().First().CellValue = new CellValue(c13Value);
        dgData.Elements<Row>().First(r => r.RowIndex?.Value == 30).Elements<Cell>().First().CellValue = new CellValue(c30Value);
        dgData.AppendChild(CreateRow(9, ("C9", c9Value, CellValues.Number)));

        SheetData crData = AddWorksheet(wbPart, sheets, "CIERRE RECEP", 2);
        PopulateEdgeCierreSheet(crData, omitCierreRowAttributes);
        wbPart.Workbook.Save();
    }

    private static void PopulateEdgeCierreSheet(SheetData crData, bool omitRowAttributes)
    {
        if (!omitRowAttributes)
        {
            crData.AppendChild(CreateRow(13, ("A13", "4300009002", CellValues.String), ("E13", "1002003002", CellValues.String)));
            return;
        }

        for (int i = 1; i <= 12; i++)
            crData.AppendChild(new Row());

        var dataRow = new Row();
        dataRow.AppendChild(new Cell { CellReference = "A13", DataType = CellValues.String, CellValue = new CellValue("4300009002") });
        dataRow.AppendChild(new Cell { CellReference = "E13", DataType = CellValues.String, CellValue = new CellValue("1002003002") });
        crData.AppendChild(dataRow);
    }

    private static void PopulateSyntheticDgSheet(SheetData dgData)
    {
        dgData.AppendChild(CreateRow(7, ("C7", "IMP-901", CellValues.String)));
        dgData.AppendChild(CreateRow(8, ("C8", "4300009001", CellValues.String)));
        dgData.AppendChild(CreateRow(10, ("C10", "SYNTHETIC VENDOR", CellValues.String)));
        dgData.AppendChild(CreateRow(13, ("C13", "10", CellValues.Number)));
        dgData.AppendChild(CreateRow(14, ("C14", "10", CellValues.Number)));
        dgData.AppendChild(CreateRow(18, ("C18", "1", CellValues.Number)));
        dgData.AppendChild(CreateRow(30, ("C30", "1", CellValues.Number)));

        var b33Row = new Row { RowIndex = 33 };
        b33Row.AppendChild(new Cell
        {
            CellReference = "B33",
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text("ELABORADO POR:\nJostin Lopez Sobalbarro"))
        });
        dgData.AppendChild(b33Row);
    }

    private static SheetData AddWorksheet(WorkbookPart wbPart, Sheets sheets, string name, uint sheetId)
    {
        WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
        var data = new SheetData();
        wsPart.Worksheet = new Worksheet(data);
        sheets.AppendChild(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = sheetId, Name = name });
        return data;
    }

    private static Row CreateRow(uint rowIndex, params (string Ref, string Value, CellValues Type)[] cells)
    {
        var row = new Row { RowIndex = rowIndex };
        foreach ((string cellRef, string value, CellValues type) in cells)
            row.AppendChild(new Cell { CellReference = cellRef, DataType = type, CellValue = new CellValue(value) });
        return row;
    }
}
