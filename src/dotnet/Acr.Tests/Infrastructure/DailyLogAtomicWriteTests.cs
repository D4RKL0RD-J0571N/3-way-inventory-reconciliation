using System.Security.Cryptography;
using Acr.Core.Models;
using Acr.Core.Services;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Safety;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Tests.Infrastructure;

public sealed class DailyLogAtomicWriteTests : IDisposable
{
    private const string ProductionDataDir = @"D:\Development\Software Development\3 Way Inventory Reconciliation\data";
    private const string ProductionDailyLogPath = @"D:\Development\Software Development\3 Way Inventory Reconciliation\Control_Diario_Operaciones_Jostin.xlsx";

    private readonly string _tempRoot;
    private readonly OperationalSafetyGuard _testingGuard;

    public DailyLogAtomicWriteTests()
    {
        _tempRoot = Path.Combine(Path.GetTempPath(), $"AcrDailyLogTests_{Guid.NewGuid():N}");
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
    public void ReadRealDailyLogCopy_ReadsAll22RowsCatalogsAndFlagsLiveAnomalies()
    {
        if (!File.Exists(ProductionDailyLogPath))
            return;

        string copy = CopyRealDailyLogToTemp();
        DailyLogLoadResult loaded = DailyLogService.Load(copy, _testingGuard);

        Assert.Equal(22, loaded.Snapshot.Entries.Count);
        Assert.Equal(5, loaded.Snapshot.TableStartRow);
        Assert.Equal(27, loaded.Snapshot.TableEndRow);

        Assert.Contains(loaded.Duplicates, d => d.NormalizedContainerId == "IMP-202" && d.Entries.Count == 2);
        Assert.Contains(loaded.SharedDocuments, s => s.DocumentKind == "PO43" && s.DocumentNumber == "4300007748");
        Assert.Contains(loaded.SharedDocuments, s => s.DocumentKind == "Pedido41" && s.DocumentNumber == "4100002329");

        DailyLogEntry row6 = loaded.Snapshot.Entries[0];
        Assert.Equal("IMP-202", row6.NormalizedContainerId);
        Assert.Equal("112345678005808817", row6.HandlingUnit);

        Assert.Equal(3, loaded.Snapshot.Catalogs.Stage1CustomsOptions.Count);
        Assert.Equal(4, loaded.Snapshot.Catalogs.Stage2SapIntakeOptions.Count);
        Assert.Equal(6, loaded.Snapshot.Catalogs.Stage5ReceptionClosureOptions.Count);
        Assert.Equal(3, loaded.Snapshot.Catalogs.SaludoOptions.Count);
    }

    [Fact]
    public void AppendNewRows_UpdatesTableRefDataValidationsDashboardAndHealsColOColR()
    {
        if (!File.Exists(ProductionDailyLogPath))
            return;

        string copy = CopyRealDailyLogToTemp();
        var batch = BuildBatchEntries(startNum: 901, count: 33);

        (AtomicTransactionResult tx, IReadOnlyList<int> rows) = DailyLogRepository.SaveEntries(copy, batch, _testingGuard);

        Assert.True(tx.BackupCreated);
        Assert.True(tx.Verified);
        Assert.NotNull(tx.BackupPath);
        Assert.True(File.Exists(tx.BackupPath!));
        Assert.Empty(Directory.GetFiles(_tempRoot, "*.tmp"));
        Assert.Equal(60, rows[^1]);

        AssertWorkbookResizedAndHealed(copy, expectedEndRow: 60);
    }

    [Fact]
    public void SafetyModes_BlockDevelopmentAndTestingWritesToProductionDailyLog()
    {
        var devGuard = new OperationalSafetyGuard(OperationalSafetyMode.Development, ProductionDataDir, ProductionDailyLogPath, _tempRoot);
        var prodGuard = new OperationalSafetyGuard(OperationalSafetyMode.Production, ProductionDataDir, ProductionDailyLogPath, _tempRoot);
        var entry = BuildValidEntry("IMP-950");

        OperationalSafetyException devEx = Assert.Throws<OperationalSafetyException>(
            () => DailyLogRepository.SaveEntry(ProductionDailyLogPath, entry, devGuard));
        Assert.Equal("DEV_DAILY_LOG_READ_ONLY", devEx.Decision.ReasonCode);

        OperationalSafetyException testEx = Assert.Throws<OperationalSafetyException>(
            () => DailyLogRepository.SaveEntry(ProductionDailyLogPath, entry, _testingGuard));
        Assert.Equal("TESTING_REDIRECT_TO_TEMP", testEx.Decision.ReasonCode);

        OperationalSafetyException bgEx = Assert.Throws<OperationalSafetyException>(
            () => prodGuard.EnsureCanWrite(ProductionDailyLogPath, OperationalActionKind.WriteAutomatedOrBackground));
        Assert.Equal("PROD_REQUIRES_EXPLICIT_USER_ACTION", bgEx.Decision.ReasonCode);
    }

    [Fact]
    public void LockfileAndExclusiveHandle_AbortBeforeMutatingTargetOrCreatingBackup()
    {
        if (!File.Exists(ProductionDailyLogPath))
            return;

        string copy = CopyRealDailyLogToTemp();
        string originalHash = Sha256Hex(copy);
        string lockFile = Path.Combine(_tempRoot, $"~${Path.GetFileName(copy)}");
        File.WriteAllText(lockFile, "lock");

        Assert.Throws<IOException>(() => DailyLogRepository.SaveEntry(copy, BuildValidEntry("IMP-960"), _testingGuard));
        Assert.Equal(originalHash, Sha256Hex(copy));
        Assert.False(File.Exists(copy + ".bak"));
        File.Delete(lockFile);

        using (var hold = new FileStream(copy, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            Assert.Throws<IOException>(() => DailyLogRepository.SaveEntry(copy, BuildValidEntry("IMP-961"), _testingGuard));
        }

        Assert.Equal(originalHash, Sha256Hex(copy));
        Assert.False(File.Exists(copy + ".bak"));
        Assert.Empty(Directory.GetFiles(_tempRoot, "*.tmp"));
    }

    [Fact]
    public void InvalidLifecycleTransition_AbortsBeforeFileTransaction()
    {
        if (!File.Exists(ProductionDailyLogPath))
            return;

        string copy = CopyRealDailyLogToTemp();
        string originalHash = Sha256Hex(copy);

        DailyLogEntry badReception = BuildValidEntry("IMP-970") with
        {
            Stage4QualityControl = QualityControlStage.NoQcWaitingDock,
            Stage5ReceptionClosure = ReceptionClosureStage.BalancedZeroDiff
        };
        DailyLogSaveResult res1 = DailyLogService.AddEntry(copy, badReception, _testingGuard);
        Assert.False(res1.Succeeded);
        Assert.Contains(res1.Validation.Errors, e => e.Code == "RECEPTION_CLOSED_REQUIRES_DUA_AND_QC_APPROVED");

        DailyLogEntry badDispatch = BuildValidEntry("IMP-971") with
        {
            Stage5ReceptionClosure = ReceptionClosureStage.PendingClosure,
            Stage6DispatchClosure = DispatchClosureStage.PreWired80PercentReadyMov641
        };
        DailyLogSaveResult res2 = DailyLogService.AddEntry(copy, badDispatch, _testingGuard);
        Assert.False(res2.Succeeded);
        Assert.Contains(res2.Validation.Errors, e => e.Code == "DISPATCH_UNLOCKED_BEFORE_RECEPTION_READY");

        DailyLogSaveResult dupRes = DailyLogService.AddEntry(copy, BuildValidEntry("IMP-349"), _testingGuard);
        Assert.False(dupRes.Succeeded);
        Assert.Contains(dupRes.Validation.Errors, e => e.Code == "DUPLICATE_CONTAINER_DETECTED");

        Assert.Equal(originalHash, Sha256Hex(copy));
        Assert.False(File.Exists(copy + ".bak"));
    }

    [Fact]
    public void FormatSapIntakeNotification_And_SharedDocumentAlert()
    {
        if (!File.Exists(ProductionDailyLogPath))
            return;

        string copy = CopyRealDailyLogToTemp();
        DailyLogEntry newEntry = BuildValidEntry("IMP-980") with
        {
            Po43Numbers = ["4300007748"],
            Pedido41 = "4100002329",
            EntregaVl31n = "180099999",
            EwmDeliveryDoc = "41499",
            HandlingUnit = "112345678005809999",
            Greeting = Saludo.BuenosDias
        };

        DailyLogSaveResult saved = DailyLogService.AddEntry(copy, newEntry, _testingGuard);

        Assert.True(saved.Succeeded);
        Assert.NotNull(saved.Notification);
        Assert.True(saved.Notification!.IsReady);
        Assert.Equal("Buenos días, Se informa que el ingreso a nivel de sistema del Contenedor 980, quedo listo.", saved.Notification.ColRValue);
        Assert.Contains(saved.SharedDocuments, s => s.DocumentKind == "PO43" && s.DocumentNumber == "4300007748" && s.NormalizedContainerIds.Contains("IMP-980"));
        Assert.Contains(saved.SharedDocuments, s => s.DocumentKind == "Pedido41" && s.DocumentNumber == "4100002329" && s.NormalizedContainerIds.Contains("IMP-980"));
    }

    [Fact]
    public void UpdateExistingRow_WithNullUpdatedDate_ClearsColCOnDiskAndReloadsNull()
    {
        if (!File.Exists(ProductionDailyLogPath))
            return;

        string copy = CopyRealDailyLogToTemp();
        DailyLogWorkbookSnapshot before = DailyLogRepository.ReadWorkbook(copy, _testingGuard);
        int idx = before.Entries.ToList().FindIndex(e => e.UpdatedDate.HasValue);
        Assert.True(idx >= 0);

        int targetRow = before.RowIndices[idx];
        DailyLogEntry cleared = before.Entries[idx] with { UpdatedDate = null };
        DailyLogRepository.SaveEntry(copy, cleared, _testingGuard, targetRow);

        DailyLogWorkbookSnapshot after = DailyLogRepository.ReadWorkbook(copy, _testingGuard);
        int afterIdx = after.RowIndices.ToList().IndexOf(targetRow);
        Assert.True(afterIdx >= 0);
        Assert.Null(after.Entries[afterIdx].UpdatedDate);
    }

    private string CopyRealDailyLogToTemp()
    {
        string dest = Path.Combine(_tempRoot, $"DailyCopy_{Guid.NewGuid():N}.xlsx");
        File.Copy(ProductionDailyLogPath, dest, overwrite: true);
        return dest;
    }

    private static List<(DailyLogEntry Entry, int? TargetRowIndex)> BuildBatchEntries(int startNum, int count)
    {
        var list = new List<(DailyLogEntry, int?)>(count);
        for (int i = 0; i < count; i++)
            list.Add((BuildValidEntry($"IMP-{startNum + i}"), null));
        return list;
    }

    private static DailyLogEntry BuildValidEntry(string containerId) =>
        new()
        {
            ContainerId = containerId,
            NormalizedContainerId = containerId,
            UpdatedDate = new DateOnly(2026, 11, 18),
            Po43Numbers = ["4300009001"],
            Pedido41 = "4100009001",
            EntregaVl31n = "180090001",
            EwmDeliveryDoc = "41490",
            HandlingUnit = "112345678005890001",
            Stage1Customs = CustomsDuaStage.DuaAndDocsReady,
            Stage2SapIntake = SapIntakeStage.ReadyAndNotified,
            Stage3BlindSetup = BlindSetupStage.ReadyComplete,
            Stage4QualityControl = QualityControlStage.QcApprovedZDescarga,
            Stage5ReceptionClosure = ReceptionClosureStage.BalancedZeroDiff,
            Stage6DispatchClosure = DispatchClosureStage.BlockedByReception,
            Observations = "Test entry",
            Greeting = Saludo.BuenasTardes
        };

    private void AssertWorkbookResizedAndHealed(string workbookPath, int expectedEndRow)
    {
        DailyLogWorkbookSnapshot snapshot = DailyLogRepository.ReadWorkbook(workbookPath, _testingGuard);
        Assert.Equal(expectedEndRow - 5, snapshot.Entries.Count);
        Assert.Equal(expectedEndRow, snapshot.TableEndRow);

        using SpreadsheetDocument doc = SpreadsheetDocument.Open(workbookPath, false);
        WorkbookPart wbPart = doc.WorkbookPart!;
        Assert.True(wbPart.Workbook!.CalculationProperties?.FullCalculationOnLoad?.Value);

        WorksheetPart maestroPart = wbPart.WorksheetParts.First(p => p.TableDefinitionParts.Any());
        Worksheet maestroWs = maestroPart.Worksheet!;
        CustomSheetView customView = maestroWs.Descendants<CustomSheetView>().First();
        Assert.Equal($"$A$5:$R${expectedEndRow}", customView.GetFirstChild<AutoFilter>()!.Reference!.Value);

        DefinedName filterDn = wbPart.Workbook.DefinedNames!.Elements<DefinedName>()
            .First(d => d.Name?.Value?.Contains("FilterData", StringComparison.Ordinal) == true &&
                        d.Text.Contains("Maestro Contenedores", StringComparison.Ordinal));
        Assert.Contains($"$A$5:$R${expectedEndRow}", filterDn.Text);

        foreach (DataValidation dv in maestroWs.Descendants<DataValidation>())
            Assert.Contains($":{dv.SequenceOfReferences!.Items.First().Value!.Split(':')[1][0]}{expectedEndRow}", dv.SequenceOfReferences.Items.First().Value);

        AssertColOAndColRFormulas(maestroWs, 27, expectedEndRow);
        AssertDashboardUpdated(wbPart, expectedEndRow);
    }

    private static void AssertColOAndColRFormulas(Worksheet maestroWs, int healedRow, int appendedRow)
    {
        SheetData sd = maestroWs.GetFirstChild<SheetData>()!;
        foreach (int targetRow in new[] { healedRow, appendedRow })
        {
            Row row = sd.Elements<Row>().First(r => r.RowIndex?.Value == (uint)targetRow);
            Cell colO = row.Elements<Cell>().First(c => c.CellReference?.Value == $"O{targetRow}");
            Cell colR = row.Elements<Cell>().First(c => c.CellReference?.Value == $"R{targetRow}");
            Assert.NotNull(colO.CellFormula);
            Assert.Null(colO.CellFormula!.FormulaType);
            Assert.NotNull(colR.CellFormula);
            Assert.Null(colR.CellFormula!.FormulaType);
            Assert.False(string.IsNullOrWhiteSpace(colR.CellValue?.Text));
        }
    }

    private static void AssertDashboardUpdated(WorkbookPart wbPart, int expectedEndRow)
    {
        Sheet dashSheet = wbPart.Workbook!.Sheets!.Elements<Sheet>()
            .First(s => string.Equals(s.Name?.Value, "Dashboard General", StringComparison.OrdinalIgnoreCase));
        var dashPart = (WorksheetPart)wbPart.GetPartById(dashSheet.Id!.Value!);
        List<string> formulas = dashPart.Worksheet!.Descendants<Cell>()
            .Where(c => c.CellFormula is not null && c.CellFormula.Text.Contains("Maestro Contenedores", StringComparison.Ordinal))
            .Select(c => c.CellFormula!.Text)
            .ToList();

        Assert.NotEmpty(formulas);
        Assert.All(formulas, f =>
        {
            Assert.Contains(expectedEndRow.ToString(System.Globalization.CultureInfo.InvariantCulture), f);
            Assert.DoesNotContain(":$I$27", f);
            Assert.DoesNotContain(":B27", f);
        });
    }

    private static string Sha256Hex(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        return Convert.ToHexString(SHA256.HashData(stream));
    }
}
