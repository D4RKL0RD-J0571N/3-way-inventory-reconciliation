using Acr.Core.Models;
using Acr.Core.SapKnowledge;
using Acr.Desktop.ViewModels;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Receiving;
using Xunit;

namespace Acr.Tests.Desktop;

public sealed class MainViewModelTests
{
    [Fact]
    public void MainViewModel_InitialState_LoadsSapCatalogAndDefaultStages()
    {
        var vm = new MainViewModel();

        Assert.NotEmpty(vm.FilteredSapTransactions);
        Assert.NotEmpty(vm.DiagnosticGuides);
        Assert.NotNull(vm.Stage1Options);
        Assert.NotNull(vm.Stage2Options);
        Assert.NotNull(vm.Stage3Options);
        Assert.NotNull(vm.Stage4Options);
        Assert.NotNull(vm.Stage5Options);
        Assert.NotNull(vm.Stage6Options);
        Assert.NotNull(vm.SaludoOptions);
        Assert.False(vm.IsBusy);
    }

    [Fact]
    public void MainViewModel_RecalculateFormulas_ComputesCorrectColumnOAndColumnR()
    {
        var vm = new MainViewModel
        {
            ContainerId = "IMP-300",
            Entrega18Number = "18001234",
            SelectedSaludo = Saludo.BuenosDias,
            Stage1Customs = CustomsDuaStage.DuaAndDocsReady,
            Stage2SapIntake = SapIntakeStage.ReadyAndNotified,
            Stage5ReceptionClosure = ReceptionClosureStage.BalancedZeroDiff,
            Stage6DispatchClosure = DispatchClosureStage.DispatchClosureReady
        };

        // When Stage 6 is ready, Col O should be "7. Cierre Total (Recep + Desp)"
        Assert.Equal("7. Cierre Total (Recep + Desp)", vm.ComputedColumnO);
        Assert.Contains("Se informa que el ingreso a nivel de sistema del Contenedor 300, quedo listo.", vm.ComputedColumnR);
    }

    [Fact]
    public void MainViewModel_RecalculateFormulas_WithNoEntrega_SetsColumnRToDash()
    {
        var vm = new MainViewModel
        {
            ContainerId = "IMP-301",
            Entrega18Number = "",
            Stage6DispatchClosure = DispatchClosureStage.BlockedByReception,
            Stage5ReceptionClosure = ReceptionClosureStage.PendingClosure,
            Stage2SapIntake = SapIntakeStage.PendingSapIntake
        };

        Assert.Equal("-", vm.ComputedColumnR);
        Assert.Equal("2. Doc SAP Listo (Pend. Ingreso)", vm.ComputedColumnO);
    }

    [Fact]
    public void MainViewModel_OpenDiagnosticGuideForDiscrepancy_JumpsToCorrectSapGuide()
    {
        var vm = new MainViewModel();
        var shortageItem = new DiscrepancyItem
        {
            MaterialCode = "1009988",
            DeclaredUnits = 100,
            PhysicalUnits = 80,
            DiscrepancyType = DiscrepancyType.Shortage
        };

        vm.OpenDiagnosticGuideForDiscrepancy(shortageItem);

        Assert.NotNull(vm.SelectedDiagnosticGuide);
        Assert.Contains("FALTANTE", vm.SelectedDiagnosticGuide.GuideId);
        Assert.Equal(7, vm.SelectedTabIndex); // Switched to Tab 8 (Index 7: SAP & Wiki)
        Assert.Equal(1, vm.WikiSubTab);        // Jumped to Guías sub-tab
    }

    [Fact]
    public void MainViewModel_CartonCalculator_CalculatesCanonicalExpressions()
    {
        var vm = new MainViewModel();

        // 1. Standard compound: 420 units, box 48, odd 36 -> 8*48 1*36 (9 boxes)
        vm.CalcUnits = 420;
        vm.CalcStdBox = 48;
        vm.CalcOddUnits = 36;
        vm.CalcHasOdds = true;
        vm.CalcOddModeIsLoose = false;

        Assert.Equal("8*48 1*36", vm.CalcPkgStr);
        Assert.Equal(9, vm.CalcTotalBoxes);
        Assert.Equal(8, vm.CalcBoxesFull);
        Assert.Equal(1, vm.CalcBoxesComp);

        // 2. Loose units mode: 8*48 +36 (8 boxes, +36 sueltas)
        vm.CalcOddModeIsLoose = true;
        Assert.Equal("8*48 +36", vm.CalcPkgStr);
        Assert.Equal(8, vm.CalcTotalBoxes);
        Assert.Equal(8, vm.CalcBoxesFull);
        Assert.Equal(0, vm.CalcBoxesComp);

        // 3. Exact full boxes: 480 units, box 48, no odds -> "48" (10 boxes)
        vm.CalcUnits = 480;
        vm.CalcHasOdds = false;
        Assert.Equal("48", vm.CalcPkgStr);
        Assert.Equal(10, vm.CalcTotalBoxes);
    }

    [Fact]
    public void MainViewModel_SapClipboardSanitizer_LoadsAndSanitizesGrid()
    {
        var vm = new MainViewModel();
        vm.LoadSapSampleCommand.Execute(null);

        Assert.NotEmpty(vm.SapRawPasteText);
        Assert.Contains("000000001455102001", vm.SapRawPasteText);

        vm.SanitizeSapClipboardCommand.Execute(null);

        Assert.NotNull(vm.SanitizedTable);
        Assert.Equal(4, vm.SapSanitizedHeaders.Count);
        Assert.Equal(4, vm.SapSanitizedRows.Count);
        Assert.Equal("1455102001", vm.SapSanitizedRows[0][0]);
        Assert.Equal("1E2", vm.SapSanitizedRows[3][0]); // Alphanumeric preserved
        Assert.Equal("-50", vm.SapSanitizedRows[3][2]); // Preserved negative dollar sign
    }

    [Fact]
    public void MainViewModel_Checklist_TracksProgressAndDynamicLabels()
    {
        var vm = new MainViewModel
        {
            KpiContainerId = "IMP-207",
            KpiDeclaredUnits = 10000,
            KpiPhysicalUnits = 9940
        };

        vm.UpdateChecklistDynamicLabels();

        Assert.Contains("IMP-207", vm.Checklist1Text);
        Assert.Contains("10,000", vm.Checklist2Text);
        Assert.Contains("9,940", vm.Checklist3Text);
        Assert.Equal(100, vm.ChecklistProgress);
        Assert.True(vm.ChecklistIsComplete);

        // Uncheck one item
        vm.Checklist7Dua = false;
        Assert.Equal(6, vm.ChecklistCompletedCount);
        Assert.False(vm.ChecklistIsComplete);
        Assert.Equal(86, vm.ChecklistProgress);
    }

    [Fact]
    public void MainViewModel_SearchSap_FiltersTransactionsCorrectly()
    {
        var vm = new MainViewModel
        {
            SapSearchQuery = "ME80FN"
        };

        vm.SearchSapCommand.Execute(null);

        Assert.Single(vm.FilteredSapTransactions);
        Assert.Equal("ME80FN", vm.FilteredSapTransactions[0].Code);
    }

    [Fact]
    public void MainViewModel_GenerateSapClipboardNotification_GeneratesFormattedSummary()
    {
        var vm = new MainViewModel
        {
            ContainerId = "IMP-400",
            Po43Number = "4300001234",
            Pedido41Number = "4100005678",
            Entrega18Number = "18009999",
            EwmTaskNumber = "998877",
            HandlingUnit = "112345678005808817"
        };

        vm.GenerateSapClipboardNotificationCommand.Execute(null);

        Assert.Contains("*RECEPCIÓN CONTENEDOR IMP-400*", vm.ClipboardSummary);
        Assert.Contains("PO: 4300001234", vm.ClipboardSummary);
        Assert.Contains("112345678005808817", vm.ClipboardSummary);
        Assert.Contains("Etapa Actual (Col O):", vm.ClipboardSummary);
        Assert.Contains("Notificación SAP (Col R):", vm.ClipboardSummary);
    }

    [Fact]
    public void MainViewModel_BlindReceiving_SetsIsBlindReceivingAndUpdatesStatusAndChecklist()
    {
        var vm = new MainViewModel();

        var blindSummary = new ComparisonSummary
        {
            TotalDeclaredUnits = 800,
            TotalPhysicalUnits = 0,
            NetUnitDifference = -800,
            TotalDeclaredCartons = 40,
            TotalPhysicalCartons = 0,
            NetCartonDifference = -40,
            CleanMaterialsCount = 0,
            DiscrepantMaterialsCount = 5,
            Items = Array.Empty<DiscrepancyItem>()
        };

        var mockResult = new ReceivingPipelineResult
        {
            ContainerId = "IMP-BLIND",
            QcReport = new QcInspectionReport
            {
                FilePath = "dummy.xlsm",
                ContainerId = "IMP-BLIND",
                Items = Array.Empty<QcInspectionItem>(),
                DetectedParts = Array.Empty<string>(),
                Warnings = Array.Empty<string>()
            },
            DeclaredItems = Array.Empty<DeclaredItem>(),
            PhysicalItems = Array.Empty<PhysicalItem>(),
            Comparison = blindSummary,
            BultosItems = Array.Empty<BultosPresentationItem>(),
            Tira1AsAtTsv = "",
            Tira2AwAyTsv = ""
        };

        vm.LastReconciliationResult = mockResult;
        vm.IsBlindReceiving = blindSummary.IsBlindReceiving;
        vm.UpdateChecklistDynamicLabels();

        Assert.True(vm.IsBlindReceiving);
        Assert.False(vm.Checklist6Bultos);
        Assert.Contains("MONTAJE A CIEGAS", vm.Checklist6Text);
    }
}
