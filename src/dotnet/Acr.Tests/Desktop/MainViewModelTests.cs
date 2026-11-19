using Acr.Core.Models;
using Acr.Core.SapKnowledge;
using Acr.Desktop.ViewModels;
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
        Assert.Equal(1, vm.SelectedTabIndex); // Switched to Tab 2 (Index 1)
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
}
