using Acr.Core.Models;
using Acr.Core.Services;

namespace Acr.Tests.Core;

public sealed class DailyLogLifecycleTests
{
    private delegate bool TryParseCatalogDelegate<TEnum>(string? raw, out TEnum parsed) where TEnum : struct, Enum;

    [Fact]
    public void CatalogMappings_RoundTripAllSixStages()
    {
        AssertEnumRoundTrip<CustomsDuaStage>(DailyLogCatalogs.ToExcelString, DailyLogCatalogs.TryParseCustomsDua);
        AssertEnumRoundTrip<SapIntakeStage>(DailyLogCatalogs.ToExcelString, DailyLogCatalogs.TryParseSapIntake);
        AssertEnumRoundTrip<BlindSetupStage>(DailyLogCatalogs.ToExcelString, DailyLogCatalogs.TryParseBlindSetup);
        AssertEnumRoundTrip<QualityControlStage>(DailyLogCatalogs.ToExcelString, DailyLogCatalogs.TryParseQualityControl);
        AssertEnumRoundTrip<ReceptionClosureStage>(DailyLogCatalogs.ToExcelString, DailyLogCatalogs.TryParseReceptionClosure);
        AssertEnumRoundTrip<DispatchClosureStage>(DailyLogCatalogs.ToExcelString, DailyLogCatalogs.TryParseDispatchClosure);
    }

    [Fact]
    public void CatalogMappings_RoundTripSaludo_ParteB_MasterStates_AndAccentTolerance()
    {
        AssertEnumRoundTrip<Saludo>(DailyLogCatalogs.ToExcelString, DailyLogCatalogs.TryParseSaludo);
        AssertEnumRoundTrip<ParteBStatus>(DailyLogCatalogs.ToExcelString, DailyLogCatalogs.TryParseParteBStatus);
        AssertEnumRoundTrip<MasterContainerState>(DailyLogCatalogs.ToExcelString, DailyLogCatalogs.TryParseMasterState);

        Assert.True(DailyLogCatalogs.TryParseCustomsDua("En Radar / Transito", out CustomsDuaStage transit));
        Assert.Equal(CustomsDuaStage.InRadarOrTransit, transit);
        Assert.True(DailyLogCatalogs.TryParseQualityControl("Pendiente Revision Final CC", out QualityControlStage qcRev));
        Assert.Equal(QualityControlStage.PendingFinalQcReview, qcRev);
    }

    [Fact]
    public void ComputeMasterState_MatchesExactColORollupPrecedence()
    {
        AssertMasterState(MasterContainerState.TotalClosureReceptionAndDispatch, new DailyLogEntry
        {
            Stage4QualityControl = QualityControlStage.PendingFinalQcReview,
            Stage5ReceptionClosure = ReceptionClosureStage.BalancedZeroDiff,
            Stage6DispatchClosure = DispatchClosureStage.DispatchClosureReady
        });
        AssertMasterState(MasterContainerState.ReceptionClosedReady, new DailyLogEntry
        {
            Stage4QualityControl = QualityControlStage.PendingFinalQcReview,
            Stage5ReceptionClosure = ReceptionClosureStage.BalancedZeroDiff,
            Stage6DispatchClosure = DispatchClosureStage.InStoreDistribution
        });
        AssertMasterState(MasterContainerState.InQualityControlAudit, new DailyLogEntry
        {
            Stage4QualityControl = QualityControlStage.PendingFinalQcReview,
            Stage5ReceptionClosure = ReceptionClosureStage.PausedWithoutQc
        });
        AssertMasterState(MasterContainerState.PausedWithoutQcBlindSetup, new DailyLogEntry
        {
            Stage2SapIntake = SapIntakeStage.ReadyAndNotified,
            Stage4QualityControl = QualityControlStage.NoQcWaitingDock,
            Stage5ReceptionClosure = ReceptionClosureStage.PausedWithoutQc
        });
        AssertMasterState(MasterContainerState.EnteredSapAwaitingClosure, new DailyLogEntry
        {
            Stage2SapIntake = SapIntakeStage.ReadyAndNotified,
            Stage4QualityControl = QualityControlStage.NoQcWaitingDock,
            Stage5ReceptionClosure = ReceptionClosureStage.PendingClosure
        });
        AssertMasterState(MasterContainerState.SapDocReadyPendingIntake, new DailyLogEntry
        {
            Stage2SapIntake = SapIntakeStage.PendingSapIntake,
            Stage4QualityControl = QualityControlStage.NoQcWaitingDock,
            Stage5ReceptionClosure = ReceptionClosureStage.PendingClosure
        });
        AssertMasterState(MasterContainerState.InBlindCountQueue, new DailyLogEntry());
    }

    private static void AssertEnumRoundTrip<TEnum>(
        Func<TEnum, string> toExcel,
        TryParseCatalogDelegate<TEnum> tryParse) where TEnum : struct, Enum
    {
        foreach (TEnum value in Enum.GetValues<TEnum>())
        {
            string text = toExcel(value);
            Assert.True(tryParse(text, out TEnum parsed));
            Assert.Equal(value, parsed);
        }
    }

    private static void AssertMasterState(MasterContainerState expected, DailyLogEntry entry) =>
        Assert.Equal(expected, DailyLogLifecycleValidator.ComputeMasterState(entry));

    [Fact]
    public void ValidateEntry_Stage2ReadyAndNotified_AllowsRetroactiveMissingTrackingIdsWithWarning()
    {
        var retroactive = new DailyLogEntry
        {
            ContainerId = "IMP-195",
            NormalizedContainerId = "IMP-195",
            EntregaVl31n = "-",
            EwmDeliveryDoc = "-",
            HandlingUnit = "-",
            Stage1Customs = CustomsDuaStage.DuaAndDocsReady,
            Stage2SapIntake = SapIntakeStage.ReadyAndNotified,
            Stage3BlindSetup = BlindSetupStage.Pending,
            Stage4QualityControl = QualityControlStage.NoQcWaitingDock,
            Stage5ReceptionClosure = ReceptionClosureStage.PendingClosure,
            Stage6DispatchClosure = DispatchClosureStage.BlockedByReception
        };

        DailyLogValidationResult result = DailyLogLifecycleValidator.ValidateEntry(retroactive);
        Assert.True(result.IsValid);
        Assert.Empty(result.Errors);
        Assert.Contains(result.Warnings, w => w.Code == "SAP_INTAKE_MISSING_TRACKING_IDS");
    }

    [Fact]
    public void ValidateEntry_BlindSetupWithoutQc_RequiresPausaSinControlCalidad()
    {
        var invalidBlind = new DailyLogEntry
        {
            ContainerId = "IMP-242",
            Stage2SapIntake = SapIntakeStage.ReadyAndNotified,
            Stage3BlindSetup = BlindSetupStage.ReadySupplierDocOnly,
            Stage4QualityControl = QualityControlStage.NoQcWaitingDock,
            Stage5ReceptionClosure = ReceptionClosureStage.PendingClosure,
            Stage6DispatchClosure = DispatchClosureStage.BlockedByReception
        };

        DailyLogValidationResult invalidResult = DailyLogLifecycleValidator.ValidateEntry(invalidBlind);
        Assert.False(invalidResult.IsValid);
        Assert.Contains(invalidResult.Errors, e => e.Code == "BLIND_SETUP_REQUIRES_PAUSA_SIN_CC");

        DailyLogEntry validBlind = invalidBlind with { Stage5ReceptionClosure = ReceptionClosureStage.PausedWithoutQc };
        DailyLogValidationResult validResult = DailyLogLifecycleValidator.ValidateEntry(validBlind);
        Assert.DoesNotContain(validResult.Errors, e => e.Code == "BLIND_SETUP_REQUIRES_PAUSA_SIN_CC");
    }

    [Fact]
    public void ValidateEntry_Stage5Cuadrado_RequiresStage1DuaReadyAndStage4QcApproved()
    {
        var imp217Override = new DailyLogEntry
        {
            ContainerId = "IMP-217",
            Stage1Customs = CustomsDuaStage.DuaAndDocsReady,
            Stage2SapIntake = SapIntakeStage.ReadyAndNotified,
            Stage4QualityControl = QualityControlStage.PendingFinalQcReview,
            Stage5ReceptionClosure = ReceptionClosureStage.BalancedZeroDiff,
            Stage6DispatchClosure = DispatchClosureStage.BlockedByReception
        };

        DailyLogValidationResult bad = DailyLogLifecycleValidator.ValidateEntry(imp217Override);
        Assert.False(bad.IsValid);
        Assert.Contains(bad.Errors, e => e.Code == "RECEPTION_CLOSED_REQUIRES_DUA_AND_QC_APPROVED");

        DailyLogEntry clean = imp217Override with { Stage4QualityControl = QualityControlStage.QcApprovedZDescarga };
        DailyLogValidationResult ok = DailyLogLifecycleValidator.ValidateEntry(clean);
        Assert.DoesNotContain(ok.Errors, e => e.Code == "RECEPTION_CLOSED_REQUIRES_DUA_AND_QC_APPROVED");
    }

    [Fact]
    public void ValidateEntry_And_ValidateTransition_BlockStage6UnlockBeforeReceptionClosedOrPartialPartB()
    {
        var prev = new DailyLogEntry
        {
            ContainerId = "IMP-300",
            Stage1Customs = CustomsDuaStage.DuaAndDocsReady,
            Stage2SapIntake = SapIntakeStage.ReadyAndNotified,
            Stage4QualityControl = QualityControlStage.NoQcWaitingDock,
            Stage5ReceptionClosure = ReceptionClosureStage.PausedWithoutQc,
            Stage6DispatchClosure = DispatchClosureStage.BlockedByReception
        };

        DailyLogEntry blockedNext = prev with { Stage6DispatchClosure = DispatchClosureStage.PreWired80PercentReadyMov641 };
        DailyLogValidationResult blockedResult = DailyLogLifecycleValidator.ValidateTransition(prev, blockedNext);
        Assert.False(blockedResult.IsValid);
        Assert.Contains(blockedResult.Errors, e => e.Code == "DISPATCH_UNLOCKED_BEFORE_RECEPTION_READY");

        DailyLogEntry allowedPartial = prev with
        {
            Stage4QualityControl = QualityControlStage.QcApprovedZDescarga,
            Stage5ReceptionClosure = ReceptionClosureStage.PartialPendingPartB,
            Stage6DispatchClosure = DispatchClosureStage.PreWired80PercentReadyMov641
        };
        DailyLogValidationResult partialResult = DailyLogLifecycleValidator.ValidateTransition(prev, allowedPartial);
        Assert.DoesNotContain(partialResult.Errors, e => e.Code == "DISPATCH_UNLOCKED_BEFORE_RECEPTION_READY");

        DailyLogEntry allowedClosed = prev with
        {
            Stage4QualityControl = QualityControlStage.QcApprovedZDescarga,
            Stage5ReceptionClosure = ReceptionClosureStage.BalancedZeroDiff,
            Stage6DispatchClosure = DispatchClosureStage.DispatchClosureReady
        };
        DailyLogValidationResult closedResult = DailyLogLifecycleValidator.ValidateTransition(prev, allowedClosed);
        Assert.DoesNotContain(closedResult.Errors, e => e.Code == "DISPATCH_UNLOCKED_BEFORE_RECEPTION_READY");
    }

    [Theory]
    [InlineData("IMP-202", "IMP-202")]
    [InlineData("IMP 81", "IMP-81")]
    [InlineData("imp - 209", "IMP-209")]
    [InlineData("IMP244", "IMP-244")]
    [InlineData("IMP- 88", "IMP-88")]
    [InlineData("IMP-201-1", "IMP-201-1")]
    [InlineData("IMP-207 (NCR)", "IMP-207")]
    [InlineData("288", "")]
    [InlineData("#REF!", "")]
    [InlineData("IMPORT-ACTIVA", "")]
    [InlineData("-", "")]
    public void NormalizeContainerId_PreservesSubIds_AndRejectsBareNumbers(string raw, string expected)
    {
        Assert.Equal(expected, DailyLogLifecycleValidator.NormalizeContainerId(raw));
    }

    [Fact]
    public void DetectDuplicateContainers_FlagsImp202Duplicate_WhileSharedPoAcrossImp320AndImp213IsNotContainerDuplicate()
    {
        var entries = new List<DailyLogEntry>
        {
            new() { SequenceNumber = 1, ContainerId = "IMP-202", Pedido41 = "4100002100", Po43Numbers = ["4300006001"] },
            new() { SequenceNumber = 16, ContainerId = "IMP-202", Pedido41 = "4100002332", Po43Numbers = ["4300006692"] },
            new() { SequenceNumber = 18, ContainerId = "IMP-320", Pedido41 = "4100002329", Po43Numbers = ["4300007748"] },
            new() { SequenceNumber = 19, ContainerId = "IMP-213", Pedido41 = "4100002329", Po43Numbers = ["4300007748"] }
        };

        IReadOnlyList<ContainerDuplicateConflict> dupes = DailyLogLifecycleValidator.DetectDuplicateContainers(entries);
        ContainerDuplicateConflict conflict = Assert.Single(dupes);
        Assert.Equal("IMP-202", conflict.NormalizedContainerId);
        Assert.Equal(2, conflict.Entries.Count);

        IReadOnlyList<SharedDocumentReference> shared = DailyLogLifecycleValidator.DetectSharedPoOrPedido(entries);
        Assert.Contains(shared, s => s.DocumentKind == "PO43" && s.DocumentNumber == "4300007748");
        Assert.Contains(shared, s => s.DocumentKind == "Pedido41" && s.DocumentNumber == "4100002329");
    }

    [Fact]
    public void FormatSapIntakeNotification_MatchesColRAndC8C9Verbatim()
    {
        var unready = new DailyLogEntry { ContainerId = "IMP-202", EntregaVl31n = "-", Greeting = Saludo.BuenosDias };
        SapIntakeNotification notReady = DailyLogLifecycleValidator.FormatSapIntakeNotification(unready);
        Assert.False(notReady.IsReady);
        Assert.Equal("-", notReady.ColRValue);

        var ready202 = new DailyLogEntry { ContainerId = "IMP-202", EntregaVl31n = "180060834", Greeting = Saludo.BuenosDias };
        SapIntakeNotification notif202 = DailyLogLifecycleValidator.FormatSapIntakeNotification(ready202);
        Assert.True(notif202.IsReady);
        Assert.Equal("Buenos días, Se informa que el ingreso a nivel de sistema del Contenedor 202, quedo listo.", notif202.ColRValue);
        Assert.Equal("Ingreso Listo a Sistema - Contenedor 202", notif202.EmailSubject);
        Assert.Equal("Buenos días\nSe informa que el ingreso a nivel de sistema del Contenedor 202, quedo listo.", notif202.ClipboardBody);

        var ready201 = new DailyLogEntry { ContainerId = "IMP-201-1", EntregaVl31n = "180060900", Greeting = Saludo.BuenasTardes };
        SapIntakeNotification notif201 = DailyLogLifecycleValidator.FormatSapIntakeNotification(ready201);
        Assert.Equal("Ingreso Listo a Sistema - Contenedor 201-1", notif201.EmailSubject);
        Assert.Equal("Buenas tardes, Se informa que el ingreso a nivel de sistema del Contenedor 201-1, quedo listo.", notif201.ColRValue);
    }

    [Fact]
    public void FormatSapIntakeNotification_InvalidContainerId_ReturnsNotReady()
    {
        var invalidContainer = new DailyLogEntry
        {
            ContainerId = "288",
            EntregaVl31n = "180060900",
            Greeting = Saludo.BuenosDias
        };

        SapIntakeNotification notification = DailyLogLifecycleValidator.FormatSapIntakeNotification(invalidContainer);
        Assert.False(notification.IsReady);
        Assert.Equal("-", notification.ColRValue);
        Assert.Equal(string.Empty, notification.EmailSubject);
        Assert.Equal(string.Empty, notification.ClipboardBody);
    }
}
