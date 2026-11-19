using Acr.Core.Models;
using Acr.Core.Services;

namespace Acr.Tests.Core;

public sealed class OperationalSafetyPolicyTests
{
    [Theory]
    [InlineData(OperationalSafetyMode.Development, OperationalActionKind.WriteAutomatedOrBackground)]
    [InlineData(OperationalSafetyMode.Development, OperationalActionKind.WriteExplicitUserSaveOrSync)]
    [InlineData(OperationalSafetyMode.Development, OperationalActionKind.WriteToNewExportPath)]
    [InlineData(OperationalSafetyMode.Testing, OperationalActionKind.WriteAutomatedOrBackground)]
    [InlineData(OperationalSafetyMode.Testing, OperationalActionKind.WriteExplicitUserSaveOrSync)]
    [InlineData(OperationalSafetyMode.Testing, OperationalActionKind.WriteToNewExportPath)]
    [InlineData(OperationalSafetyMode.Production, OperationalActionKind.WriteAutomatedOrBackground)]
    [InlineData(OperationalSafetyMode.Production, OperationalActionKind.WriteExplicitUserSaveOrSync)]
    [InlineData(OperationalSafetyMode.Production, OperationalActionKind.WriteToNewExportPath)]
    public void ProductionSourceWorkbook_IsAlwaysReadOnly_AcrossAllModes(
        OperationalSafetyMode mode,
        OperationalActionKind action)
    {
        Assert.False(OperationalSafetyPolicy.CanWriteProductionSource(mode));

        SafetyPolicyDecision decision = OperationalSafetyPolicy.EvaluateAccess(
            mode, OperationalTargetKind.ProductionSourceWorkbook, action);

        Assert.False(decision.IsAllowed);
        Assert.Equal("PRODUCTION_SOURCE_READ_ONLY", decision.ReasonCode);
        Assert.Equal(mode == OperationalSafetyMode.Testing, decision.MustRedirectToTemp);
    }

    [Fact]
    public void OperationalDailyLog_WritableOnlyInProductionWithExplicitUserAction()
    {
        Assert.False(OperationalSafetyPolicy.CanWriteOperationalDailyLog(OperationalSafetyMode.Development, true));
        Assert.False(OperationalSafetyPolicy.CanWriteOperationalDailyLog(OperationalSafetyMode.Testing, true));
        Assert.False(OperationalSafetyPolicy.CanWriteOperationalDailyLog(OperationalSafetyMode.Production, false));
        Assert.True(OperationalSafetyPolicy.CanWriteOperationalDailyLog(OperationalSafetyMode.Production, true));

        SafetyPolicyDecision devDecision = OperationalSafetyPolicy.EvaluateAccess(
            OperationalSafetyMode.Development, OperationalTargetKind.OperationalDailyLog, OperationalActionKind.WriteExplicitUserSaveOrSync);
        Assert.False(devDecision.IsAllowed);
        Assert.Equal("DEV_DAILY_LOG_READ_ONLY", devDecision.ReasonCode);

        SafetyPolicyDecision testDecision = OperationalSafetyPolicy.EvaluateAccess(
            OperationalSafetyMode.Testing, OperationalTargetKind.OperationalDailyLog, OperationalActionKind.WriteExplicitUserSaveOrSync);
        Assert.False(testDecision.IsAllowed);
        Assert.True(testDecision.MustRedirectToTemp);
        Assert.Equal("TESTING_REDIRECT_TO_TEMP", testDecision.ReasonCode);

        SafetyPolicyDecision prodBgDecision = OperationalSafetyPolicy.EvaluateAccess(
            OperationalSafetyMode.Production, OperationalTargetKind.OperationalDailyLog, OperationalActionKind.WriteAutomatedOrBackground);
        Assert.False(prodBgDecision.IsAllowed);
        Assert.Equal("PROD_REQUIRES_EXPLICIT_USER_ACTION", prodBgDecision.ReasonCode);

        SafetyPolicyDecision prodExplicitDecision = OperationalSafetyPolicy.EvaluateAccess(
            OperationalSafetyMode.Production, OperationalTargetKind.OperationalDailyLog, OperationalActionKind.WriteExplicitUserSaveOrSync);
        Assert.True(prodExplicitDecision.IsAllowed);
        Assert.False(prodExplicitDecision.MustRedirectToTemp);
        Assert.Equal("PROD_EXPLICIT_SAVE_ALLOWED", prodExplicitDecision.ReasonCode);
    }

    [Fact]
    public void TestingMode_RedirectsClosureExportsAndReportsToTemp()
    {
        Assert.True(OperationalSafetyPolicy.MustRedirectWritesToTemp(OperationalSafetyMode.Testing));
        Assert.False(OperationalSafetyPolicy.MustRedirectWritesToTemp(OperationalSafetyMode.Development));
        Assert.False(OperationalSafetyPolicy.MustRedirectWritesToTemp(OperationalSafetyMode.Production));

        SafetyPolicyDecision exportDecision = OperationalSafetyPolicy.EvaluateAccess(
            OperationalSafetyMode.Testing, OperationalTargetKind.ClosureExportOutput, OperationalActionKind.WriteToNewExportPath);
        Assert.True(exportDecision.IsAllowed);
        Assert.True(exportDecision.MustRedirectToTemp);
        Assert.Equal("TESTING_EXPORT_ISOLATED_TEMP", exportDecision.ReasonCode);

        SafetyPolicyDecision reportDecision = OperationalSafetyPolicy.EvaluateAccess(
            OperationalSafetyMode.Testing, OperationalTargetKind.LocalReportArtifact, OperationalActionKind.WriteAutomatedOrBackground);
        Assert.True(reportDecision.IsAllowed);
        Assert.True(reportDecision.MustRedirectToTemp);
    }

    [Fact]
    public void NormalizeComparablePath_And_IsUnderProtectedRoot_HandleSlashesCaseAndTrailingSeparators()
    {
        string normalized = OperationalSafetyPolicy.NormalizeComparablePath("  D:/Development/Software Development/data/IMP-81/  ");
        Assert.Equal(@"D:\DEVELOPMENT\SOFTWARE DEVELOPMENT\DATA\IMP-81", normalized);

        Assert.True(OperationalSafetyPolicy.IsUnderProtectedRoot(
            "D:/Development/Software Development/3 Way Inventory Reconciliation/data/IMP-81/IMP-81.xlsm",
            @"D:\Development\Software Development\3 Way Inventory Reconciliation\data\"));

        Assert.False(OperationalSafetyPolicy.IsUnderProtectedRoot(
            @"D:\Development\Software Development\3 Way Inventory Reconciliation\data_backup\IMP-81.xlsm",
            @"D:\Development\Software Development\3 Way Inventory Reconciliation\data"));

        Assert.True(OperationalSafetyPolicy.IsSameProtectedFile(
            "D:/Development/Control_Diario_Operaciones_Jostin.xlsx",
            @"d:\development\Control_Diario_Operaciones_Jostin.xlsx"));
    }

    [Fact]
    public void NormalizeComparablePath_CollapsesDotAndDotDotSegments_AndBlocksTraversalBypass()
    {
        Assert.True(OperationalSafetyPolicy.IsUnderProtectedRoot(
            @"D:\Development\outside\..\data\IMP-81.xlsm",
            @"D:\Development\data"));

        Assert.False(OperationalSafetyPolicy.IsUnderProtectedRoot(
            @"D:\Development\data\..\outside\file.xlsm",
            @"D:\Development\data"));

        Assert.True(OperationalSafetyPolicy.IsSameProtectedFile(
            @"D:\Development\sub\..\Control_Diario_Operaciones_Jostin.xlsx",
            @"D:\Development\Control_Diario_Operaciones_Jostin.xlsx"));

        Assert.True(OperationalSafetyPolicy.IsSameProtectedFile(
            @"D:\Development\.\Control_Diario_Operaciones_Jostin.xlsx",
            @"D:\Development\Control_Diario_Operaciones_Jostin.xlsx"));

        Assert.Equal(@"D:\DATA\IMP-81.XLSM", OperationalSafetyPolicy.NormalizeComparablePath(@"D:\..\..\data\IMP-81.xlsm"));
        Assert.Equal(@"\\SERVER\SHARE\DATA\IMP-81.XLSM", OperationalSafetyPolicy.NormalizeComparablePath(@"\\server\share\..\..\outside\..\data\IMP-81.xlsm"));
        Assert.True(OperationalSafetyPolicy.IsUnderProtectedRoot(@"D:\data\IMP-81.xlsm", @"D:\"));
    }
}
