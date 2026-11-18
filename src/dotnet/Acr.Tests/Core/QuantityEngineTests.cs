using Acr.Core.Models;
using Acr.Core.Services;

namespace Acr.Tests.Core;

public sealed class QuantityEngineTests
{
    [Fact]
    public void TestSingleTermExpression()
    {
        ParsedPackaging res = QuantityEngine.ParsePackagingExpression("50*12");
        Assert.Equal(50, res.TotalCartons);
        Assert.Equal(600, res.TotalUnits);
        Assert.Equal(50, res.FullCartons);
        Assert.Equal(12, res.UnitsPerFull);
        Assert.False(res.IsCompound);
        Assert.False(res.HasLoose);
    }

    [Fact]
    public void TestScalarExpression()
    {
        ParsedPackaging res = QuantityEngine.ParsePackagingExpression("192");
        Assert.Equal(1, res.TotalCartons);
        Assert.Equal(192, res.TotalUnits);
        Assert.Equal(1, res.FullCartons);
        Assert.Equal(192, res.UnitsPerFull);
        Assert.False(res.IsCompound);
    }

    [Fact]
    public void TestXDelimiter()
    {
        ParsedPackaging res = QuantityEngine.ParsePackagingExpression("10x24");
        Assert.Equal(10, res.TotalCartons);
        Assert.Equal(240, res.TotalUnits);
    }

    [Fact]
    public void TestCompoundTwoTerms()
    {
        ParsedPackaging res = QuantityEngine.ParsePackagingExpression("3*192 3*204");
        Assert.Equal(6, res.TotalCartons);
        Assert.Equal(1188, res.TotalUnits);
        Assert.Equal(3, res.FullCartons);
        Assert.Equal(192, res.UnitsPerFull);
        Assert.Equal(3, res.ComplementCartons);
        Assert.Equal(204, res.UnitsPerComplement);
        Assert.True(res.IsCompound);
        Assert.False(res.HasLoose);
    }

    [Fact]
    public void TestCompoundWithLooseAdditionRow37()
    {
        ParsedPackaging res = QuantityEngine.ParsePackagingExpression("6*192 +48");
        Assert.Equal(6, res.TotalCartons);
        Assert.Equal(1200, res.TotalUnits);
        Assert.Equal(6, res.FullCartons);
        Assert.Equal(192, res.UnitsPerFull);
        Assert.Equal(48, res.LooseUnits);
        Assert.True(res.IsCompound);
        Assert.True(res.HasLoose);
    }

    [Fact]
    public void TestCompoundWithPlusBetweenTerms()
    {
        ParsedPackaging res = QuantityEngine.ParsePackagingExpression("40*60 + 1*72");
        Assert.Equal(41, res.TotalCartons);
        Assert.Equal(2472, res.TotalUnits);
        Assert.Equal(40, res.FullCartons);
        Assert.Equal(60, res.UnitsPerFull);
        Assert.Equal(1, res.ComplementCartons);
        Assert.Equal(72, res.UnitsPerComplement);
        Assert.Equal(0, res.LooseUnits);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("None")]
    [InlineData("nan")]
    [InlineData("null")]
    [InlineData("   ")]
    public void TestEmptyAndNullExpressions(string? raw)
    {
        Assert.Equal(0, QuantityEngine.ParsePackagingExpression(raw).TotalUnits);
    }

    [Fact]
    public void TestStandardFullCartons()
    {
        CartonCalculationResult res = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 10,
            UnitsPerFullCarton = 24,
            Curve = 12,
            Uom = "PRK"
        });

        Assert.Equal(240, res.CalculatedUnits);
        Assert.Equal(10, res.CalculatedCartons);
        Assert.Equal(20, res.CalculatedPacks);
        Assert.Equal(0, res.LooseUnits);
        Assert.Empty(res.Flags);
    }

    [Fact]
    public void TestCompoundCartonsRow1()
    {
        CartonCalculationResult res = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 3,
            UnitsPerFullCarton = 192,
            ComplementCartons = 3,
            UnitsInComplement = 204,
            ReportedTotalCartons = 6,
            ReportedTotalUnits = 1188,
            Curve = 12,
            Uom = "PRK"
        });

        Assert.Equal(1188, res.CalculatedUnits);
        Assert.Equal(6, res.CalculatedCartons);
        Assert.Equal(99, res.CalculatedPacks);
        Assert.Equal(0, res.LooseUnits);
        Assert.Equal(0, res.FormulaDiscrepancy);
        Assert.Empty(res.Flags);
    }

    [Fact]
    public void TestRow37ExcelZeroMultiplierBugRescuing()
    {
        CartonCalculationResult res = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 6,
            UnitsPerFullCarton = 192,
            ComplementCartons = null,
            UnitsInComplement = 48,
            ReportedTotalCartons = 6,
            ReportedTotalUnits = 1200,
            ReportedDiff = -48,
            Curve = 12,
            Uom = "PRK"
        });

        Assert.Equal(1200, res.CalculatedUnits);
        Assert.Equal(6, res.CalculatedCartons);
        Assert.Equal(100, res.CalculatedPacks);
        Assert.Equal(0, res.LooseUnits);
        Assert.Equal(0, res.FormulaDiscrepancy);
        Assert.Contains(QuantityEngineFlags.FlagQcComplementWithoutCartonCount, res.Flags);
        Assert.Contains(QuantityEngineFlags.FlagNegativeReportedDiff, res.Flags);
        Assert.Contains("Alerta: 48 unidades complementarias rescatadas", res.Explanation);
    }

    [Fact]
    public void TestVarianteSueltaRow()
    {
        CartonCalculationResult res = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = null,
            ReportedTotalCartons = 0,
            ReportedTotalUnits = 10,
            Observation = "VARIANTE SUELTA",
            Uom = "UN"
        });

        Assert.Equal(10, res.CalculatedUnits);
        Assert.Equal(0, res.CalculatedCartons);
        Assert.Equal(10, res.CalculatedPacks);
        Assert.Contains(QuantityEngineFlags.FlagVarianteSueltaRow, res.Flags);
    }

    [Fact]
    public void TestSharedCartonPair()
    {
        CartonCalculationResult primary = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 1,
            UnitsPerFullCarton = 300,
            ReportedTotalCartons = 1,
            ReportedTotalUnits = 300,
            Observation = "DOS MATERIALES POR BULTO",
            Uom = "UN"
        });

        CartonCalculationResult secondary = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = null,
            UnitsPerFullCarton = 300,
            ReportedTotalCartons = 0,
            ReportedTotalUnits = 300,
            Observation = "DOS MATERIALES POR BULTO",
            Uom = "UN"
        });

        Assert.Equal(300, primary.CalculatedUnits);
        Assert.Equal(1, primary.CalculatedCartons);
        Assert.Contains(QuantityEngineFlags.FlagSharedCartonPrimary, primary.Flags);

        Assert.Equal(300, secondary.CalculatedUnits);
        Assert.Equal(0, secondary.CalculatedCartons);
        Assert.Contains(QuantityEngineFlags.FlagSharedCartonSecondary, secondary.Flags);

        Assert.Equal(1, primary.CalculatedCartons + secondary.CalculatedCartons);
        Assert.Equal(600, primary.CalculatedUnits + secondary.CalculatedUnits);
    }

    [Fact]
    public void TestUnevenPrepackCurve()
    {
        CartonCalculationResult res = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 15,
            UnitsPerFullCarton = 48,
            ComplementCartons = 1,
            UnitsInComplement = 22,
            Curve = 12,
            Uom = "PRK"
        });

        Assert.Equal(742, res.CalculatedUnits);
        Assert.Equal(16, res.CalculatedCartons);
        Assert.Equal(61, res.CalculatedPacks);
        Assert.Equal(10, res.LooseUnits);
        Assert.Contains(QuantityEngineFlags.FlagUnevenPrepackCurve, res.Flags);
    }

    [Fact]
    public void TestFormulaMismatchDetection()
    {
        CartonCalculationResult res = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 5,
            UnitsPerFullCarton = 100,
            ReportedTotalUnits = 450,
            Uom = "UN"
        });

        Assert.Equal(500, res.CalculatedUnits);
        Assert.Equal(50, res.FormulaDiscrepancy);
        Assert.Contains(QuantityEngineFlags.FlagQcFormulaMismatch, res.Flags);
    }

    [Theory]
    [InlineData(5, "MAS 1PK, 5685920")]
    [InlineData(9, "MENOS 1PK")]
    [InlineData(4, "+1pk")]
    [InlineData(4, "menos 1 pk")]
    [InlineData(4, "-1PK")]
    [InlineData(4, "MAS 1 PK")]
    public void TestObservationAdjustmentKeywords(int cartons, string observation)
    {
        CartonCalculationResult res = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = cartons,
            UnitsPerFullCarton = 192,
            Observation = observation
        });

        Assert.Contains(QuantityEngineFlags.FlagQcObservationAdjustment, res.Flags);
    }

    [Fact]
    public void TestEmptyComplementCartonsFlag()
    {
        CartonCalculationResult res = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 10,
            UnitsPerFullCarton = 50,
            ComplementCartons = 2,
            UnitsInComplement = 0
        });

        Assert.Equal(12, res.CalculatedCartons);
        Assert.Equal(500, res.CalculatedUnits);
        Assert.Contains(QuantityEngineFlags.FlagEmptyComplementCartons, res.Flags);
    }

    [Fact]
    public void TestFallbackToPackagingString()
    {
        CartonCalculationResult res = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            PackagingString = "6*192 +48",
            Curve = 12,
            Uom = "PRK"
        });

        Assert.Equal(1200, res.CalculatedUnits);
        Assert.Equal(6, res.CalculatedCartons);
        Assert.Equal(100, res.CalculatedPacks);
        Assert.Equal(0, res.LooseUnits);
        Assert.NotNull(res.FallbackPackaging);
    }

    [Fact]
    public void ComplementRescue_APlusOneAndZeroFullCartons_MatchesImp298Parity()
    {
        CartonCalculationResult plusOne = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 4,
            UnitsPerFullCarton = 60,
            ComplementCartons = null,
            UnitsInComplement = 24,
            ReportedTotalCartons = 5
        });

        Assert.Equal(5, plusOne.CalculatedCartons);
        Assert.Equal(264, plusOne.CalculatedUnits);
        Assert.Contains(QuantityEngineFlags.FlagQcComplementWithoutCartonCount, plusOne.Flags);

        CartonCalculationResult imp298ZeroFull = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 0,
            UnitsPerFullCarton = 0,
            ComplementCartons = null,
            UnitsInComplement = 24,
            ReportedTotalCartons = null
        });

        Assert.Equal(1, imp298ZeroFull.CalculatedCartons);
        Assert.Equal(24, imp298ZeroFull.CalculatedUnits);
        Assert.Contains(QuantityEngineFlags.FlagQcComplementWithoutCartonCount, imp298ZeroFull.Flags);
    }

    [Fact]
    public void WhitespaceAndUnicodeSeparators_AndThreePlusTerms_ParseAccurately()
    {
        ParsedPackaging normalizedX = QuantityEngine.ParsePackagingExpression("  10 X 24 \u00A0 ");
        Assert.Equal(10, normalizedX.TotalCartons);
        Assert.Equal(240, normalizedX.TotalUnits);

        ParsedPackaging tabbed = QuantityEngine.ParsePackagingExpression("6\t*\n192\u00A0+\r48");
        Assert.Equal(6, tabbed.TotalCartons);
        Assert.Equal(1200, tabbed.TotalUnits);
        Assert.Equal(48, tabbed.LooseUnits);

        ParsedPackaging unicodeTimes = QuantityEngine.ParsePackagingExpression("6\u00D7192 +48");
        Assert.Equal(6, unicodeTimes.TotalCartons);
        Assert.Equal(1200, unicodeTimes.TotalUnits);

        ParsedPackaging threeTerms = QuantityEngine.ParsePackagingExpression("10*12 5*6 1*3");
        Assert.Equal(16, threeTerms.TotalCartons);
        Assert.Equal(153, threeTerms.TotalUnits);
        Assert.Equal(10, threeTerms.FullCartons);
        Assert.Equal(12, threeTerms.UnitsPerFull);
        Assert.Equal(5, threeTerms.ComplementCartons);
        Assert.Equal(6, threeTerms.UnitsPerComplement);
        Assert.Equal(3, threeTerms.Terms.Count);
    }

    [Fact]
    public void Imp192RepetitionAmbiguity_PreservesParityAndFlagsForCrossCheck()
    {
        ParsedPackaging imp192 = QuantityEngine.ParsePackagingExpression("56*12+672");
        Assert.Equal(56, imp192.TotalCartons);
        Assert.Equal(1344, imp192.TotalUnits);
        Assert.Equal(672, imp192.LooseUnits);
        Assert.True(imp192.HasTotalRepetitionAmbiguity);

        CartonCalculationResult calc = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            PackagingString = "56*12+672"
        });
        Assert.Contains(QuantityEngineFlags.FlagPackagingTotalRepetitionAmbiguity, calc.Flags);

        ParsedPackaging imp200 = QuantityEngine.ParsePackagingExpression("8*72+24");
        Assert.Equal(8, imp200.TotalCartons);
        Assert.Equal(600, imp200.TotalUnits);
        Assert.Equal(24, imp200.LooseUnits);
        Assert.False(imp200.HasTotalRepetitionAmbiguity);

        ParsedPackaging tripleMultiply = QuantityEngine.ParsePackagingExpression("12*32*384");
        Assert.Equal(12, tripleMultiply.TotalCartons);
        Assert.Equal(384, tripleMultiply.TotalUnits);
        Assert.Equal(0, tripleMultiply.LooseUnits);
    }

    [Fact]
    public void CheckedOverflowProtection_ThrowsOverflowExceptionOnExtremeValues()
    {
        Assert.Throws<OverflowException>(() => QuantityEngine.ParsePackagingExpression("100000*100000"));
        Assert.Throws<OverflowException>(() => QuantityEngine.ParsePackagingExpression("99999999999"));
        Assert.Throws<OverflowException>(() => QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 100_000,
            UnitsPerFullCarton = 100_000
        }));
    }

    [Fact]
    public void SharedCartonAndVarianteSueltaEdgeCases_MatchPythonOrTruthiness()
    {
        CartonCalculationResult sharedZeroCartons = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            FullCartons = 0,
            UnitsPerFullCarton = 300,
            ReportedTotalUnits = 0,
            Observation = "DOS MATERIALES POR BULTO"
        });
        Assert.Equal(0, sharedZeroCartons.CalculatedCartons);
        Assert.Equal(300, sharedZeroCartons.CalculatedUnits);
        Assert.Equal(0, sharedZeroCartons.FormulaDiscrepancy);
        Assert.Contains(QuantityEngineFlags.FlagSharedCartonSecondary, sharedZeroCartons.Flags);

        CartonCalculationResult varianteFlagOnly = QuantityEngine.CalculateCartonQuantities(new CartonCalculationInput
        {
            IsVarianteSuelta = true,
            ReportedTotalUnits = 0,
            UnitsInComplement = 7,
            Observation = null
        });
        Assert.Equal(7, varianteFlagOnly.CalculatedUnits);
        Assert.Equal(0, varianteFlagOnly.CalculatedCartons);
        Assert.Contains(QuantityEngineFlags.FlagVarianteSueltaRow, varianteFlagOnly.Flags);
    }
}
