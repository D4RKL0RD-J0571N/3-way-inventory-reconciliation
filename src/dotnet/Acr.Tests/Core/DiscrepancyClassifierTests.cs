using Acr.Core.Models;
using Acr.Core.Services;

namespace Acr.Tests.Core;

public sealed class DiscrepancyClassifierTests
{
    [Fact]
    public void TestExactMatch()
    {
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(
            [new DeclaredItem { MaterialCode = "1001", DeclaredUnits = 500, DeclaredCartons = 10 }],
            [new PhysicalItem { MaterialCode = "1001", PhysicalUnits = 500, PhysicalCartons = 10 }]);

        Assert.Equal(1, summary.CleanMaterialsCount);
        Assert.Equal(0, summary.DiscrepantMaterialsCount);
        Assert.Equal(0, summary.TotalShortageUnits);
        Assert.Equal(0, summary.TotalOverageUnits);
        Assert.Equal(DiscrepancyType.ExactMatch, summary.Items[0].DiscrepancyType);
        Assert.Equal(0, summary.Items[0].DifferenceUnits);
    }

    [Fact]
    public void TestCardinalRuleNoOffsetting()
    {
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(
            [
                new DeclaredItem { MaterialCode = "1001", DeclaredUnits = 100, DeclaredCartons = 2 },
                new DeclaredItem { MaterialCode = "1002", DeclaredUnits = 100, DeclaredCartons = 2 }
            ],
            [
                new PhysicalItem { MaterialCode = "1001", PhysicalUnits = 80, PhysicalCartons = 2 },
                new PhysicalItem { MaterialCode = "1002", PhysicalUnits = 120, PhysicalCartons = 2 }
            ]);

        Assert.Equal(0, summary.NetUnitDifference);
        Assert.Equal(20, summary.TotalShortageUnits);
        Assert.Equal(20, summary.TotalOverageUnits);
        Assert.Equal(2, summary.DiscrepantMaterialsCount);
        Assert.Equal(1, summary.RevueltoPairsCount);

        DiscrepancyItem item1001 = summary.Items.Single(i => i.MaterialCode == "1001");
        DiscrepancyItem item1002 = summary.Items.Single(i => i.MaterialCode == "1002");

        Assert.Equal(DiscrepancyType.RevueltoCandidate, item1001.DiscrepancyType);
        Assert.Equal("1002", item1001.RevueltoPartner);
        Assert.Equal(RevueltoConfidence.Possible, item1001.RevueltoConfidence);

        Assert.Equal(DiscrepancyType.RevueltoCandidate, item1002.DiscrepancyType);
        Assert.Equal("1001", item1002.RevueltoPartner);
        Assert.Equal(RevueltoConfidence.Possible, item1002.RevueltoConfidence);
    }

    [Fact]
    public void TestUnexpectedItemReceived()
    {
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(
            [new DeclaredItem { MaterialCode = "1001", DeclaredUnits = 100, DeclaredCartons = 2 }],
            [
                new PhysicalItem { MaterialCode = "1001", PhysicalUnits = 100, PhysicalCartons = 2 },
                new PhysicalItem { MaterialCode = "9999", PhysicalUnits = 50, PhysicalCartons = 1 }
            ]);

        DiscrepancyItem unexpected = summary.Items.Single(i => i.MaterialCode == "9999");
        Assert.Equal(DiscrepancyType.UnexpectedItem, unexpected.DiscrepancyType);
        Assert.Equal(50, unexpected.DifferenceUnits);
        Assert.Equal(50, summary.TotalOverageUnits);
    }

    [Fact]
    public void TestMissingItemCompletely()
    {
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(
            [
                new DeclaredItem { MaterialCode = "1001", DeclaredUnits = 100, DeclaredCartons = 2 },
                new DeclaredItem { MaterialCode = "1002", DeclaredUnits = 200, DeclaredCartons = 4 }
            ],
            [new PhysicalItem { MaterialCode = "1001", PhysicalUnits = 100, PhysicalCartons = 2 }]);

        DiscrepancyItem missing = summary.Items.Single(i => i.MaterialCode == "1002");
        Assert.Equal(DiscrepancyType.MissingItem, missing.DiscrepancyType);
        Assert.Equal(-200, missing.DifferenceUnits);
        Assert.Equal(200, summary.TotalShortageUnits);
    }

    [Fact]
    public void TestCartonMismatchOnly()
    {
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(
            [new DeclaredItem { MaterialCode = "1001", DeclaredUnits = 100, DeclaredCartons = 5 }],
            [new PhysicalItem { MaterialCode = "1001", PhysicalUnits = 100, PhysicalCartons = 4 }]);

        Assert.Equal(DiscrepancyType.CartonMismatchOnly, summary.Items[0].DiscrepancyType);
        Assert.Equal(0, summary.Items[0].DifferenceUnits);
        Assert.Equal(-1, summary.Items[0].DifferenceCartons);
        Assert.Contains("(-1 bultos)", summary.Items[0].Explanation);
    }

    [Fact]
    public void TestDamageOnly()
    {
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(
            [new DeclaredItem { MaterialCode = "1001", DeclaredUnits = 100, DeclaredCartons = 2 }],
            [new PhysicalItem { MaterialCode = "1001", PhysicalUnits = 100, PhysicalCartons = 2, DamagedUnits = 5 }]);

        Assert.Equal(DiscrepancyType.DamageOnly, summary.Items[0].DiscrepancyType);
        Assert.Equal(5, summary.Items[0].DamagedUnits);
        Assert.Equal(5, summary.TotalDamagedUnits);
    }

    [Fact]
    public void MultiPoDuplicateSku_AggregatesWithoutSilentOverwrite()
    {
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(
            [
                new DeclaredItem { MaterialCode = "M1", OrderPo43 = "4300001", DeclaredUnits = 100, DeclaredCartons = 2 },
                new DeclaredItem { MaterialCode = "M1", OrderPo43 = "4300002", DeclaredUnits = 50, DeclaredCartons = 1 }
            ],
            [new PhysicalItem { MaterialCode = "M1", PhysicalUnits = 150, PhysicalCartons = 3 }]);

        DiscrepancyItem item = Assert.Single(summary.Items);
        Assert.Equal(DiscrepancyType.ExactMatch, item.DiscrepancyType);
        Assert.Equal(150, item.DeclaredUnits);
        Assert.Equal(150, summary.TotalDeclaredUnits);
        Assert.True(item.HasMultiplePoLines);
        Assert.Equal(2, item.PoLineCount);
        Assert.Equal("4300001, 4300002", item.OrderPo43);
        Assert.Equal(2, item.PoBreakdown.Count);
        Assert.Contains("M1", summary.MultiPoMaterials);
        Assert.Contains("MULTIPLE_PO_LINES_AGGREGATED", item.Flags);
    }

    [Fact]
    public void SplitPalletPhysicalItems_AggregateCleanly()
    {
        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(
            [new DeclaredItem { MaterialCode = "M1", DeclaredUnits = 100, DeclaredCartons = 2 }],
            [
                new PhysicalItem { MaterialCode = "M1", PhysicalUnits = 60, PhysicalCartons = 1, Flags = ["F1"] },
                new PhysicalItem { MaterialCode = "M1", PhysicalUnits = 40, PhysicalCartons = 1, Flags = ["F2"] }
            ]);

        DiscrepancyItem item = Assert.Single(summary.Items);
        Assert.Equal(DiscrepancyType.ExactMatch, item.DiscrepancyType);
        Assert.Equal(100, item.PhysicalUnits);
        Assert.Equal(2, item.PhysicalCartons);
        Assert.Equal(["F1", "F2"], item.Flags);
    }

    [Fact]
    public void RevueltoConfidence_EvaluatesConfirmedAndProbable_ByVendorCategoryAndCarton()
    {
        ComparisonSummary confirmedVendor = DiscrepancyClassifier.CompareMerchandise(
            [
                new DeclaredItem { MaterialCode = "S1", VendorMaterial = "STYLE-A", DeclaredUnits = 24, DeclaredCartons = 1 },
                new DeclaredItem { MaterialCode = "O1", VendorMaterial = "STYLE-A", DeclaredUnits = 0, DeclaredCartons = 0 }
            ],
            [
                new PhysicalItem { MaterialCode = "S1", VendorMaterial = "STYLE-A", PhysicalUnits = 12, PhysicalCartons = 1 },
                new PhysicalItem { MaterialCode = "O1", VendorMaterial = "STYLE-A", PhysicalUnits = 12, PhysicalCartons = 1 }
            ]);
        Assert.Equal(RevueltoConfidence.Confirmed, confirmedVendor.RevueltoMatches[0].Confidence);

        ComparisonSummary confirmedCatAndCarton = DiscrepancyClassifier.CompareMerchandise(
            [
                new DeclaredItem { MaterialCode = "S1", Category = "G40", DeclaredUnits = 24, DeclaredCartons = 1 },
                new DeclaredItem { MaterialCode = "O1", Category = "G40", DeclaredUnits = 24, DeclaredCartons = 1 }
            ],
            [
                new PhysicalItem { MaterialCode = "S1", Category = "G40", PhysicalUnits = 0, PhysicalCartons = 0 },
                new PhysicalItem { MaterialCode = "O1", Category = "G40", PhysicalUnits = 48, PhysicalCartons = 2 }
            ]);
        Assert.Equal(RevueltoConfidence.Confirmed, confirmedCatAndCarton.RevueltoMatches[0].Confidence);

        ComparisonSummary probableCatOnly = DiscrepancyClassifier.CompareMerchandise(
            [
                new DeclaredItem { MaterialCode = "S1", Category = "G40", DeclaredUnits = 20, DeclaredCartons = 1 },
                new DeclaredItem { MaterialCode = "O1", Category = "G40", DeclaredUnits = 20, DeclaredCartons = 1 }
            ],
            [
                new PhysicalItem { MaterialCode = "S1", Category = "G40", PhysicalUnits = 10, PhysicalCartons = 1 },
                new PhysicalItem { MaterialCode = "O1", Category = "G40", PhysicalUnits = 30, PhysicalCartons = 1 }
            ]);
        Assert.Equal(RevueltoConfidence.Probable, probableCatOnly.RevueltoMatches[0].Confidence);
    }

    [Fact]
    public void RevueltoConfidence_BreaksTiesDeterministicallyByConfidenceOverAlphabeticalOrder()
    {
        ComparisonSummary tieBreak = DiscrepancyClassifier.CompareMerchandise(
            [
                new DeclaredItem { MaterialCode = "S1", VendorMaterial = "V-MATCH", DeclaredUnits = 50, DeclaredCartons = 1 },
                new DeclaredItem { MaterialCode = "O_AAA", VendorMaterial = "OTHER", DeclaredUnits = 10, DeclaredCartons = 1 },
                new DeclaredItem { MaterialCode = "O_ZZZ", VendorMaterial = "V-MATCH", DeclaredUnits = 10, DeclaredCartons = 1 }
            ],
            [
                new PhysicalItem { MaterialCode = "S1", VendorMaterial = "V-MATCH", PhysicalUnits = 30, PhysicalCartons = 1 },
                new PhysicalItem { MaterialCode = "O_AAA", VendorMaterial = "OTHER", PhysicalUnits = 30, PhysicalCartons = 1 },
                new PhysicalItem { MaterialCode = "O_ZZZ", VendorMaterial = "V-MATCH", PhysicalUnits = 30, PhysicalCartons = 1 }
            ]);
        Assert.Equal("O_ZZZ", tieBreak.RevueltoMatches[0].OverageMaterialCode);
        Assert.Equal(RevueltoConfidence.Confirmed, tieBreak.RevueltoMatches[0].Confidence);
    }

    [Fact]
    public void BultoCero_AndClassificationPrecedence_AndDecimalCost_AndOverflow_AreEnforced()
    {
        var decl = new DeclaredItem
        {
            MaterialCode = "M1",
            DeclaredUnits = 100,
            DeclaredCartons = 2,
            UnitCost = 3.1666666666666665m
        };
        Assert.Equal(3.1666666666666665m, decl.UnitCost);

        ComparisonSummary bultoCero = DiscrepancyClassifier.CompareMerchandise(
            [decl],
            [new PhysicalItem { MaterialCode = "M1", PhysicalUnits = 0, PhysicalCartons = 0 }]);
        Assert.Equal(DiscrepancyType.MissingItem, bultoCero.Items[0].DiscrepancyType);
        Assert.Equal(0, bultoCero.Items[0].PhysicalCartons);

        ComparisonSummary precedence = DiscrepancyClassifier.CompareMerchandise(
            [
                new DeclaredItem { MaterialCode = "A", DeclaredUnits = 10, DeclaredCartons = 2 },
                new DeclaredItem { MaterialCode = "B", DeclaredUnits = 10, DeclaredCartons = 2 }
            ],
            [
                new PhysicalItem { MaterialCode = "A", PhysicalUnits = 8, PhysicalCartons = 2, DamagedUnits = 3 },
                new PhysicalItem { MaterialCode = "B", PhysicalUnits = 15, PhysicalCartons = 3 }
            ]);
        Assert.Equal(DiscrepancyType.Shortage, precedence.Items[0].DiscrepancyType);
        Assert.Equal(DiscrepancyType.Overage, precedence.Items[1].DiscrepancyType);

        Assert.Throws<OverflowException>(() => DiscrepancyClassifier.CompareMerchandise(
            [
                new DeclaredItem { MaterialCode = "OVF", DeclaredUnits = int.MaxValue, DeclaredCartons = 1 },
                new DeclaredItem { MaterialCode = "OVF", DeclaredUnits = 1, DeclaredCartons = 1 }
            ],
            []));
    }

    [Fact]
    public void CompareMerchandise_ZeroTotalMaterials_AreClassifiedAsZeroTotalAndExcludedFromActiveKpis()
    {
        var declared = new List<DeclaredItem>
        {
            new() { MaterialCode = "5033282", Description = "ESFERAS 25 PCS 6 CM C1", DeclaredUnits = 0, DeclaredCartons = 0 },
            new() { MaterialCode = "5033283", Description = "SET DE ESFERAS 25 PCS 6 CM C2", DeclaredUnits = 0, DeclaredCartons = 0 },
            new() { MaterialCode = "5033292", Description = "ESFERAS 16 PCS 8 CM C1", DeclaredUnits = 600, DeclaredCartons = 50 }
        };
        var physical = new List<PhysicalItem>
        {
            new() { MaterialCode = "5033282", Description = "ESFERAS 25 PCS 6 CM C1", PhysicalUnits = 0, PhysicalCartons = 0 },
            new() { MaterialCode = "5033283", Description = "SET DE ESFERAS 25 PCS 6 CM C2", PhysicalUnits = 95, PhysicalCartons = 5 },
            new() { MaterialCode = "5033292", Description = "ESFERAS 16 PCS 8 CM C1", PhysicalUnits = 600, PhysicalCartons = 50 }
        };

        ComparisonSummary summary = DiscrepancyClassifier.CompareMerchandise(declared, physical);

        // 5033282 is zero total (0 declared, 0 physical)
        DiscrepancyItem itemZero = summary.Items.First(i => i.MaterialCode == "5033282");
        Assert.True(itemZero.IsZeroTotal);
        Assert.Equal("SIN MOVIMIENTO (0 TOTAL)", itemZero.DiscrepancyTypeDescription);
        Assert.Equal("#F1F5F9", itemZero.BadgeBackgroundColor);
        Assert.Equal("#64748B", itemZero.BadgeForegroundColor);

        // 5033283 has 95 physical units -> NOT zero total, is UnexpectedItem
        DiscrepancyItem itemSobrante = summary.Items.First(i => i.MaterialCode == "5033283");
        Assert.False(itemSobrante.IsZeroTotal);
        Assert.Equal(DiscrepancyType.UnexpectedItem, itemSobrante.DiscrepancyType);
        Assert.Equal("SOBRANTE (NO FACTURADO)", itemSobrante.DiscrepancyTypeDescription);

        // 5033292 is an exact match (600 vs 600) -> active clean style
        DiscrepancyItem itemCuadrado = summary.Items.First(i => i.MaterialCode == "5033292");
        Assert.False(itemCuadrado.IsZeroTotal);
        Assert.Equal(DiscrepancyType.ExactMatch, itemCuadrado.DiscrepancyType);

        // KPIs: Total items is 3, but ActiveStylesCount is 2 (5033283 and 5033292)
        Assert.Equal(3, summary.Items.Count);
        Assert.Equal(2, summary.ActiveStylesCount);
        Assert.Equal(1, summary.ZeroTotalMaterialsCount);
        Assert.Equal(1, summary.CleanMaterialsCount);
        Assert.Equal(1, summary.DiscrepantMaterialsCount);
    }
}
