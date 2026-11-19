using Acr.Core.Models;
using Acr.Core.Services;
using Xunit;

namespace Acr.Tests.Core;

public sealed class BomAllocationTests
{
    [Fact]
    public void AllocateChildSkus_LargestRemainder_DistributesProportionallyAndPreservesSum()
    {
        // Prepack curve of 12 (ratio: 2, 4, 4, 2)
        var bom = new List<BomDefinitionEntry>
        {
            new() { ChildMaterialCode = "SKU-S", Ratio = 2 },
            new() { ChildMaterialCode = "SKU-M", Ratio = 4 },
            new() { ChildMaterialCode = "SKU-L", Ratio = 4 },
            new() { ChildMaterialCode = "SKU-XL", Ratio = 2 }
        };

        // Case 1: Exact multiple (120 units)
        var res1 = BomAllocationService.AllocateChildSkus(bom, 120, 120, 10, 12, "Prepack Polo", 5.50m);
        Assert.Equal(4, res1.Count);
        Assert.Equal(120, res1.Sum(c => c.ReceivedUnits));
        Assert.Equal(20, res1[0].ReceivedUnits);
        Assert.Equal(40, res1[1].ReceivedUnits);
        Assert.Equal(40, res1[2].ReceivedUnits);
        Assert.Equal(20, res1[3].ReceivedUnits);

        // Case 2: Uneven partial receipt with remainder (e.g. 50 units received)
        // Ratio total = 12
        // Exact: 50 * 2/12 = 8.333... (floor 8, frac 0.333)
        // Exact: 50 * 4/12 = 16.666... (floor 16, frac 0.666) -> gets +1
        // Exact: 50 * 4/12 = 16.666... (floor 16, frac 0.666) -> gets +1
        // Exact: 50 * 2/12 = 8.333... (floor 8, frac 0.333)
        // Floor sum = 8 + 16 + 16 + 8 = 48. Remainder = 2.
        // Remainder 2 goes to SKU-M and SKU-L (highest fractions 0.666)
        var res2 = BomAllocationService.AllocateChildSkus(bom, 120, 50, 10, 12, "Prepack Polo", 5.50m);
        Assert.Equal(4, res2.Count);
        Assert.Equal(50, res2.Sum(c => c.ReceivedUnits)); // STRICT EQUALITY
        Assert.Equal(8, res2[0].ReceivedUnits);
        Assert.Equal(17, res2[1].ReceivedUnits);
        Assert.Equal(17, res2[2].ReceivedUnits);
        Assert.Equal(8, res2[3].ReceivedUnits);

        // Case 3: 0 received units
        var res3 = BomAllocationService.AllocateChildSkus(bom, 120, 0, 10, 12, "Prepack Polo", 5.50m);
        Assert.Equal(0, res3.Sum(c => c.ReceivedUnits));
        Assert.All(res3, c => Assert.Equal(0, c.ReceivedUnits));
    }

    [Fact]
    public void AllocateChildSkus_EmptyBom_ReturnsEmptyList()
    {
        var res = BomAllocationService.AllocateChildSkus(Array.Empty<BomDefinitionEntry>(), 100, 100, 10, 10, "Parent", 1.0m);
        Assert.Empty(res);
    }
}
