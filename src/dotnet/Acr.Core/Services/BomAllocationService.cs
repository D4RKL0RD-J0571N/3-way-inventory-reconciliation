using Acr.Core.Models;

namespace Acr.Core.Services;

/// <summary>
/// Deterministic largest-remainder proportional distributor for BOM child SKU receipts.
/// Guarantees that the sum of child received units strictly matches the parent received units.
/// </summary>
public static class BomAllocationService
{
    private sealed class AllocationItem
    {
        public required BomDefinitionEntry Entry { get; init; }
        public int Floor { get; set; }
        public double Frac { get; set; }
        public int OriginalIndex { get; init; }
    }

    /// <summary>
    /// Distributes parent received units across child BOM entries using the largest-remainder method (Hamilton-Hare),
    /// guaranteeing:
    /// 1. <c>sum(child.ReceivedUnits) == parentReceivedUnits</c>
    /// 2. Remainder units are allocated to child SKUs with the largest fractional remainder.
    /// 3. In case of fractional tie, original BOM order is preserved deterministically.
    /// </summary>
    public static IReadOnlyList<ChildSkuItem> AllocateChildSkus(
        IReadOnlyList<BomDefinitionEntry> bomEntries,
        int parentInvoicedUnits,
        int parentReceivedUnits,
        int ped41Prk,
        int curve,
        string parentDesc,
        decimal unitCost,
        IReadOnlyDictionary<string, string>? childDescriptions = null)
    {
        if (bomEntries == null || bomEntries.Count == 0)
            return Array.Empty<ChildSkuItem>();

        int totalRatio = 0;
        foreach (var b in bomEntries)
        {
            if (b.Ratio > 0)
                totalRatio += b.Ratio;
        }

        if (totalRatio <= 0)
            totalRatio = curve > 0 ? curve : 1;

        var allocList = new List<AllocationItem>(bomEntries.Count);
        int floorSum = 0;

        for (int i = 0; i < bomEntries.Count; i++)
        {
            var b = bomEntries[i];
            double exact = (double)parentReceivedUnits * Math.Max(0, b.Ratio) / totalRatio;
            int floorVal = (int)Math.Floor(exact);
            floorSum += floorVal;

            allocList.Add(new AllocationItem
            {
                Entry = b,
                Floor = floorVal,
                Frac = exact - floorVal,
                OriginalIndex = i
            });
        }

        int remainder = Math.Max(0, parentReceivedUnits - floorSum);
        if (remainder > 0)
        {
            // Sort by descending fraction, then by original index
            var sorted = allocList.OrderByDescending(a => a.Frac)
                                  .ThenBy(a => a.OriginalIndex)
                                  .ToList();

            for (int i = 0; i < sorted.Count && remainder > 0; i++)
            {
                sorted[i].Floor++;
                remainder--;
            }
        }

        var result = new List<ChildSkuItem>(allocList.Count);
        foreach (var item in allocList)
        {
            int cRec = item.Floor;
            int cTheo = ped41Prk > 0
                ? ped41Prk * item.Entry.Ratio
                : (totalRatio > 0 ? (int)Math.Round((double)parentInvoicedUnits / totalRatio * item.Entry.Ratio, MidpointRounding.AwayFromZero) : 0);

            string desc = "";
            if (childDescriptions != null && childDescriptions.TryGetValue(item.Entry.ChildMaterialCode, out string? customDesc) && !string.IsNullOrWhiteSpace(customDesc))
            {
                desc = customDesc;
            }
            else
            {
                desc = string.IsNullOrWhiteSpace(parentDesc)
                    ? $"Talla ratio {item.Entry.Ratio}"
                    : $"{parentDesc} (Talla ratio {item.Entry.Ratio})";
            }

            result.Add(new ChildSkuItem
            {
                MaterialCode = item.Entry.ChildMaterialCode,
                Description = desc,
                Ratio = item.Entry.Ratio,
                TheoreticalUnits = cTheo,
                ReceivedUnits = cRec,
                UnitCost = unitCost
            });
        }

        return result;
    }
}
