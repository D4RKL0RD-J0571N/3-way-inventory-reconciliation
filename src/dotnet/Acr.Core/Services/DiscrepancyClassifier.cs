using System.Globalization;
using Acr.Core.Models;

namespace Acr.Core.Services;

/// <summary>
/// Pure deterministic discrepancy classifier enforcing the Cardinal Discrepancy Rule
/// (shortages and overages never offset across materials) and evidence-graded revuelto detection.
/// </summary>
public static class DiscrepancyClassifier
{
    private const string MultiPoAggregatedFlag = "MULTIPLE_PO_LINES_AGGREGATED";

    /// <summary>
    /// Compares declared merchandise lines against physical QC counts and produces an auditable summary.
    /// </summary>
    public static ComparisonSummary CompareMerchandise(
        IReadOnlyList<DeclaredItem> declaredItems,
        IReadOnlyList<PhysicalItem> physicalItems)
    {
        ArgumentNullException.ThrowIfNull(declaredItems);
        ArgumentNullException.ThrowIfNull(physicalItems);

        Dictionary<string, List<DeclaredItem>> declaredGroups = GroupDeclaredByMaterial(declaredItems);
        Dictionary<string, List<PhysicalItem>> physicalGroups = GroupPhysicalByMaterial(physicalItems);

        List<string> allCodes = declaredGroups.Keys
            .Union(physicalGroups.Keys, StringComparer.Ordinal)
            .OrderBy(k => k, StringComparer.Ordinal)
            .ToList();

        var items = new List<DiscrepancyItem>(allCodes.Count);
        var multiPoMaterials = new List<string>();

        foreach (string code in allCodes)
        {
            declaredGroups.TryGetValue(code, out List<DeclaredItem>? declList);
            physicalGroups.TryGetValue(code, out List<PhysicalItem>? physList);

            DiscrepancyItem item = BuildMaterialComparison(code, declList ?? [], physList ?? []);
            items.Add(item);
            if (item.HasMultiplePoLines)
                multiPoMaterials.Add(code);
        }

        List<RevueltoMatch> revueltoMatches = DetectRevueltoPairs(items);
        return BuildSummary(items, revueltoMatches, multiPoMaterials);
    }

    private static Dictionary<string, List<DeclaredItem>> GroupDeclaredByMaterial(
        IReadOnlyList<DeclaredItem> declaredItems)
    {
        var map = new Dictionary<string, List<DeclaredItem>>(StringComparer.Ordinal);
        foreach (DeclaredItem item in declaredItems)
        {
            if (!map.TryGetValue(item.MaterialCode, out List<DeclaredItem>? list))
            {
                list = [];
                map[item.MaterialCode] = list;
            }
            list.Add(item);
        }
        return map;
    }

    private static Dictionary<string, List<PhysicalItem>> GroupPhysicalByMaterial(
        IReadOnlyList<PhysicalItem> physicalItems)
    {
        var map = new Dictionary<string, List<PhysicalItem>>(StringComparer.Ordinal);
        foreach (PhysicalItem item in physicalItems)
        {
            if (!map.TryGetValue(item.MaterialCode, out List<PhysicalItem>? list))
            {
                list = [];
                map[item.MaterialCode] = list;
            }
            list.Add(item);
        }
        return map;
    }

    private static DiscrepancyItem BuildMaterialComparison(
        string materialCode,
        List<DeclaredItem> declList,
        List<PhysicalItem> physList)
    {
        checked
        {
            (int declU, int declC) = SumDeclared(declList);
            (int physU, int physC, int damU) = SumPhysical(physList);

            int diffU = physU - declU;
            int diffC = physC - declC;
            (DiscrepancyType type, string explanation) =
                DetermineDiscrepancyType(declU, physU, diffU, diffC, damU);

            List<string> flags = CollectFlags(physList, declList.Count > 1);
            (string? orderPo43, IReadOnlyList<string> poBreakdown) =
                BuildPoMetadata(declList, physList);

            return new DiscrepancyItem
            {
                MaterialCode = materialCode,
                VendorMaterial = CoalesceText(physList, declList, p => p.VendorMaterial, d => d.VendorMaterial),
                Description = CoalesceText(physList, declList, p => p.Description, d => d.Description),
                Category = CoalesceText(physList, declList, p => p.Category, d => d.Category),
                OrderPo43 = orderPo43,
                DeclaredUnits = declU,
                PhysicalUnits = physU,
                DifferenceUnits = diffU,
                DeclaredCartons = declC,
                PhysicalCartons = physC,
                DifferenceCartons = diffC,
                DamagedUnits = damU,
                DiscrepancyType = type,
                HasMultiplePoLines = declList.Count > 1,
                PoLineCount = declList.Count,
                PoBreakdown = poBreakdown,
                Flags = flags,
                Explanation = explanation
            };
        }
    }

    private static (int Units, int Cartons) SumDeclared(List<DeclaredItem> declList)
    {
        checked
        {
            int units = 0;
            int cartons = 0;
            foreach (DeclaredItem item in declList)
            {
                units += item.DeclaredUnits;
                cartons += item.DeclaredCartons;
            }
            return (units, cartons);
        }
    }

    private static (int Units, int Cartons, int Damaged) SumPhysical(List<PhysicalItem> physList)
    {
        checked
        {
            int units = 0;
            int cartons = 0;
            int damaged = 0;
            foreach (PhysicalItem item in physList)
            {
                units += item.PhysicalUnits;
                cartons += item.PhysicalCartons;
                damaged += item.DamagedUnits;
            }
            return (units, cartons, damaged);
        }
    }

    private static List<string> CollectFlags(List<PhysicalItem> physList, bool isMultiPo)
    {
        var flags = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (PhysicalItem item in physList)
        {
            foreach (string flag in item.Flags)
            {
                if (seen.Add(flag))
                    flags.Add(flag);
            }
        }
        if (isMultiPo && seen.Add(MultiPoAggregatedFlag))
            flags.Add(MultiPoAggregatedFlag);
        return flags;
    }

    private static (string? OrderPo43, IReadOnlyList<string> Breakdown) BuildPoMetadata(
        List<DeclaredItem> declList,
        List<PhysicalItem> physList)
    {
        if (declList.Count > 1)
        {
            List<string> distinctPos = declList
                .Select(d => d.OrderPo43)
                .Where(po => !string.IsNullOrWhiteSpace(po))
                .Select(po => po!.Trim())
                .Distinct(StringComparer.Ordinal)
                .ToList();
            string? joined = distinctPos.Count > 0 ? string.Join(", ", distinctPos) : null;
            List<string> breakdown = declList
                .Select(d =>
                {
                    string poLabel = string.IsNullOrWhiteSpace(d.OrderPo43) ? "SIN-PO" : d.OrderPo43.Trim();
                    return $"{poLabel}: {d.DeclaredUnits}u ({d.DeclaredCartons}b)";
                })
                .ToList();
            return (joined, breakdown);
        }

        if (declList.Count == 1)
            return (declList[0].OrderPo43, Array.Empty<string>());

        string? physPo = physList.Select(p => p.OrderPo43).FirstOrDefault(po => !string.IsNullOrWhiteSpace(po));
        return (physPo, Array.Empty<string>());
    }

    private static string CoalesceText(
        List<PhysicalItem> physList,
        List<DeclaredItem> declList,
        Func<PhysicalItem, string?> physSelector,
        Func<DeclaredItem, string?> declSelector)
    {
        foreach (PhysicalItem phys in physList)
        {
            string? val = physSelector(phys);
            if (!string.IsNullOrWhiteSpace(val))
                return val;
        }
        foreach (DeclaredItem decl in declList)
        {
            string? val = declSelector(decl);
            if (!string.IsNullOrWhiteSpace(val))
                return val;
        }
        return "";
    }

    private static (DiscrepancyType Type, string Explanation) DetermineDiscrepancyType(
        int declU,
        int physU,
        int diffU,
        int diffC,
        int damU)
    {
        checked
        {
            if (declU > 0 && physU == 0)
                return (DiscrepancyType.MissingItem, "Material declarado completamente ausente en la recepción física.");
            if (declU == 0 && physU > 0)
                return (DiscrepancyType.UnexpectedItem, $"Material no declarado físicamente recibido (+{diffU} u).");
            if (diffU < 0)
                return (DiscrepancyType.Shortage, $"Faltante físico de {Math.Abs(diffU)} unidades.");
            if (diffU > 0)
                return (DiscrepancyType.Overage, $"Sobrante físico de +{diffU} unidades.");
            if (damU > 0)
                return (DiscrepancyType.DamageOnly, $"Unidades completas pero {damU} unidades físicamente dañadas.");
            if (diffC != 0)
            {
                string signedCartons = diffC > 0
                    ? $"+{diffC.ToString(CultureInfo.InvariantCulture)}"
                    : diffC.ToString(CultureInfo.InvariantCulture);
                return (DiscrepancyType.CartonMismatchOnly, $"Unidades exactas pero diferencia en bultos ({signedCartons} bultos).");
            }
            return (DiscrepancyType.ExactMatch, "Coincidencia exacta entre declaración física y documental.");
        }
    }

    private static List<RevueltoMatch> DetectRevueltoPairs(List<DiscrepancyItem> items)
    {
        var matches = new List<RevueltoMatch>();
        var usedOverageIndices = new HashSet<int>();

        for (int sIdx = 0; sIdx < items.Count; sIdx++)
        {
            DiscrepancyItem shortage = items[sIdx];
            if (shortage.DifferenceUnits >= 0)
                continue;

            int targetQty = checked(Math.Abs(shortage.DifferenceUnits));
            int? bestOverageIdx = FindBestOverageMatch(items, shortage, targetQty, usedOverageIndices);
            if (!bestOverageIdx.HasValue)
                continue;

            int oIdx = bestOverageIdx.Value;
            usedOverageIndices.Add(oIdx);
            RevueltoMatch pair = ApplyRevueltoPairing(items, sIdx, oIdx, targetQty);
            matches.Add(pair);
        }

        return matches;
    }

    private static int? FindBestOverageMatch(
        List<DiscrepancyItem> items,
        DiscrepancyItem shortage,
        int targetQty,
        HashSet<int> usedOverageIndices)
    {
        int? bestIdx = null;
        RevueltoConfidence bestConfidence = RevueltoConfidence.Possible;

        for (int oIdx = 0; oIdx < items.Count; oIdx++)
        {
            if (usedOverageIndices.Contains(oIdx))
                continue;

            DiscrepancyItem overage = items[oIdx];
            if (overage.DifferenceUnits != targetQty)
                continue;

            (RevueltoConfidence confidence, _, _) = EvaluateRevueltoEvidence(shortage, overage, targetQty);
            if (!bestIdx.HasValue || (int)confidence < (int)bestConfidence)
            {
                bestIdx = oIdx;
                bestConfidence = confidence;
            }
        }

        return bestIdx;
    }

    private static RevueltoMatch ApplyRevueltoPairing(
        List<DiscrepancyItem> items,
        int sIdx,
        int oIdx,
        int targetQty)
    {
        DiscrepancyItem s = items[sIdx];
        DiscrepancyItem o = items[oIdx];
        (RevueltoConfidence confidence, string evidence, IReadOnlyList<string> attrs) =
            EvaluateRevueltoEvidence(s, o, targetQty);

        items[sIdx] = s with
        {
            DiscrepancyType = DiscrepancyType.RevueltoCandidate,
            RevueltoPartner = o.MaterialCode,
            RevueltoConfidence = confidence,
            RevueltoEvidence = evidence,
            Explanation = $"{s.Explanation} [Candidato a revuelto con {o.MaterialCode} (+{targetQty} u)]"
        };

        items[oIdx] = o with
        {
            DiscrepancyType = DiscrepancyType.RevueltoCandidate,
            RevueltoPartner = s.MaterialCode,
            RevueltoConfidence = confidence,
            RevueltoEvidence = evidence,
            Explanation = $"{o.Explanation} [Candidato a revuelto con {s.MaterialCode} (-{targetQty} u)]"
        };

        return new RevueltoMatch
        {
            ShortageMaterialCode = s.MaterialCode,
            OverageMaterialCode = o.MaterialCode,
            ShortageUnits = targetQty,
            OverageUnits = targetQty,
            Confidence = confidence,
            Evidence = evidence,
            MatchedAttributes = attrs
        };
    }

    private static (RevueltoConfidence Confidence, string Evidence, IReadOnlyList<string> Attributes)
        EvaluateRevueltoEvidence(DiscrepancyItem shortage, DiscrepancyItem overage, int targetQty)
    {
        var attrs = new List<string> { "ExactUnitDelta" };

        bool sameVendorMat = !string.IsNullOrWhiteSpace(shortage.VendorMaterial) &&
                             !string.IsNullOrWhiteSpace(overage.VendorMaterial) &&
                             string.Equals(shortage.VendorMaterial.Trim(), overage.VendorMaterial.Trim(),
                                 StringComparison.OrdinalIgnoreCase);
        bool sameCategory = !string.IsNullOrWhiteSpace(shortage.Category) &&
                            !string.IsNullOrWhiteSpace(overage.Category) &&
                            string.Equals(shortage.Category.Trim(), overage.Category.Trim(),
                                StringComparison.OrdinalIgnoreCase);
        bool opposingCartons = shortage.DifferenceCartons < 0 &&
                               overage.DifferenceCartons > 0 &&
                               overage.DifferenceCartons == checked(Math.Abs(shortage.DifferenceCartons));

        if (sameVendorMat) attrs.Add("MatchingVendorMaterial");
        if (sameCategory) attrs.Add("SameCategory");
        if (opposingCartons) attrs.Add("OpposingCartonDelta");

        RevueltoConfidence confidence = sameVendorMat || (sameCategory && opposingCartons)
            ? RevueltoConfidence.Confirmed
            : (sameCategory || opposingCartons ? RevueltoConfidence.Probable : RevueltoConfidence.Possible);

        string evidence = $"Delta exacto ({targetQty} u); señales: {string.Join(", ", attrs)}.";
        return (confidence, evidence, attrs);
    }

    private static ComparisonSummary BuildSummary(
        List<DiscrepancyItem> items,
        List<RevueltoMatch> revueltoMatches,
        List<string> multiPoMaterials)
    {
        checked
        {
            int totDeclU = 0, totPhysU = 0, totDeclC = 0, totPhysC = 0;
            int totShort = 0, totOver = 0, totDam = 0, cleanCount = 0;

            foreach (DiscrepancyItem item in items)
            {
                totDeclU += item.DeclaredUnits;
                totPhysU += item.PhysicalUnits;
                totDeclC += item.DeclaredCartons;
                totPhysC += item.PhysicalCartons;
                totDam += item.DamagedUnits;

                if (item.DifferenceUnits < 0)
                    totShort += Math.Abs(item.DifferenceUnits);
                else if (item.DifferenceUnits > 0)
                    totOver += item.DifferenceUnits;

                if (item.DiscrepancyType == DiscrepancyType.ExactMatch)
                    cleanCount++;
            }

            return new ComparisonSummary
            {
                TotalDeclaredUnits = totDeclU,
                TotalPhysicalUnits = totPhysU,
                NetUnitDifference = totPhysU - totDeclU,
                TotalDeclaredCartons = totDeclC,
                TotalPhysicalCartons = totPhysC,
                NetCartonDifference = totPhysC - totDeclC,
                TotalShortageUnits = totShort,
                TotalOverageUnits = totOver,
                TotalDamagedUnits = totDam,
                CleanMaterialsCount = cleanCount,
                DiscrepantMaterialsCount = items.Count - cleanCount,
                RevueltoPairsCount = revueltoMatches.Count,
                Items = items,
                RevueltoMatches = revueltoMatches,
                MultiPoMaterials = multiPoMaterials
            };
        }
    }
}
