using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Acr.Core.Models;

namespace Acr.Core.Services;

/// <summary>
/// Pure domain matcher separating Operational Business Identity (<c>(NormalizedContainerId, Pedido41)</c>)
/// from File Provenance, normalizing <c>'DATOS GENERALES'!B33</c> authorship, collapsing duplicates/drafts,
/// and enforcing idempotent matching against the Daily Log.
/// </summary>
public static class HistoricalIdentityMatcher
{
    /// <summary>Canonical normalized author name for Jostin closures.</summary>
    public const string CanonicalTargetAuthor = "Jostin Lopez Sobalbarro";

    private static readonly HashSet<string> TargetAuthorAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        "JOSTIN LOPEZ SOBALBARRO",
        "JOSTIN LOPEZ SOBALBARO",
        "JOSTIN LOPEZ"
    };

    private static readonly Regex ElaboradoPrefixPattern =
        new(@"^\s*ELABORADO\s+POR\s*:?\s*", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static readonly Regex WhitespaceRunPattern =
        new(@"\s+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex PoSplitPattern =
        new(@"[,;/\r\n\s]+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex SpacedDigitHyphenPattern =
        new(@"(?<=[0-9])\s*-\s*(?=[0-9])", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Normalizes a raw author string from <c>'DATOS GENERALES'!B33</c>, stripping labels and diacritics
    /// and mapping known Jostin variants (including the <c>IMP-326</c> <c>Sobalbaro</c> typo) to canonical form.
    /// </summary>
    public static string NormalizeAuthorName(string? rawAuthorB33)
    {
        if (string.IsNullOrWhiteSpace(rawAuthorB33))
            return "";

        string withoutPrefix = ElaboradoPrefixPattern.Replace(rawAuthorB33.Replace('\u00A0', ' '), "");
        string collapsed = WhitespaceRunPattern.Replace(withoutPrefix.Trim(), " ");
        if (collapsed.Length == 0)
            return "";

        string folded = FoldDiacritics(collapsed);
        return TargetAuthorAliases.Contains(folded) ? CanonicalTargetAuthor : folded;
    }

    /// <summary>
    /// Determines whether the <c>'DATOS GENERALES'!B33</c> cell identifies Jostin as the closure author.
    /// </summary>
    public static bool IsTargetAuthor(string? rawAuthorB33) =>
        string.Equals(NormalizeAuthorName(rawAuthorB33), CanonicalTargetAuthor, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Parses multi-PO strings (e.g. <c>"4300007309-8144"</c>, <c>"4300006764-6804-6805"</c>, slash/comma lists)
    /// into distinct expanded Purchase Order 43 numbers.
    /// </summary>
    public static IReadOnlyList<string> ParsePoList(string? rawPoText)
    {
        if (string.IsNullOrWhiteSpace(rawPoText))
            return Array.Empty<string>();

        string normalized = SpacedDigitHyphenPattern.Replace(rawPoText.Trim(), "-");
        var result = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string[] tokens = PoSplitPattern.Split(normalized);

        foreach (string token in tokens)
        {
            if (string.IsNullOrWhiteSpace(token) || token == "-" || token.StartsWith('#'))
                continue;

            ExpandHyphenatedPoToken(token, result, seen);
        }

        return result;
    }

    /// <summary>
    /// Evaluates historical workbook anomaly flags that trigger human review.
    /// </summary>
    public static IReadOnlyList<HumanReviewReason> EvaluateAnomalies(HistoricalSyncRecord record)
    {
        ArgumentNullException.ThrowIfNull(record);

        var reasons = new List<HumanReviewReason>();
        EvaluateIdentityAndSummaryAnomalies(record, reasons);
        EvaluateLineAndPoAnomalies(record, reasons);
        return reasons;
    }

    /// <summary>
    /// Collapses duplicate or draft workbooks sharing the same Business Identity into a single canonical
    /// candidate while preserving all <see cref="FileProvenance"/> records.
    /// </summary>
    public static IReadOnlyList<CollapsedHistoricalCandidate> CollapseDuplicates(
        IReadOnlyList<HistoricalSyncRecord> records)
    {
        ArgumentNullException.ThrowIfNull(records);

        var candidates = new List<CollapsedHistoricalCandidate>();
        var byContainer = new Dictionary<string, List<HistoricalSyncRecord>>(StringComparer.Ordinal);

        foreach (HistoricalSyncRecord record in records)
        {
            string container = DailyLogLifecycleValidator.NormalizeContainerId(record.Identity.NormalizedContainerId);
            if (string.IsNullOrWhiteSpace(container))
            {
                candidates.Add(BuildSingleCollapsedCandidate(record));
                continue;
            }

            if (!byContainer.TryGetValue(container, out List<HistoricalSyncRecord>? list))
            {
                list = [];
                byContainer[container] = list;
            }
            list.Add(record);
        }

        foreach (string container in byContainer.Keys.OrderBy(k => k, StringComparer.Ordinal))
        {
            foreach (List<HistoricalSyncRecord> group in PartitionContainerGroup(byContainer[container]))
                candidates.Add(CollapseGroup(group));
        }

        return candidates;
    }

    /// <summary>
    /// Matches collapsed historical candidates against existing Daily Log entries strictly by Business Identity
    /// (<c>NormalizedContainerId</c> + <c>Pedido41</c>), never by file provenance or <c>PO_43</c>.
    /// </summary>
    public static IReadOnlyList<HistoricalMatchResult> MatchAgainstDailyLog(
        IReadOnlyList<CollapsedHistoricalCandidate> candidates,
        IReadOnlyList<DailyLogEntry> existingLog)
    {
        ArgumentNullException.ThrowIfNull(candidates);
        ArgumentNullException.ThrowIfNull(existingLog);

        Dictionary<string, List<DailyLogEntry>> logByContainer = IndexDailyLogByContainer(existingLog);
        var results = new List<HistoricalMatchResult>(candidates.Count);

        foreach (CollapsedHistoricalCandidate candidate in candidates)
            results.Add(MatchSingleCandidate(candidate, logByContainer));

        return results;
    }

    private static void ExpandHyphenatedPoToken(string token, List<string> result, HashSet<string> seen)
    {
        string[] parts = token.Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
            return;

        string basePo = parts[0];
        if (!IsAllAsciiDigits(basePo))
            return;

        AddUniquePo(basePo, result, seen);
        for (int i = 1; i < parts.Length; i++)
        {
            string suffix = parts[i];
            if (!IsAllAsciiDigits(suffix))
                continue;

            string expanded = suffix.Length < basePo.Length
                ? string.Concat(basePo.AsSpan(0, basePo.Length - suffix.Length), suffix)
                : suffix;
            AddUniquePo(expanded, result, seen);
        }
    }

    private static void AddUniquePo(string po, List<string> result, HashSet<string> seen)
    {
        if (seen.Add(po))
            result.Add(po);
    }

    private static void EvaluateIdentityAndSummaryAnomalies(
        HistoricalSyncRecord record,
        List<HumanReviewReason> reasons)
    {
        if (string.IsNullOrWhiteSpace(record.Identity.NormalizedContainerId))
            reasons.Add(HumanReviewReason.MissingContainerId);

        if (string.IsNullOrWhiteSpace(record.Identity.Pedido41))
            reasons.Add(HumanReviewReason.MissingPedido41);

        if (record.TheoreticalCartons == 0 &&
            record.ReceivedCartons == 0 &&
            (record.ActualCierreRecepItemsCount > 0 || (record.DeclaredLinesCountDg ?? 0) > 0))
        {
            reasons.Add(HumanReviewReason.ZeroCartonsWithPositiveLines);
        }

        if (record.DetectedErrorTokens.Count > 0 ||
            !record.TheoreticalCartons.HasValue ||
            !record.ReceivedCartons.HasValue)
        {
            reasons.Add(HumanReviewReason.FormulaOrLiteralErrorInSummary);
        }

        if (record.TheoreticalCartons.HasValue &&
            record.ReceivedCartons.HasValue &&
            record.TheoreticalCartons.Value != record.ReceivedCartons.Value)
        {
            reasons.Add(HumanReviewReason.NegativeOrExtremeCartonDiscrepancy);
        }
    }

    private static void EvaluateLineAndPoAnomalies(
        HistoricalSyncRecord record,
        List<HumanReviewReason> reasons)
    {
        if (record.DeclaredLinesCountDg == 0 && record.ActualCierreRecepItemsCount > 0)
            reasons.Add(HumanReviewReason.LineCountMismatchDgVsCierreRecep);

        if (record.CierreRecepDistinctPoCount == 0)
            reasons.Add(HumanReviewReason.EmptyCierreRecepPoColumn);

        string path = record.Provenance.SourcePath;
        if ((path.Contains("PEND CONTROL", StringComparison.OrdinalIgnoreCase) ||
             path.Contains("SIN CONTROL", StringComparison.OrdinalIgnoreCase)) &&
            record.TheoreticalCartons is > 0 &&
            record.ReceivedCartons is > 0)
        {
            reasons.Add(HumanReviewReason.FolderStatusVsCartonMismatch);
        }

        if (record.Identity.Po43Numbers.Count > 0 &&
            record.CierreRecepDistinctPoCount > 0 &&
            record.Identity.Po43Numbers.Count != record.CierreRecepDistinctPoCount)
        {
            reasons.Add(HumanReviewReason.PoCountMismatchDgVsCierreRecep);
        }
    }

    private static CollapsedHistoricalCandidate BuildSingleCollapsedCandidate(HistoricalSyncRecord record)
    {
        IReadOnlyList<HumanReviewReason> reasons = MergeReasons(record.ReviewReasons, EvaluateAnomalies(record));
        return new CollapsedHistoricalCandidate
        {
            PrimaryRecord = record,
            AllProvenances = [record.Provenance],
            CombinedReviewReasons = reasons
        };
    }

    private static List<List<HistoricalSyncRecord>> PartitionContainerGroup(List<HistoricalSyncRecord> containerRecords)
    {
        List<string> distinctPedidos = containerRecords
            .Select(r => CleanPedido(r.Identity.Pedido41))
            .Where(p => p.Length > 0)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(p => p, StringComparer.Ordinal)
            .ToList();

        if (distinctPedidos.Count <= 1)
            return [containerRecords];

        var groups = new List<List<HistoricalSyncRecord>>();
        foreach (string pedido in distinctPedidos)
        {
            groups.Add(containerRecords.Where(r => CleanPedido(r.Identity.Pedido41) == pedido).ToList());
        }

        List<HistoricalSyncRecord> missingPedidoRecords =
            containerRecords.Where(r => CleanPedido(r.Identity.Pedido41).Length == 0).ToList();
        if (missingPedidoRecords.Count > 0)
            groups[0].AddRange(missingPedidoRecords);

        return groups;
    }

    private static CollapsedHistoricalCandidate CollapseGroup(List<HistoricalSyncRecord> group)
    {
        List<HistoricalSyncRecord> ranked = group
            .OrderByDescending(r => !r.IsDraftOrCopyFile)
            .ThenByDescending(r => r.CierreRecepDistinctPoCount > 0)
            .ThenByDescending(r => r.DetectedErrorTokens.Count == 0)
            .ThenByDescending(r => r.TheoreticalCartons is > 0 && r.TheoreticalCartons == r.ReceivedCartons)
            .ThenByDescending(r => r.Identity.RevisionDate ?? DateOnly.MinValue)
            .ThenBy(r => r.Provenance.SourcePath.Length)
            .ThenBy(r => r.Provenance.SourcePath, StringComparer.Ordinal)
            .ToList();

        HistoricalSyncRecord primary = ranked[0];
        var combinedReasons = new List<HumanReviewReason>(MergeReasons(primary.ReviewReasons, EvaluateAnomalies(primary)));

        if (group.Count > 1 && !combinedReasons.Contains(HumanReviewReason.MultipleFileVersionsForIdentity))
            combinedReasons.Add(HumanReviewReason.MultipleFileVersionsForIdentity);

        if (group.Any(r => string.IsNullOrWhiteSpace(CleanPedido(r.Identity.Pedido41))) &&
            !combinedReasons.Contains(HumanReviewReason.MissingPedido41))
        {
            combinedReasons.Add(HumanReviewReason.MissingPedido41);
        }

        return new CollapsedHistoricalCandidate
        {
            PrimaryRecord = primary,
            AllProvenances = ranked.Select(r => r.Provenance).ToList(),
            CombinedReviewReasons = combinedReasons
        };
    }

    private static Dictionary<string, List<DailyLogEntry>> IndexDailyLogByContainer(
        IReadOnlyList<DailyLogEntry> existingLog)
    {
        var index = new Dictionary<string, List<DailyLogEntry>>(StringComparer.Ordinal);
        foreach (DailyLogEntry entry in existingLog)
        {
            string normalized = !string.IsNullOrWhiteSpace(entry.NormalizedContainerId)
                ? DailyLogLifecycleValidator.NormalizeContainerId(entry.NormalizedContainerId)
                : DailyLogLifecycleValidator.NormalizeContainerId(entry.ContainerId);
            if (string.IsNullOrWhiteSpace(normalized))
                continue;

            if (!index.TryGetValue(normalized, out List<DailyLogEntry>? list))
            {
                list = [];
                index[normalized] = list;
            }
            list.Add(entry);
        }
        return index;
    }

    private static HistoricalMatchResult MatchSingleCandidate(
        CollapsedHistoricalCandidate candidate,
        Dictionary<string, List<DailyLogEntry>> logByContainer)
    {
        var reasons = new List<HumanReviewReason>(candidate.CombinedReviewReasons);
        string containerId = DailyLogLifecycleValidator.NormalizeContainerId(
            candidate.PrimaryRecord.Identity.NormalizedContainerId);

        if (string.IsNullOrWhiteSpace(containerId))
        {
            AddReasonIfMissing(reasons, HumanReviewReason.MissingContainerId);
            return new HistoricalMatchResult
            {
                Candidate = candidate,
                MatchKind = HistoricalMatchKind.ConflictOrAmbiguousRequiresReview,
                MatchReviewReasons = reasons
            };
        }

        if (!logByContainer.TryGetValue(containerId, out List<DailyLogEntry>? matches) || matches.Count == 0)
        {
            return new HistoricalMatchResult
            {
                Candidate = candidate,
                MatchKind = HistoricalMatchKind.NewRecord,
                MatchReviewReasons = reasons
            };
        }

        if (matches.Count > 1)
        {
            AddReasonIfMissing(reasons, HumanReviewReason.AmbiguousDailyLogMatch);
            return new HistoricalMatchResult
            {
                Candidate = candidate,
                MatchKind = HistoricalMatchKind.ConflictOrAmbiguousRequiresReview,
                MatchedDailyLogEntries = matches,
                MatchReviewReasons = reasons
            };
        }

        return EvaluateSingleLogMatch(candidate, matches[0], reasons);
    }

    private static HistoricalMatchResult EvaluateSingleLogMatch(
        CollapsedHistoricalCandidate candidate,
        DailyLogEntry existing,
        List<HumanReviewReason> reasons)
    {
        string candPedido = CleanPedido(candidate.PrimaryRecord.Identity.Pedido41);
        string logPedido = CleanPedido(existing.Pedido41);

        if (candPedido.Length > 0 && logPedido.Length > 0)
        {
            if (string.Equals(candPedido, logPedido, StringComparison.Ordinal))
            {
                return new HistoricalMatchResult
                {
                    Candidate = candidate,
                    MatchKind = HistoricalMatchKind.AlreadySyncedExact,
                    MatchedDailyLogEntries = [existing],
                    MatchReviewReasons = reasons
                };
            }

            AddReasonIfMissing(reasons, HumanReviewReason.AmbiguousDailyLogMatch);
            return new HistoricalMatchResult
            {
                Candidate = candidate,
                MatchKind = HistoricalMatchKind.ConflictOrAmbiguousRequiresReview,
                MatchedDailyLogEntries = [existing],
                MatchReviewReasons = reasons
            };
        }

        AddReasonIfMissing(reasons, HumanReviewReason.MissingPedido41);
        return new HistoricalMatchResult
        {
            Candidate = candidate,
            MatchKind = HistoricalMatchKind.MatchedByContainerFallbackRequiresReview,
            MatchedDailyLogEntries = [existing],
            MatchReviewReasons = reasons
        };
    }

    private static List<HumanReviewReason> MergeReasons(
        IReadOnlyList<HumanReviewReason> existing,
        IReadOnlyList<HumanReviewReason> evaluated)
    {
        var merged = new List<HumanReviewReason>(existing.Count + evaluated.Count);
        foreach (HumanReviewReason reason in existing)
            AddReasonIfMissing(merged, reason);
        foreach (HumanReviewReason reason in evaluated)
            AddReasonIfMissing(merged, reason);
        return merged;
    }

    private static void AddReasonIfMissing(List<HumanReviewReason> reasons, HumanReviewReason reason)
    {
        if (!reasons.Contains(reason))
            reasons.Add(reason);
    }

    private static string CleanPedido(string? pedido)
    {
        if (string.IsNullOrWhiteSpace(pedido))
            return "";
        string trimmed = pedido.Trim();
        return trimmed == "-" || trimmed.StartsWith('#') ? "" : trimmed;
    }

    private static bool IsAllAsciiDigits(string value)
    {
        if (value.Length == 0)
            return false;
        foreach (char ch in value)
        {
            if (!char.IsAsciiDigit(ch))
                return false;
        }
        return true;
    }

    private static string FoldDiacritics(string input)
    {
        string decomposed = input.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(decomposed.Length);
        foreach (char ch in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                sb.Append(ch);
        }
        return sb.ToString().Normalize(NormalizationForm.FormC);
    }
}
