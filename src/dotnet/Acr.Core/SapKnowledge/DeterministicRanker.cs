using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Acr.Core.Models;

namespace Acr.Core.SapKnowledge;

/// <summary>
/// Deterministic offline search request over indexed SAP knowledge entries, transactions, and guides.
/// </summary>
public sealed record SapSearchRequest
{
    /// <summary>Free-text query (transaction code, keyword, symptom, table, or synonym).</summary>
    public string? QueryText { get; init; }

    /// <summary>Optional structured diagnostic context supplying verified identifiers and discrepancy type.</summary>
    public DiagnosticContext? Context { get; init; }

    /// <summary>Optional knowledge category filter.</summary>
    public SapKnowledgeCategory? CategoryFilter { get; init; }

    /// <summary>Optional SAP object kind filter.</summary>
    public SapObjectKind? ObjectKindFilter { get; init; }

    /// <summary>Optional minimum epistemic authority tier (inclusive).</summary>
    public EpistemicTier? MinEpistemicTier { get; init; }

    /// <summary>Whether <see cref="FragmentLifecycleState.Deprecated"/> documents should be included (default false).</summary>
    public bool IncludeDeprecated { get; init; }
}

/// <summary>
/// Ranked search result paired with deterministic score breakdown and caution metadata.
/// </summary>
/// <typeparam name="T">Document type implementing <see cref="ISapRankableDocument"/>.</typeparam>
public sealed record RankedSapResult<T>(
    T Document,
    int TotalScore,
    int RelevanceScore,
    int EpistemicWeight,
    IReadOnlyList<string> MatchedReasons,
    bool RequiresCautionBanner) where T : ISapRankableDocument;

/// <summary>
/// Pure deterministic local search and ranking engine (zero LLM, zero Qdrant, zero vector DB, zero System.IO).
/// </summary>
public static class DeterministicRanker
{
    private static readonly Regex TokenSplitPattern =
        new(@"[^a-z0-9_/\-]+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly IReadOnlyDictionary<string, IReadOnlyList<string>> DomainSynonyms =
        new Dictionary<string, IReadOnlyList<string>>(StringComparer.OrdinalIgnoreCase)
        {
            ["faltante"] = ["shortage", "falta", "missing"],
            ["shortage"] = ["faltante", "falta", "missing"],
            ["sobrante"] = ["overage", "exceso", "excess"],
            ["overage"] = ["sobrante", "exceso", "excess"],
            ["revuelto"] = ["swap", "cruce", "mezclado"],
            ["swap"] = ["revuelto", "cruce"],
            ["cruce"] = ["revuelto", "swap"],
            ["bulto"] = ["carton", "caja", "bultos"],
            ["carton"] = ["bulto", "caja", "bultos"],
            ["prepack"] = ["prk", "curva", "curve"],
            ["prk"] = ["prepack", "curva"],
            ["curva"] = ["prepack", "prk", "curve"],
            ["ncr"] = ["nota de credito", "40", "credito"],
            ["nacionalizacion"] = ["pedido 41", "41", "pedido41"],
            ["tabla"] = ["table", "se16n"],
            ["table"] = ["tabla", "se16n"],
            ["descarga"] = ["zdescarga", "muelle"],
            ["zdescarga"] = ["descarga", "muelle"]
        };

    /// <summary>
    /// Ranks <paramref name="documents"/> deterministically against <paramref name="request"/>.
    /// </summary>
    public static IReadOnlyList<RankedSapResult<T>> Rank<T>(
        IReadOnlyList<T> documents,
        SapSearchRequest request) where T : ISapRankableDocument
    {
        ArgumentNullException.ThrowIfNull(documents);
        ArgumentNullException.ThrowIfNull(request);

        List<string> querySequence = TokenizeQuerySequence(request);
        List<string> queryTokens = querySequence.Distinct(StringComparer.Ordinal).ToList();
        IReadOnlyList<string> contextObjects = request.Context?.GetInferredBusinessObjects() ?? Array.Empty<string>();
        bool hasSignals = queryTokens.Count > 0 ||
                          contextObjects.Count > 0 ||
                          request.Context?.DiscrepancyType is not null;

        var results = new List<RankedSapResult<T>>();
        foreach (T doc in documents)
        {
            if (!PassesFilterGate(doc, request))
                continue;

            RankedSapResult<T>? ranked = ScoreDocument(doc, request, queryTokens, querySequence, contextObjects, hasSignals);
            if (ranked is not null)
                results.Add(ranked);
        }

        return results
            .OrderByDescending(r => r.TotalScore)
            .ThenBy(r => (int)r.Document.EpistemicTier)
            .ThenByDescending(r => (int)r.Document.LifecycleState)
            .ThenBy(r => r.Document.Id, StringComparer.Ordinal)
            .ToList();
    }

    private static bool PassesFilterGate(ISapRankableDocument doc, SapSearchRequest request)
    {
        if (!request.IncludeDeprecated && doc.LifecycleState == FragmentLifecycleState.Deprecated)
            return false;
        if (request.CategoryFilter.HasValue && doc.Category != request.CategoryFilter.Value)
            return false;
        if (request.ObjectKindFilter.HasValue && doc.ObjectKind != request.ObjectKindFilter.Value)
            return false;
        if (request.MinEpistemicTier.HasValue && doc.EpistemicTier > request.MinEpistemicTier.Value)
            return false;
        return true;
    }

    private static RankedSapResult<T>? ScoreDocument<T>(
        T doc,
        SapSearchRequest request,
        List<string> queryTokens,
        List<string> querySequence,
        IReadOnlyList<string> contextObjects,
        bool hasSignals) where T : ISapRankableDocument
    {
        var reasons = new List<string>();
        int relevance = ComputeRelevance(doc, request, queryTokens, querySequence, contextObjects, reasons);
        if (hasSignals && relevance == 0)
            return null;

        int epistemicWeight = EpistemicTierRules.GetRankingWeight(doc.EpistemicTier);
        int lifecycleWeight = EpistemicTierRules.GetLifecycleWeight(doc.LifecycleState);
        int total = checked(relevance + epistemicWeight + lifecycleWeight);
        bool caution = EpistemicTierRules.RequiresCautionBanner(doc.EpistemicTier, doc.LifecycleState);

        return new RankedSapResult<T>(doc, total, relevance, epistemicWeight, reasons, caution);
    }

    private static int ComputeRelevance(
        ISapRankableDocument doc,
        SapSearchRequest request,
        List<string> queryTokens,
        List<string> querySequence,
        IReadOnlyList<string> contextObjects,
        List<string> reasons)
    {
        List<string> titleSummarySeq = TokenizeSequence($"{doc.Title} {doc.Summary}");
        var titleSummarySet = new HashSet<string>(titleSummarySeq, StringComparer.Ordinal);
        checked
        {
            int score = ScoreCodesAndDiscrepancy(doc, request, queryTokens, querySequence, reasons);
            score += ScoreBusinessObjects(doc, queryTokens, contextObjects, reasons);
            score += ScoreKeywordsTagsSynonyms(doc, queryTokens, titleSummarySet, titleSummarySeq, reasons);
            score += ScoreTextMatches(doc, queryTokens, titleSummarySet, reasons);
            return score;
        }
    }

    private static int ScoreCodesAndDiscrepancy(
        ISapRankableDocument doc,
        SapSearchRequest request,
        List<string> queryTokens,
        List<string> querySequence,
        List<string> reasons)
    {
        int score = 0;
        var queryTokenSet = new HashSet<string>(queryTokens, StringComparer.Ordinal);
        var matchedCodes = new HashSet<string>(StringComparer.Ordinal);

        foreach (string rawCode in doc.TransactionCodes)
        {
            string normCode = FoldAndLower(rawCode);
            if (normCode.Length == 0 || matchedCodes.Contains(normCode))
                continue;

            List<string> codeSeq = TokenizeSequence(normCode);
            bool hit = codeSeq.Count <= 1
                ? queryTokenSet.Contains(normCode) || (codeSeq.Count == 1 && queryTokenSet.Contains(codeSeq[0]))
                : queryTokenSet.Contains(normCode) || ContainsTokenSequence(querySequence, codeSeq);

            if (hit && matchedCodes.Add(normCode))
            {
                score += 100;
                reasons.Add($"TCode:{normCode}");
            }
        }

        if (request.Context?.DiscrepancyType is DiscrepancyType dt && doc.DiscrepancyTypes.Contains(dt))
        {
            score += 60;
            reasons.Add($"Discrepancy:{dt}");
        }

        return score;
    }

    private static int ScoreBusinessObjects(
        ISapRankableDocument doc,
        List<string> queryTokens,
        IReadOnlyList<string> contextObjects,
        List<string> reasons)
    {
        int score = 0;
        var docObjects = new HashSet<string>(doc.BusinessObjects.Select(FoldAndLower), StringComparer.Ordinal);
        var matched = new HashSet<string>(StringComparer.Ordinal);

        foreach (string ctxObj in contextObjects)
        {
            string norm = FoldAndLower(ctxObj);
            if (docObjects.Contains(norm) && matched.Add(norm))
            {
                score += 40;
                reasons.Add($"BusinessObject:{ctxObj}");
            }
        }

        foreach (string token in queryTokens)
        {
            if (docObjects.Contains(token) && matched.Add(token))
            {
                score += 40;
                reasons.Add($"BusinessObject:{token}");
            }
        }

        return score;
    }

    private static int ScoreKeywordsTagsSynonyms(
        ISapRankableDocument doc,
        List<string> queryTokens,
        HashSet<string> titleSummarySet,
        List<string> titleSummarySeq,
        List<string> reasons)
    {
        int score = 0;
        var keywords = new HashSet<string>(doc.Keywords.Select(FoldAndLower), StringComparer.Ordinal);
        var tags = new HashSet<string>(doc.Tags.Select(FoldAndLower), StringComparer.Ordinal);
        var synonyms = new HashSet<string>(doc.Synonyms.Select(FoldAndLower), StringComparer.Ordinal);
        var matchedSynonyms = new HashSet<string>(StringComparer.Ordinal);

        foreach (string token in queryTokens)
        {
            if (keywords.Contains(token))
            {
                score += 25;
                reasons.Add($"Keyword:{token}");
            }
            if (tags.Contains(token))
            {
                score += 20;
                reasons.Add($"Tag:{token}");
            }
            if (synonyms.Contains(token) && matchedSynonyms.Add(token))
            {
                score += 15;
                reasons.Add($"Synonym:{token}");
            }
            score += ScoreExpandedSynonyms(
                token, titleSummarySet, titleSummarySeq, keywords, tags, synonyms, matchedSynonyms, reasons);
        }

        return score;
    }

    private static int ScoreExpandedSynonyms(
        string token,
        HashSet<string> titleSummarySet,
        List<string> titleSummarySeq,
        HashSet<string> keywords,
        HashSet<string> tags,
        HashSet<string> synonyms,
        HashSet<string> matchedSynonyms,
        List<string> reasons)
    {
        if (!DomainSynonyms.TryGetValue(token, out IReadOnlyList<string>? expansions))
            return 0;

        int score = 0;
        foreach (string alt in expansions)
        {
            string normAlt = FoldAndLower(alt);
            List<string> altTokens = TokenizeSequence(normAlt);
            bool textHit = altTokens.Count switch
            {
                1 => titleSummarySet.Contains(altTokens[0]),
                > 1 => ContainsTokenSequence(titleSummarySeq, altTokens),
                _ => false
            };

            bool hit = keywords.Contains(normAlt) ||
                       tags.Contains(normAlt) ||
                       synonyms.Contains(normAlt) ||
                       textHit;

            if (hit && matchedSynonyms.Add(normAlt))
            {
                score += 15;
                reasons.Add($"SynonymExpansion:{token}->{normAlt}");
            }
        }
        return score;
    }

    private static int ScoreTextMatches(
        ISapRankableDocument doc,
        List<string> queryTokens,
        HashSet<string> titleSummarySet,
        List<string> reasons)
    {
        int score = 0;
        var bodyTokens = new HashSet<string>(TokenizeSequence(doc.BodyText), StringComparer.Ordinal);

        foreach (string token in queryTokens)
        {
            if (titleSummarySet.Contains(token))
            {
                score += 12;
                reasons.Add($"TitleSummary:{token}");
            }
            if (bodyTokens.Contains(token))
            {
                score += 5;
                reasons.Add($"Body:{token}");
            }
        }

        return score;
    }

    private static List<string> TokenizeQuerySequence(SapSearchRequest request)
    {
        string combined = string.Join(" ",
            new[] { request.QueryText, request.Context?.ProblemDescription }
                .Where(s => !string.IsNullOrWhiteSpace(s)));
        return TokenizeSequence(combined);
    }

    private static List<string> TokenizeSequence(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return [];

        string folded = FoldAndLower(text);
        return TokenSplitPattern
            .Split(folded)
            .Select(t => t.Trim('-', '/'))
            .Where(t => t.Length > 0)
            .ToList();
    }

    private static bool ContainsTokenSequence(IReadOnlyList<string> haystack, IReadOnlyList<string> needle)
    {
        if (needle.Count == 0 || needle.Count > haystack.Count)
            return false;

        for (int i = 0; i <= haystack.Count - needle.Count; i++)
        {
            bool matched = true;
            for (int j = 0; j < needle.Count; j++)
            {
                if (!string.Equals(haystack[i + j], needle[j], StringComparison.Ordinal))
                {
                    matched = false;
                    break;
                }
            }
            if (matched)
                return true;
        }
        return false;
    }

    private static string FoldAndLower(string input)
    {
        string decomposed = input.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(decomposed.Length);
        foreach (char ch in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                sb.Append(ch);
        }
        return sb.ToString().Normalize(NormalizationForm.FormC).Trim().ToLowerInvariant();
    }
}
