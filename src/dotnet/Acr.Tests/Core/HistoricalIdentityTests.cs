using Acr.Core.Models;
using Acr.Core.Services;

namespace Acr.Tests.Core;

public sealed class HistoricalIdentityTests
{
    [Fact]
    public void NormalizeAuthorName_RecognizesAllJostinVariants_AndRejectsOtherAuthors()
    {
        Assert.True(HistoricalIdentityMatcher.IsTargetAuthor("ELABORADO POR: \nJostin Lopez Sobalbarro"));
        Assert.True(HistoricalIdentityMatcher.IsTargetAuthor("ELABORADO POR:\r\nJostin Lopez Sobalbaro"));
        Assert.True(HistoricalIdentityMatcher.IsTargetAuthor("Jostin Lopez"));
        Assert.True(HistoricalIdentityMatcher.IsTargetAuthor("Jostin López"));
        Assert.Equal("Jostin Lopez Sobalbarro", HistoricalIdentityMatcher.NormalizeAuthorName("Jostin Lopez Sobalbaro"));

        Assert.False(HistoricalIdentityMatcher.IsTargetAuthor("Angie Montero Rojas"));
        Assert.False(HistoricalIdentityMatcher.IsTargetAuthor("Hilario Lopez Rivera"));
    }

    [Fact]
    public void PrimaryKey_SeparatesSharedPo4300007562_IntoThreeDistinctIdentities()
    {
        var r224 = new BusinessIdentity
        {
            NormalizedContainerId = "IMP-224",
            Pedido41 = "4100002293",
            Po43Numbers = ["4300007562"]
        };
        var r225 = new BusinessIdentity
        {
            NormalizedContainerId = "IMP-225",
            Pedido41 = "4100002296",
            Po43Numbers = ["4300007562"]
        };
        var r226 = new BusinessIdentity
        {
            NormalizedContainerId = "IMP-226",
            Pedido41 = "4100002297",
            Po43Numbers = ["4300007562"]
        };

        var keys = new HashSet<BusinessIdentityKey> { r224.ToKey(), r225.ToKey(), r226.ToKey() };
        Assert.Equal(3, keys.Count);
    }

    [Fact]
    public void Idempotency_MovingOrCopyingFileWithDifferentPathFilenameHash_ProducesZeroDuplicates()
    {
        var recordOriginal = BuildRecord("IMP-104", "4100002248", ["4300008021"], @"IMP-104\IMP-104.xlsm", "SHA-AAA", 660, 660, 124, 124, 1);
        var recordMovedCopy = BuildRecord("IMP-104", "4100002248", ["4300008021"], @"Archive\Copy of IMP-104.xlsm", "SHA-BBB", 660, 660, 124, 124, 1);

        IReadOnlyList<CollapsedHistoricalCandidate> collapsed = HistoricalIdentityMatcher.CollapseDuplicates([recordOriginal, recordMovedCopy]);
        CollapsedHistoricalCandidate single = Assert.Single(collapsed);
        Assert.Equal(2, single.AllProvenances.Count);

        var existingLog = new List<DailyLogEntry>
        {
            new() { ContainerId = "IMP-104", NormalizedContainerId = "IMP-104", Pedido41 = "4100002248" }
        };

        // Re-syncing single moved/copied file against existing log is an idempotent no-op
        IReadOnlyList<CollapsedHistoricalCandidate> movedOnly = HistoricalIdentityMatcher.CollapseDuplicates([recordMovedCopy]);
        IReadOnlyList<HistoricalMatchResult> matchResults = HistoricalIdentityMatcher.MatchAgainstDailyLog(movedOnly, existingLog);
        HistoricalMatchResult match = Assert.Single(matchResults);

        Assert.Equal(HistoricalMatchKind.AlreadySyncedExact, match.MatchKind);
        Assert.True(match.IsIdempotentNoOp);
    }

    [Fact]
    public void CollapseDuplicates_Imp182_SelectsCleanFinalWinner_PreservesAllProvenances_AndFlagsMultipleVersions()
    {
        var cleanFinal = BuildRecord(
            "IMP-182", "4100002283", ["4300006764"], @"IMP-182\IMP-182.xlsm", "SHA-FINAL",
            927, 927, 136, 136, 9, isDraft: false, revisionDate: new DateOnly(2026, 11, 3));
        var draft1 = BuildRecord(
            "IMP-182", "4100002283", ["4300006764"], "Copia de IMP-182 inprogress.xlsm", "SHA-D1",
            17172, 14, 136, 136, 0, isDraft: true, revisionDate: new DateOnly(2026, 11, 2));
        var draft2 = BuildRecord(
            "IMP-182", "4100002283", ["4300006764"], @"IMP-182\IMP-182 inprogress.xlsm", "SHA-D2",
            1133, 14, 136, 136, 0, isDraft: true, revisionDate: new DateOnly(2026, 11, 3));

        IReadOnlyList<CollapsedHistoricalCandidate> collapsed =
            HistoricalIdentityMatcher.CollapseDuplicates([draft1, cleanFinal, draft2]);

        CollapsedHistoricalCandidate winner = Assert.Single(collapsed);
        Assert.Equal(@"IMP-182\IMP-182.xlsm", winner.PrimaryRecord.Provenance.SourcePath);
        Assert.Equal(3, winner.AllProvenances.Count);
        Assert.Contains(HumanReviewReason.MultipleFileVersionsForIdentity, winner.CombinedReviewReasons);
    }

    [Fact]
    public void MatchAgainstDailyLog_MissingPedido41_FallsBackToContainerIdWithMandatoryHumanReview()
    {
        var noPedidoRecord = BuildRecord("IMP-196", null, ["4300006878"], @"IMP-196\IMP-196.xlsm", "SHA-196", 519, 519, 29, 29, 1);
        IReadOnlyList<CollapsedHistoricalCandidate> candidates = HistoricalIdentityMatcher.CollapseDuplicates([noPedidoRecord]);

        var existingLog = new List<DailyLogEntry>
        {
            new() { ContainerId = "IMP-196", NormalizedContainerId = "IMP-196", Pedido41 = "4100002205" }
        };

        HistoricalMatchResult match = Assert.Single(HistoricalIdentityMatcher.MatchAgainstDailyLog(candidates, existingLog));
        Assert.Equal(HistoricalMatchKind.MatchedByContainerFallbackRequiresReview, match.MatchKind);
        Assert.True(match.RequiresHumanReview);
        Assert.False(match.IsIdempotentNoOp);
        Assert.Contains(HumanReviewReason.MissingPedido41, match.MatchReviewReasons);
    }

    [Fact]
    public void MatchAgainstDailyLog_Imp202DuplicateLogRows_TriggersAmbiguousReview()
    {
        var rec202 = BuildRecord("IMP-202", "4100002332", ["4300006692"], @"IMP-202\IMP-202.xlsm", "SHA-202", 400, 400, 12, 12, 1);
        IReadOnlyList<CollapsedHistoricalCandidate> candidates = HistoricalIdentityMatcher.CollapseDuplicates([rec202]);

        var existingLog = new List<DailyLogEntry>
        {
            new() { SequenceNumber = 1, ContainerId = "IMP-202", Pedido41 = "4100002100" },
            new() { SequenceNumber = 16, ContainerId = "IMP-202", Pedido41 = "4100002332" }
        };

        HistoricalMatchResult match = Assert.Single(HistoricalIdentityMatcher.MatchAgainstDailyLog(candidates, existingLog));
        Assert.Equal(HistoricalMatchKind.ConflictOrAmbiguousRequiresReview, match.MatchKind);
        Assert.True(match.RequiresHumanReview);
        Assert.Contains(HumanReviewReason.AmbiguousDailyLogMatch, match.MatchReviewReasons);
        Assert.Equal(2, match.MatchedDailyLogEntries.Count);
    }

    [Fact]
    public void EvaluateAnomalies_And_ParsePoList_HandleAllPhase0Exceptions()
    {
        Assert.Equal(["4300007309", "4300008144"], HistoricalIdentityMatcher.ParsePoList("4300007309-8144"));
        Assert.Equal(["4300006764", "4300006804", "4300006805"], HistoricalIdentityMatcher.ParsePoList("4300006764-6804-6805"));

        var imp195 = BuildRecord("IMP-195", "4100002284", ["4300006505"], @"IMP-195 SIN CONTROL\IMP-195.xlsm", "S1", 0, 0, 154, 154, 15);
        Assert.Contains(HumanReviewReason.ZeroCartonsWithPositiveLines, HistoricalIdentityMatcher.EvaluateAnomalies(imp195));

        var imp241 = BuildRecord("IMP-241", "4100002294", ["4300007076"], @"IMP-241 Sin Control Calidad\IMP-241.xlsm", "S2", null, null, 0, 62, 7, errors: ["#REF!"]);
        IReadOnlyList<HumanReviewReason> imp241Anomalies = HistoricalIdentityMatcher.EvaluateAnomalies(imp241);
        Assert.Contains(HumanReviewReason.FormulaOrLiteralErrorInSummary, imp241Anomalies);
        Assert.Contains(HumanReviewReason.LineCountMismatchDgVsCierreRecep, imp241Anomalies);

        var imp325 = BuildRecord("IMP-325", "4100002301", ["4300007749"], @"IMP-325 PEND CONTROL\IMP-325.xlsm", "S3", 664, 664, 6, 6, 1);
        Assert.Contains(HumanReviewReason.FolderStatusVsCartonMismatch, HistoricalIdentityMatcher.EvaluateAnomalies(imp325));

        var imp326 = BuildRecord("IMP-326", "4100002302", ["4300007309", "4300008144"], @"IMP-326\IMP-326.xlsm", "S4", 669, 669, 6, 6, 1);
        Assert.Contains(HumanReviewReason.PoCountMismatchDgVsCierreRecep, HistoricalIdentityMatcher.EvaluateAnomalies(imp326));
    }

    [Fact]
    public void ParsePoList_ExpandsSpacedHyphens_AndIgnoresStandaloneHyphen()
    {
        Assert.Equal(["4300006764", "4300006804"], HistoricalIdentityMatcher.ParsePoList("4300006764 - 6804"));
        Assert.Equal(["4300006764", "4300006804", "4300006805"], HistoricalIdentityMatcher.ParsePoList("4300006764 - 6804 - 6805"));
        Assert.Empty(HistoricalIdentityMatcher.ParsePoList("-"));
    }

    private static HistoricalSyncRecord BuildRecord(
        string containerId,
        string? pedido41,
        IReadOnlyList<string> po43Numbers,
        string sourcePath,
        string sha,
        int? theoreticalCartons,
        int? receivedCartons,
        int? dgLines,
        int crItems,
        int crPoCount,
        bool isDraft = false,
        DateOnly? revisionDate = null,
        IReadOnlyList<string>? errors = null)
    {
        return new HistoricalSyncRecord
        {
            Identity = new BusinessIdentity
            {
                NormalizedContainerId = containerId,
                Pedido41 = pedido41,
                Po43Numbers = po43Numbers,
                RevisionDate = revisionDate ?? new DateOnly(2026, 11, 3)
            },
            Provenance = new FileProvenance
            {
                SourceFileName = Path.GetFileName(sourcePath),
                SourcePath = sourcePath,
                FileSha256 = sha,
                DiscoveredAtUtc = new DateTimeOffset(2026, 11, 18, 10, 0, 0, TimeSpan.Zero),
                RawAuthorCellB33 = "ELABORADO POR: \nJostin Lopez Sobalbarro",
                NormalizedAuthor = HistoricalIdentityMatcher.CanonicalTargetAuthor,
                CoreXmlCreator = "Hilario Lopez Rivera"
            },
            TheoreticalCartons = theoreticalCartons,
            ReceivedCartons = receivedCartons,
            DeclaredLinesCountDg = dgLines,
            ActualCierreRecepItemsCount = crItems,
            CierreRecepDistinctPoCount = crPoCount,
            IsDraftOrCopyFile = isDraft,
            DetectedErrorTokens = errors ?? Array.Empty<string>()
        };
    }
}
