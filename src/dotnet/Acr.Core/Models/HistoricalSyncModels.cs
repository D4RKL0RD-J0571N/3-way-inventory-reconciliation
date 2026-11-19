namespace Acr.Core.Models;

/// <summary>
/// Composite business key separating identity from file provenance.
/// </summary>
public sealed record BusinessIdentityKey
{
    /// <summary>Normalized container identifier (e.g., <c>IMP-104</c>).</summary>
    public string NormalizedContainerId { get; }

    /// <summary>Normalized Pedido 41 number (empty string when absent).</summary>
    public string Pedido41 { get; }

    /// <summary>Initializes a case-insensitive canonical business identity key.</summary>
    public BusinessIdentityKey(string normalizedContainerId, string? pedido41)
    {
        NormalizedContainerId = (normalizedContainerId ?? "").Trim().ToUpperInvariant();
        Pedido41 = (pedido41 ?? "").Trim();
    }
}

/// <summary>
/// Operational business identity extracted from closure workbook business cells (<c>DATOS GENERALES!C7/C8/C9/C30</c>).
/// </summary>
public sealed record BusinessIdentity
{
    /// <summary>Normalized container identifier from <c>DG!C7</c>.</summary>
    public string NormalizedContainerId { get; init; } = "";

    /// <summary>Nationalization Pedido 41 from <c>DG!C9</c> (null if missing or invalid).</summary>
    public string? Pedido41 { get; init; }

    /// <summary>Purchase Order 43 numbers from <c>DG!C8</c> (never used alone as primary identity).</summary>
    public IReadOnlyList<string> Po43Numbers { get; init; } = Array.Empty<string>();

    /// <summary>Revision date from <c>DG!C30</c>.</summary>
    public DateOnly? RevisionDate { get; init; }

    /// <summary>True when both normalized container ID and Pedido 41 are populated.</summary>
    public bool HasCompositeKey =>
        !string.IsNullOrWhiteSpace(NormalizedContainerId) &&
        !string.IsNullOrWhiteSpace(Pedido41);

    /// <summary>Constructs the canonical composite key for idempotency matching.</summary>
    public BusinessIdentityKey ToKey() => new(NormalizedContainerId, Pedido41);
}

/// <summary>
/// File origin and discovery metadata strictly separated from <see cref="BusinessIdentity"/>.
/// </summary>
public sealed record FileProvenance
{
    /// <summary>Source workbook filename.</summary>
    public string SourceFileName { get; init; } = "";

    /// <summary>Relative or full source path at discovery time.</summary>
    public string SourcePath { get; init; } = "";

    /// <summary>Optional workbook identifier metadata (provenance only).</summary>
    public string? WorkbookIdentifier { get; init; }

    /// <summary>SHA-256 hash of the workbook file bytes.</summary>
    public string FileSha256 { get; init; } = "";

    /// <summary>SHA-256 hash of <c>xl/vbaProject.bin</c> if present.</summary>
    public string? VbaProjectSha256 { get; init; }

    /// <summary>UTC timestamp when the file was discovered.</summary>
    public DateTimeOffset DiscoveredAtUtc { get; init; }

    /// <summary>Raw author string extracted strictly from <c>'DATOS GENERALES'!B33</c>.</summary>
    public string? RawAuthorCellB33 { get; init; }

    /// <summary>Normalized author name resolved from <see cref="RawAuthorCellB33"/>.</summary>
    public string? NormalizedAuthor { get; init; }

    /// <summary>Creator from <c>docProps/core.xml</c> (strictly provenance, never used for authorship filtering).</summary>
    public string? CoreXmlCreator { get; init; }
}

/// <summary>
/// Reasons triggering mandatory human review during Historical Sync.
/// </summary>
public enum HumanReviewReason
{
    /// <summary>Container ID is missing or invalid.</summary>
    MissingContainerId = 0,

    /// <summary>Pedido 41 is missing, requiring fallback review.</summary>
    MissingPedido41 = 1,

    /// <summary>Multiple workbook files/drafts resolved to the same business identity.</summary>
    MultipleFileVersionsForIdentity = 2,

    /// <summary>Theoretical and received cartons are both zero despite positive line counts.</summary>
    ZeroCartonsWithPositiveLines = 3,

    /// <summary>Summary cells contain Excel error tokens (#REF!, #VALUE!, etc.) or failed to parse.</summary>
    FormulaOrLiteralErrorInSummary = 4,

    /// <summary>DG!C18 reports 0 lines while CIERRE RECEP contains positive items.</summary>
    LineCountMismatchDgVsCierreRecep = 5,

    /// <summary>CIERRE RECEP column A contains zero distinct PO numbers (e.g., draft workbook).</summary>
    EmptyCierreRecepPoColumn = 6,

    /// <summary>Folder path indicates PEND CONTROL / SIN CONTROL despite positive carton counts.</summary>
    FolderStatusVsCartonMismatch = 7,

    /// <summary>Distinct PO count in DG!C8 differs from CIERRE RECEP column A.</summary>
    PoCountMismatchDgVsCierreRecep = 8,

    /// <summary>Daily Log contains multiple matching rows or conflicting Pedido 41 values.</summary>
    AmbiguousDailyLogMatch = 9,

    /// <summary>Large or negative carton discrepancy detected between theoretical and received totals.</summary>
    NegativeOrExtremeCartonDiscrepancy = 10
}

/// <summary>
/// Extracted historical closure candidate combining separated <see cref="BusinessIdentity"/> and <see cref="FileProvenance"/>.
/// </summary>
public sealed record HistoricalSyncRecord
{
    /// <summary>Separated operational business identity.</summary>
    public required BusinessIdentity Identity { get; init; }

    /// <summary>Separated file discovery provenance.</summary>
    public required FileProvenance Provenance { get; init; }

    /// <summary>Theoretical cartons from <c>DG!C13</c> (null if error or missing).</summary>
    public int? TheoreticalCartons { get; init; }

    /// <summary>Received cartons from <c>DG!C14</c> (null if error or missing).</summary>
    public int? ReceivedCartons { get; init; }

    /// <summary>Declared lines count from <c>DG!C18</c>.</summary>
    public int? DeclaredLinesCountDg { get; init; }

    /// <summary>Actual item count scanned from <c>CIERRE RECEP</c>.</summary>
    public int ActualCierreRecepItemsCount { get; init; }

    /// <summary>Distinct PO count found in <c>CIERRE RECEP</c> column A.</summary>
    public int CierreRecepDistinctPoCount { get; init; }

    /// <summary>Raw vendor name from <c>DG!C10</c>.</summary>
    public string? VendorNameRaw { get; init; }

    /// <summary>True when workbook embeds both <c>CIERRE RECEP</c> and <c>CONTROL DATA</c>.</summary>
    public bool IsHybridQcClosure { get; init; }

    /// <summary>True when filename indicates draft/copy (<c>Copia de</c>, <c>inprogress</c>, <c>borrador</c>).</summary>
    public bool IsDraftOrCopyFile { get; init; }

    /// <summary>Excel error tokens detected in summary or header cells.</summary>
    public IReadOnlyList<string> DetectedErrorTokens { get; init; } = Array.Empty<string>();

    /// <summary>Anomaly reasons detected on this individual record.</summary>
    public IReadOnlyList<HumanReviewReason> ReviewReasons { get; init; } = Array.Empty<HumanReviewReason>();

    /// <summary>True when at least one human review reason is present.</summary>
    public bool RequiresHumanReview => ReviewReasons.Count > 0;
}

/// <summary>
/// Result of collapsing duplicate/draft historical files for the same business identity.
/// </summary>
public sealed record CollapsedHistoricalCandidate
{
    /// <summary>Highest-ranked canonical closure record selected as primary.</summary>
    public required HistoricalSyncRecord PrimaryRecord { get; init; }

    /// <summary>All provenance records (primary plus collapsed drafts/copies).</summary>
    public IReadOnlyList<FileProvenance> AllProvenances { get; init; } = Array.Empty<FileProvenance>();

    /// <summary>Combined human review reasons across primary selection and collapsing.</summary>
    public IReadOnlyList<HumanReviewReason> CombinedReviewReasons { get; init; } = Array.Empty<HumanReviewReason>();

    /// <summary>True when human review is required before committing to the Daily Log.</summary>
    public bool RequiresHumanReview => CombinedReviewReasons.Count > 0;
}

/// <summary>
/// Idempotency match outcome when comparing a collapsed historical candidate against existing Daily Log entries.
/// </summary>
public enum HistoricalMatchKind
{
    /// <summary>No matching container exists in the Daily Log; eligible as a new record.</summary>
    NewRecord = 0,

    /// <summary>Exact composite identity (NormalizedContainerId + Pedido41) already exists in the Daily Log.</summary>
    AlreadySyncedExact = 1,

    /// <summary>Matched by container ID alone because Pedido 41 was missing on either side; requires human review.</summary>
    MatchedByContainerFallbackRequiresReview = 2,

    /// <summary>Multiple Daily Log entries or conflicting Pedido 41 values detected; requires human review.</summary>
    ConflictOrAmbiguousRequiresReview = 3
}

/// <summary>
/// Final idempotency matching result pairing a collapsed historical candidate with existing Daily Log rows.
/// </summary>
public sealed record HistoricalMatchResult
{
    /// <summary>Collapsed historical candidate evaluated.</summary>
    public required CollapsedHistoricalCandidate Candidate { get; init; }

    /// <summary>Classification of how the candidate relates to the Daily Log.</summary>
    public HistoricalMatchKind MatchKind { get; init; }

    /// <summary>Matching Daily Log rows found during comparison.</summary>
    public IReadOnlyList<DailyLogEntry> MatchedDailyLogEntries { get; init; } = Array.Empty<DailyLogEntry>();

    /// <summary>Combined candidate and matching human review reasons.</summary>
    public IReadOnlyList<HumanReviewReason> MatchReviewReasons { get; init; } = Array.Empty<HumanReviewReason>();

    /// <summary>True when any human review reason applies or match is fallback/conflict.</summary>
    public bool RequiresHumanReview =>
        MatchReviewReasons.Count > 0 ||
        MatchKind is HistoricalMatchKind.MatchedByContainerFallbackRequiresReview
            or HistoricalMatchKind.ConflictOrAmbiguousRequiresReview;

    /// <summary>True when exact match already exists in the Daily Log with zero human review blockers.</summary>
    public bool IsIdempotentNoOp =>
        MatchKind == HistoricalMatchKind.AlreadySyncedExact && !RequiresHumanReview;
}
