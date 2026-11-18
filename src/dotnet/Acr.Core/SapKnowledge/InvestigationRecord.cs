namespace Acr.Core.SapKnowledge;

/// <summary>
/// Individual finding recorded during an operational investigation, strictly categorized by <see cref="RootCauseState"/>.
/// </summary>
/// <param name="FindingId">Unique finding identifier.</param>
/// <param name="Statement">Factual observation, suspected hypothesis, or confirmed root cause statement.</param>
/// <param name="State">Distinguishes observed symptom vs suspected cause vs confirmed root cause.</param>
/// <param name="EvidenceSource">Evidence provenance supporting this finding.</param>
/// <param name="RecordedUtc">Explicit UTC timestamp provided by caller.</param>
public sealed record InvestigationFinding(
    string FindingId,
    string Statement,
    RootCauseState State,
    SourceProvenance? EvidenceSource,
    DateTimeOffset RecordedUtc);

/// <summary>
/// Persistent investigation record capturing context, findings, external actions, verification,
/// and promotion gate to personal SAP knowledge entries (R5-G).
/// </summary>
public sealed record InvestigationRecord
{
    /// <summary>Unique investigation record ID.</summary>
    public required string RecordId { get; init; }

    /// <summary>Explicit UTC creation timestamp.</summary>
    public DateTimeOffset CreatedUtc { get; init; }

    /// <summary>Explicit UTC last-update timestamp.</summary>
    public DateTimeOffset UpdatedUtc { get; init; }

    /// <summary>Verified operational context attached to this investigation.</summary>
    public required DiagnosticContext Context { get; init; }

    /// <summary>Optional <see cref="DiagnosticGuide.GuideId"/> used during the investigation.</summary>
    public string? GuideId { get; init; }

    /// <summary>All recorded findings across observed, suspected, and confirmed states.</summary>
    public IReadOnlyList<InvestigationFinding> Findings { get; init; } = Array.Empty<InvestigationFinding>();

    /// <summary>Findings strictly classified as observed discrepancies.</summary>
    public IReadOnlyList<InvestigationFinding> ObservedDiscrepancies =>
        Findings.Where(f => f.State == RootCauseState.ObservedDiscrepancy).ToList();

    /// <summary>Findings strictly classified as suspected causes (unconfirmed hypotheses).</summary>
    public IReadOnlyList<InvestigationFinding> SuspectedCauses =>
        Findings.Where(f => f.State == RootCauseState.SuspectedCause).ToList();

    /// <summary>Findings strictly classified as confirmed root causes.</summary>
    public IReadOnlyList<InvestigationFinding> ConfirmedRootCauses =>
        Findings.Where(f => f.State == RootCauseState.ConfirmedRootCause).ToList();

    /// <summary>Highest confirmed/suspected/observed state reached in this investigation.</summary>
    public RootCauseState CurrentRootCauseState =>
        ConfirmedRootCauses.Count > 0
            ? RootCauseState.ConfirmedRootCause
            : (SuspectedCauses.Count > 0 ? RootCauseState.SuspectedCause : RootCauseState.ObservedDiscrepancy);

    /// <summary>Human actions executed externally in SAP or warehouse operations (ACR never executes SAP writes).</summary>
    public IReadOnlyList<string> ExternalActionsTaken { get; init; } = Array.Empty<string>();

    /// <summary>Final verification status.</summary>
    public InvestigationVerificationStatus VerificationResult { get; init; } = InvestigationVerificationStatus.Pending;

    /// <summary>Optional verification notes.</summary>
    public string? VerificationNotes { get; init; }

    /// <summary>Identifier of the personal knowledge entry if promoted.</summary>
    public string? PromotedKnowledgeEntryId { get; init; }

    /// <summary>
    /// Determines whether this investigation meets all preconditions for promotion to a personal KB entry.
    /// </summary>
    public bool IsEligibleForPersonalKbPromotion(
        bool explicitUserConfirmation,
        IReadOnlyList<SourceProvenance>? additionalReferences = null)
    {
        if (!explicitUserConfirmation ||
            ConfirmedRootCauses.Count == 0 ||
            VerificationResult != InvestigationVerificationStatus.VerifiedResolved)
        {
            return false;
        }

        return CollectProvenances(additionalReferences).Count > 0;
    }

    /// <summary>
    /// Promotes a verified investigation with explicit user confirmation into a reusable personal <see cref="SapKnowledgeEntry"/>.
    /// </summary>
    public SapKnowledgeEntry PromoteToPersonalKnowledgeEntry(
        string entryId,
        string title,
        bool explicitUserConfirmation,
        bool isEnvironmentValidated,
        IReadOnlyList<SourceProvenance>? additionalReferences = null)
    {
        if (string.IsNullOrWhiteSpace(entryId))
            throw new ArgumentException("Entry ID is required.", nameof(entryId));
        if (string.IsNullOrWhiteSpace(title))
            throw new ArgumentException("Title is required.", nameof(title));

        if (!IsEligibleForPersonalKbPromotion(explicitUserConfirmation, additionalReferences))
        {
            throw new InvalidOperationException(
                "Promotion requires explicit user confirmation, at least one ConfirmedRootCause, VerifiedResolved status, and source provenance.");
        }

        List<SourceProvenance> provenances = CollectProvenances(additionalReferences);
        return BuildPromotedEntry(entryId.Trim(), title.Trim(), isEnvironmentValidated, provenances);
    }

    private List<SourceProvenance> CollectProvenances(IReadOnlyList<SourceProvenance>? additionalReferences)
    {
        var list = new List<SourceProvenance>();
        foreach (InvestigationFinding finding in ConfirmedRootCauses)
        {
            if (finding.EvidenceSource is not null)
                list.Add(finding.EvidenceSource);
        }
        if (additionalReferences is not null)
            list.AddRange(additionalReferences);
        return list;
    }

    private SapKnowledgeEntry BuildPromotedEntry(
        string entryId,
        string title,
        bool isEnvironmentValidated,
        List<SourceProvenance> provenances)
    {
        string description = string.Join(" | ", ConfirmedRootCauses.Select(f => f.Statement));
        return new SapKnowledgeEntry
        {
            EntryId = entryId,
            Title = title,
            Description = description,
            Category = SapKnowledgeCategory.PersonalOperationalNote,
            ObjectKind = SapObjectKind.CustomWorkflow,
            EpistemicTier = isEnvironmentValidated ? EpistemicTier.EnvironmentValidated : EpistemicTier.CompanyVerified,
            LifecycleState = isEnvironmentValidated ? FragmentLifecycleState.EnvironmentValidated : FragmentLifecycleState.Reviewed,
            OperationMode = SapOperationMode.ReadOnly,
            BusinessObjects = Context.GetInferredBusinessObjects(),
            DiscrepancyTypes = Context.DiscrepancyType.HasValue ? [Context.DiscrepancyType.Value] : Array.Empty<Models.DiscrepancyType>(),
            ProcedureSteps = ExternalActionsTaken,
            VerificationSteps = string.IsNullOrWhiteSpace(VerificationNotes) ? Array.Empty<string>() : [VerificationNotes],
            SourceProvenances = provenances,
            Tags = ["personal-note", "promoted-investigation", RecordId]
        };
    }
}
