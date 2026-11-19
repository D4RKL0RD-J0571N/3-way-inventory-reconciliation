using Acr.Core.Models;

namespace Acr.Core.SapKnowledge;

/// <summary>
/// Section 1 of a <see cref="DiagnosticGuide"/>: Problem summary and affected SAP/receiving business objects.
/// </summary>
/// <param name="Summary">Concise description of the operational symptom.</param>
/// <param name="AffectedBusinessObjects">Affected business object types (e.g., PO_43, MATERIAL, HU).</param>
public sealed record DiagnosticProblemSummary(
    string Summary,
    IReadOnlyList<string> AffectedBusinessObjects);

/// <summary>
/// Section 3 item of a <see cref="DiagnosticGuide"/>: Relevant SAP transaction or table step.
/// </summary>
/// <param name="Code">Transaction or table code.</param>
/// <param name="ObjectKind">Classification as standard T-code, Z-transaction, or SAP table.</param>
/// <param name="NavigationPath">Navigation path inside SAP GUI 7.00.</param>
/// <param name="Purpose">What the analyst checks here.</param>
public sealed record DiagnosticTransactionStep(
    string Code,
    SapObjectKind ObjectKind,
    string NavigationPath,
    string Purpose);

/// <summary>
/// Section 6 item of a <see cref="DiagnosticGuide"/>: Possible cause explicitly separated from confirmed facts.
/// </summary>
/// <param name="Description">Candidate explanation or hypothesis.</param>
/// <param name="EpistemicTier">Authority level supporting this candidate cause.</param>
/// <param name="IsConfirmedFact">False unless backed by confirmed evidence in the company environment.</param>
public sealed record DiagnosticPossibleCause(
    string Description,
    EpistemicTier EpistemicTier,
    bool IsConfirmedFact);

/// <summary>
/// Section 7 item of a <see cref="DiagnosticGuide"/>: Recommended human resolution option and operational risks.
/// </summary>
/// <param name="ActionSummary">Recommended resolution action (executed externally by operator).</param>
/// <param name="OperationalRisks">Risks or side effects of taking this action.</param>
/// <param name="IsReadOnlyGuidance">True when guidance involves only read-only checks.</param>
/// <param name="RequiresSapExternalExecution">Always true for SAP modifications (ACR never executes SAP writes).</param>
public sealed record DiagnosticResolutionOption(
    string ActionSummary,
    string OperationalRisks,
    bool IsReadOnlyGuidance,
    bool RequiresSapExternalExecution = true);

/// <summary>
/// Section 10 item of a <see cref="DiagnosticGuide"/>: Documentation citation backing a recommendation.
/// </summary>
/// <param name="Provenance">Exact source document location.</param>
/// <param name="EpistemicTier">Authority tier of the cited source.</param>
/// <param name="SupportedRecommendation">Which recommendation or diagnostic step this citation backs.</param>
public sealed record DiagnosticCitation(
    SourceProvenance Provenance,
    EpistemicTier EpistemicTier,
    string SupportedRecommendation);

/// <summary>
/// Structured 10-section guided operational troubleshooting workflow (R5-C).
/// </summary>
public sealed record DiagnosticGuide : ISapRankableDocument
{
    /// <summary>Unique deterministic guide identifier.</summary>
    public required string GuideId { get; init; }

    /// <summary>Guide title.</summary>
    public required string Title { get; init; }

    /// <summary>Target receiving discrepancy classifications.</summary>
    public IReadOnlyList<DiscrepancyType> TargetDiscrepancyTypes { get; init; } = Array.Empty<DiscrepancyType>();

    /// <summary>Target business objects.</summary>
    public IReadOnlyList<string> TargetBusinessObjects { get; init; } = Array.Empty<string>();

    /// <summary>Transaction or table codes referenced in the guide.</summary>
    public IReadOnlyList<string> TransactionCodes { get; init; } = Array.Empty<string>();

    /// <summary>Section 1: Problem summary and affected business objects.</summary>
    public required DiagnosticProblemSummary ProblemSummaryAndAffectedObjects { get; init; }

    /// <summary>Section 2: Information required to investigate.</summary>
    public IReadOnlyList<string> RequiredInvestigationInfo { get; init; } = Array.Empty<string>();

    /// <summary>Section 3: Relevant SAP transactions and navigation paths.</summary>
    public IReadOnlyList<DiagnosticTransactionStep> RelevantTransactionsAndPaths { get; init; } = Array.Empty<DiagnosticTransactionStep>();

    /// <summary>Section 4: Step-by-step investigation instructions.</summary>
    public IReadOnlyList<string> InvestigationSteps { get; init; } = Array.Empty<string>();

    /// <summary>Section 5: Relevant SAP tables, fields, and relationships.</summary>
    public IReadOnlyList<SapTableFieldReference> RelevantTablesAndRelationships { get; init; } = Array.Empty<SapTableFieldReference>();

    /// <summary>Section 6: Possible causes explicitly distinguished from confirmed facts.</summary>
    public IReadOnlyList<DiagnosticPossibleCause> PossibleCauses { get; init; } = Array.Empty<DiagnosticPossibleCause>();

    /// <summary>Section 7: Recommended resolution paths and operational risks.</summary>
    public IReadOnlyList<DiagnosticResolutionOption> RecommendedResolutionsAndRisks { get; init; } = Array.Empty<DiagnosticResolutionOption>();

    /// <summary>Section 8: Required approvals or human validation.</summary>
    public IReadOnlyList<string> RequiredApprovalsAndValidation { get; init; } = Array.Empty<string>();

    /// <summary>Section 9: Verification steps to confirm resolution.</summary>
    public IReadOnlyList<string> VerificationSteps { get; init; } = Array.Empty<string>();

    /// <summary>Section 10: Documentation references supporting each recommendation.</summary>
    public IReadOnlyList<DiagnosticCitation> DocumentationReferences { get; init; } = Array.Empty<DiagnosticCitation>();

    /// <summary>SAP environment applicability.</summary>
    public SapEnvironmentInfo EnvironmentInfo { get; init; } = new();

    /// <summary>Epistemic authority tier of the guide.</summary>
    public EpistemicTier EpistemicTier { get; init; } = EpistemicTier.CompanyVerified;

    /// <summary>Governance lifecycle state.</summary>
    public FragmentLifecycleState LifecycleState { get; init; } = FragmentLifecycleState.Reviewed;

    /// <summary>Searchable keywords.</summary>
    public IReadOnlyList<string> Keywords { get; init; } = Array.Empty<string>();

    /// <summary>Domain synonyms.</summary>
    public IReadOnlyList<string> Synonyms { get; init; } = Array.Empty<string>();

    /// <summary>Classification tags.</summary>
    public IReadOnlyList<string> Tags { get; init; } = Array.Empty<string>();

    string ISapRankableDocument.Id => GuideId;
    string ISapRankableDocument.Summary => ProblemSummaryAndAffectedObjects.Summary;
    IReadOnlyList<string> ISapRankableDocument.BusinessObjects => TargetBusinessObjects;
    IReadOnlyList<DiscrepancyType> ISapRankableDocument.DiscrepancyTypes => TargetDiscrepancyTypes;
    SapKnowledgeCategory ISapRankableDocument.Category => SapKnowledgeCategory.TroubleshootingAndIncident;
    SapObjectKind ISapRankableDocument.ObjectKind => SapObjectKind.NavigationGuide;
    string ISapRankableDocument.BodyText => string.Join(" ",
        InvestigationSteps.Concat(VerificationSteps).Concat(RequiredApprovalsAndValidation));

    /// <summary>
    /// Verifies that all 10 R5-C diagnostic sections are populated.
    /// </summary>
    public bool HasAllRequiredSections() => Validate().Count == 0;

    /// <summary>
    /// Validates all 10 R5-C sections and epistemic integrity rules, returning any missing/invalid items.
    /// </summary>
    public IReadOnlyList<string> Validate()
    {
        var errors = new List<string>();
        ValidateSummaryAndSteps(errors);
        ValidateCausesResolutionsAndCitations(errors);
        return errors;
    }

    private void ValidateSummaryAndSteps(List<string> errors)
    {
        if (string.IsNullOrWhiteSpace(ProblemSummaryAndAffectedObjects.Summary) ||
            ProblemSummaryAndAffectedObjects.AffectedBusinessObjects.Count == 0 ||
            HasBlankString(ProblemSummaryAndAffectedObjects.AffectedBusinessObjects))
        {
            errors.Add("Section 1 (ProblemSummaryAndAffectedObjects) must include summary and non-blank affected business objects.");
        }
        if (RequiredInvestigationInfo.Count == 0 || HasBlankString(RequiredInvestigationInfo))
            errors.Add("Section 2 (RequiredInvestigationInfo) must not be empty or contain blank items.");
        if (RelevantTransactionsAndPaths.Count == 0 ||
            RelevantTransactionsAndPaths.Any(t =>
                string.IsNullOrWhiteSpace(t.Code) ||
                string.IsNullOrWhiteSpace(t.NavigationPath) ||
                string.IsNullOrWhiteSpace(t.Purpose)))
        {
            errors.Add("Section 3 (RelevantTransactionsAndPaths) must not be empty or contain blank fields.");
        }
        if (InvestigationSteps.Count == 0 || HasBlankString(InvestigationSteps))
            errors.Add("Section 4 (InvestigationSteps) must not be empty or contain blank steps.");
        if (RelevantTablesAndRelationships.Count == 0 ||
            RelevantTablesAndRelationships.Any(r =>
                string.IsNullOrWhiteSpace(r.TableName) ||
                r.FieldNames.Count == 0 ||
                HasBlankString(r.FieldNames)))
        {
            errors.Add("Section 5 (RelevantTablesAndRelationships) must not be empty or contain blank table/field names.");
        }
    }

    private void ValidateCausesResolutionsAndCitations(List<string> errors)
    {
        if (PossibleCauses.Count == 0 || PossibleCauses.Any(c => string.IsNullOrWhiteSpace(c.Description)))
            errors.Add("Section 6 (PossibleCauses) must not be empty or contain blank descriptions.");
        if (PossibleCauses.Any(c => c.IsConfirmedFact && !EpistemicTierRules.IsAuthoritative(c.EpistemicTier)))
            errors.Add("Section 6 violation: non-authoritative possible causes cannot be marked IsConfirmedFact.");
        if (RecommendedResolutionsAndRisks.Count == 0 ||
            RecommendedResolutionsAndRisks.Any(r =>
                string.IsNullOrWhiteSpace(r.ActionSummary) || string.IsNullOrWhiteSpace(r.OperationalRisks)))
        {
            errors.Add("Section 7 (RecommendedResolutionsAndRisks) must not be empty or contain blank entries.");
        }
        if (RequiredApprovalsAndValidation.Count == 0 || HasBlankString(RequiredApprovalsAndValidation))
            errors.Add("Section 8 (RequiredApprovalsAndValidation) must not be empty or contain blank items.");
        if (VerificationSteps.Count == 0 || HasBlankString(VerificationSteps))
            errors.Add("Section 9 (VerificationSteps) must not be empty or contain blank steps.");
        if (DocumentationReferences.Count == 0 ||
            DocumentationReferences.Any(d =>
                string.IsNullOrWhiteSpace(d.Provenance.SourceDocument) ||
                string.IsNullOrWhiteSpace(d.SupportedRecommendation)))
        {
            errors.Add("Section 10 (DocumentationReferences) must not be empty or contain blank citations.");
        }
    }

    private static bool HasBlankString(IReadOnlyList<string> items) =>
        items.Any(string.IsNullOrWhiteSpace);
}
