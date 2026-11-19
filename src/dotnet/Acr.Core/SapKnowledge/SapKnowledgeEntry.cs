using Acr.Core.Models;

namespace Acr.Core.SapKnowledge;

/// <summary>
/// Structured local SAP knowledge base entry carrying all 15 R5-A operational metadata fields,
/// provenance traceability, environment applicability, epistemic tier, and lifecycle state.
/// </summary>
public sealed record SapKnowledgeEntry : ISapRankableDocument
{
    /// <summary>Unique identifier for the knowledge entry (e.g., <c>"sap.tcode.me80fn"</c>).</summary>
    public required string EntryId { get; init; }

    /// <summary>R5-A #1: Descriptive title.</summary>
    public required string Title { get; init; }

    /// <summary>R5-A #2: Detailed description of purpose and scope.</summary>
    public required string Description { get; init; }

    /// <summary>R5-A #3: Primary transaction or table codes covered by this entry.</summary>
    public IReadOnlyList<string> TransactionCodes { get; init; } = Array.Empty<string>();

    /// <summary>R5-A #3: Related or follow-on SAP transactions.</summary>
    public IReadOnlyList<string> RelatedTransactions { get; init; } = Array.Empty<string>();

    /// <summary>R5-A #4: SAP functional module (e.g., <c>"MM"</c>, <c>"LE-WM"</c>, <c>"EWM"</c>).</summary>
    public string SapModule { get; init; } = "";

    /// <summary>R5-A #4: Functional area (e.g., <c>"Receiving"</c>, <c>"Purchasing"</c>).</summary>
    public string FunctionalArea { get; init; } = "";

    /// <summary>R5-A #5: Applicable SAP GUI and backend environment info.</summary>
    public SapEnvironmentInfo EnvironmentInfo { get; init; } = new();

    /// <summary>R5-A #6: Required inputs to execute or query the procedure.</summary>
    public IReadOnlyList<string> RequiredInputs { get; init; } = Array.Empty<string>();

    /// <summary>R5-A #7: Expected outputs or SAP screen/report results.</summary>
    public IReadOnlyList<string> ExpectedOutputs { get; init; } = Array.Empty<string>();

    /// <summary>R5-A #8: SAP GUI 7.00 navigation paths and menu instructions.</summary>
    public IReadOnlyList<string> NavigationInstructions { get; init; } = Array.Empty<string>();

    /// <summary>R5-A #9: Step-by-step operational procedure.</summary>
    public IReadOnlyList<string> ProcedureSteps { get; init; } = Array.Empty<string>();

    /// <summary>R5-A #10: Relevant SAP database tables and field relationships.</summary>
    public IReadOnlyList<SapTableFieldReference> RelevantTablesAndFields { get; init; } = Array.Empty<SapTableFieldReference>();

    /// <summary>R5-A #11: Prerequisites and upstream dependencies.</summary>
    public IReadOnlyList<string> DependenciesAndPrerequisites { get; init; } = Array.Empty<string>();

    /// <summary>R5-A #12: Read-only versus modifying operation classification.</summary>
    public SapOperationMode OperationMode { get; init; } = SapOperationMode.ReadOnly;

    /// <summary>R5-A #13: Required SAP roles, transactions, or authorization objects.</summary>
    public IReadOnlyList<string> AuthorizationRequirements { get; init; } = Array.Empty<string>();

    /// <summary>R5-A #14: Operational consequences and reversal procedures.</summary>
    public IReadOnlyList<string> ConsequencesAndReversals { get; init; } = Array.Empty<string>();

    /// <summary>R5-A #15: Verification steps confirming expected outcome.</summary>
    public IReadOnlyList<string> VerificationSteps { get; init; } = Array.Empty<string>();

    /// <summary>Source document provenance locations.</summary>
    public IReadOnlyList<SourceProvenance> SourceProvenances { get; init; } = Array.Empty<SourceProvenance>();

    /// <summary>Governance lifecycle state of this entry.</summary>
    public FragmentLifecycleState LifecycleState { get; init; } = FragmentLifecycleState.Imported;

    /// <summary>Epistemic authority tier of this entry.</summary>
    public EpistemicTier EpistemicTier { get; init; } = EpistemicTier.GeneralSapReference;

    /// <summary>Knowledge library category.</summary>
    public SapKnowledgeCategory Category { get; init; } = SapKnowledgeCategory.OperationalProcedure;

    /// <summary>Structural SAP object kind.</summary>
    public SapObjectKind ObjectKind { get; init; } = SapObjectKind.StandardTransaction;

    /// <summary>Associated business objects used in deterministic matching.</summary>
    public IReadOnlyList<string> BusinessObjects { get; init; } = Array.Empty<string>();

    /// <summary>Receiving discrepancy types addressed by this entry.</summary>
    public IReadOnlyList<DiscrepancyType> DiscrepancyTypes { get; init; } = Array.Empty<DiscrepancyType>();

    /// <summary>Searchable keywords.</summary>
    public IReadOnlyList<string> Keywords { get; init; } = Array.Empty<string>();

    /// <summary>Domain synonyms.</summary>
    public IReadOnlyList<string> Synonyms { get; init; } = Array.Empty<string>();

    /// <summary>Search tags.</summary>
    public IReadOnlyList<string> Tags { get; init; } = Array.Empty<string>();

    string ISapRankableDocument.Id => EntryId;
    string ISapRankableDocument.Summary => Description;
    string ISapRankableDocument.BodyText => string.Join(" ",
        NavigationInstructions.Concat(ProcedureSteps).Concat(VerificationSteps).Concat(ConsequencesAndReversals));
}
