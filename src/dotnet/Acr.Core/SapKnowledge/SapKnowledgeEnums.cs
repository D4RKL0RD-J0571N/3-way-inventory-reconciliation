namespace Acr.Core.SapKnowledge;

/// <summary>
/// Epistemic authority tier distinguishing verified company procedures from general SAP references
/// and unverified hypotheses.
/// </summary>
public enum EpistemicTier
{
    /// <summary>Verified corporate SOP, workbook macro, or validated company procedure.</summary>
    CompanyVerified = 0,

    /// <summary>Procedure confirmed directly in the live company SAP environment.</summary>
    EnvironmentValidated = 1,

    /// <summary>Official SAP documentation (e.g., SAP Help Portal); verify local environment compatibility.</summary>
    OfficialSapDoc = 2,

    /// <summary>General SAP MM/LE-WM/IS-Retail reference material not yet validated in the company environment.</summary>
    GeneralSapReference = 3,

    /// <summary>Unverified hypothesis or possible explanation; never presented as an established SAP procedure.</summary>
    UnverifiedHypothesis = 4
}

/// <summary>
/// Explicit governance lifecycle state for imported and authored SAP knowledge fragments.
/// </summary>
public enum FragmentLifecycleState
{
    /// <summary>Document imported into the library, not yet extracted or validated.</summary>
    Imported = 0,

    /// <summary>Content automatically extracted from source document.</summary>
    Extracted = 1,

    /// <summary>Content reviewed by the operator for accuracy and provenance.</summary>
    Reviewed = 2,

    /// <summary>Procedure confirmed directly in the company's live SAP environment.</summary>
    EnvironmentValidated = 3,

    /// <summary>Outdated information excluded from default search results.</summary>
    Deprecated = 4
}

/// <summary>
/// Diagnostic certainty state separating observed symptoms from hypotheses and confirmed root causes.
/// </summary>
public enum RootCauseState
{
    /// <summary>Verifiable symptom observed in QC, Closure, Daily Log, or SAP extract.</summary>
    ObservedDiscrepancy = 0,

    /// <summary>Candidate cause or hypothesis currently under investigation.</summary>
    SuspectedCause = 1,

    /// <summary>Confirmed root cause backed by documented verification evidence.</summary>
    ConfirmedRootCause = 2
}

/// <summary>
/// Functional categories organizing the local SAP Knowledge Library.
/// </summary>
public enum SapKnowledgeCategory
{
    /// <summary>SAP transaction reference guide.</summary>
    TransactionReferenceGuide = 0,

    /// <summary>Step-by-step operational procedure.</summary>
    OperationalProcedure = 1,

    /// <summary>Troubleshooting guide and known operational incidents.</summary>
    TroubleshootingAndIncident = 2,

    /// <summary>SAP tables, fields, and relational data model.</summary>
    TableAndDataRelationship = 3,

    /// <summary>Company-specific SOPs and customized Z-transactions.</summary>
    CompanySopAndCustomTransaction = 4,

    /// <summary>Personal operational notes and validated discoveries.</summary>
    PersonalOperationalNote = 5,

    /// <summary>Reference documentation and source links.</summary>
    ReferenceDocumentationAndSourceLink = 6
}

/// <summary>
/// Structural classification of an SAP object in the encyclopedia and knowledge base.
/// </summary>
public enum SapObjectKind
{
    /// <summary>Standard SAP transaction code (e.g., ME80FN, SE16N, VL31N, RFUI).</summary>
    StandardTransaction = 0,

    /// <summary>Company-specific custom Z-transaction or custom report (e.g., ZDescarga).</summary>
    ZTransaction = 1,

    /// <summary>SAP database table queried via SE16N (e.g., EINA, MAST, STPO).</summary>
    SapTable = 2,

    /// <summary>Company-specific operational workflow or local business convention.</summary>
    CustomWorkflow = 3,

    /// <summary>Practical SAP GUI 7.00 navigation and data discovery guide.</summary>
    NavigationGuide = 4
}

/// <summary>
/// Whether an SAP transaction or procedure is read-only or modifies SAP data.
/// </summary>
public enum SapOperationMode
{
    /// <summary>Strictly read-only query or inspection.</summary>
    ReadOnly = 0,

    /// <summary>Creates, updates, or reverses SAP documents or inventory movements.</summary>
    Modifying = 1,

    /// <summary>Display-oriented transaction that also supports conditional write actions.</summary>
    ReadOnlyWithConditionalWrite = 2
}

/// <summary>
/// Underlying SAP backend release (explicitly decoupled from SAP GUI client version).
/// </summary>
public enum SapBackendRelease
{
    /// <summary>Backend release pending live verification via SAP GUI System -> Status.</summary>
    UnknownPendingVerification = 0,

    /// <summary>SAP ECC 6.0 backend.</summary>
    Ecc600 = 1,

    /// <summary>SAP S/4HANA backend.</summary>
    S4Hana = 2,

    /// <summary>Concept or rule independent of ECC vs S/4HANA release.</summary>
    ReleaseAgnostic = 3
}

/// <summary>
/// Verification outcome recorded for an operational troubleshooting investigation.
/// </summary>
public enum InvestigationVerificationStatus
{
    /// <summary>Verification not yet completed.</summary>
    Pending = 0,

    /// <summary>Issue verified resolved with documented evidence.</summary>
    VerifiedResolved = 1,

    /// <summary>Partially verified; follow-up checks remain.</summary>
    PartiallyVerified = 2,

    /// <summary>Investigation could not resolve the issue.</summary>
    Unresolved = 3
}

/// <summary>
/// Receiving module from which a <see cref="DiagnosticContext"/> originated.
/// </summary>
public enum DiagnosticContextOrigin
{
    /// <summary>Launched from a Reconciliation discrepancy row.</summary>
    Reconciliation = 0,

    /// <summary>Launched from a Daily Log container row.</summary>
    DailyLog = 1,

    /// <summary>Launched from a Historical Sync candidate.</summary>
    HistoricalSync = 2,

    /// <summary>Created from a manual search or troubleshooting query.</summary>
    ManualQuery = 3
}

/// <summary>
/// Deterministic lifecycle transition rules for <see cref="FragmentLifecycleState"/>.
/// </summary>
public static class FragmentLifecycleRules
{
    /// <summary>
    /// Determines whether transitioning from <paramref name="from"/> to <paramref name="to"/> is permitted.
    /// </summary>
    public static bool CanTransition(FragmentLifecycleState from, FragmentLifecycleState to)
    {
        if (from == to)
            return false;

        return from switch
        {
            FragmentLifecycleState.Imported => to is FragmentLifecycleState.Extracted
                or FragmentLifecycleState.Reviewed
                or FragmentLifecycleState.Deprecated,
            FragmentLifecycleState.Extracted => to is FragmentLifecycleState.Reviewed
                or FragmentLifecycleState.Deprecated,
            FragmentLifecycleState.Reviewed => to is FragmentLifecycleState.EnvironmentValidated
                or FragmentLifecycleState.Deprecated,
            FragmentLifecycleState.EnvironmentValidated => to is FragmentLifecycleState.Reviewed
                or FragmentLifecycleState.Deprecated,
            FragmentLifecycleState.Deprecated => to == FragmentLifecycleState.Reviewed,
            _ => false
        };
    }

    /// <summary>
    /// Validates a fragment lifecycle transition and throws <see cref="InvalidOperationException"/> when disallowed.
    /// </summary>
    public static FragmentLifecycleState ValidateTransition(FragmentLifecycleState from, FragmentLifecycleState to)
    {
        if (!CanTransition(from, to))
        {
            throw new InvalidOperationException(
                $"Lifecycle transition from '{from}' to '{to}' is not permitted.");
        }
        return to;
    }
}

/// <summary>
/// Deterministic ranking and caution rules for <see cref="EpistemicTier"/>.
/// </summary>
public static class EpistemicTierRules
{
    /// <summary>
    /// Returns the deterministic search ranking weight associated with <paramref name="tier"/>.
    /// </summary>
    public static int GetRankingWeight(EpistemicTier tier) => tier switch
    {
        EpistemicTier.CompanyVerified => 30,
        EpistemicTier.EnvironmentValidated => 25,
        EpistemicTier.OfficialSapDoc => 20,
        EpistemicTier.GeneralSapReference => 10,
        EpistemicTier.UnverifiedHypothesis => 0,
        _ => 0
    };

    /// <summary>
    /// Returns the deterministic ranking adjustment associated with <paramref name="state"/>.
    /// </summary>
    public static int GetLifecycleWeight(FragmentLifecycleState state) => state switch
    {
        FragmentLifecycleState.EnvironmentValidated => 10,
        FragmentLifecycleState.Reviewed => 5,
        FragmentLifecycleState.Extracted => 2,
        FragmentLifecycleState.Imported => 0,
        FragmentLifecycleState.Deprecated => -50,
        _ => 0
    };

    /// <summary>
    /// Returns true when the tier represents authoritative verified/validated knowledge.
    /// </summary>
    public static bool IsAuthoritative(EpistemicTier tier) =>
        tier is EpistemicTier.CompanyVerified or EpistemicTier.EnvironmentValidated;

    /// <summary>
    /// Returns true when UI/diagnostics must display an environment or epistemic caution banner.
    /// </summary>
    public static bool RequiresCautionBanner(EpistemicTier tier, FragmentLifecycleState lifecycleState) =>
        tier is EpistemicTier.UnverifiedHypothesis or EpistemicTier.GeneralSapReference ||
        (tier == EpistemicTier.OfficialSapDoc && lifecycleState != FragmentLifecycleState.EnvironmentValidated);
}
