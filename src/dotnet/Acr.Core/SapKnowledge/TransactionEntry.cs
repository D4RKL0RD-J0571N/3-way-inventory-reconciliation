using Acr.Core.Models;

namespace Acr.Core.SapKnowledge;

/// <summary>
/// Encyclopedia entry describing an SAP transaction, custom Z-transaction, or SE16N-queried SAP table.
/// </summary>
public sealed record TransactionEntry : ISapRankableDocument
{
    /// <summary>Canonical transaction or table code (e.g., <c>"ME80FN"</c>, <c>"ZDescarga"</c>, <c>"EINA"</c>).</summary>
    public required string Code { get; init; }

    /// <summary>Human-readable title.</summary>
    public required string Title { get; init; }

    /// <summary>Structural object kind (<see cref="SapObjectKind.SapTable"/> for EINA/MAST/STPO, <see cref="SapObjectKind.ZTransaction"/> for ZDescarga).</summary>
    public SapObjectKind ObjectKind { get; init; } = SapObjectKind.StandardTransaction;

    /// <summary>True for standard SAP objects; false for company-specific Z-objects or custom workflows.</summary>
    public bool IsStandardSap => ObjectKind is SapObjectKind.StandardTransaction or SapObjectKind.SapTable;

    /// <summary>True when standard SAP reference has not yet been validated in the company's live environment.</summary>
    public bool RequiresCompanyBehaviorCaution =>
        IsStandardSap && EpistemicTier != EpistemicTier.EnvironmentValidated;

    /// <summary>Operational purpose of the transaction or table.</summary>
    public required string Purpose { get; init; }

    /// <summary>When an analyst should use this transaction or table.</summary>
    public string WhenToUse { get; init; } = "";

    /// <summary>SAP GUI 7.00 navigation path.</summary>
    public string NavigationPath { get; init; } = "";

    /// <summary>Required input selection parameters.</summary>
    public IReadOnlyList<string> Inputs { get; init; } = Array.Empty<string>();

    /// <summary>Produced output reports or fields.</summary>
    public IReadOnlyList<string> Outputs { get; init; } = Array.Empty<string>();

    /// <summary>Underlying SAP tables and fields.</summary>
    public IReadOnlyList<SapTableFieldReference> RelevantTables { get; init; } = Array.Empty<SapTableFieldReference>();

    /// <summary>Related operational procedures or follow-on transactions.</summary>
    public IReadOnlyList<string> RelatedProcedures { get; init; } = Array.Empty<string>();

    /// <summary>Operational risks, semantic traps (e.g. VKP2 vs EKPO-NETPR), and precautions.</summary>
    public IReadOnlyList<string> OperationalRisks { get; init; } = Array.Empty<string>();

    /// <summary>Steps to verify extracted or executed results.</summary>
    public IReadOnlyList<string> VerificationMethods { get; init; } = Array.Empty<string>();

    /// <summary>Read-only versus modifying mode.</summary>
    public SapOperationMode OperationMode { get; init; } = SapOperationMode.ReadOnly;

    /// <summary>SAP module.</summary>
    public string SapModule { get; init; } = "";

    /// <summary>Environment applicability metadata.</summary>
    public SapEnvironmentInfo EnvironmentInfo { get; init; } = new();

    /// <summary>Epistemic authority tier.</summary>
    public EpistemicTier EpistemicTier { get; init; } = EpistemicTier.GeneralSapReference;

    /// <summary>Fragment lifecycle state.</summary>
    public FragmentLifecycleState LifecycleState { get; init; } = FragmentLifecycleState.Reviewed;

    /// <summary>Traceable source provenances.</summary>
    public IReadOnlyList<SourceProvenance> SourceProvenances { get; init; } = Array.Empty<SourceProvenance>();

    /// <summary>Searchable keywords.</summary>
    public IReadOnlyList<string> Keywords { get; init; } = Array.Empty<string>();

    /// <summary>Domain synonyms.</summary>
    public IReadOnlyList<string> Synonyms { get; init; } = Array.Empty<string>();

    /// <summary>Classification tags.</summary>
    public IReadOnlyList<string> Tags { get; init; } = Array.Empty<string>();

    /// <summary>Associated business objects.</summary>
    public IReadOnlyList<string> BusinessObjects { get; init; } = Array.Empty<string>();

    /// <summary>Discrepancy types addressed by this transaction or table.</summary>
    public IReadOnlyList<DiscrepancyType> DiscrepancyTypes { get; init; } = Array.Empty<DiscrepancyType>();

    string ISapRankableDocument.Id => Code;
    string ISapRankableDocument.Summary => Purpose;
    IReadOnlyList<string> ISapRankableDocument.TransactionCodes => [Code];
    SapKnowledgeCategory ISapRankableDocument.Category =>
        ObjectKind == SapObjectKind.SapTable
            ? SapKnowledgeCategory.TableAndDataRelationship
            : (ObjectKind == SapObjectKind.ZTransaction
                ? SapKnowledgeCategory.CompanySopAndCustomTransaction
                : SapKnowledgeCategory.TransactionReferenceGuide);
    string ISapRankableDocument.BodyText => string.Join(" ",
        new[] { WhenToUse, NavigationPath }.Concat(OperationalRisks).Concat(VerificationMethods).Concat(RelatedProcedures));
}

/// <summary>
/// Domain validation rules for the SAP Transaction Encyclopedia catalog (R5-B).
/// </summary>
public static class SapTransactionCatalogRules
{
    private static readonly string[] RequiredInitialCodes =
    [
        "ME80FN", "SE16N", "MB51", "VL31N", "RFUI", "MIGO",
        "ME21N", "ME22N", "ME23N", "ZDescarga", "EINA", "MAST", "STPO"
    ];

    private static readonly HashSet<string> KnownSapTableCodes = new(StringComparer.OrdinalIgnoreCase)
    {
        "EINA", "MAST", "STPO", "EKKO", "EKPO", "MKPF", "MSEG", "MARA", "MARD", "MARM", "T161", "T156", "T320"
    };

    private static readonly HashSet<string> KnownZTransactionCodes = new(StringComparer.OrdinalIgnoreCase)
    {
        "ZDESCARGA", "Z-CARGA", "ZDECARGA", "Z FACTURAS", "Z51 -Z52"
    };

    /// <summary>Returns the required initial R5-B transaction and table codes.</summary>
    public static IReadOnlyList<string> GetRequiredInitialCodes() => RequiredInitialCodes;

    /// <summary>
    /// Validates a single <see cref="TransactionEntry"/> against structural SAP object classification rules.
    /// </summary>
    public static IReadOnlyList<string> ValidateEntry(TransactionEntry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);

        var errors = new List<string>();
        string code = entry.Code.Trim();

        if (KnownSapTableCodes.Contains(code) && entry.ObjectKind != SapObjectKind.SapTable)
            errors.Add($"Code '{code}' must be classified as SapObjectKind.SapTable, never '{entry.ObjectKind}'.");

        if (KnownZTransactionCodes.Contains(code) && entry.ObjectKind != SapObjectKind.ZTransaction)
            errors.Add($"Code '{code}' must be classified as SapObjectKind.ZTransaction, not '{entry.ObjectKind}'.");

        if (code.Equals("VKP2", StringComparison.OrdinalIgnoreCase) && !HasVkp2CostWarning(entry))
            errors.Add("VKP2 entry must carry an explicit warning that it shows retail price, never import cost (EKPO-NETPR).");

        return errors;
    }

    /// <summary>
    /// Validates an entire encyclopedia catalog for required R5-B coverage and object-kind invariants.
    /// </summary>
    public static IReadOnlyList<string> ValidateCatalog(IReadOnlyList<TransactionEntry> entries)
    {
        ArgumentNullException.ThrowIfNull(entries);

        var errors = new List<string>();
        var presentCodes = new HashSet<string>(entries.Select(e => e.Code.Trim()), StringComparer.OrdinalIgnoreCase);

        foreach (string required in RequiredInitialCodes)
        {
            if (!presentCodes.Contains(required))
                errors.Add($"Required R5-B catalog entry '{required}' is missing.");
        }

        foreach (TransactionEntry entry in entries)
            errors.AddRange(ValidateEntry(entry));

        return errors;
    }

    private static bool HasVkp2CostWarning(TransactionEntry entry) =>
        entry.OperationalRisks.Any(r =>
            r.Contains("EKPO-NETPR", StringComparison.OrdinalIgnoreCase) ||
            r.Contains("retail", StringComparison.OrdinalIgnoreCase) ||
            r.Contains("venta", StringComparison.OrdinalIgnoreCase));
}
