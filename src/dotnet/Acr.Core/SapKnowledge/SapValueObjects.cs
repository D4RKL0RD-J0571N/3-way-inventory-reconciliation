using Acr.Core.Models;

namespace Acr.Core.SapKnowledge;

/// <summary>
/// SAP environment applicability metadata explicitly separating GUI client version from backend release.
/// </summary>
public sealed record SapEnvironmentInfo
{
    /// <summary>Client GUI version (default <c>"SAP GUI 7.00"</c>); never auto-infers backend release.</summary>
    public string SapGuiVersion { get; init; } = "SAP GUI 7.00";

    /// <summary>Underlying SAP backend release (defaults to <see cref="SapBackendRelease.UnknownPendingVerification"/>).</summary>
    public SapBackendRelease BackendRelease { get; init; } = SapBackendRelease.UnknownPendingVerification;

    /// <summary>Optional corporate company code (e.g., <c>"IM 01"</c>).</summary>
    public string? CompanyCode { get; init; }

    /// <summary>Optional plant or distribution center code (e.g., <c>"C001"</c>).</summary>
    public string? PlantCode { get; init; }

    /// <summary>Additional applicability or environment notes.</summary>
    public string ApplicabilityNotes { get; init; } = "";
}

/// <summary>
/// Traceable provenance reference linking a knowledge fragment or recommendation to its source location.
/// </summary>
public sealed record SourceProvenance
{
    /// <summary>Name or title of the source document.</summary>
    public required string SourceDocument { get; init; }

    /// <summary>Source format (e.g., <c>"PDF"</c>, <c>"DOCX"</c>, <c>"XLSX"</c>, <c>"TXT"</c>, <c>"MD"</c>, <c>"HTML"</c>, <c>"PersonalNote"</c>).</summary>
    public string SourceFormat { get; init; } = "";

    /// <summary>Page number or worksheet name where the fragment appears.</summary>
    public string? PageOrSheet { get; init; }

    /// <summary>Section heading or cell coordinate where the fragment appears.</summary>
    public string? SectionOrCell { get; init; }

    /// <summary>Optional line range inside text/markdown sources.</summary>
    public string? LineRange { get; init; }

    /// <summary>Optional logical path or URI reference.</summary>
    public string? SourceUriOrPath { get; init; }
}

/// <summary>
/// Reference to an SAP database table, relevant fields, and relational usage note.
/// </summary>
public sealed record SapTableFieldReference
{
    /// <summary>SAP table name (e.g., <c>"EINA"</c>, <c>"MAST"</c>, <c>"STPO"</c>, <c>"EKPO"</c>).</summary>
    public required string TableName { get; init; }

    /// <summary>Key or relevant field names within the table.</summary>
    public IReadOnlyList<string> FieldNames { get; init; } = Array.Empty<string>();

    /// <summary>Explanation of how the table and fields relate to receiving operations.</summary>
    public string RelationshipNote { get; init; } = "";
}

/// <summary>
/// Common contract implemented by all SAP knowledge, transaction, and guide documents scored by <see cref="DeterministicRanker"/>.
/// </summary>
public interface ISapRankableDocument
{
    /// <summary>Unique deterministic identifier of the document.</summary>
    string Id { get; }

    /// <summary>Document title.</summary>
    string Title { get; }

    /// <summary>Concise summary text.</summary>
    string Summary { get; }

    /// <summary>Associated SAP transaction or table codes.</summary>
    IReadOnlyList<string> TransactionCodes { get; }

    /// <summary>Associated business object identifiers (e.g., <c>PO_43</c>, <c>MATERIAL</c>, <c>HU</c>).</summary>
    IReadOnlyList<string> BusinessObjects { get; }

    /// <summary>Receiving discrepancy types addressed by this document.</summary>
    IReadOnlyList<DiscrepancyType> DiscrepancyTypes { get; }

    /// <summary>Primary searchable keywords.</summary>
    IReadOnlyList<string> Keywords { get; }

    /// <summary>Domain synonyms.</summary>
    IReadOnlyList<string> Synonyms { get; }

    /// <summary>Classification tags.</summary>
    IReadOnlyList<string> Tags { get; }

    /// <summary>Full searchable body text assembled from steps, notes, and procedures.</summary>
    string BodyText { get; }

    /// <summary>Epistemic authority tier.</summary>
    EpistemicTier EpistemicTier { get; }

    /// <summary>Current governance lifecycle state.</summary>
    FragmentLifecycleState LifecycleState { get; }

    /// <summary>Knowledge library category.</summary>
    SapKnowledgeCategory Category { get; }

    /// <summary>SAP object classification kind.</summary>
    SapObjectKind ObjectKind { get; }
}
