namespace Acr.Core.Models;

/// <summary>
/// Classification categories for dock receiving discrepancies.
/// </summary>
public enum DiscrepancyType
{
    /// <summary>Physical units, cartons, and condition match declared expectations.</summary>
    ExactMatch = 0,

    /// <summary>Physically received units are fewer than declared units.</summary>
    Shortage = 1,

    /// <summary>Physically received units exceed declared units.</summary>
    Overage = 2,

    /// <summary>Shortage or overage paired deterministically with an equal-magnitude opposite discrepancy.</summary>
    RevueltoCandidate = 3,

    /// <summary>Units match exactly, but physical carton count differs from declared carton count.</summary>
    CartonMismatchOnly = 4,

    /// <summary>Units match declared count, but physically damaged units were recorded.</summary>
    DamageOnly = 5,

    /// <summary>Physical merchandise received without any declared quantity.</summary>
    UnexpectedItem = 6,

    /// <summary>Declared merchandise completely absent from physical dock reception.</summary>
    MissingItem = 7
}

/// <summary>
/// Evidence confidence tier for a candidate revuelto (material swap) pair.
/// </summary>
public enum RevueltoConfidence
{
    /// <summary>Exact unit delta corroborated by matching vendor material or same category plus opposing carton delta.</summary>
    Confirmed = 0,

    /// <summary>Exact unit delta corroborated by either same category or opposing carton delta.</summary>
    Probable = 1,

    /// <summary>Exact unit delta only without corroborating vendor, category, or carton swap evidence.</summary>
    Possible = 2
}

/// <summary>
/// Declared merchandise item from a Packing List, Purchase Order, or closure line.
/// </summary>
public sealed record DeclaredItem
{
    /// <summary>Primary SAP material code (SKU key).</summary>
    public required string MaterialCode { get; init; }

    /// <summary>Supplier style or vendor material reference.</summary>
    public string? VendorMaterial { get; init; }

    /// <summary>Short material description.</summary>
    public string? Description { get; init; }

    /// <summary>Merchandise hierarchy or category code.</summary>
    public string? Category { get; init; }

    /// <summary>Purchase Order 43 number for multi-PO line attribution.</summary>
    public string? OrderPo43 { get; init; }

    /// <summary>Optional source worksheet row index.</summary>
    public int? SourceRow { get; init; }

    /// <summary>Declared or invoiced units.</summary>
    public int DeclaredUnits { get; init; }

    /// <summary>Declared carton count.</summary>
    public int DeclaredCartons { get; init; }

    /// <summary>Unit of measure (<c>"PRK"</c> or <c>"UN"</c>).</summary>
    public string? Uom { get; init; } = "UN";

    /// <summary>Exact decimal monetary unit cost.</summary>
    public decimal UnitCost { get; init; }
}

/// <summary>
/// Physically received and inspected merchandise item from dock QC.
/// </summary>
public sealed record PhysicalItem
{
    /// <summary>Primary SAP material code (SKU key).</summary>
    public required string MaterialCode { get; init; }

    /// <summary>Supplier style or vendor material reference.</summary>
    public string? VendorMaterial { get; init; }

    /// <summary>Short material description.</summary>
    public string? Description { get; init; }

    /// <summary>Merchandise hierarchy or category code.</summary>
    public string? Category { get; init; }

    /// <summary>Optional Purchase Order 43 attribution.</summary>
    public string? OrderPo43 { get; init; }

    /// <summary>Optional source QC worksheet row index.</summary>
    public int? SourceRow { get; init; }

    /// <summary>Calculated physical units received.</summary>
    public int PhysicalUnits { get; init; }

    /// <summary>Calculated physical cartons received.</summary>
    public int PhysicalCartons { get; init; }

    /// <summary>Physically damaged unit count.</summary>
    public int DamagedUnits { get; init; }

    /// <summary>Unit of measure (<c>"PRK"</c> or <c>"UN"</c>).</summary>
    public string? Uom { get; init; } = "UN";

    /// <summary>Diagnostic flags propagated from <see cref="Services.QuantityEngine"/>.</summary>
    public IReadOnlyList<string> Flags { get; init; } = Array.Empty<string>();
}

/// <summary>
/// Individual material-level comparison result and discrepancy classification.
/// </summary>
public sealed record DiscrepancyItem
{
    /// <summary>Primary SAP material code.</summary>
    public string MaterialCode { get; init; } = "";

    /// <summary>Resolved vendor material reference.</summary>
    public string VendorMaterial { get; init; } = "";

    /// <summary>Resolved material description.</summary>
    public string Description { get; init; } = "";

    /// <summary>Resolved category code.</summary>
    public string Category { get; init; } = "";

    /// <summary>Single PO or comma-joined distinct PO numbers when multi-PO lines were aggregated.</summary>
    public string? OrderPo43 { get; init; }

    /// <summary>Aggregated declared units.</summary>
    public int DeclaredUnits { get; init; }

    /// <summary>Aggregated physical units.</summary>
    public int PhysicalUnits { get; init; }

    /// <summary>Signed unit difference (<c>PhysicalUnits - DeclaredUnits</c>).</summary>
    public int DifferenceUnits { get; init; }

    /// <summary>Aggregated declared cartons.</summary>
    public int DeclaredCartons { get; init; }

    /// <summary>Aggregated physical cartons.</summary>
    public int PhysicalCartons { get; init; }

    /// <summary>Signed carton difference (<c>PhysicalCartons - DeclaredCartons</c>).</summary>
    public int DifferenceCartons { get; init; }

    /// <summary>Aggregated damaged units.</summary>
    public int DamagedUnits { get; init; }

    /// <summary>Classified discrepancy category.</summary>
    public DiscrepancyType DiscrepancyType { get; init; } = DiscrepancyType.ExactMatch;

    /// <summary>Linked partner material code when classified as <see cref="DiscrepancyType.RevueltoCandidate"/>.</summary>
    public string? RevueltoPartner { get; init; }

    /// <summary>Evidence confidence level when linked as a revuelto candidate.</summary>
    public RevueltoConfidence? RevueltoConfidence { get; init; }

    /// <summary>Human-readable evidence rationale for the revuelto pairing.</summary>
    public string? RevueltoEvidence { get; init; }

    /// <summary>True when multiple declared PO lines existed for this material code.</summary>
    public bool HasMultiplePoLines { get; init; }

    /// <summary>Number of declared PO lines aggregated for this material code.</summary>
    public int PoLineCount { get; init; } = 1;

    /// <summary>Per-PO audit breakdown lines when multiple PO lines were aggregated.</summary>
    public IReadOnlyList<string> PoBreakdown { get; init; } = Array.Empty<string>();

    /// <summary>Aggregated physical and comparison flags.</summary>
    public IReadOnlyList<string> Flags { get; init; } = Array.Empty<string>();

    /// <summary>Concise Spanish explanation of the discrepancy.</summary>
    public string Explanation { get; init; } = "";

    /// <summary>Texto en español del tipo de discrepancia basado en los términos de CIERRE RECEP.</summary>
    public string DiscrepancyTypeDescription => DiscrepancyType switch
    {
        DiscrepancyType.ExactMatch => "CUADRADO (0 DIF)",
        DiscrepancyType.Shortage => "FALTANTE",
        DiscrepancyType.Overage => "SOBRANTE",
        DiscrepancyType.RevueltoCandidate => "REVUELTO",
        DiscrepancyType.CartonMismatchOnly => "DIF. BULTOS",
        DiscrepancyType.DamageOnly => "DAÑADO",
        DiscrepancyType.UnexpectedItem => "SOBRANTE (NO FACTURADO)",
        DiscrepancyType.MissingItem => "FALTANTE TOTAL",
        _ => DiscrepancyType.ToString()
    };

    /// <summary>Color de fondo semántico para el badge del tipo de discrepancia (Hex RGBA).</summary>
    public string BadgeBackgroundColor => DiscrepancyType switch
    {
        DiscrepancyType.ExactMatch => "#DCFCE7", // Verde suave
        DiscrepancyType.Shortage => "#FEE2E2",   // Rojo suave
        DiscrepancyType.MissingItem => "#FEE2E2",// Rojo suave
        DiscrepancyType.Overage => "#FEF3C7",    // Ámbar suave
        DiscrepancyType.UnexpectedItem => "#FEF3C7", // Ámbar suave
        DiscrepancyType.RevueltoCandidate => "#F3E8FF", // Púrpura suave
        DiscrepancyType.DamageOnly => "#FFE4E6", // Rosa/Rojo suave
        DiscrepancyType.CartonMismatchOnly => "#E0F2FE", // Azul cielo suave
        _ => "#F1F5F9"
    };

    /// <summary>Color de texto semántico para el badge del tipo de discrepancia.</summary>
    public string BadgeForegroundColor => DiscrepancyType switch
    {
        DiscrepancyType.ExactMatch => "#15803D", // Verde fuerte
        DiscrepancyType.Shortage => "#B91C1C",   // Rojo fuerte
        DiscrepancyType.MissingItem => "#B91C1C",// Rojo fuerte
        DiscrepancyType.Overage => "#B45309",    // Ámbar fuerte
        DiscrepancyType.UnexpectedItem => "#B45309", // Ámbar fuerte
        DiscrepancyType.RevueltoCandidate => "#7E22CE", // Púrpura fuerte
        DiscrepancyType.DamageOnly => "#BE123C", // Rosa/Rojo fuerte
        DiscrepancyType.CartonMismatchOnly => "#0369A1", // Azul fuerte
        _ => "#334155"
    };

    /// <summary>Color del borde semántico de 1px para el badge.</summary>
    public string BadgeBorderColor => DiscrepancyType switch
    {
        DiscrepancyType.ExactMatch => "#86EFAC",
        DiscrepancyType.Shortage => "#FCA5A5",
        DiscrepancyType.MissingItem => "#FCA5A5",
        DiscrepancyType.Overage => "#FCD34D",
        DiscrepancyType.UnexpectedItem => "#FCD34D",
        DiscrepancyType.RevueltoCandidate => "#D8B4FE",
        DiscrepancyType.DamageOnly => "#FDA4AF",
        DiscrepancyType.CartonMismatchOnly => "#BAE6FD",
        _ => "#CBD5E1"
    };

    /// <summary>Color de texto para la celda de Diferencia de Unidades (Verde si 0, Rojo si negativo, Ámbar si positivo).</summary>
    public string DeltaUnitsColor => DifferenceUnits switch
    {
        0 => "#16A34A",
        < 0 => "#DC2626",
        > 0 => "#D97706"
    };

    /// <summary>Color de texto para la celda de Diferencia de Bultos.</summary>
    public string DeltaCartonsColor => DifferenceCartons switch
    {
        0 => "#16A34A",
        < 0 => "#DC2626",
        > 0 => "#D97706"
    };

    /// <summary>Color de texto para unidades dañadas.</summary>
    public string DamagedUnitsColor => DamagedUnits > 0 ? "#E11D48" : "#94A3B8";

    /// <summary>Color de fondo sutil para resaltar toda la fila cuando hay discrepancia crítica.</summary>
    public string RowBackgroundColor => DiscrepancyType switch
    {
        DiscrepancyType.ExactMatch => "#FFFFFF",
        DiscrepancyType.Shortage or DiscrepancyType.MissingItem => "#FFF5F5",
        DiscrepancyType.Overage or DiscrepancyType.UnexpectedItem => "#FFFBEB",
        DiscrepancyType.RevueltoCandidate => "#FAF5FF",
        DiscrepancyType.DamageOnly => "#FFF1F2",
        DiscrepancyType.CartonMismatchOnly => "#F0F9FF",
        _ => "#FFFFFF"
    };
}

/// <summary>
/// Auditable record of a deterministic revuelto candidate pair linking a shortage and an overage.
/// </summary>
public sealed record RevueltoMatch
{
    /// <summary>Material code presenting a unit shortage.</summary>
    public string ShortageMaterialCode { get; init; } = "";

    /// <summary>Material code presenting an equal-magnitude unit overage.</summary>
    public string OverageMaterialCode { get; init; } = "";

    /// <summary>Positive shortage unit magnitude.</summary>
    public int ShortageUnits { get; init; }

    /// <summary>Positive overage unit magnitude.</summary>
    public int OverageUnits { get; init; }

    /// <summary>Deterministic confidence tier for the candidate swap.</summary>
    public RevueltoConfidence Confidence { get; init; } = RevueltoConfidence.Possible;

    /// <summary>Human-readable evidence rationale describing corroborating signals.</summary>
    public string Evidence { get; init; } = "";

    /// <summary>List of matched corroborating attribute tags.</summary>
    public IReadOnlyList<string> MatchedAttributes { get; init; } = Array.Empty<string>();
}

/// <summary>
/// High-level summary of receiving comparison across all materials enforcing the Cardinal Discrepancy Rule.
/// </summary>
public sealed record ComparisonSummary
{
    /// <summary>Total declared units across all items.</summary>
    public int TotalDeclaredUnits { get; init; }

    /// <summary>Total physical units across all items.</summary>
    public int TotalPhysicalUnits { get; init; }

    /// <summary>Net unit difference (<c>TotalPhysicalUnits - TotalDeclaredUnits</c>).</summary>
    public int NetUnitDifference { get; init; }

    /// <summary>Total declared cartons across all items.</summary>
    public int TotalDeclaredCartons { get; init; }

    /// <summary>Total physical cartons across all items.</summary>
    public int TotalPhysicalCartons { get; init; }

    /// <summary>Net carton difference (<c>TotalPhysicalCartons - TotalDeclaredCartons</c>).</summary>
    public int NetCartonDifference { get; init; }

    /// <summary>Absolute sum of all unit shortages (never offset by overages).</summary>
    public int TotalShortageUnits { get; init; }

    /// <summary>Sum of all unit overages (never offset by shortages).</summary>
    public int TotalOverageUnits { get; init; }

    /// <summary>Total damaged units across all items.</summary>
    public int TotalDamagedUnits { get; init; }

    /// <summary>Count of materials classified as <see cref="DiscrepancyType.ExactMatch"/>.</summary>
    public int CleanMaterialsCount { get; init; }

    /// <summary>Count of materials presenting any discrepancy.</summary>
    public int DiscrepantMaterialsCount { get; init; }

    /// <summary>Count of paired revuelto candidates detected.</summary>
    public int RevueltoPairsCount { get; init; }

    /// <summary>Material-level discrepancy comparison items ordered by material code.</summary>
    public IReadOnlyList<DiscrepancyItem> Items { get; init; } = Array.Empty<DiscrepancyItem>();

    /// <summary>Auditable revuelto candidate pairs with confidence and evidence.</summary>
    public IReadOnlyList<RevueltoMatch> RevueltoMatches { get; init; } = Array.Empty<RevueltoMatch>();

    /// <summary>Material codes that had multiple declared PO lines aggregated.</summary>
    public IReadOnlyList<string> MultiPoMaterials { get; init; } = Array.Empty<string>();
}
