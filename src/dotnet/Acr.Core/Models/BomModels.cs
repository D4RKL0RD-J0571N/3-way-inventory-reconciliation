namespace Acr.Core.Models;

/// <summary>
/// Child SKU line item resulting from BOM expansion of a prepack parent material.
/// </summary>
public sealed record ChildSkuItem
{
    /// <summary>Child SAP material code (size-specific SKU).</summary>
    public required string MaterialCode { get; init; }

    /// <summary>Child material description or size specification.</summary>
    public string Description { get; init; } = "";

    /// <summary>Curve ratio or quantity units per prepack pack.</summary>
    public int Ratio { get; init; } = 1;

    /// <summary>Theoretical expected units based on invoice / PO.</summary>
    public int TheoreticalUnits { get; init; }

    /// <summary>Physically received units allocated proportionally (largest-remainder).</summary>
    public int ReceivedUnits { get; init; }

    /// <summary>Signed difference (<c>ReceivedUnits - TheoreticalUnits</c>).</summary>
    public int Difference => ReceivedUnits - TheoreticalUnits;

    /// <summary>Monetary unit cost in USD.</summary>
    public decimal UnitCost { get; init; }
}

/// <summary>
/// Single BOM definition entry linking a parent prepack material to a child size SKU and ratio.
/// </summary>
public sealed record BomDefinitionEntry
{
    /// <summary>Child size SKU material code.</summary>
    public required string ChildMaterialCode { get; init; }

    /// <summary>Units of this child SKU per prepack pack (ratio).</summary>
    public int Ratio { get; init; } = 1;
}
