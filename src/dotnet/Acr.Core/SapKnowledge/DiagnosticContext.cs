using Acr.Core.Models;

namespace Acr.Core.SapKnowledge;

/// <summary>
/// Context passed from Reconciliation, Daily Log, or Historical Sync into guided SAP diagnostics.
/// Enforces the no-invention invariant: absent identifiers remain <c>null</c> and are never synthesized.
/// </summary>
public sealed record DiagnosticContext
{
    private readonly string? _containerId;
    private readonly string? _po43;
    private readonly string? _pedido41;
    private readonly string? _entrega18;
    private readonly string? _ewmDeliveryDoc;
    private readonly string? _hu;
    private readonly string? _materialCode;
    private readonly string? _vendorMaterialCode;
    private readonly string? _supplierCode;
    private readonly string? _revueltoPartner;
    private readonly string? _problemDescription;

    /// <summary>Normalized or raw container identifier (null if absent).</summary>
    public string? ContainerId { get => _containerId; init => _containerId = NullIfWhiteSpace(value); }

    /// <summary>Purchase Order 43 number (null if absent).</summary>
    public string? Po43 { get => _po43; init => _po43 = NullIfWhiteSpace(value); }

    /// <summary>Nationalization Pedido 41 number (null if absent).</summary>
    public string? Pedido41 { get => _pedido41; init => _pedido41 = NullIfWhiteSpace(value); }

    /// <summary>SAP Inbound Delivery 1800... number (null if absent).</summary>
    public string? Entrega18 { get => _entrega18; init => _entrega18 = NullIfWhiteSpace(value); }

    /// <summary>EWM delivery task document (null if absent).</summary>
    public string? EwmDeliveryDoc { get => _ewmDeliveryDoc; init => _ewmDeliveryDoc = NullIfWhiteSpace(value); }

    /// <summary>18-digit Handling Unit (null if absent).</summary>
    public string? Hu { get => _hu; init => _hu = NullIfWhiteSpace(value); }

    /// <summary>SAP Material code (null if absent).</summary>
    public string? MaterialCode { get => _materialCode; init => _materialCode = NullIfWhiteSpace(value); }

    /// <summary>Vendor material/style code (null if absent).</summary>
    public string? VendorMaterialCode { get => _vendorMaterialCode; init => _vendorMaterialCode = NullIfWhiteSpace(value); }

    /// <summary>SAP Vendor/Supplier code (null if absent).</summary>
    public string? SupplierCode { get => _supplierCode; init => _supplierCode = NullIfWhiteSpace(value); }

    /// <summary>Detected discrepancy classification if launched from reconciliation.</summary>
    public DiscrepancyType? DiscrepancyType { get; init; }

    /// <summary>Shortage magnitude in units if applicable.</summary>
    public int? ShortageUnits { get; init; }

    /// <summary>Overage magnitude in units if applicable.</summary>
    public int? OverageUnits { get; init; }

    /// <summary>Damaged units count if applicable.</summary>
    public int? DamagedUnits { get; init; }

    /// <summary>Linked revuelto partner material code (null if absent).</summary>
    public string? RevueltoPartner { get => _revueltoPartner; init => _revueltoPartner = NullIfWhiteSpace(value); }

    /// <summary>Operator problem description (null if absent).</summary>
    public string? ProblemDescription { get => _problemDescription; init => _problemDescription = NullIfWhiteSpace(value); }

    /// <summary>Originating workflow view or module.</summary>
    public DiagnosticContextOrigin SourceOrigin { get; init; } = DiagnosticContextOrigin.ManualQuery;

    /// <summary>
    /// Returns only verified non-null identifiers present in this context (never invents missing values).
    /// </summary>
    public IReadOnlyDictionary<string, string> GetProvidedIdentifiers()
    {
        var ids = new Dictionary<string, string>(StringComparer.Ordinal);
        AddIfPresent(ids, "ContainerId", ContainerId);
        AddIfPresent(ids, "Po43", Po43);
        AddIfPresent(ids, "Pedido41", Pedido41);
        AddIfPresent(ids, "Entrega18", Entrega18);
        AddIfPresent(ids, "EwmDeliveryDoc", EwmDeliveryDoc);
        AddIfPresent(ids, "Hu", Hu);
        AddIfPresent(ids, "MaterialCode", MaterialCode);
        AddIfPresent(ids, "VendorMaterialCode", VendorMaterialCode);
        AddIfPresent(ids, "SupplierCode", SupplierCode);
        AddIfPresent(ids, "RevueltoPartner", RevueltoPartner);
        return ids;
    }

    /// <summary>
    /// Returns SAP/receiving business object types evidenced strictly by non-null identifiers in this context.
    /// </summary>
    public IReadOnlyList<string> GetInferredBusinessObjects()
    {
        var objects = new List<string>();
        if (Po43 is not null) objects.Add("PO_43");
        if (Pedido41 is not null) objects.Add("PEDIDO_41");
        if (Entrega18 is not null) objects.Add("ENTREGA_18");
        if (EwmDeliveryDoc is not null) objects.Add("EWM_DOC");
        if (Hu is not null) objects.Add("HU");
        if (MaterialCode is not null) objects.Add("MATERIAL");
        if (ContainerId is not null) objects.Add("CONTENEDOR");
        if (SupplierCode is not null) objects.Add("PROVEEDOR");
        return objects;
    }

    private static string? NullIfWhiteSpace(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value.Trim();

    private static void AddIfPresent(Dictionary<string, string> dict, string key, string? value)
    {
        if (value is not null)
            dict[key] = value;
    }
}
