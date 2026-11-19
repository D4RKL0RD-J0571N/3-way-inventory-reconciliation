using Acr.Core.Models;
using Acr.Infrastructure.Safety;

namespace Acr.Infrastructure.Excel;

/// <summary>
/// Catalog option lists extracted from <c>Control_Diario_Operaciones_Jostin.xlsx!Catalogos!A1:H7</c>.
/// </summary>
public sealed record DailyLogCatalogSnapshot
{
    /// <summary>Stage 1 (<c>1. Aduana y DUA</c>) catalog strings.</summary>
    public IReadOnlyList<string> Stage1CustomsOptions { get; init; } = Array.Empty<string>();

    /// <summary>Stage 2 (<c>2. Ingreso SAP / EWM</c>) catalog strings.</summary>
    public IReadOnlyList<string> Stage2SapIntakeOptions { get; init; } = Array.Empty<string>();

    /// <summary>Stage 3 (<c>3. Montaje a Ciegas</c>) catalog strings.</summary>
    public IReadOnlyList<string> Stage3BlindSetupOptions { get; init; } = Array.Empty<string>();

    /// <summary>Stage 4 (<c>4. Control Calidad</c>) catalog strings.</summary>
    public IReadOnlyList<string> Stage4QualityControlOptions { get; init; } = Array.Empty<string>();

    /// <summary>Stage 5 (<c>5. Cierre Recepción (101)</c>) catalog strings.</summary>
    public IReadOnlyList<string> Stage5ReceptionClosureOptions { get; init; } = Array.Empty<string>();

    /// <summary>Stage 6 (<c>6. Cierre Despacho (641)</c>) catalog strings.</summary>
    public IReadOnlyList<string> Stage6DispatchClosureOptions { get; init; } = Array.Empty<string>();

    /// <summary>Greeting (<c>Saludos</c>) catalog strings.</summary>
    public IReadOnlyList<string> SaludoOptions { get; init; } = Array.Empty<string>();

    /// <summary>Deferred Part B (<c>Estado Parte B</c>) catalog strings.</summary>
    public IReadOnlyList<string> ParteBStatusOptions { get; init; } = Array.Empty<string>();
}

/// <summary>
/// Full snapshot read from a Daily Log workbook (<c>Maestro Contenedores</c> and <c>Catalogos</c>).
/// </summary>
/// <param name="FilePath">Canonicalized path of the Daily Log workbook.</param>
/// <param name="Entries">Parsed <see cref="DailyLogEntry"/> rows in worksheet order.</param>
/// <param name="RowIndices">1-based worksheet row index corresponding 1:1 with <paramref name="Entries"/>.</param>
/// <param name="RowIndexByNormalizedKey">First 1-based row index keyed by normalized container ID.</param>
/// <param name="Catalogs">Catalog lists read from <c>Catalogos!A1:H7</c>.</param>
/// <param name="TableStartRow">Starting row of <c>tblMaestroContenedores</c> (header row).</param>
/// <param name="TableEndRow">Ending row of <c>tblMaestroContenedores</c>.</param>
/// <param name="ReadWarnings">Non-blocking read warnings for historical anomalies.</param>
public sealed record DailyLogWorkbookSnapshot(
    string FilePath,
    IReadOnlyList<DailyLogEntry> Entries,
    IReadOnlyList<int> RowIndices,
    IReadOnlyDictionary<string, int> RowIndexByNormalizedKey,
    DailyLogCatalogSnapshot Catalogs,
    int TableStartRow,
    int TableEndRow,
    IReadOnlyList<string> ReadWarnings);

/// <summary>
/// Result of loading the Daily Log together with duplicate container and shared PO/Pedido diagnostics.
/// </summary>
/// <param name="Snapshot">Workbook snapshot containing entries, row indices, and catalogs.</param>
/// <param name="Duplicates">Detected duplicate container groups.</param>
/// <param name="SharedDocuments">Non-blocking shared PO 43 or Pedido 41 alerts across distinct containers.</param>
public sealed record DailyLogLoadResult(
    DailyLogWorkbookSnapshot Snapshot,
    IReadOnlyList<ContainerDuplicateConflict> Duplicates,
    IReadOnlyList<SharedDocumentReference> SharedDocuments);

/// <summary>
/// Outcome of validating and transactionally saving a Daily Log entry.
/// </summary>
public sealed record DailyLogSaveResult
{
    /// <summary>True when validation passed and the atomic workbook transaction committed.</summary>
    public bool Succeeded { get; init; }

    /// <summary>1-based Excel row index written in <c>Maestro Contenedores</c>, if committed.</summary>
    public int? WrittenRowIndex { get; init; }

    /// <summary>Normalized <see cref="DailyLogEntry"/> evaluated and persisted.</summary>
    public DailyLogEntry? NormalizedEntry { get; init; }

    /// <summary>Lifecycle validation result from <see cref="Acr.Core.Services.DailyLogLifecycleValidator"/>.</summary>
    public required DailyLogValidationResult Validation { get; init; }

    /// <summary>Atomic file transaction result when disk commit occurred.</summary>
    public AtomicTransactionResult? TransactionResult { get; init; }

    /// <summary>Duplicate container conflicts detected across the workbook.</summary>
    public IReadOnlyList<ContainerDuplicateConflict> Duplicates { get; init; } = Array.Empty<ContainerDuplicateConflict>();

    /// <summary>Non-blocking shared PO 43 / Pedido 41 alerts across distinct containers.</summary>
    public IReadOnlyList<SharedDocumentReference> SharedDocuments { get; init; } = Array.Empty<SharedDocumentReference>();

    /// <summary>Formatted SAP intake notification for the saved entry.</summary>
    public SapIntakeNotification? Notification { get; init; }

    /// <summary>Error explanation when <see cref="Succeeded"/> is false.</summary>
    public string? ErrorMessage { get; init; }
}
