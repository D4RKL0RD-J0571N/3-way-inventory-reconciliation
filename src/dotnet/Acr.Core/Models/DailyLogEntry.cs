namespace Acr.Core.Models;

/// <summary>
/// Domain representation of a row in <c>Control_Diario_Operaciones_Jostin.xlsx!Maestro Contenedores</c> (<c>A</c>–<c>R</c>).
/// </summary>
public sealed record DailyLogEntry
{
    /// <summary>Column A (<c>N°</c>): Optional sequential row number.</summary>
    public int? SequenceNumber { get; init; }

    /// <summary>Column B (<c>Contenedor (Ubic. EWM)</c>): Raw container identifier.</summary>
    public string ContainerId { get; init; } = "";

    /// <summary>Normalized canonical container identifier (e.g., <c>IMP-202</c>, <c>IMP-201-1</c>).</summary>
    public string NormalizedContainerId { get; init; } = "";

    /// <summary>Column C (<c>Fecha Act.</c>): Last update date.</summary>
    public DateOnly? UpdatedDate { get; init; }

    /// <summary>Column D (<c>PO Orden Compra (430000...)</c>): Parsed Purchase Order 43 numbers.</summary>
    public IReadOnlyList<string> Po43Numbers { get; init; } = Array.Empty<string>();

    /// <summary>Column E (<c>Pedido Nacionaliz. (410000...)</c>): Nationalization order 41 number.</summary>
    public string? Pedido41 { get; init; }

    /// <summary>Column F (<c>Entrega VL31N (1800...)</c>): SAP inbound delivery number.</summary>
    public string? EntregaVl31n { get; init; }

    /// <summary>Column G (<c>Doc. Entrega EWM (Tarea Alm.)</c>): EWM warehouse delivery task document.</summary>
    public string? EwmDeliveryDoc { get; init; }

    /// <summary>Column H (<c>Unidad Manipulación HU (18 díg)</c>): 18-digit Handling Unit identifier.</summary>
    public string? HandlingUnit { get; init; }

    /// <summary>Column I: Stage 1 (Customs and DUA).</summary>
    public CustomsDuaStage Stage1Customs { get; init; } = CustomsDuaStage.PendingDocs;

    /// <summary>Column J: Stage 2 (SAP / EWM Intake).</summary>
    public SapIntakeStage Stage2SapIntake { get; init; } = SapIntakeStage.NoDoc18Yet;

    /// <summary>Column K: Stage 3 (Blind Setup).</summary>
    public BlindSetupStage Stage3BlindSetup { get; init; } = BlindSetupStage.Pending;

    /// <summary>Column L: Stage 4 (Dock Quality Control).</summary>
    public QualityControlStage Stage4QualityControl { get; init; } = QualityControlStage.NoQcWaitingDock;

    /// <summary>Column M: Stage 5 (3-Way Reception Closure Mov 101).</summary>
    public ReceptionClosureStage Stage5ReceptionClosure { get; init; } = ReceptionClosureStage.PendingClosure;

    /// <summary>Column N: Stage 6 (Store Dispatch Closure Mov 641).</summary>
    public DispatchClosureStage Stage6DispatchClosure { get; init; } = DispatchClosureStage.BlockedByReception;

    /// <summary>Column O: Computed 7-state master container rollup.</summary>
    public MasterContainerState MasterState { get; init; } = MasterContainerState.InBlindCountQueue;

    /// <summary>Column P (<c>Observaciones y Contexto Operativo</c>): Operational notes.</summary>
    public string Observations { get; init; } = "";

    /// <summary>Column Q (<c>Saludo</c>): Greeting used for SAP intake notifications.</summary>
    public Saludo Greeting { get; init; } = Saludo.None;
}

/// <summary>
/// Severity level for Daily Log lifecycle and guardrail validation issues.
/// </summary>
public enum DailyLogValidationSeverity
{
    /// <summary>Non-blocking operational or historical anomaly warning.</summary>
    Warning = 0,

    /// <summary>Blocking structural or lifecycle transition violation.</summary>
    Error = 1
}

/// <summary>
/// Represents a single validation issue detected on a <see cref="DailyLogEntry"/>.
/// </summary>
/// <param name="Severity">Warning or Error severity.</param>
/// <param name="Code">Deterministic machine-readable issue code.</param>
/// <param name="Message">Descriptive explanation of the issue.</param>
/// <param name="ContainerId">Associated normalized or raw container identifier.</param>
public sealed record DailyLogValidationIssue(
    DailyLogValidationSeverity Severity,
    string Code,
    string Message,
    string ContainerId);

/// <summary>
/// Structured, non-throwing validation outcome for a <see cref="DailyLogEntry"/>.
/// </summary>
public sealed record DailyLogValidationResult
{
    /// <summary>True when zero blocking errors were detected.</summary>
    public bool IsValid => Errors.Count == 0;

    /// <summary>All errors and warnings produced during validation.</summary>
    public IReadOnlyList<DailyLogValidationIssue> Issues { get; init; } = Array.Empty<DailyLogValidationIssue>();

    /// <summary>Blocking validation errors.</summary>
    public IReadOnlyList<DailyLogValidationIssue> Errors =>
        Issues.Where(i => i.Severity == DailyLogValidationSeverity.Error).ToList();

    /// <summary>Non-blocking operational warnings.</summary>
    public IReadOnlyList<DailyLogValidationIssue> Warnings =>
        Issues.Where(i => i.Severity == DailyLogValidationSeverity.Warning).ToList();
}

/// <summary>
/// Represents a duplicate container conflict across multiple Daily Log rows.
/// </summary>
/// <param name="NormalizedContainerId">Duplicate normalized container ID.</param>
/// <param name="Entries">Conflicting Daily Log rows sharing this container ID.</param>
public sealed record ContainerDuplicateConflict(
    string NormalizedContainerId,
    IReadOnlyList<DailyLogEntry> Entries);

/// <summary>
/// Non-blocking alert when distinct containers legitimately reference the same PO 43 or Pedido 41.
/// </summary>
/// <param name="DocumentKind">Type of shared document (<c>"PO43"</c> or <c>"Pedido41"</c>).</param>
/// <param name="DocumentNumber">Shared document number.</param>
/// <param name="NormalizedContainerIds">Distinct containers referencing this document.</param>
public sealed record SharedDocumentReference(
    string DocumentKind,
    string DocumentNumber,
    IReadOnlyList<string> NormalizedContainerIds);

/// <summary>
/// Formatted SAP intake notification outputs matching <c>Col R</c> and <c>Comunicados y Guia SAP!C8:C9</c>.
/// </summary>
/// <param name="IsReady">True when <c>EntregaVl31n</c> is populated (not blank or <c>"-"</c>).</param>
/// <param name="ColRValue">Exact <c>Maestro Contenedores!R</c> formula output text.</param>
/// <param name="EmailSubject">Exact <c>C8</c> email subject line.</param>
/// <param name="ClipboardBody">Exact <c>C9</c> multiline clipboard message body.</param>
/// <param name="Warnings">Non-blocking formatting warnings (e.g., missing greeting).</param>
public sealed record SapIntakeNotification(
    bool IsReady,
    string ColRValue,
    string EmailSubject,
    string ClipboardBody,
    IReadOnlyList<string> Warnings);
