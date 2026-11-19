namespace Acr.Core.Models;

/// <summary>
/// Defines the runtime operational data safety mode governing read/write access.
/// </summary>
public enum OperationalSafetyMode
{
    /// <summary>Production data directory and operational Daily Log are strictly read-only.</summary>
    Development = 0,

    /// <summary>Production data and Daily Log are strictly read-only; all writes target isolated temporary copies.</summary>
    Testing = 1,

    /// <summary>Source QC and template workbooks are read-only; the Daily Log is writable only via explicit user-triggered save or sync actions.</summary>
    Production = 2
}

/// <summary>
/// Logical target category evaluated by <see cref="Services.OperationalSafetyPolicy"/> without filesystem I/O.
/// </summary>
public enum OperationalTargetKind
{
    /// <summary>Source QC or master closure workbook in the production data directory.</summary>
    ProductionSourceWorkbook = 0,

    /// <summary>The operational Daily Log workbook (Control_Diario_Operaciones_Jostin.xlsx).</summary>
    OperationalDailyLog = 1,

    /// <summary>Generated closure export workbook destination.</summary>
    ClosureExportOutput = 2,

    /// <summary>An isolated temporary copy directory used for testing or staging.</summary>
    IsolatedTempCopy = 3,

    /// <summary>A local JSON processing report or operation history artifact.</summary>
    LocalReportArtifact = 4
}

/// <summary>
/// Requested operation intent evaluated against <see cref="OperationalSafetyMode"/>.
/// </summary>
public enum OperationalActionKind
{
    /// <summary>Read-only inspection of a file or workbook.</summary>
    Read = 0,

    /// <summary>Automated or background write attempt without explicit user confirmation.</summary>
    WriteAutomatedOrBackground = 1,

    /// <summary>Explicit user-triggered save or historical sync commit action.</summary>
    WriteExplicitUserSaveOrSync = 2,

    /// <summary>Write to a newly selected export destination path.</summary>
    WriteToNewExportPath = 3
}

/// <summary>
/// Immutable result of evaluating a safety policy rule.
/// </summary>
/// <param name="IsAllowed">Whether the requested operation is permitted on the target.</param>
/// <param name="MustRedirectToTemp">Whether writes must be redirected to an isolated temporary directory.</param>
/// <param name="ReasonCode">Deterministic machine-readable policy decision code.</param>
/// <param name="Message">Human-readable explanation of the policy decision.</param>
public sealed record SafetyPolicyDecision(
    bool IsAllowed,
    bool MustRedirectToTemp,
    string ReasonCode,
    string Message);
