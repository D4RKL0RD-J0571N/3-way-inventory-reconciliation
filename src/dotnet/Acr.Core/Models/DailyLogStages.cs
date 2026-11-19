using System.Globalization;
using System.Text;

namespace Acr.Core.Models;

/// <summary>
/// Stage 1: Customs, DUA, and Radar status (<c>Maestro Contenedores!I</c>, <c>Catalogos!A2:A4</c>).
/// </summary>
public enum CustomsDuaStage
{
    /// <summary>"DUA y Docs Listos" — Customs declaration and documents ready.</summary>
    DuaAndDocsReady = 0,

    /// <summary>"En Radar / Tránsito" — Container in radar or transit.</summary>
    InRadarOrTransit = 1,

    /// <summary>"Pendiente Docs" — Pending supplier or customs documents.</summary>
    PendingDocs = 2
}

/// <summary>
/// Stage 2: SAP / EWM Intake status (<c>Maestro Contenedores!J</c>, <c>Catalogos!B2:B5</c>).
/// </summary>
public enum SapIntakeStage
{
    /// <summary>"Listo y Notificado" — SAP/EWM intake completed and notified.</summary>
    ReadyAndNotified = 0,

    /// <summary>"Pendiente Ingreso SAP" — Documents ready, pending SAP intake.</summary>
    PendingSapIntake = 1,

    /// <summary>"En Proceso EWM/RFUI" — EWM/RFUI intake in progress.</summary>
    InProgressEwmRfui = 2,

    /// <summary>"Sin Doc. 18... Aún" — Inbound delivery 18... not yet created.</summary>
    NoDoc18Yet = 3
}

/// <summary>
/// Stage 3: Blind Setup status (<c>Maestro Contenedores!K</c>, <c>Catalogos!C2:C5</c>).
/// </summary>
public enum BlindSetupStage
{
    /// <summary>"Listo (Solo Doc. Proveedor)" — Blind setup ready using supplier docs only.</summary>
    ReadySupplierDocOnly = 0,

    /// <summary>"Listo (Completo)" — Blind setup complete.</summary>
    ReadyComplete = 1,

    /// <summary>"En Cola A Ciegas" — Queued for blind setup.</summary>
    InBlindSetupQueue = 2,

    /// <summary>"Pendiente" — Blind setup pending.</summary>
    Pending = 3
}

/// <summary>
/// Stage 4: Dock Quality Control status (<c>Maestro Contenedores!L</c>, <c>Catalogos!D2:D4</c>).
/// </summary>
public enum QualityControlStage
{
    /// <summary>"CC Aprobado (ZDescarga)" — Quality Control approved via ZDescarga.</summary>
    QcApprovedZDescarga = 0,

    /// <summary>"Pendiente Revisión Final CC" — Pending final Quality Control audit review.</summary>
    PendingFinalQcReview = 1,

    /// <summary>"Sin CC (Espera Muelle)" — No QC file yet, awaiting dock unload.</summary>
    NoQcWaitingDock = 2
}

/// <summary>
/// Stage 5: 3-Way Reception Closure Mov 101 status (<c>Maestro Contenedores!M</c>, <c>Catalogos!E2:E7</c>).
/// </summary>
public enum ReceptionClosureStage
{
    /// <summary>"Cuadrado (0 Diferencias)" — Balanced 3-way reception closure with zero differences.</summary>
    BalancedZeroDiff = 0,

    /// <summary>"Pausa sin Control Calidad" — Mounted blind, paused awaiting QC.</summary>
    PausedWithoutQc = 1,

    /// <summary>"Pendiente Revisión CC" — Pending QC discrepancy verification.</summary>
    PendingQcReview = 2,

    /// <summary>"Parcial / Pendiente Parte B" — Partial reception pending Part B container.</summary>
    PartialPendingPartB = 3,

    /// <summary>"Discrepancia / NCR (40...)" — Discrepancy requiring credit note / NCR 40....</summary>
    DiscrepancyNcr40 = 4,

    /// <summary>"Pendiente Cierre" — Reception closure pending.</summary>
    PendingClosure = 5
}

/// <summary>
/// Stage 6: Store Dispatch Closure Mov 641 status (<c>Maestro Contenedores!N</c>, <c>Catalogos!F2:F5</c>).
/// </summary>
public enum DispatchClosureStage
{
    /// <summary>"Cierre Despacho Listo" — Store dispatch closure completed.</summary>
    DispatchClosureReady = 0,

    /// <summary>"Pre-Cableado 80% (Listo Mov 641)" — Pre-wired 80%, ready for Mov 641.</summary>
    PreWired80PercentReadyMov641 = 1,

    /// <summary>"En Distribución Tiendas" — Actively in store distribution.</summary>
    InStoreDistribution = 2,

    /// <summary>"Bloqueado por Recepción" — Blocked until reception closure reaches Cuadrado or Partial Part B.</summary>
    BlockedByReception = 3
}

/// <summary>
/// Greeting selector for SAP intake notification (<c>Maestro Contenedores!Q</c>, <c>Catalogos!G2:G4</c>).
/// </summary>
public enum Saludo
{
    /// <summary>"Buenos días"</summary>
    BuenosDias = 0,

    /// <summary>"Buenas tardes"</summary>
    BuenasTardes = 1,

    /// <summary>"-" (No greeting selected)</summary>
    None = 2
}

/// <summary>
/// Deferred Part B container status (<c>Catalogos!H2:H4</c>).
/// </summary>
public enum ParteBStatus
{
    /// <summary>"Pendiente Arribo" — Pending arrival.</summary>
    PendingArrival = 0,

    /// <summary>"En Muelle / Descarga" — At dock or unloading.</summary>
    AtDockUnloading = 1,

    /// <summary>"Conciliado 0 Dif" — Reconciled with zero differences.</summary>
    ReconciledZeroDiff = 2
}

/// <summary>
/// Deterministic 7-state master portfolio rollup computed in <c>Maestro Contenedores!O</c>.
/// </summary>
public enum MasterContainerState
{
    /// <summary>"1. En Cola Conteo a Ciegas"</summary>
    InBlindCountQueue = 1,

    /// <summary>"2. Doc SAP Listo (Pend. Ingreso)"</summary>
    SapDocReadyPendingIntake = 2,

    /// <summary>"3. Ingresado SAP (Espera Cierre)"</summary>
    EnteredSapAwaitingClosure = 3,

    /// <summary>"4. Pausa sin CC (Montado a Ciegas)"</summary>
    PausedWithoutQcBlindSetup = 4,

    /// <summary>"5. En Auditoría Control Calidad"</summary>
    InQualityControlAudit = 5,

    /// <summary>"6. Recepción Cerrada (Listo Recep.)"</summary>
    ReceptionClosedReady = 6,

    /// <summary>"7. Cierre Total (Recep + Desp)"</summary>
    TotalClosureReceptionAndDispatch = 7
}

/// <summary>
/// Bidirectional mappings between Daily Log domain enums and exact <c>Catalogos!A1:H7</c> / <c>Col O</c> Excel strings.
/// </summary>
public static class DailyLogCatalogs
{
    private static readonly (CustomsDuaStage Value, string Text)[] Stage1Map =
    [
        (CustomsDuaStage.DuaAndDocsReady, "DUA y Docs Listos"),
        (CustomsDuaStage.InRadarOrTransit, "En Radar / Tránsito"),
        (CustomsDuaStage.PendingDocs, "Pendiente Docs")
    ];

    private static readonly (SapIntakeStage Value, string Text)[] Stage2Map =
    [
        (SapIntakeStage.ReadyAndNotified, "Listo y Notificado"),
        (SapIntakeStage.PendingSapIntake, "Pendiente Ingreso SAP"),
        (SapIntakeStage.InProgressEwmRfui, "En Proceso EWM/RFUI"),
        (SapIntakeStage.NoDoc18Yet, "Sin Doc. 18... Aún")
    ];

    private static readonly (BlindSetupStage Value, string Text)[] Stage3Map =
    [
        (BlindSetupStage.ReadySupplierDocOnly, "Listo (Solo Doc. Proveedor)"),
        (BlindSetupStage.ReadyComplete, "Listo (Completo)"),
        (BlindSetupStage.InBlindSetupQueue, "En Cola A Ciegas"),
        (BlindSetupStage.Pending, "Pendiente")
    ];

    private static readonly (QualityControlStage Value, string Text)[] Stage4Map =
    [
        (QualityControlStage.QcApprovedZDescarga, "CC Aprobado (ZDescarga)"),
        (QualityControlStage.PendingFinalQcReview, "Pendiente Revisión Final CC"),
        (QualityControlStage.NoQcWaitingDock, "Sin CC (Espera Muelle)")
    ];

    private static readonly (ReceptionClosureStage Value, string Text)[] Stage5Map =
    [
        (ReceptionClosureStage.BalancedZeroDiff, "Cuadrado (0 Diferencias)"),
        (ReceptionClosureStage.PausedWithoutQc, "Pausa sin Control Calidad"),
        (ReceptionClosureStage.PendingQcReview, "Pendiente Revisión CC"),
        (ReceptionClosureStage.PartialPendingPartB, "Parcial / Pendiente Parte B"),
        (ReceptionClosureStage.DiscrepancyNcr40, "Discrepancia / NCR (40...)"),
        (ReceptionClosureStage.PendingClosure, "Pendiente Cierre")
    ];

    private static readonly (DispatchClosureStage Value, string Text)[] Stage6Map =
    [
        (DispatchClosureStage.DispatchClosureReady, "Cierre Despacho Listo"),
        (DispatchClosureStage.PreWired80PercentReadyMov641, "Pre-Cableado 80% (Listo Mov 641)"),
        (DispatchClosureStage.InStoreDistribution, "En Distribución Tiendas"),
        (DispatchClosureStage.BlockedByReception, "Bloqueado por Recepción")
    ];

    private static readonly (Saludo Value, string Text)[] SaludoMap =
    [
        (Saludo.BuenosDias, "Buenos días"),
        (Saludo.BuenasTardes, "Buenas tardes"),
        (Saludo.None, "-")
    ];

    private static readonly (ParteBStatus Value, string Text)[] ParteBMap =
    [
        (ParteBStatus.PendingArrival, "Pendiente Arribo"),
        (ParteBStatus.AtDockUnloading, "En Muelle / Descarga"),
        (ParteBStatus.ReconciledZeroDiff, "Conciliado 0 Dif")
    ];

    private static readonly (MasterContainerState Value, string Text)[] MasterStateMap =
    [
        (MasterContainerState.InBlindCountQueue, "1. En Cola Conteo a Ciegas"),
        (MasterContainerState.SapDocReadyPendingIntake, "2. Doc SAP Listo (Pend. Ingreso)"),
        (MasterContainerState.EnteredSapAwaitingClosure, "3. Ingresado SAP (Espera Cierre)"),
        (MasterContainerState.PausedWithoutQcBlindSetup, "4. Pausa sin CC (Montado a Ciegas)"),
        (MasterContainerState.InQualityControlAudit, "5. En Auditoría Control Calidad"),
        (MasterContainerState.ReceptionClosedReady, "6. Recepción Cerrada (Listo Recep.)"),
        (MasterContainerState.TotalClosureReceptionAndDispatch, "7. Cierre Total (Recep + Desp)")
    ];

    /// <summary>Converts a <see cref="CustomsDuaStage"/> value to its exact Excel catalog text.</summary>
    public static string ToExcelString(CustomsDuaStage stage) => FindText(Stage1Map, stage);

    /// <summary>Converts a <see cref="SapIntakeStage"/> value to its exact Excel catalog text.</summary>
    public static string ToExcelString(SapIntakeStage stage) => FindText(Stage2Map, stage);

    /// <summary>Converts a <see cref="BlindSetupStage"/> value to its exact Excel catalog text.</summary>
    public static string ToExcelString(BlindSetupStage stage) => FindText(Stage3Map, stage);

    /// <summary>Converts a <see cref="QualityControlStage"/> value to its exact Excel catalog text.</summary>
    public static string ToExcelString(QualityControlStage stage) => FindText(Stage4Map, stage);

    /// <summary>Converts a <see cref="ReceptionClosureStage"/> value to its exact Excel catalog text.</summary>
    public static string ToExcelString(ReceptionClosureStage stage) => FindText(Stage5Map, stage);

    /// <summary>Converts a <see cref="DispatchClosureStage"/> value to its exact Excel catalog text.</summary>
    public static string ToExcelString(DispatchClosureStage stage) => FindText(Stage6Map, stage);

    /// <summary>Converts a <see cref="Saludo"/> value to its exact Excel catalog text.</summary>
    public static string ToExcelString(Saludo saludo) => FindText(SaludoMap, saludo);

    /// <summary>Converts a <see cref="ParteBStatus"/> value to its exact Excel catalog text.</summary>
    public static string ToExcelString(ParteBStatus status) => FindText(ParteBMap, status);

    /// <summary>Converts a <see cref="MasterContainerState"/> value to its exact Excel Col O text.</summary>
    public static string ToExcelString(MasterContainerState state) => FindText(MasterStateMap, state);

    /// <summary>Attempts to parse a Stage 1 catalog string (case- and accent-tolerant).</summary>
    public static bool TryParseCustomsDua(string? raw, out CustomsDuaStage value) => TryFindValue(Stage1Map, raw, out value);

    /// <summary>Attempts to parse a Stage 2 catalog string (case- and accent-tolerant).</summary>
    public static bool TryParseSapIntake(string? raw, out SapIntakeStage value) => TryFindValue(Stage2Map, raw, out value);

    /// <summary>Attempts to parse a Stage 3 catalog string (case- and accent-tolerant).</summary>
    public static bool TryParseBlindSetup(string? raw, out BlindSetupStage value) => TryFindValue(Stage3Map, raw, out value);

    /// <summary>Attempts to parse a Stage 4 catalog string (case- and accent-tolerant).</summary>
    public static bool TryParseQualityControl(string? raw, out QualityControlStage value) => TryFindValue(Stage4Map, raw, out value);

    /// <summary>Attempts to parse a Stage 5 catalog string (case- and accent-tolerant).</summary>
    public static bool TryParseReceptionClosure(string? raw, out ReceptionClosureStage value) => TryFindValue(Stage5Map, raw, out value);

    /// <summary>Attempts to parse a Stage 6 catalog string (case- and accent-tolerant).</summary>
    public static bool TryParseDispatchClosure(string? raw, out DispatchClosureStage value) => TryFindValue(Stage6Map, raw, out value);

    /// <summary>Attempts to parse a Saludo catalog string (case- and accent-tolerant).</summary>
    public static bool TryParseSaludo(string? raw, out Saludo value) => TryFindValue(SaludoMap, raw, out value);

    /// <summary>Attempts to parse a Parte B status catalog string (case- and accent-tolerant).</summary>
    public static bool TryParseParteBStatus(string? raw, out ParteBStatus value) => TryFindValue(ParteBMap, raw, out value);

    /// <summary>Attempts to parse a Col O master container state string (case- and accent-tolerant).</summary>
    public static bool TryParseMasterState(string? raw, out MasterContainerState value) => TryFindValue(MasterStateMap, raw, out value);

    private static string FindText<TEnum>((TEnum Value, string Text)[] map, TEnum value) where TEnum : struct
    {
        foreach (var item in map)
        {
            if (EqualityComparer<TEnum>.Default.Equals(item.Value, value))
                return item.Text;
        }
        throw new ArgumentOutOfRangeException(nameof(value), value, "Unmapped Daily Log enum value.");
    }

    private static bool TryFindValue<TEnum>((TEnum Value, string Text)[] map, string? raw, out TEnum value) where TEnum : struct
    {
        value = default;
        if (string.IsNullOrWhiteSpace(raw))
            return false;
        string trimmed = raw.Trim();
        foreach (var item in map)
        {
            if (string.Equals(trimmed, item.Text, StringComparison.OrdinalIgnoreCase))
            {
                value = item.Value;
                return true;
            }
        }
        string folded = FoldDiacritics(trimmed);
        foreach (var item in map)
        {
            if (string.Equals(folded, FoldDiacritics(item.Text), StringComparison.OrdinalIgnoreCase))
            {
                value = item.Value;
                return true;
            }
        }
        return false;
    }

    private static string FoldDiacritics(string input)
    {
        string decomposed = input.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(decomposed.Length);
        foreach (char ch in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                sb.Append(ch);
        }
        return sb.ToString().Normalize(NormalizationForm.FormC);
    }
}
