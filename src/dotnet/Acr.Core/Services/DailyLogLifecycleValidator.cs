using System.Text.RegularExpressions;
using Acr.Core.Models;

namespace Acr.Core.Services;

/// <summary>
/// Pure domain validator and formatter for the 6-stage Daily Log lifecycle
/// (<c>Control_Diario_Operaciones_Jostin.xlsx!Maestro Contenedores</c>).
/// </summary>
public static class DailyLogLifecycleValidator
{
    private static readonly Regex ContainerPattern =
        new(@"^IMP[\s\-]*([0-9]+(?:-[0-9]+)?)(?:\s*\([^)]*\))?$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex WhitespaceRunPattern =
        new(@"\s+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly HashSet<string> InvalidTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "", "-", "0", "N/A", "NA", "NONE", "NULL", "NAN", "IMPORT-ACTIVA",
        "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?"
    };

    /// <summary>
    /// Computes the deterministic 7-state master container rollup matching <c>Maestro Contenedores!O{r}</c>.
    /// </summary>
    public static MasterContainerState ComputeMasterState(DailyLogEntry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);

        if (entry.Stage6DispatchClosure == DispatchClosureStage.DispatchClosureReady)
            return MasterContainerState.TotalClosureReceptionAndDispatch;

        if (entry.Stage5ReceptionClosure == ReceptionClosureStage.BalancedZeroDiff)
            return MasterContainerState.ReceptionClosedReady;

        string stage4Text = DailyLogCatalogs.ToExcelString(entry.Stage4QualityControl);
        if (stage4Text.Contains("Revisión", StringComparison.OrdinalIgnoreCase))
            return MasterContainerState.InQualityControlAudit;

        if (entry.Stage5ReceptionClosure == ReceptionClosureStage.PausedWithoutQc)
            return MasterContainerState.PausedWithoutQcBlindSetup;

        if (entry.Stage2SapIntake == SapIntakeStage.ReadyAndNotified)
            return MasterContainerState.EnteredSapAwaitingClosure;

        if (entry.Stage2SapIntake == SapIntakeStage.PendingSapIntake)
            return MasterContainerState.SapDocReadyPendingIntake;

        return MasterContainerState.InBlindCountQueue;
    }

    /// <summary>
    /// Validates a single Daily Log entry against all 6-stage lifecycle guardrails without throwing.
    /// </summary>
    public static DailyLogValidationResult ValidateEntry(DailyLogEntry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);

        var issues = new List<DailyLogValidationIssue>();
        string containerId = ResolveNormalizedOrRaw(entry);

        ValidateIdentityAndPreSapRules(entry, containerId, issues);
        ValidateSapIntakeIdentifiers(entry, containerId, issues);
        ValidateBlindSetupAndQcRules(entry, containerId, issues);
        ValidateClosureAndDispatchRules(entry, containerId, issues);

        return new DailyLogValidationResult { Issues = issues };
    }

    /// <summary>
    /// Validates a lifecycle state transition from <paramref name="previous"/> to <paramref name="next"/>.
    /// </summary>
    public static DailyLogValidationResult ValidateTransition(DailyLogEntry previous, DailyLogEntry next)
    {
        ArgumentNullException.ThrowIfNull(previous);
        ArgumentNullException.ThrowIfNull(next);

        DailyLogValidationResult nextResult = ValidateEntry(next);
        var issues = new List<DailyLogValidationIssue>(nextResult.Issues);
        string containerId = ResolveNormalizedOrRaw(next);

        MasterContainerState prevMaster = ComputeMasterState(previous);
        MasterContainerState nextMaster = ComputeMasterState(next);
        if (nextMaster < prevMaster)
        {
            issues.Add(new DailyLogValidationIssue(
                DailyLogValidationSeverity.Warning,
                "STAGE_REGRESSION_DETECTED",
                $"Master state regressed from '{DailyLogCatalogs.ToExcelString(prevMaster)}' to '{DailyLogCatalogs.ToExcelString(nextMaster)}'.",
                containerId));
        }

        return new DailyLogValidationResult { Issues = issues };
    }

    /// <summary>
    /// Normalizes a raw container identifier to canonical <c>IMP-XXX</c> or <c>IMP-XXX-N</c> form,
    /// returning empty string for bare numbers, Excel errors, or invalid tokens.
    /// </summary>
    public static string NormalizeContainerId(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return "";

        string cleaned = WhitespaceRunPattern.Replace(raw.Replace('\u00A0', ' ').Trim(), " ").ToUpperInvariant();
        if (InvalidTokens.Contains(cleaned) || cleaned.StartsWith('#'))
            return "";

        Match match = ContainerPattern.Match(cleaned);
        if (match.Success)
            return $"IMP-{match.Groups[1].Value}";

        return "";
    }

    /// <summary>
    /// Extracts the short container suffix matching Excel <c>SUBSTITUTE(B{r}, "IMP-", "")</c>.
    /// </summary>
    public static string ExtractShortContainerNumber(string? containerId)
    {
        string normalized = NormalizeContainerId(containerId);
        if (normalized.Length == 0)
            return "";

        return normalized.StartsWith("IMP-", StringComparison.Ordinal)
            ? normalized[4..]
            : normalized;
    }

    /// <summary>
    /// Detects duplicate container entries sharing the same normalized container identifier.
    /// </summary>
    public static IReadOnlyList<ContainerDuplicateConflict> DetectDuplicateContainers(
        IReadOnlyList<DailyLogEntry> entries)
    {
        ArgumentNullException.ThrowIfNull(entries);

        var groups = new Dictionary<string, List<DailyLogEntry>>(StringComparer.Ordinal);
        foreach (DailyLogEntry entry in entries)
        {
            string normalized = ResolveNormalizedOrRaw(entry);
            if (string.IsNullOrWhiteSpace(normalized))
                continue;

            if (!groups.TryGetValue(normalized, out List<DailyLogEntry>? list))
            {
                list = [];
                groups[normalized] = list;
            }
            list.Add(entry);
        }

        return groups
            .Where(kv => kv.Value.Count > 1)
            .OrderBy(kv => kv.Key, StringComparer.Ordinal)
            .Select(kv => new ContainerDuplicateConflict(kv.Key, kv.Value))
            .ToList();
    }

    /// <summary>
    /// Detects distinct containers referencing the same PO 43 or Pedido 41 (non-blocking operational alert).
    /// </summary>
    public static IReadOnlyList<SharedDocumentReference> DetectSharedPoOrPedido(
        IReadOnlyList<DailyLogEntry> entries)
    {
        ArgumentNullException.ThrowIfNull(entries);

        var poMap = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        var pedidoMap = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);

        foreach (DailyLogEntry entry in entries)
        {
            string containerId = ResolveNormalizedOrRaw(entry);
            if (string.IsNullOrWhiteSpace(containerId))
                continue;

            AddDocumentReferences(poMap, entry.Po43Numbers, containerId);
            if (IsPopulatedIdentifier(entry.Pedido41))
                AddDocumentReference(pedidoMap, entry.Pedido41!.Trim(), containerId);
        }

        var alerts = new List<SharedDocumentReference>();
        AppendSharedAlerts(alerts, "PO43", poMap);
        AppendSharedAlerts(alerts, "Pedido41", pedidoMap);
        return alerts;
    }

    /// <summary>
    /// Formats the SAP intake notification outputs matching <c>Maestro Contenedores!R</c> and <c>C8:C9</c>.
    /// </summary>
    public static SapIntakeNotification FormatSapIntakeNotification(DailyLogEntry entry)
    {
        ArgumentNullException.ThrowIfNull(entry);

        if (!IsPopulatedIdentifier(entry.EntregaVl31n))
            return new SapIntakeNotification(false, "-", "", "", Array.Empty<string>());

        string shortNumber = ExtractShortContainerNumber(
            !string.IsNullOrWhiteSpace(entry.NormalizedContainerId) ? entry.NormalizedContainerId : entry.ContainerId);
        if (string.IsNullOrWhiteSpace(shortNumber))
            return new SapIntakeNotification(false, "-", string.Empty, string.Empty, ["INVALID_CONTAINER_ID"]);

        string greetingText = DailyLogCatalogs.ToExcelString(entry.Greeting);
        string colR = $"{greetingText}, Se informa que el ingreso a nivel de sistema del Contenedor {shortNumber}, quedo listo.";
        string subject = $"Ingreso Listo a Sistema - Contenedor {shortNumber}";

        var warnings = new List<string>();
        string body;
        if (entry.Greeting == Saludo.None)
        {
            warnings.Add("SALUDO_NOT_SELECTED");
            body = $"Se informa que el ingreso a nivel de sistema del Contenedor {shortNumber}, quedo listo.";
        }
        else
        {
            body = $"{greetingText}\nSe informa que el ingreso a nivel de sistema del Contenedor {shortNumber}, quedo listo.";
        }

        return new SapIntakeNotification(true, colR, subject, body, warnings);
    }

    private static string ResolveNormalizedOrRaw(DailyLogEntry entry)
    {
        if (!string.IsNullOrWhiteSpace(entry.NormalizedContainerId))
            return NormalizeContainerId(entry.NormalizedContainerId);
        return NormalizeContainerId(entry.ContainerId);
    }

    private static void ValidateIdentityAndPreSapRules(
        DailyLogEntry entry,
        string containerId,
        List<DailyLogValidationIssue> issues)
    {
        if (string.IsNullOrWhiteSpace(containerId))
        {
            issues.Add(new DailyLogValidationIssue(
                DailyLogValidationSeverity.Error,
                "CONTAINER_ID_REQUIRED",
                "Normalized container ID is required and cannot be blank or bare numeric.",
                entry.ContainerId));
        }

        MasterContainerState master = ComputeMasterState(entry);
        if (master is MasterContainerState.InBlindCountQueue or MasterContainerState.SapDocReadyPendingIntake &&
            entry.Stage5ReceptionClosure != ReceptionClosureStage.PendingClosure)
        {
            issues.Add(new DailyLogValidationIssue(
                DailyLogValidationSeverity.Error,
                "PRE_SAP_STAGE5_MUST_BE_PENDING_CLOSURE",
                "Pre-SAP or blind-queue containers must keep Stage 5 in 'Pendiente Cierre'.",
                containerId));
        }
    }

    private static void ValidateSapIntakeIdentifiers(
        DailyLogEntry entry,
        string containerId,
        List<DailyLogValidationIssue> issues)
    {
        if (entry.Stage2SapIntake != SapIntakeStage.ReadyAndNotified)
            return;

        bool hasEntrega = IsPopulatedIdentifier(entry.EntregaVl31n);
        bool hasEwm = IsPopulatedIdentifier(entry.EwmDeliveryDoc);
        bool hasHu = IsPopulatedIdentifier(entry.HandlingUnit);

        if (!hasEntrega || !hasEwm || !hasHu)
        {
            issues.Add(new DailyLogValidationIssue(
                DailyLogValidationSeverity.Warning,
                "SAP_INTAKE_MISSING_TRACKING_IDS",
                "Stage 2 is 'Listo y Notificado' without all tracking IDs (Entrega 18..., EWM 414xx, HU 18-digit).",
                containerId));
        }

        if ((hasEntrega && !IsValidPrefixAndDigits(entry.EntregaVl31n!, "18", 9)) ||
            (hasEwm && !IsValidPrefixAndDigits(entry.EwmDeliveryDoc!, "414", 5)) ||
            (hasHu && !IsExactAsciiDigits(entry.HandlingUnit!, 18)))
        {
            issues.Add(new DailyLogValidationIssue(
                DailyLogValidationSeverity.Warning,
                "SAP_IDENTIFIER_FORMAT_UNEXPECTED",
                "One or more SAP tracking identifiers do not match expected format (18xxxxxxx, 414xx, 18-digit HU).",
                containerId));
        }
    }

    private static void ValidateBlindSetupAndQcRules(
        DailyLogEntry entry,
        string containerId,
        List<DailyLogValidationIssue> issues)
    {
        bool isBlindReady = entry.Stage3BlindSetup is BlindSetupStage.ReadySupplierDocOnly or BlindSetupStage.ReadyComplete;
        if (isBlindReady && entry.Stage4QualityControl == QualityControlStage.NoQcWaitingDock)
        {
            if (entry.Stage5ReceptionClosure != ReceptionClosureStage.PausedWithoutQc ||
                entry.Stage6DispatchClosure != DispatchClosureStage.BlockedByReception)
            {
                issues.Add(new DailyLogValidationIssue(
                    DailyLogValidationSeverity.Error,
                    "BLIND_SETUP_REQUIRES_PAUSA_SIN_CC",
                    "Blind setup ready without QC requires Stage 5 = 'Pausa sin Control Calidad' and Stage 6 = 'Bloqueado por Recepción'.",
                    containerId));
            }
        }

        if (entry.Stage4QualityControl == QualityControlStage.PendingFinalQcReview &&
            entry.Stage5ReceptionClosure is ReceptionClosureStage.PendingClosure or ReceptionClosureStage.PausedWithoutQc)
        {
            issues.Add(new DailyLogValidationIssue(
                DailyLogValidationSeverity.Warning,
                "QC_REVIEW_STAGE5_NOT_UPDATED",
                "Stage 4 is 'Pendiente Revisión Final CC' while Stage 5 still indicates 'Pendiente Cierre' or 'Pausa sin Control Calidad'.",
                containerId));
        }
    }

    private static void ValidateClosureAndDispatchRules(
        DailyLogEntry entry,
        string containerId,
        List<DailyLogValidationIssue> issues)
    {
        if (entry.Stage5ReceptionClosure == ReceptionClosureStage.BalancedZeroDiff &&
            (entry.Stage1Customs != CustomsDuaStage.DuaAndDocsReady ||
             entry.Stage4QualityControl != QualityControlStage.QcApprovedZDescarga))
        {
            issues.Add(new DailyLogValidationIssue(
                DailyLogValidationSeverity.Error,
                "RECEPTION_CLOSED_REQUIRES_DUA_AND_QC_APPROVED",
                "Stage 5 'Cuadrado (0 Diferencias)' requires Stage 1 = 'DUA y Docs Listos' and Stage 4 = 'CC Aprobado (ZDescarga)'.",
                containerId));
        }

        if (entry.Stage6DispatchClosure != DispatchClosureStage.BlockedByReception &&
            entry.Stage5ReceptionClosure is not (ReceptionClosureStage.BalancedZeroDiff or ReceptionClosureStage.PartialPendingPartB))
        {
            issues.Add(new DailyLogValidationIssue(
                DailyLogValidationSeverity.Error,
                "DISPATCH_UNLOCKED_BEFORE_RECEPTION_READY",
                "Stage 6 cannot leave 'Bloqueado por Recepción' until Stage 5 reaches 'Cuadrado (0 Diferencias)' or 'Parcial / Pendiente Parte B'.",
                containerId));
        }
    }

    private static bool IsPopulatedIdentifier(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;
        string trimmed = value.Trim();
        return trimmed != "-" && !trimmed.StartsWith('#');
    }

    private static bool IsValidPrefixAndDigits(string value, string prefix, int expectedLength)
    {
        string trimmed = value.Trim();
        return trimmed.Length == expectedLength &&
               trimmed.StartsWith(prefix, StringComparison.Ordinal) &&
               IsExactAsciiDigits(trimmed, expectedLength);
    }

    private static bool IsExactAsciiDigits(string value, int expectedLength)
    {
        string trimmed = value.Trim();
        if (trimmed.Length != expectedLength)
            return false;
        foreach (char ch in trimmed)
        {
            if (!char.IsAsciiDigit(ch))
                return false;
        }
        return true;
    }

    private static void AddDocumentReferences(
        Dictionary<string, HashSet<string>> map,
        IReadOnlyList<string> docs,
        string containerId)
    {
        foreach (string doc in docs)
        {
            if (IsPopulatedIdentifier(doc))
                AddDocumentReference(map, doc.Trim(), containerId);
        }
    }

    private static void AddDocumentReference(
        Dictionary<string, HashSet<string>> map,
        string docNumber,
        string containerId)
    {
        if (!map.TryGetValue(docNumber, out HashSet<string>? containers))
        {
            containers = new HashSet<string>(StringComparer.Ordinal);
            map[docNumber] = containers;
        }
        containers.Add(containerId);
    }

    private static void AppendSharedAlerts(
        List<SharedDocumentReference> alerts,
        string kind,
        Dictionary<string, HashSet<string>> map)
    {
        foreach ((string docNumber, HashSet<string> containers) in map.OrderBy(kv => kv.Key, StringComparer.Ordinal))
        {
            if (containers.Count > 1)
            {
                alerts.Add(new SharedDocumentReference(
                    kind,
                    docNumber,
                    containers.OrderBy(c => c, StringComparer.Ordinal).ToList()));
            }
        }
    }
}
