using Acr.Core.Models;
using Acr.Core.Services;
using Acr.Infrastructure.Safety;

namespace Acr.Infrastructure.Excel;

/// <summary>
/// Domain-enforced service coordinating <see cref="DailyLogLifecycleValidator"/>,
/// duplicate container detection, shared PO/Pedido alerts, SAP intake notifications,
/// and transactional persistence via <see cref="DailyLogRepository"/>.
/// </summary>
public static class DailyLogService
{
    /// <summary>
    /// Loads the Daily Log workbook along with duplicate container conflicts and shared PO/Pedido alerts.
    /// </summary>
    public static DailyLogLoadResult Load(string path, OperationalSafetyGuard guard)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(guard);

        DailyLogWorkbookSnapshot snapshot = DailyLogRepository.ReadWorkbook(path, guard);
        IReadOnlyList<ContainerDuplicateConflict> duplicates =
            DailyLogLifecycleValidator.DetectDuplicateContainers(snapshot.Entries);
        IReadOnlyList<SharedDocumentReference> sharedDocs =
            DailyLogLifecycleValidator.DetectSharedPoOrPedido(snapshot.Entries);

        return new DailyLogLoadResult(snapshot, duplicates, sharedDocs);
    }

    /// <summary>
    /// Validates and appends a new container entry, blocking duplicate container creation before touching disk.
    /// </summary>
    public static DailyLogSaveResult AddEntry(
        string path,
        DailyLogEntry entry,
        OperationalSafetyGuard guard,
        bool allowWarnings = true)
    {
        ArgumentNullException.ThrowIfNull(entry);
        ArgumentNullException.ThrowIfNull(guard);

        DailyLogEntry normalized = NormalizeEntry(entry);
        DailyLogLoadResult current = Load(path, guard);
        List<DailyLogEntry> matchingExisting = FindMatchingEntries(current.Snapshot.Entries, normalized.NormalizedContainerId);

        if (matchingExisting.Count > 0)
            return BuildDuplicateFailure(normalized, matchingExisting, current.SharedDocuments);

        return ValidateAndPersist(path, normalized, previous: null, targetRowIndex: null, allowWarnings, current, guard);
    }

    /// <summary>
    /// Validates and either updates an existing container row or appends a new container row atomically.
    /// Aborts before touching disk on blocking lifecycle errors or ambiguous duplicate targets.
    /// </summary>
    public static DailyLogSaveResult AddOrUpdateEntry(
        string path,
        DailyLogEntry entry,
        OperationalSafetyGuard guard,
        bool allowWarnings = true,
        int? targetRowIndex = null)
    {
        ArgumentNullException.ThrowIfNull(entry);
        ArgumentNullException.ThrowIfNull(guard);

        DailyLogEntry normalized = NormalizeEntry(entry);
        DailyLogLoadResult current = Load(path, guard);
        (DailyLogEntry? previous, int? resolvedRow, DailyLogSaveResult? ambiguityError) =
            ResolveTargetRow(current.Snapshot, normalized, targetRowIndex, current.SharedDocuments);

        if (ambiguityError is not null)
            return ambiguityError;

        return ValidateAndPersist(path, normalized, previous, resolvedRow, allowWarnings, current, guard);
    }

    /// <summary>
    /// Imports eligible new historical closure candidates (<see cref="HistoricalMatchKind.NewRecord"/> without
    /// human review blockers) into the Daily Log inside a single atomic transaction.
    /// </summary>
    public static IReadOnlyList<DailyLogSaveResult> ImportHistoricalCandidates(
        string path,
        IReadOnlyList<HistoricalMatchResult> matchResults,
        OperationalSafetyGuard guard)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(matchResults);
        ArgumentNullException.ThrowIfNull(guard);

        DailyLogLoadResult current = Load(path, guard);
        var knownContainers = new HashSet<string>(
            current.Snapshot.Entries.Select(e => e.NormalizedContainerId).Where(id => id.Length > 0),
            StringComparer.OrdinalIgnoreCase);

        var toCommit = new List<(DailyLogEntry Entry, int? TargetRowIndex)>();
        var results = new List<DailyLogSaveResult>();

        foreach (HistoricalMatchResult match in matchResults)
        {
            if (match.MatchKind != HistoricalMatchKind.NewRecord || match.RequiresHumanReview)
                continue;

            DailyLogEntry candidate = BuildEntryFromHistorical(match.Candidate.PrimaryRecord);
            DailyLogValidationResult validation = DailyLogLifecycleValidator.ValidateEntry(candidate);
            if (!validation.IsValid || !knownContainers.Add(candidate.NormalizedContainerId))
            {
                results.Add(BuildSkippedHistoricalResult(candidate, validation));
                continue;
            }

            toCommit.Add((candidate, null));
        }

        if (toCommit.Count > 0)
            CommitHistoricalBatch(path, toCommit, guard, results);

        return results;
    }

    private static DailyLogSaveResult ValidateAndPersist(
        string path,
        DailyLogEntry normalized,
        DailyLogEntry? previous,
        int? targetRowIndex,
        bool allowWarnings,
        DailyLogLoadResult current,
        OperationalSafetyGuard guard)
    {
        DailyLogValidationResult validation = previous is not null
            ? DailyLogLifecycleValidator.ValidateTransition(previous, normalized)
            : DailyLogLifecycleValidator.ValidateEntry(normalized);

        SapIntakeNotification notification = DailyLogLifecycleValidator.FormatSapIntakeNotification(normalized);
        if (!validation.IsValid || (!allowWarnings && validation.Warnings.Count > 0))
        {
            return new DailyLogSaveResult
            {
                Succeeded = false,
                NormalizedEntry = normalized,
                Validation = validation,
                Duplicates = current.Duplicates,
                SharedDocuments = current.SharedDocuments,
                Notification = notification,
                ErrorMessage = validation.Errors.FirstOrDefault()?.Message ?? "Validation warnings blocked save."
            };
        }

        (AtomicTransactionResult tx, int writtenRow) =
            DailyLogRepository.SaveEntry(path, normalized, guard, targetRowIndex);
        DailyLogLoadResult updated = Load(path, guard);

        return new DailyLogSaveResult
        {
            Succeeded = true,
            WrittenRowIndex = writtenRow,
            NormalizedEntry = normalized,
            Validation = validation,
            TransactionResult = tx,
            Duplicates = updated.Duplicates,
            SharedDocuments = updated.SharedDocuments,
            Notification = notification
        };
    }

    private static (DailyLogEntry? Previous, int? ResolvedRow, DailyLogSaveResult? Error) ResolveTargetRow(
        DailyLogWorkbookSnapshot snapshot,
        DailyLogEntry normalized,
        int? targetRowIndex,
        IReadOnlyList<SharedDocumentReference> sharedDocs)
    {
        if (targetRowIndex.HasValue)
        {
            int idx = IndexOfRow(snapshot.RowIndices, targetRowIndex.Value);
            DailyLogEntry? prev = idx >= 0 ? snapshot.Entries[idx] : null;
            return (prev, targetRowIndex.Value, null);
        }

        List<int> matchingIndices = FindMatchingRowPositions(snapshot.Entries, normalized.NormalizedContainerId);
        if (matchingIndices.Count == 1)
        {
            int pos = matchingIndices[0];
            return (snapshot.Entries[pos], snapshot.RowIndices[pos], null);
        }

        if (matchingIndices.Count > 1)
        {
            List<DailyLogEntry> conflicting = matchingIndices.Select(i => snapshot.Entries[i]).ToList();
            return (null, null, BuildDuplicateFailure(normalized, conflicting, sharedDocs));
        }

        return (null, null, null);
    }

    private static int IndexOfRow(IReadOnlyList<int> rowIndices, int targetRow)
    {
        for (int i = 0; i < rowIndices.Count; i++)
        {
            if (rowIndices[i] == targetRow)
                return i;
        }
        return -1;
    }

    private static List<DailyLogEntry> FindMatchingEntries(
        IReadOnlyList<DailyLogEntry> entries,
        string normalizedContainerId) =>
        string.IsNullOrWhiteSpace(normalizedContainerId)
            ? []
            : entries.Where(e => string.Equals(e.NormalizedContainerId, normalizedContainerId, StringComparison.OrdinalIgnoreCase)).ToList();

    private static List<int> FindMatchingRowPositions(
        IReadOnlyList<DailyLogEntry> entries,
        string normalizedContainerId)
    {
        var positions = new List<int>();
        if (string.IsNullOrWhiteSpace(normalizedContainerId))
            return positions;

        for (int i = 0; i < entries.Count; i++)
        {
            if (string.Equals(entries[i].NormalizedContainerId, normalizedContainerId, StringComparison.OrdinalIgnoreCase))
                positions.Add(i);
        }
        return positions;
    }

    private static DailyLogSaveResult BuildDuplicateFailure(
        DailyLogEntry normalized,
        IReadOnlyList<DailyLogEntry> existingMatches,
        IReadOnlyList<SharedDocumentReference> sharedDocs)
    {
        var issue = new DailyLogValidationIssue(
            DailyLogValidationSeverity.Error,
            "DUPLICATE_CONTAINER_DETECTED",
            $"Container '{normalized.NormalizedContainerId}' already exists in Daily Log ({existingMatches.Count} row(s)).",
            normalized.NormalizedContainerId);

        return new DailyLogSaveResult
        {
            Succeeded = false,
            NormalizedEntry = normalized,
            Validation = new DailyLogValidationResult { Issues = [issue] },
            Duplicates = [new ContainerDuplicateConflict(normalized.NormalizedContainerId, existingMatches)],
            SharedDocuments = sharedDocs,
            Notification = DailyLogLifecycleValidator.FormatSapIntakeNotification(normalized),
            ErrorMessage = issue.Message
        };
    }

    private static DailyLogEntry NormalizeEntry(DailyLogEntry entry)
    {
        string normalizedId = DailyLogLifecycleValidator.NormalizeContainerId(entry.ContainerId);
        string containerText = normalizedId.Length > 0 ? normalizedId : entry.ContainerId.Trim();
        DailyLogEntry withId = entry with
        {
            ContainerId = containerText,
            NormalizedContainerId = normalizedId
        };
        return withId with { MasterState = DailyLogLifecycleValidator.ComputeMasterState(withId) };
    }

    private static DailyLogEntry BuildEntryFromHistorical(HistoricalSyncRecord record)
    {
        var draft = new DailyLogEntry
        {
            ContainerId = record.Identity.NormalizedContainerId,
            NormalizedContainerId = record.Identity.NormalizedContainerId,
            UpdatedDate = record.Identity.RevisionDate ?? DateOnly.FromDateTime(DateTime.Today),
            Po43Numbers = record.Identity.Po43Numbers,
            Pedido41 = record.Identity.Pedido41,
            Stage1Customs = CustomsDuaStage.DuaAndDocsReady,
            Stage2SapIntake = SapIntakeStage.ReadyAndNotified,
            Stage3BlindSetup = BlindSetupStage.ReadyComplete,
            Stage4QualityControl = QualityControlStage.QcApprovedZDescarga,
            Stage5ReceptionClosure = ReceptionClosureStage.BalancedZeroDiff,
            Stage6DispatchClosure = DispatchClosureStage.BlockedByReception,
            Observations = $"Importado desde cierre histórico: {record.Provenance.SourceFileName}"
        };
        return draft with { MasterState = DailyLogLifecycleValidator.ComputeMasterState(draft) };
    }

    private static DailyLogSaveResult BuildSkippedHistoricalResult(
        DailyLogEntry candidate,
        DailyLogValidationResult validation) =>
        new()
        {
            Succeeded = false,
            NormalizedEntry = candidate,
            Validation = validation,
            ErrorMessage = validation.Errors.FirstOrDefault()?.Message ?? "Duplicate historical candidate skipped."
        };

    private static void CommitHistoricalBatch(
        string path,
        List<(DailyLogEntry Entry, int? TargetRowIndex)> toCommit,
        OperationalSafetyGuard guard,
        List<DailyLogSaveResult> results)
    {
        (AtomicTransactionResult tx, IReadOnlyList<int> writtenRows) =
            DailyLogRepository.SaveEntries(path, toCommit, guard);
        DailyLogLoadResult updated = Load(path, guard);

        for (int i = 0; i < toCommit.Count; i++)
        {
            DailyLogEntry entry = toCommit[i].Entry;
            results.Add(new DailyLogSaveResult
            {
                Succeeded = true,
                WrittenRowIndex = writtenRows[i],
                NormalizedEntry = entry,
                Validation = DailyLogLifecycleValidator.ValidateEntry(entry),
                TransactionResult = tx,
                Duplicates = updated.Duplicates,
                SharedDocuments = updated.SharedDocuments,
                Notification = DailyLogLifecycleValidator.FormatSapIntakeNotification(entry)
            });
        }
    }
}
