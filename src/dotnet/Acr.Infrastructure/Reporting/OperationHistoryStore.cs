using System.Collections.Concurrent;
using System.Text;
using System.Text.Json;
using Acr.Core.Models;
using Acr.Infrastructure.Safety;

namespace Acr.Infrastructure.Reporting;

/// <summary>
/// Categories of local operational history events recorded by ACR.
/// </summary>
public enum OperationHistoryKind
{
    /// <summary>3-way merchandise reconciliation comparison.</summary>
    Reconciliation = 0,

    /// <summary>Macro-safe corporate closure workbook export.</summary>
    ClosureExport = 1,

    /// <summary>Transactional Daily Log row insert or update.</summary>
    DailyLogUpdate = 2,

    /// <summary>Idempotent historical closure discovery and sync.</summary>
    HistoricalSync = 3
}

/// <summary>
/// Single chronological entry persisted in the local operation history JSON store.
/// </summary>
public sealed record OperationHistoryEntry
{
    /// <summary>Unique identifier of the recorded operation.</summary>
    public string OperationId { get; init; } = Guid.NewGuid().ToString("N");

    /// <summary>UTC timestamp when the operation occurred.</summary>
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;

    /// <summary>Operation category.</summary>
    public OperationHistoryKind Kind { get; init; } = OperationHistoryKind.Reconciliation;

    /// <summary>Associated container identifier.</summary>
    public string ContainerId { get; init; } = "";

    /// <summary>QC source workbook path, when applicable.</summary>
    public string? QcSourcePath { get; init; }

    /// <summary>Closure template path, when applicable.</summary>
    public string? ClosureTemplatePath { get; init; }

    /// <summary>Output file path generated or modified, when applicable.</summary>
    public string? OutputPath { get; init; }

    /// <summary>Saved JSON processing report path, when applicable.</summary>
    public string? ReportPath { get; init; }

    /// <summary>True when the operation completed cleanly.</summary>
    public bool Succeeded { get; init; }

    /// <summary>Human-readable one-line operational summary.</summary>
    public string SummaryText { get; init; } = "";

    /// <summary>Count of discrepant materials observed.</summary>
    public int DiscrepantMaterialsCount { get; init; }

    /// <summary>Count of candidate revuelto pairs detected.</summary>
    public int RevueltoPairsCount { get; init; }

    /// <summary>Error message when <see cref="Succeeded"/> is false.</summary>
    public string? ErrorMessage { get; init; }
}

/// <summary>
/// Lean local JSON persistence store for chronological <see cref="OperationHistoryEntry"/> records,
/// bounded by a FIFO cap and protected by <see cref="OperationalSafetyGuard"/> and <see cref="AtomicFileTransaction"/>.
/// </summary>
public static class OperationHistoryStore
{
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> FileLocks =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Appends <paramref name="entry"/> to the JSON history file at <paramref name="historyFilePath"/>,
    /// retaining at most <paramref name="maxEntries"/> recent records and creating a <c>.bak</c> on overwrite.
    /// </summary>
    public static string Append(
        string historyFilePath,
        OperationHistoryEntry entry,
        OperationalSafetyGuard? guard = null,
        int maxEntries = 500)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(historyFilePath);
        ArgumentNullException.ThrowIfNull(entry);
        if (maxEntries <= 0)
            throw new ArgumentOutOfRangeException(nameof(maxEntries), "maxEntries must be positive.");

        OperationalSafetyGuard effectiveGuard = guard ?? CreateDefaultPermissiveGuard();
        string safeTarget = effectiveGuard.ResolveSafeWritePath(
            historyFilePath,
            OperationalActionKind.WriteExplicitUserSaveOrSync,
            OperationalTargetKind.LocalReportArtifact);

        string canonicalKey = Path.GetFullPath(safeTarget);
        SemaphoreSlim gate = FileLocks.GetOrAdd(canonicalKey, static _ => new SemaphoreSlim(1, 1));
        gate.Wait();
        try
        {
            return AppendSynchronized(safeTarget, entry, effectiveGuard, maxEntries);
        }
        finally
        {
            gate.Release();
        }
    }

    private static string AppendSynchronized(
        string safeTarget,
        OperationHistoryEntry entry,
        OperationalSafetyGuard effectiveGuard,
        int maxEntries)
    {
        List<OperationHistoryEntry> entries = File.Exists(safeTarget)
            ? LoadAll(safeTarget, effectiveGuard).ToList()
            : [];
        entries.Add(entry);
        if (entries.Count > maxEntries)
            entries = entries.Skip(entries.Count - maxEntries).ToList();

        AtomicTransactionResult tx = AtomicFileTransaction.Execute(
            safeTarget,
            effectiveGuard,
            OperationalActionKind.WriteExplicitUserSaveOrSync,
            tempPath => WriteHistoryAtomic(tempPath, entries),
            ValidateHistoryTempFile,
            committedPath => File.Exists(committedPath) && new FileInfo(committedPath).Length > 0,
            OperationalTargetKind.LocalReportArtifact);

        return tx.TargetPath;
    }

    /// <summary>
    /// Loads all recorded <see cref="OperationHistoryEntry"/> items from <paramref name="historyFilePath"/>,
    /// returning an empty list when the file does not yet exist.
    /// </summary>
    public static IReadOnlyList<OperationHistoryEntry> LoadAll(
        string historyFilePath,
        OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(historyFilePath);
        string resolved = guard is not null ? guard.EnsureCanRead(historyFilePath) : Path.GetFullPath(historyFilePath);
        if (!File.Exists(resolved))
            return Array.Empty<OperationHistoryEntry>();

        string json = File.ReadAllText(resolved, Encoding.UTF8);
        return JsonSerializer.Deserialize<List<OperationHistoryEntry>>(json, JsonProcessingReportStore.JsonOptions)
            ?? throw new InvalidDataException($"Deserialized null operation history list from '{resolved}'.");
    }

    /// <summary>
    /// Returns the most recent <paramref name="count"/> entries ordered newest-first.
    /// </summary>
    public static IReadOnlyList<OperationHistoryEntry> GetRecent(
        string historyFilePath,
        int count,
        OperationalSafetyGuard? guard = null)
    {
        if (count <= 0)
            return Array.Empty<OperationHistoryEntry>();

        IReadOnlyList<OperationHistoryEntry> all = LoadAll(historyFilePath, guard);
        return all.OrderByDescending(e => e.TimestampUtc).Take(count).ToList();
    }

    private static OperationalSafetyGuard CreateDefaultPermissiveGuard()
    {
        string sentinelRoot = Path.Combine(Path.GetTempPath(), $"AcrNoProdHistory_{Guid.NewGuid():N}");
        return new OperationalSafetyGuard(
            OperationalSafetyMode.Production,
            sentinelRoot,
            Path.Combine(sentinelRoot, "NoDailyLog.xlsx"));
    }

    private static void WriteHistoryAtomic(string tempPath, List<OperationHistoryEntry> entries)
    {
        string json = JsonSerializer.Serialize(entries, JsonProcessingReportStore.JsonOptions);
        File.WriteAllText(tempPath, json, Encoding.UTF8);
    }

    private static bool ValidateHistoryTempFile(string tempPath)
    {
        try
        {
            string json = File.ReadAllText(tempPath, Encoding.UTF8);
            List<OperationHistoryEntry>? parsed =
                JsonSerializer.Deserialize<List<OperationHistoryEntry>>(json, JsonProcessingReportStore.JsonOptions);
            return parsed is not null;
        }
        catch (JsonException)
        {
            return false;
        }
    }
}
