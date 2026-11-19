using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Acr.Infrastructure.Safety;

namespace Acr.Infrastructure.Excel;

/// <summary>
/// Classification of an Excel workbook determined strictly from its internal worksheet signatures.
/// </summary>
public enum WorkbookKind
{
    /// <summary>Unreadable, corrupt, lockfile, or unrecognized workbook format.</summary>
    Unknown = 0,

    /// <summary>Quality Control inspection workbook containing <c>CONTROL DATA</c>.</summary>
    QcInspection = 1,

    /// <summary>Receiving closure workbook containing <c>CIERRE RECEP</c>.</summary>
    ReceivingClosure = 2,

    /// <summary>Hybrid workbook containing both <c>CONTROL DATA</c> and <c>CIERRE RECEP</c>.</summary>
    HybridQcAndClosure = 3,

    /// <summary>Operational Daily Log workbook containing <c>Maestro Contenedores</c> and <c>Catalogos</c>.</summary>
    DailyLog = 4,

    /// <summary>Valid OpenXML workbook without ACR-specific operational sheets.</summary>
    OtherWorkbook = 5
}

/// <summary>
/// Structural signature extracted from a workbook without loading full worksheet XML payloads.
/// </summary>
/// <param name="FilePath">Canonicalized workbook path.</param>
/// <param name="Kind">Detected <see cref="WorkbookKind"/>.</param>
/// <param name="SheetNames">Ordered list of worksheet names declared in <c>xl/workbook.xml</c>.</param>
/// <param name="HasVbaProject">True when <c>xl/vbaProject.bin</c> exists inside the package.</param>
/// <param name="ErrorMessage">Optional diagnostic error when inspection fails or skips a file.</param>
public sealed record WorkbookSignature(
    string FilePath,
    WorkbookKind Kind,
    IReadOnlyList<string> SheetNames,
    bool HasVbaProject,
    string? ErrorMessage = null);

/// <summary>
/// Inspects and classifies Excel OpenXML workbooks (<c>.xlsm</c>, <c>.xlsx</c>) by internal sheet signature
/// while safely skipping Office/OneDrive lockfiles (<c>~$</c>) and corrupt/legacy binary files.
/// </summary>
public static class WorkbookSignatureScanner
{
    /// <summary>
    /// Returns true if <paramref name="filePath"/> refers to an Office lockfile (<c>~$</c>) or temporary/backup file.
    /// </summary>
    public static bool IsLockOrTempFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return true;

        string fileName = Path.GetFileName(filePath);
        return fileName.StartsWith("~$", StringComparison.OrdinalIgnoreCase) ||
               fileName.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase) ||
               fileName.EndsWith(".bak", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Inspects a single workbook package (<c>xl/workbook.xml</c> and <c>xl/vbaProject.bin</c>) without
    /// loading worksheet bodies, returning <see cref="WorkbookKind.Unknown"/> on lockfiles or invalid archives.
    /// </summary>
    public static WorkbookSignature Inspect(string filePath, OperationalSafetyGuard? guard = null)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            return new WorkbookSignature("", WorkbookKind.Unknown, Array.Empty<string>(), false, "Path is empty.");

        if (IsLockOrTempFile(filePath))
            return new WorkbookSignature(filePath, WorkbookKind.Unknown, Array.Empty<string>(), false, "Skipped lock or temporary file.");

        string resolvedPath = guard is not null ? guard.EnsureCanRead(filePath) : Path.GetFullPath(filePath);
        if (!File.Exists(resolvedPath))
            return new WorkbookSignature(resolvedPath, WorkbookKind.Unknown, Array.Empty<string>(), false, "File not found.");

        try
        {
            using var stream = new FileStream(resolvedPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            IReadOnlyList<string> sheetNames = ReadSheetNames(archive);
            bool hasVba = FindEntry(archive, "xl/vbaProject.bin") is not null;
            WorkbookKind kind = ClassifySheets(sheetNames);
            return new WorkbookSignature(resolvedPath, kind, sheetNames, hasVba);
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or XmlException)
        {
            return new WorkbookSignature(resolvedPath, WorkbookKind.Unknown, Array.Empty<string>(), false, ex.Message);
        }
    }

    /// <summary>
    /// Scans <paramref name="rootDirectory"/> for <c>.xlsm</c> and <c>.xlsx</c> workbooks,
    /// skipping <c>~$</c> lockfiles and temporary/backup artifacts.
    /// </summary>
    public static IReadOnlyList<WorkbookSignature> ScanDirectory(
        string rootDirectory,
        bool recursive = true,
        OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rootDirectory);
        string resolvedRoot = guard is not null ? guard.EnsureCanRead(rootDirectory) : Path.GetFullPath(rootDirectory);
        if (!Directory.Exists(resolvedRoot))
            throw new DirectoryNotFoundException($"Directory not found: '{resolvedRoot}'.");

        List<string> candidateFiles = CollectCandidateWorkbookFiles(resolvedRoot, recursive, guard);
        var results = new List<WorkbookSignature>(candidateFiles.Count);
        foreach (string file in candidateFiles.OrderBy(path => path, StringComparer.OrdinalIgnoreCase))
            results.Add(Inspect(file, guard));

        return results;
    }

    private static List<string> CollectCandidateWorkbookFiles(
        string resolvedRoot,
        bool recursive,
        OperationalSafetyGuard? guard)
    {
        var files = new List<string>();
        var visitedDirs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Queue<string>();
        EnqueueUnvisitedDirectory(resolvedRoot, guard, visitedDirs, pending);

        var enumOptions = new EnumerationOptions
        {
            RecurseSubdirectories = false,
            AttributesToSkip = FileAttributes.System,
            IgnoreInaccessible = true
        };

        while (pending.Count > 0)
        {
            string currentDir = pending.Dequeue();
            AddDirectoryWorkbookFiles(currentDir, enumOptions, files);
            if (!recursive)
                continue;

            foreach (string subDir in EnumerateSubdirectoriesSafe(currentDir, enumOptions))
                EnqueueUnvisitedDirectory(subDir, guard, visitedDirs, pending);
        }

        return files;
    }

    private static void EnqueueUnvisitedDirectory(
        string directoryPath,
        OperationalSafetyGuard? guard,
        HashSet<string> visitedDirs,
        Queue<string> pending)
    {
        string? canonical = TryResolveCanonicalDirectory(directoryPath, guard);
        if (!string.IsNullOrEmpty(canonical) && visitedDirs.Add(canonical))
            pending.Enqueue(directoryPath);
    }

    private static string? TryResolveCanonicalDirectory(string directoryPath, OperationalSafetyGuard? guard)
    {
        try
        {
            if (guard is not null)
                return guard.CanonicalizePath(directoryPath);

            string full = Path.GetFullPath(directoryPath);
            string? target = Directory.ResolveLinkTarget(full, returnFinalTarget: true)?.FullName;
            string resolved = string.IsNullOrEmpty(target) ? full : Path.GetFullPath(target);
            return resolved.Length <= 3 && resolved.Length >= 2 && resolved[1] == ':'
                ? resolved
                : resolved.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            return null;
        }
    }

    private static void AddDirectoryWorkbookFiles(
        string currentDir,
        EnumerationOptions options,
        List<string> files)
    {
        try
        {
            foreach (string file in Directory.EnumerateFiles(currentDir, "*.*", options))
            {
                if (IsSupportedExcelExtension(file) && !IsLockOrTempFile(file))
                    files.Add(file);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
        }
    }

    private static IEnumerable<string> EnumerateSubdirectoriesSafe(string currentDir, EnumerationOptions options)
    {
        try
        {
            return Directory.EnumerateDirectories(currentDir, "*", options).ToList();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return Array.Empty<string>();
        }
    }

    private static bool IsSupportedExcelExtension(string filePath)
    {
        string ext = Path.GetExtension(filePath);
        return ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
               ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<string> ReadSheetNames(ZipArchive archive)
    {
        ZipArchiveEntry? workbookEntry = FindEntry(archive, "xl/workbook.xml");
        if (workbookEntry is null)
            return Array.Empty<string>();

        using Stream xmlStream = workbookEntry.Open();
        XDocument doc = XDocument.Load(xmlStream, LoadOptions.None);
        XNamespace mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        return doc.Descendants(mainNs + "sheet")
            .Select(el => el.Attribute("name")?.Value ?? "")
            .Where(name => name.Length > 0)
            .ToList();
    }

    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string entryPath) =>
        archive.GetEntry(entryPath) ??
        archive.Entries.FirstOrDefault(e => e.FullName.Equals(entryPath, StringComparison.OrdinalIgnoreCase));

    private static WorkbookKind ClassifySheets(IReadOnlyList<string> sheetNames)
    {
        if (sheetNames.Count == 0)
            return WorkbookKind.Unknown;

        bool hasControlData = sheetNames.Any(s => s.Contains("CONTROL DATA", StringComparison.OrdinalIgnoreCase));
        bool hasCierreRecep = sheetNames.Any(s => s.Contains("CIERRE RECEP", StringComparison.OrdinalIgnoreCase));
        bool hasMaestro = sheetNames.Any(s => s.Equals("Maestro Contenedores", StringComparison.OrdinalIgnoreCase));
        bool hasCatalogos = sheetNames.Any(s => s.Equals("Catalogos", StringComparison.OrdinalIgnoreCase));

        if (hasControlData && hasCierreRecep)
            return WorkbookKind.HybridQcAndClosure;
        if (hasControlData)
            return WorkbookKind.QcInspection;
        if (hasCierreRecep)
            return WorkbookKind.ReceivingClosure;
        if (hasMaestro && hasCatalogos)
            return WorkbookKind.DailyLog;

        return WorkbookKind.OtherWorkbook;
    }
}
