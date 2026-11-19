using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Acr.Core.Models;
using Acr.Core.Services;
using Acr.Infrastructure.Excel;
using Acr.Infrastructure.Safety;

namespace Acr.Infrastructure.Sync;

/// <summary>
/// Summary produced by scanning a directory tree for historical closure workbooks authored by the target user.
/// </summary>
/// <param name="RootDirectory">Canonicalized root directory scanned.</param>
/// <param name="TotalWorkbooksScanned">Total OpenXML workbooks inspected via <see cref="WorkbookSignatureScanner"/>.</param>
/// <param name="TargetAuthorWorkbooksMatched">Number of closure workbooks matching the target author in <c>DATOS GENERALES!B33</c>.</param>
/// <param name="RawRecords">All extracted raw closure records matching the target author.</param>
/// <param name="CollapsedCandidates">Deduplicated candidates grouped by operational Business Identity.</param>
/// <param name="MatchResults">Idempotency match outcomes against existing Daily Log entries.</param>
/// <param name="SkippedOrErrorFiles">Lockfiles, corrupt workbooks, or closures with structural extraction errors.</param>
public sealed record HistoricalSyncPreview(
    string RootDirectory,
    int TotalWorkbooksScanned,
    int TargetAuthorWorkbooksMatched,
    IReadOnlyList<HistoricalSyncRecord> RawRecords,
    IReadOnlyList<CollapsedHistoricalCandidate> CollapsedCandidates,
    IReadOnlyList<HistoricalMatchResult> MatchResults,
    IReadOnlyList<string> SkippedOrErrorFiles);

/// <summary>
/// Scans directories by worksheet signature (<c>CIERRE RECEP</c>), reads authorship strictly from
/// <c>'DATOS GENERALES'!B33</c> (never <c>sharedStrings.xml</c>), separates Business Identity from File Provenance,
/// collapses drafts/copies, and matches idempotently against Daily Log entries.
/// </summary>
public static class HistoricalClosureScanner
{
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static readonly string[] TargetDgRefs =
        ["B33", "C7", "C8", "C9", "C10", "C13", "C14", "C18", "C30"];

    private sealed record CierreRecepSummary(
        int ItemCount,
        int DistinctPoCount,
        IReadOnlyList<string> ErrorTokens);

    private readonly record struct DecodedCell(string? Text, bool IsError);

    /// <summary>
    /// Scans <paramref name="rootDirectory"/> for corporate closure workbooks, filters strictly on
    /// <c>'DATOS GENERALES'!B33</c>, evaluates anomalies, collapses duplicates, and matches against <paramref name="existingLog"/>.
    /// </summary>
    public static HistoricalSyncPreview Scan(
        string rootDirectory,
        IReadOnlyList<DailyLogEntry> existingLog,
        OperationalSafetyGuard? guard = null,
        bool recursive = true)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(rootDirectory);
        ArgumentNullException.ThrowIfNull(existingLog);

        string resolvedRoot = guard is not null ? guard.EnsureCanRead(rootDirectory) : Path.GetFullPath(rootDirectory);
        IReadOnlyList<WorkbookSignature> signatures = WorkbookSignatureScanner.ScanDirectory(resolvedRoot, recursive, guard);

        var rawRecords = new List<HistoricalSyncRecord>();
        var skippedOrError = new List<string>();
        foreach (WorkbookSignature sig in signatures)
            ProcessScannedWorkbook(sig, rawRecords, skippedOrError);

        IReadOnlyList<CollapsedHistoricalCandidate> collapsed = HistoricalIdentityMatcher.CollapseDuplicates(rawRecords);
        IReadOnlyList<HistoricalMatchResult> matchResults = HistoricalIdentityMatcher.MatchAgainstDailyLog(collapsed, existingLog);
        return new HistoricalSyncPreview(
            resolvedRoot,
            signatures.Count,
            rawRecords.Count,
            rawRecords,
            collapsed,
            matchResults,
            skippedOrError);
    }

    private static void ProcessScannedWorkbook(
        WorkbookSignature sig,
        List<HistoricalSyncRecord> rawRecords,
        List<string> skippedOrError)
    {
        if (sig.Kind == WorkbookKind.Unknown && sig.ErrorMessage is not null)
        {
            skippedOrError.Add($"{sig.FilePath}: {sig.ErrorMessage}");
            return;
        }

        if (sig.Kind is not (WorkbookKind.ReceivingClosure or WorkbookKind.HybridQcAndClosure))
            return;

        if (!TryExtractClosureRecord(sig, out HistoricalSyncRecord? record, out string? errorMessage))
        {
            if (errorMessage is not null)
                skippedOrError.Add($"{sig.FilePath}: {errorMessage}");
            return;
        }

        if (record is not null && HistoricalIdentityMatcher.IsTargetAuthor(record.Provenance.RawAuthorCellB33))
            rawRecords.Add(record);
    }

    private static bool TryExtractClosureRecord(
        WorkbookSignature sig,
        out HistoricalSyncRecord? record,
        out string? errorMessage)
    {
        record = null;
        errorMessage = null;

        try
        {
            using var stream = new FileStream(sig.FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            if (!TryResolveClosureEntries(archive, out ZipArchiveEntry? dgEntry, out ZipArchiveEntry? crEntry, out bool is1904))
            {
                errorMessage = "Missing 'DATOS GENERALES' or 'CIERRE RECEP' worksheet.";
                return false;
            }

            IReadOnlyList<string> sharedStrings = ReadSharedStrings(archive);
            Dictionary<string, DecodedCell> dgCells = ReadTargetCellsFromSheet(dgEntry!, sharedStrings);
            string? rawAuthorB33 = dgCells.TryGetValue("B33", out DecodedCell b33) ? b33.Text : null;
            if (!HistoricalIdentityMatcher.IsTargetAuthor(rawAuthorB33))
                return true;

            CierreRecepSummary crSummary = ScanCierreRecepSummary(crEntry!, sharedStrings);
            record = AssembleRecord(sig, archive, dgCells, crSummary, is1904);
            return true;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or XmlException or OverflowException or FormatException)
        {
            errorMessage = ex.Message;
            return false;
        }
    }

    private static bool TryResolveClosureEntries(
        ZipArchive archive,
        out ZipArchiveEntry? dgEntry,
        out ZipArchiveEntry? crEntry,
        out bool is1904)
    {
        dgEntry = null;
        crEntry = null;
        is1904 = false;

        ZipArchiveEntry? wbEntry = FindEntry(archive, "xl/workbook.xml");
        ZipArchiveEntry? relsEntry = FindEntry(archive, "xl/_rels/workbook.xml.rels");
        if (wbEntry is null || relsEntry is null)
            return false;

        XDocument wbDoc = LoadXmlEntry(wbEntry);
        XDocument relsDoc = LoadXmlEntry(relsEntry);
        is1904 = IsWorkbook1904(wbDoc);

        Dictionary<string, string> relMap = BuildRelationshipTargetMap(relsDoc);
        return ResolveTargetSheetEntries(archive, wbDoc, relMap, out dgEntry, out crEntry);
    }

    private static bool IsWorkbook1904(XDocument wbDoc)
    {
        string? raw = wbDoc.Descendants(XName.Get("workbookPr", MainNs))
            .Attributes("date1904")
            .FirstOrDefault()?.Value;
        return raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
    }

    private static Dictionary<string, string> BuildRelationshipTargetMap(XDocument relsDoc)
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (XElement rel in relsDoc.Root?.Elements() ?? Enumerable.Empty<XElement>())
        {
            string? id = rel.Attribute("Id")?.Value;
            string? target = rel.Attribute("Target")?.Value;
            if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                map[id] = NormalizePackageTarget(target);
        }
        return map;
    }

    private static bool ResolveTargetSheetEntries(
        ZipArchive archive,
        XDocument wbDoc,
        Dictionary<string, string> relMap,
        out ZipArchiveEntry? dgEntry,
        out ZipArchiveEntry? crEntry)
    {
        dgEntry = null;
        crEntry = null;
        XName relIdAttr = XName.Get("id", RelNs);

        foreach (XElement sheet in wbDoc.Descendants(XName.Get("sheet", MainNs)))
        {
            string name = (sheet.Attribute("name")?.Value ?? "").Trim();
            string? relId = sheet.Attribute(relIdAttr)?.Value;
            if (string.IsNullOrEmpty(relId) || !relMap.TryGetValue(relId, out string? target))
                continue;

            if (dgEntry is null && name.Equals("DATOS GENERALES", StringComparison.OrdinalIgnoreCase))
                dgEntry = FindEntry(archive, target);
            else if (crEntry is null && name.Contains("CIERRE RECEP", StringComparison.OrdinalIgnoreCase))
                crEntry = FindEntry(archive, target);
        }

        return dgEntry is not null && crEntry is not null;
    }

    private static string NormalizePackageTarget(string target)
    {
        string cleaned = target.Replace('\\', '/').TrimStart('/');
        return cleaned.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)
            ? cleaned
            : $"xl/{cleaned}";
    }

    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string entryPath) =>
        archive.GetEntry(entryPath) ??
        archive.Entries.FirstOrDefault(e => e.FullName.Equals(entryPath, StringComparison.OrdinalIgnoreCase));

    private static XDocument LoadXmlEntry(ZipArchiveEntry entry)
    {
        using Stream stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.None);
    }

    private static IReadOnlyList<string> ReadSharedStrings(ZipArchive archive)
    {
        ZipArchiveEntry? sstEntry = FindEntry(archive, "xl/sharedStrings.xml");
        if (sstEntry is null)
            return Array.Empty<string>();

        var strings = new List<string>();
        using Stream stream = sstEntry.Open();
        using XmlReader reader = XmlReader.Create(stream);
        while (!reader.EOF)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si" && reader.NamespaceURI == MainNs)
            {
                if (XNode.ReadFrom(reader) is XElement siEl)
                    strings.Add(ExtractRichText(siEl));
                continue;
            }

            if (!reader.Read())
                break;
        }
        return strings;
    }

    private static string ExtractRichText(XElement container)
    {
        var sb = new StringBuilder();
        foreach (XElement child in container.Elements())
        {
            if (child.Name.LocalName == "t")
                sb.Append(child.Value);
            else if (child.Name.LocalName == "r")
                sb.Append(child.Elements().FirstOrDefault(e => e.Name.LocalName == "t")?.Value);
        }
        return sb.ToString();
    }

    private static Dictionary<string, DecodedCell> ReadTargetCellsFromSheet(
        ZipArchiveEntry dgEntry,
        IReadOnlyList<string> sharedStrings)
    {
        var targets = new HashSet<string>(TargetDgRefs, StringComparer.OrdinalIgnoreCase);
        var found = new Dictionary<string, DecodedCell>(StringComparer.OrdinalIgnoreCase);

        using Stream stream = dgEntry.Open();
        using XmlReader reader = XmlReader.Create(stream);
        while (!reader.EOF)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c" && reader.NamespaceURI == MainNs)
            {
                string? cellRef = reader.GetAttribute("r");
                if (!string.IsNullOrEmpty(cellRef) && targets.Contains(cellRef))
                {
                    if (XNode.ReadFrom(reader) is XElement cellEl)
                        found[cellRef.ToUpperInvariant()] = DecodeCellElement(cellEl, sharedStrings);
                    continue;
                }
            }

            if (!reader.Read())
                break;
        }

        return found;
    }

    private static CierreRecepSummary ScanCierreRecepSummary(
        ZipArchiveEntry crEntry,
        IReadOnlyList<string> sharedStrings)
    {
        int itemCount = 0;
        var distinctPos = new HashSet<string>(StringComparer.Ordinal);
        var errorTokens = new List<string>();
        int emptyRows = 0;
        int lastRowIdx = 0;

        using Stream stream = crEntry.Open();
        using XmlReader reader = XmlReader.Create(stream);
        while (!reader.EOF)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row" && reader.NamespaceURI == MainNs)
            {
                if (XNode.ReadFrom(reader) is XElement rowEl &&
                    ShouldStopAfterRow(rowEl, sharedStrings, ref lastRowIdx, ref itemCount, distinctPos, errorTokens, ref emptyRows))
                {
                    break;
                }
                continue;
            }

            if (!reader.Read())
                break;
        }

        return new CierreRecepSummary(itemCount, distinctPos.Count, errorTokens);
    }

    private static bool ShouldStopAfterRow(
        XElement rowEl,
        IReadOnlyList<string> sharedStrings,
        ref int lastRowIdx,
        ref int itemCount,
        HashSet<string> distinctPos,
        List<string> errorTokens,
        ref int emptyRows)
    {
        int rowIdx = int.TryParse(rowEl.Attribute("r")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int explicitRow)
            ? explicitRow
            : lastRowIdx + 1;
        lastRowIdx = rowIdx;
        if (rowIdx < 13)
            return false;

        if (ProcessCierreRecepRow(rowEl, sharedStrings, ref itemCount, distinctPos, errorTokens))
        {
            emptyRows = 0;
            return false;
        }

        return ++emptyRows >= 20;
    }

    private static bool ProcessCierreRecepRow(
        XElement rowEl,
        IReadOnlyList<string> sharedStrings,
        ref int itemCount,
        HashSet<string> distinctPos,
        List<string> errorTokens)
    {
        string? colA = null;
        string? colE = null;
        bool colAError = false;
        bool colEError = false;
        bool hasAnyContent = false;
        int nextCol = 0;

        foreach (XElement cellEl in rowEl.Elements(XName.Get("c", MainNs)))
        {
            int colIdx = QcReader.ParseColumnIndex(cellEl.Attribute("r")?.Value, ref nextCol);
            DecodedCell decoded = DecodeCellElement(cellEl, sharedStrings);
            if (!string.IsNullOrEmpty(decoded.Text))
                hasAnyContent = true;

            if (decoded.IsError && !string.IsNullOrEmpty(decoded.Text))
                AddUniqueToken(errorTokens, decoded.Text);

            if (colIdx == 0)
            {
                colA = decoded.Text;
                colAError = decoded.IsError;
            }
            else if (colIdx == 4)
            {
                colE = decoded.Text;
                colEError = decoded.IsError;
            }
        }

        RecordCierreRecepItem(colA, colE, colAError, colEError, ref itemCount, distinctPos);
        return hasAnyContent;
    }

    private static void RecordCierreRecepItem(
        string? colA,
        string? colE,
        bool colAError,
        bool colEError,
        ref int itemCount,
        HashSet<string> distinctPos)
    {
        if (colEError || string.IsNullOrWhiteSpace(colE))
            return;

        string cleanE = colE.Trim();
        if (cleanE.Equals("MATERIAL", StringComparison.OrdinalIgnoreCase) ||
            cleanE.Contains("TOTAL", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        itemCount++;
        if (colAError || string.IsNullOrWhiteSpace(colA))
            return;

        foreach (string po in HistoricalIdentityMatcher.ParsePoList(colA))
            distinctPos.Add(po);
    }

    private static DecodedCell DecodeCellElement(XElement cellEl, IReadOnlyList<string> sharedStrings)
    {
        string? type = cellEl.Attribute("t")?.Value;
        string? valText = cellEl.Element(XName.Get("v", MainNs))?.Value;

        if (type == "e")
            return new DecodedCell(valText?.Trim(), IsError: true);

        if (type == "s")
            return DecodeSharedStringValue(valText, sharedStrings);

        if (type == "inlineStr")
        {
            XElement? isEl = cellEl.Element(XName.Get("is", MainNs));
            string? text = isEl is null ? null : ExtractRichText(isEl);
            return new DecodedCell(text, IsErrorToken(text));
        }

        if (type == "str")
            return new DecodedCell(valText, IsErrorToken(valText));

        string? numericText = NormalizeNumericText(valText);
        return new DecodedCell(numericText, IsErrorToken(numericText));
    }

    private static DecodedCell DecodeSharedStringValue(string? rawIndex, IReadOnlyList<string> sharedStrings)
    {
        if (!int.TryParse(rawIndex, NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx) ||
            idx < 0 || idx >= sharedStrings.Count)
        {
            return new DecodedCell(null, false);
        }

        string value = sharedStrings[idx];
        return new DecodedCell(value, IsErrorToken(value));
    }

    private static string? NormalizeNumericText(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;

        string trimmed = raw.Trim();
        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
            !double.IsNaN(value) && !double.IsInfinity(value))
        {
            double rounded = Math.Round(value, MidpointRounding.ToEven);
            if (Math.Abs(value - rounded) < 1e-6 && rounded >= long.MinValue && rounded < 9223372036854775808.0)
                return ((long)rounded).ToString(CultureInfo.InvariantCulture);
        }

        return trimmed;
    }

    private static bool IsErrorToken(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return text.Trim().StartsWith('#');
    }

    private static HistoricalSyncRecord AssembleRecord(
        WorkbookSignature sig,
        ZipArchive archive,
        Dictionary<string, DecodedCell> dgCells,
        CierreRecepSummary crSummary,
        bool is1904)
    {
        var errorTokens = new List<string>(crSummary.ErrorTokens);
        CollectDgErrors(dgCells, errorTokens);

        BusinessIdentity identity = BuildBusinessIdentity(dgCells, is1904);
        FileProvenance provenance = BuildFileProvenance(sig, archive, dgCells);

        var draftRecord = new HistoricalSyncRecord
        {
            Identity = identity,
            Provenance = provenance,
            TheoreticalCartons = ParseNullableInt(GetDgText(dgCells, "C13")),
            ReceivedCartons = ParseNullableInt(GetDgText(dgCells, "C14")),
            DeclaredLinesCountDg = ParseNullableInt(GetDgText(dgCells, "C18")),
            ActualCierreRecepItemsCount = crSummary.ItemCount,
            CierreRecepDistinctPoCount = crSummary.DistinctPoCount,
            VendorNameRaw = CleanVendorName(GetDgText(dgCells, "C10")),
            IsHybridQcClosure = sig.Kind == WorkbookKind.HybridQcAndClosure,
            IsDraftOrCopyFile = IsDraftOrCopyFilename(provenance.SourceFileName),
            DetectedErrorTokens = errorTokens
        };

        return draftRecord with { ReviewReasons = HistoricalIdentityMatcher.EvaluateAnomalies(draftRecord) };
    }

    private static BusinessIdentity BuildBusinessIdentity(Dictionary<string, DecodedCell> dgCells, bool is1904)
    {
        string rawContainer = GetDgText(dgCells, "C7") ?? "";
        string? rawPedido = CleanNullablePedido(GetDgText(dgCells, "C9"));
        IReadOnlyList<string> po43List = HistoricalIdentityMatcher.ParsePoList(GetDgText(dgCells, "C8"));
        DateOnly? revisionDate = DecodeExcelDate(GetDgText(dgCells, "C30"), is1904);

        return new BusinessIdentity
        {
            NormalizedContainerId = DailyLogLifecycleValidator.NormalizeContainerId(rawContainer),
            Pedido41 = rawPedido,
            Po43Numbers = po43List,
            RevisionDate = revisionDate
        };
    }

    private static FileProvenance BuildFileProvenance(
        WorkbookSignature sig,
        ZipArchive archive,
        Dictionary<string, DecodedCell> dgCells)
    {
        string fileSha256 = ComputeFileSha256(sig.FilePath);
        (string? vbaSha256, string? coreCreator) = ReadPackageMetadata(archive);
        string? rawAuthor = GetDgText(dgCells, "B33");

        return new FileProvenance
        {
            SourceFileName = Path.GetFileName(sig.FilePath),
            SourcePath = sig.FilePath,
            FileSha256 = fileSha256,
            VbaProjectSha256 = vbaSha256,
            DiscoveredAtUtc = DateTimeOffset.UtcNow,
            RawAuthorCellB33 = rawAuthor,
            NormalizedAuthor = HistoricalIdentityMatcher.NormalizeAuthorName(rawAuthor),
            CoreXmlCreator = coreCreator
        };
    }

    private static string ComputeFileSha256(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static (string? VbaSha256, string? CoreCreator) ReadPackageMetadata(ZipArchive archive)
    {
        string? vbaHash = null;
        ZipArchiveEntry? vbaEntry = FindEntry(archive, "xl/vbaProject.bin");
        if (vbaEntry is not null)
        {
            using Stream vbaStream = vbaEntry.Open();
            vbaHash = Convert.ToHexString(SHA256.HashData(vbaStream));
        }

        string? creator = ReadCoreXmlCreator(archive);
        return (vbaHash, creator);
    }

    private static string? ReadCoreXmlCreator(ZipArchive archive)
    {
        ZipArchiveEntry? coreEntry = FindEntry(archive, "docProps/core.xml");
        if (coreEntry is null)
            return null;

        XDocument doc = LoadXmlEntry(coreEntry);
        XNamespace dcNs = "http://purl.org/dc/elements/1.1/";
        string? creator = doc.Descendants(dcNs + "creator").FirstOrDefault()?.Value?.Trim();
        return string.IsNullOrEmpty(creator) ? null : creator;
    }

    private static void CollectDgErrors(Dictionary<string, DecodedCell> dgCells, List<string> errorTokens)
    {
        foreach (DecodedCell cell in dgCells.Values)
        {
            if (cell.IsError && !string.IsNullOrEmpty(cell.Text))
                AddUniqueToken(errorTokens, cell.Text);
        }
    }

    private static void AddUniqueToken(List<string> tokens, string token)
    {
        string trimmed = token.Trim();
        if (trimmed.Length > 0 && !tokens.Contains(trimmed, StringComparer.Ordinal))
            tokens.Add(trimmed);
    }

    private static string? GetDgText(Dictionary<string, DecodedCell> dgCells, string cellRef) =>
        dgCells.TryGetValue(cellRef, out DecodedCell cell) && !cell.IsError ? cell.Text : null;

    private static int? ParseNullableInt(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;
        if (!double.TryParse(raw.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ||
            double.IsNaN(value) || double.IsInfinity(value))
        {
            return null;
        }

        double rounded = Math.Round(value, MidpointRounding.ToEven);
        return rounded is >= int.MinValue and <= int.MaxValue
            ? (int)rounded
            : null;
    }

    private static string? CleanNullablePedido(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;
        string trimmed = raw.Trim();
        return trimmed == "-" || trimmed.StartsWith('#') ? null : trimmed;
    }

    private static string? CleanVendorName(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;
        string cleaned = raw.Trim();
        return cleaned.Length == 0 ? null : cleaned;
    }

    private static bool IsDraftOrCopyFilename(string fileName) =>
        fileName.Contains("copia de", StringComparison.OrdinalIgnoreCase) ||
        fileName.Contains("copy of", StringComparison.OrdinalIgnoreCase) ||
        fileName.Contains("inprogress", StringComparison.OrdinalIgnoreCase) ||
        fileName.Contains("in progress", StringComparison.OrdinalIgnoreCase) ||
        fileName.Contains("borrador", StringComparison.OrdinalIgnoreCase);

    private static DateOnly? DecodeExcelDate(string? raw, bool is1904)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;

        string trimmed = raw.Trim();
        if (double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial))
        {
            double adjusted = is1904 ? serial + 1462.0 : serial;
            if (adjusted is > -657435.0 and <= 2958465.0)
            {
                try
                {
                    return DateOnly.FromDateTime(DateTime.FromOADate(adjusted));
                }
                catch (ArgumentException)
                {
                    return null;
                }
            }
        }

        return DateOnly.TryParse(trimmed, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly parsed)
            ? parsed
            : null;
    }
}
