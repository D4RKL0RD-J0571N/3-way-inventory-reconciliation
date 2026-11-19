using System.Globalization;
using System.Text.RegularExpressions;
using Acr.Core.Models;
using Acr.Core.Services;
using Acr.Infrastructure.Safety;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Infrastructure.Excel;

/// <summary>
/// Direct OpenXML repository for reading and transactionally updating
/// <c>Control_Diario_Operaciones_Jostin.xlsx</c> (<c>Maestro Contenedores</c>, <c>Catalogos</c>,
/// <c>Dashboard General</c>, <c>tblMaestroContenedores</c>, and data validations).
/// </summary>
public static class DailyLogRepository
{
    private static readonly uint[] DefaultColumnStyles =
        [22, 21, 39, 41, 42, 42, 42, 42, 22, 22, 22, 22, 22, 22, 22, 26, 22, 26];

    private static readonly Regex TableRefRegex =
        new(@"^(\$?[A-Z]+\$?)(\d+):(\$?[A-Z]+\$?)(\d+)$", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex DashboardRangeRegex =
        new(@"('Maestro Contenedores'!\$?[A-Z]{1,3}\$?6:\$?[A-Z]{1,3}\$?)(\d+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex FilterDataRefRegex =
        new(@"('Maestro Contenedores'!\$A\$5:\$R\$)(\d+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    /// <summary>
    /// Reads all active entries, row indices, table metadata, and <c>Catalogos!A1:H7</c> options
    /// from the Daily Log workbook in read-only non-blocking mode.
    /// </summary>
    public static DailyLogWorkbookSnapshot ReadWorkbook(string path, OperationalSafetyGuard guard)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(guard);

        string canonical = guard.EnsureCanRead(path);
        EnsureValidDailyLogPackage(canonical);

        using WorkbookReadSession session = OpenXmlRelsSanitizer.OpenReadOnly(canonical);
        WorkbookPart wbPart = session.Document.WorkbookPart ?? throw new InvalidDataException($"Missing WorkbookPart in '{canonical}'.");

        List<string> sharedStrings = OpenXmlCellHelper.ReadSharedStrings(wbPart);
        bool is1904 = wbPart.Workbook?.WorkbookProperties?.Date1904?.Value == true;
        DailyLogCatalogSnapshot catalogs = ReadCatalogs(wbPart, sharedStrings);
        WorksheetPart maestroPart = ResolveMaestroPart(wbPart);

        (int tableStart, int tableEnd) = ReadTableBounds(maestroPart);
        var warnings = new List<string>();
        (List<DailyLogEntry> entries, List<int> rowIndices, Dictionary<string, int> rowByKey) =
            ReadMaestroRows(maestroPart, sharedStrings, is1904, warnings);

        return new DailyLogWorkbookSnapshot(
            canonical,
            entries,
            rowIndices,
            rowByKey,
            catalogs,
            tableStart,
            tableEnd,
            warnings);
    }

    /// <summary>
    /// Atomically persists a single <see cref="DailyLogEntry"/> into the Daily Log workbook, updating an existing
    /// row or appending after the last active row and resizing all table, filter, validation, and dashboard ranges.
    /// </summary>
    public static (AtomicTransactionResult Transaction, int WrittenRowIndex) SaveEntry(
        string path,
        DailyLogEntry entry,
        OperationalSafetyGuard guard,
        int? targetRowIndex = null)
    {
        ArgumentNullException.ThrowIfNull(entry);
        (AtomicTransactionResult tx, IReadOnlyList<int> rows) = SaveEntries(path, [(entry, targetRowIndex)], guard);
        return (tx, rows[0]);
    }

    /// <summary>
    /// Atomically persists a batch of <see cref="DailyLogEntry"/> mutations inside a single
    /// <see cref="AtomicFileTransaction"/> with pre-commit validation and post-commit verification.
    /// </summary>
    public static (AtomicTransactionResult Transaction, IReadOnlyList<int> WrittenRowIndices) SaveEntries(
        string path,
        IReadOnlyList<(DailyLogEntry Entry, int? TargetRowIndex)> mutations,
        OperationalSafetyGuard guard)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(mutations);
        ArgumentNullException.ThrowIfNull(guard);
        if (mutations.Count == 0)
            throw new ArgumentException("At least one Daily Log mutation is required.", nameof(mutations));

        string canonicalSource = guard.EnsureCanRead(path);
        EnsureValidDailyLogPackage(canonicalSource);

        var writtenRows = new List<int>();
        int expectedLastRow = 5;

        AtomicTransactionResult tx = AtomicFileTransaction.Execute(
            path,
            guard,
            OperationalActionKind.WriteExplicitUserSaveOrSync,
            tempPath => expectedLastRow = WriteMutationsToTemp(canonicalSource, tempPath, mutations, writtenRows),
            tempPath => ValidateTempWorkbook(tempPath, guard, expectedLastRow, writtenRows),
            committedPath => VerifyCommittedWorkbook(committedPath, guard, expectedLastRow));

        return (tx, writtenRows);
    }

    private static void EnsureValidDailyLogPackage(string canonical)
    {
        if (!File.Exists(canonical))
            throw new FileNotFoundException($"Daily Log workbook not found: '{canonical}'.", canonical);

        if (!IsDailyLogPackage(canonical))
            throw new InvalidDataException($"Workbook '{canonical}' is not a valid Daily Log (missing 'Maestro Contenedores' or 'Catalogos').");
    }

    private static bool IsDailyLogPackage(string path)
    {
        try
        {
            using WorkbookReadSession session = OpenXmlRelsSanitizer.OpenReadOnly(path);
            List<string> names = session.Document.WorkbookPart?.Workbook?.Sheets?.Elements<Sheet>()
                .Select(s => s.Name?.Value ?? "")
                .ToList() ?? [];
            return names.Any(n => n.Equals("Maestro Contenedores", StringComparison.OrdinalIgnoreCase)) &&
                   names.Any(n => n.Equals("Catalogos", StringComparison.OrdinalIgnoreCase));
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException)
        {
            return false;
        }
    }

    private static WorksheetPart ResolveMaestroPart(WorkbookPart wbPart) =>
        OpenXmlCellHelper.FindWorksheetPart(wbPart, "Maestro Contenedores")
        ?? throw new InvalidDataException("Daily Log is missing 'Maestro Contenedores' worksheet.");

    private static (int StartRow, int EndRow) ReadTableBounds(WorksheetPart maestroPart)
    {
        Table? table = maestroPart.TableDefinitionParts
            .Select(p => p.Table)
            .FirstOrDefault(t => t is not null &&
                string.Equals(t.Name?.Value ?? t.DisplayName?.Value, "tblMaestroContenedores", StringComparison.OrdinalIgnoreCase))
            ?? maestroPart.TableDefinitionParts.Select(p => p.Table).FirstOrDefault(t => t is not null);

        string? reference = table?.Reference?.Value;
        if (!string.IsNullOrWhiteSpace(reference))
        {
            Match m = TableRefRegex.Match(reference.Trim());
            if (m.Success &&
                int.TryParse(m.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int s) &&
                int.TryParse(m.Groups[4].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int e))
            {
                return (s, e);
            }
        }
        return (5, 5);
    }

    private static DailyLogCatalogSnapshot ReadCatalogs(WorkbookPart wbPart, IReadOnlyList<string> sharedStrings)
    {
        WorksheetPart? catPart = OpenXmlCellHelper.FindWorksheetPart(wbPart, "Catalogos");
        if (catPart?.Worksheet?.GetFirstChild<SheetData>() is not SheetData sheetData)
            return new DailyLogCatalogSnapshot();

        var columns = Enumerable.Range(0, 8).Select(_ => new List<string>()).ToArray();
        foreach (Row row in sheetData.Elements<Row>())
        {
            int rowIdx = checked((int)(row.RowIndex?.Value ?? 0));
            if (rowIdx is < 2 or > 7)
                continue;

            int fallbackCol = 0;
            foreach (Cell cell in row.Elements<Cell>())
            {
                int colIdx = QcReader.ParseColumnIndex(cell.CellReference?.Value, ref fallbackCol);
                if (colIdx is >= 0 and < 8)
                {
                    string val = OpenXmlCellHelper.DecodeCellText(cell, sharedStrings).Trim();
                    if (val.Length > 0)
                        columns[colIdx].Add(val);
                }
            }
        }

        return new DailyLogCatalogSnapshot
        {
            Stage1CustomsOptions = columns[0],
            Stage2SapIntakeOptions = columns[1],
            Stage3BlindSetupOptions = columns[2],
            Stage4QualityControlOptions = columns[3],
            Stage5ReceptionClosureOptions = columns[4],
            Stage6DispatchClosureOptions = columns[5],
            SaludoOptions = columns[6],
            ParteBStatusOptions = columns[7]
        };
    }

    private static (List<DailyLogEntry> Entries, List<int> RowIndices, Dictionary<string, int> RowByKey) ReadMaestroRows(
        WorksheetPart maestroPart,
        IReadOnlyList<string> sharedStrings,
        bool is1904,
        List<string> warnings)
    {
        var entries = new List<DailyLogEntry>();
        var rowIndices = new List<int>();
        var rowByKey = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        SheetData? sheetData = maestroPart.Worksheet?.GetFirstChild<SheetData>();
        if (sheetData is null)
            return (entries, rowIndices, rowByKey);

        foreach (Row row in sheetData.Elements<Row>())
        {
            int rowIdx = checked((int)(row.RowIndex?.Value ?? 0));
            if (rowIdx < 6)
                continue;

            string[] cols = ExtractRowColumns(row, sharedStrings);
            if (string.IsNullOrWhiteSpace(cols[1]))
                continue;

            DailyLogEntry entry = ParseRowToEntry(rowIdx, cols, is1904, warnings);
            entries.Add(entry);
            rowIndices.Add(rowIdx);
            if (!string.IsNullOrWhiteSpace(entry.NormalizedContainerId))
                rowByKey.TryAdd(entry.NormalizedContainerId, rowIdx);
        }

        return (entries, rowIndices, rowByKey);
    }

    private static string[] ExtractRowColumns(Row row, IReadOnlyList<string> sharedStrings)
    {
        var cols = new string[18];
        int fallbackCol = 0;
        foreach (Cell cell in row.Elements<Cell>())
        {
            int colIdx = QcReader.ParseColumnIndex(cell.CellReference?.Value, ref fallbackCol);
            if (colIdx is >= 0 and < 18)
                cols[colIdx] = OpenXmlCellHelper.DecodeCellText(cell, sharedStrings).Trim();
        }
        return cols;
    }

    private static DailyLogEntry ParseRowToEntry(
        int rowIdx,
        string[] cols,
        bool is1904,
        List<string> warnings)
    {
        string rawContainer = cols[1];
        string normalized = DailyLogLifecycleValidator.NormalizeContainerId(rawContainer);

        CustomsDuaStage s1 = ParseStage(cols[8], DailyLogCatalogs.TryParseCustomsDua, CustomsDuaStage.PendingDocs, rowIdx, "I", warnings);
        SapIntakeStage s2 = ParseStage(cols[9], DailyLogCatalogs.TryParseSapIntake, SapIntakeStage.NoDoc18Yet, rowIdx, "J", warnings);
        BlindSetupStage s3 = ParseStage(cols[10], DailyLogCatalogs.TryParseBlindSetup, BlindSetupStage.Pending, rowIdx, "K", warnings);
        QualityControlStage s4 = ParseStage(cols[11], DailyLogCatalogs.TryParseQualityControl, QualityControlStage.NoQcWaitingDock, rowIdx, "L", warnings);
        ReceptionClosureStage s5 = ParseStage(cols[12], DailyLogCatalogs.TryParseReceptionClosure, ReceptionClosureStage.PendingClosure, rowIdx, "M", warnings);
        DispatchClosureStage s6 = ParseStage(cols[13], DailyLogCatalogs.TryParseDispatchClosure, DispatchClosureStage.BlockedByReception, rowIdx, "N", warnings);
        Saludo greeting = ParseStage(cols[16], DailyLogCatalogs.TryParseSaludo, Saludo.None, rowIdx, "Q", warnings);

        var draft = new DailyLogEntry
        {
            SequenceNumber = QcReader.ToInt(cols[0]),
            ContainerId = rawContainer,
            NormalizedContainerId = normalized,
            UpdatedDate = ParseExcelDate(cols[2], is1904),
            Po43Numbers = HistoricalIdentityMatcher.ParsePoList(cols[3]),
            Pedido41 = CleanTrackingId(cols[4]),
            EntregaVl31n = CleanTrackingId(cols[5]),
            EwmDeliveryDoc = CleanTrackingId(cols[6]),
            HandlingUnit = CleanTrackingId(cols[7]),
            Stage1Customs = s1,
            Stage2SapIntake = s2,
            Stage3BlindSetup = s3,
            Stage4QualityControl = s4,
            Stage5ReceptionClosure = s5,
            Stage6DispatchClosure = s6,
            Observations = CleanObservations(cols[15]),
            Greeting = greeting
        };

        MasterContainerState master = DailyLogCatalogs.TryParseMasterState(cols[14], out MasterContainerState parsedMaster)
            ? parsedMaster
            : DailyLogLifecycleValidator.ComputeMasterState(draft);

        return draft with { MasterState = master };
    }

    private delegate bool CatalogTryParse<TEnum>(string? raw, out TEnum value) where TEnum : struct;

    private static TEnum ParseStage<TEnum>(
        string raw,
        CatalogTryParse<TEnum> tryParse,
        TEnum fallback,
        int rowIdx,
        string colLetter,
        List<string> warnings) where TEnum : struct
    {
        if (string.IsNullOrWhiteSpace(raw))
            return fallback;
        if (tryParse(raw, out TEnum parsed))
            return parsed;

        warnings.Add($"Row {rowIdx} Col {colLetter}: unrecognized catalog value '{raw}'; defaulted to '{fallback}'.");
        return fallback;
    }

    private static DateOnly? ParseExcelDate(string raw, bool is1904)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;

        if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial))
        {
            double adjusted = is1904 ? serial + 1462.0 : serial;
            if (adjusted is > -657435.0 and <= 2958465.0)
            {
                try { return DateOnly.FromDateTime(DateTime.FromOADate(adjusted)); }
                catch (ArgumentException) { return null; }
            }
        }

        return DateOnly.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly parsed)
            ? parsed
            : null;
    }

    private static string? CleanTrackingId(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
            return null;
        string trimmed = raw.Trim();
        return trimmed == "-" ? null : trimmed;
    }

    private static string CleanObservations(string raw) =>
        raw == "-" ? "" : raw;

    private static int WriteMutationsToTemp(
        string canonicalSource,
        string tempPath,
        IReadOnlyList<(DailyLogEntry Entry, int? TargetRowIndex)> mutations,
        List<int> writtenRows)
    {
        File.Copy(canonicalSource, tempPath, overwrite: true);
        OpenXmlRelsSanitizer.SanitizePackageFile(tempPath);
        using SpreadsheetDocument doc = SpreadsheetDocument.Open(tempPath, isEditable: true);
        WorkbookPart wbPart = doc.WorkbookPart ?? throw new InvalidDataException($"Missing WorkbookPart in '{tempPath}'.");
        WorksheetPart maestroPart = ResolveMaestroPart(wbPart);
        Worksheet maestroWs = maestroPart.Worksheet ?? throw new InvalidDataException("Missing Maestro worksheet.");
        SheetData sheetData = maestroWs.GetFirstChild<SheetData>() ?? maestroWs.AppendChild(new SheetData());

        bool is1904 = wbPart.Workbook?.WorkbookProperties?.Date1904?.Value == true;
        List<string> sharedStrings = OpenXmlCellHelper.ReadSharedStrings(wbPart);
        Dictionary<string, int> sstCache = OpenXmlCellHelper.BuildSharedStringCache(wbPart);
        (int lastDataRow, int maxSeq, Dictionary<string, List<int>> existingRows) = ScanExistingState(sheetData, sharedStrings);

        foreach ((DailyLogEntry entry, int? requestedRow) in mutations)
        {
            int rowIdx = ResolveMutationRowIndex(entry, requestedRow, existingRows, ref lastDataRow);
            int seq = entry.SequenceNumber ?? ResolveOrIncrementSeq(sheetData, rowIdx, sharedStrings, ref maxSeq);
            DailyLogEntry normalized = entry with
            {
                SequenceNumber = seq,
                NormalizedContainerId = DailyLogLifecycleValidator.NormalizeContainerId(entry.ContainerId),
                MasterState = DailyLogLifecycleValidator.ComputeMasterState(entry)
            };

            WriteDataRow(wbPart, sheetData, rowIdx, normalized, is1904, sstCache);
            writtenRows.Add(rowIdx);
            if (rowIdx > lastDataRow)
                lastDataRow = rowIdx;
        }

        List<string> updatedSst = OpenXmlCellHelper.ReadSharedStrings(wbPart);
        NormalizeAllFormulaColumns(sheetData, lastDataRow, updatedSst, is1904);
        ResizeWorkbookAnchors(wbPart, maestroPart, maestroWs, lastDataRow);
        OpenXmlCellHelper.PrepareWorkbookForFullRecalc(wbPart);

        SaveModifiedParts(wbPart, maestroPart, maestroWs);
        return lastDataRow;
    }

    private static (int LastDataRow, int MaxSeq, Dictionary<string, List<int>> RowsByKey) ScanExistingState(
        SheetData sheetData,
        IReadOnlyList<string> sharedStrings)
    {
        int lastDataRow = 5;
        int maxSeq = 0;
        var rowsByKey = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);

        foreach (Row row in sheetData.Elements<Row>())
        {
            int rowIdx = checked((int)(row.RowIndex?.Value ?? 0));
            if (rowIdx < 6)
                continue;

            Cell? bCell = OpenXmlCellHelper.TryFindCell(row, 1);
            string containerRaw = OpenXmlCellHelper.DecodeCellText(bCell, sharedStrings).Trim();
            if (containerRaw.Length == 0)
                continue;

            if (rowIdx > lastDataRow)
                lastDataRow = rowIdx;

            Cell? aCell = OpenXmlCellHelper.TryFindCell(row, 0);
            if (QcReader.ToInt(OpenXmlCellHelper.DecodeCellText(aCell, sharedStrings)) is int seq && seq > maxSeq)
                maxSeq = seq;

            string norm = DailyLogLifecycleValidator.NormalizeContainerId(containerRaw);
            if (norm.Length > 0)
            {
                if (!rowsByKey.TryGetValue(norm, out List<int>? list))
                {
                    list = [];
                    rowsByKey[norm] = list;
                }
                list.Add(rowIdx);
            }
        }

        return (lastDataRow, maxSeq, rowsByKey);
    }

    private static int ResolveMutationRowIndex(
        DailyLogEntry entry,
        int? requestedRow,
        Dictionary<string, List<int>> existingRows,
        ref int lastDataRow)
    {
        if (requestedRow.HasValue && requestedRow.Value >= 6)
            return requestedRow.Value;

        string norm = DailyLogLifecycleValidator.NormalizeContainerId(entry.ContainerId);
        if (norm.Length > 0 && existingRows.TryGetValue(norm, out List<int>? matches) && matches.Count == 1)
            return matches[0];

        int nextRow = lastDataRow + 1;
        lastDataRow = nextRow;
        if (norm.Length > 0)
            existingRows[norm] = [nextRow];
        return nextRow;
    }

    private static int ResolveOrIncrementSeq(
        SheetData sheetData,
        int rowIdx,
        IReadOnlyList<string> sharedStrings,
        ref int maxSeq)
    {
        Row row = OpenXmlCellHelper.FindOrCreateRow(sheetData, rowIdx);
        Cell? aCell = OpenXmlCellHelper.TryFindCell(row, 0);
        string text = OpenXmlCellHelper.DecodeCellText(aCell, sharedStrings);
        if (QcReader.ToInt(text) is int existing)
            return existing;
        return ++maxSeq;
    }

    private static void WriteDataRow(
        WorkbookPart wbPart,
        SheetData sheetData,
        int rowIdx,
        DailyLogEntry entry,
        bool is1904,
        Dictionary<string, int> sstCache)
    {
        Row row = OpenXmlCellHelper.FindOrCreateRow(sheetData, rowIdx);
        if (entry.SequenceNumber.HasValue)
            OpenXmlCellHelper.SetIntegerCell(row, 0, rowIdx, entry.SequenceNumber.Value, DefaultColumnStyles[0]);

        OpenXmlCellHelper.SetSharedStringCell(row, 1, rowIdx, entry.ContainerId.Trim(), wbPart, sstCache, DefaultColumnStyles[1]);

        if (entry.UpdatedDate.HasValue)
        {
            double oa = entry.UpdatedDate.Value.ToDateTime(TimeOnly.MinValue).ToOADate();
            double serial = is1904 ? oa - 1462.0 : oa;
            OpenXmlCellHelper.SetNumericCell(row, 2, rowIdx, serial, DefaultColumnStyles[2]);
        }
        else
        {
            ClearCellValuePreservingStyle(row, 2);
        }

        string poText = entry.Po43Numbers.Count > 0 ? string.Join(" / ", entry.Po43Numbers) : "-";
        OpenXmlCellHelper.SetSharedStringCell(row, 3, rowIdx, poText, wbPart, sstCache, DefaultColumnStyles[3]);
        WriteTrackingAndStageCells(wbPart, row, rowIdx, entry, sstCache);
    }

    private static void ClearCellValuePreservingStyle(Row row, int colIdx)
    {
        Cell? cell = OpenXmlCellHelper.TryFindCell(row, colIdx);
        if (cell is null)
            return;

        cell.CellFormula = null;
        cell.InlineString = null;
        cell.CellValue = null;
        cell.DataType = null;
    }

    private static void WriteTrackingAndStageCells(
        WorkbookPart wbPart,
        Row row,
        int rowIdx,
        DailyLogEntry entry,
        Dictionary<string, int> sstCache)
    {
        OpenXmlCellHelper.SetSharedStringCell(row, 4, rowIdx, FormatTracking(entry.Pedido41), wbPart, sstCache, DefaultColumnStyles[4]);
        OpenXmlCellHelper.SetSharedStringCell(row, 5, rowIdx, FormatTracking(entry.EntregaVl31n), wbPart, sstCache, DefaultColumnStyles[5]);
        OpenXmlCellHelper.SetSharedStringCell(row, 6, rowIdx, FormatTracking(entry.EwmDeliveryDoc), wbPart, sstCache, DefaultColumnStyles[6]);
        OpenXmlCellHelper.SetSharedStringCell(row, 7, rowIdx, FormatTracking(entry.HandlingUnit), wbPart, sstCache, DefaultColumnStyles[7]);

        OpenXmlCellHelper.SetSharedStringCell(row, 8, rowIdx, DailyLogCatalogs.ToExcelString(entry.Stage1Customs), wbPart, sstCache, DefaultColumnStyles[8]);
        OpenXmlCellHelper.SetSharedStringCell(row, 9, rowIdx, DailyLogCatalogs.ToExcelString(entry.Stage2SapIntake), wbPart, sstCache, DefaultColumnStyles[9]);
        OpenXmlCellHelper.SetSharedStringCell(row, 10, rowIdx, DailyLogCatalogs.ToExcelString(entry.Stage3BlindSetup), wbPart, sstCache, DefaultColumnStyles[10]);
        OpenXmlCellHelper.SetSharedStringCell(row, 11, rowIdx, DailyLogCatalogs.ToExcelString(entry.Stage4QualityControl), wbPart, sstCache, DefaultColumnStyles[11]);
        OpenXmlCellHelper.SetSharedStringCell(row, 12, rowIdx, DailyLogCatalogs.ToExcelString(entry.Stage5ReceptionClosure), wbPart, sstCache, DefaultColumnStyles[12]);
        OpenXmlCellHelper.SetSharedStringCell(row, 13, rowIdx, DailyLogCatalogs.ToExcelString(entry.Stage6DispatchClosure), wbPart, sstCache, DefaultColumnStyles[13]);

        string obs = string.IsNullOrWhiteSpace(entry.Observations) ? "-" : entry.Observations.Trim();
        OpenXmlCellHelper.SetSharedStringCell(row, 15, rowIdx, obs, wbPart, sstCache, DefaultColumnStyles[15]);
        OpenXmlCellHelper.SetSharedStringCell(row, 16, rowIdx, DailyLogCatalogs.ToExcelString(entry.Greeting), wbPart, sstCache, DefaultColumnStyles[16]);
    }

    private static string FormatTracking(string? value) =>
        string.IsNullOrWhiteSpace(value) ? "-" : value.Trim();

    private static void NormalizeAllFormulaColumns(
        SheetData sheetData,
        int lastDataRow,
        IReadOnlyList<string> sharedStrings,
        bool is1904)
    {
        var dummyWarnings = new List<string>();
        foreach (Row row in sheetData.Elements<Row>())
        {
            int rowIdx = checked((int)(row.RowIndex?.Value ?? 0));
            if (rowIdx < 6 || rowIdx > lastDataRow)
                continue;

            string[] cols = ExtractRowColumns(row, sharedStrings);
            if (string.IsNullOrWhiteSpace(cols[1]))
                continue;

            DailyLogEntry entry = ParseRowToEntry(rowIdx, cols, is1904, dummyWarnings);
            MasterContainerState state = DailyLogLifecycleValidator.ComputeMasterState(entry);
            SapIntakeNotification notif = DailyLogLifecycleValidator.FormatSapIntakeNotification(entry);

            OpenXmlCellHelper.SetFormulaCell(row, 14, rowIdx, BuildColOFormula(rowIdx), DailyLogCatalogs.ToExcelString(state), CellValues.String, DefaultColumnStyles[14]);
            OpenXmlCellHelper.SetFormulaCell(row, 17, rowIdx, BuildColRFormula(rowIdx), notif.ColRValue, CellValues.String, DefaultColumnStyles[17]);
        }
    }

    private static string BuildColOFormula(int rowIdx)
    {
        string r = rowIdx.ToString(CultureInfo.InvariantCulture);
        return $"IF(N{r}=\"Cierre Despacho Listo\",\"7. Cierre Total (Recep + Desp)\"," +
               $"IF(M{r}=\"Cuadrado (0 Diferencias)\",\"6. Recepción Cerrada (Listo Recep.)\"," +
               $"IF(ISNUMBER(SEARCH(\"Revisión\",L{r})),\"5. En Auditoría Control Calidad\"," +
               $"IF(M{r}=\"Pausa sin Control Calidad\",\"4. Pausa sin CC (Montado a Ciegas)\"," +
               $"IF(J{r}=\"Listo y Notificado\",\"3. Ingresado SAP (Espera Cierre)\"," +
               $"IF(J{r}=\"Pendiente Ingreso SAP\",\"2. Doc SAP Listo (Pend. Ingreso)\"," +
               $"\"1. En Cola Conteo a Ciegas\"))))))";
    }

    private static string BuildColRFormula(int rowIdx)
    {
        string r = rowIdx.ToString(CultureInfo.InvariantCulture);
        return $"IF(F{r}<>\"-\", Q{r} & \", Se informa que el ingreso a nivel de sistema del Contenedor \" & SUBSTITUTE(B{r}, \"IMP-\", \"\") & \", quedo listo.\", \"-\")";
    }

    private static void ResizeWorkbookAnchors(
        WorkbookPart wbPart,
        WorksheetPart maestroPart,
        Worksheet maestroWs,
        int lastDataRow)
    {
        string rowStr = lastDataRow.ToString(CultureInfo.InvariantCulture);
        UpdateTableAndAutoFilters(maestroPart, maestroWs, rowStr);
        UpdateFilterDataDefinedName(wbPart, rowStr);
        UpdateDataValidationsAndConditionalFormatting(maestroWs, lastDataRow);
        UpdateDashboardFormulas(wbPart, rowStr);
    }

    private static void UpdateTableAndAutoFilters(WorksheetPart maestroPart, Worksheet maestroWs, string rowStr)
    {
        foreach (TableDefinitionPart tablePart in maestroPart.TableDefinitionParts)
        {
            Table? table = tablePart.Table;
            if (table is null)
                continue;
            table.Reference = $"A5:R{rowStr}";
            if (table.AutoFilter is not null)
                table.AutoFilter.Reference = $"A5:R{rowStr}";
            table.Save();
        }

        if (maestroWs.GetFirstChild<AutoFilter>() is AutoFilter sheetFilter)
            sheetFilter.Reference = $"A5:R{rowStr}";

        foreach (CustomSheetView view in maestroWs.Descendants<CustomSheetView>())
        {
            if (view.GetFirstChild<AutoFilter>() is AutoFilter customFilter)
                customFilter.Reference = $"$A$5:$R${rowStr}";
        }
    }

    private static void UpdateFilterDataDefinedName(WorkbookPart wbPart, string rowStr)
    {
        DefinedNames? definedNames = wbPart.Workbook?.DefinedNames;
        if (definedNames is null)
            return;

        foreach (DefinedName dn in definedNames.Elements<DefinedName>())
        {
            if (!string.IsNullOrWhiteSpace(dn.Text) && FilterDataRefRegex.IsMatch(dn.Text))
                dn.Text = FilterDataRefRegex.Replace(dn.Text, $"${{1}}{rowStr}");
        }
    }

    private static void UpdateDataValidationsAndConditionalFormatting(Worksheet maestroWs, int lastDataRow)
    {
        int validationEnd = Math.Max(57, lastDataRow);
        foreach (DataValidation dv in maestroWs.Descendants<DataValidation>())
        {
            if (dv.SequenceOfReferences is not null)
                dv.SequenceOfReferences = RewriteSqrefEndRow(dv.SequenceOfReferences, validationEnd);
        }

        int cfEnd = Math.Max(1000, lastDataRow);
        foreach (ConditionalFormatting cf in maestroWs.Elements<ConditionalFormatting>())
        {
            if (cf.SequenceOfReferences is not null)
                cf.SequenceOfReferences = RewriteSqrefEndRow(cf.SequenceOfReferences, cfEnd);
        }
    }

    private static ListValue<StringValue> RewriteSqrefEndRow(ListValue<StringValue> sqref, int newEndRow)
    {
        string endStr = newEndRow.ToString(CultureInfo.InvariantCulture);
        var updated = new ListValue<StringValue>();
        foreach (StringValue item in sqref.Items)
        {
            string val = item.Value ?? "";
            string rewritten = TableRefRegex.IsMatch(val)
                ? TableRefRegex.Replace(val, $"${{1}}${{2}}:${{3}}{endStr}")
                : val;
            updated.Items.Add(new StringValue(rewritten));
        }
        return updated;
    }

    private static void UpdateDashboardFormulas(WorkbookPart wbPart, string rowStr)
    {
        WorksheetPart? dashPart = OpenXmlCellHelper.FindWorksheetPart(wbPart, "Dashboard General");
        Worksheet? dashWs = dashPart?.Worksheet;
        if (dashWs is null)
            return;

        bool modified = false;
        foreach (Cell cell in dashWs.Descendants<Cell>())
        {
            if (cell.CellFormula is null || string.IsNullOrWhiteSpace(cell.CellFormula.Text))
                continue;

            if (DashboardRangeRegex.IsMatch(cell.CellFormula.Text))
            {
                cell.CellFormula.Text = DashboardRangeRegex.Replace(cell.CellFormula.Text, $"${{1}}{rowStr}");
                OpenXmlCellHelper.ClearFormulaCacheOnly(cell);
                modified = true;
            }
        }

        if (modified)
            dashWs.Save();
    }

    private static void SaveModifiedParts(WorkbookPart wbPart, WorksheetPart maestroPart, Worksheet maestroWs)
    {
        maestroWs.Save();
        wbPart.SharedStringTablePart?.SharedStringTable?.Save();
        wbPart.Workbook?.Save();
    }

    private static bool ValidateTempWorkbook(
        string tempPath,
        OperationalSafetyGuard guard,
        int expectedLastRow,
        IReadOnlyList<int> writtenRows)
    {
        if (!IsDailyLogPackage(tempPath))
            return false;

        DailyLogWorkbookSnapshot snapshot = ReadWorkbook(tempPath, guard);
        if (snapshot.TableEndRow != expectedLastRow)
            return false;

        foreach (int rowIdx in writtenRows)
        {
            if (!snapshot.RowIndices.Contains(rowIdx))
                return false;
        }

        using WorkbookReadSession session = OpenXmlRelsSanitizer.OpenReadOnly(tempPath);
        WorksheetPart? maestroPart = session.Document.WorkbookPart is null ? null : ResolveMaestroPart(session.Document.WorkbookPart);
        if (maestroPart?.Worksheet is null)
            return false;

        return VerifyDataValidationsCoverRow(maestroPart.Worksheet, expectedLastRow);
    }

    private static bool VerifyDataValidationsCoverRow(Worksheet maestroWs, int expectedLastRow)
    {
        foreach (DataValidation dv in maestroWs.Descendants<DataValidation>())
        {
            string? firstRef = dv.SequenceOfReferences?.Items.FirstOrDefault()?.Value;
            if (string.IsNullOrWhiteSpace(firstRef))
                continue;

            Match m = TableRefRegex.Match(firstRef);
            if (!m.Success ||
                !int.TryParse(m.Groups[4].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int endRow) ||
                endRow < expectedLastRow)
            {
                return false;
            }
        }
        return true;
    }

    private static bool VerifyCommittedWorkbook(string committedPath, OperationalSafetyGuard guard, int expectedLastRow)
    {
        if (WorkbookSignatureScanner.Inspect(committedPath, guard).Kind != WorkbookKind.DailyLog)
            return false;

        DailyLogWorkbookSnapshot snapshot = ReadWorkbook(committedPath, guard);
        return snapshot.TableEndRow == expectedLastRow;
    }
}
