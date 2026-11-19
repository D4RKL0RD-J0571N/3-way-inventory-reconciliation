using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Xml.Linq;
using Acr.Core.Models;
using Acr.Infrastructure.Safety;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acr.Infrastructure.Excel;

/// <summary>
/// Distinguishes the origin of formula or error-token findings during closure workbook auditing.
/// </summary>
public enum FormulaErrorOrigin
{
    /// <summary>Error token was already present in the source closure template.</summary>
    PreExistingSourceError = 0,

    /// <summary>Error token was newly introduced in the exported workbook (zero allowed).</summary>
    ExporterIntroducedError = 1,

    /// <summary>Guarded optional <c>VLOOKUP</c> (<c>ISNA</c>/<c>IFNA</c>/<c>IFERROR</c>) lookup outcome.</summary>
    LegitimateOptionalNa = 2
}

/// <summary>
/// Audit record for an individual formula or error cell in <c>CIERRE RECEP</c> or <c>Bultos</c>.
/// </summary>
/// <param name="SheetName">Worksheet name where the cell resides.</param>
/// <param name="CellReference">A1 cell reference (e.g., <c>H13</c>).</param>
/// <param name="RowIndex">1-based Excel row index.</param>
/// <param name="ColumnIndex">0-based Excel column index.</param>
/// <param name="HasFormula">True when the cell defines or participates in a formula.</param>
/// <param name="IsSharedFormula">True when the formula uses <c>t="shared"</c>.</param>
/// <param name="SharedIndex">Shared formula group index (<c>si</c>) when applicable.</param>
/// <param name="FormulaText">Formula text (resolved from shared master if follower cell).</param>
/// <param name="CachedValue">Cached cell value or decoded text.</param>
/// <param name="DataType">OpenXML cell data type string.</param>
/// <param name="IsError">True when an Excel error token or <c>t="e"</c> error state is present.</param>
/// <param name="ErrorToken">Detected error token (e.g., <c>#REF!</c>, <c>#N/A</c>).</param>
/// <param name="Origin">Classified origin of the finding, when applicable.</param>
public sealed record FormulaCellAudit(
    string SheetName,
    string CellReference,
    int RowIndex,
    int ColumnIndex,
    bool HasFormula,
    bool IsSharedFormula,
    uint? SharedIndex,
    string? FormulaText,
    string? CachedValue,
    string? DataType,
    bool IsError,
    string? ErrorToken,
    FormulaErrorOrigin? Origin);

/// <summary>
/// Structural and cryptographic fingerprint of preserved workbook package parts.
/// </summary>
/// <param name="VbaProjectSha256">Uppercase SHA-256 hex digest of <c>xl/vbaProject.bin</c>, or null if absent.</param>
/// <param name="StylesXmlSha256">Uppercase SHA-256 hex digest of <c>xl/styles.xml</c>, or null if absent.</param>
/// <param name="DefinedNamesCount">Count of <c>&lt;definedName&gt;</c> entries in <c>xl/workbook.xml</c>.</param>
/// <param name="WorkbookRelsCount">Count of <c>&lt;Relationship&gt;</c> entries in <c>xl/_rels/workbook.xml.rels</c>.</param>
public sealed record WorkbookPartHashes(
    string? VbaProjectSha256,
    string? StylesXmlSha256,
    int DefinedNamesCount,
    int WorkbookRelsCount);

/// <summary>
/// Pre-print validation and 3-way formula integrity audit outcome for a closure workbook.
/// </summary>
public sealed record PrePrintValidationResult
{
    /// <summary>True when zero blocking errors and zero exporter-introduced errors were detected.</summary>
    public bool IsValid => Errors.Count == 0 && ExporterIntroducedErrors.Count == 0;

    /// <summary>Sum of physical units in column <c>AT</c> across item rows.</summary>
    public int TotalPhysicalUnits { get; init; }

    /// <summary>Sum of physical cartons in column <c>AS</c> across item rows.</summary>
    public int TotalPhysicalCartons { get; init; }

    /// <summary>Count of discrepant materials from comparison summary or validation warnings.</summary>
    public int DiscrepancyCount { get; init; }

    /// <summary>Count of unresolved or paired revuelto warnings.</summary>
    public int UnresolvedRevueltos { get; init; }

    /// <summary>All formula and error cell audits collected across <c>CIERRE RECEP</c> and <c>Bultos</c>.</summary>
    public IReadOnlyList<FormulaCellAudit> FormulaAudits { get; init; } = Array.Empty<FormulaCellAudit>();

    /// <summary>Formula or literal error cells pre-existing in the source template.</summary>
    public IReadOnlyList<FormulaCellAudit> PreExistingSourceErrors { get; init; } = Array.Empty<FormulaCellAudit>();

    /// <summary>Formula or literal error cells newly introduced by the exporter (zero allowed).</summary>
    public IReadOnlyList<FormulaCellAudit> ExporterIntroducedErrors { get; init; } = Array.Empty<FormulaCellAudit>();

    /// <summary>Guarded optional <c>VLOOKUP</c> cells classified as legitimate optional lookups.</summary>
    public IReadOnlyList<FormulaCellAudit> LegitimateOptionalNas { get; init; } = Array.Empty<FormulaCellAudit>();

    /// <summary>Non-blocking operational or pre-existing audit warnings.</summary>
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();

    /// <summary>Blocking validation or integrity errors.</summary>
    public IReadOnlyList<string> Errors { get; init; } = Array.Empty<string>();

    /// <summary>Source template package hashes when <c>sourceTemplatePath</c> is supplied.</summary>
    public WorkbookPartHashes? SourceHashes { get; init; }

    /// <summary>Output workbook package hashes.</summary>
    public WorkbookPartHashes? OutputHashes { get; init; }

    /// <summary>True when <c>xl/vbaProject.bin</c> SHA-256 matches the source template.</summary>
    public bool VbaProjectHashPreserved { get; init; } = true;

    /// <summary>True when <c>xl/styles.xml</c> SHA-256 matches the source template.</summary>
    public bool StylesPreserved { get; init; } = true;

    /// <summary>True when <c>definedNames</c> count matches the source template.</summary>
    public bool DefinedNamesPreserved { get; init; } = true;

    /// <summary>True when <c>xl/_rels/workbook.xml.rels</c> relationship count matches the source template.</summary>
    public bool RelationshipsPreserved { get; init; } = true;
}

/// <summary>
/// Direct OpenXML and ZIP-level auditor for VBA macro preservation, styles/relationships/defined-names
/// integrity, and 3-way formula error auditing across <c>CIERRE RECEP</c> and <c>Bultos</c>.
/// </summary>
public static class FormulaIntegrityAuditor
{
    private static readonly string[] ErrorTokens =
        ["#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"];

    /// <summary>
    /// Captures SHA-256 digests for <c>xl/vbaProject.bin</c> and <c>xl/styles.xml</c> plus counts of
    /// <c>definedNames</c> and <c>workbook.xml.rels</c> entries directly from the OOXML ZIP package.
    /// </summary>
    public static WorkbookPartHashes CapturePartHashes(
        string workbookPath,
        OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookPath);
        string resolved = guard is not null ? guard.EnsureCanRead(workbookPath) : Path.GetFullPath(workbookPath);
        if (!File.Exists(resolved))
            throw new FileNotFoundException($"Workbook not found for part hash capture: '{resolved}'.", resolved);

        using var stream = new FileStream(resolved, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);

        string? vbaSha = ComputeEntrySha256(archive, "xl/vbaProject.bin");
        string? stylesSha = ComputeEntrySha256(archive, "xl/styles.xml");
        int definedNamesCount = CountXmlElements(archive, "xl/workbook.xml", "definedName");
        int relsCount = CountXmlElements(archive, "xl/_rels/workbook.xml.rels", "Relationship");
        return new WorkbookPartHashes(vbaSha, stylesSha, definedNamesCount, relsCount);
    }

    /// <summary>
    /// Streams all rows of <c>CIERRE RECEP</c> and <c>Bultos</c> (when present) without row caps or
    /// trusting <c>&lt;dimension&gt;</c>, returning detailed formula and error audits.
    /// </summary>
    public static IReadOnlyList<FormulaCellAudit> ScanWorkbook(
        string workbookPath,
        OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookPath);
        string resolved = guard is not null ? guard.EnsureCanRead(workbookPath) : Path.GetFullPath(workbookPath);
        if (!File.Exists(resolved))
            throw new FileNotFoundException($"Workbook not found for formula scan: '{resolved}'.", resolved);

        using WorkbookReadSession session = OpenXmlRelsSanitizer.OpenReadOnly(resolved);
        WorkbookPart wbPart = session.Document.WorkbookPart ?? throw new InvalidDataException($"Missing WorkbookPart in '{resolved}'.");
        List<string> sharedStrings = OpenXmlCellHelper.ReadSharedStrings(wbPart);

        var audits = new List<FormulaCellAudit>();
        ScanTargetSheetIfPresent(wbPart, "CIERRE RECEP", allowContains: true, sharedStrings, audits);
        ScanTargetSheetIfPresent(wbPart, "Bultos", allowContains: false, sharedStrings, audits);
        return audits;
    }

    /// <summary>
    /// Performs a complete 3-way formula, totals, and package-preservation audit on <paramref name="outputPath"/>,
    /// optionally comparing against <paramref name="sourceTemplatePath"/> and <paramref name="expectedSummary"/>.
    /// </summary>
    public static PrePrintValidationResult ValidateClosureDocument(
        string outputPath,
        string? sourceTemplatePath = null,
        ComparisonSummary? expectedSummary = null,
        OperationalSafetyGuard? guard = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        string resolvedOut = guard is not null ? guard.EnsureCanRead(outputPath) : Path.GetFullPath(outputPath);
        if (!File.Exists(resolvedOut))
            return BuildMissingFileResult(resolvedOut);

        if (!WorkbookHasCierreRecepSheet(resolvedOut))
            return BuildMissingSheetResult();

        IReadOnlyList<FormulaCellAudit> outputAudits = ScanWorkbook(resolvedOut, guard);
        WorkbookPartHashes outputHashes = CapturePartHashes(resolvedOut, guard);
        (int physUnits, int physCartons) = SumPhysicalColumns(resolvedOut);

        return AssembleValidationResult(
            resolvedOut,
            sourceTemplatePath,
            expectedSummary,
            guard,
            outputAudits,
            outputHashes,
            physUnits,
            physCartons);
    }

    private static PrePrintValidationResult AssembleValidationResult(
        string resolvedOut,
        string? sourceTemplatePath,
        ComparisonSummary? expectedSummary,
        OperationalSafetyGuard? guard,
        IReadOnlyList<FormulaCellAudit> outputAudits,
        WorkbookPartHashes outputHashes,
        int physUnits,
        int physCartons)
    {
        var warnings = new List<string>();
        var errors = new List<string>();
        List<FormulaCellAudit> legitimateNas = outputAudits
            .Where(a => a.Origin == FormulaErrorOrigin.LegitimateOptionalNa)
            .ToList();

        (List<FormulaCellAudit> preExisting, List<FormulaCellAudit> introduced, WorkbookPartHashes? srcHashes, IReadOnlyList<string>? srcRels) =
            ClassifyErrorsAgainstSource(sourceTemplatePath, guard, outputAudits, errors, warnings);

        IReadOnlyList<string> outRels = CaptureNonTransientWorkbookRels(resolvedOut);
        (bool vbaOk, bool stylesOk, bool namesOk, bool relsOk) =
            VerifyPackagePreservation(srcHashes, outputHashes, srcRels, outRels, errors);

        AppendExpectedSummaryChecks(expectedSummary, physUnits, physCartons, warnings);

        return new PrePrintValidationResult
        {
            TotalPhysicalUnits = physUnits,
            TotalPhysicalCartons = physCartons,
            DiscrepancyCount = expectedSummary?.DiscrepantMaterialsCount ?? warnings.Count,
            UnresolvedRevueltos = expectedSummary?.RevueltoPairsCount ?? 0,
            FormulaAudits = outputAudits,
            PreExistingSourceErrors = preExisting,
            ExporterIntroducedErrors = introduced,
            LegitimateOptionalNas = legitimateNas,
            Warnings = warnings,
            Errors = errors,
            SourceHashes = srcHashes,
            OutputHashes = outputHashes,
            VbaProjectHashPreserved = vbaOk,
            StylesPreserved = stylesOk,
            DefinedNamesPreserved = namesOk,
            RelationshipsPreserved = relsOk
        };
    }

    private static (List<FormulaCellAudit> PreExisting, List<FormulaCellAudit> Introduced, WorkbookPartHashes? SourceHashes, IReadOnlyList<string>? SourceRels)
        ClassifyErrorsAgainstSource(
            string? sourceTemplatePath,
            OperationalSafetyGuard? guard,
            IReadOnlyList<FormulaCellAudit> outputAudits,
            List<string> errors,
            List<string> warnings)
    {
        List<FormulaCellAudit> outErrors = outputAudits
            .Where(a => a.IsError && a.Origin != FormulaErrorOrigin.LegitimateOptionalNa)
            .ToList();

        if (string.IsNullOrWhiteSpace(sourceTemplatePath))
        {
            var standalone = outErrors.Select(a => a with { Origin = FormulaErrorOrigin.PreExistingSourceError }).ToList();
            foreach (FormulaCellAudit err in standalone.Take(10))
                errors.Add($"Formula or literal error at {err.SheetName}!{err.CellReference}: {err.ErrorToken}");
            return (standalone, [], null, null);
        }

        string resolvedSrc = guard is not null ? guard.EnsureCanRead(sourceTemplatePath) : Path.GetFullPath(sourceTemplatePath);
        IReadOnlyList<FormulaCellAudit> srcAudits = ScanWorkbook(resolvedSrc, guard);
        WorkbookPartHashes srcHashes = CapturePartHashes(resolvedSrc, guard);
        IReadOnlyList<string> srcRels = CaptureNonTransientWorkbookRels(resolvedSrc);

        var srcErrorKeys = new HashSet<(string Sheet, string CellRef, string? Token)>(
            srcAudits
                .Where(a => a.IsError && a.Origin != FormulaErrorOrigin.LegitimateOptionalNa)
                .Select(a => (a.SheetName, a.CellReference, a.ErrorToken)));

        (List<FormulaCellAudit> pre, List<FormulaCellAudit> intro, WorkbookPartHashes hashes) =
            PartitionSourceVsIntroduced(srcAudits, outputAudits, outErrors, srcErrorKeys, srcHashes, errors, warnings);
        return (pre, intro, hashes, srcRels);
    }

    private static (List<FormulaCellAudit> PreExisting, List<FormulaCellAudit> Introduced, WorkbookPartHashes SourceHashes)
        PartitionSourceVsIntroduced(
            IReadOnlyList<FormulaCellAudit> srcAudits,
            IReadOnlyList<FormulaCellAudit> outputAudits,
            List<FormulaCellAudit> outErrors,
            HashSet<(string Sheet, string CellRef, string? Token)> srcErrorKeys,
            WorkbookPartHashes srcHashes,
            List<string> errors,
            List<string> warnings)
    {
        var preExisting = new List<FormulaCellAudit>();
        var introduced = new List<FormulaCellAudit>();

        foreach (FormulaCellAudit audit in outErrors)
        {
            if (srcErrorKeys.Contains((audit.SheetName, audit.CellReference, audit.ErrorToken)))
                preExisting.Add(audit with { Origin = FormulaErrorOrigin.PreExistingSourceError });
            else
                introduced.Add(audit with { Origin = FormulaErrorOrigin.ExporterIntroducedError });
        }

        DetectDroppedRequiredFormulas(srcAudits, outputAudits, introduced);

        if (preExisting.Count > 0)
            warnings.Add($"Detected {preExisting.Count} pre-existing source error cell(s) preserved from template.");
        foreach (FormulaCellAudit err in introduced.Take(10))
            errors.Add($"Exporter introduced error at {err.SheetName}!{err.CellReference}: {err.ErrorToken}");

        return (preExisting, introduced, srcHashes);
    }

    private static void DetectDroppedRequiredFormulas(
        IReadOnlyList<FormulaCellAudit> srcAudits,
        IReadOnlyList<FormulaCellAudit> outputAudits,
        List<FormulaCellAudit> introduced)
    {
        var outFormulaKeys = new HashSet<(string Sheet, string CellRef)>(
            outputAudits
                .Where(a => a.HasFormula)
                .Select(a => (a.SheetName.ToUpperInvariant(), a.CellReference.ToUpperInvariant())));

        var alreadyIntroduced = new HashSet<(string Sheet, string CellRef)>(
            introduced.Select(a => (a.SheetName.ToUpperInvariant(), a.CellReference.ToUpperInvariant())));

        foreach (FormulaCellAudit src in srcAudits)
        {
            if (!IsRequiredPreservedFormula(src))
                continue;

            var key = (src.SheetName.ToUpperInvariant(), src.CellReference.ToUpperInvariant());
            if (outFormulaKeys.Contains(key) || !alreadyIntroduced.Add(key))
                continue;

            introduced.Add(src with
            {
                HasFormula = false,
                IsError = true,
                ErrorToken = "#MISSING_FORMULA!",
                Origin = FormulaErrorOrigin.ExporterIntroducedError
            });
        }
    }

    private static bool IsRequiredPreservedFormula(FormulaCellAudit src)
    {
        if (!src.HasFormula)
            return false;

        if (src.SheetName.Equals("CIERRE RECEP", StringComparison.OrdinalIgnoreCase))
        {
            if (src.RowIndex is 5 or 6 && src.ColumnIndex == 1)
                return false;
            if (src.RowIndex == 7 && src.ColumnIndex == 1)
                return true;
            if (src.ColumnIndex is 7 or 46 or 51)
                return true;
            return IsTotalsAggregateFormula(src.FormulaText);
        }

        if (src.SheetName.Equals("Bultos", StringComparison.OrdinalIgnoreCase))
            return src.ColumnIndex is >= 1 and <= 9 || IsTotalsAggregateFormula(src.FormulaText);

        return false;
    }

    private static bool IsTotalsAggregateFormula(string? formulaText)
    {
        if (string.IsNullOrWhiteSpace(formulaText))
            return false;

        return formulaText.Contains("SUM(", StringComparison.OrdinalIgnoreCase) ||
               formulaText.Contains("SUBTOTAL(", StringComparison.OrdinalIgnoreCase);
    }

    private static (bool VbaOk, bool StylesOk, bool NamesOk, bool RelsOk) VerifyPackagePreservation(
        WorkbookPartHashes? srcHashes,
        WorkbookPartHashes outHashes,
        IReadOnlyList<string>? srcRels,
        IReadOnlyList<string> outRels,
        List<string> errors)
    {
        if (srcHashes is null)
            return (true, true, true, true);

        bool vbaOk = string.Equals(srcHashes.VbaProjectSha256, outHashes.VbaProjectSha256, StringComparison.Ordinal);
        bool stylesOk = string.Equals(srcHashes.StylesXmlSha256, outHashes.StylesXmlSha256, StringComparison.Ordinal);
        bool namesOk = srcHashes.DefinedNamesCount == outHashes.DefinedNamesCount;
        bool relsOk = srcRels is null || srcRels.All(r => outRels.Contains(r, StringComparer.OrdinalIgnoreCase));

        if (!vbaOk)
            errors.Add($"VBA macro binary SHA-256 mismatch: source={srcHashes.VbaProjectSha256}, output={outHashes.VbaProjectSha256}.");
        if (!stylesOk)
            errors.Add("xl/styles.xml SHA-256 changed during export.");
        if (!namesOk)
            errors.Add($"definedNames count changed: source={srcHashes.DefinedNamesCount}, output={outHashes.DefinedNamesCount}.");
        if (!relsOk)
            errors.Add("Non-transient workbook.xml.rels relationships were altered or removed.");

        return (vbaOk, stylesOk, namesOk, relsOk);
    }

    private static IReadOnlyList<string> CaptureNonTransientWorkbookRels(string workbookPath)
    {
        using var stream = new FileStream(workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        ZipArchiveEntry? entry = archive.GetEntry("xl/_rels/workbook.xml.rels") ??
            archive.Entries.FirstOrDefault(e => e.FullName.Equals("xl/_rels/workbook.xml.rels", StringComparison.OrdinalIgnoreCase));
        if (entry is null)
            return Array.Empty<string>();

        using Stream s = entry.Open();
        XDocument doc = XDocument.Load(s, LoadOptions.None);
        return doc.Descendants()
            .Where(e => e.Name.LocalName == "Relationship" &&
                        !string.IsNullOrWhiteSpace(e.Attribute("Id")?.Value) &&
                        e.Attribute("Target") is not null)
            .Select(e => $"{e.Attribute("Type")?.Value}|{e.Attribute("Target")?.Value}")
            .Where(sig => !sig.Contains("calcChain", StringComparison.OrdinalIgnoreCase) &&
                          !sig.Contains("sharedStrings", StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static void AppendExpectedSummaryChecks(
        ComparisonSummary? expectedSummary,
        int physUnits,
        int physCartons,
        List<string> warnings)
    {
        if (expectedSummary is null)
            return;

        if (physUnits != expectedSummary.TotalPhysicalUnits)
            warnings.Add($"Physical units sum in AT ({physUnits}) differs from expected ({expectedSummary.TotalPhysicalUnits}).");
        if (physCartons != expectedSummary.TotalPhysicalCartons)
            warnings.Add($"Physical cartons sum in AS ({physCartons}) differs from expected ({expectedSummary.TotalPhysicalCartons}).");
    }

    private static void ScanTargetSheetIfPresent(
        WorkbookPart wbPart,
        string sheetName,
        bool allowContains,
        IReadOnlyList<string> sharedStrings,
        List<FormulaCellAudit> audits)
    {
        WorksheetPart? wsPart = OpenXmlCellHelper.FindWorksheetPart(wbPart, sheetName, allowContains);
        if (wsPart is null)
            return;

        var sharedMasters = new Dictionary<uint, string>();
        int lastRow = 0;

        using OpenXmlReader reader = OpenXmlReader.Create(wsPart);
        while (reader.Read())
        {
            if (reader.ElementType != typeof(Row) || reader.LoadCurrentElement() is not Row rowEl)
                continue;

            int rowIdx = checked((int)(rowEl.RowIndex?.Value ?? (uint)(lastRow + 1)));
            lastRow = rowIdx;
            ScanRowCells(sheetName, rowIdx, rowEl, sharedStrings, sharedMasters, audits);
        }
    }

    private static void ScanRowCells(
        string sheetName,
        int rowIdx,
        Row rowEl,
        IReadOnlyList<string> sharedStrings,
        Dictionary<uint, string> sharedMasters,
        List<FormulaCellAudit> audits)
    {
        int fallbackCol = 0;
        foreach (Cell cell in rowEl.Elements<Cell>())
        {
            int colIdx = QcReader.ParseColumnIndex(cell.CellReference?.Value, ref fallbackCol);
            FormulaCellAudit? audit = AuditCell(sheetName, rowIdx, colIdx, cell, sharedStrings, sharedMasters);
            if (audit is not null)
                audits.Add(audit);
        }
    }

    private static FormulaCellAudit? AuditCell(
        string sheetName,
        int rowIdx,
        int colIdx,
        Cell cell,
        IReadOnlyList<string> sharedStrings,
        Dictionary<uint, string> sharedMasters)
    {
        CellFormula? formula = cell.CellFormula;
        bool hasFormula = formula is not null;
        bool isShared = formula?.FormulaType?.Value == CellFormulaValues.Shared;
        uint? sharedIndex = formula?.SharedIndex?.Value;
        string? resolvedFormula = ResolveFormulaText(formula, sharedIndex, sharedMasters);

        string cachedText = OpenXmlCellHelper.DecodeCellText(cell, sharedStrings);
        bool isErrorType = cell.DataType?.Value == CellValues.Error;
        string? token = FindCellErrorToken(resolvedFormula, cachedText, isErrorType);
        bool isOptionalVlookup = IsGuardedOptionalVlookup(resolvedFormula, token);

        if (!hasFormula && token is null && !isOptionalVlookup)
            return null;

        FormulaErrorOrigin? origin = DetermineCellOrigin(token, isOptionalVlookup);
        string cellRef = cell.CellReference?.Value ?? OpenXmlCellHelper.BuildCellRef(colIdx, rowIdx);

        return new FormulaCellAudit(
            sheetName,
            cellRef,
            rowIdx,
            colIdx,
            hasFormula,
            isShared,
            sharedIndex,
            resolvedFormula,
            cachedText.Length > 0 ? cachedText : null,
            cell.DataType?.Value.ToString(),
            token is not null,
            token,
            origin);
    }

    private static string? ResolveFormulaText(
        CellFormula? formula,
        uint? sharedIndex,
        Dictionary<uint, string> sharedMasters)
    {
        if (formula is null)
            return null;

        string? rawText = formula.Text;
        if (!string.IsNullOrWhiteSpace(rawText))
        {
            if (sharedIndex.HasValue && formula.FormulaType?.Value == CellFormulaValues.Shared)
                sharedMasters[sharedIndex.Value] = rawText;
            return rawText;
        }

        if (sharedIndex.HasValue && sharedMasters.TryGetValue(sharedIndex.Value, out string? masterText))
            return masterText;

        return rawText;
    }

    private static string? FindCellErrorToken(string? formulaText, string cachedText, bool isErrorType)
    {
        string? formulaToken = FindErrorToken(formulaText);
        if (formulaToken is not null)
            return formulaToken;

        string? valueToken = FindErrorToken(cachedText);
        if (valueToken is not null)
            return valueToken;

        return isErrorType ? "#ERROR!" : null;
    }

    private static string? FindErrorToken(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return null;

        foreach (string token in ErrorTokens)
        {
            if (text.Contains(token, StringComparison.OrdinalIgnoreCase))
                return token;
        }

        return null;
    }

    private static bool IsGuardedOptionalVlookup(string? formulaText, string? cellErrorToken)
    {
        if (string.IsNullOrWhiteSpace(formulaText))
            return false;
        if (cellErrorToken is not null && !string.Equals(cellErrorToken, "#N/A", StringComparison.OrdinalIgnoreCase))
            return false;

        string upper = formulaText.ToUpperInvariant();
        if (!upper.Contains("VLOOKUP(", StringComparison.Ordinal))
            return false;

        bool hasGuard = upper.Contains("ISNA(VLOOKUP", StringComparison.Ordinal) ||
                        upper.Contains("IFNA(VLOOKUP", StringComparison.Ordinal) ||
                        upper.Contains("IFERROR(VLOOKUP", StringComparison.Ordinal);

        return hasGuard && FindErrorToken(formulaText) is null;
    }

    private static FormulaErrorOrigin? DetermineCellOrigin(string? token, bool isOptionalVlookup)
    {
        if (isOptionalVlookup && (token is null || string.Equals(token, "#N/A", StringComparison.OrdinalIgnoreCase)))
            return FormulaErrorOrigin.LegitimateOptionalNa;
        if (token is not null)
            return FormulaErrorOrigin.PreExistingSourceError;
        return null;
    }

    private static (int Units, int Cartons) SumPhysicalColumns(string resolvedOut)
    {
        using WorkbookReadSession session = OpenXmlRelsSanitizer.OpenReadOnly(resolvedOut);
        WorkbookPart? wbPart = session.Document.WorkbookPart;
        WorksheetPart? wsPart = wbPart is null
            ? null
            : OpenXmlCellHelper.FindWorksheetPart(wbPart, "CIERRE RECEP", allowContains: true);
        if (wsPart is null)
            return (0, 0);

        List<string> sharedStrings = OpenXmlCellHelper.ReadSharedStrings(wbPart!);
        return StreamSumAsAndAt(wsPart, sharedStrings);
    }

    private static (int Units, int Cartons) StreamSumAsAndAt(WorksheetPart wsPart, IReadOnlyList<string> sharedStrings)
    {
        int totU = 0;
        int totC = 0;
        int lastRow = 0;

        using OpenXmlReader reader = OpenXmlReader.Create(wsPart);
        while (reader.Read())
        {
            if (reader.ElementType != typeof(Row) || reader.LoadCurrentElement() is not Row rowEl)
                continue;

            int rowIdx = checked((int)(rowEl.RowIndex?.Value ?? (uint)(lastRow + 1)));
            lastRow = rowIdx;
            if (rowIdx < 13)
                continue;

            AccumulateRowAsAt(rowEl, sharedStrings, ref totU, ref totC);
        }

        return (totU, totC);
    }

    private static void AccumulateRowAsAt(Row rowEl, IReadOnlyList<string> sharedStrings, ref int totU, ref int totC)
    {
        int fallbackCol = 0;
        int rowC = 0;
        int rowU = 0;
        bool skipRow = false;

        foreach (Cell cell in rowEl.Elements<Cell>())
        {
            int colIdx = QcReader.ParseColumnIndex(cell.CellReference?.Value, ref fallbackCol);
            if (colIdx != 44 && colIdx != 45)
                continue;

            if (IsAggregateFormulaCell(cell))
            {
                skipRow = true;
                break;
            }

            string decoded = OpenXmlCellHelper.DecodeCellText(cell, sharedStrings);
            if (colIdx == 44 && QcReader.ToInt(decoded) is int cVal)
                rowC = cVal;
            else if (colIdx == 45 && QcReader.ToInt(decoded) is int uVal)
                rowU = uVal;
        }

        if (!skipRow)
            checked { totC += rowC; totU += rowU; }
    }

    private static bool IsAggregateFormulaCell(Cell cell)
    {
        string? fText = cell.CellFormula?.Text;
        if (string.IsNullOrWhiteSpace(fText))
            return false;

        string trimmed = fText.TrimStart('=').TrimStart();
        return trimmed.StartsWith("SUM(", StringComparison.OrdinalIgnoreCase) ||
               trimmed.StartsWith("SUBTOTAL(", StringComparison.OrdinalIgnoreCase);
    }

    private static string? ComputeEntrySha256(ZipArchive archive, string entryPath)
    {
        ZipArchiveEntry? entry = archive.GetEntry(entryPath) ??
            archive.Entries.FirstOrDefault(e => e.FullName.Equals(entryPath, StringComparison.OrdinalIgnoreCase));
        if (entry is null)
            return null;

        using Stream s = entry.Open();
        return Convert.ToHexString(SHA256.HashData(s));
    }

    private static int CountXmlElements(ZipArchive archive, string entryPath, string localName)
    {
        ZipArchiveEntry? entry = archive.GetEntry(entryPath) ??
            archive.Entries.FirstOrDefault(e => e.FullName.Equals(entryPath, StringComparison.OrdinalIgnoreCase));
        if (entry is null)
            return 0;

        using Stream s = entry.Open();
        XDocument doc = XDocument.Load(s, LoadOptions.None);
        return doc.Descendants().Count(e => e.Name.LocalName == localName);
    }

    private static bool WorkbookHasCierreRecepSheet(string workbookPath)
    {
        try
        {
            using WorkbookReadSession session = OpenXmlRelsSanitizer.OpenReadOnly(workbookPath);
            IEnumerable<Sheet> sheets = session.Document.WorkbookPart?.Workbook?.Sheets?.Elements<Sheet>() ?? [];
            return sheets.Any(s => (s.Name?.Value ?? "").Contains("CIERRE RECEP", StringComparison.OrdinalIgnoreCase));
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or System.Xml.XmlException)
        {
            return false;
        }
    }

    private static PrePrintValidationResult BuildMissingFileResult(string resolvedOut) =>
        new()
        {
            Errors = [$"Closure file not found: '{resolvedOut}'."]
        };

    private static PrePrintValidationResult BuildMissingSheetResult() =>
        new()
        {
            Errors = ["Workbook does not contain required 'CIERRE RECEP' worksheet."]
        };
}
