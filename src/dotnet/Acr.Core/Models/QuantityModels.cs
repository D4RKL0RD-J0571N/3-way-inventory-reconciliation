namespace Acr.Core.Models;

/// <summary>
/// Deterministic audit and diagnostic flag constants emitted by <see cref="Services.QuantityEngine"/>.
/// </summary>
public static class QuantityEngineFlags
{
    /// <summary>Complement units were rescued despite null or zero complement carton count.</summary>
    public const string FlagQcComplementWithoutCartonCount = "QC_COMPLEMENT_WITHOUT_CARTON_COUNT";

    /// <summary>Calculated physical units differ from the reported Excel formula total.</summary>
    public const string FlagQcFormulaMismatch = "QC_FORMULA_MISMATCH";

    /// <summary>Primary item in a shared carton ("DOS MATERIALES POR BULTO") assigned physical cartons.</summary>
    public const string FlagSharedCartonPrimary = "SHARED_CARTON_PRIMARY";

    /// <summary>Secondary item in a shared carton assigned zero physical cartons to prevent double counting.</summary>
    public const string FlagSharedCartonSecondary = "SHARED_CARTON_SECONDARY";

    /// <summary>Loose variant row ("VARIANTE SUELTA") carrying physical units without adding cartons.</summary>
    public const string FlagVarianteSueltaRow = "VARIANTE_SUELTA_ROW";

    /// <summary>Calculated units do not divide evenly by the prepack curve.</summary>
    public const string FlagUnevenPrepackCurve = "UNEVEN_PREPACK_CURVE";

    /// <summary>Reported difference in QC sheet is negative.</summary>
    public const string FlagNegativeReportedDiff = "NEGATIVE_REPORTED_DIFF";

    /// <summary>Observation text contains manual prepack adjustment keywords (e.g., MAS 1PK / MENOS 1PK).</summary>
    public const string FlagQcObservationAdjustment = "QC_OBSERVATION_ADJUSTMENT";

    /// <summary>Complement cartons were declared with zero or missing complement units.</summary>
    public const string FlagEmptyComplementCartons = "EMPTY_COMPLEMENT_CARTONS";

    /// <summary>Packaging string loose addition equals the compound product total (IMP-192 total repetition ambiguity).</summary>
    public const string FlagPackagingTotalRepetitionAmbiguity = "PACKAGING_TOTAL_REPETITION_AMBIGUITY";
}

/// <summary>
/// Represents a single compound packaging term (<c>Cartons * UnitsPerCarton</c>).
/// </summary>
/// <param name="Cartons">Number of cartons in the term.</param>
/// <param name="UnitsPerCarton">Number of units per carton in the term.</param>
public sealed record PackagingTerm(int Cartons, int UnitsPerCarton);

/// <summary>
/// Represents the parsed breakdown of a packaging specification string.
/// </summary>
public sealed record ParsedPackaging
{
    /// <summary>Normalized trimmed packaging expression string (empty if null/sentinel).</summary>
    public string RawString { get; init; } = "";

    /// <summary>Sum of all parsed carton counts (or 1 for a pure scalar unit count).</summary>
    public int TotalCartons { get; init; }

    /// <summary>Total calculated units across all terms plus any loose units.</summary>
    public int TotalUnits { get; init; }

    /// <summary>Carton count of the primary term (or 1 for a pure scalar).</summary>
    public int FullCartons { get; init; }

    /// <summary>Units per carton of the primary term (or scalar value).</summary>
    public int UnitsPerFull { get; init; }

    /// <summary>Carton count of the second term if present, otherwise zero.</summary>
    public int ComplementCartons { get; init; }

    /// <summary>Units per carton of the second term if present, otherwise zero.</summary>
    public int UnitsPerComplement { get; init; }

    /// <summary>Trailing loose units added with '+' after compound terms.</summary>
    public int LooseUnits { get; init; }

    /// <summary>True when multiple compound terms or trailing loose units are present.</summary>
    public bool IsCompound { get; init; }

    /// <summary>True when trailing loose units are greater than zero.</summary>
    public bool HasLoose { get; init; }

    /// <summary>Ordered list of all parsed compound terms for full auditability of 3+ term expressions.</summary>
    public IReadOnlyList<PackagingTerm> Terms { get; init; } = Array.Empty<PackagingTerm>();

    /// <summary>True when trailing loose units equal the sum of compound term units (IMP-192 repetition pattern).</summary>
    public bool HasTotalRepetitionAmbiguity { get; init; }
}

/// <summary>
/// Input parameters extracted from a QC spreadsheet row or manual entry.
/// </summary>
public sealed record CartonCalculationInput
{
    /// <summary>Full carton count from QC sheet.</summary>
    public int? FullCartons { get; init; }

    /// <summary>Units per full carton from QC sheet.</summary>
    public int? UnitsPerFullCarton { get; init; }

    /// <summary>Complement carton count from QC sheet.</summary>
    public int? ComplementCartons { get; init; }

    /// <summary>Units in complement cartons or loose complement units from QC sheet.</summary>
    public int? UnitsInComplement { get; init; }

    /// <summary>Total cartons reported by the QC sheet formula or entry.</summary>
    public int? ReportedTotalCartons { get; init; }

    /// <summary>Total units reported by the QC sheet formula or entry.</summary>
    public int? ReportedTotalUnits { get; init; }

    /// <summary>Optional formula-reported unit total retained for signature parity.</summary>
    public int? ReportedFormulaUnits { get; init; }

    /// <summary>Reported difference column value from the QC sheet.</summary>
    public int? ReportedDiff { get; init; }

    /// <summary>Prepack curve size (units per PRK).</summary>
    public int? Curve { get; init; }

    /// <summary>Unit of measure (<c>"PRK"</c> or <c>"UN"</c>).</summary>
    public string? Uom { get; init; } = "UN";

    /// <summary>Operator observation text from the QC sheet.</summary>
    public string? Observation { get; init; }

    /// <summary>Optional raw packaging string used as fallback when numeric columns are blank.</summary>
    public string? PackagingString { get; init; }

    /// <summary>True when the row is marked as a loose variant ("VARIANTE SUELTA").</summary>
    public bool IsVarianteSuelta { get; init; }
}

/// <summary>
/// Deterministic result of physical merchandise and packaging calculation.
/// </summary>
public sealed record CartonCalculationResult
{
    /// <summary>Total calculated physical units.</summary>
    public int CalculatedUnits { get; init; }

    /// <summary>Total calculated physical cartons.</summary>
    public int CalculatedCartons { get; init; }

    /// <summary>Calculated prepack count when curve is known, otherwise null.</summary>
    public int? CalculatedPacks { get; init; }

    /// <summary>Remainder units modulo the effective prepack curve.</summary>
    public int LooseUnits { get; init; }

    /// <summary>Echo of <see cref="CartonCalculationInput.ReportedTotalUnits"/>.</summary>
    public int? ReportedUnits { get; init; }

    /// <summary>Echo of <see cref="CartonCalculationInput.ReportedTotalCartons"/>.</summary>
    public int? ReportedCartons { get; init; }

    /// <summary>Difference between calculated units and reported total units.</summary>
    public int FormulaDiscrepancy { get; init; }

    /// <summary>Deterministic diagnostic flag codes emitted during calculation.</summary>
    public IReadOnlyList<string> Flags { get; init; } = Array.Empty<string>();

    /// <summary>Auditable Spanish breakdown explanation.</summary>
    public string Explanation { get; init; } = "";

    /// <summary>Auditable full carton count used in the calculation.</summary>
    public int FullCartonsUsed { get; init; }

    /// <summary>Auditable units per full carton used in the calculation.</summary>
    public int UnitsPerFullCartonUsed { get; init; }

    /// <summary>Auditable product of full cartons and units per full carton.</summary>
    public int FullCartonUnits { get; init; }

    /// <summary>Auditable complement carton count used in the calculation.</summary>
    public int ComplementCartonsUsed { get; init; }

    /// <summary>Auditable complementary units used in the calculation.</summary>
    public int ComplementUnits { get; init; }

    /// <summary>Auditable effective prepack curve resolved from Curve and UOM.</summary>
    public int? EffectiveCurve { get; init; }

    /// <summary>Parsed packaging expression when fallback parsing produced positive units.</summary>
    public ParsedPackaging? FallbackPackaging { get; init; }
}
