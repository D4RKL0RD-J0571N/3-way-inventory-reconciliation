using System.Globalization;
using System.Text.RegularExpressions;
using Acr.Core.Models;

namespace Acr.Core.Services;

/// <summary>
/// Pure deterministic engine for parsing packaging expressions and calculating physical cartons,
/// units, and prepack curve integrity with checked integer arithmetic.
/// </summary>
public static class QuantityEngine
{
    private static readonly Regex CompoundTermRegex =
        new(@"([0-9]+)\s*[*xX]\s*([0-9]+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex LooseAdditionRegex =
        new(@"\+\s*([0-9]+)", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex WhitespaceRunRegex =
        new(@"\s+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly string[] ObservationAdjustmentKeywords =
        ["MAS 1PK", "MAS 1 PK", "+1PK", "MENOS 1PK", "MENOS 1 PK", "-1PK"];

    /// <summary>
    /// Parses a packaging specification expression (e.g., <c>"50*12"</c>, <c>"192"</c>, <c>"6*192 +48"</c>).
    /// </summary>
    public static ParsedPackaging ParsePackagingExpression(string? raw)
    {
        string trimmed = CleanRawString(raw);
        if (trimmed.Length == 0)
            return new ParsedPackaging();

        string normalized = NormalizeExpression(trimmed);
        List<PackagingTerm> terms = ParseCompoundTerms(normalized);
        int loose = ParseLooseAddition(normalized);

        if (terms.Count > 0)
            return BuildCompoundPackaging(trimmed, terms, loose);

        if (IsAllAsciiDigits(normalized))
        {
            int scalarVal = int.Parse(normalized, CultureInfo.InvariantCulture);
            return new ParsedPackaging
            {
                RawString = trimmed,
                TotalCartons = 1,
                TotalUnits = scalarVal,
                FullCartons = 1,
                UnitsPerFull = scalarVal,
                IsCompound = false
            };
        }

        return new ParsedPackaging { RawString = trimmed };
    }

    /// <summary>
    /// Genera la expresión canónica de empaque (ej. "8*48 1*36", "60", "0")
    /// reflejando con exactitud bultos enteros y bultos complementarios / picos.
    /// </summary>
    public static string FormatPackagingExpression(
        int fullCartons,
        int unitsPerFull,
        int compCartons,
        int unitsInComp,
        int curve,
        int totalUnits = 0,
        int totalCartons = 0)
    {
        // 1. Cero total / entrega diferida
        if (totalUnits == 0 && totalCartons == 0 && fullCartons == 0 && compCartons == 0 && unitsInComp == 0)
            return "0";

        // 2. Resolver unidades por bulto entero si no vinieron explícitas pero hay curva
        int uFull = unitsPerFull > 0 ? unitsPerFull : (curve > 0 ? curve : 0);

        // 3. Compuesto: Enteros + Complementos o Picos
        if (fullCartons > 0 && uFull > 0 && (compCartons > 0 || unitsInComp > 0))
        {
            int cCartons = compCartons;
            int cUnits = unitsInComp;

            if (cCartons <= 0 && cUnits > 0)
            {
                // Si Bultos Complementos es cero o vacío y Unidades Complementos > 0
                cCartons = (totalCartons > fullCartons) ? (totalCartons - fullCartons) : 1;
            }

            if (cCartons > 0 && cUnits > 0)
            {
                return $"{fullCartons}*{uFull} {cCartons}*{cUnits}";
            }
            if (cCartons > 0 && cUnits == 0)
            {
                return $"{fullCartons}*{uFull}";
            }
        }

        // 4. Solo bultos enteros sin complementos
        if (fullCartons > 0 && uFull > 0)
        {
            return uFull.ToString(CultureInfo.InvariantCulture);
        }

        // 5. Fallback por curva o unitario
        return curve > 1 ? curve.ToString(CultureInfo.InvariantCulture) : "1";
    }

    /// <summary>
    /// Calculates deterministic physical units, cartons, prepacks, loose remainders, and audit flags.
    /// </summary>
    public static CartonCalculationResult CalculateCartonQuantities(CartonCalculationInput input)
    {
        ArgumentNullException.ThrowIfNull(input);

        if (TryHandleVarianteSuelta(input) is { } varianteResult)
            return varianteResult;

        if (TryHandleSharedCarton(input) is { } sharedResult)
            return sharedResult;

        return CalculateStandardCartons(input);
    }

    private static string CleanRawString(string? raw)
    {
        if (raw is null)
            return "";

        string trimmed = raw.Trim();
        if (trimmed.Length == 0 ||
            trimmed.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            trimmed.Equals("nan", StringComparison.OrdinalIgnoreCase) ||
            trimmed.Equals("null", StringComparison.OrdinalIgnoreCase))
        {
            return "";
        }

        return trimmed;
    }

    private static string NormalizeExpression(string trimmed)
    {
        string replaced = trimmed.Replace('\u00D7', '*');
        return WhitespaceRunRegex.Replace(replaced, " ");
    }

    private static List<PackagingTerm> ParseCompoundTerms(string normalized)
    {
        MatchCollection matches = CompoundTermRegex.Matches(normalized);
        var terms = new List<PackagingTerm>(matches.Count);
        foreach (Match match in matches)
        {
            int cartons = int.Parse(match.Groups[1].ValueSpan, CultureInfo.InvariantCulture);
            int unitsPerCarton = int.Parse(match.Groups[2].ValueSpan, CultureInfo.InvariantCulture);
            terms.Add(new PackagingTerm(cartons, unitsPerCarton));
        }
        return terms;
    }

    private static int ParseLooseAddition(string normalized)
    {
        string stripped = CompoundTermRegex.Replace(normalized, "");
        Match match = LooseAdditionRegex.Match(stripped);
        return match.Success
            ? int.Parse(match.Groups[1].ValueSpan, CultureInfo.InvariantCulture)
            : 0;
    }

    private static ParsedPackaging BuildCompoundPackaging(string raw, List<PackagingTerm> terms, int loose)
    {
        checked
        {
            int totalCartons = 0;
            int termsUnits = 0;
            foreach (PackagingTerm term in terms)
            {
                totalCartons += term.Cartons;
                termsUnits += term.Cartons * term.UnitsPerCarton;
            }

            int totalUnits = termsUnits + loose;
            PackagingTerm first = terms[0];
            bool hasSecond = terms.Count > 1;
            return new ParsedPackaging
            {
                RawString = raw,
                TotalCartons = totalCartons,
                TotalUnits = totalUnits,
                FullCartons = first.Cartons,
                UnitsPerFull = first.UnitsPerCarton,
                ComplementCartons = hasSecond ? terms[1].Cartons : 0,
                UnitsPerComplement = hasSecond ? terms[1].UnitsPerCarton : 0,
                LooseUnits = loose,
                IsCompound = terms.Count > 1 || loose > 0,
                HasLoose = loose > 0,
                Terms = terms,
                HasTotalRepetitionAmbiguity = loose > 0 && loose == termsUnits
            };
        }
    }

    private static bool IsAllAsciiDigits(string value)
    {
        if (value.Length == 0)
            return false;
        foreach (char ch in value)
        {
            if (!char.IsAsciiDigit(ch))
                return false;
        }
        return true;
    }

    private static int NonZeroOr(int? primary, int? secondary, int fallback = 0)
    {
        if (primary is int p && p != 0)
            return p;
        if (secondary is int s && s != 0)
            return s;
        return fallback;
    }

    private static CartonCalculationResult? TryHandleVarianteSuelta(CartonCalculationInput input)
    {
        bool isVariante = input.IsVarianteSuelta ||
                          (input.Observation?.Contains("VARIANTE SUELTA", StringComparison.OrdinalIgnoreCase) ?? false);
        if (!isVariante)
            return null;

        int units = NonZeroOr(input.ReportedTotalUnits, input.UnitsInComplement, 0);
        return new CartonCalculationResult
        {
            CalculatedUnits = units,
            CalculatedCartons = 0,
            CalculatedPacks = units,
            LooseUnits = 0,
            ReportedUnits = input.ReportedTotalUnits,
            ReportedCartons = input.ReportedTotalCartons,
            FormulaDiscrepancy = 0,
            Flags = [QuantityEngineFlags.FlagVarianteSueltaRow],
            Explanation = "Fila de variante suelta: unidades físicas sin incrementar bultos.",
            ComplementUnits = (input.ReportedTotalUnits is null or 0) ? units : 0,
            EffectiveCurve = 1
        };
    }

    private static CartonCalculationResult? TryHandleSharedCarton(CartonCalculationInput input)
    {
        if (!(input.Observation?.Contains("DOS MATERIALES", StringComparison.OrdinalIgnoreCase) ?? false))
            return null;

        checked
        {
            bool isSecondary = input.FullCartons is null or 0;
            var flags = new List<string>
            {
                isSecondary ? QuantityEngineFlags.FlagSharedCartonSecondary : QuantityEngineFlags.FlagSharedCartonPrimary
            };

            int aVal = isSecondary ? 0 : input.FullCartons!.Value;
            int bVal = NonZeroOr(input.UnitsPerFullCarton, input.ReportedTotalUnits, 0);
            int calcU = isSecondary ? bVal : aVal * bVal;
            int calcC = isSecondary ? 0 : aVal;

            (int? packs, int loose, int? effCurve, List<string> curveFlags) =
                CheckCurveConsistency(calcU, input.Curve, input.Uom);
            flags.AddRange(curveFlags);

            int reportedOrCalc = (input.ReportedTotalUnits is int rtu && rtu != 0) ? rtu : calcU;
            return new CartonCalculationResult
            {
                CalculatedUnits = calcU,
                CalculatedCartons = calcC,
                CalculatedPacks = packs,
                LooseUnits = loose,
                ReportedUnits = input.ReportedTotalUnits,
                ReportedCartons = input.ReportedTotalCartons,
                FormulaDiscrepancy = calcU - reportedOrCalc,
                Flags = flags,
                Explanation = isSecondary
                    ? "Bulto compartido (secundario): 0 bultos físicos asignados para evitar doble conteo."
                    : $"Bulto compartido (primario): {calcC} bulto físico asignado.",
                FullCartonsUsed = calcC,
                UnitsPerFullCartonUsed = bVal,
                FullCartonUnits = calcU,
                EffectiveCurve = effCurve
            };
        }
    }

    private static CartonCalculationResult CalculateStandardCartons(CartonCalculationInput input)
    {
        checked
        {
            int a = input.FullCartons ?? 0;
            int b = input.UnitsPerFullCarton ?? 0;
            int fullUnits = a * b;

            (int compC, int compU, List<string> flags) =
                CalculateComplements(input.ComplementCartons, input.UnitsInComplement, a, input.ReportedTotalCartons);

            int totUnits = fullUnits + compU;
            int totCartons = a + compC;
            ParsedPackaging? fallbackPkg = TryApplyPackagingFallback(input, ref totUnits, ref totCartons, flags);

            (int? packs, int loose, int? effCurve, List<string> curveFlags) =
                CheckCurveConsistency(totUnits, input.Curve, input.Uom);
            flags.AddRange(curveFlags);
            AppendObservationAndDiffFlags(input, totUnits, flags, out int disc);

            return new CartonCalculationResult
            {
                CalculatedUnits = totUnits,
                CalculatedCartons = totCartons,
                CalculatedPacks = packs,
                LooseUnits = loose,
                ReportedUnits = input.ReportedTotalUnits,
                ReportedCartons = input.ReportedTotalCartons,
                FormulaDiscrepancy = disc,
                Flags = flags,
                Explanation = BuildStandardExplanation(a, b, fullUnits, compC, compU, totUnits, flags),
                FullCartonsUsed = a,
                UnitsPerFullCartonUsed = b,
                FullCartonUnits = fullUnits,
                ComplementCartonsUsed = compC,
                ComplementUnits = compU,
                EffectiveCurve = effCurve,
                FallbackPackaging = fallbackPkg
            };
        }
    }

    private static (int Cartons, int Units, List<string> Flags) CalculateComplements(
        int? c,
        int? d,
        int fullCartons,
        int? reportedTotalCartons)
    {
        checked
        {
            var flags = new List<string>();
            int compC = c ?? 0;
            int compD = d ?? 0;

            if (compC > 0 && compD > 0)
                return (compC, compC * compD, flags);

            if ((c is null || compC == 0) && compD > 0)
            {
                // Regla operativa: Bultos complementos es cero o vacío y unidades complemento > 0.
                // La operación cambia a sumar directamente compD (unidades complementos) sin multiplicar por cero.
                flags.Add(QuantityEngineFlags.FlagQcComplementWithoutCartonCount);
                if (reportedTotalCartons.HasValue && reportedTotalCartons.Value > fullCartons)
                    return (reportedTotalCartons.Value - fullCartons, compD, flags);
                if (reportedTotalCartons.HasValue && reportedTotalCartons.Value == fullCartons)
                    return (0, compD, flags);
                return (fullCartons > 0 ? 0 : 1, compD, flags);
            }

            if (compC > 0 && compD == 0)
            {
                flags.Add(QuantityEngineFlags.FlagEmptyComplementCartons);
                return (compC, 0, flags);
            }

            return (0, 0, flags);
        }
    }

    private static ParsedPackaging? TryApplyPackagingFallback(
        CartonCalculationInput input,
        ref int totUnits,
        ref int totCartons,
        List<string> flags)
    {
        if (totUnits != 0 || string.IsNullOrEmpty(input.PackagingString))
            return null;

        ParsedPackaging parsed = ParsePackagingExpression(input.PackagingString);
        if (parsed.TotalUnits <= 0)
            return null;

        totUnits = parsed.TotalUnits;
        totCartons = parsed.TotalCartons;
        if (parsed.HasTotalRepetitionAmbiguity)
            flags.Add(QuantityEngineFlags.FlagPackagingTotalRepetitionAmbiguity);

        return parsed;
    }

    private static (int? Packs, int Loose, int? EffectiveCurve, List<string> Flags) CheckCurveConsistency(
        int units,
        int? curve,
        string? uom)
    {
        var flags = new List<string>();
        int? effectiveCurve = (curve is int cv && cv != 0)
            ? cv
            : ((uom?.Equals("UN", StringComparison.OrdinalIgnoreCase) ?? false) ? 1 : null);

        if (effectiveCurve is > 1)
        {
            int divisor = effectiveCurve.Value;
            int packs = FloorDiv(units, divisor);
            int loose = checked(units - (packs * divisor));
            if (loose != 0)
                flags.Add(QuantityEngineFlags.FlagUnevenPrepackCurve);
            return (packs, loose, effectiveCurve, flags);
        }

        return (effectiveCurve == 1 ? units : null, 0, effectiveCurve, flags);
    }

    private static int FloorDiv(int dividend, int divisor)
    {
        int q = dividend / divisor;
        int r = dividend % divisor;
        if (r != 0 && ((r < 0) != (divisor < 0)))
            q--;
        return q;
    }

    private static void AppendObservationAndDiffFlags(
        CartonCalculationInput input,
        int totUnits,
        List<string> flags,
        out int disc)
    {
        if (!string.IsNullOrEmpty(input.Observation))
        {
            foreach (string keyword in ObservationAdjustmentKeywords)
            {
                if (input.Observation.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                {
                    flags.Add(QuantityEngineFlags.FlagQcObservationAdjustment);
                    break;
                }
            }
        }

        int? repU = input.ReportedTotalUnits;
        disc = repU.HasValue ? checked(totUnits - repU.Value) : 0;
        if (repU.HasValue && disc != 0)
            flags.Add(QuantityEngineFlags.FlagQcFormulaMismatch);
        if (input.ReportedDiff is < 0)
            flags.Add(QuantityEngineFlags.FlagNegativeReportedDiff);
    }

    private static string BuildStandardExplanation(
        int fullCartons,
        int unitsPerFull,
        int fullUnits,
        int compCartons,
        int compUnits,
        int totUnits,
        List<string> flags)
    {
        string baseText =
            $"{fullCartons} bultos enteros ({unitsPerFull} u/b = {fullUnits} u) + " +
            $"{compCartons} bultos complementos ({compUnits} u). Total = {totUnits} unidades.";

        if (flags.Contains(QuantityEngineFlags.FlagQcComplementWithoutCartonCount))
        {
            baseText += $" [Alerta: {compUnits} unidades complementarias rescatadas; " +
                        "en Excel se habrían multiplicado por cero].";
        }

        return baseText;
    }
}
