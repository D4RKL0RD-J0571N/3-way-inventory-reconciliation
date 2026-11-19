"""
Service 1: Physical Quantity & Packaging Engine (Motor de Cálculo de Cantidades Físicas).

Pure deterministic Python engine for calculating physical merchandise quantities,
carton breakdowns, and prepack curve integrity from dock receiving spreadsheets.

Zero dependencies on UI, external services, or SAP networks.
Follows DC Receiving Closure operational standards.
"""

from dataclasses import dataclass, field
import re
from typing import Any, List, Optional, Tuple


@dataclass(frozen=True)
class ParsedPackaging:
    """Represents the parsed breakdown of a packaging specification string."""
    raw_string: str
    total_cartons: int = 0
    total_units: int = 0
    full_cartons: int = 0
    units_per_full: int = 0
    complement_cartons: int = 0
    units_per_complement: int = 0
    loose_units: int = 0
    is_compound: bool = False
    has_loose: bool = False


@dataclass(frozen=True)
class CartonCalculationInput:
    """Input parameters extracted from a QC spreadsheet row or manual entry."""
    full_cartons: Optional[int] = None
    units_per_full_carton: Optional[int] = None
    complement_cartons: Optional[int] = None
    units_in_complement: Optional[int] = None
    reported_total_cartons: Optional[int] = None
    reported_total_units: Optional[int] = None
    reported_formula_units: Optional[int] = None
    reported_diff: Optional[int] = None
    curve: Optional[int] = None
    uom: Optional[str] = "UN"
    observation: Optional[str] = None
    packaging_string: Optional[str] = None
    is_variante_suelta: bool = False


@dataclass(frozen=True)
class CartonCalculationResult:
    """Deterministic result of physical merchandise and packaging calculation."""
    calculated_units: int
    calculated_cartons: int
    calculated_packs: Optional[int] = None
    loose_units: int = 0
    reported_units: Optional[int] = None
    reported_cartons: Optional[int] = None
    formula_discrepancy: int = 0
    flags: List[str] = field(default_factory=list)
    explanation: str = ""


# Deterministic Flag Codes
FLAG_QC_COMPLEMENT_WITHOUT_CARTON_COUNT = "QC_COMPLEMENT_WITHOUT_CARTON_COUNT"
FLAG_QC_FORMULA_MISMATCH = "QC_FORMULA_MISMATCH"
FLAG_SHARED_CARTON_PRIMARY = "SHARED_CARTON_PRIMARY"
FLAG_SHARED_CARTON_SECONDARY = "SHARED_CARTON_SECONDARY"
FLAG_VARIANTE_SUELTA_ROW = "VARIANTE_SUELTA_ROW"
FLAG_UNEVEN_PREPACK_CURVE = "UNEVEN_PREPACK_CURVE"
FLAG_NEGATIVE_REPORTED_DIFF = "NEGATIVE_REPORTED_DIFF"
FLAG_QC_OBSERVATION_ADJUSTMENT = "QC_OBSERVATION_ADJUSTMENT"
FLAG_EMPTY_COMPLEMENT_CARTONS = "EMPTY_COMPLEMENT_CARTONS"


def _clean_str(val: Any) -> str:
    """Normalize and clean string inputs, treating nulls as empty."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("none", "nan", "null") else s


def _parse_compound_terms(expr: str) -> List[Tuple[int, int]]:
    """
    Extract (cartons, units_per_carton) terms matching 'NxM' or 'N*M'.
    Example: '3*192 3*204' -> [(3, 192), (3, 204)]
    """
    pattern = r'(\d+)\s*[*xX]\s*(\d+)'
    matches = re.findall(pattern, expr)
    return [(int(c), int(u)) for c, u in matches]


def _parse_loose_addition(expr: str) -> int:
    """
    Extract trailing loose units added with '+' (e.g. '6*192 +48' -> 48).
    Does not match if the '+' is between two NxM terms like '40*60 + 1*72'.
    """
    stripped = re.sub(r'\d+\s*[*xX]\s*\d+', '', expr)
    match = re.search(r'\+\s*(\d+)', stripped)
    return int(match.group(1)) if match else 0


def parse_packaging_expression(expr: Any) -> ParsedPackaging:
    """
    Parse packaging expressions from packing lists or QC sheets.
    Handles scalars ('192'), single term ('50*12'), compounds ('3*192 3*204'),
    and compound + loose units ('6*192 +48').
    """
    raw = _clean_str(expr)
    if not raw:
        return ParsedPackaging(raw_string="")

    terms = _parse_compound_terms(raw)
    loose = _parse_loose_addition(raw)

    if terms:
        tot_cartons = sum(t[0] for t in terms)
        tot_units = sum(t[0] * t[1] for t in terms) + loose
        full_c, full_u = terms[0]
        comp_c = terms[1][0] if len(terms) > 1 else 0
        comp_u = terms[1][1] if len(terms) > 1 else 0
        is_comp = len(terms) > 1 or loose > 0

        return ParsedPackaging(
            raw_string=raw,
            total_cartons=tot_cartons,
            total_units=tot_units,
            full_cartons=full_c,
            units_per_full=full_u,
            complement_cartons=comp_c,
            units_per_complement=comp_u,
            loose_units=loose,
            is_compound=is_comp,
            has_loose=(loose > 0)
        )

    if raw.isdigit():
        val = int(raw)
        return ParsedPackaging(
            raw_string=raw,
            total_cartons=1,
            total_units=val,
            full_cartons=1,
            units_per_full=val,
            is_compound=False
        )

    return ParsedPackaging(raw_string=raw)


def _calculate_complements(
    c: Optional[int],
    d: Optional[int],
    a: int,
    reported_total_cartons: Optional[int]
) -> Tuple[int, int, List[str]]:
    """
    Calculate complement cartons and units, handling the Excel 0-multiplier bug.
    Returns: (complement_cartons, complement_units, flags).
    """
    flags: List[str] = []
    comp_c = c or 0
    comp_d = d or 0

    if comp_c > 0 and comp_d > 0:
        return comp_c, comp_c * comp_d, flags

    if (c is None or comp_c == 0) and comp_d > 0:
        flags.append(FLAG_QC_COMPLEMENT_WITHOUT_CARTON_COUNT)
        if reported_total_cartons is not None and reported_total_cartons == a:
            return 0, comp_d, flags
        if reported_total_cartons is not None and reported_total_cartons == a + 1:
            return 1, comp_d, flags
        return (0 if a > 0 else 1), comp_d, flags

    if comp_c > 0 and comp_d == 0:
        flags.append(FLAG_EMPTY_COMPLEMENT_CARTONS)
        return comp_c, 0, flags

    return 0, 0, flags


def _check_curve_consistency(
    units: int,
    curve: Optional[int],
    uom: Optional[str]
) -> Tuple[Optional[int], int, List[str]]:
    """
    Verify prepack curve integrity.
    Returns: (packs, loose_units, flags).
    """
    flags: List[str] = []
    effective_curve = curve or (1 if uom and uom.upper() == "UN" else None)

    if effective_curve and effective_curve > 1:
        packs = units // effective_curve
        loose = units % effective_curve
        if loose != 0:
            flags.append(FLAG_UNEVEN_PREPACK_CURVE)
        return packs, loose, flags

    return (units if effective_curve == 1 else None), 0, flags


def _check_observation_flags(obs_raw: Optional[str]) -> List[str]:
    """Identify operational adjustment keywords in observation text."""
    if not obs_raw:
        return []
    obs = obs_raw.upper()
    flags: List[str] = []
    if any(k in obs for k in ("MAS 1PK", "MAS 1 PK", "+1PK", "MENOS 1PK", "MENOS 1 PK", "-1PK")):
        flags.append(FLAG_QC_OBSERVATION_ADJUSTMENT)
    return flags


def _handle_variante_suelta(inp: CartonCalculationInput) -> Optional[CartonCalculationResult]:
    """Process loose variant rows that break out curve remainders."""
    obs_text = inp.observation or ""
    if not (inp.is_variante_suelta or "VARIANTE SUELTA" in obs_text.upper()):
        return None

    units = inp.reported_total_units or inp.units_in_complement or 0
    return CartonCalculationResult(
        calculated_units=units,
        calculated_cartons=0,
        calculated_packs=units,
        loose_units=0,
        reported_units=inp.reported_total_units,
        reported_cartons=inp.reported_total_cartons,
        formula_discrepancy=0,
        flags=[FLAG_VARIANTE_SUELTA_ROW],
        explanation="Fila de variante suelta: unidades físicas sin incrementar bultos."
    )


def _handle_shared_carton(inp: CartonCalculationInput) -> Optional[CartonCalculationResult]:
    """Process rows marked as 'DOS MATERIALES POR BULTO'."""
    obs_text = inp.observation or ""
    if "DOS MATERIALES" not in obs_text.upper():
        return None

    is_sec = (inp.full_cartons is None or inp.full_cartons == 0)
    flags = [FLAG_SHARED_CARTON_SECONDARY if is_sec else FLAG_SHARED_CARTON_PRIMARY]
    a_val = 0 if is_sec else (inp.full_cartons or 0)
    b_val = inp.units_per_full_carton or (inp.reported_total_units or 0)
    calc_u = b_val if is_sec else (a_val * b_val)
    calc_c = 0 if is_sec else a_val

    packs, loose, curve_flags = _check_curve_consistency(calc_u, inp.curve, inp.uom)
    flags.extend(curve_flags)
    exp = (
        "Bulto compartido (secundario): 0 bultos físicos asignados para evitar doble conteo."
        if is_sec else
        f"Bulto compartido (primario): {calc_c} bulto físico asignado."
    )
    return CartonCalculationResult(
        calculated_units=calc_u,
        calculated_cartons=calc_c,
        calculated_packs=packs,
        loose_units=loose,
        reported_units=inp.reported_total_units,
        reported_cartons=inp.reported_total_cartons,
        formula_discrepancy=calc_u - (inp.reported_total_units or calc_u),
        flags=flags,
        explanation=exp
    )


def calculate_carton_quantities(inp: CartonCalculationInput) -> CartonCalculationResult:
    """
    Main entry point for physical carton and unit calculation.
    Deterministic, defensive, and auditable.
    """
    # 1. Check specialized row types
    vs_res = _handle_variante_suelta(inp)
    if vs_res is not None:
        return vs_res

    sc_res = _handle_shared_carton(inp)
    if sc_res is not None:
        return sc_res

    # 2. Standard Carton Calculation
    flags: List[str] = []
    a = inp.full_cartons or 0
    b = inp.units_per_full_carton or 0
    full_units = a * b

    comp_c, comp_u, comp_flags = _calculate_complements(
        inp.complement_cartons, inp.units_in_complement, a, inp.reported_total_cartons
    )
    flags.extend(comp_flags)

    tot_units = full_units + comp_u
    tot_cartons = a + comp_c

    # Fallback to packaging expression if columns were completely blank
    if tot_units == 0 and inp.packaging_string:
        parsed = parse_packaging_expression(inp.packaging_string)
        if parsed.total_units > 0:
            tot_units, tot_cartons = parsed.total_units, parsed.total_cartons

    packs, loose, curve_flags = _check_curve_consistency(tot_units, inp.curve, inp.uom)
    flags.extend(curve_flags)
    flags.extend(_check_observation_flags(inp.observation))

    rep_u = inp.reported_total_units
    disc = (tot_units - rep_u) if rep_u is not None else 0
    if rep_u is not None and disc != 0:
        flags.append(FLAG_QC_FORMULA_MISMATCH)
    if inp.reported_diff is not None and inp.reported_diff < 0:
        flags.append(FLAG_NEGATIVE_REPORTED_DIFF)

    explanation = (
        f"{a} bultos enteros ({b} u/b = {full_units} u) + "
        f"{comp_c} bultos complementos ({comp_u} u). Total = {tot_units} unidades."
    )
    if FLAG_QC_COMPLEMENT_WITHOUT_CARTON_COUNT in flags:
        explanation += (
            f" [Alerta: {comp_u} unidades complementarias rescatadas; "
            "en Excel se habrían multiplicado por cero]."
        )

    return CartonCalculationResult(
        calculated_units=tot_units,
        calculated_cartons=tot_cartons,
        calculated_packs=packs,
        loose_units=loose,
        reported_units=rep_u,
        reported_cartons=inp.reported_total_cartons,
        formula_discrepancy=disc,
        flags=flags,
        explanation=explanation
    )
