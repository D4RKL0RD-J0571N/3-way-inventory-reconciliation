import type { PackagingBreakdown } from '../types/reconciliation';

/**
 * Port of Acr.Core/Services/QuantityEngine.cs (and src/quantity_engine.py).
 * Enforces exact carton/complement/loose unit decomposition and canonical packaging expression formatting.
 */

export const QUANTITY_FLAGS = {
  QC_COMPLEMENT_WITHOUT_CARTON_COUNT: 'QC_COMPLEMENT_WITHOUT_CARTON_COUNT',
  QC_FORMULA_MISMATCH: 'QC_FORMULA_MISMATCH',
  SHARED_CARTON_PRIMARY: 'SHARED_CARTON_PRIMARY',
  SHARED_CARTON_SECONDARY: 'SHARED_CARTON_SECONDARY',
  VARIANTE_SUELTA_ROW: 'VARIANTE_SUELTA_ROW',
  UNEVEN_PREPACK_CURVE: 'UNEVEN_PREPACK_CURVE',
  NEGATIVE_REPORTED_DIFF: 'NEGATIVE_REPORTED_DIFF',
  QC_OBSERVATION_ADJUSTMENT: 'QC_OBSERVATION_ADJUSTMENT',
  EMPTY_COMPLEMENT_CARTONS: 'EMPTY_COMPLEMENT_CARTONS',
} as const;

export function parsePackagingExpression(val: unknown): PackagingBreakdown {
  if (val === null || val === undefined) {
    return emptyBreakdown('');
  }
  const raw = String(val).trim();
  if (!raw || ['none', 'nan', 'null'].includes(raw.toLowerCase())) {
    return emptyBreakdown('');
  }

  // Extract compound terms N*M or NxM
  const termRegex = /(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)/g;
  const terms: Array<[number, number]> = [];
  let m: RegExpExecArray | null;
  while ((m = termRegex.exec(raw)) !== null) {
    terms.push([Number(m[1]), Number(m[2])]);
  }

  // Remove N*M terms and inspect trailing loose addition +L (e.g. '6*192 +48' or '8*48 +36')
  const stripped = raw.replace(/\d+(?:\.\d+)?\s*[*xX]\s*\d+(?:\.\d+)?/g, '');
  const looseMatch = stripped.match(/\+\s*(\d+(?:\.\d+)?)/);
  const looseUnits = looseMatch ? Number(looseMatch[1]) : 0;

  if (terms.length > 0) {
    const totalBoxes = terms.reduce((acc, [b]) => acc + b, 0);
    const termsUnits = terms.reduce((acc, [b, u]) => acc + b * u, 0);
    // Disambiguate '56*12+672' where +672 repeats total vs '8*48 +36' where +36 is loose units
    const effectiveLoose = looseUnits > 0 && looseUnits === termsUnits ? 0 : looseUnits;
    const totalUnits = termsUnits + effectiveLoose;
    const boxesFull = terms[0][0];
    const unitsFull = terms[0][1];
    const boxesComp = terms.length > 1 ? terms[1][0] : 0;
    const unitsComp = terms.length > 1 ? terms[1][1] : effectiveLoose;

    return {
      total_boxes: totalBoxes,
      total_units: totalUnits,
      boxes_full: boxesFull,
      units_per_box_full: unitsFull,
      boxes_comp: boxesComp,
      units_per_box_comp: unitsComp,
      loose_units: effectiveLoose,
      raw_str: raw,
      is_compound: terms.length > 1 || effectiveLoose > 0,
      has_loose: effectiveLoose > 0,
    };
  }

  // Scalar numeric string (e.g. '42')
  if (/^-?\d+(?:\.\d+)?$/.test(raw)) {
    const f = Number(raw);
    return {
      total_boxes: f,
      total_units: 0,
      boxes_full: f,
      units_per_box_full: 0,
      boxes_comp: 0,
      units_per_box_comp: 0,
      loose_units: 0,
      raw_str: raw,
      is_compound: false,
      has_loose: false,
    };
  }

  const firstNum = raw.match(/\d+(?:\.\d+)?/);
  if (firstNum) {
    const f = Number(firstNum[0]);
    return {
      total_boxes: f,
      total_units: 0,
      boxes_full: f,
      units_per_box_full: 0,
      boxes_comp: 0,
      units_per_box_comp: 0,
      loose_units: 0,
      raw_str: raw,
      is_compound: false,
      has_loose: false,
    };
  }

  return emptyBreakdown(raw);
}

/**
 * Exact port of Acr.Core QuantityEngine.FormatPackagingExpression:
 * - Zero total => "0"
 * - Full + Physical Complement carton (B_comp > 0 && U_comp > 0) => "{Full}*{UnitsPerFull} {Comp}*{UnitsInComp}" (e.g. "8*48 1*36")
 * - Full + Loose Units (B_comp <= 0 && U_comp > 0) => "{Full}*{UnitsPerFull} +{UnitsInComp}" (e.g. "8*48 +36")
 * - Only Full cartons => units per full carton (e.g. "60") or curve
 */
export function formatPackagingExpression(params: {
  fullCartons: number;
  unitsPerFull: number;
  compCartons: number;
  unitsInComp: number;
  looseUnits?: number;
  curveTotal: number;
  isZeroTotal: boolean;
  existingEmpaqueStr?: string;
}): string {
  const {
    fullCartons,
    unitsPerFull,
    compCartons,
    unitsInComp,
    looseUnits = 0,
    curveTotal,
    isZeroTotal,
    existingEmpaqueStr,
  } = params;

  if (isZeroTotal) return '0';

  const fCart = Math.round(fullCartons);
  const uFull = Math.round(unitsPerFull);
  const cCart = Math.round(compCartons);
  const uComp = Math.round(unitsInComp || looseUnits);

  // Full + Physical Complement Carton(s): e.g. "8*48 1*36"
  if (fCart > 0 && uFull > 0 && cCart > 0 && uComp > 0) {
    return `${fCart}*${uFull} ${cCart}*${uComp}`;
  }

  // Full + Loose Units (no extra physical carton): e.g. "8*48 +36"
  if (fCart > 0 && uFull > 0 && cCart <= 0 && uComp > 0) {
    return `${fCart}*${uFull} +${uComp}`;
  }

  // Only complement carton(s) with no full cartons
  if (fCart <= 0 && cCart > 0 && uComp > 0) {
    return cCart > 1 ? `${cCart}*${uComp}` : `${uComp}`;
  }

  // Only full cartons with known units per carton
  if (fCart > 0 && uFull > 0) {
    return String(uFull);
  }

  if (existingEmpaqueStr && existingEmpaqueStr.trim() !== '') {
    return existingEmpaqueStr.trim();
  }

  return String(curveTotal > 0 ? Math.round(curveTotal) : 0);
}

export interface CartonCalcInput {
  fullCartons: number;
  unitsPerFull: number;
  compCartons: number | null;
  unitsInComp: number | null;
  reportedTotalCartons: number | null;
  reportedTotalUnits: number | null;
  reportedDiff: number | null;
  curve: number;
  uom: string;
  observations: string;
  isVarianteSuelta?: boolean;
}

export interface CartonCalcResult {
  calculatedCartons: number;
  calculatedUnits: number;
  effectiveCompCartons: number;
  effectiveCompUnits: number;
  looseUnits: number;
  flags: string[];
}

export function calculateCartonQuantities(input: CartonCalcInput): CartonCalcResult {
  const flags: string[] = [];
  const obsUpper = (input.observations || '').toUpperCase();

  // Observation adjustment flags
  if (['MAS 1PK', 'MAS 1 PK', '+1PK', 'MENOS 1PK', 'MENOS 1 PK', '-1PK'].some((k) => obsUpper.includes(k))) {
    flags.push(QUANTITY_FLAGS.QC_OBSERVATION_ADJUSTMENT);
  }

  // 1. VARIANTE SUELTA
  if (input.isVarianteSuelta || obsUpper.includes('VARIANTE SUELTA')) {
    const units = input.reportedTotalUnits || input.unitsInComp || 0;
    flags.push(QUANTITY_FLAGS.VARIANTE_SUELTA_ROW);
    return {
      calculatedCartons: 0,
      calculatedUnits: units,
      effectiveCompCartons: 0,
      effectiveCompUnits: units,
      looseUnits: units,
      flags,
    };
  }

  // 2. DOS MATERIALES (shared carton)
  if (obsUpper.includes('DOS MATERIALES')) {
    if (!input.fullCartons || input.fullCartons === 0) {
      flags.push(QUANTITY_FLAGS.SHARED_CARTON_SECONDARY);
      const u = input.reportedTotalUnits || input.unitsInComp || 0;
      return {
        calculatedCartons: 0,
        calculatedUnits: u,
        effectiveCompCartons: 0,
        effectiveCompUnits: u,
        looseUnits: u,
        flags,
      };
    }
    flags.push(QUANTITY_FLAGS.SHARED_CARTON_PRIMARY);
  }

  const a = input.fullCartons || 0;
  const b = input.unitsPerFull || 0;
  const compC = input.compCartons ?? 0;
  const compD = input.unitsInComp ?? 0;

  let effCompC = 0;
  let effCompU = 0;
  let looseUnits = 0;

  if (compC > 0 && compD > 0) {
    effCompC = compC;
    effCompU = compC * compD;
  } else if ((input.compCartons === null || compC === 0) && compD > 0) {
    // Complement rescue rule (exact Acr.Core parity — crucial for IMP-207 Layout #4!)
    flags.push(QUANTITY_FLAGS.QC_COMPLEMENT_WITHOUT_CARTON_COUNT);
    if (input.reportedTotalCartons !== null && input.reportedTotalCartons === a) {
      effCompC = 0;
      effCompU = compD;
      looseUnits = compD;
    } else if (input.reportedTotalCartons !== null && input.reportedTotalCartons === a + 1) {
      effCompC = 1;
      effCompU = compD;
    } else {
      effCompC = a > 0 ? 0 : 1;
      effCompU = compD;
      if (effCompC === 0) looseUnits = compD;
    }
  } else if (compC > 0 && compD === 0) {
    flags.push(QUANTITY_FLAGS.EMPTY_COMPLEMENT_CARTONS);
    effCompC = compC;
    effCompU = 0;
  }

  const fullUnits = a * b;
  let calcUnits = fullUnits + effCompU;
  let calcCartons = a + effCompC;

  // Fallback to reported totals if full/comp breakdown columns were blank
  if (calcUnits === 0 && input.reportedTotalUnits && input.reportedTotalUnits > 0) {
    calcUnits = input.reportedTotalUnits;
  }
  if (calcCartons === 0 && input.reportedTotalCartons && input.reportedTotalCartons > 0) {
    calcCartons = input.reportedTotalCartons;
  }

  // Curve consistency check
  const effectiveCurve = input.curve > 0 ? input.curve : input.uom.toUpperCase() === 'UN' ? 1 : 0;
  if (effectiveCurve > 1 && calcUnits > 0 && calcUnits % effectiveCurve !== 0) {
    flags.push(QUANTITY_FLAGS.UNEVEN_PREPACK_CURVE);
  }

  // Reported formula mismatch check
  if (input.reportedTotalUnits !== null && input.reportedTotalUnits > 0 && calcUnits !== input.reportedTotalUnits) {
    flags.push(QUANTITY_FLAGS.QC_FORMULA_MISMATCH);
  }
  if (input.reportedDiff !== null && input.reportedDiff < 0) {
    flags.push(QUANTITY_FLAGS.NEGATIVE_REPORTED_DIFF);
  }

  return {
    calculatedCartons: calcCartons,
    calculatedUnits: calcUnits,
    effectiveCompCartons: effCompC,
    effectiveCompUnits: effCompU,
    looseUnits,
    flags,
  };
}

function emptyBreakdown(raw: string): PackagingBreakdown {
  return {
    total_boxes: 0,
    total_units: 0,
    boxes_full: 0,
    units_per_box_full: 0,
    boxes_comp: 0,
    units_per_box_comp: 0,
    loose_units: 0,
    raw_str: raw,
    is_compound: false,
    has_loose: false,
  };
}
