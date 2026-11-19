import * as XLSX from 'xlsx';
import type {
  AuditData,
  BomEntry,
  ChildSKU,
  CierreMetadata,
  PrepackItem,
  QcDetail,
} from '../types/reconciliation';
import { cleanInt, cleanMaterialCode, cleanNumber, cleanString } from './sanitizer';
import { computeSummary } from './engine';
import { validateAllItems } from './validator';
import {
  calculateCartonQuantities,
  formatPackagingExpression,
  parsePackagingExpression,
} from './quantityEngine';
import { classifyAllDiscrepancies, isItemZeroTotal } from './discrepancyClassifier';
import { computeVbaSha256 } from './macroSafeXlsmExporter';

export { parsePackagingExpression as parsePackagingString };

const IMPORT_ID_EXCLUDE_TOKENS = [
  'importadora',
  'faltante',
  'dañad',
  'danad',
  'sobrant',
  'revuelt',
  'total',
  'bulto',
  'valor',
  'costo',
  'unidad',
  'fecha',
  'diferencia',
  'ncr',
  'reclamo',
  'averi',
  'moneda',
  'monto',
  'cantidad',
  'estado',
];

function getSheetMatrix(wb: XLSX.WorkBook, sheetName: string): unknown[][] {
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  return XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: true, defval: null });
}

function getCell(matrix: unknown[][], row1: number, col1: number): unknown {
  const r = matrix[row1 - 1];
  if (!r) return null;
  return r[col1 - 1] ?? null;
}

function hasCierreSheet(wb: XLSX.WorkBook): boolean {
  return wb.SheetNames.some((s) => s.toUpperCase().includes('CIERRE RECEP'));
}

function findQcSheetName(wb: XLSX.WorkBook): string | undefined {
  return wb.SheetNames.find((s) => {
    const u = s.toUpperCase();
    return (
      u.includes('CONTROL DATA') ||
      u.includes('CONTROL_DATA') ||
      u.includes('CONTROL CALIDAD') ||
      u === 'CONTROL' ||
      u === 'QC'
    );
  });
}

function extractImportIdFromFileName(name?: string): string | null {
  if (!name) return null;
  const m = name.match(/\b(IMP)[\s\-_]*([0-9]+[A-Z0-9\-]*)(?=[_\s.\-]|$|\b)/i);
  if (m) {
    return `IMP-${m[2].toUpperCase()}`;
  }
  return null;
}

function normalizeImportIdValue(raw: string): string {
  const trimmed = raw.trim();
  if (!trimmed) return '';
  const impMatch = trimmed.match(/^IMP[\s\-_]*(\d+[A-Z0-9\-]*)$/i);
  if (impMatch) {
    return `IMP-${impMatch[1].toUpperCase()}`;
  }
  return trimmed;
}

function isPureNumeric(val: string): boolean {
  return /^-?\d+(?:\.\d+)?$/.test(val.trim());
}

/**
 * Extracts header metadata from DATOS GENERALES and CIERRE RECEP.
 * Hardened against the IMP-207 -> 288 collision bug where Row 22 ("Faltante Importación")
 * previously matched `lblLow.includes('importaci')` and overwrote "IMP-207" with "288".
 */
export function extractCierreMetadata(
  wb: XLSX.WorkBook,
  fileNames?: { cierreFileName?: string; qcFileName?: string }
): CierreMetadata {
  const meta: CierreMetadata = {
    import_id: 'IMPORT-ACTIVA',
    vendor_name: 'PROVEEDOR INTERNACIONAL',
    invoice_number: 'PENDIENTE',
    theor_packages: 0,
    received_packages: 0,
    difference_packages: 0,
  };

  let importIdLocked = false;

  const dgSheetName = wb.SheetNames.find((s) => s.toUpperCase().includes('DATOS GENERALES'));
  if (dgSheetName) {
    const dg = getSheetMatrix(wb, dgSheetName);
    for (let r = 1; r <= Math.min(30, dg.length); r++) {
      // Scan label/value pairs across columns A:B, B:C, C:D
      for (let labelCol = 1; labelCol <= 3; labelCol++) {
        const lbl = cleanString(getCell(dg, r, labelCol));
        if (!lbl) continue;
        const lblLow = lbl.toLowerCase().replace(/\s+/g, ' ').trim();
        const val = getCell(dg, r, labelCol + 1);
        const cVal = cleanString(val);

        // 1. Import ID (strictly excluding Discrepancy/Summary rows like "Faltante Importación" = 288!)
        const isImportLabel =
          (lblLow.includes('importaci') ||
            lblLow.includes('contenedor') ||
            lblLow === 'imp' ||
            lblLow.startsWith('no. imp') ||
            lblLow.startsWith('# imp')) &&
          !IMPORT_ID_EXCLUDE_TOKENS.some((tok) => lblLow.includes(tok));

        if (isImportLabel && cVal && cVal !== '0') {
          const normalized = normalizeImportIdValue(cVal);
          const candidateHasImp = /^IMP-/i.test(normalized);
          if (!importIdLocked || (candidateHasImp && !/^IMP-/i.test(meta.import_id))) {
            meta.import_id = normalized;
            importIdLocked = true;
          }
          continue;
        }

        // 2. Invoice / Nacionalización (Doc 41)
        if (
          (lblLow.includes('nacionalizaci') ||
            lblLow.includes('pedido nac') ||
            lblLow.includes('ped/nac') ||
            (lblLow.includes('factura') && !lblLow.includes('valor') && !lblLow.includes('total'))) &&
          cVal &&
          cVal !== '0' &&
          meta.invoice_number === 'PENDIENTE'
        ) {
          meta.invoice_number = cVal;
          continue;
        }

        // 3. Vendor name
        if (
          lblLow.includes('proveedor') &&
          !lblLow.includes('mat.') &&
          !lblLow.includes('codigo') &&
          !lblLow.includes('código') &&
          cVal &&
          meta.vendor_name === 'PROVEEDOR INTERNACIONAL'
        ) {
          meta.vendor_name = cVal;
          continue;
        }

        // 4. Package counts
        if (lblLow.includes('diferencia') && lblLow.includes('bulto')) {
          meta.difference_packages = cleanNumber(val);
        } else if (lblLow.includes('bultos recib')) {
          meta.received_packages = cleanNumber(val);
        } else if (lblLow.includes('bultos te')) {
          meta.theor_packages = cleanNumber(val);
        }
      }
    }
  }

  const cierreSheet = wb.SheetNames.find((s) => s.toUpperCase().includes('CIERRE RECEP'));
  if (cierreSheet) {
    const cr = getSheetMatrix(wb, cierreSheet);
    for (let r = 1; r <= Math.min(10, cr.length); r++) {
      for (let labelCol = 1; labelCol <= 3; labelCol++) {
        const lbl = cleanString(getCell(cr, r, labelCol));
        if (!lbl) continue;
        const lblLow = lbl.toLowerCase().replace(/\s+/g, ' ').trim();
        const val = getCell(cr, r, labelCol + 1);
        const cVal = cleanString(val);

        const isHeaderImportLabel =
          (lblLow.includes('contenedor') || lblLow.includes('importaci')) &&
          !IMPORT_ID_EXCLUDE_TOKENS.some((tok) => lblLow.includes(tok));

        if (isHeaderImportLabel && cVal && cVal !== '0') {
          const normalized = normalizeImportIdValue(cVal);
          if (
            meta.import_id === 'IMPORT-ACTIVA' ||
            (isPureNumeric(meta.import_id) && /^IMP-/i.test(normalized))
          ) {
            meta.import_id = normalized;
          }
        } else if (
          (lblLow.includes('ped/nac') || lblLow.includes('nacionalizaci')) &&
          meta.invoice_number === 'PENDIENTE' &&
          cVal
        ) {
          meta.invoice_number = cVal;
        } else if (lblLow.includes('bultos te') && meta.theor_packages === 0) {
          meta.theor_packages = cleanNumber(val);
        } else if (lblLow.includes('bultos recib') && meta.received_packages === 0) {
          meta.received_packages = cleanNumber(val);
        } else if (lblLow.includes('dif bultos') && meta.difference_packages === 0) {
          meta.difference_packages = cleanNumber(val);
        }
      }
    }

    // Fallback package sums with NO 500-row cap (supports IMP-187 row 612+)
    if (meta.received_packages === 0 || meta.theor_packages === 0) {
      let recSum = 0;
      let theoSum = 0;
      for (let r = 13; r <= cr.length; r++) {
        const matV = cleanMaterialCode(getCell(cr, r, 5));
        if (!matV) continue;
        recSum += parsePackagingExpression(getCell(cr, r, 45)).total_boxes;
        theoSum += parsePackagingExpression(getCell(cr, r, 49)).total_boxes;
      }
      if (meta.received_packages === 0 && recSum > 0) meta.received_packages = recSum;
      if (meta.theor_packages === 0 && theoSum > 0) meta.theor_packages = theoSum;
      if (meta.difference_packages === 0) meta.difference_packages = meta.received_packages - meta.theor_packages;
    }
  }

  // Fallback to filename if import_id is still placeholder or suspiciously pure-numeric without IMP prefix
  const fileFallback =
    extractImportIdFromFileName(fileNames?.cierreFileName) ||
    extractImportIdFromFileName(fileNames?.qcFileName);

  if (fileFallback) {
    if (meta.import_id === 'IMPORT-ACTIVA' || isPureNumeric(meta.import_id)) {
      meta.import_id = fileFallback;
    }
  } else if (isPureNumeric(meta.import_id) && meta.import_id !== '0') {
    meta.import_id = `IMP-${meta.import_id}`;
  }

  return meta;
}

function isSummaryFooterCode(matCode: string): boolean {
  const u = matCode.trim().toUpperCase();
  if (!u) return true;
  return /^(TOTAL|SUBTOTAL|SUMA|RESUMEN|FIRMA|GRAN TOTAL|OBSERVACION)/.test(u);
}

/**
 * Port of Acr.Infrastructure/Excel/QcReader.cs.
 * - Dynamic header row scan with collision guards (commit 14240a8 for IMP-207 Layout #4).
 * - Never filters valid product descriptions containing protected words like "LAPIZ COLOR SUMA".
 * - Ignores "ESQUEMA ALISTO" sheets when detecting multi-part containers.
 * - Accumulates duplicate material rows (split pallets) instead of overwriting.
 */
export function parseQualityControlWorkbook(wbOrBuffer: ArrayBuffer | XLSX.WorkBook): {
  boms: Record<string, BomEntry[]>;
  qcDetails: Record<string, QcDetail>;
  detectedParts: string[];
  warnings: string[];
} {
  const wb = wbOrBuffer instanceof ArrayBuffer ? XLSX.read(wbOrBuffer, { type: 'array' }) : wbOrBuffer;
  const warnings: string[] = [];

  // 1. Parse BOM from BASE DE DATOS
  const bdSheetName = wb.SheetNames.find((s) =>
    ['BASE DE DATOS', 'BASE_DATOS', 'BOM', 'CURVAS'].some((k) => s.toUpperCase().includes(k))
  );
  const boms: Record<string, BomEntry[]> = {};
  if (bdSheetName) {
    const rows = getSheetMatrix(wb, bdSheetName);
    let currentBomId: string | null = null;
    let currentParent: string | null = null;
    let emptyStreak = 0;

    for (let i = 4; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length < 5) {
        emptyStreak++;
        if (emptyStreak > 35) break;
        continue;
      }
      const bomId = cleanString(row[4]);
      const matVal = row[8] ?? row[5];
      const matClean = cleanMaterialCode(matVal);
      const ratioVal = row[9] ?? row[6];
      const umVal = cleanString(row[10]).toUpperCase();

      if (!bomId && !matClean) {
        emptyStreak++;
        if (emptyStreak > 35) break;
        continue;
      }
      emptyStreak = 0;

      if (bomId && matClean) {
        if (bomId !== currentBomId || umVal === 'PRK') {
          currentBomId = bomId;
          currentParent = matClean;
          if (!boms[currentParent]) boms[currentParent] = [];
        } else {
          const ratio = cleanInt(ratioVal, 1);
          if (currentParent && matClean !== currentParent) {
            boms[currentParent].push({ child: matClean, ratio });
          }
        }
      }
    }
  }

  // 2. Dynamic header row & column detection in CONTROL DATA
  const qcSheetName = findQcSheetName(wb) ?? wb.SheetNames[0];
  const qcRows = getSheetMatrix(wb, qcSheetName);

  let headerRowIdx = 5; // Default Row 6 (0-indexed 5)
  for (let r = 0; r < Math.min(15, qcRows.length); r++) {
    const candidate = qcRows[r];
    if (!candidate) continue;
    const normCells = candidate.map((c) => cleanString(c).toUpperCase());
    const hasMat = normCells.some((c) => c === 'MATERIAL' || (c.includes('MATERIAL') && !c.includes('PROVEEDOR') && !c.includes('DIFERENCIA')));
    const hasMetric = normCells.some((c) => c.includes('BULTOS') || c.includes('UNIDADES') || c.includes('CURVA'));
    if (hasMat && hasMetric) {
      headerRowIdx = r;
      break;
    }
  }

  let colMat = 3;
  let colDesc = 11;
  let colUm = 10;
  let colQty = 4;
  let colCurva = 34;
  let colBultos = 31;
  let colUnidades = 32;
  let colBultosEnteros = 27;
  let colUnidEnteros = 28;
  let colBultosComp: number | null = null;
  let colUnidComp: number | null = null;
  let colTotalPrk = 33;
  let colDiffUnits: number | null = null;
  let colObs: number | null = null;

  const headerRow = qcRows[headerRowIdx];
  let descFound = false;
  if (headerRow) {
    headerRow.forEach((val, cIdx) => {
      const sVal = cleanString(val);
      if (!sVal) return;
      const vLow = sVal.toLowerCase().replace(/\s+/g, ' ').trim();

      // Material (exclude PROVEEDOR & DIFERENCIA)
      if (
        (vLow === 'material' || (vLow.includes('material') && !vLow.includes('proveedor') && !vLow.includes('diferencia'))) &&
        cIdx < 15
      ) {
        colMat = cIdx;
      } else if (vLow.includes('descripci') && !descFound) {
        colDesc = cIdx;
        descFound = true;
      } else if (vLow === 'um' || vLow === 'u.m.' || vLow === 'u/m') {
        colUm = cIdx;
      } else if (vLow.includes('bultos enteros') || vLow.includes('bulto entero')) {
        colBultosEnteros = cIdx;
      } else if (vLow.includes('unidades por bulto') || vLow.includes('unidades por bultos') || vLow.includes('unid. por bulto')) {
        colUnidEnteros = cIdx;
      } else if (
        // CRITICAL (commit 14240a8 / IMP-207 Layout #4):
        // comp_cartons MUST exclude any header containing "unidades" or "unid"
        (vLow.includes('bultos complemento') || vLow.includes('bulto complemento')) &&
        !vLow.includes('unidad') &&
        !vLow.includes('unid')
      ) {
        colBultosComp = cIdx;
      } else if (
        (vLow.includes('unidades bultos complemento') ||
          vLow.includes('unidades complemento') ||
          vLow.includes('unid') && vLow.includes('complemento'))
      ) {
        colUnidComp = cIdx;
      } else if (vLow.includes('total de bultos') || vLow === 'total bultos') {
        colBultos = cIdx;
      } else if (vLow.includes('total de unidades') || vLow === 'total unidades') {
        colUnidades = cIdx;
      } else if (vLow.includes('total prk') || vLow.includes('total prepack')) {
        colTotalPrk = cIdx;
      } else if (vLow.includes('curva') && cIdx > 15) {
        colCurva = cIdx;
      } else if (vLow.includes('cantidad de unidades') || vLow === 'cantidad') {
        colQty = cIdx;
      } else if (
        // Collision guard: diff_units MUST ignore "MATERIAL" ("Diferencia en Material SAP")
        vLow.includes('diferencia') &&
        !vLow.includes('material') &&
        !vLow.includes('bulto')
      ) {
        colDiffUnits = cIdx;
      } else if (vLow.includes('observaci') || vLow.includes('comentario')) {
        colObs = cIdx;
      }
    });
  }

  // Detect Layout #4 (deleted BULTOS COMPLEMENTOS column)
  if (colBultosComp === null && colUnidComp !== null) {
    warnings.push(
      'Detectado Formato QC Layout #4 (sin columna "Bultos Complementos"): se activó la regla de rescate matemático de complementos/sueltas.'
    );
  }

  // 3. Detect multi-part distribution sheets — EXCLUDING "ESQUEMA ALISTO" (prevents 21 false PENDIENTE_PARTE_B on IMP-207!)
  const partMaterials: Record<string, string> = {};
  const detectedPartsSet = new Set<string>();
  for (const sname of wb.SheetNames) {
    const sUpper = sname.toUpperCase();
    if (sUpper.includes('ESQUEMA') || sUpper.includes('ALISTO')) {
      continue;
    }
    const mPart = sname.match(/PARTE\s+([A-Z0-9]+)/i);
    if (mPart) {
      const pCode = mPart[1].toUpperCase();
      detectedPartsSet.add(pCode);
      const pRows = getSheetMatrix(wb, sname);
      for (let r = 0; r < pRows.length; r++) {
        const rowP = pRows[r];
        if (!rowP) continue;
        for (let c = 0; c < Math.min(8, rowP.length); c++) {
          const valS = String(rowP[c] ?? '').trim();
          if (/^\d{7,}$/.test(valS)) {
            const cleanM = cleanMaterialCode(valS);
            if (cleanM && !partMaterials[cleanM]) {
              partMaterials[cleanM] = pCode;
            }
          }
        }
      }
    }
  }

  const detectedParts = Array.from(detectedPartsSet).sort();
  const qcDetails: Record<string, QcDetail> = {};
  let emptyStreak = 0;

  for (let i = headerRowIdx + 1; i < qcRows.length; i++) {
    const row = qcRows[i];
    if (!row || colMat >= row.length) {
      emptyStreak++;
      if (emptyStreak > 40) break;
      continue;
    }

    const rawMat = row[colMat];
    const rawMatStr = cleanString(rawMat);
    if (!rawMatStr || isSummaryFooterCode(rawMatStr)) {
      emptyStreak++;
      if (emptyStreak > 40) break;
      continue;
    }

    const mClean = cleanMaterialCode(rawMat);
    if (!mClean || isSummaryFooterCode(mClean)) {
      emptyStreak++;
      if (emptyStreak > 40) break;
      continue;
    }
    emptyStreak = 0;

    const desc = cleanString(row[colDesc]);
    const um = cleanString(row[colUm]).toUpperCase() || 'PRK';
    const qty = cleanNumber(row[colQty]);
    const curve = cleanInt(row[colCurva], um === 'UN' ? 1 : 12);
    const fullCartons = cleanNumber(row[colBultosEnteros]);
    const unitsPerFull = cleanNumber(row[colUnidEnteros]);
    const rawCompCartons = colBultosComp !== null ? row[colBultosComp] : null;
    const compCartons =
      colBultosComp !== null && rawCompCartons !== null && cleanString(rawCompCartons) !== ''
        ? cleanNumber(rawCompCartons)
        : null;
    const rawUnitsComp = colUnidComp !== null ? row[colUnidComp] : null;
    const unitsInComp =
      colUnidComp !== null && rawUnitsComp !== null && cleanString(rawUnitsComp) !== ''
        ? cleanNumber(rawUnitsComp)
        : null;
    const reportedBultos = row[colBultos] !== null && row[colBultos] !== undefined ? cleanNumber(row[colBultos]) : null;
    const reportedUnidades =
      row[colUnidades] !== null && row[colUnidades] !== undefined ? cleanNumber(row[colUnidades]) : null;
    const reportedDiff =
      colDiffUnits !== null && row[colDiffUnits] !== null && row[colDiffUnits] !== undefined
        ? cleanNumber(row[colDiffUnits])
        : null;
    const totalPrk = cleanNumber(row[colTotalPrk]);
    const obsText = colObs !== null ? cleanString(row[colObs]) : '';

    const calc = calculateCartonQuantities({
      fullCartons,
      unitsPerFull,
      compCartons,
      unitsInComp,
      reportedTotalCartons: reportedBultos,
      reportedTotalUnits: reportedUnidades,
      reportedDiff,
      curve,
      uom: um,
      observations: obsText,
      isVarianteSuelta: desc.toUpperCase().includes('VARIANTE SUELTA'),
    });

    const assignedPart = partMaterials[mClean] ?? (detectedParts.length > 0 ? 'A' : '');
    const existing = qcDetails[mClean];

    if (existing) {
      // Accumulate split-pallet duplicate rows cleanly instead of overwriting
      existing.qty += qty;
      existing.bultos += calc.calculatedCartons;
      existing.unidades += calc.calculatedUnits;
      existing.bultos_enteros += fullCartons;
      if (existing.unidades_por_bulto_entero === 0 && unitsPerFull > 0) {
        existing.unidades_por_bulto_entero = unitsPerFull;
      }
      existing.bultos_complementos += calc.effectiveCompCartons;
      existing.unidades_por_bulto_comp += calc.effectiveCompUnits;
      existing.loose_units += calc.looseUnits;
      existing.total_prk += totalPrk;
      if (obsText && !existing.observations.includes(obsText)) {
        existing.observations = existing.observations ? `${existing.observations} | ${obsText}` : obsText;
      }
      for (const fl of calc.flags) {
        if (!existing.flags.includes(fl)) existing.flags.push(fl);
      }
    } else {
      qcDetails[mClean] = {
        desc,
        um,
        qty,
        curve,
        bultos: calc.calculatedCartons,
        unidades: calc.calculatedUnits,
        bultos_enteros: fullCartons,
        unidades_por_bulto_entero: unitsPerFull,
        bultos_complementos: calc.effectiveCompCartons,
        unidades_por_bulto_comp: calc.effectiveCompUnits,
        loose_units: calc.looseUnits,
        total_prk: totalPrk,
        part: assignedPart,
        observations: obsText,
        flags: [...calc.flags],
      };
    }
  }

  return { boms, qcDetails, detectedParts, warnings };
}

/**
 * Largest-remainder proportional allocation of Child SKUs so sum(received_units) === parentReceived.
 */
function allocateChildSkus(
  bomEntries: BomEntry[],
  parentInvoicedUnits: number,
  parentReceivedUnits: number,
  ped41Prk: number,
  curve: number,
  parentDesc: string,
  unitCost: number,
  qcDetails: Record<string, QcDetail>
): ChildSKU[] {
  if (!bomEntries.length) return [];
  const totalRatio = bomEntries.reduce((s, b) => s + Math.max(0, b.ratio), 0) || curve || 1;

  const rawRecAlloc = bomEntries.map((b) => {
    const exact = (parentReceivedUnits * b.ratio) / totalRatio;
    return {
      entry: b,
      floor: Math.floor(exact),
      frac: exact - Math.floor(exact),
    };
  });

  let remainder = Math.max(0, Math.round(parentReceivedUnits) - rawRecAlloc.reduce((s, r) => s + r.floor, 0));
  const sortedByFrac = [...rawRecAlloc].sort((a, b) => b.frac - a.frac);
  for (let i = 0; i < sortedByFrac.length && remainder > 0; i++) {
    sortedByFrac[i].floor += 1;
    remainder -= 1;
  }

  return rawRecAlloc.map(({ entry: b, floor: cRec }) => {
    const cTheo =
      ped41Prk > 0
        ? Math.round(ped41Prk * b.ratio)
        : totalRatio > 0
          ? Math.round((parentInvoicedUnits / totalRatio) * b.ratio)
          : 0;
    const cDesc = qcDetails[b.child]?.desc || `${parentDesc} (Talla ${b.ratio})`;
    return {
      material_code: b.child,
      description: cDesc,
      ratio: b.ratio,
      theoretical_units: cTheo,
      received_units: cRec,
      difference: cRec - cTheo,
      unit_cost: unitCost,
    };
  });
}

export async function runWorkbookAudit(
  qcBufferInput: ArrayBuffer | null,
  cierreBufferInput: ArrayBuffer,
  options: { qcFileName?: string; cierreFileName?: string; isBlindSetup?: boolean } = {}
): Promise<AuditData> {
  const tStart = performance.now();
  const warnings: string[] = [];

  let cierreBuffer = cierreBufferInput;
  let qcBuffer = qcBufferInput;
  let cierreFileName = options.cierreFileName ?? 'Cierre_Oficial.xlsm';
  let qcFileName = options.qcFileName ?? (options.isBlindSetup || !qcBuffer ? 'Sin QC (Modo Pausa)' : 'Control_Calidad.xlsm');

  let cierreWb = XLSX.read(cierreBuffer, { type: 'array' });
  let qcWb = qcBuffer ? XLSX.read(qcBuffer, { type: 'array' }) : null;

  // Auto-detect swapped workbooks by worksheet signature (CONTROL DATA vs CIERRE RECEP)
  if (!hasCierreSheet(cierreWb) && qcWb && hasCierreSheet(qcWb)) {
    const tmpWb = cierreWb;
    cierreWb = qcWb;
    qcWb = tmpWb;

    const tmpBuf = cierreBuffer;
    cierreBuffer = qcBuffer!;
    qcBuffer = tmpBuf;

    const tmpName = cierreFileName;
    cierreFileName = qcFileName;
    qcFileName = tmpName;

    warnings.push(
      'Detección automática de firma de hojas: los archivos de Control de Calidad y Cierre Recepción estaban intercambiados y fueron reasignados automáticamente.'
    );
  }

  if (!hasCierreSheet(cierreWb)) {
    throw new Error(
      "No se encontró la hoja 'CIERRE RECEP' en la plantilla de cierre. Verifica que hayas cargado el libro oficial de recepción."
    );
  }

  // Support hybrid workbooks (e.g. IMP-218 containing both CIERRE RECEP and CONTROL DATA)
  let isBlindSetup = Boolean(options.isBlindSetup);
  if (!qcWb && !isBlindSetup) {
    if (findQcSheetName(cierreWb)) {
      qcWb = cierreWb;
      qcFileName = `${cierreFileName} (Libro Híbrido)`;
      warnings.push(
        'Libro híbrido detectado: se encontró la hoja CONTROL DATA dentro de la misma plantilla de cierre.'
      );
    } else {
      isBlindSetup = true;
    }
  }

  let boms: Record<string, BomEntry[]> = {};
  let qcDetails: Record<string, QcDetail> = {};
  let detectedParts: string[] = [];

  if (!isBlindSetup && qcWb) {
    const parsedQc = parseQualityControlWorkbook(qcWb);
    boms = parsedQc.boms;
    qcDetails = parsedQc.qcDetails;
    detectedParts = parsedQc.detectedParts;
    warnings.push(...parsedQc.warnings);
  }

  const meta = extractCierreMetadata(cierreWb, { cierreFileName, qcFileName });
  const cierreSheetName = cierreWb.SheetNames.find((s) => s.toUpperCase().includes('CIERRE RECEP'))!;
  const matrix = getSheetMatrix(cierreWb, cierreSheetName);

  // Default 1-indexed column positions matching CIERRE RECEP 58-column corporate layout (A-BF)
  const colMap: Record<string, number> = {
    po_43: 1,
    cat: 3,
    vendor_mat: 4,
    mat: 5,
    ped43_u: 6,
    ped41_u: 7,
    rec_u: 8,
    mov101: 9,
    rev_sal: 10,
    rev_dest: 11,
    rev_ent: 12,
    rev_orig: 13,
    falt: 14,
    dan: 15,
    sob: 16,
    mov641: 21,
    partes_bcd: 24,
    desc: 30,
    unit_cost: 32,
    as_bultos: 45,
    at_unidades: 46,
    aw_bultos: 49,
    ax_unidades: 50,
    ay_empaque: 51,
    curve: 54,
    um: 55,
    ped43_prk: 56,
    ped41_prk: 57,
  };

  // Dynamic header scan on Row 12
  const headerRow12 = matrix[11] ?? [];
  for (let c = 1; c <= Math.min(headerRow12.length, 65); c++) {
    const hdr = cleanString(headerRow12[c - 1]);
    if (!hdr) continue;
    const hLow = hdr.toLowerCase().replace(/\s+/g, ' ').trim();
    if (hLow === 'material') colMap.mat = c;
    else if (hLow.includes('mat. proveedor') || hLow.includes('material proveedor')) colMap.vendor_mat = c;
    else if (hLow.includes('orden de compra')) colMap.po_43 = c;
    else if (hLow.includes('cantidad pedido 43') && c < 20) colMap.ped43_u = c;
    else if (hLow.includes('cantidad pedido 41') && c < 20) colMap.ped41_u = c;
    else if (hLow.includes('recibido real')) colMap.rec_u = c;
    else if (hLow.includes('mov-101') && c < 20) colMap.mov101 = c;
    else if (hLow.includes('revuelto') && hLow.includes('saliendo')) colMap.rev_sal = c;
    else if (hLow === 'destino' && c >= 9 && c <= 15) colMap.rev_dest = c;
    else if (hLow.includes('revuelto') && hLow.includes('entrando')) colMap.rev_ent = c;
    else if (hLow === 'origen' && c >= 9 && c <= 15) colMap.rev_orig = c;
    else if (hLow.includes('faltante') && c < 20) colMap.falt = c;
    else if ((hLow.includes('dañado') || hLow.includes('danado')) && c < 20) colMap.dan = c;
    else if (hLow.includes('sobrante') && c < 20) colMap.sob = c;
    else if ((hLow.includes('partes') || hLow.includes('b-c-d') || hLow.includes('enviada a partes')) && c < 30)
      colMap.partes_bcd = c;
    else if (hLow.includes('total bultos') && hLow.includes('descarga')) colMap.as_bultos = c;
    else if (hLow.includes('unidades en bultos')) colMap.at_unidades = c;
    else if (hLow.includes('bultos') && (hLow.includes('packing') || hLow.includes('paking'))) colMap.aw_bultos = c;
    else if (hLow.includes('cantidad unidades') && c > 40) colMap.ax_unidades = c;
    else if (hLow.includes('empaque') && c > 40) colMap.ay_empaque = c;
    else if (hLow.includes('curva') && c > 40) colMap.curve = c;
    else if (hLow === 'um' && c > 40) colMap.um = c;
    else if ((hLow.includes('cantidad pedido 43') || hLow.includes('po 43')) && c > 40) colMap.ped43_prk = c;
    else if ((hLow.includes('cantidad pedido 41') || hLow.includes('ped 41')) && c > 40) colMap.ped41_prk = c;
  }

  const items: PrepackItem[] = [];

  for (let r = 13; r <= matrix.length; r++) {
    const rawMatCell = getCell(matrix, r, colMap.mat);
    const rawMatText = cleanString(rawMatCell);
    if (!rawMatText || isSummaryFooterCode(rawMatText)) continue;

    const matCode = cleanMaterialCode(rawMatCell);
    if (!matCode || isSummaryFooterCode(matCode)) continue;

    const vendorMat = cleanString(getCell(matrix, r, colMap.vendor_mat));
    const po43 = cleanString(getCell(matrix, r, colMap.po_43));
    const category = cleanString(getCell(matrix, r, colMap.cat));
    let desc = cleanString(getCell(matrix, r, colMap.desc));

    const ped43U = cleanNumber(getCell(matrix, r, colMap.ped43_u));
    const ped41U = cleanNumber(getCell(matrix, r, colMap.ped41_u));
    const rawRecU = getCell(matrix, r, colMap.rec_u);
    const mov101 = cleanNumber(getCell(matrix, r, colMap.mov101));
    const mov641 = cleanNumber(getCell(matrix, r, colMap.mov641));

    const revSal = cleanNumber(getCell(matrix, r, colMap.rev_sal));
    const revDest = cleanMaterialCode(getCell(matrix, r, colMap.rev_dest));
    const revEnt = cleanNumber(getCell(matrix, r, colMap.rev_ent));
    const revOrig = cleanMaterialCode(getCell(matrix, r, colMap.rev_orig));

    let falt = cleanNumber(getCell(matrix, r, colMap.falt));
    let dan = cleanNumber(getCell(matrix, r, colMap.dan));
    let sob = cleanNumber(getCell(matrix, r, colMap.sob));

    const unitCost = cleanNumber(getCell(matrix, r, colMap.unit_cost));
    const rawUm = getCell(matrix, r, colMap.um);
    const umStr = cleanString(rawUm).toUpperCase() || 'PRK';
    const curve = cleanInt(getCell(matrix, r, colMap.curve), umStr === 'UN' ? 1 : 12);

    const ped43Prk = cleanNumber(getCell(matrix, r, colMap.ped43_prk));
    const ped41Prk = cleanNumber(getCell(matrix, r, colMap.ped41_prk));

    const rawAs = getCell(matrix, r, colMap.as_bultos);
    const rawAt = getCell(matrix, r, colMap.at_unidades);
    const rawAw = getCell(matrix, r, colMap.aw_bultos);
    const rawAx = getCell(matrix, r, colMap.ax_unidades);
    const rawAy = getCell(matrix, r, colMap.ay_empaque);

    const pkgAs = parsePackagingExpression(rawAs);
    const pkgAw = parsePackagingExpression(rawAw);
    const pkgAy = parsePackagingExpression(rawAy);
    const empaqueStr = cleanString(rawAy);

    const qcItem = qcDetails[matCode];

    let bultosDesc = pkgAs.total_boxes;
    let unidadesDesc = cleanNumber(rawAt);
    if (bultosDesc === 0 && qcItem && qcItem.bultos > 0) bultosDesc = qcItem.bultos;
    if (unidadesDesc === 0 && qcItem && qcItem.unidades > 0) unidadesDesc = qcItem.unidades;
    else if (unidadesDesc === 0 && pkgAs.total_units > 0) unidadesDesc = pkgAs.total_units;
    else if (unidadesDesc === 0 && pkgAy.total_units > 0) unidadesDesc = pkgAy.total_units;

    let bultosPl = pkgAw.total_boxes;
    let unidadesPl = cleanNumber(rawAx);
    if (bultosPl === 0 && bultosDesc > 0) bultosPl = bultosDesc;
    if (unidadesPl === 0 && unidadesDesc > 0) unidadesPl = unidadesDesc;
    else if (unidadesPl === 0 && ped41U > 0) unidadesPl = ped41U;

    // Extract full/comp/loose breakdown from QC first, else AS compound, else AY compound
    let bEnteros = 0;
    let uEnteros = 0;
    let bComp = 0;
    let uComp = 0;
    let looseUnits = 0;

    if (qcItem && (qcItem.bultos_enteros > 0 || qcItem.bultos_complementos > 0 || qcItem.loose_units > 0)) {
      bEnteros = qcItem.bultos_enteros;
      uEnteros = qcItem.unidades_por_bulto_entero;
      bComp = qcItem.bultos_complementos;
      uComp = qcItem.unidades_por_bulto_comp;
      looseUnits = qcItem.loose_units;
    } else if (pkgAs.is_compound || pkgAs.units_per_box_full > 0) {
      bEnteros = pkgAs.boxes_full;
      uEnteros = pkgAs.units_per_box_full;
      bComp = pkgAs.boxes_comp;
      uComp = pkgAs.units_per_box_comp;
      looseUnits = pkgAs.loose_units;
    } else if (pkgAy.is_compound || pkgAy.units_per_box_full > 0) {
      bEnteros = pkgAy.boxes_full;
      uEnteros = pkgAy.units_per_box_full;
      bComp = pkgAy.boxes_comp;
      uComp = pkgAy.units_per_box_comp;
      looseUnits = pkgAy.loose_units;
    } else if (bultosDesc > 0) {
      bEnteros = bultosDesc;
      uEnteros = unidadesDesc > 0 ? Math.round(unidadesDesc / bultosDesc) : curve;
    }

    let uPerB = curve;
    if (bultosDesc > 0 && unidadesDesc > 0) uPerB = Math.round((unidadesDesc / bultosDesc) * 100) / 100;
    else if (bultosPl > 0 && unidadesPl > 0) uPerB = Math.round((unidadesPl / bultosPl) * 100) / 100;
    else if (/^\d+(?:\.\d+)?$/.test(empaqueStr)) uPerB = Number(empaqueStr);

    const difBultos = bultosDesc - bultosPl;

    let recU = 0;
    if (rawRecU !== null && rawRecU !== undefined && !['', '0', 'None'].includes(String(rawRecU).trim())) {
      recU = cleanNumber(rawRecU);
    } else if (unidadesDesc > 0) {
      recU = unidadesDesc;
    } else if (qcItem && qcItem.unidades > 0) {
      recU = qcItem.unidades;
    }

    if (!desc && qcItem) desc = qcItem.desc;

    const isInQc = Boolean(qcItem && qcItem.unidades > 0);
    let assignedPart = qcItem?.part ?? '';

    if (!assignedPart) {
      const partesBcdVal = colMap.partes_bcd ? cleanNumber(getCell(matrix, r, colMap.partes_bcd)) : 0;
      if (partesBcdVal > 0) {
        assignedPart = 'B';
      } else {
        const obsText = `${cleanString(getCell(matrix, r, colMap.desc))} ${cleanString(getCell(matrix, r, 20))} ${cleanString(getCell(matrix, r, 29))}`;
        const pMatch = obsText.match(/PARTE\s+([A-Z0-9]+)/i);
        if (pMatch) assignedPart = pMatch[1].toUpperCase();
      }
    }

    if (!assignedPart && detectedParts.length > 0) {
      if (isInQc || Boolean(boms[matCode])) {
        assignedPart = 'A';
      } else {
        const otherParts = detectedParts.filter((p) => p !== 'A');
        assignedPart = otherParts[0] ?? 'B';
      }
    }

    const rowBlind =
      isBlindSetup ||
      (Object.keys(qcDetails).length === 0 && Object.keys(boms).length === 0 && rawRecU === null && unidadesDesc === 0 && ped41U > 0);

    if (rowBlind) {
      recU = 0;
      falt = 0;
      dan = 0;
      sob = 0;
    }

    const zeroTotal = isItemZeroTotal({
      quantity_inv_41_units: ped41U,
      received_real_units: recU,
      bultos_packing_list: bultosPl,
      bultos_descarga: bultosDesc,
      danado: dan,
    });

    const canonicalPackaging = formatPackagingExpression({
      fullCartons: bEnteros,
      unitsPerFull: uEnteros,
      compCartons: bComp,
      unitsInComp: uComp,
      looseUnits,
      curveTotal: curve,
      isZeroTotal: zeroTotal,
      existingEmpaqueStr: empaqueStr,
    });

    const children = boms[matCode]
      ? allocateChildSkus(boms[matCode], ped41U, recU, ped41Prk, curve, desc, unitCost, qcDetails)
      : [];

    items.push({
      source_row: r,
      material_code: matCode,
      vendor_material: vendorMat,
      description: desc,
      category,
      order_po_43: po43,
      quantity_po_43_prk: ped43Prk,
      quantity_po_43_units: ped43U,
      quantity_inv_41_prk: ped41Prk,
      quantity_inv_41_units: ped41U,
      received_real_units: recU,
      curve_total: curve,
      unit_cost: unitCost,
      container_part: assignedPart,
      sap_mov_101: mov101,
      sap_mov_641: mov641,
      children,
      faltante: falt,
      danado: dan,
      sobrante: sob,
      revuelto: Math.max(revSal, revEnt),
      revuelto_saliendo: revSal,
      revuelto_destino: revDest,
      revuelto_entrando: revEnt,
      revuelto_origen: revOrig,
      revuelto_partner: revDest || revOrig || undefined,
      is_zero_total: zeroTotal,
      discrepancy_type: 'ExactMatch',
      discrepancy_label: 'CUADRADO (0 DIF)',
      status: 'LIMPIO',
      ncr_amount: 0,
      observations: '',
      bultos_descarga: bultosDesc,
      unidades_descarga: unidadesDesc,
      bultos_packing_list: bultosPl,
      unidades_packing_list: unidadesPl,
      empaque_str: empaqueStr || canonicalPackaging,
      canonical_packaging: canonicalPackaging,
      bultos_str: pkgAs.raw_str || String(Math.round(bultosDesc)),
      um: umStr,
      bultos_enteros: bEnteros,
      unidades_por_bulto_entero: uEnteros,
      bultos_complementos: bComp,
      unidades_por_bulto_comp: uComp,
      loose_units: looseUnits,
      unidades_por_bulto: uPerB,
      dif_bultos: difBultos,
      quantity_flags: qcItem?.flags ? [...qcItem.flags] : [],
      is_blind_setup: rowBlind,
    });
  }

  // Deterministic 3-Way Classification (Acr.Core DiscrepancyClassifier parity)
  const classification = classifyAllDiscrepancies(items, isBlindSetup);
  const vbaInfo = await computeVbaSha256(cierreBuffer);

  const summary = computeSummary(items, meta, {
    revueltoPairsCount: classification.revueltoPairsCount,
    totalShortageUnits: classification.totalShortageUnits,
    totalOverageUnits: classification.totalOverageUnits,
    vbaPresent: vbaInfo.vbaPresent,
    vbaSha256: vbaInfo.sha256,
  });

  const validation = validateAllItems(items);
  const elapsedSeconds = (performance.now() - tStart) / 1000;

  return {
    items,
    summary,
    boms,
    qcDetails,
    validation,
    warnings,
    cierreRawBytes: cierreBuffer,
    cierreFileName,
    qcFileName,
    isBlindSetup,
    elapsedSeconds,
  };
}

export function exportRowsToExcel(rows: Record<string, unknown>[], sheetName: string, fileName: string): void {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, fileName);
}

export function exportOfficialSquaredWorkbook(audit: AuditData, excludeZeroTotal = true): string {
  const { items, summary, qcDetails } = audit;

  const cierreRows = items.map((i) => {
    const qcItem = qcDetails[i.material_code];
    const isOtherPart =
      i.discrepancy_type === 'PendingPart' ||
      i.status.startsWith('PENDIENTE_PARTE') ||
      (Boolean(i.container_part) && i.container_part !== 'A' && i.received_real_units === 0);

    const bultosDesc =
      i.is_zero_total || i.received_real_units === 0 || isOtherPart
        ? 0
        : i.bultos_descarga > 0
          ? Math.round(i.bultos_descarga)
          : qcItem && qcItem.bultos > 0
            ? Math.round(qcItem.bultos)
            : Math.round(i.quantity_inv_41_prk);

    const unidadesDesc =
      i.is_zero_total || i.received_real_units === 0 || isOtherPart
        ? 0
        : Math.round(i.unidades_descarga > 0 ? i.unidades_descarga : i.received_real_units);

    const unidadesPl =
      i.is_zero_total || isOtherPart
        ? 0
        : Math.round(i.unidades_packing_list > 0 ? i.unidades_packing_list : i.quantity_inv_41_units);

    const bultosPl =
      i.is_zero_total || isOtherPart
        ? 0
        : Math.round(i.bultos_packing_list > 0 ? i.bultos_packing_list : bultosDesc);

    const emp =
      i.is_zero_total || isOtherPart
        ? '0'
        : i.canonical_packaging || i.empaque_str || String(i.curve_total || 0);

    return {
      FILA_EXCEL: i.source_row,
      'ORDEN DE COMPRA': i.order_po_43,
      PROVEEDOR: summary.vendor_name,
      CATEGORIA: i.category,
      'MAT. PROVEEDOR': i.vendor_material,
      MATERIAL: i.material_code,
      'CANTIDAD PEDIDO 43': Math.round(i.quantity_po_43_units),
      'CANTIDAD PEDIDO 41': Math.round(i.quantity_inv_41_units),
      'RECIBIDO REAL & FISICO': Math.round(i.received_real_units),
      'MOV-101': Math.round(i.sap_mov_101),
      FALTANTE: Math.round(i.faltante),
      DAÑADO: Math.round(i.danado),
      SOBRANTE: Math.round(i.sobrante),
      REVUELTO: Math.round(i.revuelto),
      DESCRIPCION: i.description,
      'VALOR UNITARIO $': i.unit_cost,
      'BULTOS DESCARGA (AS)': bultosDesc,
      'UNIDADES DESCARGA (AT)': unidadesDesc,
      'DIF FISICO VS RECIBIDO (AU)': Math.round(unidadesDesc - i.received_real_units),
      'BULTOS PL (AW)': bultosPl,
      'UNIDADES PL (AX)': unidadesPl,
      'EMPAQUE (AY)': emp,
      'DIF FACT VS RECIBIDO (AZ)': Math.round(unidadesPl - i.received_real_units),
      CURVA: i.curve_total,
      UM: i.um,
      CLASIFICACION: i.discrepancy_label,
    };
  });

  const bultosSource = excludeZeroTotal ? items.filter((it) => !it.is_zero_total) : items;
  const bultosRows = bultosSource.map((it) => ({
    CATEGORIA: it.category || 'TEXTIL',
    'CANTIDAD DE BULTOS Z-Carga': Math.round(it.bultos_descarga),
    'CANTIDAD DE UNIDADES Z-Carga': Math.round(it.unidades_descarga || it.received_real_units),
    MATERIAL: it.material_code,
    'RECIBIDO REAL & FISICO': Math.round(it.received_real_units),
    Curva: it.curve_total,
    UM: it.um || 'PRK',
    DESCRIPCION: it.description,
    'EMPAQUE UNIDADES Z-Carga': it.canonical_packaging || '0',
    'Bultos Packing List': Math.round(it.bultos_packing_list || it.bultos_descarga),
    'Diferencia Bultos': Math.round(it.dif_bultos),
    ESTADO: it.discrepancy_label,
  }));

  const wb = XLSX.utils.book_new();
  const wsMeta = XLSX.utils.aoa_to_sheet([
    ['CAMPO', 'VALOR'],
    ['IMPORTACION', summary.import_id],
    ['PROVEEDOR', summary.vendor_name],
    ['PEDIDO NACIONALIZACION (41)', summary.invoice_number],
    ['ESTILOS ACTIVOS', summary.active_styles_count],
    ['MATERIALES SIN MOVIMIENTO (0 TOTAL)', summary.zero_total_count],
    ['UNIDADES FACTURADAS (41)', summary.total_ped_41_units],
    ['UNIDADES RECIBIDAS MUELLE', summary.total_received_units],
    ['FALTANTE CARDINAL (UN)', summary.total_shortage_units],
    ['SOBRANTE CARDINAL (UN)', summary.total_overage_units],
    ['PARES REVUELTO CANDIDATOS', summary.revuelto_pairs_count],
    ['BULTOS RECIBIDOS', summary.received_packages],
    ['BULTOS TEORICOS', summary.theor_packages],
    ['NOTA DE CREDITO (USD)', summary.total_ncr_amount],
    ['ESTADO GENERAL', summary.status],
  ]);
  XLSX.utils.book_append_sheet(wb, wsMeta, 'DATOS GENERALES');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cierreRows), 'CIERRE RECEP');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bultosRows), 'Bultos');

  const suffix = summary.is_blind_setup ? 'Montaje_Ciego_Pausa' : 'Cierre_Oficial_Cuadrado';
  const fileName = `${summary.import_id}_${suffix}.xlsx`;
  XLSX.writeFile(wb, fileName);
  return fileName;
}
