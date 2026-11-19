import JSZip from 'jszip';
import type { AuditData, PrepackItem } from '../types/reconciliation';
import { exportOfficialSquaredWorkbook } from './excelParser';

export interface VbaIntegrityResult {
  vbaPresent: boolean;
  sha256?: string;
}

export interface MacroSafeExportResult {
  fileName: string;
  vbaPreserved: boolean;
  vbaSha256?: string;
  fallbackXlsx: boolean;
  patchedRows: number;
  omittedZeroRows: number;
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  const hashBuffer = await crypto.subtle.digest('SHA-256', copy.buffer);
  return Array.from(new Uint8Array(hashBuffer))
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('');
}

export async function computeVbaSha256(buffer: ArrayBuffer): Promise<VbaIntegrityResult> {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const vbaEntry = zip.file('xl/vbaProject.bin');
    if (!vbaEntry) {
      return { vbaPresent: false };
    }
    const bytes = await vbaEntry.async('uint8array');
    const sha256 = await sha256Hex(bytes);
    return { vbaPresent: true, sha256 };
  } catch {
    return { vbaPresent: false };
  }
}

function colLettersToIndex(col: string): number {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

function splitCellRef(ref: string): { col: string; row: number } {
  const m = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!m) return { col: 'A', row: 1 };
  return { col: m[1].toUpperCase(), row: Number(m[2]) };
}

function findOrCreateRow(doc: Document, sheetData: Element, rowNum: number): Element {
  const rows = Array.from(sheetData.children).filter((el) => el.localName === 'row');
  for (const r of rows) {
    const rAttr = Number(r.getAttribute('r'));
    if (rAttr === rowNum) return r;
    if (rAttr > rowNum) {
      const created = doc.createElementNS(sheetData.namespaceURI, 'row');
      created.setAttribute('r', String(rowNum));
      sheetData.insertBefore(created, r);
      return created;
    }
  }
  const created = doc.createElementNS(sheetData.namespaceURI, 'row');
  created.setAttribute('r', String(rowNum));
  sheetData.appendChild(created);
  return created;
}

function findOrCreateCell(doc: Document, sheetData: Element, rowNum: number, colLetters: string): Element {
  const rowEl = findOrCreateRow(doc, sheetData, rowNum);
  const targetRef = `${colLetters.toUpperCase()}${rowNum}`;
  const targetColIdx = colLettersToIndex(colLetters.toUpperCase());

  const cells = Array.from(rowEl.children).filter((el) => el.localName === 'c');
  for (const c of cells) {
    const rAttr = (c.getAttribute('r') || '').toUpperCase();
    if (rAttr === targetRef) return c;
    const { col } = splitCellRef(rAttr);
    if (colLettersToIndex(col) > targetColIdx) {
      const created = doc.createElementNS(sheetData.namespaceURI, 'c');
      created.setAttribute('r', targetRef);
      rowEl.insertBefore(created, c);
      return created;
    }
  }
  const created = doc.createElementNS(sheetData.namespaceURI, 'c');
  created.setAttribute('r', targetRef);
  rowEl.appendChild(created);
  return created;
}

function clearCellContent(cell: Element): void {
  while (cell.firstChild) {
    cell.removeChild(cell.firstChild);
  }
  cell.removeAttribute('t');
}

function setNumericCell(doc: Document, sheetData: Element, rowNum: number, colLetters: string, value: number): void {
  const cell = findOrCreateCell(doc, sheetData, rowNum, colLetters);
  clearCellContent(cell);
  const v = doc.createElementNS(sheetData.namespaceURI, 'v');
  const cleanVal = Number.isFinite(value) ? (Number.isInteger(value) ? String(value) : String(Math.round(value * 100) / 100)) : '0';
  v.textContent = cleanVal;
  cell.appendChild(v);
}

function setInlineStringCell(doc: Document, sheetData: Element, rowNum: number, colLetters: string, text: string): void {
  const cell = findOrCreateCell(doc, sheetData, rowNum, colLetters);
  clearCellContent(cell);
  cell.setAttribute('t', 'inlineStr');
  const isEl = doc.createElementNS(sheetData.namespaceURI, 'is');
  const tEl = doc.createElementNS(sheetData.namespaceURI, 't');
  tEl.textContent = text;
  isEl.appendChild(tEl);
  cell.appendChild(isEl);
}

function setFormulaCell(
  doc: Document,
  sheetData: Element,
  rowNum: number,
  colLetters: string,
  formula: string,
  preserveExistingFormula: boolean
): void {
  const cell = findOrCreateCell(doc, sheetData, rowNum, colLetters);
  if (preserveExistingFormula) {
    const existingF = Array.from(cell.children).find((el) => el.localName === 'f');
    if (existingF && existingF.textContent && existingF.textContent.trim().length > 0) {
      return;
    }
  }
  clearCellContent(cell);
  const f = doc.createElementNS(sheetData.namespaceURI, 'f');
  f.textContent = formula;
  cell.appendChild(f);
}

function triggerBlobDownload(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}

function resolveSheetTargets(workbookXml: string, relsXml: string): Map<string, string> {
  const parser = new DOMParser();
  const wbDoc = parser.parseFromString(workbookXml, 'application/xml');
  const relsDoc = parser.parseFromString(relsXml, 'application/xml');

  const relMap = new Map<string, string>();
  const relNodes = Array.from(relsDoc.getElementsByTagName('*')).filter((el) => el.localName === 'Relationship');
  for (const rel of relNodes) {
    const id = rel.getAttribute('Id');
    const target = rel.getAttribute('Target');
    if (id && target) {
      const normalized = target.startsWith('/') ? target.slice(1) : target.startsWith('xl/') ? target : `xl/${target}`;
      relMap.set(id, normalized);
    }
  }

  const sheetMap = new Map<string, string>();
  const sheetNodes = Array.from(wbDoc.getElementsByTagName('*')).filter((el) => el.localName === 'sheet');
  for (const sh of sheetNodes) {
    const name = sh.getAttribute('name') || '';
    const rId =
      sh.getAttribute('r:id') ||
      sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') ||
      '';
    const target = relMap.get(rId);
    if (name && target) {
      sheetMap.set(name, target);
    }
  }

  return sheetMap;
}

function resolveRowValues(item: PrepackItem): {
  bultosDesc: string | number;
  unidadesDesc: number;
  bultosPl: number;
  unidadesPl: number;
  empaqueAy: string;
} {
  const isOtherPart =
    item.discrepancy_type === 'PendingPart' ||
    item.status.startsWith('PENDIENTE_PARTE') ||
    (Boolean(item.container_part) && item.container_part !== 'A' && item.received_real_units === 0);

  if (item.is_zero_total || isOtherPart) {
    return {
      bultosDesc: 0,
      unidadesDesc: 0,
      bultosPl: 0,
      unidadesPl: 0,
      empaqueAy: '0',
    };
  }

  const unidadesDesc =
    item.received_real_units === 0
      ? 0
      : Math.round(item.unidades_descarga > 0 ? item.unidades_descarga : item.received_real_units);

  const bultosDesc: string | number =
    unidadesDesc === 0
      ? 0
      : item.bultos_descarga > 0
        ? Math.round(item.bultos_descarga)
        : Math.round(item.quantity_inv_41_prk);

  const unidadesPl = Math.round(
    item.unidades_packing_list > 0
      ? item.unidades_packing_list
      : item.quantity_inv_41_units > 0
        ? item.quantity_inv_41_units
        : unidadesDesc
  );

  const bultosPl = Math.round(
    item.bultos_packing_list > 0 ? item.bultos_packing_list : typeof bultosDesc === 'number' ? bultosDesc : 0
  );

  const empaqueAy = item.canonical_packaging || item.empaque_str || String(item.curve_total || 0);

  return { bultosDesc, unidadesDesc, bultosPl, unidadesPl, empaqueAy };
}

/**
 * Surgical Macro-Safe OpenXML `.xlsm` Exporter.
 * Preserves `xl/vbaProject.bin`, all relationships, styles, charts, and native Excel formulas (`=SUM(...)`, `+B6-B5`, `+AT{r}-H{r}`, `+AX{r}-H{r}`).
 * Verifies SHA-256 of `xl/vbaProject.bin` before and after packaging.
 */
export async function exportMacroSafeXlsm(
  audit: AuditData,
  excludeZeroTotal = true
): Promise<MacroSafeExportResult> {
  const omittedZeroRows = excludeZeroTotal ? audit.items.filter((i) => i.is_zero_total).length : 0;

  if (!audit.cierreRawBytes) {
    const fileName = exportOfficialSquaredWorkbook(audit, excludeZeroTotal);
    return {
      fileName,
      vbaPreserved: false,
      fallbackXlsx: true,
      patchedRows: audit.items.length,
      omittedZeroRows,
    };
  }

  const zip = await JSZip.loadAsync(audit.cierreRawBytes);
  const vbaEntryBefore = zip.file('xl/vbaProject.bin');
  const preSha = vbaEntryBefore ? await sha256Hex(await vbaEntryBefore.async('uint8array')) : undefined;

  const wbXmlEntry = zip.file('xl/workbook.xml');
  const relsXmlEntry = zip.file('xl/_rels/workbook.xml.rels');

  if (!wbXmlEntry || !relsXmlEntry) {
    const fileName = exportOfficialSquaredWorkbook(audit, excludeZeroTotal);
    return {
      fileName,
      vbaPreserved: false,
      fallbackXlsx: true,
      patchedRows: audit.items.length,
      omittedZeroRows,
    };
  }

  const wbXml = await wbXmlEntry.async('string');
  const relsXml = await relsXmlEntry.async('string');
  const sheetTargets = resolveSheetTargets(wbXml, relsXml);

  const cierreEntry = Array.from(sheetTargets.entries()).find(([name]) =>
    name.toUpperCase().includes('CIERRE RECEP')
  );
  const bultosEntry = Array.from(sheetTargets.entries()).find(
    ([name]) => name.trim().toUpperCase() === 'BULTOS'
  );

  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  let patchedRows = 0;

  // 1. Surgically patch CIERRE RECEP anchored to item.source_row
  if (cierreEntry) {
    const [, cierrePath] = cierreEntry;
    const cierreFile = zip.file(cierrePath);
    if (cierreFile) {
      const xmlStr = await cierreFile.async('string');
      const doc = parser.parseFromString(xmlStr, 'application/xml');
      const sheetData = Array.from(doc.getElementsByTagName('*')).find((el) => el.localName === 'sheetData');

      if (sheetData) {
        // Header package totals (B5 = theor, B6 = received, B7 = formula +B6-B5 preserved)
        if (audit.summary.theor_packages > 0) {
          setNumericCell(doc, sheetData, 5, 'B', Math.round(audit.summary.theor_packages));
        }
        if (!audit.summary.is_blind_setup) {
          setNumericCell(doc, sheetData, 6, 'B', Math.round(audit.summary.received_packages));
        }
        setFormulaCell(doc, sheetData, 7, 'B', '+B6-B5', true);

        for (const item of audit.items) {
          const r = item.source_row;
          if (!r || r < 13) continue;

          const vals = resolveRowValues(item);

          if (typeof vals.bultosDesc === 'number') {
            setNumericCell(doc, sheetData, r, 'AS', vals.bultosDesc);
          } else {
            setInlineStringCell(doc, sheetData, r, 'AS', vals.bultosDesc);
          }
          setNumericCell(doc, sheetData, r, 'AT', vals.unidadesDesc);
          setFormulaCell(doc, sheetData, r, 'AU', `+AT${r}-H${r}`, true);

          setNumericCell(doc, sheetData, r, 'AW', vals.bultosPl);
          setNumericCell(doc, sheetData, r, 'AX', vals.unidadesPl);

          if (/^\d+(?:\.\d+)?$/.test(vals.empaqueAy)) {
            setNumericCell(doc, sheetData, r, 'AY', Number(vals.empaqueAy));
          } else {
            setInlineStringCell(doc, sheetData, r, 'AY', vals.empaqueAy);
          }
          setFormulaCell(doc, sheetData, r, 'AZ', `+AX${r}-H${r}`, true);
          patchedRows++;
        }

        zip.file(cierrePath, serializer.serializeToString(doc));
      }
    }
  }

  // 2. Update Bultos sheet formula links (omitting zero-total rows if requested)
  if (bultosEntry && cierreEntry) {
    const [cierreName] = cierreEntry;
    const [, bultosPath] = bultosEntry;
    const bultosFile = zip.file(bultosPath);
    if (bultosFile) {
      const xmlStr = await bultosFile.async('string');
      const doc = parser.parseFromString(xmlStr, 'application/xml');
      const sheetData = Array.from(doc.getElementsByTagName('*')).find((el) => el.localName === 'sheetData');

      if (sheetData) {
        const exportItems = excludeZeroTotal
          ? audit.items.filter((i) => !i.is_zero_total)
          : audit.items;

        exportItems.forEach((item, idx) => {
          const targetRow = 2 + idx;
          const srcRow = item.source_row;
          if (!srcRow || srcRow < 13) return;
          const refPrefix = `'${cierreName.replace(/'/g, "''")}'!`;
          setFormulaCell(doc, sheetData, targetRow, 'B', `${refPrefix}C${srcRow}`, false);
          setFormulaCell(doc, sheetData, targetRow, 'C', `${refPrefix}AS${srcRow}`, false);
          setFormulaCell(doc, sheetData, targetRow, 'D', `${refPrefix}AT${srcRow}`, false);
          setFormulaCell(doc, sheetData, targetRow, 'E', `${refPrefix}E${srcRow}`, false);
          setFormulaCell(doc, sheetData, targetRow, 'F', `${refPrefix}H${srcRow}`, false);
          setFormulaCell(doc, sheetData, targetRow, 'G', `${refPrefix}BB${srcRow}`, false);
          setFormulaCell(doc, sheetData, targetRow, 'H', `${refPrefix}BC${srcRow}`, false);
          setFormulaCell(doc, sheetData, targetRow, 'I', `${refPrefix}AD${srcRow}`, false);
          setFormulaCell(doc, sheetData, targetRow, 'J', `${refPrefix}AY${srcRow}`, false);
        });

        zip.file(bultosPath, serializer.serializeToString(doc));
      }
    }
  }

  // Generate output preserving zip binary entries verbatim
  const outArrayBuffer = await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  // Post-export VBA SHA-256 verification
  const postIntegrity = await computeVbaSha256(outArrayBuffer);
  if (preSha && (!postIntegrity.vbaPresent || postIntegrity.sha256 !== preSha)) {
    throw new Error(
      'Fallo de integridad binaria: el hash SHA-256 de xl/vbaProject.bin cambió durante la exportación. Exportación abortada por seguridad.'
    );
  }

  const ext = preSha || audit.cierreFileName.toLowerCase().endsWith('.xlsm') ? 'xlsm' : 'xlsx';
  const suffix = audit.summary.is_blind_setup ? 'Montaje_Ciego_Pausa' : 'Cierre_Oficial_MacroSafe';
  const safeId = (audit.summary.import_id || 'IMPORT').replace(/[^\w\-]+/g, '_');
  const fileName = `${safeId}_${suffix}.${ext}`;

  const mime =
    ext === 'xlsm'
      ? 'application/vnd.ms-excel.sheet.macroEnabled.12'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  triggerBlobDownload(new Blob([outArrayBuffer], { type: mime }), fileName);

  return {
    fileName,
    vbaPreserved: Boolean(preSha && postIntegrity.sha256 === preSha),
    vbaSha256: postIntegrity.sha256,
    fallbackXlsx: false,
    patchedRows,
    omittedZeroRows,
  };
}
