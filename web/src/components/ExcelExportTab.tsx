import React, { useMemo, useState } from 'react';
import { Check, Copy, Download, ShieldCheck, Zap, AlertTriangle } from 'lucide-react';
import type { AuditData, Language } from '../types/reconciliation';
import { exportOfficialSquaredWorkbook, exportRowsToExcel } from '../utils/excelParser';
import { exportMacroSafeXlsm, type MacroSafeExportResult } from '../utils/macroSafeXlsmExporter';
import { formatExactInt } from '../utils/sanitizer';
import { t } from '../utils/i18n';

interface ExcelExportTabProps {
  audit: AuditData;
  lang: Language;
}

export const ExcelExportTab: React.FC<ExcelExportTabProps> = ({ audit, lang }) => {
  const { items, summary, qcDetails } = audit;
  const [subtab, setSubtab] = useState<'block1' | 'block2'>('block2');
  const [excludeZeroTotal, setExcludeZeroTotal] = useState(true);
  const [copied1, setCopied1] = useState(false);
  const [copied2, setCopied2] = useState(false);
  const [exportingXlsm, setExportingXlsm] = useState(false);
  const [exportResult, setExportResult] = useState<MacroSafeExportResult | null>(null);
  const [exportError, setExportError] = useState<string | null>(null);

  const zeroTotalCount = useMemo(() => items.filter((i) => i.is_zero_total).length, [items]);

  const activeItems = useMemo(
    () => (excludeZeroTotal ? items.filter((i) => !i.is_zero_total) : items),
    [items, excludeZeroTotal]
  );

  const coreRows = useMemo(
    () =>
      activeItems.map((i) => ({
        'Fila Excel': i.source_row,
        'ORDEN DE COMPRA': i.order_po_43,
        Proveedor: summary.vendor_name,
        Categoría: i.category,
        'MAT. PROVEEDOR': i.vendor_material,
        MATERIAL: i.material_code,
        'CANTIDAD PEDIDO 43': Math.round(i.quantity_po_43_units),
        'CANTIDAD PEDIDO 41': Math.round(i.quantity_inv_41_units),
        'RECIBIDO REAL & FISICO': Math.round(i.received_real_units),
        'CANTIDAD INGRESO MOV-101': Math.round(i.sap_mov_101),
        FALTANTE: Math.round(i.faltante),
        DAÑADO: Math.round(i.danado),
        SOBRANTE: Math.round(i.sobrante),
        REVUELTO: Math.round(i.revuelto),
        DESCRIPCION: i.description,
        'Valor unitario PEDIDO NAC $': i.unit_cost,
        Curva: i.curve_total,
        'Empaque Canónico': i.canonical_packaging,
        'CANTIDAD PEDIDO 43 (PRK)': Math.round(i.quantity_po_43_prk),
        'CANTIDAD PEDIDO 41 (PRK)': Math.round(i.quantity_inv_41_prk),
        CLASIFICACION: i.discrepancy_label,
      })),
    [activeItems, summary.vendor_name]
  );

  const { descargaRows, tsvAsAt, tsvAwAy } = useMemo(() => {
    const rows: Record<string, unknown>[] = [];
    const lines1: string[] = [];
    const lines2: string[] = [];

    activeItems.forEach((i) => {
      const qcItem = qcDetails[i.material_code];
      const isOtherPart =
        i.discrepancy_type === 'PendingPart' ||
        i.status.startsWith('PENDIENTE_PARTE') ||
        (Boolean(i.container_part) && i.container_part !== 'A' && i.received_real_units === 0);

      let unidadesDesc = 0;
      let bultosDesc: string | number = 0;
      if (!(i.is_zero_total || i.received_real_units === 0 || isOtherPart)) {
        unidadesDesc = Math.round(
          i.unidades_descarga > 0
            ? i.unidades_descarga
            : qcItem && qcItem.unidades > 0
              ? qcItem.unidades
              : i.received_real_units
        );
        if (i.bultos_descarga > 0) {
          bultosDesc = Math.round(i.bultos_descarga);
        } else if (qcItem && qcItem.bultos > 0) {
          bultosDesc = Math.round(qcItem.bultos);
        } else {
          bultosDesc =
            i.quantity_inv_41_prk > 0
              ? Math.round(i.quantity_inv_41_prk)
              : unidadesDesc > 0
                ? Math.round(unidadesDesc / Math.max(i.curve_total, 1))
                : 0;
        }
      }

      let unidadesPl = 0;
      let bultosPl: string | number = 0;
      let curvaEmpaque = '0';
      if (!(i.is_zero_total || isOtherPart)) {
        unidadesPl = Math.round(
          i.unidades_packing_list > 0
            ? i.unidades_packing_list
            : i.quantity_inv_41_units > 0
              ? i.quantity_inv_41_units
              : unidadesDesc
        );
        bultosPl = i.bultos_packing_list > 0 ? Math.round(i.bultos_packing_list) : bultosDesc;
        curvaEmpaque = i.canonical_packaging || i.empaque_str || String(i.curve_total > 0 ? i.curve_total : 0);
      }

      const difAu = Math.round(unidadesDesc - i.received_real_units);
      const difAz = Math.round(unidadesPl - i.received_real_units);

      rows.push({
        'Fila Excel': i.source_row,
        'Material SAP': i.material_code,
        'Bultos Descarga (AS)': bultosDesc,
        'Unidades Descarga (AT)': unidadesDesc,
        'Dif AT-H (AU)': difAu,
        'Bultos PL (AW)': bultosPl,
        'Unidades PL (AX)': unidadesPl,
        'Empaque (AY)': curvaEmpaque,
        'Dif AX-H (AZ)': difAz,
      });

      const bDescStr = typeof bultosDesc === 'number' ? formatExactInt(bultosDesc) : String(bultosDesc);
      const bPlStr = typeof bultosPl === 'number' ? formatExactInt(bultosPl) : String(bultosPl);

      lines1.push(`${bDescStr}\t${formatExactInt(unidadesDesc)}`);
      lines2.push(`${bPlStr}\t${formatExactInt(unidadesPl)}\t${curvaEmpaque}`);
    });

    return {
      descargaRows: rows,
      tsvAsAt: lines1.join('\n'),
      tsvAwAy: lines2.join('\n'),
    };
  }, [activeItems, qcDetails]);

  const copyText = async (text: string, which: 1 | 2) => {
    try {
      await navigator.clipboard.writeText(text);
      if (which === 1) {
        setCopied1(true);
        setTimeout(() => setCopied1(false), 2000);
      } else {
        setCopied2(true);
        setTimeout(() => setCopied2(false), 2000);
      }
    } catch {
      // Clipboard blocked
    }
  };

  const handleMacroSafeExport = async () => {
    setExportingXlsm(true);
    setExportError(null);
    setExportResult(null);
    try {
      const res = await exportMacroSafeXlsm(audit, excludeZeroTotal);
      setExportResult(res);
    } catch (err) {
      setExportError(err instanceof Error ? err.message : String(err));
    } finally {
      setExportingXlsm(false);
    }
  };

  return (
    <div className="space-y-5">
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <h3 className="text-lg font-bold text-white">{t('export_title', lang)}</h3>
          <p className="mt-1 text-xs text-slate-300">
            {lang === 'es'
              ? 'Genera tiras de pegado instantáneo ZDESCARGA (AS:AT) y PACKING LIST (AW:AY) y exporta el archivo .xlsm Macro-Safe preservando xl/vbaProject.bin y todas las fórmulas nativas.'
              : 'Generates instant ZDESCARGA (AS:AT) & PACKING LIST (AW:AY) paste strips and exports a Macro-Safe .xlsm preserving xl/vbaProject.bin and native Excel formulas.'}
          </p>
        </div>

        <label className="flex cursor-pointer items-center gap-2 rounded-md border border-slate-700 bg-slate-800/80 px-3 py-1.5 text-xs font-semibold text-slate-200">
          <input
            type="checkbox"
            checked={excludeZeroTotal}
            onChange={(e) => setExcludeZeroTotal(e.target.checked)}
            className="rounded border-slate-600 bg-slate-900 accent-sky-500"
          />
          <span>
            {lang === 'es'
              ? 'Omitir materiales en cero de Tira 1, Tira 2 y hoja Bultos'
              : 'Exclude zero-total items from Strips & Bultos'}
            {zeroTotalCount > 0 ? ` (${zeroTotalCount})` : ''}
          </span>
        </label>
      </div>

      {/* Export status banners */}
      {exportResult && (
        <div className="flex items-start gap-2.5 rounded-md border border-emerald-500/40 bg-emerald-500/10 p-3.5 text-xs text-emerald-200">
          <ShieldCheck className="mt-0.5 h-4 w-4 shrink-0 text-emerald-400" />
          <div>
            <p className="font-bold">
              {exportResult.fallbackXlsx
                ? lang === 'es'
                  ? `Libro Excel generado: ${exportResult.fileName} (contenedor de demostración sin plantilla .xlsm binaria)`
                  : `Excel workbook generated: ${exportResult.fileName} (demo container without binary .xlsm template)`
                : lang === 'es'
                  ? `Exportación Macro-Safe completada: ${exportResult.fileName} (${exportResult.patchedRows} filas ancladas a su fila original)`
                  : `Macro-Safe export complete: ${exportResult.fileName} (${exportResult.patchedRows} rows anchored to source row)`}
            </p>
            {exportResult.vbaPreserved && exportResult.vbaSha256 && (
              <p className="mt-1 font-mono text-[11px] text-emerald-300">
                xl/vbaProject.bin SHA-256 verificado (100% intacto): {exportResult.vbaSha256}
              </p>
            )}
          </div>
        </div>
      )}

      {exportError && (
        <div className="flex items-start gap-2 rounded-md border border-red-500/40 bg-red-500/15 p-3 text-xs text-red-200">
          <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0 text-red-400" />
          <span>{exportError}</span>
        </div>
      )}

      {/* Subtabs */}
      <div className="flex flex-wrap gap-2 border-b border-slate-700/60 pb-2">
        <button
          type="button"
          onClick={() => setSubtab('block2')}
          className={`rounded-t-md px-4 py-2 text-xs font-bold transition ${
            subtab === 'block2'
              ? 'border-b-2 border-sky-400 bg-slate-800 text-sky-300'
              : 'text-slate-400 hover:bg-slate-800/50'
          }`}
        >
          📦 {lang === 'es' ? 'Bloque 2: ZDESCARGA y PACKING LIST (Cols AS:AZ)' : 'Block 2: ZDESCARGA & PACKING LIST (Cols AS:AZ)'}
        </button>
        <button
          type="button"
          onClick={() => setSubtab('block1')}
          className={`rounded-t-md px-4 py-2 text-xs font-bold transition ${
            subtab === 'block1'
              ? 'border-b-2 border-sky-400 bg-slate-800 text-sky-300'
              : 'text-slate-400 hover:bg-slate-800/50'
          }`}
        >
          📊 {lang === 'es' ? 'Bloque 1: Columnas Principales (A - BB)' : 'Block 1: Core Columns (A - BB)'}
        </button>
      </div>

      {subtab === 'block1' ? (
        <div className="space-y-4">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <h4 className="text-sm font-bold text-white">
              {lang === 'es'
                ? `Columnas Principales de Conciliación (${coreRows.length} filas activas)`
                : `Core Reconciliation Columns (${coreRows.length} active rows)`}
            </h4>
            <button
              type="button"
              onClick={() =>
                exportRowsToExcel(
                  coreRows,
                  'CIERRE_RECEP_VALORES',
                  summary.is_blind_setup
                    ? `${summary.import_id}_Montaje_Ciego_Pausa.xlsx`
                    : `${summary.import_id}_Valores_Cierre_Recep_Limpio.xlsx`
                )
              }
              className="flex items-center gap-2 rounded-md bg-gradient-to-br from-[#0284C7] to-[#0369A1] px-4 py-2 text-xs font-bold text-white shadow hover:from-[#0369A1] hover:to-[#075985]"
            >
              <Download className="h-4 w-4" />
              {lang === 'es' ? 'Descargar Columnas Principales (.xlsx)' : 'Download Core Columns (.xlsx)'}
            </button>
          </div>

          <div className="overflow-x-auto rounded-lg border border-slate-700/70 bg-slate-900/60">
            <table className="min-w-full divide-y divide-slate-800 text-left text-xs">
              <thead className="bg-slate-800 text-[11px] font-bold text-slate-300 uppercase">
                <tr>
                  {Object.keys(coreRows[0] ?? {}).map((col) => (
                    <th key={col} className="whitespace-nowrap px-3 py-2">
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/70 text-slate-200">
                {coreRows.map((row, idx) => (
                  <tr key={idx} className="hover:bg-slate-800/50">
                    {Object.values(row).map((val, cIdx) => (
                      <td key={cIdx} className="whitespace-nowrap px-3 py-2">
                        {String(val)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div className="space-y-5">
          {summary.is_blind_setup && (
            <div className="rounded-md border border-amber-500/40 bg-amber-500/10 p-3 text-xs text-amber-200">
              ⚠️{' '}
              {lang === 'es'
                ? 'Aviso de Modo Pausa (Sin Control de Calidad): Los valores de descarga permanecen en 0 hasta recibir el reporte físico del muelle.'
                : 'Pause Mode Notice (No QC): Dock discharge values remain at 0 until physical dock tally is uploaded.'}
            </div>
          )}

          <div className="rounded-md border border-sky-500/40 bg-sky-500/10 p-4 text-xs leading-relaxed text-sky-200">
            <p className="font-bold">
              💡 {lang === 'es' ? 'Regla de Pegado Oficial (Acr.Core):' : 'Official Paste Rule (Acr.Core):'}
            </p>
            <ol className="mt-1.5 list-inside list-decimal space-y-1">
              <li>
                {lang === 'es'
                  ? 'Tira 1 (AS:AT) pega Total Bultos Descarga y Unidades en Bultos. La fórmula nativa AU (=+AT-H) se conserva intacta.'
                  : 'Strip 1 (AS:AT) pastes Dock Cartons and Units. Native Excel formula AU (=+AT-H) stays intact.'}
              </li>
              <li>
                {lang === 'es'
                  ? 'Tira 2 (AW:AY) pega Bultos PL, Unidades PL y Empaque Canónico (ej. 8*48 1*36 o 8*48 +36).'
                  : 'Strip 2 (AW:AY) pastes PL Cartons, PL Units, and Canonical Packaging (e.g. 8*48 1*36 or 8*48 +36).'}
              </li>
              <li>
                {lang === 'es'
                  ? 'Con la casilla activa, los materiales Sin Movimiento (0 total) se omiten automáticamente de Tira 1, Tira 2 y la hoja Bultos.'
                  : 'When enabled, Zero-Total items are excluded automatically from Strip 1, Strip 2, and the Bultos sheet.'}
              </li>
            </ol>
          </div>

          {/* Preview Table */}
          <div className="overflow-x-auto rounded-lg border border-slate-700/70 bg-slate-900/60">
            <table className="min-w-full divide-y divide-slate-800 text-left text-xs">
              <thead className="bg-slate-800 text-[11px] font-bold text-slate-300 uppercase">
                <tr>
                  {Object.keys(descargaRows[0] ?? {}).map((col) => (
                    <th key={col} className="whitespace-nowrap px-3 py-2">
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/70 text-slate-200">
                {descargaRows.map((row, idx) => (
                  <tr key={idx} className="hover:bg-slate-800/50">
                    {Object.values(row).map((val, cIdx) => (
                      <td key={cIdx} className="whitespace-nowrap px-3 py-2 font-mono">
                        {String(val)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Dual Copy Strips */}
          <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
            <div className="space-y-2 rounded-lg border border-slate-700/60 bg-[#1E293B] p-4">
              <div className="flex items-center justify-between gap-2">
                <div>
                  <h5 className="text-sm font-bold text-sky-300">📌 Tira 1: ZDESCARGA (Cols AS:AT)</h5>
                  <p className="text-[11px] text-slate-400">
                    [TOTAL BULTOS DESCARGA | UNIDADES EN BULTOS] · {activeItems.length} filas
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => copyText(tsvAsAt, 1)}
                  className="flex items-center gap-1.5 rounded-md bg-sky-600 px-3 py-1.5 text-xs font-bold text-white hover:bg-sky-500"
                >
                  {copied1 ? <Check className="h-3.5 w-3.5" /> : <Copy className="h-3.5 w-3.5" />}
                  {copied1 ? '¡Copiado!' : 'Copiar AS13'}
                </button>
              </div>
              <textarea
                readOnly
                rows={6}
                value={tsvAsAt}
                className="w-full rounded border border-slate-700 bg-slate-950 p-2.5 font-mono text-xs text-emerald-300"
              />
            </div>

            <div className="space-y-2 rounded-lg border border-slate-700/60 bg-[#1E293B] p-4">
              <div className="flex items-center justify-between gap-2">
                <div>
                  <h5 className="text-sm font-bold text-sky-300">📌 Tira 2: PACKING LIST (Cols AW:AY)</h5>
                  <p className="text-[11px] text-slate-400">
                    [Bultos Packing List | UNIDADES | Empaque Canónico] · {activeItems.length} filas
                  </p>
                </div>
                <button
                  type="button"
                  onClick={() => copyText(tsvAwAy, 2)}
                  className="flex items-center gap-1.5 rounded-md bg-sky-600 px-3 py-1.5 text-xs font-bold text-white hover:bg-sky-500"
                >
                  {copied2 ? <Check className="h-3.5 w-3.5" /> : <Copy className="h-3.5 w-3.5" />}
                  {copied2 ? '¡Copiado!' : 'Copiar AW13'}
                </button>
              </div>
              <textarea
                readOnly
                rows={6}
                value={tsvAwAy}
                className="w-full rounded border border-slate-700 bg-slate-950 p-2.5 font-mono text-xs text-emerald-300"
              />
            </div>
          </div>

          {/* Export Buttons */}
          <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
            <button
              type="button"
              onClick={() =>
                exportRowsToExcel(
                  descargaRows,
                  'ZDESCARGA_Y_PACKING_LIST',
                  `${summary.import_id}_ZDESCARGA_PACKING_LIST_AS_AZ.xlsx`
                )
              }
              className="flex items-center justify-center gap-2 rounded-md border border-slate-600 bg-slate-800 px-4 py-3 text-xs font-bold text-white transition hover:border-sky-400 hover:text-sky-300"
            >
              <Download className="h-4 w-4" />
              {lang === 'es' ? 'Descargar Tiras (.xlsx)' : 'Download Strips (.xlsx)'}
            </button>

            <button
              type="button"
              onClick={() => exportOfficialSquaredWorkbook(audit, excludeZeroTotal)}
              className="flex items-center justify-center gap-2 rounded-md border border-sky-500/50 bg-sky-900/40 px-4 py-3 text-xs font-bold text-sky-200 transition hover:bg-sky-900/60"
            >
              <Download className="h-4 w-4" />
              {lang === 'es' ? 'Resumen Multi-Hoja (.xlsx)' : 'Multi-Sheet Summary (.xlsx)'}
            </button>

            <button
              type="button"
              disabled={exportingXlsm}
              onClick={handleMacroSafeExport}
              className="flex items-center justify-center gap-2 rounded-md bg-gradient-to-br from-emerald-600 to-teal-700 px-4 py-3 text-xs font-bold text-white shadow-md transition hover:from-emerald-500 hover:to-teal-600 disabled:opacity-50"
            >
              <Zap className="h-4 w-4" />
              {exportingXlsm
                ? lang === 'es'
                  ? 'Verificando SHA-256 VBA...'
                  : 'Verifying VBA SHA-256...'
                : audit.cierreRawBytes
                  ? lang === 'es'
                    ? 'Exportar .xlsm Macro-Safe (Preserva VBA)'
                    : 'Export Macro-Safe .xlsm (Preserve VBA)'
                  : lang === 'es'
                    ? 'Descargar Libro Oficial Cuadrado (.xlsx)'
                    : 'Download Official Squared Workbook (.xlsx)'}
            </button>
          </div>
        </div>
      )}
    </div>
  );
};
