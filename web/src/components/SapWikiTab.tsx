import React, { useMemo, useState } from 'react';
import { BookOpen, Search, Wrench } from 'lucide-react';
import type { Language } from '../types/reconciliation';
import { SAP_DIAGNOSTIC_GUIDES, SAP_TRANSACTIONS } from '../data/sapCatalog';

interface SapWikiTabProps {
  lang: Language;
}

export const SapWikiTab: React.FC<SapWikiTabProps> = ({ lang }) => {
  const [subtab, setSubtab] = useState<'sap' | 'guides' | 'wiki'>('sap');
  const [sapSearch, setSapSearch] = useState('');

  const filteredTx = useMemo(() => {
    const q = sapSearch.toLowerCase().trim();
    if (!q) return SAP_TRANSACTIONS;
    return SAP_TRANSACTIONS.filter(
      (tx) =>
        tx.code.toLowerCase().includes(q) ||
        tx.title.toLowerCase().includes(q) ||
        tx.purpose.toLowerCase().includes(q) ||
        tx.whenToUse.toLowerCase().includes(q)
    );
  }, [sapSearch]);

  return (
    <div className="space-y-5">
      <div>
        <h3 className="text-lg font-bold text-white">
          {lang === 'es'
            ? '📖 Enciclopedia SAP, Guías Diagnósticas & Manual Operativo'
            : '📖 SAP T-Code Encyclopedia, Diagnostic Guides & Operating Wiki'}
        </h3>
        <p className="text-xs text-slate-400">
          {lang === 'es'
            ? 'Referencia offline de transacciones SAP (ME80FN, SE16N, MB51, VL31N, RFUI, MIGO), protocolos de investigación y flujo paso a paso:'
            : 'Offline reference of SAP T-Codes (ME80FN, SE16N, MB51, VL31N, RFUI, MIGO), troubleshooting guides, and step-by-step SOP:'}
        </p>
      </div>

      <div className="flex flex-wrap gap-2 border-b border-slate-700/60 pb-2">
        <button
          type="button"
          onClick={() => setSubtab('sap')}
          className={`flex items-center gap-1.5 rounded-t-md px-4 py-2 text-xs font-bold transition ${
            subtab === 'sap'
              ? 'border-b-2 border-sky-400 bg-slate-800 text-sky-300'
              : 'text-slate-400 hover:bg-slate-800/50'
          }`}
        >
          <BookOpen className="h-3.5 w-3.5" />
          {lang === 'es' ? `Transacciones y Tablas SAP (${SAP_TRANSACTIONS.length})` : `SAP T-Codes & Tables (${SAP_TRANSACTIONS.length})`}
        </button>
        <button
          type="button"
          onClick={() => setSubtab('guides')}
          className={`flex items-center gap-1.5 rounded-t-md px-4 py-2 text-xs font-bold transition ${
            subtab === 'guides'
              ? 'border-b-2 border-sky-400 bg-slate-800 text-sky-300'
              : 'text-slate-400 hover:bg-slate-800/50'
          }`}
        >
          <Wrench className="h-3.5 w-3.5" />
          {lang === 'es' ? `Guías de Diagnóstico (${SAP_DIAGNOSTIC_GUIDES.length})` : `Diagnostic Guides (${SAP_DIAGNOSTIC_GUIDES.length})`}
        </button>
        <button
          type="button"
          onClick={() => setSubtab('wiki')}
          className={`rounded-t-md px-4 py-2 text-xs font-bold transition ${
            subtab === 'wiki'
              ? 'border-b-2 border-sky-400 bg-slate-800 text-sky-300'
              : 'text-slate-400 hover:bg-slate-800/50'
          }`}
        >
          📘 {lang === 'es' ? 'Manual Operativo & Casos (ES/EN)' : 'Operating Guide & SOP (ES/EN)'}
        </button>
      </div>

      {subtab === 'sap' && (
        <div className="space-y-4">
          <div className="relative max-w-md">
            <Search className="pointer-events-none absolute top-2.5 left-3 h-4 w-4 text-slate-400" />
            <input
              type="text"
              value={sapSearch}
              onChange={(e) => setSapSearch(e.target.value)}
              placeholder={
                lang === 'es'
                  ? 'Buscar transacción o tabla (ej. ME80FN, MB51, STPO, RFUI)...'
                  : 'Search T-Code or table (e.g. ME80FN, MB51, STPO, RFUI)...'
              }
              className="w-full rounded-md border border-slate-600 bg-slate-900 py-2 pr-3 pl-9 text-sm text-white placeholder-slate-400"
            />
          </div>

          <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
            {filteredTx.map((tx) => (
              <div
                key={tx.code}
                className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-4 shadow"
              >
                <div className="flex items-center justify-between gap-2">
                  <span className="rounded bg-sky-500/20 px-2.5 py-0.5 font-mono text-sm font-extrabold text-sky-300">
                    {tx.code}
                  </span>
                  <span className="rounded border border-slate-600 bg-slate-800 px-2 py-0.5 text-[10px] font-semibold text-slate-300">
                    {tx.objectKind}
                  </span>
                </div>
                <h4 className="mt-2 text-sm font-bold text-white">{tx.title}</h4>
                <p className="mt-1 text-xs leading-relaxed text-slate-300">{tx.purpose}</p>
                <p className="mt-2 text-xs text-emerald-300">
                  <span className="font-bold">Cuándo usar:</span> {tx.whenToUse}
                </p>
                {tx.navigationPath && (
                  <p className="mt-1.5 font-mono text-[11px] text-slate-400">{tx.navigationPath}</p>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      {subtab === 'guides' && (
        <div className="space-y-4">
          {SAP_DIAGNOSTIC_GUIDES.map((guide) => (
            <div
              key={guide.guideId}
              className="rounded-lg border border-slate-700/60 border-l-4 border-l-sky-500 bg-[#1E293B] p-4 shadow"
            >
              <div className="flex flex-wrap items-center justify-between gap-2">
                <h4 className="text-sm font-bold text-white">
                  {guide.guideId} — {guide.title}
                </h4>
                <div className="flex flex-wrap gap-1.5">
                  {guide.transactionCodes.map((code) => (
                    <span
                      key={code}
                      className="rounded bg-sky-500/20 px-2 py-0.5 font-mono text-[11px] font-bold text-sky-300"
                    >
                      {code}
                    </span>
                  ))}
                </div>
              </div>
              <p className="mt-1.5 text-xs text-slate-300">{guide.problemSummary}</p>
              <ul className="mt-3 space-y-1.5 text-xs text-slate-200">
                {guide.investigationSteps.map((step, idx) => (
                  <li key={idx} className="rounded bg-slate-900/60 px-3 py-2">
                    {step}
                  </li>
                ))}
              </ul>
            </div>
          ))}
        </div>
      )}

      {subtab === 'wiki' && (
        <div className="grid grid-cols-1 gap-5 lg:grid-cols-2">
          <div className="space-y-3 rounded-lg border border-slate-700/60 bg-[#1E293B] p-5 text-xs leading-relaxed text-slate-300">
            <h4 className="text-base font-bold text-sky-300">🇪🇸 Flujo de Trabajo: Cómo Cerrar una Importación</h4>
            <p className="font-bold text-white">1. Los 3 Niveles Numéricos Independientes:</p>
            <ul className="list-inside list-disc space-y-1">
              <li>
                <strong>Nivel 1 — Bultos Físicos:</strong> Cajas manipuladas en muelle (Col AS), Packing List (Col AW) y Control de Calidad (Enteros + Complementos).
              </li>
              <li>
                <strong>Nivel 2 — Prepacks (PRK / Curva):</strong> Unidad de surtido comercial (Cols BD/BE, Curva BB, UM BC).
              </li>
              <li>
                <strong>Nivel 3 — Unidades Totales:</strong> Prendas facturadas y nacionalizadas (Cols F, G, H, AT, AX).
              </li>
            </ul>
            <p className="font-bold text-white">2. Trampas Frecuentes en Bodega que ACR Resuelve:</p>
            <ul className="list-inside list-disc space-y-1">
              <li>
                <strong>Desplazamiento de Columnas:</strong> Detección dinámica por nombres de encabezado en fila 6 y fila 12 (inmune a columnas extra como Tarima o Ubicación).
              </li>
              <li>
                <strong>Fórmulas Borradas en Columna H:</strong> Recuperación automática desde Col AT y Control de Calidad.
              </li>
              <li>
                <strong>Montaje a Ciegas (Modo Pausa):</strong> Bloqueo automático de NCR si aún no existe conteo físico de muelle.
              </li>
            </ul>
          </div>

          <div className="space-y-3 rounded-lg border border-slate-700/60 bg-[#1E293B] p-5 text-xs leading-relaxed text-slate-300">
            <h4 className="text-base font-bold text-sky-300">🇬🇧 3-Way Reconciliation Engine: Standard Operating Procedure</h4>
            <p className="font-bold text-white">1. What is 3-Way Reconciliation?</p>
            <ul className="list-inside list-disc space-y-1">
              <li>
                <strong>Purchase Orders (PO 43):</strong> Goods requested from international suppliers.
              </li>
              <li>
                <strong>Nationalization Invoice (Order 41 / DUA):</strong> Quantities officially declared and billed.
              </li>
              <li>
                <strong>Dock Physical Receipts (Control de Calidad):</strong> Actual cartons, prepacks, and loose units unloaded and tallied at warehouse docks.
              </li>
            </ul>
            <p className="font-bold text-white">2. Key Operational Capabilities:</p>
            <ul className="list-inside list-disc space-y-1">
              <li>
                <strong>Prepack BOM Curve Explosion:</strong> Automatically expands parent assortment codes into child SKU size breakdowns.
              </li>
              <li>
                <strong>10-Second Square Strips:</strong> Generates TSV paste blocks for Columns AS:AT and AW:AY so variance columns drop to 0 instantly.
              </li>
              <li>
                <strong>100% Local Memory Processing:</strong> Zero external server transmission or cloud leakage.
              </li>
            </ul>
          </div>
        </div>
      )}
    </div>
  );
};
