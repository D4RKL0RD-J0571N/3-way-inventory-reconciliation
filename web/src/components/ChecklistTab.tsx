import React, { useState } from 'react';
import type { Language, ReconciliationSummary } from '../types/reconciliation';
import { t } from '../utils/i18n';

interface ChecklistTabProps {
  summary: ReconciliationSummary;
  lang: Language;
}

export const ChecklistTab: React.FC<ChecklistTabProps> = ({ summary, lang }) => {
  const [checks, setChecks] = useState<boolean[]>([true, true, true, true, true, true, true]);

  const items = [
    lang === 'es'
      ? "1. Portada de Proveedor (Hoja 'PORTADA') — Carátula oficial con nombre del proveedor y código de importación."
      : "1. Vendor Cover Page ('PORTADA' sheet) — Official cover with vendor name and import ID.",
    lang === 'es'
      ? `2. Datos Generales (Hoja 'DATOS GENERALES') — Resumen con POs (${summary.total_ped_43_units.toLocaleString()} UN), Factura 41 (${summary.total_ped_41_units.toLocaleString()} UN) y totales.`
      : `2. General Data ('DATOS GENERALES' sheet) — Summary with POs (${summary.total_ped_43_units.toLocaleString()} UN), Order 41 (${summary.total_ped_41_units.toLocaleString()} UN) and totals.`,
    lang === 'es'
      ? `3. Carátula del Cierre Oficial (Hoja 'CIERRE RECEP') — Balance general con unidades recibidas (${summary.total_received_units.toLocaleString()} UN) y bultos.`
      : `3. Official Closure Cover ('CIERRE RECEP' sheet) — General balance with received units (${summary.total_received_units.toLocaleString()} UN) and cartons.`,
    lang === 'es'
      ? `4. Formulario de Nota de Crédito (Hoja 'PARA N.C RECEP.') — Estado: ${
          summary.total_ncr_amount === 0
            ? 'LIMPIO ($0.00)'
            : `RECLAMO POR $${summary.total_ncr_amount.toLocaleString(undefined, { minimumFractionDigits: 2 })}`
        }.`
      : `4. Credit Note Claim Form ('PARA N.C RECEP.' sheet) — Status: ${
          summary.total_ncr_amount === 0
            ? 'CLEAN ($0.00)'
            : `CLAIM FOR $${summary.total_ncr_amount.toLocaleString(undefined, { minimumFractionDigits: 2 })}`
        }.`,
    lang === 'es'
      ? '5. Respaldo en ERP de ingreso pedido 41 (Mov 101 en MB51) — Impresión de MB51, clases 101-102.'
      : '5. ERP Goods Receipt Confirmation (Movement 101 in MB51) — MB51 report printout, movements 101-102.',
    lang === 'es'
      ? "6. Lista de Empaque de Bultos y Unidades en Excel (Hoja 'Bultos') — Desglose físico de cartones enteros y complementos."
      : "6. Carton Packing List ('Bultos' sheet) — Physical breakdown of full master cartons and loose packs.",
    lang === 'es'
      ? '7. DUA (Documentación Única Aduanera) — Copia aduanera sellada con peso bruto y valor CIF.'
      : '7. Customs Clearance (DUA) — Stamped customs declaration with gross weight and CIF value.',
  ];

  const completed = checks.filter(Boolean).length;
  const pct = Math.round((completed / items.length) * 100);

  return (
    <div className="space-y-5">
      <div>
        <h3 className="text-lg font-bold text-white">{t('checklist_title', lang)}</h3>
        <p className="text-xs text-slate-400">
          {lang === 'es'
            ? 'Marca cada documento conforme lo imprimas para asegurar que la carpeta física esté en el orden reglamentario:'
            : 'Check off each document as you print it to ensure the physical binder is organized in proper order:'}
        </p>
      </div>

      <div className="space-y-2.5">
        {items.map((label, idx) => (
          <label
            key={idx}
            className="flex cursor-pointer items-start gap-3 rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5 transition hover:border-sky-500/40"
          >
            <input
              type="checkbox"
              checked={checks[idx]}
              onChange={() =>
                setChecks((prev) => prev.map((v, i) => (i === idx ? !v : v)))
              }
              className="mt-0.5 h-4 w-4 rounded border-slate-500 accent-sky-500"
            />
            <span className={`text-sm ${checks[idx] ? 'text-slate-100' : 'text-slate-400 line-through'}`}>
              {label}
            </span>
          </label>
        ))}
      </div>

      <div className="space-y-2">
        <div className="h-3 w-full overflow-hidden rounded-full bg-slate-800">
          <div
            className={`h-full transition-all duration-300 ${
              completed === items.length ? 'bg-emerald-500' : 'bg-sky-500'
            }`}
            style={{ width: `${pct}%` }}
          />
        </div>

        {completed === items.length ? (
          <div className="rounded-md border border-emerald-500/40 bg-emerald-500/15 p-3 text-sm font-semibold text-emerald-300">
            🎉{' '}
            {lang === 'es'
              ? '¡Expediente físico 100% completo! Todos los documentos impresos y listos para revisión.'
              : 'Physical binder 100% complete! All documents printed and ready for review.'}
          </div>
        ) : (
          <div className="rounded-md border border-sky-500/40 bg-sky-500/10 p-3 text-sm text-sky-200">
            {lang === 'es'
              ? `Progreso del expediente: ${completed} de ${items.length} documentos verificados (${pct}%).`
              : `Binder progress: ${completed} of ${items.length} documents verified (${pct}%).`}
          </div>
        )}
      </div>
    </div>
  );
};
