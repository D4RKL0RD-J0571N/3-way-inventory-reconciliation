import React, { useMemo, useState } from 'react';
import type { Language, PrepackItem } from '../types/reconciliation';
import { t } from '../utils/i18n';

interface CurvesTabProps {
  items: PrepackItem[];
  lang: Language;
}

export const CurvesTab: React.FC<CurvesTabProps> = ({ items, lang }) => {
  const parentsWithBom = useMemo(() => items.filter((i) => i.children.length > 0), [items]);
  const [selectedCode, setSelectedCode] = useState<string>(parentsWithBom[0]?.material_code ?? '');

  const selectedItem =
    parentsWithBom.find((i) => i.material_code === selectedCode) ?? parentsWithBom[0] ?? null;

  return (
    <div className="space-y-5">
      <div>
        <h3 className="text-lg font-bold text-white">{t('curves_title', lang)}</h3>
        <p className="text-xs text-slate-400">{t('curves_caption', lang)}</p>
      </div>

      {selectedItem ? (
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-xs font-semibold text-slate-400">
              {lang === 'es' ? 'Selecciona Prepack:' : 'Select Prepack:'}
            </label>
            <select
              value={selectedItem.material_code}
              onChange={(e) => setSelectedCode(e.target.value)}
              className="w-full rounded-md border border-slate-600 bg-slate-900 px-3 py-2.5 text-sm font-semibold text-white"
            >
              {parentsWithBom.map((i) => (
                <option key={i.material_code} value={i.material_code}>
                  📦 {i.material_code} | {i.vendor_material} | {i.description} (Curva {i.curve_total} UN)
                </option>
              ))}
            </select>
          </div>

          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
            <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5">
              <p className="text-[11px] font-bold text-slate-400 uppercase">
                {lang === 'es' ? 'PREPACKS FACTURADOS' : 'INVOICED PREPACKS'}
              </p>
              <p className="mt-1 text-xl font-extrabold text-white">
                {Math.round(selectedItem.quantity_inv_41_prk)} PRK
              </p>
            </div>
            <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5">
              <p className="text-[11px] font-bold text-slate-400 uppercase">
                {lang === 'es' ? 'CURVA PREPACK (UN/PRK)' : 'PREPACK CURVE (UN/PRK)'}
              </p>
              <p className="mt-1 text-xl font-extrabold text-white">{selectedItem.curve_total} UN/PRK</p>
            </div>
            <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5">
              <p className="text-[11px] font-bold text-slate-400 uppercase">
                {lang === 'es' ? 'TOTAL PIEZAS ESPERADAS' : 'TOTAL EXPECTED UNITS'}
              </p>
              <p className="mt-1 text-xl font-extrabold text-white">
                {selectedItem.quantity_inv_41_units.toLocaleString()} UN
              </p>
            </div>
            <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5">
              <p className="text-[11px] font-bold text-slate-400 uppercase">
                {lang === 'es' ? 'ESTADO LÍNEA' : 'LINE STATUS'}
              </p>
              <p className="mt-1 text-xl font-extrabold text-sky-300">{selectedItem.status}</p>
            </div>
          </div>

          <div>
            <h4 className="mb-2 text-sm font-bold text-slate-100">
              {lang === 'es' ? 'Reparto Matemático por Talla:' : 'Mathematical Breakdown by Size:'}
            </h4>
            <div className="overflow-x-auto rounded-lg border border-slate-700/70 bg-slate-900/60">
              <table className="min-w-full divide-y divide-slate-800 text-left text-xs">
                <thead className="bg-slate-800 text-[11px] font-bold text-slate-300 uppercase">
                  <tr>
                    <th className="px-3 py-2.5">{lang === 'es' ? 'Código SKU (Hijo)' : 'Child SKU Code'}</th>
                    <th className="px-3 py-2.5">{lang === 'es' ? 'Descripción / Talla' : 'Description / Size'}</th>
                    <th className="px-3 py-2.5">{lang === 'es' ? 'Ratio Curva' : 'Curve Ratio'}</th>
                    <th className="px-3 py-2.5">{lang === 'es' ? 'Unidades Teóricas' : 'Theoretical Units'}</th>
                    <th className="px-3 py-2.5">{lang === 'es' ? 'Unidades Físicas' : 'Physical Received'}</th>
                    <th className="px-3 py-2.5">{lang === 'es' ? 'Diferencia' : 'Difference'}</th>
                    <th className="px-3 py-2.5">{lang === 'es' ? 'Costo Unitario ($)' : 'Unit Cost ($)'}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800/70 text-slate-200">
                  {selectedItem.children.map((c) => (
                    <tr key={c.material_code} className="hover:bg-slate-800/50">
                      <td className="px-3 py-2 font-mono text-sky-300">{c.material_code}</td>
                      <td className="px-3 py-2">{c.description}</td>
                      <td className="px-3 py-2 font-bold">{c.ratio}</td>
                      <td className="px-3 py-2">{c.theoretical_units.toLocaleString()}</td>
                      <td className="px-3 py-2">{c.received_units.toLocaleString()}</td>
                      <td
                        className={`px-3 py-2 font-bold ${
                          c.difference === 0 ? 'text-emerald-400' : 'text-red-400'
                        }`}
                      >
                        {c.difference > 0 ? `+${c.difference}` : c.difference}
                      </td>
                      <td className="px-3 py-2">${c.unit_cost.toFixed(2)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      ) : (
        <div className="rounded-md border border-sky-500/40 bg-sky-500/10 p-4 text-sm text-sky-200">
          {lang === 'es'
            ? 'No se encontraron curvas o desgloses de tallas (BOMs) para los materiales de este libro.'
            : 'No size curves or BOM breakdowns were found for the materials in this workbook.'}
        </div>
      )}
    </div>
  );
};
