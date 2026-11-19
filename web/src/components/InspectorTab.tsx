import React, { useMemo, useState } from 'react';
import { Search } from 'lucide-react';
import type { Language, PrepackItem, ValidationReport } from '../types/reconciliation';
import { filterItems } from '../utils/engine';
import { t } from '../utils/i18n';

interface InspectorTabProps {
  items: PrepackItem[];
  validation: ValidationReport;
  lang: Language;
}

export const InspectorTab: React.FC<InspectorTabProps> = ({ items, validation, lang }) => {
  const [query, setQuery] = useState('');
  const [partFilter, setPartFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [catFilter, setCatFilter] = useState('ALL');
  const [hideZeroTotal, setHideZeroTotal] = useState(true);
  const [selectedCode, setSelectedCode] = useState<string>(items[0]?.material_code ?? '');

  const distinctParts = useMemo(
    () => Array.from(new Set(items.map((i) => (i.container_part || '').trim().toUpperCase()).filter(Boolean))).sort(),
    [items]
  );
  const distinctStatuses = useMemo(
    () => Array.from(new Set(items.map((i) => i.discrepancy_label || i.status).filter(Boolean))).sort(),
    [items]
  );
  const distinctCats = useMemo(
    () => Array.from(new Set(items.map((i) => i.category).filter((c) => c && c !== '-'))).sort(),
    [items]
  );

  const matching = useMemo(
    () => filterItems(items, partFilter, statusFilter, catFilter, query, hideZeroTotal),
    [items, partFilter, statusFilter, catFilter, query, hideZeroTotal]
  );

  const selectedItem = useMemo(() => {
    return matching.find((m) => m.material_code === selectedCode) ?? matching[0] ?? null;
  }, [matching, selectedCode]);

  const itemIssues = useMemo(() => {
    if (!selectedItem) return [];
    return [...validation.ERROR, ...validation.WARNING, ...validation.INFO].filter(
      (iss) => iss.material_code === selectedItem.material_code
    );
  }, [selectedItem, validation]);

  const statusBannerStyle = (item: PrepackItem) => {
    if (item.is_zero_total || item.discrepancy_type === 'ZeroTotal')
      return 'border-slate-600 border-l-slate-400 bg-slate-800/80 text-slate-300';
    switch (item.discrepancy_type) {
      case 'ExactMatch':
        return 'border-emerald-500/40 border-l-emerald-500 bg-emerald-500/15 text-emerald-300';
      case 'PendingPart':
      case 'CartonMismatchOnly':
        return 'border-blue-500/40 border-l-blue-500 bg-blue-500/15 text-blue-300';
      case 'Overage':
      case 'UnexpectedItem':
        return 'border-purple-500/40 border-l-purple-500 bg-purple-500/15 text-purple-300';
      case 'RevueltoCandidate':
      case 'BlindSetup':
        return 'border-amber-500/40 border-l-amber-500 bg-amber-500/15 text-amber-200';
      default:
        return 'border-red-500/40 border-l-red-500 bg-red-500/15 text-red-300';
    }
  };

  return (
    <div className="space-y-5">
      <div className="flex flex-wrap items-start justify-between gap-2">
        <div>
          <h3 className="text-lg font-bold text-white">{t('inspector_title', lang)}</h3>
          <p className="text-xs text-slate-400">
            {lang === 'es'
              ? 'Inspecciona cualquier material padre, su empaque canónico, sus curvas de compra y su balance físico sin desplazarte entre hojas:'
              : 'Inspect any parent style, canonical packaging, purchase curves, and dock physical receipts on a single screen:'}
          </p>
        </div>

        <label className="flex cursor-pointer items-center gap-2 rounded-md border border-slate-700 bg-slate-800/80 px-3 py-1.5 text-xs font-semibold text-slate-200">
          <input
            type="checkbox"
            checked={hideZeroTotal}
            onChange={(e) => setHideZeroTotal(e.target.checked)}
            className="rounded border-slate-600 bg-slate-900 accent-sky-500"
          />
          <span>{lang === 'es' ? 'Ocultar en cero (0 total)' : 'Hide zero-total'}</span>
        </label>
      </div>

      {/* Filters */}
      <div className="grid grid-cols-1 gap-3 md:grid-cols-4">
        <div className="relative md:col-span-2">
          <Search className="pointer-events-none absolute top-2.5 left-3 h-4 w-4 text-slate-400" />
          <input
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder={
              lang === 'es'
                ? 'Buscar código SAP, estilo proveedor, PO o descripción...'
                : 'Search SAP code, vendor style, PO or description...'
            }
            className="w-full rounded-md border border-slate-600 bg-slate-900 py-2 pr-3 pl-9 text-sm text-white placeholder-slate-400"
          />
        </div>

        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
          className="rounded-md border border-slate-600 bg-slate-900 px-3 py-2 text-sm text-white"
        >
          <option value="ALL">
            {lang === 'es' ? `Todas Clasificaciones` : `All Classifications`}
          </option>
          {distinctStatuses.map((s) => (
            <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>

        <div className="flex gap-2">
          {distinctParts.length > 1 && (
            <select
              value={partFilter}
              onChange={(e) => setPartFilter(e.target.value)}
              className="flex-1 rounded-md border border-slate-600 bg-slate-900 px-2.5 py-2 text-sm text-white"
            >
              <option value="ALL">{lang === 'es' ? 'Todas Partes' : 'All Parts'}</option>
              {distinctParts.map((p) => (
                <option key={p} value={p}>
                  {lang === 'es' ? `Parte ${p}` : `Part ${p}`}
                </option>
              ))}
            </select>
          )}
          {distinctCats.length > 1 && (
            <select
              value={catFilter}
              onChange={(e) => setCatFilter(e.target.value)}
              className="flex-1 rounded-md border border-slate-600 bg-slate-900 px-2.5 py-2 text-sm text-white"
            >
              <option value="ALL">{lang === 'es' ? 'Categorías' : 'Categories'}</option>
              {distinctCats.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
          )}
        </div>
      </div>

      {/* Material Selector */}
      {matching.length > 0 && selectedItem ? (
        <div className="space-y-4">
          <div>
            <label className="mb-1 block text-xs font-semibold text-slate-400">
              {lang === 'es'
                ? `Selecciona un estilo (${matching.length} coincidencias disponibles):`
                : `Select a style (${matching.length} matches available):`}
            </label>
            <select
              value={selectedItem.material_code}
              onChange={(e) => setSelectedCode(e.target.value)}
              className="w-full rounded-md border border-sky-500/40 bg-slate-900 px-3 py-2.5 text-sm font-semibold text-white"
            >
              {matching.map((it) => {
                const icon = it.is_zero_total
                  ? '⬜'
                  : it.discrepancy_type === 'ExactMatch'
                    ? '🟢'
                    : it.discrepancy_type === 'PendingPart'
                      ? '🔵'
                      : it.discrepancy_type === 'RevueltoCandidate'
                        ? '🟠'
                        : it.discrepancy_type === 'Overage' || it.discrepancy_type === 'UnexpectedItem'
                          ? '🟣'
                          : '🔴';
                return (
                  <option key={`${it.material_code}-${it.source_row}`} value={it.material_code}>
                    {icon} [F.{it.source_row}] {it.material_code} | {it.vendor_material} | {it.description}
                    {it.container_part ? ` (Parte ${it.container_part})` : ''}
                  </option>
                );
              })}
            </select>
          </div>

          {/* Status Banner */}
          <div
            className={`rounded-md border border-l-4 px-4 py-3 text-sm font-semibold ${statusBannerStyle(
              selectedItem
            )}`}
          >
            <div className="flex flex-wrap items-center justify-between gap-2">
              <span>
                {selectedItem.discrepancy_label} — {selectedItem.observations}
              </span>
              {selectedItem.ncr_amount > 0 && (
                <span className="font-extrabold">
                  NCR: ${selectedItem.ncr_amount.toLocaleString(undefined, { minimumFractionDigits: 2 })} USD
                </span>
              )}
            </div>
            {selectedItem.quantity_flags.length > 0 && (
              <div className="mt-2 flex flex-wrap gap-1.5">
                {selectedItem.quantity_flags.map((flag) => (
                  <span
                    key={flag}
                    className="rounded border border-sky-400/30 bg-slate-900/70 px-2 py-0.5 font-mono text-[10px] text-sky-300"
                  >
                    {flag}
                  </span>
                ))}
              </div>
            )}
          </div>

          {/* 4 Detail Cards */}
          <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-4">
            <div className="rounded-lg border border-slate-700/60 border-t-4 border-t-sky-500 bg-[#1E293B] p-4 shadow">
              <h4 className="border-b border-slate-700 pb-1.5 text-xs font-bold tracking-wider text-sky-400 uppercase">
                {lang === 'es' ? '1. Identificación' : '1. Identification'}
              </h4>
              <div className="mt-3 space-y-2 text-xs">
                <Row label={lang === 'es' ? 'Fila CIERRE RECEP:' : 'Excel Row:'} value={`Fila ${selectedItem.source_row}`} />
                <Row label={lang === 'es' ? 'Código SAP:' : 'SAP Code:'} value={selectedItem.material_code} />
                <Row label={lang === 'es' ? 'Estilo Prov:' : 'Vendor Style:'} value={selectedItem.vendor_material || '—'} />
                <Row label={lang === 'es' ? 'Categoría:' : 'Category:'} value={selectedItem.category || '—'} />
                <Row label={lang === 'es' ? 'Descripción:' : 'Description:'} value={selectedItem.description} />
                <Row
                  label={lang === 'es' ? 'Contenedor:' : 'Container:'}
                  value={
                    selectedItem.container_part
                      ? `${lang === 'es' ? 'Parte' : 'Part'} ${selectedItem.container_part}`
                      : lang === 'es'
                        ? 'Contenedor Único'
                        : 'Single Container'
                  }
                />
              </div>
            </div>

            <div className="rounded-lg border border-slate-700/60 border-t-4 border-t-sky-500 bg-[#1E293B] p-4 shadow">
              <h4 className="border-b border-slate-700 pb-1.5 text-xs font-bold tracking-wider text-sky-400 uppercase">
                {lang === 'es' ? '2. Empaque & Compras (PO 43)' : '2. Packaging & PO 43'}
              </h4>
              <div className="mt-3 space-y-2 text-xs">
                <Row label={lang === 'es' ? 'Orden Compra:' : 'Purchase Order:'} value={selectedItem.order_po_43 || '—'} />
                <Row label="Prepacks (PRK):" value={`${selectedItem.quantity_po_43_prk.toLocaleString()} PRK`} />
                <Row
                  label={lang === 'es' ? 'Bultos Físicos:' : 'Physical Cartons:'}
                  value={`${selectedItem.bultos_descarga.toLocaleString()} Bultos`}
                />
                <Row
                  label={lang === 'es' ? 'Empaque Canónico:' : 'Canonical Pack:'}
                  value={selectedItem.canonical_packaging}
                />
                <Row
                  label={lang === 'es' ? 'Enteros / Pico / Sueltas:' : 'Full / Odd / Loose:'}
                  value={`${selectedItem.bultos_enteros}E / ${selectedItem.bultos_complementos}P / +${selectedItem.loose_units}`}
                />
                <Row label={lang === 'es' ? 'Curva Prepack:' : 'Prepack Curve:'} value={`${selectedItem.curve_total} UN/PRK`} />
              </div>
            </div>

            <div className="rounded-lg border border-slate-700/60 border-t-4 border-t-sky-500 bg-[#1E293B] p-4 shadow">
              <h4 className="border-b border-slate-700 pb-1.5 text-xs font-bold tracking-wider text-sky-400 uppercase">
                {lang === 'es' ? '3. Factura 41 & Físico' : '3. Invoice 41 & Dock'}
              </h4>
              <div className="mt-3 space-y-2 text-xs">
                <Row
                  label={lang === 'es' ? 'Facturado (41):' : 'Invoiced (41):'}
                  value={`${selectedItem.quantity_inv_41_units.toLocaleString()} UN`}
                />
                <Row
                  label={lang === 'es' ? 'Recibido Muelle:' : 'Dock Received:'}
                  value={`${selectedItem.received_real_units.toLocaleString()} UN`}
                />
                <Row
                  label={lang === 'es' ? 'Costo Unitario:' : 'Unit Cost:'}
                  value={`$${selectedItem.unit_cost.toFixed(2)} USD`}
                />
                <Row
                  label={lang === 'es' ? 'Total Factura:' : 'Total Invoice:'}
                  value={`$${(selectedItem.quantity_inv_41_units * selectedItem.unit_cost).toLocaleString(undefined, { minimumFractionDigits: 2 })} USD`}
                />
                <Row
                  label={lang === 'es' ? 'Diferencia:' : 'Discrepancy:'}
                  value={`${Math.round(selectedItem.received_real_units - selectedItem.quantity_inv_41_units)} UN`}
                />
              </div>
            </div>

            <div className="rounded-lg border border-slate-700/60 border-t-4 border-t-sky-500 bg-[#1E293B] p-4 shadow">
              <h4 className="border-b border-slate-700 pb-1.5 text-xs font-bold tracking-wider text-sky-400 uppercase">
                {lang === 'es' ? '4. Balances y Ajustes' : '4. Balances & Adjustments'}
              </h4>
              <div className="mt-3 space-y-2 text-xs">
                <Row label="SAP Mov 101:" value={`${selectedItem.sap_mov_101.toLocaleString()} UN`} />
                <Row label={lang === 'es' ? 'Faltante:' : 'Shortage:'} value={`${selectedItem.faltante.toLocaleString()} UN`} />
                <Row label={lang === 'es' ? 'Dañado / Avería:' : 'Damaged:'} value={`${selectedItem.danado.toLocaleString()} UN`} />
                <Row label={lang === 'es' ? 'Sobrante:' : 'Surplus:'} value={`${selectedItem.sobrante.toLocaleString()} UN`} />
                {selectedItem.revuelto_partner && (
                  <Row
                    label={lang === 'es' ? 'Par Revuelto:' : 'Revuelto Pair:'}
                    value={selectedItem.revuelto_partner}
                  />
                )}
                <Row
                  label={lang === 'es' ? 'Monto NCR:' : 'NCR Amount:'}
                  value={`$${selectedItem.ncr_amount.toLocaleString(undefined, { minimumFractionDigits: 2 })} USD`}
                />
              </div>
            </div>
          </div>

          {/* Integrity Validation Alerts for this material */}
          {itemIssues.length > 0 && (
            <div className="space-y-2">
              {itemIssues.map((iss, idx) => (
                <div
                  key={`${iss.issue_type}-${idx}`}
                  className={`rounded-md border px-3.5 py-2.5 text-xs ${
                    iss.severity === 'ERROR'
                      ? 'border-red-500/40 bg-red-500/10 text-red-200'
                      : iss.severity === 'WARNING'
                        ? 'border-amber-500/40 bg-amber-500/10 text-amber-200'
                        : 'border-sky-500/40 bg-sky-500/10 text-sky-200'
                  }`}
                >
                  <span className="font-bold">[{iss.severity}] {iss.issue_type}:</span> {iss.message}{' '}
                  <span className="opacity-80">— {iss.recommendation}</span>
                </div>
              ))}
            </div>
          )}

          {/* Child SKUs */}
          {selectedItem.children.length > 0 && (
            <div className="space-y-2">
              <h4 className="text-sm font-bold text-slate-100">
                📐{' '}
                {lang === 'es'
                  ? 'Desglose de Tallas Hijas ancladas a este Prepack:'
                  : 'Child SKU Breakdown linked to this Prepack:'}
              </h4>
              <div className="overflow-x-auto rounded-lg border border-slate-700/70 bg-slate-900/60">
                <table className="min-w-full divide-y divide-slate-800 text-left text-xs">
                  <thead className="bg-slate-800 text-[11px] font-bold text-slate-300 uppercase">
                    <tr>
                      <th className="px-3 py-2">{lang === 'es' ? 'Material Hijo (Talla)' : 'Child SKU'}</th>
                      <th className="px-3 py-2">{lang === 'es' ? 'Descripción / Talla' : 'Description / Size'}</th>
                      <th className="px-3 py-2">{lang === 'es' ? 'Ratio Curva' : 'Ratio'}</th>
                      <th className="px-3 py-2">{lang === 'es' ? 'Esperado (UN)' : 'Expected (UN)'}</th>
                      <th className="px-3 py-2">{lang === 'es' ? 'Recibido Físico (UN)' : 'Physical (UN)'}</th>
                      <th className="px-3 py-2">{lang === 'es' ? 'Diferencia' : 'Diff'}</th>
                      <th className="px-3 py-2">{lang === 'es' ? 'Costo Unit.' : 'Unit Cost'}</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/70 text-slate-200">
                    {selectedItem.children.map((c) => (
                      <tr key={c.material_code} className="hover:bg-slate-800/50">
                        <td className="px-3 py-2 font-mono">{c.material_code}</td>
                        <td className="px-3 py-2">{c.description}</td>
                        <td className="px-3 py-2">{c.ratio}</td>
                        <td className="px-3 py-2">{c.theoretical_units}</td>
                        <td className="px-3 py-2">{c.received_units}</td>
                        <td
                          className={`px-3 py-2 font-bold ${
                            c.difference === 0 ? 'text-emerald-400' : 'text-red-400'
                          }`}
                        >
                          {c.difference}
                        </td>
                        <td className="px-3 py-2">${c.unit_cost.toFixed(2)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>
      ) : (
        <div className="rounded-md border border-amber-500/40 bg-amber-500/10 p-4 text-sm text-amber-200">
          {lang === 'es'
            ? `No se encontraron materiales que coincidan con la búsqueda: '${query}'.`
            : `No materials matched your search: '${query}'.`}
        </div>
      )}
    </div>
  );
};

const Row: React.FC<{ label: string; value: string }> = ({ label, value }) => (
  <div className="flex items-start justify-between gap-2">
    <span className="shrink-0 text-slate-400">{label}</span>
    <span className="text-right font-bold text-slate-100">{value}</span>
  </div>
);
