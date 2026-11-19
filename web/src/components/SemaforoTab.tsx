import React, { useMemo, useState } from 'react';
import { Search } from 'lucide-react';
import type { Language, PrepackItem, ValidationReport } from '../types/reconciliation';
import { filterItems } from '../utils/engine';
import { t } from '../utils/i18n';

interface SemaforoTabProps {
  items: PrepackItem[];
  validation: ValidationReport;
  lang: Language;
}

export const SemaforoTab: React.FC<SemaforoTabProps> = ({ items, validation, lang }) => {
  const [hideZeroTotal, setHideZeroTotal] = useState(true);
  const [partFilter, setPartFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState('ALL');
  const [catFilter, setCatFilter] = useState('ALL');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const pageSize = 50;

  const zeroTotalCount = useMemo(() => items.filter((i) => i.is_zero_total).length, [items]);

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

  const filtered = useMemo(
    () => filterItems(items, partFilter, statusFilter, catFilter, search, hideZeroTotal),
    [items, partFilter, statusFilter, catFilter, search, hideZeroTotal]
  );

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const paginated = filtered.slice((page - 1) * pageSize, page * pageSize);

  const getRowBg = (item: PrepackItem) => {
    if (item.is_zero_total || item.discrepancy_type === 'ZeroTotal') {
      return 'bg-slate-800/40 text-slate-400 hover:bg-slate-800/60';
    }
    switch (item.discrepancy_type) {
      case 'ExactMatch':
        return 'bg-emerald-500/8 hover:bg-emerald-500/15';
      case 'PendingPart':
      case 'BlindSetup':
      case 'CartonMismatchOnly':
        return 'bg-blue-500/8 hover:bg-blue-500/15';
      case 'Overage':
      case 'UnexpectedItem':
        return 'bg-purple-500/10 hover:bg-purple-500/18';
      case 'RevueltoCandidate':
        return 'bg-amber-500/12 hover:bg-amber-500/20';
      default:
        return 'bg-red-500/15 hover:bg-red-500/20';
    }
  };

  const renderStatusBadge = (item: PrepackItem) => {
    if (item.is_zero_total || item.discrepancy_type === 'ZeroTotal') {
      return (
        <span className="inline-flex items-center rounded border border-slate-600 bg-slate-700/60 px-2 py-0.5 text-[11px] font-bold text-slate-300">
          SIN MOVIMIENTO (0 TOTAL)
        </span>
      );
    }
    switch (item.discrepancy_type) {
      case 'ExactMatch':
        return (
          <span className="inline-flex items-center rounded border border-emerald-500/40 bg-emerald-500/15 px-2 py-0.5 text-[11px] font-bold text-emerald-300">
            🟢 {item.discrepancy_label}
          </span>
        );
      case 'PendingPart':
        return (
          <span className="inline-flex items-center rounded border border-blue-500/40 bg-blue-500/15 px-2 py-0.5 text-[11px] font-bold text-blue-300">
            🔵 {item.discrepancy_label}
          </span>
        );
      case 'BlindSetup':
        return (
          <span className="inline-flex items-center rounded border border-amber-500/40 bg-amber-500/15 px-2 py-0.5 text-[11px] font-bold text-amber-300">
            🟡 {item.discrepancy_label}
          </span>
        );
      case 'CartonMismatchOnly':
        return (
          <span className="inline-flex items-center rounded border border-sky-500/40 bg-sky-500/15 px-2 py-0.5 text-[11px] font-bold text-sky-300">
            📦 {item.discrepancy_label}
          </span>
        );
      case 'Overage':
      case 'UnexpectedItem':
        return (
          <span className="inline-flex items-center rounded border border-purple-500/40 bg-purple-500/15 px-2 py-0.5 text-[11px] font-bold text-purple-300">
            🟣 {item.discrepancy_label}
          </span>
        );
      case 'RevueltoCandidate':
        return (
          <span className="inline-flex items-center gap-1 rounded border border-amber-500/40 bg-amber-500/15 px-2 py-0.5 text-[11px] font-bold text-amber-300">
            🟠 {item.discrepancy_label}
            {item.revuelto_partner && (
              <span className="font-mono text-[10px] text-amber-200">↔ {item.revuelto_partner}</span>
            )}
          </span>
        );
      default:
        return (
          <span className="inline-flex items-center rounded border border-red-500/40 bg-red-500/15 px-2 py-0.5 text-[11px] font-bold text-red-300">
            🔴 {item.discrepancy_label}
          </span>
        );
    }
  };

  const totalAlerts = validation.ERROR.length + validation.WARNING.length;

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-start justify-between gap-2">
        <div>
          <h3 className="text-lg font-bold text-white">{t('semaforo_title', lang)}</h3>
          <p className="text-xs text-slate-400">
            {lang === 'es'
              ? '🟢 Cuadrado | 🟠 Revuelto Candidato | 🟣 Sobrante | 🔴 Faltante / Dañado | 🔵 Parte Posterior | ⬜ Sin Movimiento (0 Total)'
              : '🟢 Exact Match | 🟠 Revuelto Candidate | 🟣 Overage | 🔴 Shortage / Damaged | 🔵 Pending Part | ⬜ Zero Total'}
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <label className="flex cursor-pointer items-center gap-2 rounded-md border border-slate-700 bg-slate-800/80 px-3 py-1.5 text-xs font-semibold text-slate-200">
            <input
              type="checkbox"
              checked={hideZeroTotal}
              onChange={(e) => {
                setHideZeroTotal(e.target.checked);
                setPage(1);
              }}
              className="rounded border-slate-600 bg-slate-900 accent-sky-500"
            />
            <span>
              {lang === 'es'
                ? 'Ocultar materiales en cero (0 fact / 0 fís)'
                : 'Hide zero-total materials (0 inv / 0 phys)'}
              {zeroTotalCount > 0 ? ` (${zeroTotalCount})` : ''}
            </span>
          </label>

          {totalAlerts > 0 && (
            <div className="flex items-center gap-2 text-xs font-bold">
              {validation.ERROR.length > 0 && (
                <span className="rounded border border-red-500/40 bg-red-500/15 px-2.5 py-1 text-red-300">
                  {validation.ERROR.length} {lang === 'es' ? 'Errores de Integridad' : 'Integrity Errors'}
                </span>
              )}
              {validation.WARNING.length > 0 && (
                <span className="rounded border border-amber-500/40 bg-amber-500/15 px-2.5 py-1 text-amber-300">
                  {validation.WARNING.length} {lang === 'es' ? 'Advertencias' : 'Warnings'}
                </span>
              )}
            </div>
          )}
        </div>
      </div>

      {/* Filters */}
      <div className="grid grid-cols-1 gap-3 md:grid-cols-4">
        {distinctParts.length > 1 && (
          <select
            value={partFilter}
            onChange={(e) => {
              setPartFilter(e.target.value);
              setPage(1);
            }}
            className="rounded-md border border-slate-600 bg-slate-900 px-3 py-2 text-sm text-white"
          >
            <option value="ALL">
              {lang === 'es' ? `Todos Contenedores` : `All Containers`}
            </option>
            {distinctParts.map((p) => (
              <option key={p} value={p}>
                {lang === 'es' ? `Contenedor Parte ${p}` : `Container Part ${p}`}
              </option>
            ))}
          </select>
        )}

        <select
          value={statusFilter}
          onChange={(e) => {
            setStatusFilter(e.target.value);
            setPage(1);
          }}
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

        {distinctCats.length > 1 && (
          <select
            value={catFilter}
            onChange={(e) => {
              setCatFilter(e.target.value);
              setPage(1);
            }}
            className="rounded-md border border-slate-600 bg-slate-900 px-3 py-2 text-sm text-white"
          >
            <option value="ALL">
              {lang === 'es' ? `Todas Categorías` : `All Categories`}
            </option>
            {distinctCats.map((c) => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>
        )}

        <div className="relative">
          <Search className="pointer-events-none absolute top-2.5 left-3 h-4 w-4 text-slate-400" />
          <input
            type="text"
            value={search}
            onChange={(e) => {
              setSearch(e.target.value);
              setPage(1);
            }}
            placeholder={
              lang === 'es'
                ? 'Filtrar por código, estilo, empaque o revuelto...'
                : 'Filter by code, style, packaging or revuelto...'
            }
            className="w-full rounded-md border border-slate-600 bg-slate-900 py-2 pr-3 pl-9 text-sm text-white placeholder-slate-400"
          />
        </div>
      </div>

      {/* Table */}
      <div className="overflow-x-auto rounded-lg border border-slate-700/70 bg-slate-900/60">
        <table className="min-w-full divide-y divide-slate-800 text-left text-xs">
          <thead className="bg-slate-800/95 text-[11px] font-bold text-slate-300 uppercase">
            <tr>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Fila' : 'Row'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Contenedor' : 'Container'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Material SAP' : 'SAP Code'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Estilo Prov' : 'Vendor Style'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Descripción' : 'Description'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Empaque' : 'Pack'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Facturado' : 'Invoiced'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Recibido' : 'Received'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Faltante' : 'Shortage'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Sobrante' : 'Overage'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Dañado' : 'Damaged'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Costo Unit' : 'Unit Cost'}</th>
              <th className="whitespace-nowrap px-3 py-2.5">{lang === 'es' ? 'Clasificación' : 'Classification'}</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800/70 text-slate-100">
            {paginated.map((i, idx) => (
              <tr key={`${i.material_code}-${i.source_row}-${idx}`} className={getRowBg(i)}>
                <td className="whitespace-nowrap px-3 py-2 font-mono text-slate-400">{i.source_row}</td>
                <td className="whitespace-nowrap px-3 py-2">
                  {i.container_part ? `Parte ${i.container_part}` : lang === 'es' ? 'Único' : 'Single'}
                </td>
                <td className="whitespace-nowrap px-3 py-2 font-mono font-semibold text-sky-300">{i.material_code}</td>
                <td className="whitespace-nowrap px-3 py-2">{i.vendor_material}</td>
                <td className="max-w-xs truncate px-3 py-2" title={i.description}>
                  {i.description}
                </td>
                <td className="whitespace-nowrap px-3 py-2 font-mono text-sky-300">{i.canonical_packaging}</td>
                <td className="whitespace-nowrap px-3 py-2">{Math.round(i.quantity_inv_41_units)} UN</td>
                <td className="whitespace-nowrap px-3 py-2 font-bold">{Math.round(i.received_real_units)} UN</td>
                <td className={`whitespace-nowrap px-3 py-2 ${i.faltante > 0 ? 'font-bold text-red-400' : ''}`}>
                  {Math.round(i.faltante)} UN
                </td>
                <td className={`whitespace-nowrap px-3 py-2 ${i.sobrante > 0 ? 'font-bold text-purple-300' : ''}`}>
                  {Math.round(i.sobrante)} UN
                </td>
                <td className={`whitespace-nowrap px-3 py-2 ${i.danado > 0 ? 'font-bold text-amber-300' : ''}`}>
                  {Math.round(i.danado)} UN
                </td>
                <td className="whitespace-nowrap px-3 py-2">${i.unit_cost.toFixed(2)}</td>
                <td className="whitespace-nowrap px-3 py-2">{renderStatusBadge(i)}</td>
              </tr>
            ))}
            {paginated.length === 0 && (
              <tr>
                <td colSpan={13} className="px-4 py-6 text-center text-slate-400">
                  {lang === 'es' ? 'Sin materiales para el filtro seleccionado.' : 'No materials match this filter.'}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {totalPages > 1 && (
        <div className="flex items-center justify-between text-xs text-slate-400">
          <span>
            {lang === 'es'
              ? `Página ${page} de ${totalPages} (${filtered.length} materiales)`
              : `Page ${page} of ${totalPages} (${filtered.length} items)`}
          </span>
          <div className="flex gap-2">
            <button
              type="button"
              disabled={page <= 1}
              onClick={() => setPage((p) => Math.max(1, p - 1))}
              className="rounded border border-slate-700 bg-slate-800 px-3 py-1 disabled:opacity-40"
            >
              {lang === 'es' ? 'Anterior' : 'Prev'}
            </button>
            <button
              type="button"
              disabled={page >= totalPages}
              onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
              className="rounded border border-slate-700 bg-slate-800 px-3 py-1 disabled:opacity-40"
            >
              {lang === 'es' ? 'Siguiente' : 'Next'}
            </button>
          </div>
        </div>
      )}
    </div>
  );
};
