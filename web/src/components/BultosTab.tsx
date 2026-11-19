import React, { useMemo, useState } from 'react';
import { Calculator, ChevronDown, ChevronUp, Download, Search } from 'lucide-react';
import type { Language, PrepackItem, ReconciliationSummary } from '../types/reconciliation';
import { exportRowsToExcel } from '../utils/excelParser';
import { t } from '../utils/i18n';

interface BultosTabProps {
  items: PrepackItem[];
  summary: ReconciliationSummary;
  lang: Language;
}

export const BultosTab: React.FC<BultosTabProps> = ({ items, summary, lang }) => {
  const [strictCustoms, setStrictCustoms] = useState(false);
  const [hideZeroTotal, setHideZeroTotal] = useState(true);
  const [calcOpen, setCalcOpen] = useState(false);
  const [calcUnits, setCalcUnits] = useState(420);
  const [calcCurve, setCalcCurve] = useState(12);
  const [calcStdBox, setCalcStdBox] = useState(48);
  const [hasOdds, setHasOdds] = useState(true);
  const [oddMode, setOddMode] = useState<'carton' | 'loose'>('carton');
  const [oddUnitsInput, setOddUnitsInput] = useState(36);

  const [catFilter, setCatFilter] = useState('ALL');
  const [search, setSearch] = useState('');
  const [page, setPage] = useState(1);
  const pageSize = 25;

  const visibleItems = useMemo(
    () => (hideZeroTotal ? items.filter((i) => !i.is_zero_total) : items),
    [items, hideZeroTotal]
  );

  const zeroTotalCount = useMemo(() => items.filter((i) => i.is_zero_total).length, [items]);

  const totalBDesc = summary.total_bultos_descarga > 0 ? summary.total_bultos_descarga : summary.received_packages;
  const totalBPl = summary.total_bultos_pl > 0 ? summary.total_bultos_pl : summary.theor_packages;
  const difB = totalBDesc - totalBPl;
  const totalBEnteros = items.reduce((s, it) => s + it.bultos_enteros, 0);
  const totalBComp = items.reduce((s, it) => s + it.bultos_complementos, 0);

  // Calculator math aligned with Acr.Core QuantityEngine.FormatPackagingExpression
  const calcRes = useMemo(() => {
    const safeStd = Math.max(1, calcStdBox);
    const safeCurve = Math.max(1, calcCurve);
    let boxesFull = 0;
    let boxesComp = 0;
    let oddFinal = hasOdds ? oddUnitsInput : 0;
    let pkgStr = '';

    if (hasOdds && oddFinal > 0 && calcUnits > oddFinal) {
      const fullBoxUnits = calcUnits - oddFinal;
      boxesFull = Math.floor(fullBoxUnits / safeStd);
      const rem = fullBoxUnits % safeStd;
      if (rem > 0) oddFinal += rem;

      if (oddMode === 'loose') {
        boxesComp = 0;
        pkgStr = `${boxesFull}*${safeStd} +${oddFinal}`;
      } else {
        boxesComp = 1;
        pkgStr = `${boxesFull}*${safeStd} 1*${oddFinal}`;
      }
    } else {
      boxesFull = Math.floor(calcUnits / safeStd);
      const rem = calcUnits % safeStd;
      if (rem > 0) {
        if (oddMode === 'loose') {
          boxesComp = 0;
          pkgStr = `${boxesFull}*${safeStd} +${rem}`;
        } else {
          boxesComp = 1;
          pkgStr = `${boxesFull}*${safeStd} 1*${rem}`;
        }
      } else {
        boxesComp = 0;
        pkgStr = String(safeStd);
      }
    }

    return {
      totalBoxes: boxesFull + boxesComp,
      boxesFull,
      boxesComp,
      resPrk: calcUnits / safeCurve,
      curveOk: calcUnits % safeCurve === 0,
      pkgStr,
    };
  }, [calcUnits, calcCurve, calcStdBox, hasOdds, oddMode, oddUnitsInput]);

  const categories = useMemo(() => {
    const map: Record<string, number> = {};
    visibleItems.forEach((it) => {
      const c = it.category || 'TEXTIL';
      map[c] = (map[c] ?? 0) + 1;
    });
    return Object.entries(map).sort(([a], [b]) => a.localeCompare(b));
  }, [visibleItems]);

  const rows = useMemo(() => {
    return visibleItems.map((it) => {
      const isOtherP =
        it.discrepancy_type === 'PendingPart' ||
        it.status.startsWith('PENDIENTE_PARTE') ||
        (Boolean(it.container_part) && it.container_part !== 'A' && it.received_real_units === 0);
      const isZeroRec = it.is_zero_total || it.received_real_units === 0 || isOtherP;
      const empaqueDisp = isZeroRec ? '0' : it.canonical_packaging || it.empaque_str || String(it.curve_total || 0);

      if (strictCustoms) {
        const bDisp = isZeroRec
          ? '0 Enteros'
          : it.bultos_complementos > 0
            ? `${Math.round(it.bultos_enteros)} E + ${Math.round(it.bultos_complementos)} P`
            : it.loose_units > 0
              ? `${Math.round(it.bultos_enteros)} E (+${Math.round(it.loose_units)} sueltas)`
              : `${Math.round(it.bultos_enteros)} Enteros`;
        return {
          _isZeroTotal: it.is_zero_total,
          CATEGORIA: it.category || 'TEXTIL',
          'CANTIDAD DE BULTOS Z-Carga': bDisp,
          'Bultos Enteros': isZeroRec ? 0 : Math.round(it.bultos_enteros),
          'Bultos Pico': isZeroRec ? 0 : Math.round(it.bultos_complementos),
          'CANTIDAD DE UNIDADES Z-Carga': isZeroRec ? 0 : Math.round(it.unidades_descarga || it.received_real_units),
          MATERIAL: it.material_code,
          'RECIBIDO REAL & FISICO': isZeroRec ? 0 : Math.round(it.received_real_units),
          Curva: it.curve_total,
          UM: it.um || 'PRK',
          DESCRIPCION: it.description,
          'EMPAQUE UNIDADES Z-Carga': empaqueDisp,
          'Bultos Packing List': isZeroRec ? 0 : Math.round(it.bultos_packing_list || it.bultos_descarga),
          'Diferencia Bultos': isZeroRec ? 0 : Math.round(it.dif_bultos),
          ESTADO: it.discrepancy_label,
        };
      }

      const bDisp = isZeroRec
        ? 0
        : it.bultos_descarga > 0
          ? Math.round(it.bultos_descarga)
          : Math.round(it.quantity_inv_41_prk);

      return {
        _isZeroTotal: it.is_zero_total,
        CATEGORIA: it.category || 'TEXTIL',
        'CANTIDAD DE BULTOS Z-Carga': bDisp,
        'CANTIDAD DE UNIDADES Z-Carga': isZeroRec ? 0 : Math.round(it.unidades_descarga || it.received_real_units),
        MATERIAL: it.material_code,
        'RECIBIDO REAL & FISICO': isZeroRec ? 0 : Math.round(it.received_real_units),
        Curva: it.curve_total,
        UM: it.um || 'PRK',
        DESCRIPCION: it.description,
        'EMPAQUE UNIDADES Z-Carga': empaqueDisp,
        'Bultos Packing List': isZeroRec ? 0 : Math.round(it.bultos_packing_list || it.bultos_descarga),
        'Diferencia Bultos': isZeroRec ? 0 : Math.round(it.dif_bultos),
        ESTADO: it.discrepancy_label,
      };
    });
  }, [visibleItems, strictCustoms]);

  const filteredRows = useMemo(() => {
    return rows.filter((r) => {
      if (catFilter !== 'ALL' && r.CATEGORIA !== catFilter) return false;
      if (search.trim()) {
        const q = search.toLowerCase().trim();
        return (
          String(r.MATERIAL).toLowerCase().includes(q) ||
          String(r.DESCRIPCION).toLowerCase().includes(q) ||
          String(r['CANTIDAD DE BULTOS Z-Carga']).toLowerCase().includes(q) ||
          String(r['EMPAQUE UNIDADES Z-Carga']).toLowerCase().includes(q) ||
          String(r.ESTADO).toLowerCase().includes(q)
        );
      }
      return true;
    });
  }, [rows, catFilter, search]);

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize));
  const paginatedRows = filteredRows.slice((page - 1) * pageSize, page * pageSize);
  const columns = useMemo(() => {
    const sample = paginatedRows[0] ?? rows[0];
    return sample ? Object.keys(sample).filter((k) => !k.startsWith('_')) : [];
  }, [paginatedRows, rows]);

  const handleExportBultos = () => {
    const cleanRows = filteredRows.map(({ _isZeroTotal, ...rest }) => {
      void _isZeroTotal;
      return rest;
    });
    exportRowsToExcel(cleanRows, 'Bultos', `${summary.import_id}_Hoja_Bultos_Oficial.xlsx`);
  };

  return (
    <div className="space-y-5">
      <div>
        <h3 className="text-lg font-bold text-white">{t('bultos_title', lang)}</h3>
        <p className="text-xs text-slate-400">{t('bultos_desc', lang)}</p>
      </div>

      {/* Mode toggle + explanation */}
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-12">
        <div className="flex flex-wrap items-center gap-3 rounded-lg border border-slate-700/60 bg-slate-800/70 p-3 lg:col-span-5">
          <button
            type="button"
            onClick={() => setStrictCustoms(false)}
            className={`flex-1 rounded-md px-3 py-2 text-xs font-bold transition ${
              !strictCustoms
                ? 'bg-sky-600 text-white shadow'
                : 'bg-slate-900/60 text-slate-300 hover:bg-slate-700'
            }`}
          >
            📦 {lang === 'es' ? 'Conteo de Muelle (Consolidado)' : 'Dock Count (Consolidated)'}
          </button>
          <button
            type="button"
            onClick={() => setStrictCustoms(true)}
            className={`flex-1 rounded-md px-3 py-2 text-xs font-bold transition ${
              strictCustoms
                ? 'bg-sky-600 text-white shadow'
                : 'bg-slate-900/60 text-slate-300 hover:bg-slate-700'
            }`}
          >
            ⚖️ {lang === 'es' ? 'Estricto Aduanero (Enteros + Picos)' : 'Strict Customs (Full + Odds)'}
          </button>
        </div>

        <div
          className={`flex items-center rounded-lg border px-4 py-2.5 text-xs lg:col-span-7 ${
            strictCustoms
              ? 'border-blue-500/40 bg-blue-500/10 text-blue-200'
              : 'border-emerald-500/40 bg-emerald-500/10 text-emerald-200'
          }`}
        >
          {strictCustoms
            ? lang === 'es'
              ? '⚖️ Modo Aduanero Activo: Distingue bultos enteros, bultos complemento (8*48 1*36) y unidades sueltas sin caja extra (8*48 +36) con rigor DUA.'
              : '⚖️ Strict Customs Active: Distinguishes full cartons, physical complement cartons (8*48 1*36), and loose units (8*48 +36).'
            : lang === 'es'
              ? '📦 Modo Conteo de Muelle Activo: Empaque canónico listo para hoja Bultos y Z-Carga sin distorsionar totales físicos.'
              : '📦 Dock Count Active: Canonical packaging ready for Bultos sheet and Z-Carga.'}
        </div>
      </div>

      {/* 5 Carton KPI Cards */}
      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-5">
        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5">
          <p className="text-[11px] font-bold text-slate-400 uppercase">
            {strictCustoms
              ? lang === 'es'
                ? 'BULTOS ADUANEROS'
                : 'CUSTOMS CARTONS'
              : t('kpi_bultos_descarga', lang)}
          </p>
          <p className="mt-1 text-xl font-extrabold text-white">
            {(strictCustoms ? totalBEnteros + totalBComp : totalBDesc).toLocaleString()}{' '}
            {lang === 'es' ? 'Bultos' : 'Cartons'}
          </p>
          <p className="mt-1 text-xs text-sky-400">
            {strictCustoms
              ? `${totalBEnteros.toLocaleString()} ${lang === 'es' ? 'Enteros' : 'Full'} + ${totalBComp.toLocaleString()} ${lang === 'es' ? 'Pico' : 'Odds'}`
              : lang === 'es'
                ? 'Descarga Muelle'
                : 'Dock Discharge'}
          </p>
        </div>

        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5">
          <p className="text-[11px] font-bold text-slate-400 uppercase">{t('kpi_bultos_pl', lang)}</p>
          <p className="mt-1 text-xl font-extrabold text-white">
            {totalBPl.toLocaleString()} {lang === 'es' ? 'Bultos' : 'Cartons'}
          </p>
          <p className="mt-1 text-xs text-slate-400">
            {lang === 'es' ? 'Packing List (Proveedor)' : 'Vendor Packing List'}
          </p>
        </div>

        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5">
          <p className="text-[11px] font-bold text-slate-400 uppercase">{t('kpi_dif_bultos', lang)}</p>
          <p className="mt-1 text-xl font-extrabold text-white">
            {difB > 0 ? '+' : ''}
            {difB.toLocaleString()} {lang === 'es' ? 'Bultos' : 'Cartons'}
          </p>
          <p className={`mt-1 text-xs font-semibold ${difB === 0 ? 'text-emerald-400' : 'text-red-400'}`}>
            {difB === 0
              ? lang === 'es'
                ? 'Cuadrado Exacto (0 Dif)'
                : 'Exact Match (0 Diff)'
              : `${lang === 'es' ? 'Descuadre' : 'Variance'}: ${difB}`}
          </p>
        </div>

        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5">
          <p className="text-[11px] font-bold text-slate-400 uppercase">{t('kpi_total_prk', lang)}</p>
          <p className="mt-1 text-xl font-extrabold text-white">{summary.total_prk_41.toLocaleString()} PRK</p>
          <p className="mt-1 text-xs text-slate-400">
            {lang === 'es' ? 'Curva Promedio' : 'Avg Curve'}:{' '}
            {(summary.total_received_units / Math.max(summary.total_prk_41, 1)).toFixed(1)} UN/PRK
          </p>
        </div>

        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5">
          <p className="text-[11px] font-bold text-slate-400 uppercase">{t('kpi_total_units_bultos', lang)}</p>
          <p className="mt-1 text-xl font-extrabold text-white">
            {summary.total_received_units.toLocaleString()} UN
          </p>
          <p className="mt-1 text-xs text-slate-400">
            {lang === 'es' ? 'Promedio' : 'Average'}:{' '}
            {(summary.total_received_units / Math.max(totalBDesc, 1)).toFixed(1)} UN/
            {lang === 'es' ? 'Bulto' : 'Carton'}
          </p>
        </div>
      </div>

      {/* Collapsible Interactive Packaging Calculator */}
      <div className="rounded-lg border border-slate-700/60 bg-[#1E293B]">
        <button
          type="button"
          onClick={() => setCalcOpen((prev) => !prev)}
          className="flex w-full items-center justify-between px-4 py-3 text-left text-sm font-bold text-sky-300 hover:bg-slate-800/70"
        >
          <span className="flex items-center gap-2">
            <Calculator className="h-4 w-4" />
            {t('calc_bultos_title', lang)}
          </span>
          {calcOpen ? <ChevronUp className="h-4 w-4" /> : <ChevronDown className="h-4 w-4" />}
        </button>

        {calcOpen && (
          <div className="space-y-4 border-t border-slate-700/60 p-4">
            <p className="text-xs text-slate-400">{t('calc_bultos_caption', lang)}</p>
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-4">
              <div>
                <label className="block text-xs font-semibold text-slate-300">
                  {lang === 'es' ? 'Total Unidades a Empacar:' : 'Total Units to Pack:'}
                </label>
                <input
                  type="number"
                  min={1}
                  value={calcUnits}
                  onChange={(e) => setCalcUnits(Math.max(1, Number(e.target.value) || 1))}
                  className="mt-1 w-full rounded border border-slate-600 bg-slate-900 px-3 py-1.5 text-sm text-white"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-300">
                  {lang === 'es' ? 'Curva (Unidades por Prepack):' : 'Curve (Units per Prepack):'}
                </label>
                <input
                  type="number"
                  min={1}
                  value={calcCurve}
                  onChange={(e) => setCalcCurve(Math.max(1, Number(e.target.value) || 1))}
                  className="mt-1 w-full rounded border border-slate-600 bg-slate-900 px-3 py-1.5 text-sm text-white"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-300">
                  {lang === 'es' ? 'Unidades por Caja Estándar:' : 'Units per Standard Carton:'}
                </label>
                <input
                  type="number"
                  min={1}
                  value={calcStdBox}
                  onChange={(e) => setCalcStdBox(Math.max(1, Number(e.target.value) || 1))}
                  className="mt-1 w-full rounded border border-slate-600 bg-slate-900 px-3 py-1.5 text-sm text-white"
                />
              </div>
              <div className="space-y-1.5">
                <label className="flex items-center gap-2 text-xs font-semibold text-slate-300">
                  <input
                    type="checkbox"
                    checked={hasOdds}
                    onChange={(e) => setHasOdds(e.target.checked)}
                    className="rounded border-slate-600 bg-slate-900"
                  />
                  {lang === 'es' ? '¿Tiene pico / sueltas?' : 'Has odd carton / loose?'}
                </label>
                {hasOdds && (
                  <div className="flex gap-2">
                    <input
                      type="number"
                      min={0}
                      value={oddUnitsInput}
                      onChange={(e) => setOddUnitsInput(Math.max(0, Number(e.target.value) || 0))}
                      className="w-24 rounded border border-slate-600 bg-slate-900 px-2.5 py-1.5 text-sm text-white"
                    />
                    <select
                      value={oddMode}
                      onChange={(e) => setOddMode(e.target.value as 'carton' | 'loose')}
                      className="flex-1 rounded border border-slate-600 bg-slate-900 px-2 py-1.5 text-xs text-white"
                    >
                      <option value="carton">
                        {lang === 'es' ? 'Bulto físico (1*36)' : 'Physical carton (1*36)'}
                      </option>
                      <option value="loose">
                        {lang === 'es' ? 'Sueltas (+36)' : 'Loose units (+36)'}
                      </option>
                    </select>
                  </div>
                )}
              </div>
            </div>

            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
              <div className="rounded bg-slate-900/70 p-3">
                <p className="text-[11px] font-bold text-slate-400 uppercase">
                  {lang === 'es' ? 'Bultos Físicos Totales' : 'Total Physical Cartons'}
                </p>
                <p className="mt-1 text-lg font-extrabold text-white">{calcRes.totalBoxes} Bultos</p>
              </div>
              <div className="rounded bg-slate-900/70 p-3">
                <p className="text-[11px] font-bold text-slate-400 uppercase">
                  {lang === 'es' ? 'Enteros / Complementos' : 'Full / Odds'}
                </p>
                <p className="mt-1 text-lg font-extrabold text-white">
                  {calcRes.boxesFull} Enteros + {calcRes.boxesComp} Pico
                </p>
              </div>
              <div className="rounded bg-slate-900/70 p-3">
                <p className="text-[11px] font-bold text-slate-400 uppercase">Prepacks</p>
                <p className="mt-1 text-lg font-extrabold text-white">{calcRes.resPrk.toFixed(0)} PRK</p>
                <p className={`text-xs font-semibold ${calcRes.curveOk ? 'text-emerald-400' : 'text-red-400'}`}>
                  {calcRes.curveOk
                    ? lang === 'es'
                      ? 'Curva Exacta'
                      : 'Exact Curve Multiple'
                    : lang === 'es'
                      ? 'Alerta: No múltiplo de curva'
                      : 'Warning: Not a curve multiple'}
                </p>
              </div>
              <div className="rounded bg-slate-900/70 p-3">
                <p className="text-[11px] font-bold text-slate-400 uppercase">
                  {lang === 'es' ? 'Cadena Canónica (AY / Bultos)' : 'Canonical Packaging (AY)'}
                </p>
                <code className="mt-1 inline-block rounded bg-slate-800 px-2.5 py-1 text-sm font-bold text-sky-300">
                  {calcRes.pkgStr}
                </code>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Official Bultos Sheet Table */}
      <div className="space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div>
            <h4 className="text-base font-bold text-white">{t('bultos_sheet_title', lang)}</h4>
            <p className="text-xs text-slate-400">{t('bultos_sheet_caption', lang)}</p>
          </div>
          <div className="flex flex-wrap items-center gap-3">
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
                  ? `Ocultar materiales en cero (0 fact / 0 fís)`
                  : `Hide zero-total materials (0 inv / 0 phys)`}
                {zeroTotalCount > 0 ? ` (${zeroTotalCount})` : ''}
              </span>
            </label>

            <button
              type="button"
              onClick={handleExportBultos}
              className="flex items-center gap-2 rounded-md border border-slate-600 bg-slate-800 px-3 py-2 text-xs font-bold text-slate-100 transition hover:border-sky-400 hover:text-sky-300"
            >
              <Download className="h-4 w-4" />
              {lang === 'es' ? "Descargar Hoja 'Bultos' (.xlsx)" : "Download 'Bultos' Sheet (.xlsx)"}
            </button>
          </div>
        </div>

        {/* Filters */}
        <div className="grid grid-cols-1 gap-3 md:grid-cols-3">
          <select
            value={catFilter}
            onChange={(e) => {
              setCatFilter(e.target.value);
              setPage(1);
            }}
            className="rounded-md border border-slate-600 bg-slate-900 px-3 py-2 text-sm text-white"
          >
            <option value="ALL">
              {lang === 'es' ? `Todas las Categorías (${rows.length})` : `All Categories (${rows.length})`}
            </option>
            {categories.map(([cat, cnt]) => (
              <option key={cat} value={cat}>
                {cat} ({cnt})
              </option>
            ))}
          </select>

          <div className="relative md:col-span-2">
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
                  ? 'Buscar por material, descripción, empaque o bultos...'
                  : 'Search by material, description, packaging or cartons...'
              }
              className="w-full rounded-md border border-slate-600 bg-slate-900 py-2 pr-3 pl-9 text-sm text-white placeholder-slate-400"
            />
          </div>
        </div>

        {/* Data Table */}
        <div className="overflow-x-auto rounded-lg border border-slate-700/70 bg-slate-900/60">
          <table className="min-w-full divide-y divide-slate-800 text-left text-xs">
            <thead className="bg-slate-800/90 text-[11px] font-bold text-slate-300 uppercase">
              <tr>
                {columns.map((col) => (
                  <th key={col} className="whitespace-nowrap px-3 py-2.5">
                    {col}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800/70 text-slate-200">
              {paginatedRows.map((row, idx) => (
                <tr
                  key={`${row.MATERIAL}-${idx}`}
                  className={row._isZeroTotal ? 'bg-slate-800/40 text-slate-400' : 'hover:bg-slate-800/50'}
                >
                  {columns.map((col) => {
                    const val = String(row[col as keyof typeof row] ?? '');
                    if (col === 'EMPAQUE UNIDADES Z-Carga') {
                      return (
                        <td key={col} className="whitespace-nowrap px-3 py-2 font-mono font-bold text-sky-300">
                          {val}
                        </td>
                      );
                    }
                    if (col === 'ESTADO' && row._isZeroTotal) {
                      return (
                        <td key={col} className="whitespace-nowrap px-3 py-2">
                          <span className="rounded border border-slate-600 bg-slate-700/60 px-2 py-0.5 text-[11px] font-semibold text-slate-300">
                            {val}
                          </span>
                        </td>
                      );
                    }
                    return (
                      <td key={col} className="whitespace-nowrap px-3 py-2">
                        {val}
                      </td>
                    );
                  })}
                </tr>
              ))}
              {paginatedRows.length === 0 && (
                <tr>
                  <td colSpan={columns.length || 1} className="px-4 py-6 text-center text-slate-400">
                    {lang === 'es' ? 'Sin resultados para los filtros actuales.' : 'No matching rows.'}
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
                ? `Página ${page} de ${totalPages} (${filteredRows.length} registros)`
                : `Page ${page} of ${totalPages} (${filteredRows.length} records)`}
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
    </div>
  );
};
