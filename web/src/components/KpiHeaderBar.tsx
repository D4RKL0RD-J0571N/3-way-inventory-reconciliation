import React from 'react';
import { ShieldCheck } from 'lucide-react';
import type { Language, ReconciliationSummary } from '../types/reconciliation';
import { t } from '../utils/i18n';

interface KpiHeaderBarProps {
  summary: ReconciliationSummary;
  lang: Language;
}

export const KpiHeaderBar: React.FC<KpiHeaderBarProps> = ({ summary, lang }) => {
  const isBlind = summary.is_blind_setup || summary.status === 'PAUSA_SIN_CONTROL_CALIDAD';

  const partsEntries = Object.entries(summary.container_parts_breakdown).sort(([a], [b]) => a.localeCompare(b));
  let breakdownStr = `${summary.clean_materials_count} ${lang === 'es' ? 'cuadrados' : 'clean'}`;
  if (partsEntries.length > 1) {
    breakdownStr = partsEntries.map(([part, cnt]) => `${cnt} P.${part}`).join(' | ');
  } else if (partsEntries.length === 1) {
    const [onlyP, cnt] = partsEntries[0];
    breakdownStr =
      lang === 'es'
        ? onlyP
          ? `Parte ${onlyP} (${cnt} activos)`
          : `${cnt} activos`
        : onlyP
          ? `Part ${onlyP} (${cnt} active)`
          : `${cnt} active`;
  }

  const hasCardinalDiscrepancy = summary.total_shortage_units > 0 || summary.total_overage_units > 0;
  const pkgRec = summary.received_packages;
  const pkgTheor = summary.theor_packages;
  const pkgDif = summary.difference_packages;

  return (
    <div className="space-y-4">
      {/* Header Banner */}
      {isBlind ? (
        <div className="rounded-md border-l-4 border-amber-500 bg-gradient-to-r from-[#78350f] to-[#451a03] px-5 py-3.5 shadow-lg">
          <h2 className="text-base font-bold text-amber-300">
            ⚠️{' '}
            {lang === 'es'
              ? 'MODO MONTAJE CIEGO DESDE SAP (EN PAUSA)'
              : 'BLIND SAP SETUP MODE (ON PAUSE)'}
          </h2>
          <p className="mt-1 text-xs leading-relaxed text-amber-100">
            {lang === 'es'
              ? 'Regla Operativa: No se cuenta con archivo de Control de Calidad de muelle. Los pedidos de SAP (PO 43 / Doc. 41) han sido montados pero la importación permanece en PAUSA / BORRADOR. El cierre oficial se encuentra bloqueado y no se emiten reclamos de faltante (NCR) hasta que muelle remita el conteo físico.'
              : 'Operational Rule: No dock Quality Control workbook was provided. SAP orders (PO 43 / Doc. 41) are mounted in PAUSE / DRAFT state. Official closure and NCR claims remain locked until physical dock count is submitted.'}
          </p>
        </div>
      ) : (
        <div className="flex flex-wrap items-center justify-between gap-3 rounded-lg bg-gradient-to-br from-[#0F4C81] to-[#1E3A8A] px-5 py-3.5 shadow-lg">
          <div>
            <h1 className="text-lg font-bold text-white">{t('app_badge', lang)}</h1>
            <p className="mt-0.5 text-xs text-slate-200">
              {t('app_subtitle', lang)} | <span className="font-semibold text-sky-200">{summary.import_id}</span> (
              {summary.vendor_name})
            </p>
          </div>
          {summary.vba_present && (
            <div
              className="flex items-center gap-2 rounded-md border border-emerald-400/40 bg-emerald-950/50 px-3 py-1.5 text-[11px] font-semibold text-emerald-200"
              title={summary.vba_sha256 ? `xl/vbaProject.bin SHA-256: ${summary.vba_sha256}` : undefined}
            >
              <ShieldCheck className="h-4 w-4 text-emerald-400" />
              <span>
                {lang === 'es' ? 'Macro-Safe VBA Activo' : 'Macro-Safe VBA Active'}
                {summary.vba_sha256 ? ` · ${summary.vba_sha256.slice(0, 8)}` : ''}
              </span>
            </div>
          )}
        </div>
      )}

      {/* 6-Card Executive KPI Strip */}
      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6">
        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5 shadow">
          <p className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
            {lang === 'en' ? 'IMPORT / INVOICE' : 'IMPORTACIÓN / FACTURA'}
          </p>
          <p className="mt-1 text-xl font-extrabold text-white">{summary.import_id}</p>
          <p className="mt-1 text-xs font-medium text-sky-400">Doc. 41: {summary.invoice_number}</p>
        </div>

        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5 shadow">
          <p className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
            {lang === 'en' ? 'ACTIVE STYLES' : 'ESTILOS ACTIVOS'}
          </p>
          <p className="mt-1 text-xl font-extrabold text-white">
            {summary.active_styles_count} {lang === 'en' ? 'Active' : 'Activos'}
          </p>
          <p className="mt-1 text-xs font-medium text-slate-300">
            {breakdownStr}
            {summary.zero_total_count > 0 && (
              <span className="ml-1 text-slate-400">
                (+{summary.zero_total_count} {lang === 'es' ? 'en 0' : 'zero'})
              </span>
            )}
          </p>
        </div>

        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5 shadow">
          <p className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
            {t('kpi_invoiced_units', lang)}
          </p>
          <p className="mt-1 text-xl font-extrabold text-white">{summary.total_ped_41_units.toLocaleString()} UN</p>
          <p className="mt-1 text-xs font-medium text-slate-300">
            {lang === 'en' ? 'Declared DUA' : 'Declarado en DUA'}
          </p>
        </div>

        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5 shadow">
          <p className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
            {t('kpi_received_units', lang)}
          </p>
          {isBlind ? (
            <>
              <p className="mt-1 text-xl font-extrabold text-amber-300">0 UN</p>
              <p className="mt-1 text-xs font-medium text-slate-400">
                {lang === 'es' ? 'En Pausa / Pendiente QC' : 'On Pause / Pending QC'}
              </p>
            </>
          ) : (
            <>
              <p className="mt-1 text-xl font-extrabold text-white">
                {summary.total_received_units.toLocaleString()} UN
              </p>
              {!hasCardinalDiscrepancy ? (
                <p className="mt-1 text-xs font-semibold text-emerald-400">
                  {t('kpi_net_diff_squared', lang)}
                </p>
              ) : (
                <div className="mt-1 flex flex-wrap gap-x-2 gap-y-0.5 text-[11px] font-bold">
                  {summary.total_shortage_units > 0 && (
                    <span className="text-red-400">
                      -{summary.total_shortage_units.toLocaleString()} Falt.
                    </span>
                  )}
                  {summary.total_overage_units > 0 && (
                    <span className="text-purple-300">
                      +{summary.total_overage_units.toLocaleString()} Sobr.
                    </span>
                  )}
                  {summary.revuelto_pairs_count > 0 && (
                    <span className="text-amber-300">
                      {summary.revuelto_pairs_count} Rev.
                    </span>
                  )}
                </div>
              )}
            </>
          )}
        </div>

        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5 shadow">
          <p className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
            {t('kpi_dock_bultos', lang)}
          </p>
          {isBlind ? (
            <>
              <p className="mt-1 text-xl font-extrabold text-amber-300">
                0 {lang === 'en' ? 'Cartons' : 'Bultos'}
              </p>
              <p className="mt-1 text-xs font-medium text-slate-400">
                {lang === 'es' ? 'Sin Conteo Físico' : 'No Dock Count'}
              </p>
            </>
          ) : (
            <>
              <p className="mt-1 text-xl font-extrabold text-white">
                {pkgRec.toLocaleString()} {lang === 'en' ? 'Cartons' : 'Bultos'}
              </p>
              <p className={`mt-1 text-xs font-semibold ${pkgDif === 0 ? 'text-emerald-400' : 'text-amber-300'}`}>
                {lang === 'en' ? 'Theor' : 'Teóricos'}: {pkgTheor.toLocaleString()} (Dif: {pkgDif > 0 ? '+' : ''}
                {pkgDif.toLocaleString()})
              </p>
            </>
          )}
        </div>

        <div className="rounded-lg border border-slate-700/60 bg-[#1E293B] p-3.5 shadow">
          <p className="text-[11px] font-bold tracking-wider text-slate-400 uppercase">
            {lang === 'en' ? 'CREDIT NOTE (NCR)' : 'NOTA DE CRÉDITO (NCR)'}
          </p>
          {isBlind ? (
            <>
              <p className="mt-1 text-xl font-extrabold text-slate-300">$0.00 USD</p>
              <p className="mt-1 text-xs font-semibold text-amber-300">
                {lang === 'es' ? 'BLOQUEADO (Sin QC)' : 'LOCKED (No QC)'}
              </p>
            </>
          ) : (
            <>
              <p className="mt-1 text-xl font-extrabold text-white">
                ${summary.total_ncr_amount.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
              </p>
              <p
                className={`mt-1 text-xs font-semibold ${
                  summary.status === 'LIMPIO' ? 'text-emerald-400' : 'text-red-400'
                }`}
              >
                {summary.status === 'LIMPIO'
                  ? lang === 'en'
                    ? 'CLEAN (0 Diff)'
                    : 'LIMPIO (0 Discrepancias)'
                  : lang === 'en'
                    ? 'Active Discrepancies'
                    : 'Con Discrepancias'}
              </p>
            </>
          )}
        </div>
      </div>
    </div>
  );
};
