import React, { useState } from 'react';
import { Upload, FolderOpen, Play, ShieldCheck, AlertTriangle, FileSpreadsheet } from 'lucide-react';
import type { AuditData, Language } from '../types/reconciliation';
import { SAMPLE_CONTAINERS } from '../data/sampleDatasets';
import { runWorkbookAudit } from '../utils/excelParser';
import { t } from '../utils/i18n';

interface FileExplorerViewProps {
  lang: Language;
  onAuditLoaded: (data: AuditData) => void;
}

export const FileExplorerView: React.FC<FileExplorerViewProps> = ({ lang, onAuditLoaded }) => {
  const [activeSubtab, setActiveSubtab] = useState<'demo' | 'upload'>('demo');
  const [selectedSampleId, setSelectedSampleId] = useState<string>(SAMPLE_CONTAINERS[0].id);
  const [qcFile, setQcFile] = useState<File | null>(null);
  const [cierreFile, setCierreFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);

  const handleRunSample = (sampleId: string) => {
    setErrorMsg(null);
    const chosen = SAMPLE_CONTAINERS.find((s) => s.id === sampleId);
    if (chosen) {
      onAuditLoaded(chosen.build());
    }
  };

  const handleRunUpload = async (blindMode: boolean) => {
    if (!cierreFile) return;
    setLoading(true);
    setErrorMsg(null);
    try {
      const cierreBuf = await cierreFile.arrayBuffer();
      const qcBuf = !blindMode && qcFile ? await qcFile.arrayBuffer() : null;
      const result = await runWorkbookAudit(qcBuf, cierreBuf, {
        qcFileName: qcFile?.name,
        cierreFileName: cierreFile.name,
        isBlindSetup: blindMode,
      });
      onAuditLoaded(result);
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="space-y-6">
      {/* Enterprise App Header */}
      <div className="rounded-lg bg-gradient-to-br from-[#0F4C81] to-[#1E3A8A] px-5 py-4 shadow-lg">
        <h1 className="text-xl font-bold text-white">{t('app_badge', lang)}</h1>
        <p className="mt-1 text-xs text-slate-200">{t('app_subtitle', lang)}</p>
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-12">
        {/* Left column: Explorer / Uploader */}
        <div className="space-y-5 lg:col-span-8">
          <div className="rounded-xl border border-slate-700/60 bg-gradient-to-br from-slate-900/95 to-slate-800/95 p-5 shadow-xl">
            <h2 className="text-lg font-bold text-sky-400">{t('explorer_title', lang)}</h2>
            <p className="mt-1 text-sm leading-relaxed text-slate-300">{t('explorer_caption', lang)}</p>
          </div>

          {/* Subtabs */}
          <div className="flex flex-wrap gap-2 border-b border-slate-700/60 pb-2">
            <button
              type="button"
              onClick={() => setActiveSubtab('demo')}
              className={`flex items-center gap-2 rounded-t-md px-4 py-2 text-sm font-semibold transition ${
                activeSubtab === 'demo'
                  ? 'border-b-2 border-sky-400 bg-slate-800 text-sky-300'
                  : 'text-slate-400 hover:bg-slate-800/60 hover:text-slate-200'
              }`}
            >
              <FolderOpen className="h-4 w-4" />
              {t('tab_demo_explorer', lang)}
            </button>
            <button
              type="button"
              onClick={() => setActiveSubtab('upload')}
              className={`flex items-center gap-2 rounded-t-md px-4 py-2 text-sm font-semibold transition ${
                activeSubtab === 'upload'
                  ? 'border-b-2 border-sky-400 bg-slate-800 text-sky-300'
                  : 'text-slate-400 hover:bg-slate-800/60 hover:text-slate-200'
              }`}
            >
              <Upload className="h-4 w-4" />
              {t('tab_upload_files', lang)}
            </button>
          </div>

          {errorMsg && (
            <div className="flex items-start gap-2 rounded-md border border-red-500/40 border-l-4 border-l-red-500 bg-red-500/15 p-3 text-sm text-red-300">
              <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0" />
              <span>{errorMsg}</span>
            </div>
          )}

          {activeSubtab === 'demo' ? (
            <div className="space-y-4">
              <div className="grid grid-cols-1 gap-3">
                {SAMPLE_CONTAINERS.map((sample) => {
                  const isSelected = selectedSampleId === sample.id;
                  const badgeStyle =
                    sample.badgeColor === 'emerald'
                      ? 'bg-emerald-500/15 text-emerald-300 border-emerald-500/30'
                      : sample.badgeColor === 'rose'
                        ? 'bg-red-500/15 text-red-300 border-red-500/30'
                        : 'bg-amber-500/15 text-amber-300 border-amber-500/30';
                  return (
                    <button
                      key={sample.id}
                      type="button"
                      onClick={() => setSelectedSampleId(sample.id)}
                      className={`w-full rounded-lg border p-4 text-left transition ${
                        isSelected
                          ? 'border-sky-500 bg-slate-800/90 ring-1 ring-sky-500/50'
                          : 'border-slate-700/70 bg-slate-800/40 hover:border-slate-500 hover:bg-slate-800/70'
                      }`}
                    >
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <span className="font-bold text-slate-100">
                          {lang === 'es' ? sample.labelEs : sample.labelEn}
                        </span>
                        <span className={`rounded border px-2 py-0.5 text-xs font-bold ${badgeStyle}`}>
                          {lang === 'es' ? sample.badgeEs : sample.badgeEn}
                        </span>
                      </div>
                      <p className="mt-1.5 text-xs leading-relaxed text-slate-400">
                        {lang === 'es' ? sample.descriptionEs : sample.descriptionEn}
                      </p>
                    </button>
                  );
                })}
              </div>

              <button
                type="button"
                onClick={() => handleRunSample(selectedSampleId)}
                className="flex w-full items-center justify-center gap-2 rounded-md bg-gradient-to-br from-[#0284C7] to-[#0369A1] px-4 py-3 text-sm font-bold text-white shadow-md transition hover:from-[#0369A1] hover:to-[#075985]"
              >
                <Play className="h-4 w-4" />
                {t('start_audit_btn', lang)} ({selectedSampleId})
              </button>
            </div>
          ) : (
            <div className="space-y-4">
              <div className="rounded-md border border-sky-500/30 bg-sky-500/10 px-3.5 py-2.5 text-xs text-sky-200">
                {lang === 'es'
                  ? '🔍 Detección automática por firma de hojas (CONTROL DATA vs CIERRE RECEP): compatible con Layout #4 (IMP-207), archivos intercambiados y libros híbridos (IMP-218).'
                  : '🔍 Automatic worksheet-signature detection (CONTROL DATA vs CIERRE RECEP): supports Layout #4 (IMP-207), swapped files, and hybrid workbooks (IMP-218).'}
              </div>

              <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                {/* QC File Dropzone */}
                <label className="flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed border-slate-600 bg-slate-800/50 p-5 text-center transition hover:border-sky-400 hover:bg-slate-800">
                  <FileSpreadsheet className="h-8 w-8 text-sky-400" />
                  <span className="mt-2 text-xs font-bold text-slate-200">{t('upload_qc_label', lang)}</span>
                  <span className="mt-1 text-xs text-slate-400">
                    {qcFile
                      ? `✅ ${qcFile.name}`
                      : lang === 'es'
                        ? 'Haz clic para seleccionar .xlsm / .xlsx (Opcional para Modo Pausa o Híbrido)'
                        : 'Click to browse .xlsm / .xlsx (Optional for Pause Mode or Hybrid)'}
                  </span>
                  <input
                    type="file"
                    accept=".xlsx,.xlsm,.xls"
                    className="hidden"
                    onChange={(e) => setQcFile(e.target.files?.[0] ?? null)}
                  />
                </label>

                {/* Closure File Dropzone */}
                <label className="flex cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed border-slate-600 bg-slate-800/50 p-5 text-center transition hover:border-sky-400 hover:bg-slate-800">
                  <FileSpreadsheet className="h-8 w-8 text-emerald-400" />
                  <span className="mt-2 text-xs font-bold text-slate-200">{t('upload_cierre_label', lang)}</span>
                  <span className="mt-1 text-xs text-slate-400">
                    {cierreFile
                      ? `✅ ${cierreFile.name}`
                      : lang === 'es'
                        ? 'Haz clic para seleccionar plantilla CIERRE RECEP (.xlsm / .xlsx)'
                        : 'Click to browse CIERRE RECEP template (.xlsm / .xlsx)'}
                  </span>
                  <input
                    type="file"
                    accept=".xlsx,.xlsm,.xls"
                    className="hidden"
                    onChange={(e) => setCierreFile(e.target.files?.[0] ?? null)}
                  />
                </label>
              </div>

              {cierreFile && !qcFile && (
                <div className="rounded-md border border-amber-500/40 border-l-4 border-l-amber-500 bg-amber-500/10 p-3 text-xs text-amber-200">
                  {lang === 'es'
                    ? 'ℹ️ Sin archivo de Control de Calidad separado: si el libro contiene CONTROL DATA (híbrido) se auditará automáticamente; de lo contrario puedes montar a ciegas en Modo Pausa.'
                    : 'ℹ️ No separate Quality Control file provided: if the workbook contains CONTROL DATA (hybrid) it will reconcile automatically; otherwise you may run Blind Setup in Pause Mode.'}
                </div>
              )}

              <div className="flex flex-wrap gap-3">
                {cierreFile && (
                  <button
                    type="button"
                    disabled={loading}
                    onClick={() => handleRunUpload(false)}
                    className="flex flex-1 items-center justify-center gap-2 rounded-md bg-gradient-to-br from-[#0284C7] to-[#0369A1] px-4 py-3 text-sm font-bold text-white shadow-md transition hover:from-[#0369A1] hover:to-[#075985] disabled:opacity-50"
                  >
                    <Play className="h-4 w-4" />
                    {loading
                      ? lang === 'es'
                        ? 'Auditando libros de Excel...'
                        : 'Auditing Excel workbooks...'
                      : qcFile
                        ? t('start_audit_btn', lang)
                        : lang === 'es'
                          ? 'Auditar Libro (Auto-Detectar Híbrido o Pausa)'
                          : 'Audit Workbook (Auto-Detect Hybrid or Pause)'}
                  </button>
                )}

                {cierreFile && !qcFile && (
                  <button
                    type="button"
                    disabled={loading}
                    onClick={() => handleRunUpload(true)}
                    className="flex flex-1 items-center justify-center gap-2 rounded-md border border-amber-500/50 bg-amber-600/20 px-4 py-3 text-sm font-bold text-amber-200 transition hover:bg-amber-600/30 disabled:opacity-50"
                  >
                    🟡{' '}
                    {lang === 'es'
                      ? 'Forzar Montaje a Ciegas (Modo Pausa)'
                      : 'Force Blind Setup (Pause Mode)'}
                  </button>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Right column: Operating Guide */}
        <div className="lg:col-span-4">
          <div className="rounded-xl border border-slate-700/60 bg-slate-800/70 p-5 shadow-lg">
            <h3 className="text-base font-bold text-slate-100">{t('sidebar_guide_title', lang)}</h3>
            <ul className="mt-4 space-y-3 text-sm leading-relaxed text-slate-300">
              <li>{t('sidebar_step1', lang)}</li>
              <li>{t('sidebar_step2', lang)}</li>
              <li>{t('sidebar_step3', lang)}</li>
            </ul>
            <hr className="my-4 border-slate-700" />
            <div className="flex items-start gap-2.5 text-xs text-slate-300">
              <ShieldCheck className="mt-0.5 h-5 w-5 shrink-0 text-emerald-400" />
              <div>
                <p className="font-bold text-emerald-300">{t('sidebar_privacy_title', lang)}</p>
                <p className="mt-1 text-slate-400">{t('sidebar_privacy_text', lang)}</p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
