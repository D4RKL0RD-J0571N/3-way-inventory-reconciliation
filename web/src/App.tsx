import React, { useState } from 'react';
import { RefreshCw, Shield, Menu, X, Info } from 'lucide-react';
import type { AuditData, Language } from './types/reconciliation';
import { t } from './utils/i18n';
import { FileExplorerView } from './components/FileExplorerView';
import { KpiHeaderBar } from './components/KpiHeaderBar';
import { BultosTab } from './components/BultosTab';
import { InspectorTab } from './components/InspectorTab';
import { SemaforoTab } from './components/SemaforoTab';
import { CurvesTab } from './components/CurvesTab';
import { SapValidatorTab } from './components/SapValidatorTab';
import { ExcelExportTab } from './components/ExcelExportTab';
import { ChecklistTab } from './components/ChecklistTab';
import { SapWikiTab } from './components/SapWikiTab';

type TabKey =
  | 'bultos'
  | 'inspector'
  | 'semaforo'
  | 'curves'
  | 'clipboard'
  | 'export'
  | 'checklist'
  | 'wiki';

export const App: React.FC = () => {
  const [lang, setLang] = useState<Language>('es');
  const [auditData, setAuditData] = useState<AuditData | null>(null);
  const [activeTab, setActiveTab] = useState<TabKey>('bultos');
  const [sidebarOpen, setSidebarOpen] = useState(false);

  const tabs: Array<{ key: TabKey; labelKey: string }> = [
    { key: 'bultos', labelKey: 'tab_bultos' },
    { key: 'inspector', labelKey: 'tab_inspector' },
    { key: 'semaforo', labelKey: 'tab_reconcile' },
    { key: 'curves', labelKey: 'tab_curves' },
    { key: 'clipboard', labelKey: 'tab_clipboard' },
    { key: 'export', labelKey: 'tab_excel_export' },
    { key: 'checklist', labelKey: 'tab_checklist' },
    { key: 'wiki', labelKey: 'tab_wiki' },
  ];

  const handleAuditLoaded = (data: AuditData) => {
    setAuditData(data);
    setActiveTab('bultos');
    setSidebarOpen(false);
  };

  return (
    <div className="flex min-h-screen bg-[#0F172A] text-[#F8FAFC]">
      {/* Mobile Sidebar Backdrop */}
      {sidebarOpen && (
        <div
          className="fixed inset-0 z-30 bg-black/60 lg:hidden"
          onClick={() => setSidebarOpen(false)}
        />
      )}

      {/* Left Sidebar (matching Streamlit sidebar layout) */}
      <aside
        className={`fixed inset-y-0 left-0 z-40 flex w-64 shrink-0 flex-col border-r border-slate-800 bg-[#111827] p-4 transition-transform lg:static lg:translate-x-0 ${
          sidebarOpen ? 'translate-x-0' : '-translate-x-full'
        }`}
      >
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-gradient-to-br from-sky-500 to-blue-700 text-white shadow">
              <Shield className="h-5 w-5" />
            </div>
            <div>
              <h2 className="text-base font-extrabold tracking-tight text-white">ReconcileAI</h2>
              <p className="text-[11px] font-medium text-sky-400">ACR 3-Way Engine</p>
            </div>
          </div>
          <button
            type="button"
            onClick={() => setSidebarOpen(false)}
            className="rounded p-1 text-slate-400 hover:bg-slate-800 hover:text-white lg:hidden"
            aria-label="Close sidebar"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        <p className="mt-3 text-xs leading-relaxed text-slate-400">{t('app_subtitle', lang)}</p>

        {/* Language Selector */}
        <div className="mt-4">
          <label className="block text-[11px] font-bold tracking-wider text-slate-400 uppercase">
            {t('language_label', lang)}
          </label>
          <select
            value={lang}
            onChange={(e) => setLang(e.target.value as Language)}
            className="mt-1.5 w-full rounded-md border border-slate-700 bg-slate-900 px-3 py-2 text-xs font-semibold text-white"
          >
            <option value="es">🇪🇸 Español</option>
            <option value="en">🇬🇧 English</option>
          </select>
        </div>

        <hr className="my-4 border-slate-800" />

        {/* Active Audit Summary in Sidebar */}
        {auditData && (
          <div className="space-y-3">
            <div className="rounded-md border border-emerald-500/50 bg-[#0F172A] px-3 py-2">
              <div className="flex items-center justify-between">
                <span className="text-[11px] font-semibold text-emerald-400">
                  ⚡ {lang === 'es' ? 'Velocidad de Carga:' : 'Load Time:'}
                </span>
                <span className="text-xs font-bold text-white">{auditData.elapsedSeconds.toFixed(2)}s</span>
              </div>
              <p className="mt-0.5 text-[10px] text-slate-400">
                {lang === 'es' ? 'Procesamiento local en memoria' : 'In-memory local processing'}
              </p>
            </div>

            <div className="rounded-md border border-emerald-500/30 bg-emerald-500/10 p-3 text-xs">
              <p className="font-bold text-emerald-300">
                📦 {t('field_status', lang)}: {auditData.summary.import_id}
              </p>
              <p className="mt-1 truncate text-[11px] text-slate-300" title={auditData.summary.vendor_name}>
                {lang === 'es' ? 'Proveedor' : 'Vendor'}: {auditData.summary.vendor_name}
              </p>
              <p className="mt-0.5 text-[11px] text-slate-300">
                {lang === 'es' ? 'Estilos Activos' : 'Active Styles'}: {auditData.summary.active_styles_count}
                {auditData.summary.zero_total_count > 0 && (
                  <span className="text-slate-400"> (+{auditData.summary.zero_total_count} en 0)</span>
                )}
              </p>
            </div>

            <button
              type="button"
              onClick={() => setAuditData(null)}
              className="flex w-full items-center justify-center gap-2 rounded-md border border-slate-600 bg-slate-800 px-3 py-2 text-xs font-bold text-slate-100 transition hover:border-sky-400 hover:text-sky-300"
            >
              <RefreshCw className="h-3.5 w-3.5" />
              {t('reload_audit_btn', lang)}
            </button>
          </div>
        )}

        <div className="mt-auto pt-6 text-[11px] text-slate-500">
          <p className="font-semibold text-slate-400">ACR Web Edition</p>
          <p className="mt-0.5">
            {lang === 'es'
              ? '100% agnóstico · Paridad Acr.Core .NET 10'
              : '100% agnostic · Acr.Core .NET 10 parity'}
          </p>
        </div>
      </aside>

      {/* Main Fluid Widescreen Content */}
      <main className="flex-1 overflow-x-hidden">
        {/* Mobile top bar */}
        <div className="flex items-center justify-between border-b border-slate-800 bg-[#111827] px-4 py-2.5 lg:hidden">
          <button
            type="button"
            onClick={() => setSidebarOpen(true)}
            className="flex items-center gap-2 rounded-md border border-slate-700 bg-slate-800 px-2.5 py-1.5 text-xs font-bold text-white"
          >
            <Menu className="h-4 w-4" />
            ReconcileAI
          </button>
          <span className="text-xs font-semibold text-sky-400">
            {auditData ? auditData.summary.import_id : 'ACR 3-Way'}
          </span>
        </div>

        <div className="mx-auto max-w-[1600px] px-4 py-5 sm:px-6 lg:px-8">
          {auditData === null ? (
            <FileExplorerView lang={lang} onAuditLoaded={handleAuditLoaded} />
          ) : (
            <div className="space-y-5">
              <KpiHeaderBar summary={auditData.summary} lang={lang} />

              {auditData.warnings.length > 0 && (
                <div className="space-y-1.5">
                  {auditData.warnings.map((w, idx) => (
                    <div
                      key={idx}
                      className="flex items-start gap-2 rounded-md border border-sky-500/30 bg-sky-500/10 px-3.5 py-2 text-xs text-sky-200"
                    >
                      <Info className="mt-0.5 h-4 w-4 shrink-0 text-sky-400" />
                      <span>{w}</span>
                    </div>
                  ))}
                </div>
              )}

              <hr className="border-slate-800" />

              {/* Navigation Tabs */}
              <div className="flex flex-wrap gap-1.5 border-b border-slate-700/60 pb-2">
                {tabs.map((tab) => (
                  <button
                    key={tab.key}
                    type="button"
                    onClick={() => setActiveTab(tab.key)}
                    className={`rounded-t-md px-3.5 py-2 text-xs font-bold whitespace-nowrap transition ${
                      activeTab === tab.key
                        ? 'border-b-2 border-sky-400 bg-slate-800 text-sky-300 shadow-sm'
                        : 'text-slate-400 hover:bg-slate-800/60 hover:text-slate-100'
                    }`}
                  >
                    {t(tab.labelKey, lang)}
                  </button>
                ))}
              </div>

              {/* Tab Content */}
              <div className="pb-10">
                {activeTab === 'bultos' && (
                  <BultosTab items={auditData.items} summary={auditData.summary} lang={lang} />
                )}
                {activeTab === 'inspector' && (
                  <InspectorTab
                    items={auditData.items}
                    validation={auditData.validation}
                    lang={lang}
                  />
                )}
                {activeTab === 'semaforo' && (
                  <SemaforoTab
                    items={auditData.items}
                    validation={auditData.validation}
                    lang={lang}
                  />
                )}
                {activeTab === 'curves' && <CurvesTab items={auditData.items} lang={lang} />}
                {activeTab === 'clipboard' && <SapValidatorTab lang={lang} />}
                {activeTab === 'export' && <ExcelExportTab audit={auditData} lang={lang} />}
                {activeTab === 'checklist' && (
                  <ChecklistTab summary={auditData.summary} lang={lang} />
                )}
                {activeTab === 'wiki' && <SapWikiTab lang={lang} />}
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  );
};

export default App;
