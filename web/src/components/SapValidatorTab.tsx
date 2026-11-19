import React, { useState } from 'react';
import { Check, Copy, Sparkles } from 'lucide-react';
import type { Language } from '../types/reconciliation';
import { cleanMaterialCode, cleanNumber, cleanString } from '../utils/sanitizer';
import { t } from '../utils/i18n';

interface SapValidatorTabProps {
  lang: Language;
}

const SAMPLE_SAP_PASTE = `Material\tDescripcion\tCantidad\tBultos
000000001455102001\tCAMISETA ALGODON BASICA\t600.00\t10
1.455102002E+09\tJEANS SLIM FIT INDIGO\t1,200\t20
1455102003.0\tBLUSA POPELINA BLANCA\t(48)\t1`;

export const SapValidatorTab: React.FC<SapValidatorTabProps> = ({ lang }) => {
  const [pasteArea, setPasteArea] = useState('');
  const [parsed, setParsed] = useState<{
    headers: string[];
    rows: string[][];
    separatorLabel: string;
  } | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);

  const handleValidate = () => {
    setError(null);
    const trimmed = pasteArea.trim();
    if (!trimmed) {
      setError(
        lang === 'es'
          ? 'El cuadro de texto está vacío. Copia datos en SAP con Ctrl+C y pégalos aquí.'
          : 'Text area is empty. Copy table data from SAP with Ctrl+C and paste it here.'
      );
      setParsed(null);
      return;
    }

    const lines = trimmed.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (lines.length < 2) {
      setError(
        lang === 'es'
          ? 'Se necesitan al menos 2 líneas (encabezado + 1 fila de datos).'
          : 'At least 2 lines are required (header + 1 data row).'
      );
      return;
    }

    const firstLine = lines[0];
    const sep = firstLine.includes('\t')
      ? '\t'
      : firstLine.includes('|')
        ? '|'
        : firstLine.includes(';')
          ? ';'
          : ',';
    const sepLabel = sep === '\t' ? (lang === 'es' ? 'Tabulación' : 'Tab') : sep;

    const rawHeaders = firstLine
      .split(sep)
      .map((c) => cleanString(c))
      .filter((c, idx, arr) => !(c === '' && (idx === 0 || idx === arr.length - 1) && sep === '|'));

    const cleanedRows = lines.slice(1).map((line) => {
      const cells = line
        .split(sep)
        .filter((c, idx, arr) => !(c.trim() === '' && (idx === 0 || idx === arr.length - 1) && sep === '|'));
      return rawHeaders.map((hdr, idx) => {
        const val = cells[idx] ?? '';
        const hLow = hdr.toLowerCase();
        if (['material', 'producto', 'sku'].some((k) => hLow.includes(k))) {
          return cleanMaterialCode(val);
        }
        if (['cant', 'unidades', 'qty', 'recibido', 'bultos'].some((k) => hLow.includes(k))) {
          return String(cleanNumber(val));
        }
        return cleanString(val);
      });
    });

    setParsed({ headers: rawHeaders, rows: cleanedRows, separatorLabel: sepLabel });
  };

  const handleCopyCleanTsv = async () => {
    if (!parsed) return;
    const tsv = [parsed.headers.join('\t'), ...parsed.rows.map((r) => r.join('\t'))].join('\n');
    try {
      await navigator.clipboard.writeText(tsv);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // Fallback ignored if clipboard blocked
    }
  };

  return (
    <div className="space-y-4">
      <div>
        <h3 className="text-lg font-bold text-white">{t('validator_title', lang)}</h3>
        <p className="text-xs text-slate-400">{t('validator_caption', lang)}</p>
      </div>

      <div className="space-y-2">
        <div className="flex items-center justify-between">
          <label className="text-xs font-semibold text-slate-300">
            {lang === 'es' ? 'Pega los datos de SAP con Ctrl + V:' : 'Paste raw SAP data with Ctrl + V:'}
          </label>
          <button
            type="button"
            onClick={() => setPasteArea(SAMPLE_SAP_PASTE)}
            className="text-xs font-semibold text-sky-400 hover:text-sky-300"
          >
            {lang === 'es' ? 'Cargar ejemplo SAP sucio' : 'Load dirty SAP example'}
          </button>
        </div>

        <textarea
          rows={6}
          value={pasteArea}
          onChange={(e) => setPasteArea(e.target.value)}
          placeholder={
            lang === 'es'
              ? 'Pega aquí los datos de SAP (separados por tabulaciones, barras | o comas)...'
              : 'Paste raw SAP tabular data here (tab-separated, pipe |, or comma-separated)...'
          }
          className="w-full rounded-md border border-slate-600 bg-slate-900 p-3 font-mono text-xs text-slate-100 placeholder-slate-500"
        />

        <div className="flex flex-wrap gap-3">
          <button
            type="button"
            onClick={handleValidate}
            className="flex items-center gap-2 rounded-md bg-gradient-to-br from-[#0284C7] to-[#0369A1] px-4 py-2 text-xs font-bold text-white shadow hover:from-[#0369A1] hover:to-[#075985]"
          >
            <Sparkles className="h-4 w-4" />
            {t('validator_clean_btn', lang)}
          </button>

          {parsed && (
            <button
              type="button"
              onClick={handleCopyCleanTsv}
              className="flex items-center gap-2 rounded-md border border-emerald-500/40 bg-emerald-500/15 px-4 py-2 text-xs font-bold text-emerald-300 hover:bg-emerald-500/25"
            >
              {copied ? <Check className="h-4 w-4" /> : <Copy className="h-4 w-4" />}
              {copied
                ? lang === 'es'
                  ? '¡TSV Limpio Copiado!'
                  : 'Clean TSV Copied!'
                : lang === 'es'
                  ? 'Copiar TSV Limpio para Excel'
                  : 'Copy Clean TSV for Excel'}
            </button>
          )}
        </div>
      </div>

      {error && (
        <div className="rounded-md border border-amber-500/40 bg-amber-500/10 p-3 text-xs text-amber-200">{error}</div>
      )}

      {parsed && (
        <div className="space-y-3">
          <div className="rounded-md border border-emerald-500/40 bg-emerald-500/15 p-3 text-xs text-emerald-200">
            ✅{' '}
            {lang === 'es'
              ? `Verificación Exitosa: Se detectaron ${parsed.rows.length} filas limpias con separador '${parsed.separatorLabel}'. Notación científica y ceros a la izquierda neutralizados.`
              : `Validation Successful: Detected ${parsed.rows.length} clean rows using '${parsed.separatorLabel}' delimiter. Scientific notation and leading zeros neutralized.`}
          </div>

          <div className="overflow-x-auto rounded-lg border border-slate-700/70 bg-slate-900/60">
            <table className="min-w-full divide-y divide-slate-800 text-left text-xs">
              <thead className="bg-slate-800 text-[11px] font-bold text-slate-300 uppercase">
                <tr>
                  {parsed.headers.map((h, i) => (
                    <th key={`${h}-${i}`} className="px-3 py-2">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800/70 text-slate-200">
                {parsed.rows.slice(0, 20).map((row, rIdx) => (
                  <tr key={rIdx} className="hover:bg-slate-800/50">
                    {row.map((cell, cIdx) => (
                      <td key={cIdx} className="px-3 py-2 font-mono">
                        {cell}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};
