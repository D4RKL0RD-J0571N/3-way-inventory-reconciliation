import type { Language } from '../types/reconciliation';

export const TRANSLATIONS: Record<string, { es: string; en: string }> = {
  app_title: {
    es: 'ReconcileAI — Copiloto 3-Way de Inventarios y Recepción',
    en: 'ReconcileAI — 3-Way Inventory Reconciliation & Receiving Copilot',
  },
  app_badge: {
    es: '🛡️ ReconcileAI — Copiloto Universal de Conciliación 3-Way',
    en: '🛡️ ReconcileAI — Universal 3-Way Reconciliation Copilot',
  },
  app_subtitle: {
    es: 'Auditoría precisa de Órdenes de Compra (PO), Facturación de Nacionalización y Recepción Física en Muelle.',
    en: 'Precision audit of Purchase Orders (PO), Customs Invoices, and Warehouse Dock Physical Receipts.',
  },
  language_label: {
    es: '🌐 Idioma / Language',
    en: '🌐 Language / Idioma',
  },
  explorer_title: {
    es: '📁 Explorador y Selección de Archivos',
    en: '📁 File Explorer & Selection',
  },
  explorer_caption: {
    es: 'Elige un escenario de demostración listo para auditar o carga directamente tus dos libros de Excel (.xlsm / .xlsx) de Control de Calidad y Plantilla de Cierre.',
    en: 'Pick a ready-to-audit sample container or upload your Quality Control and Official Closure Excel workbooks (.xlsm / .xlsx) directly.',
  },
  tab_demo_explorer: {
    es: '📦 Contenedores de Demostración (Listos)',
    en: '📦 Sample Containers (Ready to Run)',
  },
  tab_upload_files: {
    es: '📤 Subir Archivos Excel (.xlsm / .xlsx)',
    en: '📤 Upload Excel Files (.xlsm / .xlsx)',
  },
  upload_qc_label: {
    es: '1. Archivo de Control de Calidad (Muelle / Curvas)',
    en: '1. Quality Control Workbook (Dock / Curves)',
  },
  upload_cierre_label: {
    es: '2. Plantilla de Cierre Oficial (CIERRE RECEP)',
    en: '2. Official Closure Template (CIERRE RECEP)',
  },
  start_audit_btn: {
    es: '🚀 Iniciar Auditoría de estos Archivos',
    en: '🚀 Start Audit on these Files',
  },
  reload_audit_btn: {
    es: '🔄 Cambiar / Recargar Auditoría',
    en: '🔄 Switch / Reload Audit',
  },
  sidebar_guide_title: {
    es: '🧭 Guía de Flujo Operativo',
    en: '🧭 Operating Workflow Guide',
  },
  sidebar_step1: {
    es: '1. Cruce Heurístico: Vincula los materiales padre (Prepacks) con sus tallas hijas (SKUs) respetando la curva de compra.',
    en: '1. Heuristic Cross-Referencing: Links parent Prepacks with child SKUs according to purchasing curves.',
  },
  sidebar_step2: {
    es: '2. Detección de Faltantes & Averías: Identifica cualquier descuadre entre lo facturado y lo recibido físicamente para generar la Nota de Crédito (NCR).',
    en: '2. Shortage & Damage Detection: Identifies any variance between billed and physical dock receipts for Credit Note (NCR) filing.',
  },
  sidebar_step3: {
    es: '3. Solución a diferencias artificiales: Genera las dos tiras exactas (AS:AT y AW:AY) o el libro oficial cuadrado en 1 clic.',
    en: '3. Zero Artificial Discrepancies: Generates exact paste strips (AS:AT & AW:AY) or the official squared workbook in 1 click.',
  },
  sidebar_privacy_title: {
    es: '🔒 Privacidad & Confidencialidad',
    en: '🔒 Privacy & Confidentiality',
  },
  sidebar_privacy_text: {
    es: 'Esta herramienta funciona 100% en memoria local del navegador y no almacena ni envía datos a servidores externos.',
    en: 'This tool operates 100% in local browser memory and does not store or transmit data to external servers.',
  },
  kpi_dock_bultos: {
    es: 'BULTOS / CARTONES (MUELLE)',
    en: 'DOCK CARTONS / BULTOS',
  },
  kpi_invoiced_units: {
    es: 'UNIDADES FACTURADAS (PED 41)',
    en: 'INVOICED UNITS (ORDER 41)',
  },
  kpi_received_units: {
    es: 'UNIDADES RECIBIDAS (MUELLE)',
    en: 'RECEIVED UNITS (DOCK)',
  },
  kpi_net_diff_squared: {
    es: '100% Cuadrado (0 Dif)',
    en: '100% Squared (0 Diff)',
  },
  tab_bultos: {
    es: '📦 Cuadre & Bultos',
    en: '📦 Cartons & Packaging',
  },
  tab_inspector: {
    es: '🔍 Ficha 360° Material',
    en: '🔍 360° Inspector',
  },
  tab_reconcile: {
    es: '🚦 Tabla Semáforo',
    en: '🚦 Status Semáforo',
  },
  tab_curves: {
    es: '📐 Explosión Curvas',
    en: '📐 Curve Breakdown',
  },
  tab_clipboard: {
    es: '⚡ Validador SAP',
    en: '⚡ SAP Validator',
  },
  tab_excel_export: {
    es: '📋 Copiar a Excel',
    en: '📋 Official Excel Export',
  },
  tab_checklist: {
    es: '🖨️ Checklist',
    en: '🖨️ Checklist',
  },
  tab_wiki: {
    es: '📖 SAP & Wiki',
    en: '📖 SAP & Wiki',
  },
  bultos_title: {
    es: '📦 Auditoría y Cuadre de Bultos Físicos (Muelle vs Packing List)',
    en: '📦 Physical Carton Reconciliation & Audit (Dock vs Packing List)',
  },
  bultos_desc: {
    es: 'Conciliación exacta de cajas físicas recibidas en muelle contra el packing list del proveedor. Previene confusiones entre bultos, prepacks (PRK) y unidades totales.',
    en: 'Exact reconciliation of physical cartons received at the dock against vendor packing list. Prevents discrepancies between cartons, prepacks (PRK), and total units.',
  },
  kpi_bultos_descarga: {
    es: 'BULTOS MUELLE (Z-CARGA)',
    en: 'DOCK CARTONS (DISCHARGE)',
  },
  kpi_bultos_pl: {
    es: 'BULTOS PACKING LIST',
    en: 'PACKING LIST CARTONS',
  },
  kpi_dif_bultos: {
    es: 'DIFERENCIA BULTOS FÍSICOS',
    en: 'PHYSICAL CARTON DIFF',
  },
  kpi_total_prk: {
    es: 'TOTAL PREPACKS (PRK)',
    en: 'TOTAL PREPACKS (PRK)',
  },
  kpi_total_units_bultos: {
    es: 'TOTAL UNIDADES FÍSICAS',
    en: 'TOTAL PHYSICAL UNITS',
  },
  calc_bultos_title: {
    es: '🧮 Calculadora de Bultos, Curvas y Empaque Oficial',
    en: '🧮 Carton, Curve & Official Packaging Calculator',
  },
  calc_bultos_caption: {
    es: "Calcula el número de bultos completos y complementos (picos), valida la curva y genera la cadena oficial de empaque (ej. '39*60 1*78' o '50*12').",
    en: "Calculates complete cartons and complement cartons (odds), validates curve, and generates official packaging string (e.g. '39*60 1*78' or '50*12').",
  },
  bultos_sheet_title: {
    es: "📋 Vista Oficial de Hoja 'Bultos' (Plantilla Corporativa)",
    en: "📋 Official 'Bultos' Sheet View (Corporate Template)",
  },
  bultos_sheet_caption: {
    es: "Estructura idéntica a la pestaña 'Bultos' del libro de cierre oficial, con enlaces 1-a-1 a CIERRE RECEP.",
    en: "Identical structure to the official closure workbook 'Bultos' tab, linked 1-to-1 with CIERRE RECEP.",
  },
  inspector_title: {
    es: '🔍 Ficha 360° del Material (Auditoría Instantánea)',
    en: '🔍 360° Material Inspector (Instant Audit)',
  },
  semaforo_title: {
    es: '🚦 Matriz de Conciliación 3-Way (Semáforo de Cierre)',
    en: '🚦 3-Way Reconciliation Matrix (Status Semáforo)',
  },
  curves_title: {
    es: '📐 Explosión Heurística de Curvas por Prepack',
    en: '📐 Heuristic Prepack Curve Breakdown',
  },
  curves_caption: {
    es: 'Relación entre listas de materiales padre (BOM / Prepacks) y tallas hijas recepcionadas en muelle:',
    en: 'Relationship between parent prepack BOMs and child sizes received at docks:',
  },
  validator_title: {
    es: '⚡ Validador y Sanitizador de Portapapeles SAP (Pre-Macro)',
    en: '⚡ SAP Clipboard Validator & Sanitizer (Pre-Macro)',
  },
  validator_caption: {
    es: 'Pega aquí cualquier tabla copiada directamente desde SAP GUI (ME80FN, MB51, SE16N) para desactivar la notación científica y sanear saltos de línea antes de llevarla a Excel:',
    en: 'Paste any table copied directly from SAP GUI (ME80FN, MB51, SE16N) here to neutralize scientific notation and sanitize line breaks before pasting into Excel:',
  },
  validator_clean_btn: {
    es: '⚡ Auditar y Limpiar Portapapeles',
    en: '⚡ Audit & Clean Clipboard',
  },
  export_title: {
    es: '📋 Copiar a Plantilla Oficial de Cierre (Excel .xlsm / .xlsx)',
    en: '📋 Copy to Official Closure Workbook (Excel .xlsm / .xlsx)',
  },
  export_download_btn: {
    es: '⚡ Descargar Libro Oficial Cuadrado (.xlsm con Macros VBA)',
    en: '⚡ Download Official Squared Workbook (.xlsm with VBA Macros)',
  },
  checklist_title: {
    es: '🖨️ Checklist de Impresión Física (Expediente Oficial)',
    en: '🖨️ Physical Binder Printing Checklist (Official File)',
  },
  field_status: {
    es: 'Importación Activa',
    en: 'Active Import',
  },
};

export function t(key: string, lang: Language, fallback?: string): string {
  const entry = TRANSLATIONS[key];
  if (entry) return entry[lang] || entry.es;
  return fallback ?? key;
}
