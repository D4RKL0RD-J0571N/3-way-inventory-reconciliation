"""
Internationalization (i18n) module for ReconcileAI.
Provides centralized bilingual support (Spanish and English),
with Spanish ('es') set as the system default.
Manages UI string translations, localized table column headers,
and dynamic client-side translation for Streamlit dataframe toolbars.
"""
from typing import Dict, Any, Optional
import streamlit as st

# Supported Languages Registry
DEFAULT_LANGUAGE = "es"
LANGUAGES = {
    "es": "🇪🇸 Español",
    "en": "🇬🇧 English"
}

# Master Translation Dictionary
TRANSLATIONS: Dict[str, Dict[str, str]] = {
    # Application Title & Global Headers
    "app_title": {
        "es": "ReconcileAI — Copiloto 3-Way de Inventarios y Recepción",
        "en": "ReconcileAI — 3-Way Inventory Reconciliation & Receiving Copilot"
    },
    "app_badge": {
        "es": "🛡️ ReconcileAI — Copiloto Universal de Conciliación 3-Way",
        "en": "🛡️ ReconcileAI — Universal 3-Way Reconciliation Copilot"
    },
    "app_subtitle": {
        "es": "Auditoría precisa de Órdenes de Compra (PO), Facturación de Nacionalización y Recepción Física en Muelle.",
        "en": "Precision audit of Purchase Orders (PO), Customs Invoices, and Warehouse Dock Physical Receipts."
    },
    "language_label": {
        "es": "🌐 Idioma / Language",
        "en": "🌐 Language / Idioma"
    },

    # Sidebar / File Explorer
    "explorer_title": {
        "es": "📁 Explorador y Selección de Archivos",
        "en": "📁 File Explorer & Selection"
    },
    "explorer_caption": {
        "es": "Elige directamente los dos archivos de trabajo. Puedes explorar cualquier carpeta de tu disco o arrastrar los archivos desde el Explorador de Windows.",
        "en": "Select your two working spreadsheets directly. Browse any local folder or drag and drop files from Windows Explorer."
    },
    "tab_disk_explorer": {
        "es": "📁 Explorador de Disco (Carpetas Locales)",
        "en": "📁 Disk Explorer (Local Folders)"
    },
    "tab_upload_files": {
        "es": "📤 Subir Archivos (Examinar / Drag & Drop)",
        "en": "📤 Upload Files (Browse / Drag & Drop)"
    },
    "folder_input_label": {
        "es": "Carpeta del Contenedor / Importación:",
        "en": "Container / Import Folder:"
    },
    "quick_folders_label": {
        "es": "Carpetas rápidas:",
        "en": "Quick Folders:"
    },
    "use_folder_btn": {
        "es": "Usar esta",
        "en": "Use this"
    },
    "files_found_in": {
        "es": "Archivos encontrados en {folder} ({count} libros Excel):",
        "en": "Files found in {folder} ({count} Excel workbooks):"
    },
    "qc_file_select_label": {
        "es": "1. Archivo de Control de Calidad (Muelle / Curvas):",
        "en": "1. Quality Control File (Dock / Curves):"
    },
    "cierre_file_select_label": {
        "es": "2. Plantilla de Cierre (Balance / Pedidos):",
        "en": "2. Closure Template (Balance / Orders):"
    },
    "upload_qc_label": {
        "es": "1. Arrastra aquí el archivo de Control de Calidad (Muelle)",
        "en": "1. Drag & drop Quality Control spreadsheet here (Dock)"
    },
    "upload_cierre_label": {
        "es": "2. Arrastra aquí la Plantilla de Cierre Oficial (Balance)",
        "en": "2. Drag & drop Official Closure template here (Balance)"
    },
    "upload_ready": {
        "es": "✅ Archivos cargados exitosamente: {qc} y {cierre}",
        "en": "✅ Files loaded successfully: {qc} and {cierre}"
    },
    "start_audit_btn": {
        "es": "🚀 Iniciar Auditoría de estos Archivos",
        "en": "🚀 Start Audit on these Files"
    },
    "reload_audit_btn": {
        "es": "🔄 Recargar Auditoría",
        "en": "🔄 Reload Audit"
    },

    # Sidebar Guides
    "sidebar_guide_title": {
        "es": "🧭 Guía de Flujo Operativo",
        "en": "🧭 Operating Workflow Guide"
    },
    "sidebar_step1": {
        "es": "1. **Cruce Heurístico:** Vincula los materiales padre (Prepacks) con sus tallas hijas (SKUs) respetando la curva de compra.",
        "en": "1. **Heuristic Cross-Referencing:** Links parent Prepacks with child SKUs according to purchasing curves."
    },
    "sidebar_step2": {
        "es": "2. **Detección de Faltantes & Averías:** Identifica cualquier descuadre entre lo facturado y lo recibido físicamente para generar la Nota de Crédito (NCR).",
        "en": "2. **Shortage & Damage Detection:** Identifies any variance between billed and physical dock receipts for Credit Note (NCR) filing."
    },
    "sidebar_step3": {
        "es": "3. **Solución a diferencias artificiales:** Genera las dos tiras exactas (AS:AT y AW:AY) o el libro oficial cuadrado en 1 clic.",
        "en": "3. **Zero Artificial Discrepancies:** Generates exact paste strips (AS:AT & AW:AY) or the official squared workbook in 1 click."
    },
    "sidebar_privacy_title": {
        "es": "🔒 Privacidad & Confidencialidad",
        "en": "🔒 Privacy & Confidentiality"
    },
    "sidebar_privacy_text": {
        "es": "Esta herramienta funciona 100% en memoria local y no almacena ni envía datos a servidores externos.",
        "en": "This tool operates 100% in local memory and does not store or transmit data to external servers."
    },

    # Top KPI Metrics Bar
    "kpi_theoretical_bultos": {
        "es": "Bultos Teóricos (PL)",
        "en": "Theoretical Cartons (PL)"
    },
    "kpi_theoretical_bultos_sub": {
        "es": "Total según Packing List",
        "en": "Total per Packing List"
    },
    "kpi_dock_bultos": {
        "es": "Bultos / Cartones (Muelle)",
        "en": "Dock Cartons / Bultos"
    },
    "kpi_dock_bultos_sub": {
        "es": "{diff:+d} dif vs teórico",
        "en": "{diff:+d} diff vs theoretical"
    },
    "kpi_invoiced_units": {
        "es": "Unidades Facturadas (Ped 41)",
        "en": "Invoiced Units (Order 41)"
    },
    "kpi_invoiced_units_sub": {
        "es": "Total nacionalizado DUA",
        "en": "Total customs declared"
    },
    "kpi_received_units": {
        "es": "Unidades Recibidas (Muelle)",
        "en": "Received Units (Dock)"
    },
    "kpi_received_units_sub": {
        "es": "Total físico descargado",
        "en": "Total physical received"
    },
    "kpi_erp_balance": {
        "es": "Balance Neto (101 - 102)",
        "en": "Net ERP Balance (101 - 102)"
    },
    "kpi_erp_balance_sub": {
        "es": "Ingresos registrados ERP",
        "en": "Registered ERP receipts"
    },
    "kpi_net_diff": {
        "es": "Diferencia Neta / NCR",
        "en": "Net Discrepancy / NCR"
    },
    "kpi_net_diff_squared": {
        "es": "100% Cuadrado (0 Dif)",
        "en": "100% Squared (0 Diff)"
    },
    "kpi_net_diff_shortage": {
        "es": "Faltante Físico / Avería",
        "en": "Physical Shortage / Damage"
    },

    # Navigation Tabs
    "tab_inspector": {
        "es": "🔍 Ficha 360° Material",
        "en": "🔍 360° Material Inspector"
    },
    "tab_reconcile": {
        "es": "🚦 Tabla Semáforo (Cierre)",
        "en": "🚦 Status Semáforo (Closure)"
    },
    "tab_curves": {
        "es": "📐 Explosión de Curvas",
        "en": "📐 Curve Breakdown"
    },
    "tab_clipboard": {
        "es": "⚡ Validador SAP (Pre-Macro)",
        "en": "⚡ SAP Clipboard Validator"
    },
    "tab_excel_export": {
        "es": "📋 Copiar a Excel Oficial",
        "en": "📋 Copy to Official Excel"
    },
    "tab_checklist": {
        "es": "🖨️ Checklist Expediente",
        "en": "🖨️ Binder Checklist"
    },
    "tab_wiki": {
        "es": "📖 Wiki y Guía de Uso (ES/EN)",
        "en": "📖 User Wiki & Guide (ES/EN)"
    },

    # Tab 0: Ficha 360 Material Inspector
    "inspector_title": {
        "es": "🔍 Ficha 360° del Material (Auditoría Instantánea)",
        "en": "🔍 360° Material Inspector (Instant Audit)"
    },
    "inspector_search_label": {
        "es": "Escribe el código SAP o estilo de proveedor para inspeccionar:",
        "en": "Enter SAP material code or vendor style to inspect:"
    },
    "inspector_search_placeholder": {
        "es": "Ej. 1455100001, 1456051 o nombre de producto...",
        "en": "e.g. 1455100001, 1456051 or product name..."
    },
    "inspector_filter_label": {
        "es": "Selecciona el material de la lista ({count} encontrados):",
        "en": "Select material from list ({count} found):"
    },
    "card_general_info": {
        "es": "📦 Información General del Estilo",
        "en": "📦 General Style Information"
    },
    "card_qty_summary": {
        "es": "📊 Resumen de Cantidades Auditadas",
        "en": "📊 Audited Quantity Summary"
    },
    "card_costs_valuation": {
        "es": "💰 Costos y Valorización",
        "en": "💰 Costs & Valuation"
    },
    "card_child_curves": {
        "es": "📐 Explosión de Tallas Hijas y Curva de Compra",
        "en": "📐 Child SKU Sizes & Purchasing Curve Breakdown"
    },
    "field_material_sap": {"es": "Material SAP:", "en": "SAP Material:"},
    "field_vendor_style": {"es": "Estilo Proveedor:", "en": "Vendor Style:"},
    "field_description": {"es": "Descripción:", "en": "Description:"},
    "field_unit_measure": {"es": "Unidad Medida:", "en": "Unit of Measure:"},
    "field_status": {"es": "Estado:", "en": "Status:"},
    "field_po_units": {"es": "Pedido Compra (PO):", "en": "Purchase Order (PO):"},
    "field_invoiced_units": {"es": "Factura Nacionalización:", "en": "Nationalization Invoice:"},
    "field_physical_units": {"es": "Recibido Físico (Muelle):", "en": "Physical Dock Received:"},
    "field_cartons_dock": {"es": "Bultos / Cajas (Muelle):", "en": "Dock Cartons / Bultos:"},
    "field_difference_units": {"es": "Diferencia Neta:", "en": "Net Discrepancy:"},
    "field_unit_cost": {"es": "Costo Unitario ($):", "en": "Unit Cost ($):"},
    "field_prepack_cost": {"es": "Costo por Prepack:", "en": "Prepack Cost:"},
    "field_total_invoiced_val": {"es": "Valor Total Facturado:", "en": "Total Invoiced Value:"},
    "field_ncr_amount": {"es": "Monto Reclamo NCR:", "en": "NCR Claim Amount:"},

    # Tab 1: Status Semáforo
    "semaforo_title": {
        "es": "🚦 Matriz de Conciliación 3-Way (Semáforo de Cierre)",
        "en": "🚦 3-Way Reconciliation Matrix (Status Semáforo)"
    },
    "filter_all": {"es": "Todos ({count})", "en": "All ({count})"},
    "filter_squared": {"es": "🟢 Cuadrados ({count})", "en": "🟢 Squared ({count})"},
    "filter_part_b": {"es": "🔵 Parte B ({count})", "en": "🔵 Part B ({count})"},
    "filter_shortage": {"es": "🔴 Faltantes / NCR ({count})", "en": "🔴 Shortages / NCR ({count})"},

    # Tab 2: Curves
    "curves_title": {
        "es": "📐 Explosión Heurística de Curvas por Prepack",
        "en": "📐 Heuristic Prepack Curve Breakdown"
    },
    "curves_caption": {
        "es": "Relación entre listas de materiales padre (BOM / Prepacks) y tallas hijas recepcionadas en muelle:",
        "en": "Relationship between parent prepack BOMs and child sizes received at docks:"
    },

    # Tab 3: SAP Clipboard Validator
    "validator_title": {
        "es": "⚡ Validador y Sanitizador de Portapapeles SAP (Pre-Macro)",
        "en": "⚡ SAP Clipboard Validator & Sanitizer (Pre-Macro)"
    },
    "validator_caption": {
        "es": "Pega aquí cualquier tabla copiada directamente desde SAP GUI (ME80FN, MB51, SE16N) para desactivar la notación científica y sanear saltos de línea antes de llevarla a Excel:",
        "en": "Paste any table copied directly from SAP GUI (ME80FN, MB51, SE16N) here to neutralize scientific notation and sanitize line breaks before pasting into Excel:"
    },
    "validator_textarea_label": {
        "es": "Pega los datos de SAP con Ctrl + V:",
        "en": "Paste raw SAP data with Ctrl + V:"
    },
    "validator_clean_btn": {
        "es": "⚡ Auditar y Limpiar Portapapeles",
        "en": "⚡ Audit & Clean Clipboard"
    },

    # Tab 4: Official Excel Export & Strips
    "export_title": {
        "es": "📋 Copiar a Plantilla Oficial de Cierre (Excel .xlsm)",
        "en": "📋 Copy to Official Closure Workbook (Excel .xlsm)"
    },
    "export_option1_title": {
        "es": "⚡ Opción 1: Descarga Directa del Libro Oficial Cuadrado (Recomendada - 1 Clic)",
        "en": "⚡ Option 1: 1-Click Direct Download of Official Squared Workbook (Recommended)"
    },
    "export_option1_desc": {
        "es": "Genera el archivo `.xlsm` oficial directamente en memoria con las columnas `AS:AT` y `AW:AY` ya escritas, fórmulas auto-reparadas y todas las macros VBA intactas.",
        "en": "Generates the official `.xlsm` workbook in memory with columns `AS:AT` and `AW:AY` pre-filled, formulas self-healed, and all VBA macros intact."
    },
    "export_download_btn": {
        "es": "⚡ Descargar Libro Oficial Cuadrado (.xlsm con Macros VBA)",
        "en": "⚡ Download Official Squared Workbook (.xlsm with VBA Macros)"
    },
    "export_option2_title": {
        "es": "📋 Opción 2: Tiras de Pegado Rápido (Manual Controlado - 10 Segundos)",
        "en": "📋 Option 2: Fast Paste Strips (Manual Controlled - 10 Seconds)"
    },
    "strip1_title": {
        "es": "📦 Tira 1: ZDESCARGA (Pegar en Columna AS13)",
        "en": "📦 Strip 1: ZDESCARGA (Paste at Column AS13)"
    },
    "strip1_help": {
        "es": "Copia este bloque con Ctrl+A y Ctrl+C, luego pégalo en la celda AS13 de la hoja CIERRE RECEP.",
        "en": "Select all and copy (Ctrl+A, Ctrl+C), then paste into cell AS13 of the CIERRE RECEP sheet."
    },
    "strip2_title": {
        "es": "📦 Tira 2: PACKING LIST (Pegar en Columna AW13)",
        "en": "📦 Strip 2: PACKING LIST (Paste at Column AW13)"
    },
    "strip2_help": {
        "es": "Copia este bloque con Ctrl+A y Ctrl+C, luego pégalo en la celda AW13 de la hoja CIERRE RECEP.",
        "en": "Select all and copy (Ctrl+A, Ctrl+C), then paste into cell AW13 of the CIERRE RECEP sheet."
    },

    # Tab 5: Binder Checklist
    "checklist_title": {
        "es": "🖨️ Checklist de Impresión Física (Expediente Oficial)",
        "en": "🖨️ Physical Binder Printing Checklist (Official File)"
    },
    "checklist_progress": {
        "es": "Progreso del expediente: {completed} de {total} documentos verificados.",
        "en": "Binder progress: {completed} of {total} documents verified."
    },

    # Table Column Headers
    "col_material": {"es": "MATERIAL", "en": "MATERIAL"},
    "col_status": {"es": "ESTADO", "en": "STATUS"},
    "col_po_43": {"es": "PEDIDO COMPRA (PO 43)", "en": "PURCHASE ORDER (PO 43)"},
    "col_invoiced_41": {"es": "FACTURADO (PED 41)", "en": "INVOICED (ORDER 41)"},
    "col_received_dock": {"es": "RECIBIDO MUELLE", "en": "DOCK RECEIVED"},
    "col_diff": {"es": "DIFERENCIA", "en": "DISCREPANCY"},
    "col_bultos": {"es": "BULTOS", "en": "CARTONS"},
    "col_desc": {"es": "DESCRIPCION", "en": "DESCRIPTION"},
    "col_unit_cost": {"es": "COSTO UNIT ($)", "en": "UNIT COST ($)"},
    "col_total_val": {"es": "VALOR TOTAL ($)", "en": "TOTAL VALUE ($)"},
    "col_ncr_val": {"es": "MONTO NCR ($)", "en": "NCR AMOUNT ($)"},

    # Tooltip and Toolbar Action Translations (for Streamlit table hover)
    "tooltip_fullscreen": {
        "es": "Pantalla completa",
        "en": "Fullscreen"
    },
    "tooltip_close_fullscreen": {
        "es": "Salir de pantalla completa",
        "en": "Close fullscreen"
    },
    "tooltip_download_csv": {
        "es": "Descargar como CSV",
        "en": "Download as CSV"
    },
    "tooltip_search": {
        "es": "Buscar en tabla",
        "en": "Search table"
    },
    "tooltip_columns": {
        "es": "Columnas",
        "en": "Columns"
    }
}


def get_current_language() -> str:
    """
    Returns the active language code ('es' or 'en').
    Defaults to 'es' (Spanish).
    """
    if "app_language" not in st.session_state:
        st.session_state["app_language"] = DEFAULT_LANGUAGE
    return st.session_state["app_language"]


def set_language(lang_code: str) -> None:
    """Updates the active language in session state."""
    if lang_code in LANGUAGES:
        st.session_state["app_language"] = lang_code


def t(key: str, default: Optional[str] = None, **kwargs: Any) -> str:
    """
    Translates a key into the currently active language.
    Falls back to Spanish, default string, or the key itself.
    Interpolates any kwargs passed (e.g. {count}, {folder}).
    """
    lang = get_current_language()
    entry = TRANSLATIONS.get(key)
    if entry:
        template = entry.get(lang) or entry.get("es") or (default if default is not None else key)
    else:
        template = default if default is not None else key

    if kwargs:
        try:
            return template.format(**kwargs)
        except Exception:
            return template
    return template


def render_language_selector() -> str:
    """
    Renders an elegant language selector in the UI.
    Defaults to Spanish ('es').
    """
    current = get_current_language()
    options = list(LANGUAGES.keys())
    current_index = options.index(current) if current in options else 0

    selected = st.sidebar.selectbox(
        t("language_label"),
        options=options,
        index=current_index,
        format_func=lambda code: LANGUAGES.get(code, code),
        key="language_selector_widget"
    )

    if selected != current:
        st.session_state["app_language"] = selected
        st.rerun()

    return selected


def get_table_toolbar_js(lang: str) -> str:
    """
    Returns a client-side JavaScript snippet that dynamically translates
    the Streamlit/BaseWeb table toolbar tooltips and button labels
    ('Fullscreen', 'Download as CSV', 'Search', etc.) to match the selected language.
    """
    t_full = TRANSLATIONS["tooltip_fullscreen"].get(lang, "Pantalla completa")
    t_close = TRANSLATIONS["tooltip_close_fullscreen"].get(lang, "Salir de pantalla completa")
    t_csv = TRANSLATIONS["tooltip_download_csv"].get(lang, "Descargar como CSV")
    t_search = TRANSLATIONS["tooltip_search"].get(lang, "Buscar")

    return f"""
    <script>
    (function() {{
        const map = {{
            "fullscreen": "{t_full}",
            "close fullscreen": "{t_close}",
            "download as csv": "{t_csv}",
            "search": "{t_search}",
            "search table": "{t_search}",
            "buscar en tabla": "{t_search}"
        }};

        function getRootDoc() {{
            try {{
                if (window.parent && window.parent.document && window.parent.document.body) {{
                    return window.parent.document;
                }}
            }} catch(e) {{}}
            return window.document;
        }}

        function translateToolbars() {{
            try {{
                const doc = getRootDoc();
                if (!doc) return;

                // 1. Update aria-labels and titles on toolbar buttons
                const buttons = doc.querySelectorAll('div[data-testid="stElementToolbarButton"] button, button[data-testid="stElementToolbarButton"], div[data-testid="stElementToolbar"] button');
                buttons.forEach(btn => {{
                    const label = btn.getAttribute('aria-label') || btn.getAttribute('title') || '';
                    const lower = label.trim().toLowerCase();
                    if (map[lower]) {{
                        btn.setAttribute('aria-label', map[lower]);
                        btn.setAttribute('title', map[lower]);
                    }}
                }});

                // 2. Translate any visible tooltips rendered in BaseWeb or popovers
                const tooltips = doc.querySelectorAll('[role="tooltip"], div[data-baseweb="tooltip"], div[data-baseweb="popover"] [role="tooltip"]');
                tooltips.forEach(tip => {{
                    const raw = (tip.textContent || tip.innerText || '').trim();
                    const lower = raw.toLowerCase();
                    if (map[lower]) {{
                        tip.textContent = map[lower];
                    }}
                }});
            }} catch (err) {{}}
        }}

        // Run immediately
        translateToolbars();

        try {{
            const doc = getRootDoc();
            const rootWin = (window.parent && window.parent.document) ? window.parent : window;
            if (doc && !rootWin._reconcileToolbarObserverAttached) {{
                rootWin._reconcileToolbarObserverAttached = true;
                
                // Mutation observer for newly mounted popovers/tooltips
                const observer = new MutationObserver(function() {{
                    translateToolbars();
                }});
                observer.observe(doc.body, {{ childList: true, subtree: true }});

                // Mouseover listener for instant translation on hover
                doc.body.addEventListener('mouseover', function() {{
                    translateToolbars();
                }}, true);
            }}
        }} catch(e) {{}}
    }})();
    </script>
    """
