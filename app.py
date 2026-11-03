"""
Interactive Web Application for 3-Way Inventory Reconciliation.
Universal Supply Chain, Logistics & Goods Receipt Audit Engine.
Full-width Widescreen & Professional Dark Theme.
Agnostic architecture designed for multi-vendor and multi-ERP import audits.
Direct file explorer & manual selection: zero guessing, zero blind searching.
"""
import streamlit as st
import pandas as pd
import openpyxl
import os
import io
from pathlib import Path
import streamlit.components.v1 as components

from src.config import list_available_local_imports
from src.parsers import parse_quality_control_boms, parse_cierre_recepcion
from src.engine import compute_summary, filter_items
from src.validator import validate_all_items
from src.sanitizer import clean_material_code, clean_number, clean_string
from src.models import PrepackItem
from src.i18n import t, get_current_language, render_language_selector, get_table_toolbar_js
from src.ui_helpers import inject_skeleton_css, render_skeleton_dashboard, render_skeleton_table, paginate_dataframe

# High-Resolution Widescreen Layout
st.set_page_config(
    page_title=t("app_title"),
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom High-Contrast, Enterprise Dark-Mode & Fluid Widescreen CSS
st.markdown("""
<style>
    /* Clean Header Toolbar & Hide Deploy Button */
    header[data-testid="stHeader"],
    .stAppHeader {
        background: transparent !important;
        height: 2.75rem !important;
        pointer-events: none;
    }
    header[data-testid="stHeader"] *,
    .stAppHeader * {
        pointer-events: auto;
    }
    div[data-testid="stAppDeployButton"] {
        display: none !important;
    }

    /* Fullscreen & High-Resolution True Fluid Layout with Guaranteed Top Clearance */
    .stMainBlockContainer,
    .block-container,
    div[data-testid="stMainBlockContainer"],
    div[data-testid="stAppViewBlockContainer"] {
        padding-top: 5.5rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 99% !important;
        width: 99% !important;
    }

    @media (min-width: 1200px) {
        .stMainBlockContainer,
        .block-container,
        div[data-testid="stMainBlockContainer"] {
            padding-top: 5.5rem !important;
            max-width: 99% !important;
            width: 99% !important;
            padding-left: 2rem !important;
            padding-right: 2rem !important;
        }
    }

    /* Streamlit tabs styling */
    div[data-baseweb="tab-list"] {
        gap: 8px !important;
        width: 100% !important;
        flex-wrap: wrap !important;
        border-bottom: 1px solid rgba(148, 163, 184, 0.25) !important;
        padding-bottom: 6px !important;
        margin-bottom: 1.25rem !important;
    }
    button[data-baseweb="tab"] {
        padding: 9px 18px !important;
        font-size: 13.5px !important;
        font-weight: 600 !important;
        border-radius: 6px 6px 0 0 !important;
        white-space: nowrap !important;
        transition: all 0.15s ease-in-out !important;
    }

    /* High-contrast Buttons with Explanatory Typography */
    button[kind="primary"] {
        background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        border: none !important;
        border-radius: 6px !important;
        padding: 8px 16px !important;
        box-shadow: 0 4px 6px -1px rgba(2, 132, 199, 0.4) !important;
        transition: all 0.15s ease-in-out !important;
    }
    button[kind="primary"]:hover {
        background: linear-gradient(135deg, #0369A1 0%, #075985 100%) !important;
        box-shadow: 0 6px 12px -2px rgba(2, 132, 199, 0.5) !important;
        transform: translateY(-1px);
    }
    button[kind="secondary"] {
        background-color: #1E293B !important;
        color: #F8FAFC !important;
        font-weight: 600 !important;
        border: 1px solid rgba(148, 163, 184, 0.3) !important;
        border-radius: 6px !important;
    }
    button[kind="secondary"]:hover {
        background-color: #334155 !important;
        border-color: #38BDF8 !important;
        color: #38BDF8 !important;
    }

    /* Metric Cards across top row */
    div[data-testid="stMetric"] {
        background-color: var(--secondary-background-color, #1E293B);
        border: 1px solid rgba(148, 163, 184, 0.2);
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
    }
    div[data-testid="stMetricLabel"] p {
        font-size: 11.5px !important;
        font-weight: 700 !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
        color: #94A3B8 !important;
    }
    div[data-testid="stMetricValue"] {
        font-size: 22px !important;
        font-weight: 800 !important;
    }

    /* Enterprise App Header (Rich Blue Gradient) */
    .app-header {
        background: linear-gradient(135deg, #0F4C81 0%, #1E3A8A 100%);
        color: #FFFFFF !important;
        padding: 14px 20px;
        border-radius: 8px;
        margin-top: 0.5rem !important;
        margin-bottom: 14px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.25);
    }
    .app-header h1 {
        color: #FFFFFF !important;
        font-size: 20px;
        font-weight: 700;
        margin: 0;
        padding: 0;
    }
    .app-header p {
        color: #E2E8F0 !important;
        font-size: 12.5px;
        margin: 4px 0 0 0;
    }

    /* Detail Cards (Ficha 360) */
    .detail-card {
        background-color: var(--secondary-background-color, #1E293B);
        border: 1px solid rgba(148, 163, 184, 0.2);
        border-top: 3px solid #0284C7;
        color: var(--text-color, #F8FAFC);
        border-radius: 8px;
        padding: 14px 16px;
        margin-bottom: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
    }
    .detail-title {
        color: #38BDF8;
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        margin-bottom: 10px;
        border-bottom: 1px solid rgba(148, 163, 184, 0.2);
        padding-bottom: 5px;
    }
    .detail-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 7px;
        font-size: 13px;
    }
    .detail-label {
        color: #94A3B8;
        font-weight: 500;
    }
    .detail-value {
        color: #F8FAFC;
        font-weight: 700;
        text-align: right;
    }

    /* Status Banners */
    .status-banner-clean {
        background: rgba(16, 185, 129, 0.15);
        border: 1px solid rgba(16, 185, 129, 0.4);
        border-left: 5px solid #10B981;
        color: #34D399;
        padding: 11px 16px;
        border-radius: 6px;
        font-size: 13.5px;
        font-weight: 600;
        margin-bottom: 16px;
    }
    .status-banner-info {
        background: rgba(59, 130, 246, 0.15);
        border: 1px solid rgba(59, 130, 246, 0.4);
        border-left: 5px solid #3B82F6;
        color: #60A5FA;
        padding: 11px 16px;
        border-radius: 6px;
        font-size: 13.5px;
        font-weight: 600;
        margin-bottom: 16px;
    }
    .status-banner-warn {
        background: rgba(239, 68, 68, 0.15);
        border: 1px solid rgba(239, 68, 68, 0.4);
        border-left: 5px solid #EF4444;
        color: #F87171;
        padding: 11px 16px;
        border-radius: 6px;
        font-size: 13.5px;
        font-weight: 600;
        margin-bottom: 16px;
    }

    /* Welcome Hero Card */
    .welcome-card {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.95) 0%, rgba(30, 41, 59, 0.95) 100%);
        border: 1px solid rgba(148, 163, 184, 0.25);
        border-radius: 10px;
        padding: 20px 24px;
        margin-bottom: 20px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
    }
</style>
""", unsafe_allow_html=True)

# Inject animated skeleton shimmer styles
inject_skeleton_css()

# Session State for Loaded Import Data
if "audit_data" not in st.session_state:
    st.session_state.audit_data = None

@st.cache_data(show_spinner="Procesando y auditando libros de Excel...")
def run_audit(qc_content, cierre_content):
    if isinstance(qc_content, bytes):
        qc_file = io.BytesIO(qc_content)
        qc_raw = qc_content
    else:
        qc_file = qc_content
        qc_raw = Path(qc_content).read_bytes()

    if isinstance(cierre_content, bytes):
        cierre_file = io.BytesIO(cierre_content)
        cierre_raw = cierre_content
    else:
        cierre_file = cierre_content
        cierre_raw = Path(cierre_content).read_bytes()

    boms, qc_details = parse_quality_control_boms(qc_file)
    items, meta = parse_cierre_recepcion(cierre_file, boms, qc_details, return_meta=True)
    summary = compute_summary(items, metadata=meta)
    validation = validate_all_items(items)
    return {
        "items": items,
        "summary": summary,
        "boms": boms,
        "qc_details": qc_details,
        "validation": validation,
        "cierre_raw_bytes": cierre_raw,
        "qc_raw_bytes": qc_raw
    }

# Sidebar Navigation & Management
st.sidebar.image("https://img.icons8.com/color/96/warehouse.png", width=56)
st.sidebar.markdown("### ReconcileAI")
st.sidebar.caption(t("app_subtitle"))

# Language Selector (Spanish default, English optional)
current_lang = render_language_selector()
st.sidebar.divider()

# Inject Dynamic Client-Side Tooltip Translator for Table Toolbars
components.html(get_table_toolbar_js(current_lang), height=0, width=0)

if st.session_state.audit_data is not None:
    current_summary = st.session_state.audit_data["summary"]
    st.sidebar.success(f"📦 **{t('field_status')}:** {current_summary.import_id}")
    st.sidebar.caption(f"Proveedor / Vendor: {current_summary.vendor_name[:25]}")
    st.sidebar.caption(f"Materiales / Styles: {current_summary.total_materials}")
    st.sidebar.divider()
    if st.sidebar.button(t("reload_audit_btn"), use_container_width=True):
        st.session_state.audit_data = None
        st.rerun()

# ----------------- VISTA 1: EXPLORADOR DIRECTO DE ARCHIVOS -----------------
if st.session_state.audit_data is None:
    st.markdown(f"""
    <div class="app-header">
        <h1>{t("app_badge")}</h1>
        <p>{t("app_subtitle")}</p>
    </div>
    """, unsafe_allow_html=True)

    col_welcome_left, col_welcome_right = st.columns([1.25, 0.75], gap="large")

    with col_welcome_left:
        st.markdown(f"""
        <div class="welcome-card">
            <h3 style="margin-top:0; color:#38BDF8;">{t("explorer_title")}</h3>
            <p style="color:#CBD5E1; font-size:13.5px; line-height:1.6; margin-bottom:0;">
                {t("explorer_caption")}
            </p>
        </div>
        """, unsafe_allow_html=True)

        tab_local_explorer, tab_uploader = st.tabs([
            t("tab_disk_explorer"),
            t("tab_upload_files")
        ])

        with tab_local_explorer:
            local_folders = list_available_local_imports()
            default_folder = f"data/{local_folders[0]}" if local_folders else "data"
            
            col_f1, col_f2 = st.columns([3, 1])
            with col_f1:
                folder_input = st.text_input(
                    t("folder_input_label"),
                    value=default_folder,
                    help="Ruta de carpeta / Folder path"
                )
            with col_f2:
                if local_folders:
                    sel_quick_folder = st.selectbox(t("quick_folders_label"), local_folders)
                    if st.button(t("use_folder_btn")):
                        folder_input = f"data/{sel_quick_folder}"
                        st.rerun()

            p = Path(folder_input)
            if p.exists() and p.is_dir():
                all_excel_files = sorted([
                    f.name for f in p.iterdir()
                    if f.is_file() and f.suffix.lower() in [".xlsm", ".xlsx", ".xls"] and not f.name.startswith("~$")
                ])
                if all_excel_files:
                    st.caption(t("files_found_in", folder=folder_input, count=len(all_excel_files)))
                    
                    qc_candidates_idx = [i for i, f in enumerate(all_excel_files) if any(k in f.lower() for k in ["calidad", "qc", "conteo"])]
                    cierre_candidates_idx = [i for i, f in enumerate(all_excel_files) if any(k in f.lower() for k in ["cierre", "plantilla", "inprogress", "recep"]) and "final" not in f.lower()]
                    
                    qc_default_idx = qc_candidates_idx[0] if qc_candidates_idx else 0
                    cierre_default_idx = cierre_candidates_idx[0] if cierre_candidates_idx else min(1, len(all_excel_files)-1)
                    
                    col_sel1, col_sel2 = st.columns(2)
                    with col_sel1:
                        sel_qc_file = st.selectbox(
                            t("qc_file_select_label"),
                            options=all_excel_files,
                            index=qc_default_idx,
                            key="sel_qc_local"
                        )
                    with col_sel2:
                        sel_cierre_file = st.selectbox(
                            t("cierre_file_select_label"),
                            options=all_excel_files,
                            index=cierre_default_idx,
                            key="sel_cierre_local"
                        )
                        
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button(t("start_audit_btn"), type="primary", use_container_width=True):
                        qc_full_path = p / sel_qc_file
                        cierre_full_path = p / sel_cierre_file
                        placeholder = st.empty()
                        with placeholder.container():
                            with st.status(
                                "⚡ " + ("Auditando libros y cuadrando diferencias..." if current_lang == "es" else "Auditing workbooks and calculating reconciliation..."),
                                expanded=True
                            ) as status_box:
                                st.write("📐 " + ("Extrayendo curvas de muelle y listas de materiales (BOMs)..." if current_lang == "es" else "Extracting dock curves and Bills of Materials (BOMs)..."))
                                render_skeleton_dashboard()
                                try:
                                    res = run_audit(str(qc_full_path), str(cierre_full_path))
                                    status_box.update(
                                        label="✅ " + ("¡Auditoría 3-Way completada con éxito!" if current_lang == "es" else "3-Way audit completed successfully!"),
                                        state="complete",
                                        expanded=False
                                    )
                                    st.session_state.audit_data = res
                                    st.rerun()
                                except Exception as e_run:
                                    status_box.update(label="❌ Error", state="error")
                                    st.error(f"Error: {e_run}")
                else:
                    st.warning(f"No se encontraron libros de Excel (.xlsm, .xlsx, .xls) dentro de '{folder_input}'.")
            else:
                st.error(f"La ruta '{folder_input}' no existe o no es un directorio válido.")

        with tab_uploader:
            uploaded_qc = st.file_uploader(
                t("upload_qc_label"),
                type=["xlsm", "xlsx"]
            )
            uploaded_cierre = st.file_uploader(
                t("upload_cierre_label"),
                type=["xlsm", "xlsx"]
            )

            if uploaded_qc and uploaded_cierre:
                if st.button(t("start_audit_btn"), type="primary", use_container_width=True, key="btn_upload_start"):
                    placeholder_up = st.empty()
                    with placeholder_up.container():
                        with st.status(
                            "⚡ " + ("Auditando libros y cuadrando diferencias..." if current_lang == "es" else "Auditing workbooks and calculating reconciliation..."),
                            expanded=True
                        ) as status_box:
                            st.write("📐 " + ("Extrayendo curvas y listas de materiales (BOMs)..." if current_lang == "es" else "Extracting curves and Bills of Materials (BOMs)..."))
                            render_skeleton_dashboard()
                            try:
                                res = run_audit(uploaded_qc.getvalue(), uploaded_cierre.getvalue())
                                status_box.update(
                                    label="✅ " + ("¡Auditoría 3-Way completada con éxito!" if current_lang == "es" else "3-Way audit completed successfully!"),
                                    state="complete",
                                    expanded=False
                                )
                                st.session_state.audit_data = res
                                st.rerun()
                            except Exception as e_up:
                                status_box.update(label="❌ Error", state="error")
                                st.error(f"Error al procesar archivos: {e_up}")

    with col_welcome_right:
        st.markdown(f"### {t('sidebar_guide_title')}")
        st.markdown(f"""
        {t('sidebar_step1')}
        
        {t('sidebar_step2')}
        
        {t('sidebar_step3')}
        
        ---
        **{t('sidebar_privacy_title')}:**  
        {t('sidebar_privacy_text')}
        """)

    st.stop()

# ----------------- VISTA 2: DASHBOARD DE AUDITORÍA ACTIVA -----------------
audit_res = st.session_state.audit_data
items = audit_res["items"]
summary = audit_res["summary"]
boms = audit_res["boms"]
qc_details = audit_res["qc_details"]
validation = audit_res["validation"]

# Header Banner (Agnostic & Corporate-Independent)
st.markdown(f"""
<div class="app-header">
    <h1>{t("app_badge")}</h1>
    <p>{t("app_subtitle")} | {summary.import_id} ({summary.vendor_name})</p>
</div>
""", unsafe_allow_html=True)

# Top KPIs Row (Expanded across full widescreen - 6 cards)
k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1:
    st.metric(
        label="IMPORT / INVOICE" if current_lang == "en" else "IMPORTACIÓN / FACTURA",
        value=summary.import_id,
        delta=f"Doc. 41: {summary.invoice_number}",
        help="Código único / Unique import reference"
    )
with k2:
    if summary.container_parts_breakdown:
        breakdown_str = " | ".join(f"{cnt} P.{part}" for part, cnt in sorted(summary.container_parts_breakdown.items()))
    else:
        breakdown_str = f"{summary.part_a_materials} P.A | {summary.part_b_materials} P.B"
    st.metric(
        label="AUDITED STYLES" if current_lang == "en" else "MATERIALES AUDITADOS",
        value=f"{summary.total_materials} {'Styles' if current_lang == 'en' else 'Estilos'}",
        delta=breakdown_str,
        help="Total parent codes"
    )
with k3:
    st.metric(
        label=t("kpi_invoiced_units"),
        value=f"{summary.total_ped_41_units:,.0f} UN",
        delta="Declared DUA" if current_lang == "en" else "Declarado en DUA",
        help=t("kpi_invoiced_units_sub")
    )
with k4:
    is_squared = (summary.total_ped_41_units == summary.total_received_units)
    dif_units = summary.total_received_units - summary.total_ped_41_units
    delta_txt = (t("kpi_net_diff_squared") if is_squared else f"Dif: {dif_units:,.0f} UN")
    st.metric(
        label=t("kpi_received_units"),
        value=f"{summary.total_received_units:,.0f} UN",
        delta=delta_txt,
        delta_color="normal" if is_squared else "inverse",
        help=t("kpi_received_units_sub")
    )
with k5:
    pkg_rec = summary.received_packages
    pkg_theor = summary.theor_packages
    pkg_dif = summary.difference_packages
    if pkg_rec > 0 or pkg_theor > 0:
        pkg_val = f"{pkg_rec:,.0f} {'Cartons' if current_lang == 'en' else 'Bultos'}"
        pkg_delta = f"{'Theor' if current_lang == 'en' else 'Teóricos'}: {pkg_theor:,.0f} (Dif: {pkg_dif:,.0f})"
    else:
        pkg_val = f"0 {'Cartons' if current_lang == 'en' else 'Bultos'}"
        pkg_delta = "Verificado en muelle" if current_lang == "es" else "Dock verified"
    st.metric(
        label=t("kpi_dock_bultos"),
        value=pkg_val,
        delta=pkg_delta,
        delta_color="normal" if pkg_dif == 0 else "off",
        help="Bultos muelle vs packing list"
    )
with k6:
    st.metric(
        label="CREDIT NOTE (NCR)" if current_lang == "en" else "NOTA DE CRÉDITO (NCR)",
        value=f"${summary.total_ncr_amount:,.2f} USD",
        delta=("CLEAN (0 Diff)" if current_lang == "en" else "LIMPIO (0 Discrepancias)") if summary.status == "LIMPIO" else ("Active Claim" if current_lang == "en" else "Con Reclamo Activo"),
        delta_color="normal" if summary.status == "LIMPIO" else "inverse",
        help="Reclamo NCR / Claim amount"
    )

st.divider()

# Navigation Tabs (Short, punchy, full width - zero cut-offs)
tab_inspector, tab_reconcile, tab_curves, tab_clipboard, tab_excel_export, tab_checklist, tab_wiki = st.tabs([
    t("tab_inspector"),
    t("tab_reconcile"),
    t("tab_curves"),
    t("tab_clipboard"),
    t("tab_excel_export"),
    t("tab_checklist"),
    t("tab_wiki")
])

# ----------------- TAB 0: FICHA 360 FLUIDA -----------------
with tab_inspector:
    st.markdown(f"### {t('inspector_title')}")
    st.caption("Inspecciona cualquier material padre, sus curvas de compra y su balance físico sin desplazarte entre hojas:" if current_lang == "es" else "Inspect any parent style, purchase curves, and dock physical receipts on a single screen:")
    
    col_search1, col_search2 = st.columns([3, 1])
    with col_search1:
        query_search = st.text_input(
            "Buscar por código SAP, estilo de proveedor o descripción:" if current_lang == "es" else "Search by SAP code, vendor style or description:",
            value="",
            placeholder="Escribe código (ej: 1455103), estilo (ej: 25B03), o prenda..." if current_lang == "es" else "Enter code (e.g. 1455103), style (e.g. 25B03), or product...",
            help="Filtra en tiempo real por cualquier dato del material padre o de sus tallas hijas." if current_lang == "es" else "Real-time search across parent or child SKU codes.",
            key="f360_search"
        )
    distinct_parts_360 = sorted(list(set((it.container_part or "A").upper() for it in items)))
    if current_lang == "es":
        part_options_360 = ["Todos"] + [f"Contenedor {p} (Parte {p})" for p in distinct_parts_360]
    else:
        part_options_360 = ["All"] + [f"Container {p} (Part {p})" for p in distinct_parts_360]

    with col_search2:
        part_filter_360 = st.selectbox(
            "Filtrar Contenedor:" if current_lang == "es" else "Filter Container:",
            part_options_360,
            index=0,
            help="Separa los materiales por parte o viaje de descarga." if current_lang == "es" else "Filter items by shipment part.",
            key="f360_part"
        )
        
    selected_p_360 = None
    for p in distinct_parts_360:
        if f"({p})" in part_filter_360 or f"Parte {p}" in part_filter_360 or f"Part {p}" in part_filter_360 or f"Contenedor {p}" in part_filter_360 or f"Container {p}" in part_filter_360:
            selected_p_360 = p
            break

    q_norm = query_search.lower().strip()
    matching = []
    for it in items:
        it_part = (it.container_part or "A").upper()
        if selected_p_360 and it_part != selected_p_360:
            continue
        if not q_norm or (
            q_norm in it.material_code.lower() or
            q_norm in it.vendor_material.lower() or
            q_norm in it.description.lower() or
            any(q_norm in c.material_code.lower() for c in it.children)
        ):
            matching.append(it)
            
    if matching:
        if len(matching) > 1:
            opt_labels = {
                f"📦 {it.material_code} | {it.vendor_material} | {it.description[:45]} ({'Parte' if current_lang == 'es' else 'Part'} {it.container_part})": it
                for it in matching
            }
            selected_label = st.selectbox(
                f"Selecciona un estilo ({len(matching)} coincidencias disponibles):" if current_lang == "es" else f"Select a style ({len(matching)} matches available):",
                options=list(opt_labels.keys()),
                key="f360_selected_mat",
                help="Despliega la lista para buscar o elegir directamente el material a auditar." if current_lang == "es" else "Dropdown to pick audited material."
            )
            selected_item = opt_labels[selected_label]
        else:
            selected_item = matching[0]
            st.info(f"{'Mostrando resultado único:' if current_lang == 'es' else 'Showing single match:'} **{selected_item.material_code}** — {selected_item.description}")
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Status banner
        if selected_item.status == "LIMPIO":
            st.markdown(f'<div class="status-banner-clean">✅ {"ESTADO: LIMPIO" if current_lang == "es" else "STATUS: CLEAN"} — {selected_item.observations}</div>', unsafe_allow_html=True)
        elif selected_item.status.startswith("PENDIENTE_PARTE_") or "CONTENEDOR" in selected_item.status:
            part_lbl = selected_item.container_part
            st.markdown(f'<div class="status-banner-info">🚢 {"ESTADO: CONTENEDOR" if current_lang == "es" else "STATUS: CONTAINER"} {part_lbl} — {selected_item.observations}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="status-banner-warn">⚠️ {"ESTADO:" if current_lang == "es" else "STATUS:"} {selected_item.status} — {selected_item.observations} ({"Reclamo NCR:" if current_lang == "es" else "NCR Claim:"} ${selected_item.ncr_amount:,.2f} USD)</div>', unsafe_allow_html=True)
            
        # 4 High-Contrast Adaptive Cards
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
            <div class="detail-card">
                <div class="detail-title">{"1. Identificación" if current_lang == "es" else "1. Identification"}</div>
                <div class="detail-row"><span class="detail-label">{"Código SAP:" if current_lang == "es" else "SAP Code:"}</span><span class="detail-value">{selected_item.material_code}</span></div>
                <div class="detail-row"><span class="detail-label">{"Estilo Prov:" if current_lang == "es" else "Vendor Style:"}</span><span class="detail-value">{selected_item.vendor_material}</span></div>
                <div class="detail-row"><span class="detail-label">{"Categoría:" if current_lang == "es" else "Category:"}</span><span class="detail-value">{selected_item.category}</span></div>
                <div class="detail-row"><span class="detail-label">{"Descripción:" if current_lang == "es" else "Description:"}</span><span class="detail-value" style="font-size:12px;">{selected_item.description[:32]}</span></div>
                <div class="detail-row"><span class="detail-label">{"Contenedor:" if current_lang == "es" else "Container:"}</span><span class="detail-value">{"Parte" if current_lang == "es" else "Part"} {selected_item.container_part}</span></div>
            </div>
            """, unsafe_allow_html=True)
            
        with c2:
            st.markdown(f"""
            <div class="detail-card">
                <div class="detail-title">{"2. Compras (PO 43)" if current_lang == "es" else "2. Purchasing (PO 43)"}</div>
                <div class="detail-row"><span class="detail-label">{"Orden Compra:" if current_lang == "es" else "Purchase Order:"}</span><span class="detail-value">{selected_item.order_po_43}</span></div>
                <div class="detail-row"><span class="detail-label">Prepacks (PRK):</span><span class="detail-value">{selected_item.quantity_po_43_prk:,.0f} PRK</span></div>
                <div class="detail-row"><span class="detail-label">{"Total Unidades:" if current_lang == "es" else "Total Units:"}</span><span class="detail-value">{selected_item.quantity_po_43_units:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">{"Curva por Bulto:" if current_lang == "es" else "Units / Carton:"}</span><span class="detail-value">{selected_item.curve_total} UN/PRK</span></div>
                <div class="detail-row"><span class="detail-label">{"Unidad Medida:" if current_lang == "es" else "UoM:"}</span><span class="detail-value">PRK / UN</span></div>
            </div>
            """, unsafe_allow_html=True)
            
        with c3:
            st.markdown(f"""
            <div class="detail-card">
                <div class="detail-title">{"3. Factura 41 & Físico" if current_lang == "es" else "3. Invoice 41 & Dock"}</div>
                <div class="detail-row"><span class="detail-label">{"Facturado (41):" if current_lang == "es" else "Invoiced (41):"}</span><span class="detail-value">{selected_item.quantity_inv_41_units:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">{"Recibido Muelle:" if current_lang == "es" else "Dock Received:"}</span><span class="detail-value">{selected_item.received_real_units:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">{"Costo Unitario:" if current_lang == "es" else "Unit Cost:"}</span><span class="detail-value">${selected_item.unit_cost:.2f} USD</span></div>
                <div class="detail-row"><span class="detail-label">{"Total Factura:" if current_lang == "es" else "Total Invoice:"}</span><span class="detail-value">${(selected_item.quantity_inv_41_units * selected_item.unit_cost):,.2f} USD</span></div>
                <div class="detail-row"><span class="detail-label">{"Diferencia:" if current_lang == "es" else "Discrepancy:"}</span><span class="detail-value">{int(selected_item.received_real_units - selected_item.quantity_inv_41_units)} UN</span></div>
            </div>
            """, unsafe_allow_html=True)
            
        with c4:
            st.markdown(f"""
            <div class="detail-card">
                <div class="detail-title">{"4. Balances y Ajustes" if current_lang == "es" else "4. Balances & Adjustments"}</div>
                <div class="detail-row"><span class="detail-label">SAP Mov 101:</span><span class="detail-value">{selected_item.sap_mov_101:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">{"Faltante:" if current_lang == "es" else "Shortage:"}</span><span class="detail-value">{selected_item.faltante:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">{"Dañado / Avería:" if current_lang == "es" else "Damaged:"}</span><span class="detail-value">{selected_item.danado:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">{"Sobrante:" if current_lang == "es" else "Surplus:"}</span><span class="detail-value">{selected_item.sobrante:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">{"Monto NCR:" if current_lang == "es" else "NCR Amount:"}</span><span class="detail-value">${selected_item.ncr_amount:,.2f} USD</span></div>
            </div>
            """, unsafe_allow_html=True)

        if selected_item.children:
            st.markdown("#### 📐 " + ("Desglose de Tallas Hijas ancladas a este Prepack:" if current_lang == "es" else "Child SKU Breakdown linked to this Prepack:"))
            c_rows = []
            for c in selected_item.children:
                c_rows.append({
                    "Material Hijo (Talla)" if current_lang == "es" else "Child SKU (Size)": c.material_code,
                    "Descripción / Talla" if current_lang == "es" else "Description / Size": c.description,
                    "Ratio Curva" if current_lang == "es" else "Curve Ratio": c.ratio,
                    "Esperado (UN)" if current_lang == "es" else "Expected (UN)": c.theoretical_units,
                    "Recibido Físico (UN)" if current_lang == "es" else "Physical Received (UN)": c.received_units,
                    "Diferencia" if current_lang == "es" else "Difference": c.difference,
                    "Costo Unitario ($)" if current_lang == "es" else "Unit Cost ($)": f"${c.unit_cost:.2f}"
                })
            st.dataframe(pd.DataFrame(c_rows), use_container_width=True)
    else:
        st.warning(f"No se encontraron materiales que coincidan con la búsqueda: '{query_search}'." if current_lang == "es" else f"No materials matched your search: '{query_search}'.")

# ----------------- TAB 1: TABLA SEMÁFORO -----------------
with tab_reconcile:
    st.subheader(t("semaforo_title"))
    st.caption("🟢 Verde = Conciliación Limpia | 🔵 Azul = Pendiente Contenedor B | 🔴 Rojo = Faltante / Dañado Reclamable" if current_lang == "es" else "🟢 Green = Clean Reconciled | 🔵 Blue = Pending Container B | 🔴 Red = Shortage / Claimable Damage")
    
    distinct_parts_t = sorted(list(set((it.container_part or "A").upper() for it in items)))
    if current_lang == "es":
        part_options_t = ["Todos"] + [f"Contenedor {p} (Parte {p})" for p in distinct_parts_t]
    else:
        part_options_t = ["All"] + [f"Container {p} (Part {p})" for p in distinct_parts_t]

    col_t1, col_t2, col_t3 = st.columns([1.5, 1.5, 3])
    with col_t1:
        part_table = st.selectbox("Filtrar Contenedor:" if current_lang == "es" else "Filter Container:", part_options_t, index=0, key="t_part")
        p_val = "ALL"
        for p in distinct_parts_t:
            if f"({p})" in part_table or f"Parte {p}" in part_table or f"Part {p}" in part_table:
                p_val = p
                break

    with col_t2:
        stat_options = ["Todos" if current_lang == "es" else "All", "LIMPIO"]
        for p in distinct_parts_t:
            if p != "A":
                stat_options.append(f"PENDIENTE_PARTE_{p}")
        stat_options.extend(["FALTANTE", "DAÑADO", "SOBRANTE"])
        status_table = st.selectbox("Filtrar Estado:" if current_lang == "es" else "Filter Status:", stat_options, index=0, key="t_stat")
        s_val = "ALL" if ("Todos" in status_table or "All" in status_table) else status_table

    with col_t3:
        search_table = st.text_input("Buscar en la grilla:" if current_lang == "es" else "Search table:", placeholder="Filtrar por código, estilo o descripción..." if current_lang == "es" else "Filter by code, style or description...", key="t_search")
        
    items_table = filter_items(items, part_filter=p_val, status_filter=s_val, search_query=search_table)
    
    col_cont = "Contenedor" if current_lang == "es" else "Container"
    col_mat = "Material SAP" if current_lang == "es" else "SAP Code"
    col_style = "Estilo Prov" if current_lang == "es" else "Vendor Style"
    col_desc = "Descripción" if current_lang == "es" else "Description"
    col_po = "PO (43)"
    col_prk = "Ped. 41 (PRK)" if current_lang == "es" else "Order 41 (PRK)"
    col_curve = "Curva" if current_lang == "es" else "Curve"
    col_inv = "Facturado UN" if current_lang == "es" else "Invoiced UN"
    col_rec = "Recibido UN" if current_lang == "es" else "Received UN"
    col_sap = "SAP Mov 101"
    col_short = "Faltante" if current_lang == "es" else "Shortage"
    col_dam = "Dañado" if current_lang == "es" else "Damaged"
    col_cost = "Costo Unit" if current_lang == "es" else "Unit Cost"
    col_stat = "Estado" if current_lang == "es" else "Status"

    t_rows = []
    for i in items_table:
        is_clean = (i.status == "LIMPIO")
        is_pending = (i.status.startswith("PENDIENTE_PARTE_") or "CONTENEDOR" in i.status)
        if is_clean:
            status_badge = "🟢 LIMPIO" if current_lang == "es" else "🟢 CLEAN"
        elif is_pending:
            status_badge = f"🔵 CONT. {i.container_part}" if current_lang == "es" else f"🔵 PART {i.container_part}"
        else:
            status_badge = f"🔴 {i.status}"

        t_rows.append({
            col_cont: f"Parte {i.container_part}" if current_lang == "es" else f"Part {i.container_part}",
            col_mat: i.material_code,
            col_style: i.vendor_material,
            col_desc: i.description,
            col_po: i.order_po_43,
            col_prk: int(i.quantity_inv_41_prk),
            col_curve: i.curve_total,
            col_inv: int(i.quantity_inv_41_units),
            col_rec: int(i.received_real_units),
            col_sap: int(i.sap_mov_101),
            col_short: int(i.faltante),
            col_dam: int(i.danado),
            col_cost: i.unit_cost,
            col_stat: status_badge
        })
    df_t = pd.DataFrame(t_rows)
    
    def highlight_status(row):
        st_val = row[col_stat]
        if "LIMPIO" in st_val or "CLEAN" in st_val:
            return ["background-color: rgba(16, 185, 129, 0.08)"] * len(row)
        elif "CONT." in st_val or "PART" in st_val:
            return ["background-color: rgba(59, 130, 246, 0.08)"] * len(row)
        else:
            return ["background-color: rgba(239, 68, 68, 0.15)"] * len(row)

    df_display = paginate_dataframe(df_t, key_prefix="semaforo", page_size_options=[50, 100, 250, 500])
    styled_df = df_display.style.apply(highlight_status, axis=1)
    
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=540,
        column_config={
            col_cost: st.column_config.NumberColumn(format="$%.2f"),
            col_prk: st.column_config.NumberColumn(format="%d PRK"),
            col_inv: st.column_config.NumberColumn(format="%d UN"),
            col_rec: st.column_config.NumberColumn(format="%d UN"),
            col_sap: st.column_config.NumberColumn(format="%d UN"),
            col_short: st.column_config.NumberColumn(format="%d UN"),
            col_dam: st.column_config.NumberColumn(format="%d UN"),
        }
    )

# ----------------- TAB 2: EXPLOSIÓN DE CURVAS -----------------
with tab_curves:
    st.subheader(t("curves_title"))
    st.caption(t("curves_caption"))
    
    parent_choices = [f"📦 {i.material_code} | {i.vendor_material} | {i.description[:45]} (Curva {i.curve_total} UN)" for i in items if i.children]
    if parent_choices:
        sel_p = st.selectbox("Selecciona Prepack:" if current_lang == "es" else "Select Prepack:", parent_choices, key="curv_p")
        sel_code_curv = sel_p.split(" | ")[0].replace("📦 ", "").strip()
        curv_item = next((i for i in items if i.material_code == sel_code_curv), None)
        
        if curv_item:
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("PREPACKS FACTURADOS" if current_lang == "es" else "INVOICED PREPACKS", f"{curv_item.quantity_inv_41_prk:.0f} PRK")
            k2.metric("CURVA POR BULTO" if current_lang == "es" else "CURVE PER CARTON", f"{curv_item.curve_total} UN/PRK")
            k3.metric("TOTAL PIEZAS ESPERADAS" if current_lang == "es" else "TOTAL EXPECTED UNITS", f"{curv_item.quantity_inv_41_units:,.0f} UN")
            k4.metric("ESTADO LÍNEA" if current_lang == "es" else "LINE STATUS", curv_item.status)
            
            st.markdown("#### " + ("Reparto Matemático por Talla:" if current_lang == "es" else "Mathematical Breakdown by Size:"))
            c_rows_v = []
            for c in curv_item.children:
                c_rows_v.append({
                    "Código SKU (Hijo)" if current_lang == "es" else "Child SKU Code": c.material_code,
                    "Descripción / Talla" if current_lang == "es" else "Description / Size": c.description,
                    "Ratio Curva" if current_lang == "es" else "Curve Ratio": c.ratio,
                    "Unidades Teóricas" if current_lang == "es" else "Theoretical Units": c.theoretical_units,
                    "Unidades Físicas" if current_lang == "es" else "Physical Received": c.received_units,
                    "Diferencia" if current_lang == "es" else "Difference": c.difference,
                    "Costo Unitario ($)" if current_lang == "es" else "Unit Cost ($)": f"${c.unit_cost:.2f}"
                })
            st.dataframe(pd.DataFrame(c_rows_v), use_container_width=True)

# ----------------- TAB 3: VERIFICADOR PRE-MACRO -----------------
with tab_clipboard:
    st.subheader(t("validator_title"))
    st.markdown(t("validator_caption"))
    
    paste_area = st.text_area(
        t("validator_textarea_label"),
        height=160,
        placeholder="Pega aquí los datos de SAP (separados por tabulaciones, barras | o comas)..." if current_lang == "es" else "Paste raw SAP tabular data here (tab-separated, pipe |, or comma-separated)...",
        help="Acepta texto copiado directamente desde el GUI de SAP o exportaciones de tabla." if current_lang == "es" else "Accepts raw copied text from SAP GUI grids or clipboard.",
        key="paste_area_input"
    )
    
    col_btn1, col_btn2 = st.columns([2, 3])
    with col_btn1:
        validate_clicked = st.button(t("validator_clean_btn"), type="primary")
        
    if validate_clicked:
        if paste_area.strip():
            try:
                first_line = paste_area.strip().split("\n")[0]
                sep = "\t" if "\t" in first_line else ("|" if "|" in first_line else (";" if ";" in first_line else ","))
                df_raw = pd.read_csv(io.StringIO(paste_area), sep=sep)
                df_raw.columns = [clean_string(c) for c in df_raw.columns]
                
                for c in df_raw.columns:
                    if any(k in c.lower() for k in ["material", "producto", "sku"]):
                        df_raw[c] = df_raw[c].apply(clean_material_code)
                    elif any(k in c.lower() for k in ["cant", "unidades", "qty", "recibido", "bultos"]):
                        df_raw[c] = df_raw[c].apply(clean_number)
                    else:
                        df_raw[c] = df_raw[c].apply(clean_string)
                        
                st.success(f"✅ **Verificación Exitosa:** Se detectaron **{len(df_raw)} filas limpias** con separador '{'Tabulación' if sep == '\t' else sep}'. La tabla está lista para pegarse en la macro oficial sin riesgo de desfasar columnas.")
                st.dataframe(df_raw.head(15), use_container_width=True)
            except Exception as e_p:
                st.error(f"Error al analizar el contenido del portapapeles: {e_p}")
        else:
            st.warning("El cuadro de texto está vacío. Copia datos en SAP con Ctrl+C y pégalos aquí.")

# ----------------- TAB 4: GENERADOR PARA EXCEL OFICIAL -----------------
with tab_excel_export:
    st.subheader(t("export_title"))
    st.markdown("""
    Este módulo resuelve el problema del llenado manual en la hoja **`CIERRE RECEP`**:
    - **Bloque 1:** Genera las columnas principales (`A` hasta `BB`).
    - **Bloque 2 (ZDESCARGA y PACKING LIST):** Entrega las tiras de datos listas para copiar y pegar en las columnas **`AS:AT`** y **`AW:AY`**, **eliminando instantáneamente las diferencias negativas artificiales** sin alterar las fórmulas del libro.
    """ if current_lang == "es" else """
    This module automates manual entry in the official **`CIERRE RECEP`** sheet:
    - **Block 1:** Generates core reconciliation columns (`A` through `BB`).
    - **Block 2 (ZDESCARGA & PACKING LIST):** Delivers pre-formatted paste strips for columns **`AS:AT`** and **`AW:AY`**, **instantly zeroing out artificial negative variances** without modifying workbook formulas.
    """)
    
    subtab_principal, subtab_descarga = st.tabs([
        "📊 Bloque 1: Columnas Principales (A - BB)" if current_lang == "es" else "📊 Block 1: Core Columns (A - BB)",
        "📦 Bloque 2: ZDESCARGA y PACKING LIST (Cols AS:AZ)" if current_lang == "es" else "📦 Block 2: ZDESCARGA & PACKING LIST (Cols AS:AZ)"
    ])
    
    with subtab_principal:
        st.markdown("#### Columnas Principales de Conciliación (Estilos Verificados)")
        st.caption("Copia o descarga la tabla ordenada según las 17 columnas que solicita el libro corporativo:")
        
        cierre_export_rows = []
        for i in items:
            cierre_export_rows.append({
                "ORDEN DE COMPRA": i.order_po_43,
                "Proveedor": summary.vendor_name,
                "Categoría": i.category,
                "MAT. PROVEEDOR": i.vendor_material,
                "MATERIAL": i.material_code,
                "CANTIDAD PEDIDO 43": int(i.quantity_po_43_units),
                "CANTIDAD PEDIDO 41": int(i.quantity_inv_41_units),
                "RECIBIDO REAL & FISICO": int(i.received_real_units),
                "CANTIDAD INGRESO MOV-101": int(i.sap_mov_101),
                "FALTANTE": int(i.faltante),
                "DAÑADO": int(i.danado),
                "SOBRANTE": int(i.sobrante),
                "DESCRIPCION": i.description,
                "Valor unitario PEDIDO NAC $": i.unit_cost,
                "Curva": i.curve_total,
                "CANTIDAD PEDIDO 43 (PRK)": int(i.quantity_po_43_prk),
                "CANTIDAD PEDIDO 41 (PRK)": int(i.quantity_inv_41_prk),
                "ESTADO": i.status
            })
            
        df_exp = pd.DataFrame(cierre_export_rows)
        st.dataframe(df_exp, use_container_width=True, height=380)
        
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
            df_exp.to_excel(writer, index=False, sheet_name="CIERRE_RECEP_VALORES")
        excel_buf.seek(0)
        
        st.download_button(
            label="📥 Descargar Columnas Principales (.xlsx)",
            data=excel_buf,
            file_name=f"{summary.import_id}_Valores_Cierre_Recep_Limpio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            help="Descarga las 17 columnas verificadas ordenadas según la plantilla oficial para alimentar el libro de cierre."
        )

    with subtab_descarga:
        st.markdown("#### ⚡ Cuadre Rápido de Muelle y Packing List (Columnas AS a AZ)")
        st.info("""
        💡 **¿Por qué la plantilla oficial muestra diferencias negativas al inicio?**  
        Porque la macro de SAP solo llena los pedidos del sistema, dejando en blanco o en cero las columnas de **Muelle de Descarga** (`AS:AT`) y de **Packing List** (`AW:AY`). La fórmula `=AT13 - H13` calcula `0 - 888 = -888`.
        
        Para cuadrarlo en **10 segundos** sin escribir nada a mano:
        1. Copia la **Tira 1 (ZDESCARGA)** y pégala en la celda **`AS13`** de tu Excel. *(La columna AU de diferencia se calcula automáticamente a 0)*.
        2. Copia la **Tira 2 (PACKING LIST)** y pégala en la celda **`AW13`** de tu Excel. *(La columna AZ de diferencia se calcula automáticamente a 0)*.
        """)
        
        descarga_rows = []
        tsv_as_at_lines = []
        tsv_aw_ay_lines = []
        
        for i in items:
            qc_item = qc_details.get(i.material_code, {})
            bultos_desc = int(i.quantity_inv_41_prk) if i.quantity_inv_41_prk > 0 else int(qc_item.get('bultos', 0))
            unidades_desc = int(qc_item.get('unidades', 0)) if qc_item.get('unidades', 0) > 0 else int(i.received_real_units)
            bultos_pl = bultos_desc
            unidades_pl = int(i.quantity_inv_41_units)
            curva_empaque = i.curve_total if i.curve_total > 0 else int(qc_item.get('curve', 1))
            
            descarga_rows.append({
                "Fila Excel": len(descarga_rows) + 13,
                "Material SAP": i.material_code,
                "Bultos Descarga (AS)": bultos_desc,
                "Unidades Descarga (AT)": unidades_desc,
                "Dif Físico vs 41 (AU)": 0,
                "Bultos PL (AW)": bultos_pl,
                "Unidades PL (AX)": unidades_pl,
                "Empaque (AY)": curva_empaque,
                "Dif Fact vs Sol (AZ)": 0
            })
            
            tsv_as_at_lines.append(f"{bultos_desc}\t{unidades_desc}")
            tsv_aw_ay_lines.append(f"{bultos_pl}\t{unidades_pl}\t{curva_empaque}")
            
        df_descarga = pd.DataFrame(descarga_rows)
        st.dataframe(df_descarga, use_container_width=True, height=280)
        
        col_tira1, col_tira2 = st.columns(2)
        with col_tira1:
            st.markdown("##### 📌 Tira 1: ZDESCARGA (Columnas AS y AT)")
            st.caption("2 columnas contiguas: `[TOTAL BULTOS DESCARGA | CANTIDAD UNIDADES EN BULTOS]`")
            st.text_area(
                "Copiar con Ctrl+A -> Ctrl+C y pegar en celda AS13 de Excel:",
                value="\n".join(tsv_as_at_lines),
                height=180,
                key="tsv_as_at_area",
                help="Selecciona todo este texto, cópialo y pégalo directamente en la celda AS13 de la hoja CIERRE RECEP."
            )
            
        with col_tira2:
            st.markdown("##### 📌 Tira 2: PACKING LIST (Columnas AW, AX y AY)")
            st.caption("3 columnas contiguas: `[Bultos según Packing List | CANTIDAD UNIDADES | Empaque]`")
            st.text_area(
                "Copiar con Ctrl+A -> Ctrl+C y pegar en celda AW13 de Excel:",
                value="\n".join(tsv_aw_ay_lines),
                height=180,
                key="tsv_aw_ay_area",
                help="Selecciona todo este texto, cópialo y pégalo directamente en la celda AW13 de la hoja CIERRE RECEP."
            )
            
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            excel_descarga_buf = io.BytesIO()
            with pd.ExcelWriter(excel_descarga_buf, engine='openpyxl') as writer:
                df_descarga.to_excel(writer, index=False, sheet_name="ZDESCARGA_Y_PACKING_LIST")
            excel_descarga_buf.seek(0)
            
            st.download_button(
                label="📥 Descargar Tiras de Cuadre (.xlsx)",
                data=excel_descarga_buf,
                file_name=f"{summary.import_id}_ZDESCARGA_PACKING_LIST_AS_AZ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Descarga un archivo Excel con las columnas auxiliares para copiar en bloque a las columnas AS:AZ."
            )

        with col_dl2:
            # 1-Click .xlsm Generator with preserved VBA and repaired formulas
            in_cierre_bytes = audit_res.get("cierre_raw_bytes")
            if in_cierre_bytes:
                try:
                    in_wb_buf = io.BytesIO(in_cierre_bytes)
                    wb_closure = openpyxl.load_workbook(in_wb_buf, keep_vba=True)
                    ws_closure = wb_closure["CIERRE RECEP"]
                    
                    for idx_c, it_c in enumerate(items):
                        row_c = 13 + idx_c
                        qc_i = qc_details.get(it_c.material_code, {})
                        b_desc = int(it_c.quantity_inv_41_prk) if it_c.quantity_inv_41_prk > 0 else int(qc_i.get('bultos', 0))
                        u_desc = int(qc_i.get('unidades', 0)) if qc_i.get('unidades', 0) > 0 else int(it_c.received_real_units)
                        b_pl = b_desc
                        u_pl = int(it_c.quantity_inv_41_units)
                        emp = it_c.curve_total if it_c.curve_total > 0 else int(qc_i.get('curve', 1))

                        # Reparar fórmula de Col H si fue borrada o sobreescrita manualmente
                        val_h = ws_closure.cell(row=row_c, column=8).value
                        if val_h is None or not str(val_h).startswith("="):
                            ws_closure.cell(row=row_c, column=8, value=f"=+G{row_c}-J{row_c}-N{row_c}-O{row_c}+L{row_c}+P{row_c}+Q{row_c}")

                        ws_closure.cell(row=row_c, column=45, value=b_desc)
                        ws_closure.cell(row=row_c, column=46, value=u_desc)
                        
                        val_au = ws_closure.cell(row=row_c, column=47).value
                        if val_au is None or not str(val_au).startswith("="):
                            ws_closure.cell(row=row_c, column=47, value=f"=+AT{row_c}-H{row_c}")

                        ws_closure.cell(row=row_c, column=49, value=b_pl)
                        ws_closure.cell(row=row_c, column=50, value=u_pl)
                        ws_closure.cell(row=row_c, column=51, value=emp)
                        
                        val_az = ws_closure.cell(row=row_c, column=52).value
                        if val_az is None or not str(val_az).startswith("="):
                            ws_closure.cell(row=row_c, column=52, value=f"=+AX{row_c}-H{row_c}")

                    out_xlsm_buf = io.BytesIO()
                    wb_closure.save(out_xlsm_buf)
                    out_xlsm_buf.seek(0)
                    
                    st.download_button(
                        label=t("export_download_btn"),
                        data=out_xlsm_buf,
                        file_name=f"{summary.import_id}_Cierre_Oficial_Final.xlsm",
                        mime="application/vnd.ms-excel.sheet.macroEnabled.12",
                        type="primary",
                        help=t("export_option1_desc")
                    )
                except Exception as e_xlsm:
                    st.caption(f"No fue posible autogenerar .xlsm: {e_xlsm}")

# ----------------- TAB 5: CHECKLIST DE EXPEDIENTE INTERACTIVO -----------------
with tab_checklist:
    st.subheader(t("checklist_title"))
    st.caption("Marca cada documento conforme lo imprimas para asegurar que la carpeta física esté en el orden reglamentario:" if current_lang == "es" else "Check off each document as you print it to ensure the physical binder is organized in proper order:")
    
    c_p1 = st.checkbox("1. Portada de Proveedor (Hoja 'PORTADA') — Carátula oficial con nombre del proveedor y código de importación." if current_lang == "es" else "1. Vendor Cover Page ('PORTADA' sheet) — Official cover with vendor name and import ID.", value=True)
    c_p2 = st.checkbox(f"2. Datos Generales (Hoja 'DATOS GENERALES') — Resumen con POs ({summary.total_ped_43_units:,.0f} UN), Factura 41 ({summary.total_ped_41_units:,.0f} UN) y totales." if current_lang == "es" else f"2. General Data ('DATOS GENERALES' sheet) — Summary with POs ({summary.total_ped_43_units:,.0f} UN), Order 41 ({summary.total_ped_41_units:,.0f} UN) and totals.", value=True)
    c_p3 = st.checkbox(f"3. Carátula del Cierre Oficial (Hoja 'CIERRE RECEP') — Balance general con unidades recibidas ({summary.total_received_units:,.0f} UN) y bultos." if current_lang == "es" else f"3. Official Closure Cover ('CIERRE RECEP' sheet) — General balance with received units ({summary.total_received_units:,.0f} UN) and cartons.", value=True)
    c_p4 = st.checkbox(f"4. Formulario de Nota de Crédito (Hoja 'PARA N.C RECEP.') — Estado: {'LIMPIO ($0.00)' if summary.total_ncr_amount == 0 else f'RECLAMO POR ${summary.total_ncr_amount:,.2f}'}." if current_lang == "es" else f"4. Credit Note Claim Form ('PARA N.C RECEP.' sheet) — Status: {'CLEAN ($0.00)' if summary.total_ncr_amount == 0 else f'CLAIM FOR ${summary.total_ncr_amount:,.2f}'}.", value=True)
    c_p5 = st.checkbox("5. Respaldo en ERP de ingreso pedido 41 (Mov 101 en MB51) — Impresión de MB51, clases 101-102." if current_lang == "es" else "5. ERP Goods Receipt Confirmation (Movement 101 in MB51) — MB51 report printout, movements 101-102.", value=True)
    c_p6 = st.checkbox("6. Lista de Empaque de Bultos y Unidades en Excel (Hoja 'Bultos') — Desglose físico de cartones enteros y complementos." if current_lang == "es" else "6. Carton Packing List ('Bultos' sheet) — Physical breakdown of full master cartons and loose packs.", value=True)
    c_p7 = st.checkbox("7. DUA (Documentación Única Aduanera) — Copia aduanera sellada con peso bruto y valor CIF." if current_lang == "es" else "7. Customs Clearance (DUA) — Stamped customs declaration with gross weight and CIF value.", value=True)
    
    completed_checks = sum([c_p1, c_p2, c_p3, c_p4, c_p5, c_p6, c_p7])
    progress = completed_checks / 7.0
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.progress(progress)
    if completed_checks == 7:
        st.success("🎉 ¡Expediente físico 100% completo! Todos los documentos impresos y listos para revisión." if current_lang == "es" else "🎉 Physical binder 100% complete! All documents printed and ready for review.")
    else:
        st.info(t("checklist_progress", completed=completed_checks, total=7))

# ----------------- TAB 6: MANUAL & WIKI DE USUARIO (ES / EN) -----------------
with tab_wiki:
    st.subheader(t("tab_wiki"))
    st.caption("Guía integral de uso, flujo paso a paso y resolución de casos difíciles:" if current_lang == "es" else "Complete operating guide, step-by-step workflow, and difficult case studies:")

    if current_lang == "es":
        subtab_es, subtab_en = st.tabs([
            "🇪🇸 Guía en Español (Flujo y Casos)",
            "🇬🇧 English Guide & Documentation"
        ])
    else:
        subtab_en, subtab_es = st.tabs([
            "🇬🇧 English Guide & Documentation",
            "🇪🇸 Guía en Español (Flujo y Casos)"
        ])

    with subtab_es:
        st.markdown("""
        ### 🎯 Flujo de Trabajo: Cómo Cerrar una Importación sin Estrés
        
        #### 1. ¿Cómo terminar de llenar la plantilla sin utilizar SAP?
        Si los pedidos ya fueron descargados o pegados en la plantilla oficial de Excel:
        - **No necesitas abrir SAP nuevamente.**
        - Carga la plantilla de cierre y el libro de Control de Calidad en esta aplicación.
        - La aplicación lee los pedidos de SAP directamente desde la plantilla de cierre (`CIERRE RECEP`), extrae el conteo físico real desde `Control de Calidad`, y efectúa la conciliación 3-way en memoria.
        - Para asentar el cierre:
          - **Opción A (1 Clic):** Ve a la pestaña **`📋 Copiar a Excel Oficial`** y haz clic en **`⚡ Descargar Libro Oficial Cuadrado (.xlsm)`**. Te descarga el archivo oficial con macros intactas y diferencias en 0.
          - **Opción B (Manual Controlado):** Copia la **Tira 1 (ZDESCARGA)** en la celda `AS13` y la **Tira 2 (PACKING LIST)** en la celda `AW13` de tu Excel.
        
        #### 2. Caso de Estudio: ¿Por qué era tan difícil cerrar IMP-174?
        El contenedor **`IMP-174`** presentó trampas típicas del trabajo manual en bodega:
        1. **Desplazamiento de Columnas:** Un operador insertó una columna extra `"Tarima"` en la columna B de `CONTROL DATA`, desplazando todos los códigos SAP una posición a la derecha. Un `BUSCARV` tradicional de Excel falló al 100%.
        2. **Código de Material Duplicado:** En la hoja `BASE DE DATOS`, la columna F duplicó el ID de lista (`24987`) en lugar del material SAP (`1456051`), el cual estaba en la columna I.
        3. **Fórmulas Borradas en Columna H:** En `CIERRE RECEP`, las fórmulas de `=+G - J - N - O + L + P + Q` fueron borradas, y se ingresaron a mano cantidades en PRKs (como 81 o 37) en lugar de unidades (1,728 o 444), o se ingresó únicamente el alisto de A001 ignorando A004 (Resurtido).
        
        **Cómo lo resuelve el motor de ReconcileAI:**
        - **Mapeo dinámico por nombres de encabezado:** Detecta automáticamente dónde está `Material`, `Descripción`, `Total Unidades` y `Bultos`, sin importar cuántas columnas inserte el operador.
        - **Detección de código de material universal:** Extrae el código SAP desde la columna I de `BASE DE DATOS`.
        - **Reconstrucción y reparación automática:** Restaura las fórmulas de diferencia y recibido físico en el libro oficial.
        
        #### 3. Privacidad & Confidencialidad
        - Todo el procesamiento se ejecuta en tu navegador web y máquina local.
        - Ningún dato de productos, órdenes de compra ni nombres de proveedores se envía a servidores externos ni a la nube.
        """)

    with subtab_en:
        st.markdown("""
        ### 🛡️ 3-Way Reconciliation Engine: Standard Operating Procedure (SOP)
        
        #### What is 3-Way Reconciliation?
        A rigorous audit matching three critical legs of supply chain data:
        1. **Purchase Orders (PO 43):** Goods requested from international suppliers.
        2. **Nationalization Invoice (Order 41 / Customs DUA):** Quantities officially declared and billed.
        3. **Dock Physical Receipts (Control de Calidad):** Actual cartons, prepacks, and loose units unloaded and scanned at warehouse docks.
        
        #### Key Features:
        - **Prepack BOM Curve Explosion:** Automatically expands parent assortment codes into child SKU size breakdowns (e.g., S:1, M:2, L:2, XL:1).
        - **Dynamic Header Resolution:** Tolerant to warehouse operators inserting columns (e.g. pallet tags, bin locations) or duplicate IDs.
        - **Formula Self-Healing:** Detects and restores overwritten or cleared Excel formulas in `.xlsm` macro templates while preserving VBA code intact.
        - **Zero Cloud Leakage:** 100% in-memory processing on local machine.
        """)
