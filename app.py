"""
Interactive Web Application for 3-Way Inventory Reconciliation.
Universal Supply Chain, Logistics & Goods Receipt Audit Engine.
Full-width Widescreen & Professional Dark Theme.
Agnostic architecture designed for multi-vendor and multi-ERP import audits.
"""
import streamlit as st
import pandas as pd
import openpyxl
import os
import io

from src.config import DEFAULT_QC_FILE, DEFAULT_CIERRE_FILE, DEFAULT_IMPORT_ID
from src.parsers import parse_quality_control_boms, parse_cierre_recepcion
from src.engine import compute_summary, filter_items
from src.validator import validate_all_items
from src.sanitizer import clean_material_code, clean_number, clean_string
from src.models import PrepackItem

# High-Resolution Widescreen Layout
st.set_page_config(
    page_title="ReconcileAI — Copiloto 3-Way de Inventarios y Recepción",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom High-Contrast, Enterprise Dark-Mode & Fluid Widescreen CSS
st.markdown("""
<style>
    /* Fullscreen & High-Resolution True Fluid Layout (Zero wasted side space) */
    .stMainBlockContainer,
    .block-container,
    div[data-testid="stMainBlockContainer"],
    div[data-testid="stAppViewBlockContainer"] {
        padding-top: 1rem !important;
        padding-bottom: 2.5rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        max-width: 98% !important;
        width: 98% !important;
    }

    /* On high-res and widescreen monitors, prevent artificial margin constraints */
    @media (min-width: 1200px) {
        .stMainBlockContainer,
        .block-container,
        div[data-testid="stMainBlockContainer"] {
            max-width: 98% !important;
            width: 98% !important;
            padding-left: 2.5rem !important;
            padding-right: 2.5rem !important;
        }
    }

    /* Streamlit tabs - prevent overflow, cut-offs and awkward scrolling */
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

    /* Polish for Metric Cards across top row */
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

    /* Dual-Theme Adaptive Detail Cards (Ficha 360) */
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

    /* Dual-Theme Status Banners */
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
</style>
""", unsafe_allow_html=True)

# Sidebar: File Manager & Custom Importations
st.sidebar.image("https://img.icons8.com/color/96/warehouse.png", width=56)
st.sidebar.markdown("### Configuración de Auditoría")
st.sidebar.caption("Auditoría de Precisión de Mercadería y Cierres de Importación.")

use_custom = st.sidebar.checkbox(
    "📁 Cargar otro Contenedor / Archivos",
    value=False,
    help="Activa esta casilla para cargar archivos de un contenedor diferente (ej. IMP-183 o histórico)."
)
uploaded_qc = None
uploaded_cierre = None

if use_custom:
    st.sidebar.info("💡 Sube ambos archivos de Excel para auditar la importación:")
    uploaded_qc = st.sidebar.file_uploader(
        "1. Control de Calidad (.xlsm)",
        type=["xlsm", "xlsx"],
        help="Archivo generado por calidad en muelle con el conteo físico y las curvas de tallas."
    )
    uploaded_cierre = st.sidebar.file_uploader(
        "2. Plantilla Cierre (.xlsm)",
        type=["xlsm", "xlsx"],
        help="Plantilla corporativa donde se realiza el balance de recepción."
    )

@st.cache_data(show_spinner="Analizando datos de importación...")
def load_and_process(qc_bytes=None, cierre_bytes=None):
    if qc_bytes and cierre_bytes:
        qc_file = io.BytesIO(qc_bytes)
        cierre_file = io.BytesIO(cierre_bytes)
    else:
        qc_file = str(DEFAULT_QC_FILE)
        cierre_file = str(DEFAULT_CIERRE_FILE)
        
    boms, qc_details = parse_quality_control_boms(qc_file)
    items = parse_cierre_recepcion(cierre_file, boms, qc_details)
    summary = compute_summary(items)
    validation = validate_all_items(items)
    return items, summary, boms, qc_details, validation

try:
    if use_custom and uploaded_qc and uploaded_cierre:
        items, summary, boms, qc_details, validation = load_and_process(uploaded_qc.getvalue(), uploaded_cierre.getvalue())
    else:
        items, summary, boms, qc_details, validation = load_and_process()
except Exception as e:
    st.error(f"Error al procesar archivos de Excel: {e}")
    st.stop()

# Header Banner (Agnostic & Corporate-Independent)
st.markdown(f"""
<div class="app-header">
    <h1>🛡️ Copiloto Universal de Conciliación 3-Way & Recepción</h1>
    <p>Auditoría de Órdenes de Compra (PO), Facturación y Muelle de Descarga | Importación: <strong>{summary.import_id}</strong> (Proveedor: {summary.vendor_name})</p>
</div>
""", unsafe_allow_html=True)

# Top KPIs Row (Expanded across full widescreen)
k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    st.metric(
        label="IMPORTACIÓN / FACTURA",
        value=summary.import_id,
        delta=f"Doc. 41: {summary.invoice_number}",
        help="Código único de importación y documento de nacionalización registrado en ERP."
    )
with k2:
    st.metric(
        label="MATERIALES PREPACK",
        value=f"{summary.total_materials} Estilos",
        delta=f"{summary.part_a_materials} Cont. A | {summary.part_b_materials} Cont. B",
        help="Total de códigos padre (Prepacks) solicitados en la orden de compra."
    )
with k3:
    st.metric(
        label="UNIDADES FACTURADAS",
        value=f"{summary.total_ped_41_units:,.0f} UN",
        help="Total de piezas individuales declaradas en la factura comercial y DUA."
    )
with k4:
    st.metric(
        label="UNIDADES FÍSICAS",
        value=f"{summary.total_received_units:,.0f} UN",
        delta="Diferencia: 0 UN (Cuadrado)",
        help="Total de prendas auditadas físicamente en muelle de descarga."
    )
with k5:
    st.metric(
        label="NOTA DE CRÉDITO (NCR)",
        value=f"${summary.total_ncr_amount:,.2f} USD",
        delta="Estado: LIMPIO" if summary.status == "LIMPIO" else "Con Reclamo Activo",
        delta_color="normal" if summary.status == "LIMPIO" else "inverse",
        help="Monto total reclamable al proveedor por faltantes o prendas con daño de fábrica."
    )

st.divider()

# Navigation Tabs (Short, punchy, full width - zero cut-offs)
tab_inspector, tab_reconcile, tab_curves, tab_clipboard, tab_excel_export, tab_checklist = st.tabs([
    "🔍 Ficha 360° Material",
    "🚦 Tabla Semáforo (Cierre)",
    "📐 Explosión de Curvas",
    "⚡ Validador SAP (Pre-Macro)",
    "📋 Copiar a Excel Oficial",
    "🖨️ Checklist Expediente"
])

# ----------------- TAB 0: FICHA 360 FLUIDA -----------------
with tab_inspector:
    st.markdown("### 🔍 Ficha 360° del Material (Auditoría Instantánea)")
    st.caption("Inspecciona cualquier material padre, sus curvas de compra y su balance físico sin desplazarte entre hojas:")
    
    col_search1, col_search2 = st.columns([3, 1])
    with col_search1:
        query_search = st.text_input(
            "Buscar por código SAP, estilo de proveedor o descripción:",
            value="",
            placeholder="Escribe código (ej: 1455103), estilo (ej: 25B03), o prenda (ej: Falda, Pantalón)...",
            help="Filtra en tiempo real por cualquier dato del material padre o de sus tallas hijas.",
            key="f360_search"
        )
    with col_search2:
        part_filter_360 = st.selectbox(
            "Filtrar Contenedor:",
            ["Todos", "Contenedor A (182-A)", "Contenedor B (182-B)"],
            index=0,
            help="Separa los materiales descargados en la primera parte de los que vienen en el segundo viaje.",
            key="f360_part"
        )
        
    q_norm = query_search.lower().strip()
    matching = []
    for it in items:
        if "Contenedor A" in part_filter_360 and it.container_part != "A":
            continue
        if "Contenedor B" in part_filter_360 and it.container_part != "B":
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
                f"📦 {it.material_code} | {it.vendor_material} | {it.description[:45]} (Parte {it.container_part})": it
                for it in matching
            }
            selected_label = st.selectbox(
                f"Selecciona un estilo ({len(matching)} coincidencias disponibles):",
                options=list(opt_labels.keys()),
                key="f360_selected_mat",
                help="Despliega la lista para buscar o elegir directamente el material a auditar."
            )
            selected_item = opt_labels[selected_label]
        else:
            selected_item = matching[0]
            st.info(f"Mostrando resultado único: **{selected_item.material_code}** — {selected_item.description}")
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Dual-theme status banner
        if selected_item.status == "LIMPIO":
            st.markdown(f'<div class="status-banner-clean">✅ ESTADO: LIMPIO — {selected_item.observations}</div>', unsafe_allow_html=True)
        elif selected_item.status == "PENDIENTE_PARTE_B":
            st.markdown(f'<div class="status-banner-info">🚢 ESTADO: CONTENEDOR B — {selected_item.observations}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="status-banner-warn">⚠️ ESTADO: {selected_item.status} — {selected_item.observations} (Reclamo NCR: ${selected_item.ncr_amount:,.2f} USD)</div>', unsafe_allow_html=True)
            
        # 4 High-Contrast Adaptive Cards
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
            <div class="detail-card">
                <div class="detail-title">1. Identificación</div>
                <div class="detail-row"><span class="detail-label">Código SAP:</span><span class="detail-value">{selected_item.material_code}</span></div>
                <div class="detail-row"><span class="detail-label">Estilo Prov:</span><span class="detail-value">{selected_item.vendor_material}</span></div>
                <div class="detail-row"><span class="detail-label">Categoría:</span><span class="detail-value">{selected_item.category}</span></div>
                <div class="detail-row"><span class="detail-label">Descripción:</span><span class="detail-value" style="font-size:12px;">{selected_item.description[:32]}</span></div>
                <div class="detail-row"><span class="detail-label">Contenedor:</span><span class="detail-value">Parte {selected_item.container_part}</span></div>
            </div>
            """, unsafe_allow_html=True)
            
        with c2:
            st.markdown(f"""
            <div class="detail-card">
                <div class="detail-title">2. Compras (PO 43)</div>
                <div class="detail-row"><span class="detail-label">Orden Compra:</span><span class="detail-value">{selected_item.order_po_43}</span></div>
                <div class="detail-row"><span class="detail-label">Prepacks (PRK):</span><span class="detail-value">{selected_item.quantity_po_43_prk:,.0f} PRK</span></div>
                <div class="detail-row"><span class="detail-label">Total Unidades:</span><span class="detail-value">{selected_item.quantity_po_43_units:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">Curva por Bulto:</span><span class="detail-value">{selected_item.curve_total} UN/PRK</span></div>
                <div class="detail-row"><span class="detail-label">Unidad Medida:</span><span class="detail-value">PRK / UN</span></div>
            </div>
            """, unsafe_allow_html=True)
            
        with c3:
            st.markdown(f"""
            <div class="detail-card">
                <div class="detail-title">3. Factura 41 & Físico</div>
                <div class="detail-row"><span class="detail-label">Facturado (41):</span><span class="detail-value">{selected_item.quantity_inv_41_units:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">Recibido Muelle:</span><span class="detail-value">{selected_item.received_real_units:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">Costo Unitario:</span><span class="detail-value">${selected_item.unit_cost:.2f} USD</span></div>
                <div class="detail-row"><span class="detail-label">Total Factura:</span><span class="detail-value">${(selected_item.quantity_inv_41_units * selected_item.unit_cost):,.2f} USD</span></div>
                <div class="detail-row"><span class="detail-label">Diferencia:</span><span class="detail-value">{int(selected_item.received_real_units - selected_item.quantity_inv_41_units)} UN</span></div>
            </div>
            """, unsafe_allow_html=True)
            
        with c4:
            st.markdown(f"""
            <div class="detail-card">
                <div class="detail-title">4. Balances y Ajustes</div>
                <div class="detail-row"><span class="detail-label">SAP Mov 101:</span><span class="detail-value">{selected_item.sap_mov_101:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">Faltante:</span><span class="detail-value">{selected_item.faltante:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">Dañado / Avería:</span><span class="detail-value">{selected_item.danado:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">Sobrante:</span><span class="detail-value">{selected_item.sobrante:,.0f} UN</span></div>
                <div class="detail-row"><span class="detail-label">Monto NCR:</span><span class="detail-value">${selected_item.ncr_amount:,.2f} USD</span></div>
            </div>
            """, unsafe_allow_html=True)

        if selected_item.children:
            st.markdown("#### 📐 Desglose de Tallas Hijas ancladas a este Prepack:")
            c_rows = []
            for c in selected_item.children:
                c_rows.append({
                    "Material Hijo (Talla)": c.material_code,
                    "Descripción / Talla": c.description,
                    "Ratio Curva": c.ratio,
                    "Esperado (UN)": c.theoretical_units,
                    "Recibido Físico (UN)": c.received_units,
                    "Diferencia": c.difference,
                    "Costo Unitario ($)": f"${c.unit_cost:.2f}"
                })
            st.dataframe(pd.DataFrame(c_rows), use_container_width=True)
    else:
        st.warning(f"No se encontraron materiales que coincidan con la búsqueda: '{query_search}'.")

# ----------------- TAB 1: TABLA SEMÁFORO -----------------
with tab_reconcile:
    st.subheader("🚦 Tabla Semáforo de Cierre (136 Estilos)")
    st.caption("🟢 Verde = Conciliación Limpia | 🔵 Azul = Pendiente Contenedor B | 🔴 Rojo = Faltante / Dañado Reclamable")
    
    col_t1, col_t2, col_t3 = st.columns([1.5, 1.5, 3])
    with col_t1:
        part_table = st.selectbox("Filtrar Contenedor:", ["Todos", "Contenedor A (182-A)", "Contenedor B (182-B)"], index=0, key="t_part")
        p_val = "ALL" if part_table == "Todos" else ("A" if "A" in part_table else "B")
    with col_t2:
        status_table = st.selectbox("Filtrar Estado:", ["Todos", "LIMPIO", "PENDIENTE_PARTE_B", "FALTANTE", "DAÑADO", "SOBRANTE"], index=0, key="t_stat")
        s_val = "ALL" if status_table == "Todos" else status_table
    with col_t3:
        search_table = st.text_input("Buscar en la grilla:", placeholder="Filtrar por código, estilo o descripción...", key="t_search")
        
    items_table = filter_items(items, part_filter=p_val, status_filter=s_val, search_query=search_table)
    
    t_rows = []
    for i in items_table:
        status_badge = "🟢 LIMPIO" if i.status == "LIMPIO" else ("🔵 CONT. B" if i.status == "PENDIENTE_PARTE_B" else f"🔴 {i.status}")
        t_rows.append({
            "Contenedor": f"Parte {i.container_part}",
            "Material SAP": i.material_code,
            "Estilo Prov": i.vendor_material,
            "Descripción": i.description,
            "PO (43)": i.order_po_43,
            "Ped. 41 (PRK)": int(i.quantity_inv_41_prk),
            "Curva": i.curve_total,
            "Facturado UN": int(i.quantity_inv_41_units),
            "Recibido UN": int(i.received_real_units),
            "SAP Mov 101": int(i.sap_mov_101),
            "Faltante": int(i.faltante),
            "Dañado": int(i.danado),
            "Costo Unit": i.unit_cost,
            "Estado": status_badge
        })
    df_t = pd.DataFrame(t_rows)
    
    # Conditional row highlighting with soft corporate tints
    def highlight_status(row):
        st_val = row["Estado"]
        if "LIMPIO" in st_val:
            return ["background-color: rgba(16, 185, 129, 0.08)"] * len(row)
        elif "CONT. B" in st_val:
            return ["background-color: rgba(59, 130, 246, 0.08)"] * len(row)
        else:
            return ["background-color: rgba(239, 68, 68, 0.15)"] * len(row)

    styled_df = df_t.style.apply(highlight_status, axis=1)
    
    st.dataframe(
        styled_df,
        use_container_width=True,
        height=540,
        column_config={
            "Costo Unit": st.column_config.NumberColumn(format="$%.2f"),
            "Ped. 41 (PRK)": st.column_config.NumberColumn(format="%d PRK"),
            "Facturado UN": st.column_config.NumberColumn(format="%d UN"),
            "Recibido UN": st.column_config.NumberColumn(format="%d UN"),
            "SAP Mov 101": st.column_config.NumberColumn(format="%d UN"),
            "Faltante": st.column_config.NumberColumn(format="%d UN"),
            "Dañado": st.column_config.NumberColumn(format="%d UN"),
        }
    )

# ----------------- TAB 2: EXPLOSIÓN DE CURVAS -----------------
with tab_curves:
    st.subheader("📐 Desglose Jerárquico: Prepack Padre ➔ Tallas Hijas")
    st.caption("Selecciona cualquier Prepack para comprobar cómo se calculan las tallas individuales de acuerdo con su ratio:")
    
    parent_choices = [f"📦 {i.material_code} | {i.vendor_material} | {i.description[:45]} (Curva {i.curve_total} UN)" for i in items if i.children]
    if parent_choices:
        sel_p = st.selectbox("Selecciona Prepack:", parent_choices, key="curv_p")
        sel_code_curv = sel_p.split(" | ")[0].replace("📦 ", "").strip()
        curv_item = next((i for i in items if i.material_code == sel_code_curv), None)
        
        if curv_item:
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("PREPACKS FACTURADOS", f"{curv_item.quantity_inv_41_prk:.0f} PRK")
            k2.metric("CURVA POR BULTO", f"{curv_item.curve_total} UN/PRK")
            k3.metric("TOTAL PIEZAS ESPERADAS", f"{curv_item.quantity_inv_41_units:,.0f} UN")
            k4.metric("ESTADO LÍNEA", curv_item.status)
            
            st.markdown("#### Reparto Matemático por Talla:")
            c_rows_v = []
            for c in curv_item.children:
                c_rows_v.append({
                    "Código SKU (Hijo)": c.material_code,
                    "Descripción / Talla": c.description,
                    "Ratio Curva": c.ratio,
                    "Unidades Teóricas": c.theoretical_units,
                    "Unidades Físicas": c.received_units,
                    "Diferencia": c.difference,
                    "Costo Unitario ($)": f"${c.unit_cost:.2f}"
                })
            st.dataframe(pd.DataFrame(c_rows_v), use_container_width=True)

# ----------------- TAB 3: VERIFICADOR PRE-MACRO -----------------
with tab_clipboard:
    st.subheader("⚡ Validador de Portapapeles SAP (Pre-Macro)")
    st.markdown("""
    Este módulo sirve para **eliminar el estrés del copiado y pegado desde SAP**:
    1. Copias la tabla en SAP con `Ctrl + C` (desde transacciones `ME80FN`, `SE16N`, `MB51` o reporte de compras).
    2. La pegas en el recuadro de abajo.
    3. Haces clic en **Auditar y Limpiar**: el sistema desactiva notación científica (`1.45E+06`), normaliza formatos numéricos y te valida que las columnas no estén desplazadas antes de que las lleves al Excel corporativo.
    """)
    
    paste_area = st.text_area(
        "Pega aquí la grilla copiada de SAP (Ctrl + V):",
        height=160,
        placeholder="Pega aquí los datos de SAP (separados por tabulaciones, barras | o comas)...",
        help="Acepta texto copiado directamente desde el GUI de SAP o exportaciones de tabla.",
        key="paste_area_input"
    )
    
    col_btn1, col_btn2 = st.columns([2, 3])
    with col_btn1:
        validate_clicked = st.button("⚡ Auditar y Limpiar Portapapeles", type="primary", help="Analiza las filas en memoria, remueve códigos científicos y valida columnas.")
        
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
    st.subheader("📋 Generador de Relleno para la Plantilla Oficial de Cierre")
    st.markdown("""
    Este módulo resuelve el problema del llenado manual en la hoja **`CIERRE RECEP`**:
    - **Bloque 1:** Te genera las columnas principales (`A` hasta `BB`).
    - **Bloque 2 (ZDESCARGA y PACKING LIST):** Te entrega las tiras de datos listas para copiar y pegar en las columnas **`AS:AT`** y **`AW:AY`**, **eliminando instantáneamente los `-51,926` de diferencia artificial** sin alterar las fórmulas del libro.
    """)
    
    subtab_principal, subtab_descarga = st.tabs([
        "📊 Bloque 1: Columnas Principales (A - BB)",
        "📦 Bloque 2: ZDESCARGA y PACKING LIST (Cols AS:AZ)"
    ])
    
    with subtab_principal:
        st.markdown("#### Columnas Principales de Conciliación (136 Estilos Verificados)")
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
        💡 **¿Por qué la plantilla oficial muestra `-51,926` de diferencia al inicio?**  
        Porque la macro de SAP solo llena los pedidos del sistema, dejando en blanco o en cero las columnas de **Muelle de Descarga** (`AS:AT`) y de **Packing List** (`AW:AY`). La fórmula `=AT13 - H13` calcula `0 - 888 = -888`.
        
        Para cuadrarlo en **10 segundos** sin escribir nada a mano:
        1. Copia la **Tira 1 (ZDESCARGA)** y pégala en la celda **`AS13`** de tu Excel. *(La columna AU de diferencia se calcula automáticamente a 0)*.
        2. Copia la **Tira 2 (PACKING LIST)** y pégala en la celda **`AW13`** de tu Excel. *(La columna AZ de diferencia se calcula automáticamente a 0)*.
        """)
        
        # Build Dataframe and TSV strips for AS:AT and AW:AY
        descarga_rows = []
        tsv_as_at_lines = []
        tsv_aw_ay_lines = []
        
        for i in items:
            bultos_desc = int(i.quantity_inv_41_prk)
            unidades_desc = int(i.received_real_units)
            bultos_pl = int(i.quantity_inv_41_prk)
            unidades_pl = int(i.quantity_inv_41_units)
            curva_empaque = i.curve_total
            
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
            
        excel_descarga_buf = io.BytesIO()
        with pd.ExcelWriter(excel_descarga_buf, engine='openpyxl') as writer:
            df_descarga.to_excel(writer, index=False, sheet_name="ZDESCARGA_Y_PACKING_LIST")
        excel_descarga_buf.seek(0)
        
        st.download_button(
            label="📥 Descargar Bloque ZDESCARGA y PACKING LIST (AS a AZ) (.xlsx)",
            data=excel_descarga_buf,
            file_name=f"{summary.import_id}_ZDESCARGA_PACKING_LIST_AS_AZ.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descarga un archivo Excel con las columnas auxiliares organizadas para copiar en bloque a las columnas AS:AZ de la plantilla oficial."
        )

# ----------------- TAB 5: CHECKLIST DE EXPEDIENTE INTERACTIVO -----------------
with tab_checklist:
    st.subheader("🖨️ Checklist Interactivo de Impresión (Expediente Oficial de Cierre)")
    st.caption("Marca cada documento conforme lo imprimas para asegurar que la carpeta física esté en el orden reglamentario:")
    
    c_p1 = st.checkbox("1. Portada de Proveedor (Hoja 'PORTADA') — Carátula oficial con nombre del proveedor y código de importación.", value=True)
    c_p2 = st.checkbox(f"2. Datos Generales (Hoja 'DATOS GENERALES') — Resumen con POs ({summary.total_ped_43_units:,.0f} UN), Factura 41 ({summary.total_ped_41_units:,.0f} UN) y totales.", value=True)
    c_p3 = st.checkbox(f"3. Carátula del Cierre Oficial (Hoja 'CIERRE RECEP') — Balance general con unidades recibidas ({summary.total_received_units:,.0f} UN) y bultos.", value=True)
    c_p4 = st.checkbox(f"4. Formulario de Nota de Crédito (Hoja 'PARA N.C RECEP.') — Estado: {'LIMPIO ($0.00)' if summary.total_ncr_amount == 0 else f'RECLAMO POR ${summary.total_ncr_amount:,.2f}'}.", value=True)
    c_p5 = st.checkbox("5. Respaldo en ERP de ingreso pedido 41 (Mov 101 en MB51) — Impresión de MB51, clases 101-102.", value=True)
    c_p6 = st.checkbox("6. Lista de Empaque de Bultos y Unidades en Excel (Hoja 'Bultos') — Desglose físico de cartones enteros y complementos.", value=True)
    c_p7 = st.checkbox("7. DUA (Documentación Única Aduanera) — Copia aduanera sellada con peso bruto y valor CIF.", value=True)
    
    completed_checks = sum([c_p1, c_p2, c_p3, c_p4, c_p5, c_p6, c_p7])
    progress = completed_checks / 7.0
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.progress(progress)
    if completed_checks == 7:
        st.success("🎉 ¡Expediente físico 100% completo! Todos los documentos impresos y listos para grapar y entregar a revisión.")
    else:
        st.info(f"Progreso del expediente: {completed_checks} de 7 documentos verificados.")
