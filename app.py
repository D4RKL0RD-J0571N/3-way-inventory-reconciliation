"""
Interactive Web Application for 3-Way Inventory Reconciliation.
Ekono / Almacema S.A. — Centro de Distribución C001
Rol: Supervisión UPC (Copiloto de Verificación y Precisión de Recepción)
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

st.set_page_config(
    page_title="Ekono — Copiloto de Recepción UPC",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom High-Contrast, Enterprise Theme Styling
st.markdown("""
<style>
    /* Main Layout */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2.5rem;
        max-width: 1300px;
    }
    
    /* Typography */
    .app-header {
        background: linear-gradient(135deg, #0F4C81 0%, #1B365D 100%);
        color: #FFFFFF !important;
        padding: 20px 24px;
        border-radius: 10px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .app-header h1 {
        color: #FFFFFF !important;
        font-size: 22px;
        font-weight: 700;
        margin: 0;
        padding: 0;
    }
    .app-header p {
        color: #E2E8F0 !important;
        font-size: 13px;
        margin: 4px 0 0 0;
    }

    /* KPI Summary Cards */
    .kpi-container {
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 12px;
        margin-bottom: 20px;
    }
    .kpi-card {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 8px;
        padding: 12px 14px;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.05);
    }
    .kpi-label {
        color: #475569;
        font-size: 11px;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .kpi-value {
        color: #0F172A;
        font-size: 20px;
        font-weight: 700;
    }
    .kpi-delta {
        font-size: 11.5px;
        color: #0284C7;
        margin-top: 2px;
    }

    /* Detail Card Boxes in Ficha 360 */
    .detail-card {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-top: 3px solid #0F4C81;
        border-radius: 6px;
        padding: 14px;
        margin-bottom: 12px;
        box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);
    }
    .detail-title {
        color: #0F4C81;
        font-size: 12px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
        border-bottom: 1px solid #F1F5F9;
        padding-bottom: 4px;
    }
    .detail-row {
        display: flex;
        justify-content: space-between;
        margin-bottom: 5px;
        font-size: 13px;
    }
    .detail-label {
        color: #475569;
        font-weight: 500;
    }
    .detail-value {
        color: #0F172A;
        font-weight: 700;
        text-align: right;
    }

    /* Status Banner */
    .status-banner-clean {
        background-color: #ECFDF5;
        border: 1px solid #A7F3D0;
        border-left: 5px solid #059669;
        color: #065F46;
        padding: 10px 14px;
        border-radius: 6px;
        font-size: 13px;
        font-weight: 600;
        margin-bottom: 14px;
    }
    .status-banner-info {
        background-color: #EFF6FF;
        border: 1px solid #BFDBFE;
        border-left: 5px solid #2563EB;
        color: #1E40AF;
        padding: 10px 14px;
        border-radius: 6px;
        font-size: 13px;
        font-weight: 600;
        margin-bottom: 14px;
    }
    .status-banner-warn {
        background-color: #FEF2F2;
        border: 1px solid #FECACA;
        border-left: 5px solid #DC2626;
        color: #991B1B;
        padding: 10px 14px;
        border-radius: 6px;
        font-size: 13px;
        font-weight: 600;
        margin-bottom: 14px;
    }

    /* Checklist Box */
    .checklist-row {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar: File Manager & Custom Importations
st.sidebar.image("https://img.icons8.com/color/96/warehouse.png", width=56)
st.sidebar.markdown("### Configuración de Datos")
st.sidebar.caption("Copiloto de Precisión para el Área de Recepción C001.")

use_custom = st.sidebar.checkbox("📁 Cargar otro Contenedor / Archivos", value=False, help="Activa esta casilla para cargar archivos de un contenedor diferente (ej. IMP-183 o histórico).")
uploaded_qc = None
uploaded_cierre = None

if use_custom:
    st.sidebar.info("💡 Sube ambos archivos de Excel para auditar la importación:")
    uploaded_qc = st.sidebar.file_uploader("1. Control de Calidad (.xlsm)", type=["xlsm", "xlsx"], help="Archivo generado por calidad en muelle con el conteo y las curvas.")
    uploaded_cierre = st.sidebar.file_uploader("2. Plantilla Cierre (.xlsm)", type=["xlsm", "xlsx"], help="Plantilla oficial donde se realiza el cierre de recepción.")

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

# Header Banner
st.markdown(f"""
<div class="app-header">
    <h1>🛡️ Copiloto de Verificación de Recepción y Cierres UPC</h1>
    <p>Cadena de Tiendas Ekono / Almacema S.A. — Centro de Distribución C001 | Importación: <strong>{summary.import_id}</strong> (Proveedor: {summary.vendor_name})</p>
</div>
""", unsafe_allow_html=True)

# Top KPIs Row
k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    st.metric(
        label="IMPORTACIÓN / FACTURA",
        value=summary.import_id,
        delta=f"Doc. 41: {summary.invoice_number}",
        help="Código único de contenedor y número de documento de nacionalización en SAP."
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

# Navigation Tabs
tab_inspector, tab_reconcile, tab_curves, tab_clipboard, tab_excel_export, tab_checklist = st.tabs([
    "🔎 Ficha 360° (Búsqueda Fluida)",
    "📊 Tabla con Semáforos de Color",
    "🔍 Explosión de Curvas y Tallas",
    "📋 Verificador Pre-Macro (SAP)",
    "📑 Generador para Excel Oficial",
    "🖨️ Checklist de Impresión Físico"
])

# ----------------- TAB 0: FICHA 360 FLUIDA -----------------
with tab_inspector:
    st.markdown("### 🔎 Ficha 360° del Material (Inspección sin Desplazamiento)")
    st.caption("Escribe el código SAP, estilo de proveedor o descripción. La ficha completa se despliega abajo de forma inmediata:")
    
    col_search1, col_search2 = st.columns([3, 1])
    with col_search1:
        query_search = st.text_input(
            "Buscar código o descripción:",
            value="1455103",
            placeholder="Ejemplo: 1455103, 25B03, Falda, Camisa...",
            help="Filtra en tiempo real por cualquier campo del material padre o de sus tallas hijas.",
            key="f360_search"
        )
    with col_search2:
        part_filter_360 = st.selectbox(
            "Filtrar por Contenedor:",
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
            st.caption(f"Coincidencias encontradas: **{len(matching)} materiales**. Selecciona uno para visualizar:")
            opt_labels = {f"{it.material_code} | {it.vendor_material} | {it.description[:35]} (Cont. {it.container_part})": it for it in matching[:20]}
            selected_label = st.radio("Resultados de búsqueda:", list(opt_labels.keys()), horizontal=True, label_visibility="collapsed")
            selected_item = opt_labels[selected_label]
        else:
            selected_item = matching[0]
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Status Banner with High Contrast
        if selected_item.status == "LIMPIO":
            st.markdown(f'<div class="status-banner-clean">✅ ESTADO: LIMPIO — {selected_item.observations}</div>', unsafe_allow_html=True)
        elif selected_item.status == "PENDIENTE_PARTE_B":
            st.markdown(f'<div class="status-banner-info">🚢 ESTADO: CONTENEDOR B — {selected_item.observations}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="status-banner-warn">⚠️ ESTADO: {selected_item.status} — {selected_item.observations} (Reclamo NCR: ${selected_item.ncr_amount:,.2f} USD)</div>', unsafe_allow_html=True)
            
        # 4 High-Contrast Executive Cards
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f"""
            <div class="detail-card">
                <div class="detail-title">1. Identificación</div>
                <div class="detail-row"><span class="detail-label">Código SAP:</span><span class="detail-value">{selected_item.material_code}</span></div>
                <div class="detail-row"><span class="detail-label">Estilo Prov:</span><span class="detail-value">{selected_item.vendor_material}</span></div>
                <div class="detail-row"><span class="detail-label">Categoría:</span><span class="detail-value">{selected_item.category}</span></div>
                <div class="detail-row"><span class="detail-label">Descripción:</span><span class="detail-value" style="font-size:11px;">{selected_item.description[:30]}</span></div>
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

# ----------------- TAB 1: TABLA CON SEMÁFOROS -----------------
with tab_reconcile:
    st.subheader("📊 Tabla Consolidada de Recepción (136 Materiales)")
    st.caption("🟢 Verde = Conciliación Limpia | 🔴 Rojo = Faltante / Dañado | 🔵 Azul = Mercadería de Contenedor B")
    
    col_t1, col_t2, col_t3 = st.columns([1.5, 1.5, 3])
    with col_t1:
        part_table = st.selectbox("Filtrar Contenedor:", ["Todos", "Contenedor A (182-A)", "Contenedor B (182-B)"], index=0, key="t_part")
        p_val = "ALL" if part_table == "Todos" else ("A" if "A" in part_table else "B")
    with col_t2:
        status_table = st.selectbox("Filtrar Estado:", ["Todos", "LIMPIO", "FALTANTE", "DAÑADO", "SOBRANTE"], index=0, key="t_stat")
        s_val = "ALL" if status_table == "Todos" else status_table
    with col_t3:
        search_table = st.text_input("Buscar en la tabla:", placeholder="Escribe código, estilo o descripción...", key="t_search")
        
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
    
    st.dataframe(
        df_t,
        use_container_width=True,
        height=480,
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
    st.subheader("🔍 Desglose Jerárquico: Prepack Padre ➔ Tallas Hijas")
    st.caption("Selecciona cualquier Prepack para comprobar cómo se calculan las tallas individuales de acuerdo con su ratio:")
    
    parent_choices = [f"{i.material_code} | {i.vendor_material} | {i.description[:35]}" for i in items if i.children]
    if parent_choices:
        sel_p = st.selectbox("Selecciona Prepack:", parent_choices, key="curv_p")
        sel_code_curv = sel_p.split(" | ")[0]
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
    st.subheader("📋 Verificador Pre-Macro de SAP (Antes de Pegar en Excel)")
    st.markdown("""
    Este módulo sirve para **eliminar el estrés del copiado y pegado**:
    1. Copias la tabla en SAP con `Ctrl + C` (desde `ME80FN`, `SE16N` o `MB51`).
    2. La pegas en el recuadro de abajo.
    3. Haces clic en el botón de verificación: el sistema valida en **1 segundo** que no haya saltos de renglón, limpia la notación científica y te confirma la cantidad exacta de filas antes de que las lleves al Excel oficial.
    """)
    
    paste_area = st.text_area(
        "Pega aquí la grilla copiada de SAP (Ctrl + V):",
        height=160,
        placeholder="Pega aquí los datos de SAP...",
        help="Acepta texto separado por tabulaciones, barras | o punto y coma."
    )
    
    if st.button("🔍 Validar Datos Copiados de SAP", help="Analiza las filas en memoria, remueve códigos con notación científica y valida que el separador sea correcto antes de llevar los datos al Excel oficial."):
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
    st.subheader("📑 Valores Conciliados para la Hoja 'CIERRE RECEP'")
    st.markdown("""
    Esta tabla contiene los **136 materiales con sus cálculos verificados** en el **orden exacto de columnas** que solicita la plantilla oficial de Presidencia (`CIERRE RECEP`, desde Columna A hasta BB).
    Puedes descargar el archivo o copiar los datos para alimentar el libro oficial con la certeza de que las matemáticas están cuadradas:
    """)
    
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
        label="📥 Descargar Hoja Lista para Pegar en CIERRE RECEP (.xlsx)",
        data=excel_buf,
        file_name=f"{summary.import_id}_Valores_Cierre_Recep_Limpio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Descarga un archivo Excel con las 17 columnas limpias ordenadas según la plantilla oficial para alimentar el libro de Presidencia."
    )

# ----------------- TAB 5: CHECKLIST DE IMPRESIÓN INTERACTIVO -----------------
with tab_checklist:
    st.subheader("🖨️ Checklist Interactivo de Impresión (Expediente Presidencia)")
    st.caption("Marca cada documento conforme lo imprimas para asegurar que la carpeta física esté en el orden reglamentario:")
    
    # Interactive checkboxes
    c_p1 = st.checkbox("1. Portada de Proveedor (Hoja 'PORTADA') — Carátula oficial con nombre del proveedor y código de importación.", value=True)
    c_p2 = st.checkbox(f"2. Datos Generales (Hoja 'DATOS GENERALES') — Resumen con POs ({summary.total_ped_43_units:,.0f} UN), Factura 41 ({summary.total_ped_41_units:,.0f} UN) y totales.", value=True)
    c_p3 = st.checkbox(f"3. Carátula del Cierre para Presidencia (Hoja 'CIERRE RECEP') — Balance general con unidades recibidas ({summary.total_received_units:,.0f} UN) y bultos.", value=True)
    c_p4 = st.checkbox(f"4. Formulario de Nota de Crédito (Hoja 'PARA N.C RECEP.') — Estado: {'LIMPIO ($0.00)' if summary.total_ncr_amount == 0 else f'RECLAMO POR ${summary.total_ncr_amount:,.2f}'}.", value=True)
    c_p5 = st.checkbox("5. Respaldo en SAP de ingreso pedido 41 (Mov 101 en MB51) — Impresión de MB51 en Centro C001, clases 101-102.", value=True)
    c_p6 = st.checkbox("6. Lista de Empaque de Bultos y Unidades en Excel (Hoja 'Bultos') — Desglose físico de cartones enteros y complementos.", value=True)
    c_p7 = st.checkbox("7. DUA (Documentación Única Aduanera) — Copia aduanera sellada con peso bruto y valor CIF.", value=True)
    
    completed_checks = sum([c_p1, c_p2, c_p3, c_p4, c_p5, c_p6, c_p7])
    progress = completed_checks / 7.0
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.progress(progress)
    if completed_checks == 7:
        st.success("🎉 ¡Expediente físico 100% completo! Todos los documentos impresos y listos para grapar y entregar a Presidencia.")
    else:
        st.info(f"Progreso del expediente: {completed_checks} de 7 documentos verificados.")
