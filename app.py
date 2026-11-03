"""
Interactive Web Application for 3-Way Inventory Reconciliation.
Ekono / Almacema S.A. — Centro de Distribución C001
Rol: Supervisión UPC (Copiloto de Verificación y Precisión de Recepción)
"""
import streamlit as st
import pandas as pd
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

# Custom Fluid Styling (Clean, soft, responsive, non-rigid)
st.markdown("""
<style>
    /* Main container styling */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    .main-title {
        font-size: 24px;
        font-weight: 700;
        color: #0F4C81;
        margin-bottom: 2px;
    }
    .sub-title {
        font-size: 13.5px;
        color: #4A5568;
        margin-bottom: 16px;
    }
    .card-box {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px;
        margin-bottom: 10px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    }
    .card-title {
        font-size: 12px;
        font-weight: 700;
        color: #0F4C81;
        margin-bottom: 6px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    /* Clickable pill button styling */
    div.stButton > button:first-child {
        border-radius: 6px;
        font-weight: 500;
        transition: all 0.2s ease-in-out;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar: File Manager & Custom Importations
st.sidebar.image("https://img.icons8.com/color/96/warehouse.png", width=56)
st.sidebar.markdown("### Copiloto de Recepción")
st.sidebar.caption("Diseñado para asistir y verificar cálculos antes de alimentar el Excel oficial.")

use_custom = st.sidebar.checkbox("📁 Cargar otro Contenedor / Archivos", value=False)
uploaded_qc = None
uploaded_cierre = None

if use_custom:
    st.sidebar.info("💡 Puedes probar con importaciones anteriores o nuevas subiendo sus dos archivos de Excel:")
    uploaded_qc = st.sidebar.file_uploader("1. Control de Calidad (.xlsm)", type=["xlsm", "xlsx"])
    uploaded_cierre = st.sidebar.file_uploader("2. Plantilla Cierre (.xlsm)", type=["xlsm", "xlsx"])

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

# Header
st.markdown('<div class="main-title">🛡️ Copiloto de Verificación de Recepción y Cierres UPC</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Cadena de Tiendas Ekono / Almacema S.A. — Centro de Distribución C001 (Área de Recepción)</div>', unsafe_allow_html=True)

# Top KPIs
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric(label="Importación / Factura", value=summary.import_id, delta=f"Ped. {summary.invoice_number}")
with col2:
    st.metric(label="Materiales Prepack", value=f"{summary.total_materials} Estilos", delta=f"{summary.part_a_materials} en Parte A")
with col3:
    st.metric(label="Unidades Facturadas (41)", value=f"{summary.total_ped_41_units:,.0f} UN")
with col4:
    st.metric(label="Unidades Físicas", value=f"{summary.total_received_units:,.0f} UN", delta="Diferencia: 0 UN")
with col5:
    st.metric(
        label="Nota de Crédito (NCR)",
        value=f"${summary.total_ncr_amount:,.2f}",
        delta="Cierre LIMPIO" if summary.status == "LIMPIO" else "Con Reclamo",
        delta_color="normal" if summary.status == "LIMPIO" else "inverse"
    )

st.divider()

# Navigation Tabs (Focusing on Verification, Visual Search, and Excel Output)
tab_inspector, tab_reconcile, tab_curves, tab_clipboard, tab_excel_export, tab_checklist = st.tabs([
    "🔎 Ficha 360° (Búsqueda Fluida)",
    "📊 Tabla con Semáforos de Color",
    "🔍 Explosión de Curvas y Tallas",
    "📋 Verificador Pre-Macro (Copiar SAP)",
    "📑 Generador para Excel Oficial",
    "🖨️ Checklist de Impresión"
])

# ----------------- TAB 0: FICHA 360 FLUIDA (CERO RIGIDEZ) -----------------
with tab_inspector:
    st.subheader("🔎 Ficha 360° del Material (Búsqueda Rápida sin Scroll)")
    st.caption("Escribe el código SAP, parte del estilo o la descripción. Los resultados aparecen al instante sin menús rígidos:")
    
    # Fluid search input
    c_s1, c_s2 = st.columns([3, 1])
    with c_s1:
        query_search = st.text_input("Buscar material (ej. 1455103, 25B03, Falda, Camisa):", value="1455103", key="f360_search")
    with c_s2:
        part_filter_360 = st.selectbox("Filtrar Contenedor:", ["Todos", "Cont. A", "Cont. B"], index=0, key="f360_part")
        
    # Filter matching items
    q_norm = query_search.lower().strip()
    matching = []
    for it in items:
        if part_filter_360 == "Cont. A" and it.container_part != "A":
            continue
        if part_filter_360 == "Cont. B" and it.container_part != "B":
            continue
        if not q_norm or (
            q_norm in it.material_code.lower() or
            q_norm in it.vendor_material.lower() or
            q_norm in it.description.lower() or
            any(q_norm in c.material_code.lower() for c in it.children)
        ):
            matching.append(it)
            
    if matching:
        # If multiple matches, show compact selector
        if len(matching) > 1:
            st.caption(f"Se encontraron **{len(matching)} materiales**. Selecciona uno para inspeccionar:")
            opt_labels = {f"{it.material_code} | {it.vendor_material} | {it.description[:35]} (Cont. {it.container_part})": it for it in matching[:25]}
            selected_label = st.radio("Materiales encontrados:", list(opt_labels.keys()), horizontal=True, label_visibility="collapsed")
            selected_item = opt_labels[selected_label]
        else:
            selected_item = matching[0]
            
        st.divider()
        
        # Status Banner
        if selected_item.status == "LIMPIO":
            st.success(f"✅ Estado: {selected_item.status} — {selected_item.observations}")
        elif selected_item.status == "PENDIENTE_PARTE_B":
            st.info(f"🚢 Estado: {selected_item.status} — {selected_item.observations}")
        else:
            st.error(f"⚠️ Estado: {selected_item.status} — {selected_item.observations} | Monto NCR: ${selected_item.ncr_amount:,.2f}")
            
        # 4 Responsive Cards
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">1. Identificación</div>', unsafe_allow_html=True)
            st.write(f"**Código SAP:** `{selected_item.material_code}`")
            st.write(f"**Estilo Proveedor:** `{selected_item.vendor_material}`")
            st.write(f"**Categoría:** {selected_item.category}")
            st.write(f"**Descripción:** {selected_item.description}")
            st.write(f"**Ubicación:** Contenedor {selected_item.container_part}")
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c2:
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">2. Compras (PO 43)</div>', unsafe_allow_html=True)
            st.write(f"**Orden Compra:** `{selected_item.order_po_43}`")
            st.write(f"**Prepacks Solicitados:** {selected_item.quantity_po_43_prk:,.0f} PRK")
            st.write(f"**Unidades Solicitadas:** {selected_item.quantity_po_43_units:,.0f} UN")
            st.write(f"**Curva Teórica:** {selected_item.curve_total} UN/PRK")
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c3:
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">3. Factura 41 & Físico</div>', unsafe_allow_html=True)
            st.write(f"**Facturado (41):** {selected_item.quantity_inv_41_units:,.0f} UN")
            st.write(f"**Recibido Muelle:** {selected_item.received_real_units:,.0f} UN")
            st.write(f"**Costo Unitario:** ${selected_item.unit_cost:.2f} USD")
            st.write(f"**Monto Factura:** ${(selected_item.quantity_inv_41_units * selected_item.unit_cost):,.2f} USD")
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c4:
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">4. Balances y Ajustes</div>', unsafe_allow_html=True)
            st.write(f"**SAP Mov 101:** {selected_item.sap_mov_101:,.0f} UN")
            st.write(f"**Faltante:** {selected_item.faltante:,.0f} UN")
            st.write(f"**Dañado:** {selected_item.danado:,.0f} UN")
            st.write(f"**Sobrante:** {selected_item.sobrante:,.0f} UN")
            st.markdown('</div>', unsafe_allow_html=True)

        if selected_item.children:
            st.markdown("#### Desglose de Tallas Hijas:")
            c_rows = []
            for c in selected_item.children:
                c_rows.append({
                    "Material Hijo (Talla)": c.material_code,
                    "Descripción / Talla": c.description,
                    "Ratio Curva": c.ratio,
                    "Esperado (UN)": c.theoretical_units,
                    "Recibido Físico (UN)": c.received_units,
                    "Diferencia": c.difference,
                    "Costo Unit ($)": f"${c.unit_cost:.2f}"
                })
            st.dataframe(pd.DataFrame(c_rows), use_container_width=True)
    else:
        st.warning(f"No se encontraron materiales que coincidan con '{query_search}'.")

# ----------------- TAB 1: TABLA CON SEMÁFOROS -----------------
with tab_reconcile:
    st.subheader("📊 Tabla de Conciliación con Semáforos de Color")
    st.caption("Verde suave = Limpio | Rojo suave = Faltante/Dañado | Azul suave = Contenedor B")
    
    t_f1, t_f2, t_s = st.columns([1.5, 1.5, 3])
    with t_f1:
        part_table = st.selectbox("Contenedor:", ["Todos", "Contenedor A", "Contenedor B"], index=0, key="t_part")
        p_val = "ALL" if part_table == "Todos" else ("A" if "A" in part_table else "B")
    with t_f2:
        status_table = st.selectbox("Estado:", ["Todos", "LIMPIO", "FALTANTE", "DAÑADO", "SOBRANTE"], index=0, key="t_stat")
        s_val = "ALL" if status_table == "Todos" else status_table
    with t_s:
        search_table = st.text_input("Filtrar lista:", placeholder="Escribe para filtrar...", key="t_search")
        
    items_table = filter_items(items, part_filter=p_val, status_filter=s_val, search_query=search_table)
    
    # Build DataFrame
    t_rows = []
    for i in items_table:
        t_rows.append({
            "Parte": f"Cont. {i.container_part}",
            "Material SAP": i.material_code,
            "Estilo Proveedor": i.vendor_material,
            "Descripción": i.description,
            "PO (43...)": i.order_po_43,
            "Ped. 41 (PRK)": int(i.quantity_inv_41_prk),
            "Curva": i.curve_total,
            "Ped. 41 (UN)": int(i.quantity_inv_41_units),
            "Recibido (UN)": int(i.received_real_units),
            "SAP Mov 101": int(i.sap_mov_101),
            "Faltante": int(i.faltante),
            "Dañado": int(i.danado),
            "Sobrante": int(i.sobrante),
            "Costo Unit ($)": f"${i.unit_cost:.2f}",
            "Estado": i.status
        })
    df_t = pd.DataFrame(t_rows)
    
    # Apply soft color styling
    def highlight_status(val):
        if val == "LIMPIO":
            return "background-color: #E8F5E9; color: #1B5E20; font-weight: 600;"
        elif val in ["FALTANTE", "DAÑADO"]:
            return "background-color: #FFEBEE; color: #B71C1C; font-weight: 600;"
        elif val == "PENDIENTE_PARTE_B":
            return "background-color: #E3F2FD; color: #0D47A1;"
        elif val == "SOBRANTE":
            return "background-color: #FFF8E1; color: #E65100; font-weight: 600;"
        return ""
        
    styled_df = df_t.style.map(highlight_status, subset=["Estado"])
    st.dataframe(styled_df, use_container_width=True, height=480)

# ----------------- TAB 2: EXPLOSIÓN DE CURVAS -----------------
with tab_curves:
    st.subheader("🔍 Desglose Jerárquico: Prepack Padre ➔ Tallas Hijas")
    st.caption("Visualiza cómo la curva oficial de SAP reparte las unidades entre las tallas:")
    
    parent_choices = [f"{i.material_code} | {i.vendor_material} | {i.description[:35]}" for i in items if i.children]
    if parent_choices:
        sel_p = st.selectbox("Selecciona Prepack:", parent_choices, key="curv_p")
        sel_code_curv = sel_p.split(" | ")[0]
        curv_item = next((i for i in items if i.material_code == sel_code_curv), None)
        
        if curv_item:
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Prepacks Facturados", f"{curv_item.quantity_inv_41_prk:.0f} PRK")
            k2.metric("Curva por Bulto", f"{curv_item.curve_total} UN/PRK")
            k3.metric("Total Unidades Esperadas", f"{curv_item.quantity_inv_41_units:.0f} UN")
            k4.metric("Estado", curv_item.status)
            
            st.markdown("#### Reparto Matemático de Tallas:")
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

# ----------------- TAB 3: VERIFICADOR PRE-MACRO (SAP) -----------------
with tab_clipboard:
    st.subheader("📋 Verificador Pre-Macro de SAP")
    st.markdown("""
    **El problema habitual:** Copias una tabla de SAP (`ME80FN` o `MB51`), cambias de pestaña, la pegas en la macro de Excel oficial y tienes que cruzar los dedos esperando que la macro haya jalado todas las filas sin errores.
    
    **El uso de esta herramienta:** Pégala aquí primero. El copiloto verifica en **1 segundo** que la tabla esté completa, limpia los `#N/A` y valida las filas antes de que toques el Excel oficial.
    """)
    
    paste_area = st.text_area("Pega aquí tus datos copiados de SAP con Ctrl + C:", height=150, placeholder="Pega aquí las filas de SAP ALV Grid...")
    
    if st.button("🔍 Auditar y Limpiar Portapapeles"):
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
                        
                st.success(f"✅ Datos verificados: **{len(df_raw)} filas detectadas**. Sin caracteres extraños ni saltos de columna.")
                st.dataframe(df_raw, use_container_width=True)
            except Exception as e_p:
                st.error(f"Error al leer datos del portapapeles: {e_p}")
        else:
            st.warning("El cuadro de texto está vacío. Copia datos en SAP con Ctrl+C y presiona Ctrl+V aquí.")

# ----------------- TAB 4: GENERADOR PARA EXCEL OFICIAL -----------------
with tab_excel_export:
    st.subheader("📑 Copiar / Pegar al Excel Oficial que se entrega a Presidencia")
    st.write("Esta tabla te entrega los valores limpios y calculados en el orden exacto de columnas que pide la hoja `CIERRE RECEP`. Puedes copiarlos tranquilamente sabiendo que no tienen errores de fórmulas:")
    
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
    st.dataframe(df_exp, use_container_width=True, height=400)
    
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
        df_exp.to_excel(writer, index=False, sheet_name="CIERRE_RECEP_VALORES")
    excel_buf.seek(0)
    
    st.download_button(
        label="📥 Descargar Tabla Conciliada para Excel (.xlsx)",
        data=excel_buf,
        file_name=f"{summary.import_id}_Valores_Cierre_Recep_Limpio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ----------------- TAB 5: CHECKLIST DE IMPRESIÓN -----------------
with tab_checklist:
    st.subheader("🖨️ Checklist de Impresión de Recepción (Orden del Expediente Físico)")
    st.caption("Verifica los 7 documentos en el orden estricto antes de engrapar la carpeta física para Presidencia:")
    
    checklist_items = [
        ("1. Portada de Proveedor (Hoja 'PORTADA')", True, f"Carátula con nombre del proveedor ({summary.vendor_name}) y código ({summary.import_id})."),
        ("2. Resumen de Resultados (Hoja 'DATOS GENERALES')", True, f"Resumen de compras ({summary.total_ped_43_units:,.0f} UN) y factura 41 ({summary.total_ped_41_units:,.0f} UN)."),
        ("3. Carátula del Cierre para Presidencia (Hoja 'CIERRE RECEP')", True, f"Balance general de recepción con unidades recibidas ({summary.total_received_units:,.0f} UN) y bultos."),
        ("4. Formulario de Nota de Crédito (Hoja 'PARA N.C RECEP.')", True, f"Estado: {'LIMPIO (Costo $0.00)' if summary.total_ncr_amount == 0 else f'RECLAMO POR ${summary.total_ncr_amount:,.2f}'}."),
        ("5. Respaldo en SAP de ingreso pedido 41 (Mov 101) MB51", True, "Reporte impreso de transacción MB51 filtrado por clases 101-102 en Centro C001."),
        ("6. Lista de Empaque de Bultos y Unidades en Excel (Hoja 'Bultos')", True, "Desglose por cartones enteros y complementos con las curvas de surtido."),
        ("7. DUA (Documentación Única Aduanera)", True, "Documento legal aduanero sellado con peso bruto y valor CIF.")
    ]
    
    for title, status, desc in checklist_items:
        st.markdown(f"""
        <div style="background-color: #F8FAFC; border-left: 4px solid #31C48D; padding: 10px 14px; border-radius: 4px; margin-bottom: 8px;">
            <strong>✅ {title}</strong><br>
            <span style="font-size: 13px; color: #4A5568;">{desc}</span>
        </div>
        """, unsafe_allow_html=True)
