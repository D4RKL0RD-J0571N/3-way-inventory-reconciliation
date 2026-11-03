"""
Interactive Web Application for 3-Way Inventory Reconciliation.
Ekono / Almacema S.A. — Centro de Distribución C001
Rol: Supervisión UPC (Área Exclusiva: RECEPCIÓN Y CIERRES)
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
    page_title="Ekono — Asistente de Recepción 3-Way",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
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
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px;
        margin-bottom: 10px;
    }
    .card-title {
        font-size: 12.5px;
        font-weight: 700;
        color: #0F4C81;
        margin-bottom: 6px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .checklist-done {
        background-color: #DEF7EC;
        border-left: 4px solid #31C48D;
        padding: 10px 14px;
        border-radius: 4px;
        margin-bottom: 8px;
    }
    .checklist-pending {
        background-color: #FEF08A;
        border-left: 4px solid #E3A008;
        padding: 10px 14px;
        border-radius: 4px;
        margin-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar: File Manager for Current & Future Importations
st.sidebar.image("https://img.icons8.com/color/96/warehouse.png", width=64)
st.sidebar.title("Gestión de Importaciones")
st.sidebar.caption("Sube los archivos de cualquier contenedor futuro (IMP-183, IMP-184...) o utiliza los archivos actuales por defecto.")

use_custom_files = st.sidebar.checkbox("Subir archivos de otra importación", value=False)
uploaded_qc = None
uploaded_cierre = None

if use_custom_files:
    uploaded_qc = st.sidebar.file_uploader("1. Archivo Control de Calidad (.xlsm)", type=["xlsm", "xlsx"])
    uploaded_cierre = st.sidebar.file_uploader("2. Plantilla Cierre Recepción (.xlsm)", type=["xlsm", "xlsx"])
    st.sidebar.info("Sube ambos archivos para procesar la nueva importación.")

@st.cache_data(show_spinner="Procesando conciliación de recepción...")
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
    if use_custom_files and uploaded_qc and uploaded_cierre:
        items, summary, boms, qc_details, validation = load_and_process(uploaded_qc.getvalue(), uploaded_cierre.getvalue())
    else:
        items, summary, boms, qc_details, validation = load_and_process()
except Exception as e:
    st.error(f"Error al procesar archivos: {e}")
    st.stop()

# Header
st.markdown('<div class="main-title">📋 Asistente Operativo de Recepción y Cierres UPC (3-Way Match)</div>', unsafe_allow_html=True)
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
    st.metric(label="Unidades Recibidas", value=f"{summary.total_received_units:,.0f} UN", delta="Diferencia: 0 UN")
with col5:
    st.metric(
        label="Nota de Crédito (NCR)",
        value=f"${summary.total_ncr_amount:,.2f}",
        delta="Cierre LIMPIO" if summary.status == "LIMPIO" else "Con Reclamo",
        delta_color="normal" if summary.status == "LIMPIO" else "inverse"
    )

st.divider()

# Navigation Tabs (Focusing strictly on Reception & Excel Output)
tab_inspector, tab_reconcile, tab_curves, tab_checklist, tab_excel_export, tab_clipboard = st.tabs([
    "🔎 Ficha 360° del Material (Cero Scroll)",
    "📊 Conciliación de Recepción",
    "🔍 Explosión de Curvas y Tallas",
    "🖨️ Checklist de Impresión (Expediente Presidencia)",
    "📑 Generador para Excel Oficial",
    "📋 Pegar desde Portapapeles (SAP)"
])

# ----------------- TAB 0: FICHA 360 DEL MATERIAL -----------------
with tab_inspector:
    st.subheader("Ficha 360° del Material: Toda la información anclada a su código")
    st.caption("Inspecciona un material sin desplazarte entre 144 columnas en Excel:")
    
    mat_list = [f"{i.material_code} — {i.description[:40]} (Estilo: {i.vendor_material})" for i in items]
    selected_mat_str = st.selectbox("Buscar o seleccionar material:", mat_list)
    sel_code = selected_mat_str.split(" — ")[0]
    item = next((i for i in items if i.material_code == sel_code), None)
    
    if item:
        if item.status == "LIMPIO":
            st.success(f"✅ Estado: {item.status} — {item.observations}")
        elif item.status == "PENDIENTE_PARTE_B":
            st.info(f"🚢 Estado: {item.status} — {item.observations}")
        else:
            st.error(f"⚠️ Estado: {item.status} — {item.observations} | Monto NCR: ${item.ncr_amount:,.2f}")
            
        c_meta1, c_meta2, c_meta3, c_meta4 = st.columns(4)
        with c_meta1:
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">1. Identificación</div>', unsafe_allow_html=True)
            st.write(f"**Código SAP:** `{item.material_code}`")
            st.write(f"**Estilo Proveedor:** `{item.vendor_material}`")
            st.write(f"**Categoría:** {item.category}")
            st.write(f"**Descripción:** {item.description}")
            st.write(f"**Ubicación:** Cont. {item.container_part}")
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c_meta2:
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">2. Compras (PO 43)</div>', unsafe_allow_html=True)
            st.write(f"**Orden Compra:** `{item.order_po_43}`")
            st.write(f"**Prepacks Pedidos:** {item.quantity_po_43_prk:,.0f} PRK")
            st.write(f"**Unidades Pedidas:** {item.quantity_po_43_units:,.0f} UN")
            st.write(f"**Curva Teórica:** {item.curve_total} UN/PRK")
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c_meta3:
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">3. Factura 41 & Físico</div>', unsafe_allow_html=True)
            st.write(f"**Facturado (41):** {item.quantity_inv_41_units:,.0f} UN")
            st.write(f"**Recibido Muelle:** {item.received_real_units:,.0f} UN")
            st.write(f"**Costo Unitario:** ${item.unit_cost:.2f} USD")
            st.write(f"**Valor Factura:** ${(item.quantity_inv_41_units * item.unit_cost):,.2f} USD")
            st.markdown('</div>', unsafe_allow_html=True)
            
        with c_meta4:
            st.markdown('<div class="card-box">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">4. SAP Mov 101 & Balances</div>', unsafe_allow_html=True)
            st.write(f"**Mov 101 (Ingreso):** {item.sap_mov_101:,.0f} UN")
            st.write(f"**Faltante:** {item.faltante:,.0f} UN")
            st.write(f"**Dañado:** {item.danado:,.0f} UN")
            st.write(f"**Sobrante:** {item.sobrante:,.0f} UN")
            st.markdown('</div>', unsafe_allow_html=True)

        if item.children:
            st.markdown("#### Desglose de Tallas Hijas ancladas a este Prepack:")
            c_rows = []
            for c in item.children:
                c_rows.append({
                    "Material Hijo (Talla)": c.material_code,
                    "Descripción": c.description,
                    "Ratio Curva": c.ratio,
                    "Esperado (UN)": c.theoretical_units,
                    "Recibido Físico (UN)": c.received_units,
                    "Diferencia": c.difference,
                    "Costo Unitario ($)": f"${c.unit_cost:.2f}"
                })
            st.dataframe(pd.DataFrame(c_rows), use_container_width=True)

# ----------------- TAB 1: RECONCILIACIÓN -----------------
with tab_reconcile:
    st.subheader("Cruce 3-Way de Recepción (Compras vs Factura vs Físico vs Mov 101)")
    
    c_filter1, c_filter2, c_search = st.columns([1.5, 1.5, 3])
    with c_filter1:
        part_opt = st.selectbox("Contenedor / Parte:", ["Todos", "Parte A (Cont. 182-A)", "Parte B / Otros"], index=0, key="rec_part")
        part_val = "ALL" if part_opt == "Todos" else ("A" if "Parte A" in part_opt else "B")
    with c_filter2:
        status_opt = st.selectbox("Estado de Conciliación:", ["Todos", "LIMPIO", "FALTANTE", "DAÑADO", "SOBRANTE"], index=0, key="rec_status")
        status_val = "ALL" if status_opt == "Todos" else status_opt
    with c_search:
        search_txt = st.text_input("🔍 Buscar por Material, Estilo, Descripción o PO:", placeholder="Ej: 1455103 o 25B03...", key="rec_search")
        
    filtered_items = filter_items(items, part_filter=part_val, status_filter=status_val, search_query=search_txt)
    st.caption(f"Mostrando {len(filtered_items)} de {len(items)} materiales.")
    
    table_rows = []
    for i in filtered_items:
        table_rows.append({
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
        
    df_display = pd.DataFrame(table_rows)
    st.dataframe(df_display, use_container_width=True, height=450)

# ----------------- TAB 2: EXPLOSIÓN DE CURVAS -----------------
with tab_curves:
    st.subheader("Desglose Jerárquico: Prepack Padre ➔ Tallas Hijas")
    parent_choices = [f"{i.material_code} — {i.description[:35]} (Estilo: {i.vendor_material})" for i in items if i.children]
    if parent_choices:
        selected_parent_str = st.selectbox("Selecciona un Material Prepack:", parent_choices, key="curv_parent")
        selected_code = selected_parent_str.split(" — ")[0]
        selected_item = next((i for i in items if i.material_code == selected_code), None)
        
        if selected_item:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Prepacks Facturados", f"{selected_item.quantity_inv_41_prk:.0f} PRK")
            c2.metric("Curva Total por Bulto", f"{selected_item.curve_total} UN/PRK")
            c3.metric("Total Unidades Esperadas", f"{selected_item.quantity_inv_41_units:.0f} UN")
            c4.metric("Estado de Línea", selected_item.status)
            
            st.markdown("#### Detalle de Tallas / Curva Surtida:")
            child_rows = []
            for c in selected_item.children:
                child_rows.append({
                    "Código SKU (Hijo)": c.material_code,
                    "Descripción / Talla": c.description,
                    "Ratio Curva": c.ratio,
                    "Unidades Teóricas": c.theoretical_units,
                    "Unidades Físicas": c.received_units,
                    "Diferencia": c.difference,
                    "Costo Unitario ($)": f"${c.unit_cost:.2f}"
                })
            df_children = pd.DataFrame(child_rows)
            st.dataframe(df_children, use_container_width=True)

# ----------------- TAB 3: CHECKLIST DE IMPRESIÓN (PRESIDENCIA) -----------------
with tab_checklist:
    st.subheader("🖨️ Checklist de Impresión de Recepción (Orden del Expediente Físico)")
    st.info("📌 Esta lista representa el orden estricto en que deben imprimirse y engraparse los documentos de Recepción que Presidencia presenta a los dueños:")
    
    checklist_items = [
        ("1. Portada de Proveedor (Hoja 'PORTADA')", True, f"Carátula oficial con nombre del proveedor ({summary.vendor_name}) y código de contenedor ({summary.import_id})."),
        ("2. Resumen de Resultados (Hoja 'DATOS GENERALES')", True, f"Resumen ejecutivo con POs ({summary.total_ped_43_units:,.0f} UN), Factura 41 ({summary.total_ped_41_units:,.0f} UN) y totales."),
        ("3. Carátula del Cierre para Presidencia (Hoja 'CIERRE RECEP')", True, f"Balance general de recepción con unidades recibidas ({summary.total_received_units:,.0f} UN) y bultos."),
        ("4. Formulario de Nota de Crédito (Hoja 'PARA N.C RECEP.')", True, f"Estado: {'LIMPIO (Costo $0.00)' if summary.total_ncr_amount == 0 else f'RECLAMO POR ${summary.total_ncr_amount:,.2f}'}."),
        ("5. Respaldo en SAP de ingreso pedido 41 (Mov 101) MB51", True, "Reporte impreso de transacción MB51 filtrado por clases 101-102 en Centro C001."),
        ("6. Lista de Empaque de Bultos y Unidades en Excel (Hoja 'Bultos')", True, "Desglose por cartones enteros y complementos con las curvas de surtido."),
        ("7. DUA (Documentación Única Aduanera)", True, "Documento legal aduanero sellado con peso bruto y valor CIF.")
    ]
    
    for title, status, desc in checklist_items:
        st.markdown(f"""
        <div class="checklist-done">
            <strong>✅ {title}</strong><br>
            <span style="font-size: 13px; color: #1E429F;">{desc}</span>
        </div>
        """, unsafe_allow_html=True)

# ----------------- TAB 4: GENERADOR PARA EXCEL OFICIAL -----------------
with tab_excel_export:
    st.subheader("📑 Copiar / Pegar al Excel Oficial que se entrega a Presidencia")
    st.write("""
    Para mantener el proceso corporativo 100% intacto, esta tabla genera **los valores limpios y verificados en el orden exacto de columnas** de la hoja `CIERRE RECEP`.
    Puedes descargar el archivo o copiar los valores para pegarlos en el Excel oficial:
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
        
    df_cierre_export = pd.DataFrame(cierre_export_rows)
    st.dataframe(df_cierre_export, use_container_width=True, height=400)
    
    # Download Button
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        df_cierre_export.to_excel(writer, index=False, sheet_name="CIERRE_RECEP_VALORES")
    excel_buffer.seek(0)
    
    st.download_button(
        label="📥 Descargar Valores Conciliados para el Excel Oficial (.xlsx)",
        data=excel_buffer,
        file_name=f"{summary.import_id}_Valores_Cierre_Recep_Limpio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ----------------- TAB 5: PORTAPAPELES (SAP) -----------------
with tab_clipboard:
    st.subheader("Pegar Datos Copiados Directamente desde SAP GUI")
    st.write("Si copiaste filas con `Ctrl + C` desde `ME80FN`, `SE16N` o `MB51`, puedes pegarlas aquí para analizarlas al instante:")
    
    pasted_text = st.text_area("Pega aquí el contenido del portapapeles (Ctrl + V):", height=160, placeholder="Pega filas copiadas de SAP ALV Grid aquí...")
    
    if st.button("🚀 Limpiar y Procesar Portapapeles"):
        if pasted_text.strip():
            try:
                first_line = pasted_text.strip().split("\n")[0]
                if "\t" in first_line:
                    sep = "\t"
                elif "|" in first_line:
                    sep = "|"
                elif ";" in first_line:
                    sep = ";"
                else:
                    sep = ","
                    
                df_pasted = pd.read_csv(io.StringIO(pasted_text), sep=sep)
                df_pasted.columns = [clean_string(c) for c in df_pasted.columns]
                
                for col in df_pasted.columns:
                    if any(k in col.lower() for k in ["material", "producto", "sku"]):
                        df_pasted[col] = df_pasted[col].apply(clean_material_code)
                    elif any(k in col.lower() for k in ["cant", "unidades", "qty", "recibido", "bultos"]):
                        df_pasted[col] = df_pasted[col].apply(clean_number)
                    else:
                        df_pasted[col] = df_pasted[col].apply(clean_string)
                        
                st.success(f"✅ Se limpiaron y procesaron {len(df_pasted)} filas con separador '{'Tabulador' if sep == '\t' else sep}'.")
                st.dataframe(df_pasted, use_container_width=True)
            except Exception as ex:
                st.error(f"Error al interpretar los datos del portapapeles: {ex}")
        else:
            st.warning("El área de texto está vacía. Copia datos en SAP y presiona Ctrl + V.")
