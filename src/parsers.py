"""
Robust parsers for Excel files (Cierre Recepción, Control de Calidad, Packing Lists).
Equipped with dynamic header detection to resist column insertions (e.g. Tarima),
corrupted formulas, and dirty strings.
"""
import openpyxl
from typing import Dict, List, Tuple, Any, Optional, Union
from src.models import PrepackItem, ChildSKU
from src.sanitizer import (
    clean_string,
    clean_material_code,
    clean_number,
    clean_int
)

def extract_cierre_metadata(wb_or_path: Union[str, openpyxl.Workbook]) -> Dict[str, Any]:
    """
    Extracts high-level container/import metadata from 'DATOS GENERALES' or top rows of 'CIERRE RECEP'.
    Handles label variances cleanly and never hardcodes company or container names.
    """
    if isinstance(wb_or_path, str):
        wb = openpyxl.load_workbook(wb_or_path, data_only=True)
    else:
        wb = wb_or_path

    meta = {
        "import_id": "IMPORT-ACTIVA",
        "vendor_name": "PROVEEDOR INTERNACIONAL",
        "invoice_number": "PENDIENTE",
        "theor_packages": 0.0,
        "received_packages": 0.0,
        "difference_packages": 0.0,
    }

    # 1. Parse from 'DATOS GENERALES' if present
    if 'DATOS GENERALES' in wb.sheetnames:
        ws_dg = wb['DATOS GENERALES']
        for r in range(1, 25):
            lbl = clean_string(ws_dg.cell(r, 2).value)
            if not lbl:
                continue
            lbl_lower = lbl.lower()
            val = ws_dg.cell(r, 3).value
            
            # Avoid matching 'importadora' when looking for import_id
            if 'importaci' in lbl_lower and 'importadora' not in lbl_lower:
                c_val = clean_string(val)
                if c_val and c_val != "0":
                    meta['import_id'] = c_val
            elif 'nacionalizaci' in lbl_lower or 'pedido nac' in lbl_lower:
                c_val = clean_string(val)
                if c_val:
                    meta['invoice_number'] = c_val
            elif 'proveedor' in lbl_lower:
                c_val = clean_string(val)
                if c_val:
                    meta['vendor_name'] = c_val
            elif 'diferencia' in lbl_lower and 'bulto' in lbl_lower:
                meta['difference_packages'] = clean_number(val)
            elif 'bultos recib' in lbl_lower:
                meta['received_packages'] = clean_number(val)
            elif 'bultos te' in lbl_lower:
                meta['theor_packages'] = clean_number(val)

    # 2. Check top rows of 'CIERRE RECEP' if missing any fields
    if 'CIERRE RECEP' in wb.sheetnames:
        ws_cr = wb['CIERRE RECEP']
        for r in range(1, 10):
            lbl = clean_string(ws_cr.cell(r, 1).value)
            if not lbl:
                continue
            lbl_lower = lbl.lower()
            val = ws_cr.cell(r, 2).value
            
            if 'contenedor' in lbl_lower and meta['import_id'] in ['IMPORT-ACTIVA', '0', '']:
                c_val = clean_string(val)
                if c_val:
                    meta['import_id'] = c_val
            elif ('ped/nac' in lbl_lower or 'nacionalizaci' in lbl_lower) and meta['invoice_number'] == 'PENDIENTE':
                c_val = clean_string(val)
                if c_val:
                    meta['invoice_number'] = c_val
            elif 'bultos te' in lbl_lower and meta['theor_packages'] == 0:
                meta['theor_packages'] = clean_number(val)
            elif 'bultos recib' in lbl_lower and meta['received_packages'] == 0:
                meta['received_packages'] = clean_number(val)
            elif 'dif bultos' in lbl_lower and meta['difference_packages'] == 0:
                meta['difference_packages'] = clean_number(val)

    return meta

def parse_quality_control_boms(qc_source: Any) -> Tuple[Dict[str, List[Dict]], Dict[str, Dict]]:
    """
    Extracts BOM definitions and child details from Control de Calidad sheet BASE DE DATOS & CONTROL DATA.
    Uses dynamic header detection in CONTROL DATA and resilient column indices in BASE DE DATOS.
    Returns:
        (boms_by_parent, qc_details_by_material)
    """
    wb = openpyxl.load_workbook(qc_source, data_only=True)

    # 1. Parse BOM from BASE DE DATOS
    # In both IMP-182 and IMP-174:
    # Col 5 (E): LISTA / BOM ID
    # Col 9 (I): SAP Material Code (fallback to Col 6 (F))
    # Col 10 (J): Ratio/Curva (fallback to Col 7 (G))
    # Col 11 (K): UM ('PRK' for parent, 'UN' for child)
    ws_bd = wb['BASE DE DATOS']
    boms_by_parent: Dict[str, List[Dict]] = {}
    current_bom_id = None
    current_parent = None

    for r in range(5, ws_bd.max_row + 1):
        bom_id = clean_string(ws_bd.cell(r, 5).value) # Col E: BOM ID
        mat_val = ws_bd.cell(r, 9).value or ws_bd.cell(r, 6).value
        mat_clean = clean_material_code(mat_val)
        ratio_val = ws_bd.cell(r, 10).value or ws_bd.cell(r, 7).value
        um_val = clean_string(ws_bd.cell(r, 11).value).upper()

        if bom_id and mat_clean:
            # A new parent row is identified by PRK or BOM ID change
            if bom_id != current_bom_id or um_val == 'PRK':
                current_bom_id = bom_id
                current_parent = mat_clean
                if current_parent not in boms_by_parent:
                    boms_by_parent[current_parent] = []
            else:
                ratio = clean_int(ratio_val, default=1)
                if current_parent and mat_clean != current_parent:
                    boms_by_parent[current_parent].append({
                        'child': mat_clean,
                        'ratio': ratio
                    })

    # 2. Dynamic column detection in CONTROL DATA
    ws_qc = wb['CONTROL DATA']
    col_mat = None
    col_desc = None
    col_um = None
    col_qty = None
    col_curva = None
    col_bultos = None
    col_unidades = None

    # Scan row 6 headers
    for c in range(1, ws_qc.max_column + 1):
        val = clean_string(ws_qc.cell(6, c).value)
        if not val:
            continue
        v_low = val.lower()
        if v_low == 'material':
            col_mat = c
        elif 'descripci' in v_low and not col_desc:
            col_desc = c
        elif v_low == 'um':
            col_um = c
        elif 'total de bultos' in v_low or v_low == 'total bultos':
            col_bultos = c
        elif 'total de unidades' in v_low or v_low == 'total unidades':
            col_unidades = c
        elif 'curva' in v_low and c > 20 and not col_curva:
            col_curva = c
        elif 'cantidad de unidades' in v_low and not col_qty:
            col_qty = c

    # Fallbacks if columns were not found by label
    col_mat = col_mat or 4
    col_desc = col_desc or 12
    col_um = col_um or 11
    col_qty = col_qty or 5
    col_bultos = col_bultos or 32
    col_unidades = col_unidades or 33
    col_curva = col_curva or 35

    qc_details: Dict[str, Dict] = {}
    for r in range(7, ws_qc.max_row + 1):
        m_clean = clean_material_code(ws_qc.cell(r, col_mat).value)
        if m_clean:
            # We record total_unidades, bultos, and curve for complete dock audit
            qc_details[m_clean] = {
                'desc': clean_string(ws_qc.cell(r, col_desc).value),
                'um': clean_string(ws_qc.cell(r, col_um).value).upper(),
                'qty': clean_number(ws_qc.cell(r, col_qty).value),
                'curve': clean_int(ws_qc.cell(r, col_curva).value),
                'bultos': clean_number(ws_qc.cell(r, col_bultos).value),
                'unidades': clean_number(ws_qc.cell(r, col_unidades).value),
            }

    return boms_by_parent, qc_details

class ItemList(list):
    """
    Standard Python list subclass with attached container metadata.
    """
    def __init__(self, items=None, metadata=None):
        super().__init__(items or [])
        self.metadata = metadata or {}

def parse_cierre_recepcion(
    cierre_source: Any,
    boms: Dict[str, List[Dict]],
    qc_details: Dict[str, Dict],
    return_meta: bool = False
) -> Union[ItemList, Tuple[ItemList, Dict[str, Any]]]:
    """
    Parses sheet CIERRE RECEP into structured PrepackItem objects using defensive data sanitization.
    Recovers from missing or corrupted Col H formulas by cross-referencing discharge and QC.
    """
    wb = openpyxl.load_workbook(cierre_source, data_only=True)
    meta = extract_cierre_metadata(wb)
    ws = wb['CIERRE RECEP']

    # Dynamic column mapping on row 12 of CIERRE RECEP
    col_map = {
        'po_43': 1,
        'cat': 3,
        'vendor_mat': 4,
        'mat': 5,
        'ped43_u': 6,
        'ped41_u': 7,
        'rec_u': 8,
        'mov101': 9,
        'falt': 14,
        'dan': 15,
        'sob': 16,
        'mov641': 21,
        'desc': 30,
        'unit_cost': 32,
        'as_bultos': 45,
        'at_unidades': 46,
        'aw_bultos': 49,
        'ax_unidades': 50,
        'curve': 54,
        'ped43_prk': 56,
        'ped41_prk': 57,
    }

    # Scan row 12 to refine positions if present
    for c in range(1, min(ws.max_column + 1, 65)):
        hdr = clean_string(ws.cell(12, c).value)
        if not hdr:
            continue
        h_low = hdr.lower()
        if 'material' == h_low:
            col_map['mat'] = c
        elif 'mat. proveedor' in h_low or 'material proveedor' in h_low:
            col_map['vendor_mat'] = c
        elif 'orden de compra' in h_low:
            col_map['po_43'] = c
        elif 'cantidad pedido 43' in h_low and c < 20:
            col_map['ped43_u'] = c
        elif 'cantidad pedido 41' in h_low and c < 20:
            col_map['ped41_u'] = c
        elif 'recibido real' in h_low:
            col_map['rec_u'] = c
        elif 'mov-101' in h_low and c < 20:
            col_map['mov101'] = c
        elif 'total bultos' in h_low and 'descarga' in h_low:
            col_map['as_bultos'] = c
        elif 'unidades en bultos' in h_low:
            col_map['at_unidades'] = c

    items: List[PrepackItem] = []

    for r in range(13, ws.max_row + 1):
        mat_code = clean_material_code(ws.cell(r, col_map['mat']).value)
        if not mat_code:
            continue

        vendor_mat = clean_string(ws.cell(r, col_map['vendor_mat']).value)
        po_43 = clean_string(ws.cell(r, col_map['po_43']).value)
        category = clean_string(ws.cell(r, col_map['cat']).value)
        desc = clean_string(ws.cell(r, col_map['desc']).value)

        ped43_u = clean_number(ws.cell(r, col_map['ped43_u']).value)
        ped41_u = clean_number(ws.cell(r, col_map['ped41_u']).value)
        raw_rec_u = ws.cell(r, col_map['rec_u']).value
        mov101 = clean_number(ws.cell(r, col_map['mov101']).value)
        mov641 = clean_number(ws.cell(r, col_map['mov641']).value)

        falt = clean_number(ws.cell(r, col_map['falt']).value)
        dan = clean_number(ws.cell(r, col_map['dan']).value)
        sob = clean_number(ws.cell(r, col_map['sob']).value)

        unit_cost = clean_number(ws.cell(r, col_map['unit_cost']).value)
        curve = clean_int(ws.cell(r, col_map['curve']).value, default=12)

        ped43_prk = clean_number(ws.cell(r, col_map['ped43_prk']).value)
        ped41_prk = clean_number(ws.cell(r, col_map['ped41_prk']).value)

        at_unidades = clean_number(ws.cell(r, col_map['at_unidades']).value)

        # Physical receipt resolution:
        # 1. If Col H (RECIBIDO REAL) has a non-zero value, respect it.
        # 2. If Col H is empty/zero (common when unclosed), fallback to Col AT (DESCARGA UNIDADES).
        # 3. If still empty, fallback to QC CONTROL DATA 'TOTAL DE UNIDADES'.
        if raw_rec_u is not None and str(raw_rec_u).strip() not in ["", "0", "None"]:
            rec_u = clean_number(raw_rec_u)
        elif at_unidades > 0:
            rec_u = at_unidades
        elif mat_code in qc_details and qc_details[mat_code].get('unidades', 0) > 0:
            rec_u = qc_details[mat_code]['unidades']
        else:
            rec_u = 0.0

        # If description is missing in closure sheet, fill from QC
        if not desc and mat_code in qc_details:
            desc = qc_details[mat_code].get('desc', '')

        # Determine container part (Part A if in BOMs or arrived in QC dock, otherwise Part B)
        is_in_qc = mat_code in qc_details and qc_details[mat_code].get('unidades', 0) > 0
        is_in_bom = mat_code in boms
        container_part = "A" if (is_in_bom or is_in_qc) else "B"

        # Determine status
        diff = rec_u - ped41_u
        if container_part == "B" and rec_u == 0:
            status = "PENDIENTE_PARTE_B"
            obs = "Mercadería en Contenedor Parte B o pendiente de descarga."
        elif falt > 0:
            status = "FALTANTE"
            obs = f"Faltante registrado de {falt:,.0f} unidades."
        elif dan > 0:
            status = "DAÑADO"
            obs = f"Prendas dañadas registradas: {dan:,.0f} unidades."
        elif sob > 0:
            status = "SOBRANTE"
            obs = f"Sobrante registrado de {sob:,.0f} unidades."
        elif diff < 0:
            status = "FALTANTE"
            obs = f"Diferencia negativa de {abs(diff):,.0f} unidades."
        elif diff > 0:
            status = "SOBRANTE"
            obs = f"Diferencia positiva de {diff:,.0f} unidades."
        else:
            status = "LIMPIO"
            obs = "Conciliación exacta sin discrepancias."

        ncr_amount = (falt + dan) * unit_cost

        # Build children list if BOM exists
        children: List[ChildSKU] = []
        if mat_code in boms:
            for b in boms[mat_code]:
                c_code = b['child']
                c_ratio = b['ratio']
                if ped41_prk > 0:
                    c_theo = int(round(ped41_prk * c_ratio))
                elif curve > 0:
                    c_theo = int(round((ped41_u / curve) * c_ratio))
                else:
                    c_theo = 0

                c_desc = qc_details.get(c_code, {}).get('desc', f"{desc} (Talla {c_ratio})")
                c_rec = int(c_theo) if status == "LIMPIO" else 0
                children.append(ChildSKU(
                    material_code=c_code,
                    description=c_desc,
                    ratio=c_ratio,
                    theoretical_units=c_theo,
                    received_units=c_rec,
                    difference=c_rec - c_theo,
                    unit_cost=unit_cost
                ))

        item = PrepackItem(
            material_code=mat_code,
            vendor_material=vendor_mat,
            description=desc,
            category=category,
            order_po_43=po_43,
            quantity_po_43_prk=ped43_prk,
            quantity_po_43_units=ped43_u,
            quantity_inv_41_prk=ped41_prk,
            quantity_inv_41_units=ped41_u,
            received_real_units=rec_u,
            curve_total=curve,
            unit_cost=unit_cost,
            container_part=container_part,
            sap_mov_101=mov101,
            sap_mov_641=mov641,
            children=children,
            faltante=falt,
            danado=dan,
            sobrante=sob,
            status=status,
            ncr_amount=ncr_amount,
            observations=obs
        )
        items.append(item)

    result_list = ItemList(items, metadata=meta)

    if return_meta:
        return result_list, meta
    return result_list
