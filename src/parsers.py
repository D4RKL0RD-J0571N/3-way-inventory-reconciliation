"""
Robust parsers for Excel files (Cierre Recepción, Control de Calidad, Packing Lists).
Equipped with dynamic header detection to resist column insertions (e.g. Tarima),
corrupted formulas, and dirty strings.
"""
import re
import openpyxl
from typing import Dict, List, Tuple, Any, Optional, Union
from src.models import PrepackItem, ChildSKU, PackagingBreakdown
from src.sanitizer import (
    clean_string,
    clean_material_code,
    clean_number,
    clean_int
)

def parse_packaging_string(val: Any) -> PackagingBreakdown:
    """
    Parses carton counts, packaging strings, and compound breakdowns (e.g. '39*60 1*78', '50*12', 42).
    """
    if val is None:
        return PackagingBreakdown()
    s = str(val).strip()
    if not s or s.lower() in ["none", "nan", "null"]:
        return PackagingBreakdown()

    # Try direct numeric
    try:
        f = float(s)
        return PackagingBreakdown(
            total_boxes=f,
            total_units=0.0,
            boxes_full=f,
            raw_str=s,
            is_compound=False
        )
    except ValueError:
        pass

    # Find patterns like '39*60 1*78' or '39 * 60 + 1 * 48' or '10x12'
    parts = re.findall(r'(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)', s)
    if parts:
        total_boxes = sum(float(b) for b, u in parts)
        total_units = sum(float(b) * float(u) for b, u in parts)
        boxes_full = float(parts[0][0])
        units_full = float(parts[0][1])
        boxes_comp = float(parts[1][0]) if len(parts) > 1 else 0.0
        units_comp = float(parts[1][1]) if len(parts) > 1 else 0.0
        
        return PackagingBreakdown(
            total_boxes=total_boxes,
            total_units=total_units,
            boxes_full=boxes_full,
            units_per_box_full=units_full,
            boxes_comp=boxes_comp,
            units_per_box_comp=units_comp,
            raw_str=s,
            is_compound=len(parts) > 1 or total_boxes != boxes_full
        )

    # Fallback: extract first number
    m = re.search(r'\d+(?:\.\d+)?', s)
    if m:
        f = float(m.group(0))
        return PackagingBreakdown(
            total_boxes=f,
            boxes_full=f,
            raw_str=s,
            is_compound=False
        )

    return PackagingBreakdown(raw_str=s)

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

        # Fallback: compute package sums directly from rows if still zero
        if meta['received_packages'] == 0.0 or meta['theor_packages'] == 0.0:
            rec_sum = 0.0
            theo_sum = 0.0
            for r in range(13, min(ws_cr.max_row + 1, 500)):
                mat_v = clean_material_code(ws_cr.cell(r, 5).value)
                if not mat_v:
                    continue
                v_as = ws_cr.cell(r, 45).value
                v_aw = ws_cr.cell(r, 49).value
                rec_sum += parse_packaging_string(v_as).total_boxes
                theo_sum += parse_packaging_string(v_aw).total_boxes
            if meta['received_packages'] == 0.0 and rec_sum > 0:
                meta['received_packages'] = rec_sum
            if meta['theor_packages'] == 0.0 and theo_sum > 0:
                meta['theor_packages'] = theo_sum
            if meta['difference_packages'] == 0.0:
                meta['difference_packages'] = meta['received_packages'] - meta['theor_packages']

    return meta

def parse_quality_control_boms(qc_source: Any) -> Tuple[Dict[str, List[Dict]], Dict[str, Dict]]:
    """
    Extracts BOM definitions and child details from Control de Calidad sheet BASE DE DATOS & CONTROL DATA.
    Uses dynamic header detection in CONTROL DATA and resilient column indices in BASE DE DATOS.
    Returns:
        (boms_by_parent, qc_details_by_material)
    """
    # Open workbook in streaming read_only mode for ultra-fast XML parsing (50x speedup)
    try:
        wb = openpyxl.load_workbook(qc_source, data_only=True, read_only=True)
    except Exception:
        wb = openpyxl.load_workbook(qc_source, data_only=True)

    # 1. Parse BOM from BASE DE DATOS (if present)
    bd_sheet_name = next((s for s in wb.sheetnames if any(k in s.upper() for k in ['BASE DE DATOS', 'BASE_DATOS', 'BOM', 'CURVAS'])), None)
    boms_by_parent: Dict[str, List[Dict]] = {}
    if bd_sheet_name:
        ws_bd = wb[bd_sheet_name]
        current_bom_id = None
        current_parent = None
        empty_streak = 0

        for row in ws_bd.iter_rows(min_row=5, values_only=True):
            if not row or len(row) < 5:
                empty_streak += 1
                if empty_streak > 30:
                    break
                continue
            bom_id = clean_string(row[4]) if len(row) > 4 else ""
            mat_val = (row[8] if len(row) > 8 else None) or (row[5] if len(row) > 5 else None)
            mat_clean = clean_material_code(mat_val)
            ratio_val = (row[9] if len(row) > 9 else None) or (row[6] if len(row) > 6 else None)
            um_val = clean_string(row[10] if len(row) > 10 else "").upper()

            if not bom_id and not mat_clean:
                empty_streak += 1
                if empty_streak > 30:
                    break
                continue
            empty_streak = 0

            if bom_id and mat_clean:
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
    qc_sheet_name = next((s for s in wb.sheetnames if any(k in s.upper() for k in ['CONTROL DATA', 'CONTROL_DATA', 'CONTROL', 'QC'])), wb.sheetnames[0])
    ws_qc = wb[qc_sheet_name]
    col_mat, col_part, col_desc, col_um, col_qty = 3, None, 11, 10, 4 # 0-indexed fallbacks
    col_curva, col_bultos, col_unidades = 34, 31, 32
    col_bultos_enteros, col_unid_enteros = 27, 28
    col_bultos_comp, col_unid_comp, col_total_prk = 29, 30, 33

    # Scan row 6 headers
    headers_row = None
    for r_idx, row in enumerate(ws_qc.iter_rows(min_row=1, max_row=6, values_only=True), start=1):
        if r_idx == 6:
            headers_row = row
            break

    if headers_row:
        for c_idx, val in enumerate(headers_row):
            s_val = clean_string(val)
            if not s_val:
                continue
            v_low = s_val.lower()
            if v_low == 'material':
                col_mat = c_idx
            elif any(k in v_low for k in ['parte', 'contenedor', 'container', 'lote']):
                col_part = c_idx
            elif 'descripci' in v_low and col_desc == 11:
                col_desc = c_idx
            elif v_low == 'um':
                col_um = c_idx
            elif 'bultos enteros' in v_low:
                col_bultos_enteros = c_idx
            elif 'unidades por bulto' in v_low or 'unidades por bultos' in v_low:
                col_unid_enteros = c_idx
            elif 'bultos complemento' in v_low:
                col_bultos_comp = c_idx
            elif 'unidades bultos complemento' in v_low or 'unidades complemento' in v_low:
                col_unid_comp = c_idx
            elif 'total de bultos' in v_low or v_low == 'total bultos':
                col_bultos = c_idx
            elif 'total de unidades' in v_low or v_low == 'total unidades':
                col_unidades = c_idx
            elif 'total prk' in v_low:
                col_total_prk = c_idx
            elif 'curva' in v_low and c_idx > 20:
                col_curva = c_idx
            elif 'cantidad de unidades' in v_low:
                col_qty = c_idx

    # Detect multi-part distribution sheets dynamically (e.g. DIST. PARTE A, PARTE B, PARTE C...)
    part_materials: Dict[str, str] = {}
    detected_parts = set()
    for sname in wb.sheetnames:
        m_part = re.search(r'PARTE\s+([A-Z0-9]+)', sname, re.IGNORECASE)
        if m_part:
            p_code = m_part.group(1).upper()
            detected_parts.add(p_code)
            try:
                ws_part = wb[sname]
                for r_p, row_p in enumerate(ws_part.iter_rows(min_row=1, max_row=500, values_only=True), start=1):
                    if not row_p:
                        continue
                    for val_p in row_p[:6]:
                        val_s = str(val_p or '').strip()
                        if val_s.isdigit() and len(val_s) >= 7:
                            clean_m = clean_material_code(val_s)
                            if clean_m and clean_m not in part_materials:
                                part_materials[clean_m] = p_code
            except Exception:
                pass

    qc_details: Dict[str, Dict] = {}
    empty_streak = 0
    for row in ws_qc.iter_rows(min_row=7, values_only=True):
        if not row or col_mat >= len(row):
            empty_streak += 1
            if empty_streak > 30:
                break
            continue
        m_clean = clean_material_code(row[col_mat])
        if not m_clean:
            empty_streak += 1
            if empty_streak > 30:
                break
            continue
        empty_streak = 0

        assigned_part = part_materials.get(m_clean, "A" if detected_parts else "")

        qc_details[m_clean] = {
            'desc': clean_string(row[col_desc]) if col_desc < len(row) else '',
            'um': clean_string(row[col_um]).upper() if col_um < len(row) else 'PRK',
            'qty': clean_number(row[col_qty]) if col_qty < len(row) else 0.0,
            'curve': clean_int(row[col_curva]) if col_curva < len(row) else 12,
            'bultos': clean_number(row[col_bultos]) if col_bultos < len(row) else 0.0,
            'unidades': clean_number(row[col_unidades]) if col_unidades < len(row) else 0.0,
            'bultos_enteros': clean_number(row[col_bultos_enteros]) if col_bultos_enteros and col_bultos_enteros < len(row) else 0.0,
            'unidades_por_bulto_entero': clean_number(row[col_unid_enteros]) if col_unid_enteros and col_unid_enteros < len(row) else 0.0,
            'bultos_complementos': clean_number(row[col_bultos_comp]) if col_bultos_comp and col_bultos_comp < len(row) else 0.0,
            'unidades_por_bulto_comp': clean_number(row[col_unid_comp]) if col_unid_comp and col_unid_comp < len(row) else 0.0,
            'total_prk': clean_number(row[col_total_prk]) if col_total_prk and col_total_prk < len(row) else 0.0,
            'part': assigned_part,
        }

    if detected_parts:
        qc_details['_detected_parts'] = sorted(list(detected_parts))

    try:
        wb.close()
    except Exception:
        pass

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
    boms: Optional[Dict[str, List[Dict]]] = None,
    qc_details: Optional[Dict[str, Dict]] = None,
    return_meta: bool = False,
    is_blind_setup: bool = False
) -> Union[ItemList, Tuple[ItemList, Dict[str, Any]]]:
    """
    Parses sheet CIERRE RECEP into structured PrepackItem objects using defensive data sanitization.
    Supports blind setup mode (montaje a ciegas desde SAP en pausa sin Control de Calidad).
    Recovers from missing or corrupted Col H formulas by cross-referencing discharge and QC.
    Accurately segregates physical cartons, prepacks (PRK), and total units.
    """
    boms_map = boms or {}
    qc_map = qc_details or {}
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
        'ay_empaque': 51,
        'curve': 54,
        'um': 55,
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
        elif 'faltante' in h_low and c < 20:
            col_map['falt'] = c
        elif any(k in h_low for k in ['dañado', 'danado']) and c < 20:
            col_map['dan'] = c
        elif 'sobrante' in h_low and c < 20:
            col_map['sob'] = c
        elif any(k in h_low for k in ['partes', 'b-c-d', 'enviada a partes']) and c < 30:
            col_map['partes_bcd'] = c
        elif 'total bultos' in h_low and 'descarga' in h_low:
            col_map['as_bultos'] = c
        elif 'unidades en bultos' in h_low:
            col_map['at_unidades'] = c
        elif 'bultos' in h_low and 'packing' in h_low:
            col_map['aw_bultos'] = c
        elif 'cantidad unidades' in h_low and c > 40:
            col_map['ax_unidades'] = c
        elif 'empaque' in h_low and c > 40:
            col_map['ay_empaque'] = c
        elif 'curva' in h_low and c > 40:
            col_map['curve'] = c
        elif 'um' == h_low and c > 40:
            col_map['um'] = c
        elif 'cantidad pedido 43' in h_low and c > 40:
            col_map['ped43_prk'] = c
        elif 'cantidad pedido 41' in h_low and c > 40:
            col_map['ped41_prk'] = c

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

        # Parse packaging and cartons
        raw_as = ws.cell(r, col_map['as_bultos']).value
        raw_at = ws.cell(r, col_map['at_unidades']).value
        raw_aw = ws.cell(r, col_map['aw_bultos']).value
        raw_ax = ws.cell(r, col_map['ax_unidades']).value
        raw_ay = ws.cell(r, col_map['ay_empaque']).value
        raw_um = ws.cell(r, col_map['um']).value

        pkg_as = parse_packaging_string(raw_as)
        pkg_aw = parse_packaging_string(raw_aw)
        empaque_str = clean_string(raw_ay)
        um_str = clean_string(raw_um).upper() or "PRK"

        # Bultos Descarga (Muelle)
        bultos_desc = pkg_as.total_boxes
        unidades_desc = clean_number(raw_at)
        if bultos_desc == 0.0 and mat_code in qc_map and qc_map[mat_code].get('bultos', 0) > 0:
            bultos_desc = qc_map[mat_code]['bultos']
        if unidades_desc == 0.0 and mat_code in qc_map and qc_map[mat_code].get('unidades', 0) > 0:
            unidades_desc = qc_map[mat_code]['unidades']
        elif unidades_desc == 0.0 and pkg_as.total_units > 0:
            unidades_desc = pkg_as.total_units

        # Bultos Packing List
        bultos_pl = pkg_aw.total_boxes
        unidades_pl = clean_number(raw_ax)
        if bultos_pl == 0.0 and bultos_desc > 0:
            bultos_pl = bultos_desc
        if unidades_pl == 0.0 and unidades_desc > 0:
            unidades_pl = unidades_desc
        elif unidades_pl == 0.0 and ped41_u > 0:
            unidades_pl = ped41_u

        # Cartons Breakdown (Full cartons & Complement cartons)
        b_enteros = pkg_as.boxes_full
        u_enteros = pkg_as.units_per_box_full
        b_comp = pkg_as.boxes_comp
        u_comp = pkg_as.units_per_box_comp

        if b_enteros == 0.0 and mat_code in qc_map:
            b_enteros = qc_map[mat_code].get('bultos_enteros', 0.0)
            u_enteros = qc_map[mat_code].get('unidades_por_bulto_entero', 0.0)
            b_comp = qc_map[mat_code].get('bultos_complementos', 0.0)
            u_comp = qc_map[mat_code].get('unidades_por_bulto_comp', 0.0)

        # Average units per carton
        if bultos_desc > 0 and unidades_desc > 0:
            u_per_b = round(unidades_desc / bultos_desc, 2)
        elif bultos_pl > 0 and unidades_pl > 0:
            u_per_b = round(unidades_pl / bultos_pl, 2)
        elif empaque_str and empaque_str.replace('.', '', 1).isdigit():
            u_per_b = float(empaque_str)
        else:
            u_per_b = float(curve)

        dif_bultos = bultos_desc - bultos_pl

        # Physical receipt resolution:
        # 1. If Col H (RECIBIDO REAL) has a non-zero value, respect it.
        # 2. If Col H is empty/zero (common when unclosed), fallback to Col AT (DESCARGA UNIDADES).
        # 3. If still empty, fallback to QC CONTROL DATA 'TOTAL DE UNIDADES'.
        if raw_rec_u is not None and str(raw_rec_u).strip() not in ["", "0", "None"]:
            rec_u = clean_number(raw_rec_u)
        elif unidades_desc > 0:
            rec_u = unidades_desc
        elif mat_code in qc_map and qc_map[mat_code].get('unidades', 0) > 0:
            rec_u = qc_map[mat_code]['unidades']
        else:
            rec_u = 0.0

        # If description is missing in closure sheet, fill from QC
        if not desc and mat_code in qc_map:
            desc = qc_map[mat_code].get('desc', '')

        # Determine container part dynamically
        is_in_qc = mat_code in qc_map and qc_map[mat_code].get('unidades', 0) > 0
        detected_parts = qc_map.get('_detected_parts', [])
        assigned_part = qc_map.get(mat_code, {}).get('part', '')

        # Check observations or action columns in CIERRE RECEP if unassigned
        if not assigned_part:
            partes_bcd_val = clean_number(ws.cell(r, col_map['partes_bcd']).value) if col_map.get('partes_bcd') else 0.0
            if partes_bcd_val > 0:
                assigned_part = "B"
            else:
                obs_text = f"{clean_string(ws.cell(r, col_map.get('desc', 30)).value)} {clean_string(ws.cell(r, 20).value)} {clean_string(ws.cell(r, 29).value)}"
                p_match = re.search(r'PARTE\s+([A-Z0-9]+)', obs_text, re.IGNORECASE)
                if p_match:
                    assigned_part = p_match.group(1).upper()

        if not assigned_part:
            if is_in_qc or (mat_code in boms_map):
                assigned_part = "A" if detected_parts else ""
            elif detected_parts:
                other_parts = [p for p in detected_parts if p != "A"]
                assigned_part = other_parts[0] if other_parts else "B"
            else:
                assigned_part = ""

        container_part = assigned_part

        # Determine status dynamically based on physical count vs commercial invoice
        if is_blind_setup or (not qc_map and not boms_map and raw_rec_u is None and unidades_desc == 0):
            status = "PENDIENTE_CONTROL_CALIDAD"
            obs = "Montaje a ciegas desde SAP. En pausa: no se puede emitir cierre sin Control de Calidad."
            rec_u = 0.0
            falt = 0.0
            dan = 0.0
            sob = 0.0
            diff = 0.0
        else:
            diff = rec_u - ped41_u
            if container_part and container_part != "A" and rec_u == 0:
                status = f"PENDIENTE_PARTE_{container_part}"
                obs = f"Mercadería en Contenedor Parte {container_part} o pendiente de descarga."
            elif not container_part and rec_u == 0 and ped41_u > 0:
                status = "FALTANTE"
                obs = f"No recibido en muelle (faltante de {ped41_u:,.0f} unidades)."
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
        if mat_code in boms_map:
            for b in boms_map[mat_code]:
                c_code = b['child']
                c_ratio = b['ratio']
                if ped41_prk > 0:
                    c_theo = int(round(ped41_prk * c_ratio))
                elif curve > 0:
                    c_theo = int(round((ped41_u / curve) * c_ratio))
                else:
                    c_theo = 0

                c_desc = qc_map.get(c_code, {}).get('desc', f"{desc} (Talla {c_ratio})")
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
            observations=obs,
            bultos_descarga=bultos_desc,
            unidades_descarga=unidades_desc,
            bultos_packing_list=bultos_pl,
            unidades_packing_list=unidades_pl,
            empaque_str=empaque_str,
            bultos_str=pkg_as.raw_str or str(bultos_desc),
            um=um_str,
            bultos_enteros=b_enteros,
            unidades_por_bulto_entero=u_enteros,
            bultos_complementos=b_comp,
            unidades_por_bulto_comp=u_comp,
            unidades_por_bulto=u_per_b,
            dif_bultos=dif_bultos,
            is_blind_setup=is_blind_setup,
        )
        items.append(item)

    try:
        wb.close()
    except Exception:
        pass

    result_list = ItemList(items, metadata=meta)

    if return_meta:
        return result_list, meta
    return result_list
