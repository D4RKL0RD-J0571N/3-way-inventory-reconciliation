"""
Robust parsers for Excel files (Cierre Recepción, Control de Calidad, Packing Lists).
Now equipped with DataSanitizer to handle dirty data, weird formats, and bad strings.
"""
import openpyxl
from typing import Dict, List, Tuple
from src.models import PrepackItem, ChildSKU
from src.sanitizer import (
    clean_string,
    clean_material_code,
    clean_number,
    clean_int
)

def parse_quality_control_boms(qc_path: str) -> Tuple[Dict[str, List[Dict]], Dict[str, Dict]]:
    """
    Extracts BOM definitions and child details from Control de Calidad sheet BASE DE DATOS & CONTROL DATA.
    Returns:
        (boms_by_parent, qc_details_by_material)
    """
    wb = openpyxl.load_workbook(qc_path, data_only=True)
    
    # 1. Parse BOM from BASE DE DATOS grouped by LISTA (BOM ID in Col E)
    ws_bd = wb['BASE DE DATOS']
    boms_by_parent: Dict[str, List[Dict]] = {}
    current_bom_id = None
    current_parent = None
    
    for r in range(5, ws_bd.max_row + 1):
        bom_id = ws_bd.cell(r, 5).value # Col E: LISTA (BOM ID)
        mat_val = ws_bd.cell(r, 6).value # Col F: Material
        ratio_val = ws_bd.cell(r, 7).value # Col G: CURVA
        
        bom_clean = clean_string(bom_id)
        mat_clean = clean_material_code(mat_val)
        
        if bom_clean and mat_clean:
            if bom_clean != current_bom_id:
                # First row of a new BOM group is the Parent Prepack
                current_bom_id = bom_clean
                current_parent = mat_clean
                boms_by_parent[current_parent] = []
            else:
                # Subsequent rows in the same BOM group are the Child SKUs
                ratio = clean_int(ratio_val, default=1)
                if current_parent and mat_clean != current_parent:
                    boms_by_parent[current_parent].append({
                        'child': mat_clean,
                        'ratio': ratio
                    })
                
    # 2. Parse descriptions and counts from CONTROL DATA
    ws_qc = wb['CONTROL DATA']
    qc_details: Dict[str, Dict] = {}
    for r in range(7, ws_qc.max_row + 1):
        m_clean = clean_material_code(ws_qc.cell(r, 4).value)
        if m_clean:
            qc_details[m_clean] = {
                'desc': clean_string(ws_qc.cell(r, 12).value),
                'um': clean_string(ws_qc.cell(r, 11).value).upper(),
                'qty': clean_number(ws_qc.cell(r, 5).value),
                'curve': clean_int(ws_qc.cell(r, 8).value)
            }
            
    return boms_by_parent, qc_details

def parse_cierre_recepcion(cierre_path: str, boms: Dict[str, List[Dict]], qc_details: Dict[str, Dict]) -> List[PrepackItem]:
    """
    Parses sheet CIERRE RECEP into structured PrepackItem objects using defensive data sanitization.
    """
    wb = openpyxl.load_workbook(cierre_path, data_only=True)
    ws = wb['CIERRE RECEP']
    
    items: List[PrepackItem] = []
    
    for r in range(13, ws.max_row + 1):
        mat_code = clean_material_code(ws.cell(r, 5).value) # Col E
        if not mat_code:
            continue
            
        vendor_mat = clean_string(ws.cell(r, 4).value)
        po_43 = clean_string(ws.cell(r, 1).value)
        category = clean_string(ws.cell(r, 3).value)
        desc = clean_string(ws.cell(r, 30).value) # Col AD
        
        ped43_u = clean_number(ws.cell(r, 6).value)
        ped41_u = clean_number(ws.cell(r, 7).value)
        rec_u = clean_number(ws.cell(r, 8).value)
        mov101 = clean_number(ws.cell(r, 9).value)
        mov641 = clean_number(ws.cell(r, 21).value)
        
        falt = clean_number(ws.cell(r, 14).value)
        dan = clean_number(ws.cell(r, 15).value)
        sob = clean_number(ws.cell(r, 16).value)
        
        unit_cost = clean_number(ws.cell(r, 32).value) # Col AF: Valor unitario PEDIDO NAC $
        curve = clean_int(ws.cell(r, 54).value, default=12) # Col BB: Curva
        
        ped43_prk = clean_number(ws.cell(r, 56).value) # Col BD
        ped41_prk = clean_number(ws.cell(r, 57).value) # Col BE
        
        # Determine container part (Part A if in BOMs of 182-A, otherwise Part B)
        container_part = "A" if mat_code in boms else "B"
        
        # Determine status
        diff = rec_u - ped41_u
        if container_part == "B" and rec_u == 0:
            status = "PENDIENTE_PARTE_B"
            obs = "Mercadería en Contenedor 182-B (CZC7933B) o pendiente de descarga."
        elif falt > 0:
            status = "FALTANTE"
            obs = f"Faltante de {falt:,.0f} unidades."
        elif dan > 0:
            status = "DAÑADO"
            obs = f"Prendas dañadas: {dan:,.0f} unidades."
        elif sob > 0:
            status = "SOBRANTE"
            obs = f"Sobrante de {sob:,.0f} unidades."
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
        
    return items
