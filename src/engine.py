"""
Reconciliation engine for 3-way matching and analytics.
"""
from typing import List, Dict, Any, Optional
from src.models import PrepackItem, ReconciliationSummary

def compute_summary(
    items: List[PrepackItem],
    import_id: Optional[str] = None,
    vendor_name: Optional[str] = None,
    invoice_number: Optional[str] = None,
    metadata: Optional[Dict[str, Any]] = None
) -> ReconciliationSummary:
    """
    Computes global reconciliation KPIs. Dynamically extracts metadata from items or arguments.
    """
    meta = metadata or getattr(items, 'metadata', {})
    
    final_import_id = import_id or meta.get('import_id') or "IMPORT-ACTIVA"
    final_vendor = vendor_name or meta.get('vendor_name') or "PROVEEDOR INTERNACIONAL"
    final_invoice = invoice_number or meta.get('invoice_number') or "PENDIENTE"
    
    theor_pkg = meta.get('theor_packages', 0.0)
    rec_pkg = meta.get('received_packages', 0.0)
    dif_pkg = meta.get('difference_packages', 0.0)

    total_mat = len(items)
    container_parts_breakdown = {}
    for i in items:
        p = str(i.container_part or "").strip().upper()
        if p:
            container_parts_breakdown[p] = container_parts_breakdown.get(p, 0) + 1

    part_a_mat = container_parts_breakdown.get("A", 0)
    part_b_mat = container_parts_breakdown.get("B", 0)
    
    total_ped_43 = sum(i.quantity_po_43_units for i in items)
    total_ped_41 = sum(i.quantity_inv_41_units for i in items)
    total_rec = sum(i.received_real_units for i in items)
    
    total_falt = sum(i.faltante for i in items)
    total_dan = sum(i.danado for i in items)
    total_sob = sum(i.sobrante for i in items)
    total_ncr = sum(i.ncr_amount for i in items)
    
    total_bultos_desc = sum(i.bultos_descarga for i in items)
    total_bultos_pl = sum(i.bultos_packing_list for i in items)
    dif_total_b = total_bultos_desc - total_bultos_pl
    total_prk_41 = sum(i.quantity_inv_41_prk for i in items)
    total_prk_43 = sum(i.quantity_po_43_prk for i in items)

    # Reconcile header packages with item-level totals if missing
    if theor_pkg == 0.0 and total_bultos_pl > 0:
        theor_pkg = total_bultos_pl
    if rec_pkg == 0.0 and total_bultos_desc > 0:
        rec_pkg = total_bultos_desc
    if dif_pkg == 0.0 and (rec_pkg != 0 or theor_pkg != 0):
        dif_pkg = rec_pkg - theor_pkg
    
    is_blind = any(getattr(i, 'is_blind_setup', False) or i.status == "PENDIENTE_CONTROL_CALIDAD" for i in items)
    if is_blind:
        status = "PAUSA_SIN_CONTROL_CALIDAD"
    else:
        status = "LIMPIO" if (total_falt == 0 and total_dan == 0) else "CON_DISCREPANCIAS"
    
    return ReconciliationSummary(
        import_id=final_import_id,
        vendor_name=final_vendor,
        invoice_number=final_invoice,
        total_materials=total_mat,
        part_a_materials=part_a_mat,
        part_b_materials=part_b_mat,
        total_ped_43_units=total_ped_43,
        total_ped_41_units=total_ped_41,
        total_received_units=total_rec,
        total_faltante_units=total_falt,
        total_danado_units=total_dan,
        total_sobrante_units=total_sob,
        total_ncr_amount=total_ncr,
        status=status,
        theor_packages=theor_pkg,
        received_packages=rec_pkg,
        difference_packages=dif_pkg,
        container_parts_breakdown=container_parts_breakdown,
        total_bultos_descarga=total_bultos_desc,
        total_bultos_pl=total_bultos_pl,
        dif_total_bultos=dif_total_b,
        total_prk_41=total_prk_41,
        total_prk_43=total_prk_43,
        is_blind_setup=is_blind
    )

def filter_items(
    items: List[PrepackItem],
    part_filter: str = "ALL", # "ALL", "A", "B", "C", "D", "E", "F"...
    status_filter: str = "ALL",
    category_filter: str = "ALL",
    search_query: str = ""
) -> List[PrepackItem]:
    """
    Filters items by container part, status, category, and search keyword (searches parent or child SKUs).
    Supports arbitrary container parts (Parts A through Z+) and dynamic statuses extracted from live data.
    """
    filtered = items
    if part_filter != "ALL":
        p_clean = part_filter.strip().upper()
        if p_clean in ["NONE", "SIN_PARTE", "SINGLE", "UNICO", "ÚNICO"]:
            filtered = [i for i in filtered if not str(i.container_part or "").strip()]
        else:
            filtered = [i for i in filtered if str(i.container_part or "").strip().upper() == p_clean]
        
    if status_filter != "ALL":
        filtered = [i for i in filtered if str(i.status or "").strip().upper() == status_filter.strip().upper()]

    if category_filter != "ALL":
        filtered = [i for i in filtered if str(i.category or "").strip().upper() == category_filter.strip().upper()]
        
    if search_query:
        q = search_query.lower().strip()
        matched = []
        for i in filtered:
            # Search parent attributes
            parent_match = (
                q in i.material_code.lower() or
                q in i.vendor_material.lower() or
                q in i.description.lower() or
                q in i.order_po_43.lower() or
                q in i.category.lower() or
                q in str(i.container_part).lower() or
                q in str(i.status).lower()
            )
            # Search children SKUs
            child_match = any(q in c.material_code.lower() or q in c.description.lower() for c in i.children)
            if parent_match or child_match:
                matched.append(i)
        filtered = matched
        
    return filtered
