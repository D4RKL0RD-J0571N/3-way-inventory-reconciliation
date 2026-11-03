"""
Reconciliation engine for 3-way matching and analytics.
"""
from typing import List, Dict, Any
from src.models import PrepackItem, ReconciliationSummary

def compute_summary(items: List[PrepackItem], import_id: str = "IMP-182", vendor_name: str = "CENTRE ZONE LIMITED", invoice_number: str = "4100002283") -> ReconciliationSummary:
    """
    Computes global reconciliation KPIs.
    """
    total_mat = len(items)
    part_a_mat = sum(1 for i in items if i.container_part == "A")
    part_b_mat = sum(1 for i in items if i.container_part == "B")
    
    total_ped_43 = sum(i.quantity_po_43_units for i in items)
    total_ped_41 = sum(i.quantity_inv_41_units for i in items)
    total_rec = sum(i.received_real_units for i in items)
    
    total_falt = sum(i.faltante for i in items)
    total_dan = sum(i.danado for i in items)
    total_sob = sum(i.sobrante for i in items)
    total_ncr = sum(i.ncr_amount for i in items)
    
    status = "LIMPIO" if (total_falt == 0 and total_dan == 0) else "CON_DISCREPANCIAS"
    
    return ReconciliationSummary(
        import_id=import_id,
        vendor_name=vendor_name,
        invoice_number=invoice_number,
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
        status=status
    )

def filter_items(
    items: List[PrepackItem],
    part_filter: str = "ALL", # "ALL", "A", "B"
    status_filter: str = "ALL",
    search_query: str = ""
) -> List[PrepackItem]:
    """
    Filters items by container part, status, and search keyword (searches parent or child SKUs).
    """
    filtered = items
    if part_filter in ["A", "B"]:
        filtered = [i for i in filtered if i.container_part == part_filter]
        
    if status_filter != "ALL":
        filtered = [i for i in filtered if i.status == status_filter]
        
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
                q in i.category.lower()
            )
            # Search children SKUs
            child_match = any(q in c.material_code.lower() or q in c.description.lower() for c in i.children)
            if parent_match or child_match:
                matched.append(i)
        filtered = matched
        
    return filtered
