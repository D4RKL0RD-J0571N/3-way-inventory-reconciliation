"""
Data models for 3-Way Inventory Reconciliation.
"""
from dataclasses import dataclass, field
from typing import List, Optional, Dict

@dataclass
class ChildSKU:
    material_code: str
    description: str
    ratio: int
    theoretical_units: int
    received_units: int
    difference: int = 0
    unit_cost: float = 0.0

@dataclass
class PrepackItem:
    material_code: str
    vendor_material: str
    description: str
    category: str
    order_po_43: str
    quantity_po_43_prk: float
    quantity_po_43_units: float
    quantity_inv_41_prk: float
    quantity_inv_41_units: float
    received_real_units: float
    curve_total: int
    unit_cost: float
    container_part: str = "A" # "A" or "B"
    sap_mov_101: float = 0.0
    sap_mov_641: float = 0.0
    children: List[ChildSKU] = field(default_factory=list)
    faltante: float = 0.0
    danado: float = 0.0
    sobrante: float = 0.0
    revuelto: float = 0.0
    status: str = "LIMPIO" # "LIMPIO", "FALTANTE", "SOBRANTE", "DAÑADO", "REVUELTO", "PENDIENTE_PARTE_B"
    ncr_amount: float = 0.0
    observations: str = ""

@dataclass
class ReconciliationSummary:
    import_id: str
    vendor_name: str
    invoice_number: str
    total_materials: int
    part_a_materials: int
    part_b_materials: int
    total_ped_43_units: float
    total_ped_41_units: float
    total_received_units: float
    total_faltante_units: float
    total_danado_units: float
    total_sobrante_units: float
    total_ncr_amount: float
    status: str
    theor_packages: float = 0.0
    received_packages: float = 0.0
    difference_packages: float = 0.0
    container_parts_breakdown: Dict[str, int] = field(default_factory=dict)
