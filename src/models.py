"""
Data models for 3-Way Inventory Reconciliation.
"""
from dataclasses import dataclass, field
from typing import List, Optional, Dict

@dataclass
class PackagingBreakdown:
    total_boxes: float = 0.0
    total_units: float = 0.0
    boxes_full: float = 0.0
    units_per_box_full: float = 0.0
    boxes_comp: float = 0.0
    units_per_box_comp: float = 0.0
    raw_str: str = ""
    is_compound: bool = False

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
    container_part: str = "A" # Dynamic part identifier: "A", "B", "C", "D", "E", "F"... or "" if single container
    sap_mov_101: float = 0.0
    sap_mov_641: float = 0.0
    children: List[ChildSKU] = field(default_factory=list)
    faltante: float = 0.0
    danado: float = 0.0
    sobrante: float = 0.0
    revuelto: float = 0.0
    status: str = "LIMPIO" # Dynamic status: "LIMPIO", "FALTANTE", "SOBRANTE", "DAÑADO", "REVUELTO", "PENDIENTE_PARTE_X"...
    ncr_amount: float = 0.0
    observations: str = ""
    # Explicit Bultos & Packaging Fields (Muelle vs Packing List)
    bultos_descarga: float = 0.0
    unidades_descarga: float = 0.0
    bultos_packing_list: float = 0.0
    unidades_packing_list: float = 0.0
    empaque_str: str = ""
    bultos_str: str = ""
    um: str = "PRK"
    bultos_enteros: float = 0.0
    unidades_por_bulto_entero: float = 0.0
    bultos_complementos: float = 0.0
    unidades_por_bulto_comp: float = 0.0
    unidades_por_bulto: float = 0.0
    dif_bultos: float = 0.0

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
    total_bultos_descarga: float = 0.0
    total_bultos_pl: float = 0.0
    dif_total_bultos: float = 0.0
    total_prk_41: float = 0.0
    total_prk_43: float = 0.0
