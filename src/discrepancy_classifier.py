"""
Service 3: Material-Level Comparison & Discrepancy Classifier (Comparador de Materiales y Clasificador de Discrepancias).

Performs strict, deterministic comparison between Supplier Declarations (Packing List/PO)
and Physical Receiving QC counts at the individual material code level.
Enforces the Cardinal Discrepancy Rule: differences between distinct material codes
must NEVER offset each other.
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, List, Optional, Set, Tuple


class DiscrepancyType(str, Enum):
    """Classification categories for dock receiving discrepancies."""
    EXACT_MATCH = "EXACT_MATCH"
    SHORTAGE = "SHORTAGE"
    OVERAGE = "OVERAGE"
    REVUELTO_CANDIDATE = "REVUELTO_CANDIDATE"
    CARTON_MISMATCH_ONLY = "CARTON_MISMATCH_ONLY"
    DAMAGE_ONLY = "DAMAGE_ONLY"
    UNEXPECTED_ITEM = "UNEXPECTED_ITEM"
    MISSING_ITEM = "MISSING_ITEM"


@dataclass(frozen=True)
class DeclaredItem:
    """Declared merchandise item from Packing List or Purchase Order."""
    material_code: str
    vendor_material: Optional[str] = None
    description: Optional[str] = None
    category: Optional[str] = None
    declared_units: int = 0
    declared_cartons: int = 0
    uom: Optional[str] = "UN"
    unit_cost: float = 0.0


@dataclass(frozen=True)
class PhysicalItem:
    """Physically received and inspected merchandise item from dock QC."""
    material_code: str
    vendor_material: Optional[str] = None
    description: Optional[str] = None
    category: Optional[str] = None
    physical_units: int = 0
    physical_cartons: int = 0
    damaged_units: int = 0
    uom: Optional[str] = "UN"
    flags: List[str] = field(default_factory=list)


@dataclass
class MaterialDiscrepancy:
    """Individual material-level comparison result and classification."""
    material_code: str
    vendor_material: str
    description: str
    category: str
    declared_units: int
    physical_units: int
    difference_units: int
    declared_cartons: int
    physical_cartons: int
    difference_cartons: int
    damaged_units: int = 0
    discrepancy_type: DiscrepancyType = DiscrepancyType.EXACT_MATCH
    revuelto_partner: Optional[str] = None
    flags: List[str] = field(default_factory=list)
    explanation: str = ""

    @property
    def is_zero_total(self) -> bool:
        return (
            self.declared_units == 0 and
            self.physical_units == 0 and
            self.declared_cartons == 0 and
            self.physical_cartons == 0 and
            self.damaged_units == 0
        )


@dataclass
class ComparisonSummary:
    """High-level summary of receiving comparison across all items."""
    total_declared_units: int = 0
    total_physical_units: int = 0
    net_unit_difference: int = 0
    total_declared_cartons: int = 0
    total_physical_cartons: int = 0
    net_carton_difference: int = 0
    total_shortage_units: int = 0
    total_overage_units: int = 0
    total_damaged_units: int = 0
    clean_materials_count: int = 0
    discrepant_materials_count: int = 0
    revuelto_pairs_count: int = 0
    items: List[MaterialDiscrepancy] = field(default_factory=list)

    @property
    def active_styles_count(self) -> int:
        return sum(1 for it in self.items if not it.is_zero_total)

    @property
    def zero_total_materials_count(self) -> int:
        return sum(1 for it in self.items if it.is_zero_total)


def _determine_discrepancy_type(
    diff_u: int,
    diff_c: int,
    dam: int,
    decl_u: int,
    phys_u: int
) -> DiscrepancyType:
    """Classify the nature of the difference for a single material."""
    if decl_u > 0 and phys_u == 0:
        return DiscrepancyType.MISSING_ITEM
    if decl_u == 0 and phys_u > 0:
        return DiscrepancyType.UNEXPECTED_ITEM
    if diff_u < 0:
        return DiscrepancyType.SHORTAGE
    if diff_u > 0:
        return DiscrepancyType.OVERAGE
    if diff_u == 0 and dam > 0:
        return DiscrepancyType.DAMAGE_ONLY
    if diff_u == 0 and diff_c != 0:
        return DiscrepancyType.CARTON_MISMATCH_ONLY
    return DiscrepancyType.EXACT_MATCH


def _build_explanation(
    disc_type: DiscrepancyType,
    diff_u: int,
    diff_c: int,
    dam: int
) -> str:
    """Generate concise human-readable explanation in Spanish."""
    if disc_type == DiscrepancyType.EXACT_MATCH:
        return "Coincidencia exacta entre declaración física y documental."
    if disc_type == DiscrepancyType.MISSING_ITEM:
        return "Material declarado completamente ausente en la recepción física."
    if disc_type == DiscrepancyType.UNEXPECTED_ITEM:
        return f"Material no declarado físicamente recibido (+{diff_u} u)."
    if disc_type == DiscrepancyType.SHORTAGE:
        return f"Faltante físico de {abs(diff_u)} unidades."
    if disc_type == DiscrepancyType.OVERAGE:
        return f"Sobrante físico de +{diff_u} unidades."
    if disc_type == DiscrepancyType.CARTON_MISMATCH_ONLY:
        return f"Unidades exactas pero diferencia en bultos ({diff_c:+d} bultos)."
    if disc_type == DiscrepancyType.DAMAGE_ONLY:
        return f"Unidades completas pero {dam} unidades físicamente dañadas."
    return ""


def _compare_single_material(
    code: str,
    decl: Optional[DeclaredItem],
    phys: Optional[PhysicalItem]
) -> MaterialDiscrepancy:
    """Compare declared vs physical values for a single material code."""
    decl_u = decl.declared_units if decl else 0
    decl_c = decl.declared_cartons if decl else 0
    phys_u = phys.physical_units if phys else 0
    phys_c = phys.physical_cartons if phys else 0
    dam_u = phys.damaged_units if phys else 0

    diff_u = phys_u - decl_u
    diff_c = phys_c - decl_c

    v_mat = (phys.vendor_material if phys else None) or (decl.vendor_material if decl else "") or ""
    desc = (phys.description if phys else None) or (decl.description if decl else "") or ""
    cat = (phys.category if phys else None) or (decl.category if decl else "") or ""
    flags = list(phys.flags) if phys else []

    disc_type = _determine_discrepancy_type(diff_u, diff_c, dam_u, decl_u, phys_u)
    exp = _build_explanation(disc_type, diff_u, diff_c, dam_u)

    return MaterialDiscrepancy(
        material_code=code,
        vendor_material=v_mat,
        description=desc,
        category=cat,
        declared_units=decl_u,
        physical_units=phys_u,
        difference_units=diff_u,
        declared_cartons=decl_c,
        physical_cartons=phys_c,
        difference_cartons=diff_c,
        damaged_units=dam_u,
        discrepancy_type=disc_type,
        flags=flags,
        explanation=exp
    )


def _detect_revuelto_pairs(items: List[MaterialDiscrepancy]) -> int:
    """
    Identify potential material swaps (revueltos) where material A has -X
    and material B has +X units. Never offsets quantities; merely links them.
    """
    shortages = [it for it in items if it.difference_units < 0]
    overages = [it for it in items if it.difference_units > 0]
    paired_shortages: Set[str] = set()
    paired_overages: Set[str] = set()
    pairs_count = 0

    for s in shortages:
        if s.material_code in paired_shortages:
            continue
        target_qty = abs(s.difference_units)
        # Search overages with exact matching opposite quantity
        for o in overages:
            if o.material_code in paired_overages:
                continue
            if o.difference_units == target_qty:
                # Candidate match: link both items
                s.discrepancy_type = DiscrepancyType.REVUELTO_CANDIDATE
                o.discrepancy_type = DiscrepancyType.REVUELTO_CANDIDATE
                s.revuelto_partner = o.material_code
                o.revuelto_partner = s.material_code
                s.explanation += f" [Candidato a revuelto con {o.material_code} (+{o.difference_units} u)]"
                o.explanation += f" [Candidato a revuelto con {s.material_code} ({s.difference_units} u)]"
                paired_shortages.add(s.material_code)
                paired_overages.add(o.material_code)
                pairs_count += 1
                break
    return pairs_count


def compare_merchandise(
    declared_items: List[DeclaredItem],
    physical_items: List[PhysicalItem]
) -> ComparisonSummary:
    """
    Main entry point for comparing declared vs physical merchandise.
    Compares every distinct material code and detects candidate revueltos.
    """
    decl_map: Dict[str, DeclaredItem] = {d.material_code: d for d in declared_items}
    phys_map: Dict[str, PhysicalItem] = {p.material_code: p for p in physical_items}
    all_codes = sorted(set(decl_map.keys()) | set(phys_map.keys()))

    discrepancies: List[MaterialDiscrepancy] = []
    for code in all_codes:
        discrepancies.append(
            _compare_single_material(code, decl_map.get(code), phys_map.get(code))
        )

    revuelto_count = _detect_revuelto_pairs(discrepancies)

    tot_decl_u = sum(d.declared_units for d in declared_items)
    tot_phys_u = sum(p.physical_units for p in physical_items)
    tot_decl_c = sum(d.declared_cartons for d in declared_items)
    tot_phys_c = sum(p.physical_cartons for p in physical_items)

    shortage_u = sum(abs(it.difference_units) for it in discrepancies if it.difference_units < 0)
    overage_u = sum(it.difference_units for it in discrepancies if it.difference_units > 0)
    damaged_u = sum(it.damaged_units for it in discrepancies)

    clean_count = sum(1 for it in discrepancies if it.discrepancy_type == DiscrepancyType.EXACT_MATCH and not it.is_zero_total)
    active_count = sum(1 for it in discrepancies if not it.is_zero_total)
    disc_count = active_count - clean_count

    return ComparisonSummary(
        total_declared_units=tot_decl_u,
        total_physical_units=tot_phys_u,
        net_unit_difference=tot_phys_u - tot_decl_u,
        total_declared_cartons=tot_decl_c,
        total_physical_cartons=tot_phys_c,
        net_carton_difference=tot_phys_c - tot_decl_c,
        total_shortage_units=shortage_u,
        total_overage_units=overage_u,
        total_damaged_units=damaged_u,
        clean_materials_count=clean_count,
        discrepant_materials_count=disc_count,
        revuelto_pairs_count=revuelto_count,
        items=discrepancies
    )
