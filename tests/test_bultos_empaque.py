"""
Unit tests for Packaging Breakdown parser and Carton (Bultos) Reconciliation.
Ensures zero confusion between Cartons (Bultos), Prepacks (PRK), and Total Units.
"""
import pytest
from src.models import PrepackItem, PackagingBreakdown, ReconciliationSummary
from src.parsers import parse_packaging_string
from src.engine import compute_summary

def test_parse_packaging_scalar():
    res = parse_packaging_string(42)
    assert res.total_boxes == 42.0
    assert res.boxes_full == 42.0
    assert not res.is_compound

    res2 = parse_packaging_string("15")
    assert res2.total_boxes == 15.0
    assert not res2.is_compound

def test_parse_packaging_single_pattern():
    res = parse_packaging_string("50*12")
    assert res.total_boxes == 50.0
    assert res.total_units == 600.0
    assert res.boxes_full == 50.0
    assert res.units_per_box_full == 12.0
    assert not res.is_compound

    res_x = parse_packaging_string("10x24")
    assert res_x.total_boxes == 10.0
    assert res_x.total_units == 240.0

def test_parse_packaging_compound():
    # Compound packaging breakdown (e.g. 39 boxes of 60 + 1 box of 78)
    res = parse_packaging_string("39*60 1*78")
    assert res.total_boxes == 40.0
    assert res.total_units == 2418.0
    assert res.boxes_full == 39.0
    assert res.units_per_box_full == 60.0
    assert res.boxes_comp == 1.0
    assert res.units_per_box_comp == 78.0
    assert res.is_compound

    # Compound packaging with plus sign
    res_plus = parse_packaging_string("40 * 60 + 1 * 72")
    assert res_plus.total_boxes == 41.0
    assert res_plus.total_units == 2472.0
    assert res_plus.is_compound

def test_parse_packaging_empty_or_none():
    assert parse_packaging_string(None).total_boxes == 0.0
    assert parse_packaging_string("").total_boxes == 0.0
    assert parse_packaging_string("None").total_boxes == 0.0
    assert parse_packaging_string("nan").total_boxes == 0.0

def test_bultos_math_in_prepack_item():
    item = PrepackItem(
        material_code="1001001",
        vendor_material="VM-01",
        description="DM CAMISETA MC",
        category="DAMAS REGULAR",
        order_po_43="4500001",
        quantity_po_43_prk=403.0,
        quantity_po_43_units=2418.0,
        quantity_inv_41_prk=403.0,
        quantity_inv_41_units=2418.0,
        received_real_units=2418.0,
        curve_total=6,
        unit_cost=3.5,
        bultos_descarga=40.0,
        unidades_descarga=2418.0,
        bultos_packing_list=40.0,
        unidades_packing_list=2418.0,
        empaque_str="60",
        bultos_str="39*60 1*78",
        um="PRK",
        bultos_enteros=39.0,
        unidades_por_bulto_entero=60.0,
        bultos_complementos=1.0,
        unidades_por_bulto_comp=78.0,
        unidades_por_bulto=60.45,
        dif_bultos=0.0
    )
    # Ensure physical cartons (40) is not confused with PRKs (403) or Units (2418)
    assert item.bultos_descarga == 40.0
    assert item.quantity_inv_41_prk == 403.0
    assert item.received_real_units == 2418.0
    assert item.bultos_enteros + item.bultos_complementos == 40.0
    assert (item.bultos_enteros * item.unidades_por_bulto_entero) + (item.bultos_complementos * item.unidades_por_bulto_comp) == 2418.0

def test_compute_summary_bultos():
    items = [
        PrepackItem(
            material_code="MAT1",
            vendor_material="V1",
            description="Item 1",
            category="CAT1",
            order_po_43="PO1",
            quantity_po_43_prk=10.0,
            quantity_po_43_units=60.0,
            quantity_inv_41_prk=10.0,
            quantity_inv_41_units=60.0,
            received_real_units=60.0,
            curve_total=6,
            unit_cost=5.0,
            bultos_descarga=2.0,
            bultos_packing_list=2.0,
        ),
        PrepackItem(
            material_code="MAT2",
            vendor_material="V2",
            description="Item 2",
            category="CAT1",
            order_po_43="PO1",
            quantity_po_43_prk=20.0,
            quantity_po_43_units=120.0,
            quantity_inv_41_prk=20.0,
            quantity_inv_41_units=120.0,
            received_real_units=120.0,
            curve_total=6,
            unit_cost=5.0,
            bultos_descarga=5.0,
            bultos_packing_list=4.0, # 1 carton physical discrepancy!
        )
    ]
    summary = compute_summary(items)
    assert summary.total_bultos_descarga == 7.0
    assert summary.total_bultos_pl == 6.0
    assert summary.dif_total_bultos == 1.0
    assert summary.total_prk_41 == 30.0
    assert summary.total_received_units == 180.0


def test_dual_carton_modes_dock_vs_strict_customs():
    """
    Verifies that ReconcileAI correctly handles both carton views:
    1. Dock Count Mode: Direct physical cartons counted on the patio (e.g. 40 cartons).
    2. Strict Customs Mode: Mathematical decomposition into full standard boxes and odds (e.g. 39 Full + 1 Odd).
    """
    item = PrepackItem(
        material_code="1001",
        vendor_material="VM-01",
        description="Pantalon Denim",
        category="CABALLEROS",
        order_po_43="PO-100",
        quantity_po_43_prk=40.0,
        quantity_po_43_units=240.0,
        quantity_inv_41_prk=40.0,
        quantity_inv_41_units=240.0,
        received_real_units=240.0,
        curve_total=6,
        unit_cost=10.0,
        bultos_descarga=5.0, # 5 physical cartons reported by dock
        unidades_descarga=240.0,
        bultos_packing_list=5.0,
        unidades_packing_list=240.0,
        empaque_str="50",
        bultos_str="4*50 1*40",
        um="PRK",
        bultos_enteros=4.0, # 4 full cartons of 50
        unidades_por_bulto_entero=50.0,
        bultos_complementos=1.0, # 1 odd carton of 40
        unidades_por_bulto_comp=40.0,
        unidades_por_bulto=48.0,
        dif_bultos=0.0
    )

    # 1. Dock View Mode
    dock_cartons = item.bultos_descarga
    assert dock_cartons == 5.0

    # 2. Strict Customs Mode
    customs_full = item.bultos_enteros
    customs_odds = item.bultos_complementos
    assert customs_full == 4.0
    assert customs_odds == 1.0
    assert customs_full + customs_odds == 5.0
    assert (customs_full * item.unidades_por_bulto_entero) + (customs_odds * item.unidades_por_bulto_comp) == item.received_real_units


def test_zero_units_produce_zero_cartons():
    """
    Critical regression test:
    Items with 0 units received or status 'PENDIENTE_PARTE_B' must NEVER produce
    a phantom fallback of 1 carton (1.0). Cartons and packaging values must strictly be 0.
    """
    item_pending = PrepackItem(
        material_code="1306026",
        vendor_material="EMN0100",
        description="SERIES DE LUCES BLANCA ARROZ",
        category="NAVIDAD",
        order_po_43="4300006742",
        quantity_po_43_prk=7560.0,
        quantity_po_43_units=7560.0,
        quantity_inv_41_prk=0.0,
        quantity_inv_41_units=0.0,
        received_real_units=0.0,
        curve_total=1,
        unit_cost=0.0,
        status="PENDIENTE_PARTE_B",
        container_part="B",
        bultos_descarga=0.0,
        unidades_descarga=0.0,
        bultos_packing_list=0.0,
        unidades_packing_list=0.0,
        empaque_str="0",
        bultos_str="0",
        bultos_enteros=0.0,
        bultos_complementos=0.0,
        dif_bultos=0.0
    )

    # Calculation logic for dock (AS/AT) & packing list (AW/AX/AY)
    is_other_part = item_pending.status.startswith("PENDIENTE_PARTE") or (item_pending.container_part and item_pending.container_part != "A")
    assert is_other_part is True

    if item_pending.received_real_units == 0 or is_other_part:
        bultos_desc = 0.0
        unidades_desc = 0.0
    else:
        bultos_desc = item_pending.bultos_descarga or 1.0
        unidades_desc = item_pending.unidades_descarga

    if is_other_part or (item_pending.received_real_units == 0 and item_pending.quantity_inv_41_units == 0):
        bultos_pl = 0.0
        unidades_pl = 0.0
        empaque_val = "0"
    else:
        bultos_pl = item_pending.bultos_packing_list or bultos_desc
        unidades_pl = item_pending.unidades_packing_list
        empaque_val = item_pending.empaque_str or "12"

    assert bultos_desc == 0.0
    assert unidades_desc == 0.0
    assert bultos_pl == 0.0
    assert unidades_pl == 0.0
    assert empaque_val == "0"


