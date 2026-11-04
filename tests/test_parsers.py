"""
Unit and integration tests for dynamic parsers and multi-container resilience.
"""
import openpyxl
from pathlib import Path
from src.parsers import parse_quality_control_boms, parse_cierre_recepcion, extract_cierre_metadata
from src.engine import compute_summary
from src.config import resolve_local_import_files

def test_dynamic_header_detection_with_inserted_column(tmp_path):
    """
    Simulates warehouse operator inserting 'Tarima' at Col B in CONTROL DATA.
    Verifies that the dynamic parser detects Material, Description, and Totals accurately.
    """
    wb = openpyxl.Workbook()
    ws_bd = wb.active
    ws_bd.title = "BASE DE DATOS"
    ws_bd.cell(4, 5, "LISTA")
    ws_bd.cell(4, 9, "MATERIAL")
    ws_bd.cell(4, 10, "CURVA")
    ws_bd.cell(4, 11, "UM")
    
    # Parent Prepack
    ws_bd.cell(5, 5, "BOM1")
    ws_bd.cell(5, 9, "1000001")
    ws_bd.cell(5, 10, 6)
    ws_bd.cell(5, 11, "PRK")
    
    # Child SKU
    ws_bd.cell(6, 5, "BOM1")
    ws_bd.cell(6, 9, "1000001001")
    ws_bd.cell(6, 10, 2)
    ws_bd.cell(6, 11, "UN")
    
    ws_qc = wb.create_sheet(title="CONTROL DATA")
    # Shifted headers: Col 2 is Tarima!
    ws_qc.cell(6, 1, "SECCION")
    ws_qc.cell(6, 2, "Tarima")
    ws_qc.cell(6, 3, "CATEGORIA")
    ws_qc.cell(6, 4, "Material proveedor")
    ws_qc.cell(6, 5, "Material") # Col 5 shifted!
    ws_qc.cell(6, 12, "UM")
    ws_qc.cell(6, 13, "DESCRIPCION")
    ws_qc.cell(6, 34, "TOTAL DE BULTOS")
    ws_qc.cell(6, 35, "TOTAL DE UNIDADES")
    
    ws_qc.cell(7, 5, "1000001")
    ws_qc.cell(7, 12, "PRK")
    ws_qc.cell(7, 13, "CAMISA ALGODON")
    ws_qc.cell(7, 34, 10)
    ws_qc.cell(7, 35, 600)
    
    qc_file = tmp_path / "QC_test.xlsm"
    wb.save(qc_file)
    
    boms, details = parse_quality_control_boms(str(qc_file))
    
    assert "1000001" in boms
    assert len(boms["1000001"]) == 1
    assert boms["1000001"][0]["child"] == "1000001001"
    assert boms["1000001"][0]["ratio"] == 2
    
    assert "1000001" in details
    assert details["1000001"]["desc"] == "CAMISA ALGODON"
    assert details["1000001"]["bultos"] == 10
    assert details["1000001"]["unidades"] == 600

def test_metadata_extractor(tmp_path):
    """
    Verifies extraction of container ID, invoice number, vendor, and package counts.
    """
    wb = openpyxl.Workbook()
    ws_dg = wb.active
    ws_dg.title = "DATOS GENERALES"
    
    ws_dg.cell(7, 2, "Importación")
    ws_dg.cell(7, 3, "CONT-TEST-001")
    ws_dg.cell(9, 2, "Pedido de Nacionalización :")
    ws_dg.cell(9, 3, "4100099999")
    ws_dg.cell(10, 2, "Proveedor")
    ws_dg.cell(10, 3, "GLOBAL TEXTILES CORP")
    ws_dg.cell(13, 2, "Bultos Teóricos de esta parte")
    ws_dg.cell(13, 3, 500)
    ws_dg.cell(14, 2, "Bultos Recibidos de esta parte")
    ws_dg.cell(14, 3, 500)
    ws_dg.cell(15, 2, "Diferencias de bultos teóricos vrs recibidos")
    ws_dg.cell(15, 3, 0)
    
    meta = extract_cierre_metadata(wb)
    assert meta["import_id"] == "CONT-TEST-001"
    assert meta["invoice_number"] == "4100099999"
    assert meta["vendor_name"] == "GLOBAL TEXTILES CORP"
    assert meta["theor_packages"] == 500
    assert meta["received_packages"] == 500
    assert meta["difference_packages"] == 0

def test_synthetic_full_reconciliation_clean(tmp_path):
    """
    End-to-end reconciliation test using synthetic mock spreadsheets.
    """
    # 1. Create mock QC workbook
    qc_wb = openpyxl.Workbook()
    ws_cd = qc_wb.active
    ws_cd.title = "CONTROL DATA"
    ws_cd.append(["Material", "Descripción", "Bultos", "Total Unidades"])
    ws_cd.append(["1001001", "Mock Item 1", 20, 240])
    ws_cd.append(["1001002", "Mock Item 2", 30, 360])

    ws_bd = qc_wb.create_sheet(title="BASE DE DATOS")
    for _ in range(4):
        ws_bd.append([])
    # BOM definition
    ws_bd.append(["", "", "", "", "BOM-01", "1001001", 12, "PRK", "1001001", 12, "PRK"])
    ws_bd.append(["", "", "", "", "BOM-01", "1001001001", 6, "UN", "1001001001", 6, "UN"])
    ws_bd.append(["", "", "", "", "BOM-01", "1001001002", 6, "UN", "1001001002", 6, "UN"])
    
    qc_file = tmp_path / "QC_Mock.xlsx"
    qc_wb.save(qc_file)

    # 2. Create mock Cierre workbook
    cierre_wb = openpyxl.Workbook()
    ws_cr = cierre_wb.active
    ws_cr.title = "CIERRE RECEP"
    ws_cr.cell(1, 1, "CIERRE RECEPCION")
    
    headers = [
        "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U",
        "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN",
        "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF"
    ]
    # Place headers at row 12
    for col_idx, h in enumerate(headers, 1):
        ws_cr.cell(12, col_idx, h)
        
    # Row 13: Item 1001001
    ws_cr.cell(12, 5, "Material")
    ws_cr.cell(13, 1, "4300001")     # A: PO
    ws_cr.cell(13, 5, "1001001")     # E: Material
    ws_cr.cell(13, 4, "Mock Item 1") # D: Desc
    ws_cr.cell(13, 6, 240)           # F: Ped 43 UN
    ws_cr.cell(13, 7, 240)           # G: Ped 41 UN
    ws_cr.cell(13, 8, 240)           # H: Rec UN
    ws_cr.cell(13, 11, 5.0)          # K: Unit cost
    ws_cr.cell(13, 54, 12)           # BB: Curva
    ws_cr.cell(13, 55, "PRK")        # BC: UM
    ws_cr.cell(13, 57, 20)           # BE: Ped 41 PRK
    
    # DG sheet
    ws_dg = cierre_wb.create_sheet(title="DATOS GENERALES")
    ws_dg.cell(7, 2, "Importación")
    ws_dg.cell(7, 3, "CONT-SYNTHETIC-01")
    
    cierre_file = tmp_path / "Cierre_Mock.xlsx"
    cierre_wb.save(cierre_file)

    boms, qc_details = parse_quality_control_boms(str(qc_file))
    items, meta = parse_cierre_recepcion(str(cierre_file), boms, qc_details, return_meta=True)
    summary = compute_summary(items, metadata=meta)
    
    assert summary.import_id == "CONT-SYNTHETIC-01"
    assert summary.total_materials == 1
    assert summary.total_received_units == 240
    assert summary.status == "LIMPIO"


def test_dynamic_multi_part_c_d_extraction(tmp_path):
    """
    Verifies that imports with 3 or more parts (e.g. Part A, Part B, Part C, Part D)
    dynamically detect all parts from distribution sheets without hardcoded limitations.
    """
    qc_wb = openpyxl.Workbook()
    ws_cd = qc_wb.active
    ws_cd.title = "CONTROL DATA"
    ws_cd.cell(6, 5, "Material")
    ws_cd.cell(6, 13, "Descripción")
    ws_cd.cell(6, 34, "Total Bultos")
    ws_cd.cell(6, 35, "Total Unidades")

    items_data = [("A", "1001001"), ("B", "1001002"), ("C", "1001003"), ("D", "1001004")]
    for idx, (p, mat_id) in enumerate(items_data, start=7):
        ws_cd.cell(idx, 5, mat_id)
        ws_cd.cell(idx, 13, f"Item Part {p}")
        ws_cd.cell(idx, 34, 10)
        ws_cd.cell(idx, 35, 120)

    # Add distribution sheets for Parts A, B, C, D
    for part_letter, mat_id in items_data:
        ws_p = qc_wb.create_sheet(title=f"DIST. PARTE {part_letter}")
        ws_p.cell(1, 1, mat_id)

    qc_file = tmp_path / "QC_MultiPart.xlsx"
    qc_wb.save(qc_file)

    boms, qc_details = parse_quality_control_boms(str(qc_file))
    assert qc_details.get('_detected_parts') == ["A", "B", "C", "D"]
    assert qc_details["1001001"]["part"] == "A"
    assert qc_details["1001002"]["part"] == "B"
    assert qc_details["1001003"]["part"] == "C"
    assert qc_details["1001004"]["part"] == "D"


def test_single_container_fallback_never_fabricates_part_b(tmp_path):
    """
    Verifies that a single container import never assumes Part B for unreceived goods,
    correctly assigning FALTANTE status instead.
    """
    qc_wb = openpyxl.Workbook()
    ws_cd = qc_wb.active
    ws_cd.title = "CONTROL DATA"
    ws_cd.append(["Material", "Descripción", "Bultos", "Total Unidades"])
    ws_cd.append(["2001001", "Single Container Item 1", 10, 100])
    qc_file = tmp_path / "QC_Single.xlsx"
    qc_wb.save(qc_file)

    cierre_wb = openpyxl.Workbook()
    ws_cr = cierre_wb.active
    ws_cr.title = "CIERRE RECEP"
    ws_cr.cell(1, 1, "CIERRE RECEPCION")
    for c in range(1, 60):
        ws_cr.cell(12, c, f"Col_{c}")
    ws_cr.cell(12, 5, "Material")

    # Item 1: Received
    ws_cr.cell(13, 1, "PO-100")
    ws_cr.cell(13, 5, "2001001")
    ws_cr.cell(13, 4, "Item Received")
    ws_cr.cell(13, 6, 100)
    ws_cr.cell(13, 7, 100)
    ws_cr.cell(13, 8, 100)

    # Item 2: Unreceived (0 units received)
    ws_cr.cell(14, 1, "PO-100")
    ws_cr.cell(14, 5, "2001002")
    ws_cr.cell(14, 4, "Item Missing")
    ws_cr.cell(14, 6, 50)
    ws_cr.cell(14, 7, 50)
    ws_cr.cell(14, 8, 0) # 0 received

    cierre_file = tmp_path / "Cierre_Single.xlsx"
    cierre_wb.save(cierre_file)

    boms, qc_details = parse_quality_control_boms(str(qc_file))
    items = parse_cierre_recepcion(str(cierre_file), boms, qc_details)

    assert len(items) == 2
    item_clean = items[0]
    item_missing = items[1]

    assert item_clean.status == "LIMPIO"
    # Crucial test: item_missing must NOT be PENDIENTE_PARTE_B because it's a single container!
    assert item_missing.status == "FALTANTE"
    assert "PENDIENTE_PARTE_B" not in item_missing.status


def test_dynamic_filter_items_arbitrary_parts_and_statuses():
    """
    Verifies that filter_items dynamically filters by arbitrary container parts (C, D, etc.),
    dynamic statuses, and categories without any hardcoded restrictions.
    """
    from src.models import PrepackItem, ChildSKU
    from src.engine import filter_items

    item_a = PrepackItem(
        material_code="1001",
        vendor_material="V-01",
        description="Camisa Parte A",
        category="TEXTIL",
        order_po_43="4301",
        quantity_po_43_prk=10,
        quantity_po_43_units=120,
        quantity_inv_41_prk=10,
        quantity_inv_41_units=120,
        received_real_units=120,
        curve_total=12,
        unit_cost=10.0,
        container_part="A",
        status="LIMPIO"
    )

    item_c = PrepackItem(
        material_code="1002",
        vendor_material="V-02",
        description="Pantalón Parte C",
        category="MODA",
        order_po_43="4302",
        quantity_po_43_prk=5,
        quantity_po_43_units=60,
        quantity_inv_41_prk=5,
        quantity_inv_41_units=60,
        received_real_units=0,
        curve_total=12,
        unit_cost=15.0,
        container_part="C",
        status="PENDIENTE_PARTE_C"
    )

    item_d = PrepackItem(
        material_code="1003",
        vendor_material="V-03",
        description="Zapatos Parte D",
        category="CALZADO",
        order_po_43="4303",
        quantity_po_43_prk=8,
        quantity_po_43_units=48,
        quantity_inv_41_prk=8,
        quantity_inv_41_units=48,
        received_real_units=40,
        curve_total=6,
        unit_cost=25.0,
        container_part="D",
        status="DAÑADO",
        danado=8
    )

    items = [item_a, item_c, item_d]

    # Test filtering by arbitrary Part C
    res_c = filter_items(items, part_filter="C")
    assert len(res_c) == 1
    assert res_c[0].material_code == "1002"

    # Test filtering by dynamic status PENDIENTE_PARTE_C
    res_stat = filter_items(items, status_filter="PENDIENTE_PARTE_C")
    assert len(res_stat) == 1
    assert res_stat[0].material_code == "1002"

    # Test filtering by category CALZADO
    res_cat = filter_items(items, category_filter="CALZADO")
    assert len(res_cat) == 1
    assert res_cat[0].material_code == "1003"

    # Test search query matching container part
    res_q_part = filter_items(items, search_query="Parte C")
    assert len(res_q_part) == 1
    assert res_q_part[0].material_code == "1002"


def test_blind_setup_mode_without_qc(tmp_path):
    """
    Verifies the operational business rule:
    When an import has no Quality Control report, official closure cannot be performed.
    The system allows mounting the SAP macro blindly and leaving it in PAUSE mode.
    - All items are marked PENDIENTE_CONTROL_CALIDAD
    - No fake shortages or fake NCR claims are fabricated
    - Summary status is PAUSA_SIN_CONTROL_CALIDAD with is_blind_setup=True
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CIERRE RECEP"
    for c in range(1, 60):
        ws.cell(12, c, f"Col_{c}")

    ws.cell(12, 4, "Material")
    ws.cell(12, 5, "Estilo Prov")
    ws.cell(12, 11, "Texto breve")
    ws.cell(12, 13, "Cantidad facturada (UN)")
    ws.cell(12, 14, "Costo Unitario")
    ws.cell(12, 19, "Total Pedido 41 (PRK)")
    ws.cell(12, 23, "Curva")

    # Row 13: 100 units from SAP
    ws.cell(13, 4, "9990001")
    ws.cell(13, 5, "STYLE-BLIND-01")
    ws.cell(13, 11, "Camiseta Básica")
    ws.cell(13, 13, 100)
    ws.cell(13, 14, 15.5)
    ws.cell(13, 19, 10)
    ws.cell(13, 23, 10)

    # DATOS GENERALES sheet
    ws_dg = wb.create_sheet(title="DATOS GENERALES")
    ws_dg.cell(7, 2, "Importación")
    ws_dg.cell(7, 3, "IMP-BLIND-TEST")

    cierre_path = tmp_path / "Cierre_Blind.xlsx"
    wb.save(cierre_path)

    # Parse in blind setup mode
    items, meta = parse_cierre_recepcion(str(cierre_path), boms=None, qc_details=None, return_meta=True, is_blind_setup=True)
    summary = compute_summary(items, metadata=meta)

    assert len(items) == 1
    item = items[0]
    assert item.material_code == "9990001"
    assert item.status == "PENDIENTE_CONTROL_CALIDAD"
    assert item.is_blind_setup is True
    assert item.received_real_units == 0.0
    assert item.faltante == 0.0
    assert item.danado == 0.0
    assert item.ncr_amount == 0.0
    assert "En pausa: no se puede emitir cierre sin Control de Calidad" in item.observations

    assert summary.is_blind_setup is True
    assert summary.status == "PAUSA_SIN_CONTROL_CALIDAD"
    assert summary.total_received_units == 0.0
    assert summary.total_ncr_amount == 0.0

