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
