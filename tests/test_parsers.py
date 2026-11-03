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
    ws_dg.cell(7, 3, "IMP-999")
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
    assert meta["import_id"] == "IMP-999"
    assert meta["invoice_number"] == "4100099999"
    assert meta["vendor_name"] == "GLOBAL TEXTILES CORP"
    assert meta["theor_packages"] == 500
    assert meta["received_packages"] == 500
    assert meta["difference_packages"] == 0

def test_local_imp182_regression():
    """
    Regression test against local IMP-182 files if present on machine.
    """
    qc_path, cierre_path = resolve_local_import_files("IMP-182")
    if not qc_path or not cierre_path:
        return
    
    boms, qc_details = parse_quality_control_boms(str(qc_path))
    items, meta = parse_cierre_recepcion(str(cierre_path), boms, qc_details, return_meta=True)
    summary = compute_summary(items, metadata=meta)
    
    assert summary.import_id == "IMP-182"
    assert summary.total_materials == 136
    assert round(summary.total_ped_41_units) == 51926
    assert round(summary.total_received_units) == 51926
    assert summary.status == "LIMPIO"

def test_local_imp174_integration():
    """
    Integration test against local IMP-174 files if present on machine.
    """
    qc_path, cierre_path = resolve_local_import_files("IMP-174")
    if not qc_path or not cierre_path:
        return
    
    boms, qc_details = parse_quality_control_boms(str(qc_path))
    items, meta = parse_cierre_recepcion(str(cierre_path), boms, qc_details, return_meta=True)
    summary = compute_summary(items, metadata=meta)
    
    assert summary.import_id == "IMP-174"
    assert summary.total_materials == 100
    assert summary.total_ped_41_units == 61607
    assert len(boms) == 44 # 44 Prepacks
