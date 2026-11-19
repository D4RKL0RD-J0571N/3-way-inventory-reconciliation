"""
Unit and integration tests for Service 2: QC Spreadsheet Ingestion Reader (src/qc_reader.py).

Validates header mapping, single-pass streaming, metadata extraction,
and real container QC parsing.
"""

import os
import openpyxl
import pytest
from src.qc_reader import read_qc_inspection_file, QCInspectionReport
from src.quantity_engine import (
    FLAG_QC_COMPLEMENT_WITHOUT_CARTON_COUNT,
    FLAG_SHARED_CARTON_PRIMARY,
    FLAG_SHARED_CARTON_SECONDARY,
    FLAG_VARIANTE_SUELTA_ROW,
)


@pytest.fixture
def sample_qc_workbook(tmp_path):
    """Create a minimal synthetic QC workbook mimicking the real production structure."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CONTROL DATA"

    # Metadata rows
    ws.append(["", "", "", "", "", "CONTENEDOR", "TEST-999"])
    ws.append(["", "", "", "", "", "RESPONSABLE:", "Analista Pruebas"])
    ws.append(["", "", "", "", "", "NACIONALIZACIÓN:", "4100009999"])
    ws.append([])
    ws.append(["SECCIÓN DE REPARTO", "", "", "", "", "", "", "RECIBO TOTAL DE BULTOS"])

    # Header row (row 6)
    headers = [
        "NUMERO DE TARIMA", "CATEGORIA", "Material proveedor", "Material",
        "CANTIDAD DE UNIDADES", "CONFIRMACION MATERIAL SAP", "Diferencia en Material SAP",
        "CURVA SISTEMA", "CONFIRMAR CURVA", "CURVA FISICA", "UM", "DESCRIPCIÓN",
        "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA",
        "BULTOS ENTEROS", "UNIDADES POR BULTOS COMPLETOS", "BULTOS COMPLEMENTOS",
        "UNIDADES BULTOS COMPLEMENTOS", "AF", "DIFERENCIA EN UNIDADES", "TOTAL DE BULTOS",
        "TOTAL DE UNIDADES", "TOTAL PRK", "CURVA", "AL", "OBSERVACIONES"
    ]
    ws.append(headers)

    # Row 7: Standard item
    ws.append([
        "1", "CABALLERO", "VP-01", "1001", 600, "C", "", 12, "C", 12, "PRK", "BOXER REGULAR",
        "", "", "", "", "", "", "", "", "", "",
        10, 60, "", "", "", "", 10, 600, 50, 12, "", ""
    ])

    # Row 8: Row 37 Bug: Complement units present but complement cartons blank
    ws.append([
        "2", "CABALLERO", "VP-02", "1002", 1200, "C", "", 12, "C", 12, "PRK", "BOXER LARGO NEGRO",
        "", "", "", "", "", "", "", "", "", "",
        6, 192, "", 48, "", -48, 6, 1200, 100, 12, "", ""
    ])

    # Row 9: Primary shared carton
    ws.append([
        "3", "CABALLERO", "VP-03", "1003", 300, "C", "", 1, "C", 1, "UN", "CORBATA AZUL",
        "", "", "", "", "", "", "", "", "", "",
        1, 300, "", "", "", "", 1, 300, "", 1, "", "DOS MATERIALES POR BULTO"
    ])

    # Row 10: Secondary shared carton
    ws.append([
        "4", "CABALLERO", "VP-04", "1004", 300, "C", "", 1, "C", 1, "UN", "CORBATA GRIS",
        "", "", "", "", "", "", "", "", "", "",
        "", 300, "", "", "", "", 0, 300, "", 1, "", "DOS MATERIALES POR BULTO"
    ])

    # Row 11: Variante suelta
    ws.append([
        "VARIANTE SUELTA", "CABALLERO", "VP-05", "1005", 10, "C", "", 1, "C", 1, "UN", "VARIANTE TALLA S",
        "", "", "", "", "", "", "", "", "", "",
        "", "", "", "", "", "", 0, 10, "", 1, "", ""
    ])

    file_path = str(tmp_path / "test_qc.xlsx")
    wb.save(file_path)
    wb.close()
    return file_path


class TestQCReader:
    """Test suite for QC workbook ingestion."""

    def test_read_synthetic_qc_file(self, sample_qc_workbook):
        report = read_qc_inspection_file(sample_qc_workbook)

        assert report.container_id == "TEST-999"
        assert report.metadata.get("responsible") == "Analista Pruebas"
        assert report.metadata.get("nationalization_po") == "4100009999"
        assert len(report.items) == 5

        # Item 1: Standard
        it1 = report.items[0]
        assert it1.material_code == "1001"
        assert it1.calculation_result.calculated_units == 600
        assert it1.calculation_result.calculated_cartons == 10

        # Item 2: Rescued complement units
        it2 = report.items[1]
        assert it2.material_code == "1002"
        assert it2.calculation_result.calculated_units == 1200
        assert it2.calculation_result.calculated_cartons == 6
        assert FLAG_QC_COMPLEMENT_WITHOUT_CARTON_COUNT in it2.calculation_result.flags

        # Item 3 & 4: Shared carton pair
        it3, it4 = report.items[2], report.items[3]
        assert it3.calculation_result.calculated_cartons == 1
        assert it4.calculation_result.calculated_cartons == 0
        assert it3.calculation_result.calculated_units == 300
        assert it4.calculation_result.calculated_units == 300
        assert FLAG_SHARED_CARTON_PRIMARY in it3.calculation_result.flags
        assert FLAG_SHARED_CARTON_SECONDARY in it4.calculation_result.flags

        # Item 5: Variante suelta
        it5 = report.items[4]
        assert it5.calculation_result.calculated_units == 10
        assert it5.calculation_result.calculated_cartons == 0
        assert FLAG_VARIANTE_SUELTA_ROW in it5.calculation_result.flags

        # Aggregate checks: 10 + 6 + 1 + 0 + 0 = 17 cartons
        assert report.total_calculated_cartons == 17
        # 600 + 1200 + 300 + 300 + 10 = 2410 units
        assert report.total_calculated_units == 2410

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            read_qc_inspection_file("non_existent_file.xlsm")

    def test_real_container_imp_103_integration(self):
        """Integration test with real production file if present on disk."""
        real_path = r"D:\Development\Software Development\3 Way Inventory Reconciliation\data\IMP-103\Informacion General\imp 103 control calidad corregirdo.xlsm"
        if not os.path.exists(real_path):
            pytest.skip("Real data directory not available in this environment")

        report = read_qc_inspection_file(real_path)
        assert report.container_id == "103"
        assert report.total_calculated_cartons == 1038
        assert report.total_reported_cartons == 1038
        assert len(report.items) == 267
        assert len(report.flagged_items) > 0
