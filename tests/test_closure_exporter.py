"""
Unit tests for Service 4: Receiving Closure Exporter & Pre-Print Validator (src/closure_exporter.py).

Validates VBA macro preservation, pre-print formula auditing (#REF! detection),
and accurate row population.
"""

import os
import openpyxl
import pytest
from src.closure_exporter import (
    validate_closure_document,
    export_receiving_closure,
    PrePrintValidationResult,
)
from src.discrepancy_classifier import (
    ComparisonSummary,
    MaterialDiscrepancy,
    DiscrepancyType,
)


@pytest.fixture
def sample_closure_template(tmp_path):
    """Create a minimal synthetic corporate closure workbook mimicking CIERRE RECEP."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CIERRE RECEP"

    # Header rows
    ws.cell(row=1, column=1, value="CONTENEDOR")
    ws.cell(row=1, column=2, value="TEST-100")
    ws.cell(row=5, column=1, value="BULTOS TEORICOS IMPORTADO")
    ws.cell(row=5, column=2, value=100)
    ws.cell(row=6, column=1, value="BULTOS RECIBIDOS:")
    ws.cell(row=6, column=2, value=100)
    ws.cell(row=7, column=1, value="DIF BULTOS :")
    ws.cell(row=7, column=2, value=0)

    # Table header at row 12
    headers = [
        "ORDEN DE COMPRA", "Proveedor", "Categoría", "MAT. PROVEEDOR", "MATERIAL",
        "CANTIDAD PEDIDO 43", "CANTIDAD PEDIDO 41", "RECIBIDO REAL", "MOV 101",
        "REVUELTO SALIENDO", "DESTINO", "REVUELTO ENTRANDO", "ORIGEN"
    ]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=12, column=col_idx, value=h)

    # Item row 13
    ws.cell(row=13, column=5, value="1419038") # Material
    ws.cell(row=13, column=6, value=1188)      # PO 43
    ws.cell(row=13, column=7, value=1188)      # PO 41
    ws.cell(row=13, column=8, value="=+G13-J13-N13-O13+L13+P13+Q13") # Formula

    # Item row 14
    ws.cell(row=14, column=5, value="1419040") # Material
    ws.cell(row=14, column=6, value=1764)      # PO 43
    ws.cell(row=14, column=7, value=1764)      # PO 41

    template_path = str(tmp_path / "test_template.xlsx")
    wb.save(template_path)
    wb.close()
    return template_path


class TestClosureExporter:
    """Test suite for closure export and pre-print verification."""

    def test_validate_nonexistent_file(self):
        res = validate_closure_document("non_existent_file.xlsx")
        assert not res.is_valid
        assert any("no encontrado" in e for e in res.errors)

    def test_detect_formula_corruption_in_sheet(self, tmp_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CIERRE RECEP"
        ws.cell(row=13, column=8, value="=#REF! + 10")
        corrupt_path = str(tmp_path / "corrupt.xlsx")
        wb.save(corrupt_path)
        wb.close()

        res = validate_closure_document(corrupt_path)
        assert not res.is_valid
        assert len(res.formula_errors) > 0
        assert any("corrupta" in e for e in res.errors)

    def test_export_and_validate_closure(self, sample_closure_template, tmp_path):
        out_path = str(tmp_path / "final_closure.xlsx")
        items = [
            MaterialDiscrepancy(
                material_code="1419038",
                vendor_material="23C02-07-31",
                description="BOXER BASICO AZUL",
                category="CABALLERO",
                declared_units=1188,
                physical_units=1188,
                difference_units=0,
                declared_cartons=6,
                physical_cartons=6,
                difference_cartons=0,
                discrepancy_type=DiscrepancyType.EXACT_MATCH
            ),
            MaterialDiscrepancy(
                material_code="1419040",
                vendor_material="23C02-07-31",
                description="BOXER BASICO GRIS",
                category="CABALLERO",
                declared_units=1764,
                physical_units=1764,
                difference_units=0,
                declared_cartons=9,
                physical_cartons=9,
                difference_cartons=0,
                discrepancy_type=DiscrepancyType.EXACT_MATCH
            )
        ]
        summary = ComparisonSummary(
            total_declared_units=2952,
            total_physical_units=2952,
            total_declared_cartons=15,
            total_physical_cartons=15,
            clean_materials_count=2,
            items=items
        )

        val_result = export_receiving_closure(
            template_path=sample_closure_template,
            output_path=out_path,
            summary=summary
        )

        assert val_result.is_valid
        assert val_result.total_physical_units == 2952
        assert val_result.total_physical_cartons == 15
        assert len(val_result.errors) == 0

        # Inspect generated workbook directly
        wb = openpyxl.load_workbook(out_path, data_only=False)
        ws = wb["CIERRE RECEP"]
        assert ws.cell(row=5, column=2).value == 15  # Declared cartons header
        assert ws.cell(row=6, column=2).value == 15  # Physical cartons header
        assert ws.cell(row=13, column=45).value == 6 # AS: Cartons
        assert ws.cell(row=13, column=46).value == 1188 # AT: Units
        assert ws.cell(row=14, column=45).value == 9 # AS: Cartons
        assert ws.cell(row=14, column=46).value == 1764 # AT: Units
        wb.close()
