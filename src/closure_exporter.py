"""
Service 4: Receiving Closure Exporter & Pre-Print Validator (Generador de Cierre de Recepción y Validador Pre-Impresión).

Populates the corporate receiving closure workbook ('CIERRE RECEP' and 'Bultos')
while strictly preserving existing VBA macros, styles, and formulas.
Performs pre-print integrity checks to catch formula bugs (#REF!, mismatched totals)
before documents are printed or submitted.
"""

from dataclasses import dataclass, field
import os
import shutil
from typing import Any, Dict, List, Optional, Tuple
import openpyxl

from src.discrepancy_classifier import ComparisonSummary, MaterialDiscrepancy, DiscrepancyType
from src.qc_reader import QCInspectionReport


@dataclass
class PrePrintValidationResult:
    """Pre-print validation outcome checking document integrity before finalization."""
    is_valid: bool
    total_physical_units: int
    total_physical_cartons: int
    discrepancy_count: int
    unresolved_revueltos: int
    formula_errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)


def _safe_cell_val(val: Any) -> str:
    """Return clean string for formula and error detection."""
    return str(val).strip().upper() if val is not None else ""


def _check_formula_errors_in_sheet(ws: Any) -> List[str]:
    """Scan sheet cells for Excel calculation errors (#REF!, #VALUE!, #N/A, #DIV/0!)."""
    error_tokens = ("#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?")
    found_errors: List[str] = []
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        for c_idx, cell in enumerate(row, start=1):
            if cell is not None:
                s = str(cell).upper()
                if any(err in s for err in error_tokens):
                    found_errors.append(f"Celda R{r_idx}C{c_idx}: {cell}")
    return found_errors


def validate_closure_document(
    file_path: str,
    expected_units: Optional[int] = None,
    expected_cartons: Optional[int] = None
) -> PrePrintValidationResult:
    """
    Audit an Excel closure document before printing or submission.
    Verifies that formulas are intact and totals match expectations.
    """
    if not os.path.exists(file_path):
        return PrePrintValidationResult(
            is_valid=False,
            total_physical_units=0,
            total_physical_cartons=0,
            discrepancy_count=0,
            unresolved_revueltos=0,
            errors=[f"Archivo no encontrado: {file_path}"]
        )

    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=False)
    try:
        if "CIERRE RECEP" not in wb.sheetnames:
            return PrePrintValidationResult(
                is_valid=False,
                total_physical_units=0,
                total_physical_cartons=0,
                discrepancy_count=0,
                unresolved_revueltos=0,
                errors=["El libro no contiene la hoja requerida 'CIERRE RECEP'."]
            )

        ws = wb["CIERRE RECEP"]
        formula_errors = _check_formula_errors_in_sheet(ws)
        warnings: List[str] = []
        errors: List[str] = []

        if formula_errors:
            errors.extend(f"Fórmula corrupta detectada: {e}" for e in formula_errors[:5])

        # Read summary row totals if present (Row 5 & 6)
        r5_bultos = ws.cell(row=5, column=2).value
        r6_bultos = ws.cell(row=6, column=2).value

        # Calculate item sums in column AT (unidades) and AS (bultos)
        tot_u = 0
        tot_c = 0
        for row in ws.iter_rows(min_row=13, max_col=47, values_only=True):
            if len(row) >= 46:
                u_val = row[45] # AT (col 46, 0-indexed 45)
                c_val = row[44] # AS (col 45, 0-indexed 44)
                if isinstance(u_val, (int, float)):
                    tot_u += int(round(u_val))
                if isinstance(c_val, (int, float)):
                    tot_c += int(round(c_val))

        if expected_units is not None and tot_u != expected_units:
            warnings.append(
                f"Suma de unidades físicas en columna AT ({tot_u}) difiere de lo esperado ({expected_units})."
            )

        if expected_cartons is not None and tot_c != expected_cartons:
            warnings.append(
                f"Suma de bultos físicos en columna AS ({tot_c}) difiere de lo esperado ({expected_cartons})."
            )

        is_valid = (len(errors) == 0)
        return PrePrintValidationResult(
            is_valid=is_valid,
            total_physical_units=tot_u,
            total_physical_cartons=tot_c,
            discrepancy_count=len(warnings),
            unresolved_revueltos=0,
            formula_errors=formula_errors,
            warnings=warnings,
            errors=errors
        )
    finally:
        wb.close()


def _populate_item_row(
    ws: Any,
    r: int,
    item: MaterialDiscrepancy,
    qc_record: Optional[Any]
) -> None:
    """Write verified values into columns H, J, L, AS, AT, AU, AW, AX, AY of CIERRE RECEP."""
    # Repair formula in Column H (col 8) if broken
    val_h = ws.cell(row=r, column=8).value
    if val_h is None or not str(val_h).startswith("="):
        ws.cell(row=r, column=8, value=f"=+G{r}-J{r}-N{r}-O{r}+L{r}+P{r}+Q{r}")

    # Revueltos (Cols J & L) if paired
    if item.discrepancy_type == DiscrepancyType.REVUELTO_CANDIDATE:
        if item.difference_units < 0:
            ws.cell(row=r, column=10, value=abs(item.difference_units)) # Col J (Revuelto Saliendo)
            ws.cell(row=r, column=11, value=item.revuelto_partner or "") # Col K (Destino)
        elif item.difference_units > 0:
            ws.cell(row=r, column=12, value=item.difference_units)      # Col L (Revuelto Entrando)
            ws.cell(row=r, column=13, value=item.revuelto_partner or "") # Col M (Origen)

    # Column AS (45): TOTAL BULTOS DESCARGA
    ws.cell(row=r, column=45, value=item.physical_cartons)
    # Column AT (46): CANTIDAD DE UNIDADES EN BULTOS
    ws.cell(row=r, column=46, value=item.physical_units)
    # Column AU (47): Diferencia Físico vs Pedido Nac
    ws.cell(row=r, column=47, value=f"=+AT{r}-H{r}")

    # Column AW (49): Bultos según Packing List
    ws.cell(row=r, column=49, value=item.declared_cartons)
    # Column AX (50): CANTIDAD UNIDADES (Packing List)
    ws.cell(row=r, column=50, value=item.declared_units)

    # Column AY (51): Empaque
    curve = qc_record.curve_system if qc_record else None
    empaque_val = str(curve) if curve and curve > 1 else "1"
    ws.cell(row=r, column=51, value=empaque_val)


def export_receiving_closure(
    template_path: str,
    output_path: str,
    summary: ComparisonSummary,
    qc_report: Optional[QCInspectionReport] = None
) -> PrePrintValidationResult:
    """
    Populate and validate corporate receiving closure workbook.
    Preserves all VBA macros and standard formatting.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template closure file not found: {template_path}")

    # Copy template to output path preserving VBA macros
    shutil.copy2(template_path, output_path)

    wb = openpyxl.load_workbook(output_path, keep_vba=True)
    try:
        ws = wb["CIERRE RECEP"]

        # Update header summary values (Rows 5-7)
        ws.cell(row=5, column=2, value=summary.total_declared_cartons)
        ws.cell(row=6, column=2, value=summary.total_physical_cartons)
        ws.cell(row=7, column=2, value=f"=+B6-B5")

        qc_map = {it.material_code: it for it in qc_report.items} if qc_report else {}
        item_map = {it.material_code: it for it in summary.items}

        # Iterate over table rows in template and match by Material code in Column E (col 5)
        for r in range(13, 300):
            mat_val = ws.cell(row=r, column=5).value
            if mat_val is None:
                continue
            mat_code = str(mat_val).strip()
            if mat_code in item_map:
                disc_item = item_map[mat_code]
                _populate_item_row(ws, r, disc_item, qc_map.get(mat_code))

        wb.save(output_path)
    finally:
        wb.close()

    # Perform pre-print validation on the generated document
    return validate_closure_document(
        output_path,
        expected_units=summary.total_physical_units,
        expected_cartons=summary.total_physical_cartons
    )
