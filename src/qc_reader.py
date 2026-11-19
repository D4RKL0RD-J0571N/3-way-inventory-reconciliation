"""
Service 2: QC Spreadsheet Ingestion Reader (Lector de Hojas de Control de Calidad).

Robust, streaming reader for Quality Control workbooks (specifically 'CONTROL DATA' sheets)
supporting macro-enabled (.xlsm) and standard (.xlsx) files.
Extracts material lines, child size breakdowns, packaging counts, and runs
deterministic validation via Service 1 (quantity_engine).
Single-pass streaming architecture with early-exit prevents reading millions of empty rows.
"""

from dataclasses import dataclass, field
import os
import re
from typing import Any, Dict, List, Optional, Tuple
import openpyxl

from src.quantity_engine import (
    CartonCalculationInput,
    CartonCalculationResult,
    calculate_carton_quantities,
)


@dataclass
class QCItemRecord:
    """Structured representation of an item line from the QC inspection sheet."""
    row_index: int
    pallet_number: Optional[str]
    category: str
    vendor_material: str
    material_code: str
    description: str
    uom: str
    declared_units: Optional[int] = None
    curve_system: Optional[int] = None
    curve_physical: Optional[int] = None
    full_cartons: Optional[int] = None
    units_per_full_carton: Optional[int] = None
    complement_cartons: Optional[int] = None
    units_in_complement: Optional[int] = None
    reported_total_cartons: Optional[int] = None
    reported_total_units: Optional[int] = None
    reported_total_prk: Optional[int] = None
    reported_diff: Optional[int] = None
    observation: Optional[str] = None
    is_variante_suelta: bool = False
    is_child_sku: bool = False
    parent_material: Optional[str] = None
    calculation_result: Optional[CartonCalculationResult] = None


@dataclass
class QCInspectionReport:
    """Complete container QC report containing metadata, items, and audit flags."""
    container_id: str
    file_path: str
    metadata: Dict[str, Any] = field(default_factory=dict)
    items: List[QCItemRecord] = field(default_factory=list)
    total_calculated_units: int = 0
    total_calculated_cartons: int = 0
    total_reported_units: int = 0
    total_reported_cartons: int = 0
    flagged_items: List[QCItemRecord] = field(default_factory=list)


def _to_int(val: Any) -> Optional[int]:
    """Safely convert cell value to integer, handling floats, None, and strings."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return int(round(val))
    s = str(val).strip()
    if not s or s.lower() in ("none", "nan", "null", "-", "c", "x"):
        return None
    try:
        return int(round(float(s)))
    except (ValueError, TypeError):
        return None


def _to_clean_str(val: Any) -> str:
    """Normalize string, stripping whitespace and formatting."""
    if val is None:
        return ""
    s = str(val).strip()
    return "" if s.lower() in ("none", "nan", "null") else s


def _extract_row_metadata(row: Tuple[Any, ...], meta: Dict[str, Any]) -> None:
    """Extract metadata keywords from rows 1-5."""
    r_str = [str(c).strip() if c is not None else "" for c in row]
    for idx, text in enumerate(r_str):
        tu = text.upper()
        if "CONTENEDOR" in tu and idx + 1 < len(r_str) and r_str[idx + 1]:
            meta["container_id"] = r_str[idx + 1]
        elif "RESPONSABLE" in tu and idx + 1 < len(r_str) and r_str[idx + 1]:
            meta["responsible"] = r_str[idx + 1]
        elif "NACIONALIZACIÓN" in tu or "NACIONALIZACION" in tu:
            val = r_str[idx + 1] if idx + 1 < len(r_str) else ""
            if val:
                meta["nationalization_po"] = val
        elif "FECHA" in tu and idx + 1 < len(r_str) and r_str[idx + 1]:
            meta["date"] = r_str[idx + 1]


def _build_column_mapping(header_row: Tuple[Any, ...]) -> Dict[str, int]:
    """Map standardized field names to column indices from header row."""
    mapping: Dict[str, int] = {}
    field_keywords = {
        "pallet": ["TARIMA", "NUMERO DE TARIMA"],
        "category": ["CATEGORIA", "CATEGORÍA", "UEN"],
        "vendor_material": ["MATERIAL PROVEEDOR", "MAT. PROVEEDOR"],
        "material": ["MATERIAL"],
        "declared_units": ["CANTIDAD DE UNIDADES", "CANTIDAD"],
        "curve_system": ["CURVA SISTEMA"],
        "curve_physical": ["CURVA FISICA", "CURVA FÍSICA"],
        "uom": ["UM", "UNIDAD DE MEDIDA"],
        "description": ["DESCRIPCIÓN", "DESCRIPCION"],
        "full_cartons": ["BULTOS ENTEROS"],
        "units_per_full": ["UNIDADES POR BULTOS COMPLETOS", "UNIDADES POR BULTO"],
        "comp_cartons": ["BULTOS COMPLEMENTOS", "BULTOS COMPLEMENTO"],
        "units_comp": ["UNIDADES BULTOS COMPLEMENTOS", "UNIDADES EN BULTOS COMPLEMENTOS"],
        "total_cartons": ["TOTAL DE BULTOS", "TOTAL BULTOS"],
        "total_units": ["TOTAL DE UNIDADES", "TOTAL UNIDADES"],
        "total_prk": ["TOTAL PRK"],
        "diff_units": ["DIFERENCIA EN UNIDADES", "DIFERENCIA"],
        "observation": ["OBSERVACIONES", "OBSERVACION"],
    }

    for idx, cell in enumerate(header_row):
        if not cell:
            continue
        c_clean = str(cell).replace("\n", " ").strip().upper()
        for field_name, keywords in field_keywords.items():
            if field_name not in mapping:
                for kw in keywords:
                    if kw in c_clean:
                        if field_name == "material" and "PROVEEDOR" in c_clean:
                            continue
                        if field_name == "comp_cartons" and "UNIDADES" in c_clean:
                            continue
                        if field_name == "diff_units" and "MATERIAL" in c_clean:
                            continue
                        mapping[field_name] = idx
                        break
    return mapping


def _is_footer_or_total_row(mat_str: str, desc_str: str) -> bool:
    """Determine whether row is a footer summary row rather than item merchandise."""
    mat_upper = mat_str.upper()
    desc_upper = desc_str.upper()
    if (
        "TOTAL" in mat_upper
        or "RESUMEN" in mat_upper
        or "FIRMA" in mat_upper
        or mat_upper.startswith("SUMA")
        or desc_upper.startswith("TOTAL")
        or desc_upper.startswith("RESUMEN")
    ):
        return True
    if not re.search(r"\d", mat_str) and not desc_str:
        return True
    return False


def _parse_row_to_record(
    r_idx: int,
    row: Tuple[Any, ...],
    mapping: Dict[str, int],
    last_parent_mat: Optional[str]
) -> Optional[QCItemRecord]:
    """Parse a single spreadsheet row into a QCItemRecord."""
    def get_val(key: str) -> Any:
        idx = mapping.get(key)
        return row[idx] if idx is not None and idx < len(row) else None

    mat_raw = _to_clean_str(get_val("material"))
    desc = _to_clean_str(get_val("description"))
    if not mat_raw or _is_footer_or_total_row(mat_raw, desc):
        return None

    tarima = _to_clean_str(get_val("pallet"))
    is_vs = "VARIANTE SUELTA" in tarima.upper()
    is_child = (len(mat_raw) >= 10 and not tarima and not is_vs)
    uom = _to_clean_str(get_val("uom")) or ("PRK" if get_val("total_prk") else "UN")
    obs = _to_clean_str(get_val("observation"))

    rec = QCItemRecord(
        row_index=r_idx,
        pallet_number=tarima if tarima else None,
        category=_to_clean_str(get_val("category")),
        vendor_material=_to_clean_str(get_val("vendor_material")),
        material_code=mat_raw,
        description=desc,
        uom=uom,
        declared_units=_to_int(get_val("declared_units")),
        curve_system=_to_int(get_val("curve_system")),
        curve_physical=_to_int(get_val("curve_physical")),
        full_cartons=_to_int(get_val("full_cartons")),
        units_per_full_carton=_to_int(get_val("units_per_full")),
        complement_cartons=_to_int(get_val("comp_cartons")),
        units_in_complement=_to_int(get_val("units_comp")),
        reported_total_cartons=_to_int(get_val("total_cartons")),
        reported_total_units=_to_int(get_val("total_units")),
        reported_total_prk=_to_int(get_val("total_prk")),
        reported_diff=_to_int(get_val("diff_units")),
        observation=obs if obs else None,
        is_variante_suelta=is_vs,
        is_child_sku=is_child,
        parent_material=last_parent_mat if is_child else None
    )

    calc_inp = CartonCalculationInput(
        full_cartons=rec.full_cartons,
        units_per_full_carton=rec.units_per_full_carton,
        complement_cartons=rec.complement_cartons,
        units_in_complement=rec.units_in_complement,
        reported_total_cartons=rec.reported_total_cartons,
        reported_total_units=rec.reported_total_units,
        reported_diff=rec.reported_diff,
        curve=rec.curve_system or rec.curve_physical,
        uom=rec.uom,
        observation=rec.observation,
        is_variante_suelta=rec.is_variante_suelta
    )
    rec.calculation_result = calculate_carton_quantities(calc_inp)
    return rec


def _find_qc_sheet_name(sheetnames: List[str]) -> str:
    """Find the best matching QC inspection sheet name."""
    for name in sheetnames:
        if "CONTROL DATA" in name.upper():
            return name
    for name in sheetnames:
        if "CONTROL" in name.upper() or "CALIDAD" in name.upper():
            return name
    return sheetnames[0]


def _sum_report_totals(items: List[QCItemRecord]) -> Tuple[int, int, int, int]:
    """Compute aggregate sums for units and cartons across primary merchandise items."""
    calc_u = sum(
        it.calculation_result.calculated_units for it in items
        if it.calculation_result and not it.is_child_sku
    )
    calc_c = sum(
        it.calculation_result.calculated_cartons for it in items
        if it.calculation_result and not it.is_child_sku
    )
    rep_u = sum(it.reported_total_units or 0 for it in items if not it.is_child_sku)
    rep_c = sum(it.reported_total_cartons or 0 for it in items if not it.is_child_sku)
    return calc_u, calc_c, rep_u, rep_c


def read_qc_inspection_file(file_path: str) -> QCInspectionReport:
    """
    Open and parse a QC inspection spreadsheet in a single forward streaming pass.
    Early-exit after 15 consecutive empty rows ensures sub-second execution.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"QC file not found at: {file_path}")

    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    try:
        ws = wb[_find_qc_sheet_name(wb.sheetnames)]
        metadata: Dict[str, Any] = {}
        mapping: Dict[str, int] = {}
        header_idx: Optional[int] = None
        items: List[QCItemRecord] = []
        flagged: List[QCItemRecord] = []
        last_parent: Optional[str] = None
        empty_count = 0

        for r_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if not row or not any(c is not None for c in row):
                if header_idx is not None:
                    empty_count += 1
                    if empty_count >= 15:
                        break
                continue
            empty_count = 0

            if r_idx <= 5:
                _extract_row_metadata(row, metadata)
                continue
            if header_idx is None:
                row_str = " ".join(str(c).upper() for c in row if c is not None)
                if "MATERIAL" in row_str and ("BULTOS" in row_str or "UNIDADES" in row_str):
                    header_idx = r_idx
                    mapping = _build_column_mapping(row)
                continue

            record = _parse_row_to_record(r_idx, row, mapping, last_parent)
            if record:
                if not record.is_child_sku:
                    last_parent = record.material_code
                items.append(record)
                if record.calculation_result and record.calculation_result.flags:
                    flagged.append(record)

        container_id = metadata.get("container_id") or os.path.basename(file_path)
        tot_calc_u, tot_calc_c, tot_rep_u, tot_rep_c = _sum_report_totals(items)

        return QCInspectionReport(
            container_id=str(container_id),
            file_path=file_path,
            metadata=metadata,
            items=items,
            total_calculated_units=tot_calc_u,
            total_calculated_cartons=tot_calc_c,
            total_reported_units=tot_rep_u,
            total_reported_cartons=tot_rep_c,
            flagged_items=flagged
        )
    finally:
        wb.close()
