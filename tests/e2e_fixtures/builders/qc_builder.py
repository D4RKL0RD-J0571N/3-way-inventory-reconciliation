"""Synthetic Quality Control (.xlsm) fixture builder covering 15-layout variants and edge cases."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl

from openxml_utils import finalize_workbook

CANONICAL_HEADERS: dict[int, str] = {
    1: "Numero de Linea",
    2: "Tarima",
    3: "CATEGORIA",
    4: "Material proveedor",
    5: "Material",
    6: "CANTIDAD DE UNIDADES",
    9: "CURVA SISTEMA",
    11: "CURVA FISICA",
    12: "UM",
    13: "DESCRIPCIÓN",
    23: "CONTEO PIEZAS",
    29: "BULTOS ENTEROS",
    30: "UNIDADES POR BULTOS COMPLETOS",
    31: "BULTOS COMPLEMENTOS",
    32: "UNIDADES BULTOS COMPLEMENTOS",
    33: "TOTAL DE BULTOS",
    34: "TOTAL DE UNIDADES",
    35: "TOTAL PRK",
    36: "CURVA",
    37: "DIFERENCIA EN UNIDADES",
    38: "OBSERVACIONES",
}


def _populate_qc_metadata(ws: Any, container_id: str, po_43: str) -> None:
    """Write rows 1-5 metadata and merged header banners."""
    ws["A1"] = "IMPORTADORA ALMACEMA, S.A. — CONTROL DE CALIDAD"
    ws["A2"] = "CONTENEDOR:"
    ws["B2"] = container_id
    ws["D2"] = "RESPONSABLE:"
    ws["E2"] = "Inspector QC Sintético"
    ws["A3"] = "NACIONALIZACION:"
    ws["B3"] = po_43
    ws["D3"] = "FECHA:"
    ws["E3"] = "2026-11-18"
    ws["A4"] = "BODEGA C001 — MUELLE RECEPCION"
    ws["A5"] = "DATOS GENERALES DE INSPECCION"
    for row_idx in (1, 4, 5):
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)


def _apply_layout_headers(layout: str) -> dict[int, str]:
    """Return column-index-to-header mapping for the requested QC layout."""
    if layout == "canonical":
        return dict(CANONICAL_HEADERS)
    if layout == "no_tarima":
        return {(c - 1 if c > 2 else c): h for c, h in CANONICAL_HEADERS.items() if c != 2}
    if layout == "shifted_spacer":
        return {(c + 1 if c >= 33 else c): h for c, h in CANONICAL_HEADERS.items()}
    if layout == "no_comp_cartons":
        return {(c - 1 if c > 31 else c): h for c, h in CANONICAL_HEADERS.items() if c != 31}
    if layout == "numero_tarima":
        hdr = dict(CANONICAL_HEADERS)
        hdr[2] = "NUMERO DE TARIMA"
        return hdr
    if layout == "material_col_g":
        shifted: dict[int, str] = {}
        for c, h in CANONICAL_HEADERS.items():
            shifted[c + 2 if c >= 5 else c] = h
        return shifted
    raise ValueError(f"Unsupported QC layout: {layout}")


def _write_qc_headers(ws: Any, headers: dict[int, str]) -> dict[str, int]:
    """Write Row-6 headers and return reverse mapping from canonical header name to col."""
    reverse: dict[str, int] = {}
    for col_idx, title in headers.items():
        ws.cell(row=6, column=col_idx, value=title)
        reverse[title] = col_idx
    return reverse


def _write_qc_rows(
    ws: Any,
    col_by_title: dict[str, int],
    rows: list[dict[str, Any]],
    start_row: int = 7,
) -> tuple[int, dict[str, Any]]:
    """Write item rows starting at start_row and return (next_row, cached_values)."""
    cached: dict[str, Any] = {}
    cur_row = start_row
    for item in rows:
        for header_title, val in item.items():
            if header_title.startswith("_"):
                continue
            col_idx = col_by_title.get(header_title)
            if col_idx is None:
                continue
            cell = ws.cell(row=cur_row, column=col_idx, value=val)
            if isinstance(val, str) and val.startswith("=") and "_cached_full" in item:
                if header_title == "BULTOS ENTEROS":
                    cached[cell.coordinate] = item["_cached_full"]
        cur_row += 1
    return cur_row, cached


def _write_qc_footers(
    ws: Any,
    col_by_title: dict[str, int],
    footer_row: int,
    total_cartons: int,
    total_units: int,
) -> None:
    """Append TOTAL and FIRMA footer rows."""
    mat_col = col_by_title.get("Material", 5)
    cart_col = col_by_title.get("TOTAL DE BULTOS", 33)
    unit_col = col_by_title.get("TOTAL DE UNIDADES", 34)
    ws.cell(row=footer_row, column=mat_col, value="TOTAL")
    ws.cell(row=footer_row, column=cart_col, value=total_cartons)
    ws.cell(row=footer_row, column=unit_col, value=total_units)
    ws.cell(row=footer_row + 1, column=mat_col, value="FIRMA INSPECTOR:")
    ws.cell(row=footer_row + 1, column=mat_col + 1, value="Control Calidad C001")


def _canonical_parent_rows(use_formula: bool) -> list[dict[str, Any]]:
    """Build primary parent rows for clean canonical QC fixtures."""
    bulto_val: Any = "=W7/AD7" if use_formula else 10
    return [
        {
            "Numero de Linea": 1, "Tarima": "T-01", "CATEGORIA": "CABALLERO",
            "Material proveedor": "VS-101", "Material": "1455103",
            "CANTIDAD DE UNIDADES": 480, "CURVA SISTEMA": 12, "CURVA FISICA": 12,
            "UM": "PRK", "DESCRIPCIÓN": "Camisa Caballero Prepack",
            "CONTEO PIEZAS": 480, "BULTOS ENTEROS": bulto_val, "_cached_full": 10,
            "UNIDADES POR BULTOS COMPLETOS": 48, "BULTOS COMPLEMENTOS": 0,
            "UNIDADES BULTOS COMPLEMENTOS": 0, "TOTAL DE BULTOS": 10,
            "TOTAL DE UNIDADES": 480, "TOTAL PRK": 40, "CURVA": 12,
            "DIFERENCIA EN UNIDADES": 0, "OBSERVACIONES": "SIN NOVEDAD",
        },
        {
            "Numero de Linea": 4, "Tarima": "T-02", "CATEGORIA": "DAMA",
            "Material proveedor": "VS-202", "Material": "5033284",
            "CANTIDAD DE UNIDADES": 600, "CURVA SISTEMA": 1, "CURVA FISICA": 1,
            "UM": "UN", "DESCRIPCIÓN": "Blusa Dama Basica",
            "CONTEO PIEZAS": 576, "BULTOS ENTEROS": 8,
            "UNIDADES POR BULTOS COMPLETOS": 72, "BULTOS COMPLEMENTOS": 1,
            "UNIDADES BULTOS COMPLEMENTOS": 24, "TOTAL DE BULTOS": 9,
            "TOTAL DE UNIDADES": 600, "CURVA": 1,
            "DIFERENCIA EN UNIDADES": 0, "OBSERVACIONES": "8*72+24",
        },
    ]


def _canonical_sample_rows(use_formula: bool = True) -> list[dict[str, Any]]:
    """Build deterministic clean QC item rows including child SKUs."""
    parents = _canonical_parent_rows(use_formula)
    children = [
        {
            "Numero de Linea": 2, "Tarima": None, "CATEGORIA": "CABALLERO",
            "Material proveedor": "VS-101-S", "Material": "1455103001",
            "CANTIDAD DE UNIDADES": 240, "CURVA SISTEMA": 6, "UM": "UN",
            "DESCRIPCIÓN": "Camisa Talla S (Hijo)",
        },
        {
            "Numero de Linea": 3, "Tarima": None, "CATEGORIA": "CABALLERO",
            "Material proveedor": "VS-101-M", "Material": "1455103002",
            "CANTIDAD DE UNIDADES": 240, "CURVA SISTEMA": 6, "UM": "UN",
            "DESCRIPCIÓN": "Camisa Talla M (Hijo)",
        },
    ]
    return [parents[0], children[0], children[1], parents[1]]


def _edge_case_complement_and_variant_rows() -> list[dict[str, Any]]:
    """Build Complement Rescue, Empty Complement, and Variante Suelta edge rows."""
    return [
        {
            "Numero de Linea": 1, "Tarima": "T-01", "CATEGORIA": "CABALLERO",
            "Material proveedor": "EC-01", "Material": "2001001",
            "CANTIDAD DE UNIDADES": 1248, "CURVA SISTEMA": 12, "UM": "PRK",
            "DESCRIPCIÓN": "Complement Rescue a=6", "BULTOS ENTEROS": 6,
            "UNIDADES POR BULTOS COMPLETOS": 192, "BULTOS COMPLEMENTOS": None,
            "UNIDADES BULTOS COMPLEMENTOS": 48, "TOTAL DE BULTOS": 6,
            "TOTAL DE UNIDADES": 1200, "TOTAL PRK": 100,
            "DIFERENCIA EN UNIDADES": -48, "OBSERVACIONES": "6*192 +48",
        },
        {
            "Numero de Linea": 2, "Tarima": "T-02", "CATEGORIA": "CABALLERO",
            "Material proveedor": "EC-02", "Material": "2001002",
            "CANTIDAD DE UNIDADES": 1200, "CURVA SISTEMA": 12, "UM": "PRK",
            "DESCRIPCIÓN": "Complement Rescue a+1", "BULTOS ENTEROS": 6,
            "UNIDADES POR BULTOS COMPLETOS": 192, "BULTOS COMPLEMENTOS": 0,
            "UNIDADES BULTOS COMPLEMENTOS": 48, "TOTAL DE BULTOS": 7,
            "TOTAL DE UNIDADES": 1200, "TOTAL PRK": 100,
            "DIFERENCIA EN UNIDADES": 0, "OBSERVACIONES": "RESCATE A+1",
        },
        {
            "Numero de Linea": 3, "Tarima": "T-03", "CATEGORIA": "DAMA",
            "Material proveedor": "EC-03", "Material": "2001003",
            "CANTIDAD DE UNIDADES": 24, "CURVA SISTEMA": 1, "UM": "UN",
            "DESCRIPCIÓN": "Complement Rescue a=0", "BULTOS ENTEROS": 0,
            "UNIDADES POR BULTOS COMPLETOS": 0, "BULTOS COMPLEMENTOS": None,
            "UNIDADES BULTOS COMPLEMENTOS": 24, "TOTAL DE BULTOS": 1,
            "TOTAL DE UNIDADES": 24, "DIFERENCIA EN UNIDADES": 0,
        },
        {
            "Numero de Linea": 4, "Tarima": "T-04", "CATEGORIA": "DAMA",
            "Material proveedor": "EC-04", "Material": "2001004",
            "CANTIDAD DE UNIDADES": 500, "CURVA SISTEMA": 1, "UM": "UN",
            "DESCRIPCIÓN": "Empty Complement Cartons", "BULTOS ENTEROS": 10,
            "UNIDADES POR BULTOS COMPLETOS": 50, "BULTOS COMPLEMENTOS": 2,
            "UNIDADES BULTOS COMPLEMENTOS": 0, "TOTAL DE BULTOS": 12,
            "TOTAL DE UNIDADES": 500, "DIFERENCIA EN UNIDADES": 0,
        },
        {
            "Numero de Linea": 5, "Tarima": "VARIANTE SUELTA", "CATEGORIA": "INFANTIL",
            "Material proveedor": "EC-05", "Material": "2001005",
            "CANTIDAD DE UNIDADES": 10, "CURVA SISTEMA": 1, "UM": "UN",
            "DESCRIPCIÓN": "Variante Suelta Item", "TOTAL DE BULTOS": 0,
            "TOTAL DE UNIDADES": 10, "DIFERENCIA EN UNIDADES": 0,
        },
    ]


def _edge_case_shared_and_curve_rows() -> list[dict[str, Any]]:
    """Build Shared Carton and Uneven Prepack Curve edge rows."""
    return [
        {
            "Numero de Linea": 6, "Tarima": "T-06", "CATEGORIA": "DAMA",
            "Material proveedor": "EC-06A", "Material": "2001006",
            "CANTIDAD DE UNIDADES": 300, "CURVA SISTEMA": 1, "UM": "UN",
            "DESCRIPCIÓN": "Bulto Compartido Primario", "BULTOS ENTEROS": 1,
            "UNIDADES POR BULTOS COMPLETOS": 300, "TOTAL DE BULTOS": 1,
            "TOTAL DE UNIDADES": 300, "OBSERVACIONES": "DOS MATERIALES POR BULTO",
        },
        {
            "Numero de Linea": 7, "Tarima": "T-06", "CATEGORIA": "DAMA",
            "Material proveedor": "EC-06B", "Material": "2001007",
            "CANTIDAD DE UNIDADES": 300, "CURVA SISTEMA": 1, "UM": "UN",
            "DESCRIPCIÓN": "Bulto Compartido Secundario", "BULTOS ENTEROS": None,
            "UNIDADES POR BULTOS COMPLETOS": 300, "TOTAL DE BULTOS": 0,
            "TOTAL DE UNIDADES": 300, "OBSERVACIONES": "DOS MATERIALES POR BULTO",
        },
        {
            "Numero de Linea": 8, "Tarima": "T-08", "CATEGORIA": "CABALLERO",
            "Material proveedor": "EC-08", "Material": "2001008",
            "CANTIDAD DE UNIDADES": 742, "CURVA SISTEMA": 12, "UM": "PRK",
            "DESCRIPCIÓN": "Uneven Prepack Curve", "BULTOS ENTEROS": 15,
            "UNIDADES POR BULTOS COMPLETOS": 48, "BULTOS COMPLEMENTOS": 1,
            "UNIDADES BULTOS COMPLEMENTOS": 22, "TOTAL DE BULTOS": 16,
            "TOTAL DE UNIDADES": 742, "TOTAL PRK": 61,
        },
    ]


def _edge_case_obs_and_token_rows() -> list[dict[str, Any]]:
    """Build Observation Adjustment, IMP298/192, and sentinel token rows."""
    return [
        {
            "Numero de Linea": 9, "Tarima": "T-09", "CATEGORIA": "CABALLERO",
            "Material proveedor": "EC-09", "Material": "2001009",
            "CANTIDAD DE UNIDADES": 96, "CURVA SISTEMA": 12, "UM": "PRK",
            "DESCRIPCIÓN": "Obs Adjustment Mas 1PK", "BULTOS ENTEROS": 4,
            "UNIDADES POR BULTOS COMPLETOS": 24, "TOTAL DE BULTOS": 4,
            "TOTAL DE UNIDADES": 96, "OBSERVACIONES": "MAS 1PK",
        },
        {
            "Numero de Linea": 10, "Tarima": "T-10", "CATEGORIA": "CABALLERO",
            "Material proveedor": "EC-10", "Material": "2001010",
            "CANTIDAD DE UNIDADES": 96, "CURVA SISTEMA": 12, "UM": "PRK",
            "DESCRIPCIÓN": "Obs Adjustment Menos 1PK", "BULTOS ENTEROS": 4,
            "UNIDADES POR BULTOS COMPLETOS": 24, "TOTAL DE BULTOS": 4,
            "TOTAL DE UNIDADES": 96, "OBSERVACIONES": "MENOS 1PK",
        },
        {
            "Numero de Linea": 11, "Tarima": "T-11", "CATEGORIA": "DAMA",
            "Material proveedor": "EC-11", "Material": "2001011",
            "CANTIDAD DE UNIDADES": 252, "CURVA SISTEMA": 12, "UM": "PRK",
            "DESCRIPCIÓN": "IMP298 Mismatch Pattern", "BULTOS ENTEROS": 10,
            "UNIDADES POR BULTOS COMPLETOS": 24, "BULTOS COMPLEMENTOS": 1,
            "UNIDADES BULTOS COMPLEMENTOS": 12, "TOTAL DE BULTOS": 10,
            "TOTAL DE UNIDADES": 240, "OBSERVACIONES": "OMITIO COMPLEMENTO",
        },
        {
            "Numero de Linea": 12, "Tarima": "T-12", "CATEGORIA": "DAMA",
            "Material proveedor": "EC-12", "Material": "2001012",
            "CANTIDAD DE UNIDADES": 672, "CURVA SISTEMA": 12, "UM": "PRK",
            "DESCRIPCIÓN": "IMP192 Ambiguous Plus Token", "BULTOS ENTEROS": 56,
            "UNIDADES POR BULTOS COMPLETOS": 12, "TOTAL DE BULTOS": 56,
            "TOTAL DE UNIDADES": 672, "OBSERVACIONES": "56*12+672",
        },
        {
            "Numero de Linea": 13, "Tarima": "T-13", "CATEGORIA": "INFANTIL",
            "Material proveedor": "EC-13", "Material": "2001013",
            "CANTIDAD DE UNIDADES": 100, "CURVA SISTEMA": 1, "CURVA FISICA": "nan",
            "UM": "UN", "DESCRIPCIÓN": "Cell Token Sentinel Row",
            "BULTOS ENTEROS": 5, "UNIDADES POR BULTOS COMPLETOS": 20,
            "BULTOS COMPLEMENTOS": "-", "UNIDADES BULTOS COMPLEMENTOS": "C",
            "TOTAL DE BULTOS": 5, "TOTAL DE UNIDADES": 100,
            "DIFERENCIA EN UNIDADES": "X", "OBSERVACIONES": "TOKENS - C X nan",
        },
    ]


def _edge_case_rows() -> list[dict[str, Any]]:
    """Build combined rows exercising every QuantityEngine and QcReader edge case."""
    return (
        _edge_case_complement_and_variant_rows()
        + _edge_case_shared_and_curve_rows()
        + _edge_case_obs_and_token_rows()
    )


def _bloat_rows(count: int = 19) -> list[dict[str, Any]]:
    """Build 19 clean item rows for qc_bloat_streaming.xlsm."""
    rows: list[dict[str, Any]] = []
    for idx in range(1, count + 1):
        rows.append({
            "Numero de Linea": idx,
            "Tarima": f"T-{idx:02d}",
            "CATEGORIA": "CABALLERO",
            "Material proveedor": f"BL-{idx:03d}",
            "Material": f"3001{idx:03d}",
            "CANTIDAD DE UNIDADES": 120,
            "CURVA SISTEMA": 12,
            "UM": "PRK",
            "DESCRIPCIÓN": f"Item Streaming Bloat {idx:02d}",
            "BULTOS ENTEROS": 5,
            "UNIDADES POR BULTOS COMPLETOS": 24,
            "BULTOS COMPLEMENTOS": 0,
            "UNIDADES BULTOS COMPLEMENTOS": 0,
            "TOTAL DE BULTOS": 5,
            "TOTAL DE UNIDADES": 120,
            "TOTAL PRK": 10,
            "DIFERENCIA EN UNIDADES": 0,
        })
    return rows


def _reconciliation_discrepancy_rows() -> list[dict[str, Any]]:
    """Build 5 physical QC rows paired with closure_discrepancy_template.xlsm."""
    specs = [
        (1, "T-01", "CABALLERO", "VS-101", "1455103", 480, 12, "PRK", "Camisa Caballero Prepack", 10, 48, 0, 0, 10, 480, 40, 0),
        (2, "T-02", "DAMA", "VS-202", "5033284", 600, 1, "UN", "Blusa Dama Basica", 8, 72, 0, 0, 8, 576, None, -24),
        (3, "T-03", "CABALLERO", "VS-303", "1455104", 120, 12, "PRK", "Pantalon Caballero Slim", 5, 24, 1, 12, 6, 132, 11, 12),
        (4, "T-04", "CABALLERO", "VS-404", "1455105", 240, 12, "PRK", "Polo Caballero Basico A", 8, 24, 1, 12, 9, 204, 17, -36),
        (5, "T-05", "CABALLERO", "VS-404", "1455106", 240, 12, "PRK", "Polo Caballero Basico B", 11, 24, 1, 12, 12, 276, 23, 36),
    ]
    rows: list[dict[str, Any]] = []
    for ln, tar, cat, vmat, mat, qty, curve, uom, desc, fc, uf, cc, uc, tb, tu, prk, diff in specs:
        row: dict[str, Any] = {
            "Numero de Linea": ln, "Tarima": tar, "CATEGORIA": cat,
            "Material proveedor": vmat, "Material": mat, "CANTIDAD DE UNIDADES": qty,
            "CURVA SISTEMA": curve, "CURVA FISICA": curve, "UM": uom, "DESCRIPCIÓN": desc,
            "BULTOS ENTEROS": fc, "UNIDADES POR BULTOS COMPLETOS": uf,
            "BULTOS COMPLEMENTOS": cc, "UNIDADES BULTOS COMPLEMENTOS": uc,
            "TOTAL DE BULTOS": tb, "TOTAL DE UNIDADES": tu,
            "CURVA": curve, "DIFERENCIA EN UNIDADES": diff,
        }
        if prk is not None:
            row["TOTAL PRK"] = prk
        rows.append(row)
    return rows


def _create_qc_workbook(
    output_path: Path,
    *,
    container_id: str,
    po_43: str,
    layout: str,
    rows: list[dict[str, Any]],
    use_shared_strings: bool = True,
    bloat_row_count: int = 0,
    vba_variant: str = "qc_default",
) -> dict[str, Any]:
    """Create and finalize a single synthetic QC .xlsm workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CONTROL DATA"
    _populate_qc_metadata(ws, container_id, po_43)
    headers = _apply_layout_headers(layout)
    col_by_title = _write_qc_headers(ws, headers)
    next_row, cached = _write_qc_rows(ws, col_by_title, rows)
    tot_c = sum(int(r.get("TOTAL DE BULTOS") or 0) for r in rows)
    tot_u = sum(int(r.get("TOTAL DE UNIDADES") or 0) for r in rows)
    _write_qc_footers(ws, col_by_title, next_row + 1, total_cartons=tot_c, total_units=tot_u)
    bloat_map = {"CONTROL DATA": bloat_row_count} if bloat_row_count > 0 else None
    cached_map = {"CONTROL DATA": cached} if cached else None
    return finalize_workbook(
        wb,
        output_path,
        macro_enabled=True,
        use_shared_strings=use_shared_strings,
        vba_variant=vba_variant,
        cached_values=cached_map,
        bloat_sheets=bloat_map,
    )


def build_all_qc_fixtures(out_dir: Path, bloat_rows: int = 25000) -> dict[str, dict[str, Any]]:
    """Build all QC synthetic fixtures under out_dir/qc/ and return metadata."""
    qc_dir = out_dir / "qc"
    qc_dir.mkdir(parents=True, exist_ok=True)
    specs: list[tuple[str, str, str, list[dict[str, Any]], bool, int]] = [
        ("qc_canonical_clean.xlsm", "IMP-901", "canonical", _canonical_sample_rows(True), True, 0),
        ("qc_no_tarima_layout.xlsm", "IMP-902", "no_tarima", _canonical_sample_rows(False), True, 0),
        ("qc_shifted_spacer_layout.xlsm", "IMP-903", "shifted_spacer", _canonical_sample_rows(False), True, 0),
        ("qc_no_comp_cartons_layout.xlsm", "IMP-904", "no_comp_cartons", _canonical_sample_rows(False), True, 0),
        ("qc_numero_tarima_layout.xlsm", "IMP-905", "numero_tarima", _canonical_sample_rows(False), True, 0),
        ("qc_material_col_g_layout.xlsm", "IMP-906", "material_col_g", _canonical_sample_rows(False), True, 0),
        ("qc_edge_cases.xlsm", "IMP-907", "canonical", _edge_case_rows(), True, 0),
        ("qc_inlinestr.xlsm", "IMP-908", "canonical", _canonical_sample_rows(False), False, 0),
        ("qc_bloat_streaming.xlsm", "IMP-100", "canonical", _bloat_rows(19), True, bloat_rows),
        ("qc_reconciliation_discrepancies.xlsm", "IMP-909", "canonical", _reconciliation_discrepancy_rows(), True, 0),
        ("SYNTH_IMP81_NOCALIDAD.xlsm", "IMP-81", "canonical", _canonical_sample_rows(False), True, 0),
    ]
    results: dict[str, dict[str, Any]] = {}
    for fname, cid, layout, rows, use_ss, bloat in specs:
        rel = f"qc/{fname}"
        results[rel] = _create_qc_workbook(
            qc_dir / fname,
            container_id=cid,
            po_43="4300008100",
            layout=layout,
            rows=rows,
            use_shared_strings=use_ss,
            bloat_row_count=bloat,
            vba_variant=f"qc_{fname}",
        )
    return results
