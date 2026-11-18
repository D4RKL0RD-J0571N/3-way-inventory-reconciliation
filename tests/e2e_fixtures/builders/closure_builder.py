"""Synthetic Master Closure (.xlsm), legacy .xls, and lockfile builders."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName

from openxml_utils import (
    FIXED_ISO_TIME,
    build_legacy_xls_bytes,
    finalize_workbook,
    sha256_bytes,
    write_lockfile,
)

CIERRE_RECEP_HEADERS: dict[int, str | None] = {
    1: "ORDEN DE COMPRA",
    2: "Proveedor",
    3: "Categoría",
    4: "MAT. PROVEEDOR",
    5: "MATERIAL",
    6: "CANTIDAD PEDIDO 43",
    7: "CANTIDAD PEDIDO 41",
    8: "RECIBIDO REAL & FISICO",
    9: "CANTIDAD INGRESO MOV-101",
    10: "REVUELTO SALIENDO",
    11: "DESTINO",
    12: "REVUELTO ENTRANDO",
    13: "ORIGEN",
    14: "FALTANTE",
    15: "DAÑADO",
    16: "SOBRANTE",
    17: "UNIDADES CON MOV 501 EN TIENDA",
    18: "# DE MOVIMIENTO",
    19: "INDICADOR  DE  INCONSIST",
    20: "ACCIONES A SEGUIR U OBSERVACIONES",
    21: "CANTIDAD FACTURADA MOV-641",
    22: "DIFERENC MOV 101 VRS MOV 641",
    23: "MATERIAL NUEVO EN CREACIONES",
    24: "CANTIDAD ENVIADA  A PARTES ( B-C-D )",
    25: "STOCK  A004 RESURTIDO",
    26: "UNIDADES EN PROCESO DE FACTURAR (643)",
    27: "UNIDADES QUE APLICAN PARA NOTA DE CREDITO",
    28: "REVISION FINAL DE STOCK",
    29: "OBSERVACIONES GENERALES",
    30: "DESCRIPCION (TEXTO BREVE)",
    31: "VALOR UNITARIO",
    32: "Valor unitario PEDIDO NAC $",
    33: "Valor Revuelto",
    34: "Valor faltante",
    35: "Valor Dañado",
    36: "Valor Sobrante",
    37: "VALOR TOTAL PEDIDO 43",
    38: "VALOR TOTAL PEDIDO 41",
    39: "COD PROVEEDOR",
    40: "CATEGORIA",
    41: "COD CATEGORIA",
    42: None,
    43: None,
    44: None,
    45: "TOTAL BULTOS DESCARGA",
    46: "CANTIDAD DE UNIDADES EN BULTOS",
    47: "DIFERENCIA UNIDADES FISICO VS PEDIDO NAC",
    48: None,
    49: "Bultos según Paking List",
    50: "CANTIDAD UNIDADES",
    51: "Empaque",
    52: "Diferencia unidades facturadas vs solicitadas",
    53: None,
    54: "Curva",
    55: "UM",
    56: "CANTIDAD PEDIDO 43",
    57: "CANTIDAD PEDIDO 41",
    58: "CANTIDAD INGRESO MOV-101",
}


def populate_prov_y_categ(ws: Any) -> None:
    """Populate PROV Y CATEG lookup sheet for guarded VLOOKUP tests."""
    ws["A1"] = "COD PROVEEDOR"
    ws["B1"] = "PROVEEDOR"
    ws["A2"] = 1100000577
    ws["B2"] = "PROVEEDOR TEXTIL SINTETICO S.A."
    ws["A3"] = 1100000578
    ws["B3"] = "CONFECCIONES MODELO S.A."
    ws["F1"] = "COD CATEGORIA"
    ws["G1"] = "CATEGORIA"
    ws["F2"] = "G4"
    ws["G2"] = "CABALLERO"
    ws["F3"] = "G5"
    ws["G3"] = "DAMA"
    ws["F4"] = "G6"
    ws["G4"] = "INFANTIL"


def populate_datos_generales(
    ws: Any,
    *,
    container_id: str,
    po_43: str,
    pedido_41: str,
    author_b33: str,
    provider: str = "PROVEEDOR TEXTIL SINTETICO S.A.",
    bultos_teo: Any = 19,
    bultos_rec: Any = 19,
    lineas_c18: int = 2,
    fecha_rev: str = "2026-11-18",
) -> None:
    """Populate DATOS GENERALES sheet with operational identity and B33 author."""
    ws["B7"] = "CONTENEDOR"
    ws["C7"] = container_id
    ws["B8"] = "ORDEN DE COMPRA"
    ws["C8"] = po_43
    ws["B9"] = "PEDIDO NACIONALIZACION"
    ws["C9"] = pedido_41
    ws["B10"] = "PROVEEDOR"
    ws["C10"] = provider
    ws["B13"] = "BULTOS TEORICOS"
    ws["C13"] = bultos_teo
    ws["B14"] = "BULTOS RECIBIDOS"
    ws["C14"] = bultos_rec
    ws["B18"] = "CANTIDAD DE LINEAS"
    ws["C18"] = lineas_c18
    ws["B30"] = "FECHA REVISION"
    ws["C30"] = fecha_rev
    ws["B33"] = author_b33


def populate_bultos_sheet(ws: Any, item_count: int, stale_rows: int = 15) -> None:
    """Populate Bultos sheet (B..J) with linked rows plus stale trailing rows."""
    headers = [
        "Categoría", "Bultos", "Unidades", "Material",
        "Recibido", "Curva", "UM", "Descripción", "Empaque",
    ]
    for offset, title in enumerate(headers):
        ws.cell(row=2, column=2 + offset, value=title)
    total_rows = max(item_count, stale_rows)
    for idx in range(total_rows):
        b_row = 3 + idx
        cr_row = 13 + idx
        if idx < item_count:
            ws.cell(row=b_row, column=2, value=f"='CIERRE RECEP'!C{cr_row}")
            ws.cell(row=b_row, column=3, value=f"='CIERRE RECEP'!AS{cr_row}")
            ws.cell(row=b_row, column=4, value=f"='CIERRE RECEP'!AT{cr_row}")
            ws.cell(row=b_row, column=5, value=f"='CIERRE RECEP'!E{cr_row}")
            ws.cell(row=b_row, column=6, value=f"='CIERRE RECEP'!H{cr_row}")
            ws.cell(row=b_row, column=7, value=f"='CIERRE RECEP'!BB{cr_row}")
            ws.cell(row=b_row, column=8, value=f"='CIERRE RECEP'!BC{cr_row}")
            ws.cell(row=b_row, column=9, value=f"='CIERRE RECEP'!AD{cr_row}")
            ws.cell(row=b_row, column=10, value=f"='CIERRE RECEP'!AY{cr_row}")
        else:
            ws.cell(row=b_row, column=2, value="STALE_TEMPLATE_ROW")
            ws.cell(row=b_row, column=3, value=0)
            ws.cell(row=b_row, column=4, value=0)
            ws.cell(row=b_row, column=5, value=f"STALE_{b_row}")


def init_cierre_recep_headers(ws: Any, b5_val: Any = 19, b6_val: Any = 19) -> None:
    """Write summary header cells B5:B7, Row 11 AP helper, and 58 Row-12 headers."""
    ws["A5"] = "Bultos Packing List"
    ws["B5"] = b5_val
    ws["A6"] = "Bultos Descarga"
    ws["B6"] = b6_val
    ws["A7"] = "Diferencia Bultos"
    ws["B7"] = "=+B6-B5"
    ws.cell(row=11, column=44, value="AP")
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx, title in CIERRE_RECEP_HEADERS.items():
        cell = ws.cell(row=12, column=col_idx, value=title)
        if title is not None:
            cell.fill = header_fill
            cell.font = header_font


def write_cierre_item_row(ws: Any, row_idx: int, item: dict[str, Any]) -> None:
    """Write one 58-column CIERRE RECEP item row with authentic Excel formulas."""
    r = row_idx
    ws.cell(row=r, column=1, value=item.get("po_43", "4300008100"))
    ws.cell(row=r, column=2, value=f"=IF(ISNA(VLOOKUP(AM{r},'PROV Y CATEG'!$A$2:$B$194,2,0))=TRUE,\"\",VLOOKUP(AM{r},'PROV Y CATEG'!$A$2:$B$194,2,0))")
    ws.cell(row=r, column=3, value=f"=IF(ISNA(VLOOKUP(AO{r},'PROV Y CATEG'!$F$2:$G$76,2,0))=TRUE,\"\",VLOOKUP(AO{r},'PROV Y CATEG'!$F$2:$G$76,2,0))")
    ws.cell(row=r, column=4, value=item.get("vendor_material", "VS-101"))
    ws.cell(row=r, column=5, value=item["material_code"])
    ws.cell(row=r, column=6, value=f"=BD{r}*BB{r}")
    ws.cell(row=r, column=7, value=f"=BE{r}*BB{r}")
    ws.cell(row=r, column=8, value=f"=+G{r}-J{r}-N{r}-O{r}+L{r}+P{r}+Q{r}")
    ws.cell(row=r, column=9, value=f"=BF{r}*BB{r}")
    ws.cell(row=r, column=19, value=f"=+J{r}+L{r}+N{r}+O{r}+P{r}")
    ws.cell(row=r, column=21, value=item.get("mov_641", 0))
    ws.cell(row=r, column=22, value=f"=+U{r}-I{r}")
    _write_cierre_item_tail(ws, r, item)


def _write_cierre_item_tail(ws: Any, r: int, item: dict[str, Any]) -> None:
    """Write columns X..BF for a single CIERRE RECEP item row."""
    for col_idx in (24, 25, 26):
        ws.cell(row=r, column=col_idx, value=0)
    ws.cell(row=r, column=27, value=f"=+N{r}+O{r}")
    ws.cell(row=r, column=28, value=f"=V{r}+X{r}+Y{r}+Z{r}")
    ws.cell(row=r, column=30, value=item.get("description", "Prenda Sintetica"))
    ws.cell(row=r, column=31, value=f"=IFERROR(AK{r}/F{r},0)")
    ws.cell(row=r, column=32, value=f"=IFERROR(AL{r}/G{r},0)")
    ws.cell(row=r, column=33, value=f"=+J{r}*AE{r}")
    ws.cell(row=r, column=34, value=f"=+N{r}*AE{r}")
    ws.cell(row=r, column=35, value=f"=+O{r}*AE{r}")
    ws.cell(row=r, column=36, value=f"=+P{r}*AE{r}")
    ws.cell(row=r, column=37, value=item.get("val_43", 1200.0))
    ws.cell(row=r, column=38, value=item.get("val_41", 1200.0))
    ws.cell(row=r, column=39, value=item.get("cod_prov", 1100000577))
    ws.cell(row=r, column=40, value=item.get("cat_code", "G40403001"))
    ws.cell(row=r, column=41, value=f"=MID(AN{r},1,2)")
    ws.cell(row=r, column=43, value=f"=+D{r}")
    ws.cell(row=r, column=44, value=f"=+E{r}")
    ws.cell(row=r, column=45, value=item.get("as_cartons"))
    ws.cell(row=r, column=46, value=item.get("at_units"))
    ws.cell(row=r, column=47, value=f"=+AT{r}-H{r}")
    ws.cell(row=r, column=49, value=item.get("aw_cartons"))
    ws.cell(row=r, column=50, value=item.get("ax_units"))
    ws.cell(row=r, column=51, value=item.get("ay_empaque"))
    ws.cell(row=r, column=52, value=f"=+AX{r}-H{r}")
    ws.cell(row=r, column=54, value=item.get("curve", 1))
    ws.cell(row=r, column=55, value=item.get("uom", "UN"))
    ws.cell(row=r, column=56, value=item.get("qty_43", 10))
    ws.cell(row=r, column=57, value=item.get("qty_41", 10))
    ws.cell(row=r, column=58, value=item.get("qty_101", 10))


def add_closure_defined_names(wb: openpyxl.Workbook, max_row: int) -> None:
    """Register _FilterDatabase and Print_Area defined names on CIERRE RECEP."""
    wb.defined_names.add(
        DefinedName("_xlnm._FilterDatabase", attr_text="'CIERRE RECEP'!$A$12:$BF$12", localSheetId=0)
    )
    wb.defined_names.add(
        DefinedName("_xlnm.Print_Area", attr_text=f"'CIERRE RECEP'!$A$1:$BF${max_row}", localSheetId=0)
    )


def _assemble_closure_sheets(
    wb: openpyxl.Workbook,
    *,
    container_id: str,
    po_43: str,
    pedido_41: str,
    author_b33: str,
    items: list[dict[str, Any]],
    stale_bultos_rows: int,
    error_mode: str | None,
    include_control_data: bool,
    dg_provider: str,
    dg_bultos_teo: Any,
    dg_bultos_rec: Any,
    dg_lineas_c18: int | None,
) -> dict[str, dict[str, Any]] | None:
    """Populate CIERRE RECEP, Bultos, PROV Y CATEG, DATOS GENERALES, and optional CONTROL DATA."""
    ws_cr = wb.active
    ws_cr.title = "CIERRE RECEP"
    tot_c = sum(int(i.get("as_cartons") or 0) for i in items)
    init_cierre_recep_headers(ws_cr, b5_val=tot_c, b6_val=tot_c)
    for idx, item in enumerate(items):
        write_cierre_item_row(ws_cr, 13 + idx, item)
    cached_map: dict[str, dict[str, Any]] | None = None
    if error_mode == "preexisting":
        _inject_preexisting_closure_errors(ws_cr)
        cached_map = {
            "CIERRE RECEP": {
                "B5": "#REF!", "B6": "#REF!", "B7": "#REF!",
                "V14": "#VALUE!", "AB14": "#VALUE!", "AT16": "#N/A", "AU16": "#N/A", "AS17": "#REF!",
            }
        }
    populate_bultos_sheet(wb.create_sheet("Bultos"), len(items), stale_rows=stale_bultos_rows)
    populate_prov_y_categ(wb.create_sheet("PROV Y CATEG"))
    populate_datos_generales(
        wb.create_sheet("DATOS GENERALES"),
        container_id=container_id, po_43=po_43, pedido_41=pedido_41,
        author_b33=author_b33, provider=dg_provider,
        bultos_teo=tot_c if dg_bultos_teo is None else dg_bultos_teo,
        bultos_rec=tot_c if dg_bultos_rec is None else dg_bultos_rec,
        lineas_c18=len(items) if dg_lineas_c18 is None else dg_lineas_c18,
    )
    if include_control_data:
        ws_qc = wb.create_sheet("CONTROL DATA")
        ws_qc["A2"] = "CONTENEDOR:"
        ws_qc["B2"] = container_id
    add_closure_defined_names(wb, max(30, 12 + len(items)))
    return cached_map


def build_closure_workbook(
    output_path: Path,
    *,
    container_id: str,
    po_43: str,
    pedido_41: str,
    author_b33: str,
    items: list[dict[str, Any]],
    vba_variant: str = "closure_default",
    stale_bultos_rows: int = 15,
    omit_dimension: bool = False,
    error_mode: str | None = None,
    include_control_data: bool = False,
    dg_provider: str = "PROVEEDOR TEXTIL SINTETICO S.A.",
    dg_bultos_teo: Any = None,
    dg_bultos_rec: Any = None,
    dg_lineas_c18: int | None = None,
    core_creator: str = "Hilario Lopez Rivera",
    core_created: str = "2016-09-05T16:30:28Z",
    extra_shared_strings: list[str] | None = None,
) -> dict[str, Any]:
    """Assemble and deterministically finalize a synthetic master closure .xlsm."""
    wb = openpyxl.Workbook()
    cached_map = _assemble_closure_sheets(
        wb,
        container_id=container_id, po_43=po_43, pedido_41=pedido_41,
        author_b33=author_b33, items=items, stale_bultos_rows=stale_bultos_rows,
        error_mode=error_mode, include_control_data=include_control_data,
        dg_provider=dg_provider, dg_bultos_teo=dg_bultos_teo,
        dg_bultos_rec=dg_bultos_rec, dg_lineas_c18=dg_lineas_c18,
    )
    omit_set = {"CIERRE RECEP"} if omit_dimension else None
    return finalize_workbook(
        wb, output_path, macro_enabled=True, use_shared_strings=True,
        vba_variant=vba_variant, extra_shared_strings=extra_shared_strings,
        cached_values=cached_map, omit_dimension_sheets=omit_set,
        core_creator=core_creator, core_last_modified_by="ACR Fixture Generator",
        core_created=core_created, core_modified=FIXED_ISO_TIME,
    )


def _inject_preexisting_closure_errors(ws: Any) -> None:
    """Inject IMP-241/244/256/360/90 style pre-existing template errors."""
    ws["B5"] = "=#REF!"
    ws["B6"] = "=#REF!"
    ws["B7"] = "=+B6-B5"
    ws["U14"] = "X"
    ws["AY29"] = "#N/A"
    ws["E16"] = "9999999"
    ws["AT16"] = "=VLOOKUP(AR16,'PROV Y CATEG'!$A$2:$B$10,2,0)"
    ws["AS17"] = "=VLOOKUP(AR17,#REF!,2,0)"


def _default_clean_items() -> list[dict[str, Any]]:
    """Return 2 clean items matching qc_canonical_clean.xlsm plus optional VLOOKUP miss."""
    return [
        {
            "po_43": "4300008100", "vendor_material": "VS-101", "material_code": "1455103",
            "description": "Camisa Caballero Prepack", "cod_prov": 1100000577,
            "cat_code": "G40403001", "as_cartons": 10, "at_units": 480,
            "aw_cartons": 10, "ax_units": 480, "ay_empaque": "48",
            "curve": 12, "uom": "PRK", "qty_43": 40, "qty_41": 40, "qty_101": 40,
        },
        {
            "po_43": "4300008100", "vendor_material": "VS-202", "material_code": "5033284",
            "description": "Blusa Dama Basica", "cod_prov": 1100000999,
            "cat_code": "G50102001", "as_cartons": 9, "at_units": 600,
            "aw_cartons": 9, "ax_units": 600, "ay_empaque": "8*72+24",
            "curve": 1, "uom": "UN", "qty_43": 600, "qty_41": 600, "qty_101": 600,
        },
    ]


def _over500_items() -> list[dict[str, Any]]:
    """Generate 548 items (Rows 13..560) with multi-PO duplicate SKU 1455103."""
    items: list[dict[str, Any]] = []
    for idx in range(548):
        row_num = 13 + idx
        if row_num == 15:
            po, mat = "4300001001", "1455103"
        elif row_num == 350:
            po, mat = "4300001002", "1455103"
        else:
            po, mat = "4300008200", f"700{idx:04d}"
        items.append({
            "po_43": po, "vendor_material": f"VP-{idx:03d}", "material_code": mat,
            "description": f"Articulo Largo {row_num}", "cod_prov": 1100000577,
            "cat_code": "G40403001", "as_cartons": 2, "at_units": 48,
            "aw_cartons": 2, "ax_units": 48, "ay_empaque": "24",
            "curve": 12, "uom": "PRK", "qty_43": 4, "qty_41": 4, "qty_101": 4,
        })
    return items


def _error_fixture_items() -> list[dict[str, Any]]:
    """Build 18 items so Rows 14, 16, 17, 29 all sit on populated rows."""
    items = _default_clean_items()
    for idx in range(16):
        items.append({
            "po_43": "4300008300", "vendor_material": f"ER-{idx:02d}",
            "material_code": f"8001{idx:03d}", "description": f"Error Test Row {15 + idx}",
            "cod_prov": 1100000577, "cat_code": "G40403001",
            "as_cartons": 1, "at_units": 24, "aw_cartons": 1, "ax_units": 24,
            "ay_empaque": "24", "curve": 1, "uom": "UN",
            "qty_43": 24, "qty_41": 24, "qty_101": 24,
        })
    return items


def _discrepancy_template_items() -> list[dict[str, Any]]:
    """Build 5 declared items matching qc_reconciliation_discrepancies.xlsm."""
    return [
        {
            "po_43": "4300008100", "vendor_material": "VS-101", "material_code": "1455103",
            "description": "Camisa Caballero Prepack", "cod_prov": 1100000577,
            "cat_code": "G40403001", "as_cartons": 10, "at_units": 480,
            "aw_cartons": 10, "ax_units": 480, "ay_empaque": "48",
            "curve": 12, "uom": "PRK", "qty_43": 40, "qty_41": 40, "qty_101": 40,
        },
        {
            "po_43": "4300008100", "vendor_material": "VS-202", "material_code": "5033284",
            "description": "Blusa Dama Basica", "cod_prov": 1100000999,
            "cat_code": "G50102001", "as_cartons": 9, "at_units": 600,
            "aw_cartons": 9, "ax_units": 600, "ay_empaque": "8*72+24",
            "curve": 1, "uom": "UN", "qty_43": 600, "qty_41": 600, "qty_101": 600,
        },
        {
            "po_43": "4300008100", "vendor_material": "VS-303", "material_code": "1455104",
            "description": "Pantalon Caballero Slim", "cod_prov": 1100000577,
            "cat_code": "G40403001", "as_cartons": 5, "at_units": 120,
            "aw_cartons": 5, "ax_units": 120, "ay_empaque": "24",
            "curve": 12, "uom": "PRK", "qty_43": 10, "qty_41": 10, "qty_101": 10,
        },
        {
            "po_43": "4300008100", "vendor_material": "VS-404", "material_code": "1455105",
            "description": "Polo Caballero Basico A", "cod_prov": 1100000577,
            "cat_code": "G40403001", "as_cartons": 10, "at_units": 240,
            "aw_cartons": 10, "ax_units": 240, "ay_empaque": "24",
            "curve": 12, "uom": "PRK", "qty_43": 20, "qty_41": 20, "qty_101": 20,
        },
        {
            "po_43": "4300008100", "vendor_material": "VS-404", "material_code": "1455106",
            "description": "Polo Caballero Basico B", "cod_prov": 1100000577,
            "cat_code": "G40403001", "as_cartons": 10, "at_units": 240,
            "aw_cartons": 10, "ax_units": 240, "ay_empaque": "24",
            "curve": 12, "uom": "PRK", "qty_43": 20, "qty_41": 20, "qty_101": 20,
        },
    ]


def build_closure_missing_cierre(output_path: Path) -> dict[str, Any]:
    """Build valid .xlsm closure missing CIERRE RECEP for E2E-F12-T2-05."""
    wb = openpyxl.Workbook()
    ws_bultos = wb.active
    ws_bultos.title = "Bultos"
    populate_bultos_sheet(ws_bultos, item_count=2, stale_rows=15)
    populate_prov_y_categ(wb.create_sheet("PROV Y CATEG"))
    populate_datos_generales(
        wb.create_sheet("DATOS GENERALES"),
        container_id="IMP-999",
        po_43="4300008999",
        pedido_41="4100002999",
        author_b33="ELABORADO POR: \nJostin Lopez Sobalbarro",
    )
    return finalize_workbook(
        wb, output_path, macro_enabled=True, use_shared_strings=True,
        vba_variant="closure_missing_cierre",
        core_creator="Hilario Lopez Rivera",
        core_last_modified_by="ACR Fixture Generator",
        core_created="2016-09-05T16:30:28Z", core_modified=FIXED_ISO_TIME,
    )


def _build_standard_closure_fixtures(c_dir: Path) -> dict[str, dict[str, Any]]:
    """Build clean, discrepancy-matched, over-500, and error closure .xlsm files."""
    results: dict[str, dict[str, Any]] = {}
    results["closures/closure_template_clean.xlsm"] = build_closure_workbook(
        c_dir / "closure_template_clean.xlsm",
        container_id="IMP-901", po_43="4300008100", pedido_41="4100002400",
        author_b33="ELABORADO POR: \nJostin Lopez Sobalbarro",
        items=_default_clean_items(), vba_variant="closure_clean", stale_bultos_rows=15,
    )
    results["closures/closure_discrepancy_template.xlsm"] = build_closure_workbook(
        c_dir / "closure_discrepancy_template.xlsm",
        container_id="IMP-909", po_43="4300008100", pedido_41="4100002409",
        author_b33="ELABORADO POR: \nJostin Lopez Sobalbarro",
        items=_discrepancy_template_items(), vba_variant="closure_discrepancy", stale_bultos_rows=15,
    )
    results["closures/closure_over500_rows.xlsm"] = build_closure_workbook(
        c_dir / "closure_over500_rows.xlsm",
        container_id="IMP-187", po_43="4300001001-4300001002", pedido_41="4100002500",
        author_b33="ELABORADO POR: \nJostin Lopez Sobalbarro",
        items=_over500_items(), vba_variant="closure_over500",
        stale_bultos_rows=573, omit_dimension=True,
    )
    results["closures/closure_preexisting_errors.xlsm"] = build_closure_workbook(
        c_dir / "closure_preexisting_errors.xlsm",
        container_id="IMP-241", po_43="4300008300", pedido_41="4100002410",
        author_b33="ELABORADO POR: \nJostin Lopez Sobalbarro",
        items=_error_fixture_items(), vba_variant="closure_errors", error_mode="preexisting",
    )
    return results


def _build_trap_closure_fixtures(c_dir: Path) -> dict[str, dict[str, Any]]:
    """Build sin-control, missing-CIERRE-RECEP, legacy .xls, and lockfile traps."""
    results: dict[str, dict[str, Any]] = {}
    results["closures/IMP-184_sin_control_calidad.xlsm"] = build_closure_workbook(
        c_dir / "IMP-184_sin_control_calidad.xlsm",
        container_id="IMP-184", po_43="4300008400", pedido_41="4100002420",
        author_b33="ELABORADO POR: \nJostin Lopez Sobalbarro",
        items=_default_clean_items(), vba_variant="closure_imp184",
    )
    results["closures/closure_missing_cierre.xlsm"] = build_closure_missing_cierre(
        c_dir / "closure_missing_cierre.xlsm"
    )
    xls_path = c_dir / "SYNTH_LEGACY.xls"
    xls_bytes = build_legacy_xls_bytes("SYNTH_LEGACY_BIFF8_CLOSURE_TRAP")
    xls_path.write_bytes(xls_bytes)
    results["closures/SYNTH_LEGACY.xls"] = {
        "path": str(xls_path), "file_sha256": sha256_bytes(xls_bytes),
        "vba_sha256": None, "size_bytes": len(xls_bytes),
    }
    lock_path = write_lockfile(c_dir / "~$SYNTH_LOCK.xlsm", "ACR Lock Trap")
    lock_bytes = lock_path.read_bytes()
    results["closures/~$SYNTH_LOCK.xlsm"] = {
        "path": str(lock_path), "file_sha256": sha256_bytes(lock_bytes),
        "vba_sha256": None, "size_bytes": len(lock_bytes),
    }
    return results


def build_all_closure_fixtures(out_dir: Path) -> dict[str, dict[str, Any]]:
    """Build all master closure and non-ZIP trap fixtures under out_dir/closures/."""
    c_dir = out_dir / "closures"
    c_dir.mkdir(parents=True, exist_ok=True)
    results = _build_standard_closure_fixtures(c_dir)
    results.update(_build_trap_closure_fixtures(c_dir))
    return results
