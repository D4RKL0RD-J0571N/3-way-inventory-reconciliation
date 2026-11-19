"""Synthetic 5-sheet Daily Log (.xlsx) and lockfile builder."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from openxml_utils import finalize_workbook, sha256_bytes, write_lockfile

CATALOGOS_COLUMNS: list[tuple[str, list[str]]] = [
    ("1. Aduana y DUA", ["DUA y Docs Listos", "En Radar / Tránsito", "Pendiente Docs"]),
    ("2. Ingreso SAP / EWM", [
        "Listo y Notificado", "Pendiente Ingreso SAP",
        "En Proceso EWM/RFUI", "Sin Doc. 18... Aún",
    ]),
    ("3. Montaje a Ciegas", [
        "Listo (Solo Doc. Proveedor)", "Listo (Completo)",
        "En Cola A Ciegas", "Pendiente",
    ]),
    ("4. Control Calidad", [
        "CC Aprobado (ZDescarga)", "Pendiente Revisión Final CC", "Sin CC (Espera Muelle)",
    ]),
    ("5. Cierre Recepción (101)", [
        "Cuadrado (0 Diferencias)", "Pausa sin Control Calidad",
        "Pendiente Revisión CC", "Parcial / Pendiente Parte B",
        "Discrepancia / NCR (40...)", "Pendiente Cierre",
    ]),
    ("6. Cierre Despacho (641)", [
        "Cierre Despacho Listo", "Pre-Cableado 80% (Listo Mov 641)",
        "En Distribución Tiendas", "Bloqueado por Recepción",
    ]),
    ("Saludos", ["Buenos días", "Buenas tardes", "-"]),
    ("Estado Parte B", ["Pendiente Arribo", "En Muelle / Descarga", "Conciliado 0 Dif"]),
]

MAESTRO_HEADERS: list[str] = [
    "N°", "Contenedor (Ubic. EWM)", "Fecha Act.", "PO Orden Compra (430000...)",
    "Pedido Nacionaliz. (410000...)", "Entrega VL31N (1800...)",
    "Doc. Entrega EWM (Tarea Alm.)", "Unidad Manipulación HU (18 díg)",
    "1. Aduana y DUA", "2. Ingreso SAP / EWM", "3. Montaje a Ciegas",
    "4. Control Calidad (Muelle)", "5. Cierre Recepción (Mov 101)",
    "6. Cierre Despacho (Mov 641)", "Estado Maestro Actual (Auto)",
    "Observaciones y Contexto Operativo", "Saludo", "Notificación Automática SAP",
]


def _col_o_formula(r: int) -> str:
    """Return exact 7-state Estado Maestro Actual rollup formula for row r."""
    return (
        f'=IF(N{r}="Cierre Despacho Listo","7. Cierre Total (Recep + Desp)",'
        f'IF(M{r}="Cuadrado (0 Diferencias)","6. Recepción Cerrada (Listo Recep.)",'
        f'IF(ISNUMBER(SEARCH("Revisión",L{r})),"5. En Auditoría Control Calidad",'
        f'IF(M{r}="Pausa sin Control Calidad","4. Pausa sin CC (Montado a Ciegas)",'
        f'IF(J{r}="Listo y Notificado","3. Ingresado SAP (Espera Cierre)",'
        f'IF(J{r}="Pendiente Ingreso SAP","2. Doc SAP Listo (Pend. Ingreso)",'
        f'"1. En Cola Conteo a Ciegas"))))))'
    )


def _col_r_formula(r: int) -> str:
    """Return exact SAP intake notification formula for row r."""
    return (
        f'=IF(F{r}<>"-",Q{r}&", Se informa que el ingreso a nivel de sistema '
        f'del Contenedor "&SUBSTITUTE(B{r},"IMP-","")&", quedo listo.","-")'
    )


def _populate_dashboard_general(ws: Any) -> None:
    """Populate Dashboard General with stage names and KPI formulas."""
    ws["B2"] = "CONTROL DIARIO DE OPERACIONES — RECEPCION C001"
    ws["B8"] = '=COUNTIF(\'Maestro Contenedores\'!B6:B57,"IMP-*")'
    ws["J7"] = '=COUNTIF(\'Maestro Contenedores\'!L6:L57,"*Revisión*")'
    ws["J8"] = "IMP-259 e IMP-201-1"
    stages = [
        "1. Aduana, DUA y Radar",
        "2. Ingreso SAP / EWM (VL31N & RFUI)",
        "3. Montaje a Ciegas (Doc. Proveedor)",
        "4. Control Calidad Muelle (ZDescarga)",
        "5. Cierre Recepción 3-Way (Mov 101)",
        "6. Cierre Despacho Tiendas (Mov 641)",
    ]
    for idx, label in enumerate(stages):
        ws.cell(row=13 + idx, column=2, value=label)
    ws["B45"] = '(AC22*AD22) + SI(AE22=""; AF22; AE22*AF22)'


def _populate_catalogos(ws: Any) -> None:
    """Populate Catalogos!A1:H7 with exact Phase 0 dropdown lists."""
    for col_idx, (header, values) in enumerate(CATALOGOS_COLUMNS, start=1):
        ws.cell(row=1, column=col_idx, value=header)
        for row_offset, val in enumerate(values, start=2):
            ws.cell(row=row_offset, column=col_idx, value=val)


def _populate_parte_b(ws: Any) -> None:
    """Populate Control Parte B y Bultos with tblParteB (B5:H15)."""
    headers = [
        "Contenedor", "PO", "Pedido 41", "Bultos Pend.",
        "Estado Parte B", "Fecha Est.", "Notas",
    ]
    for idx, title in enumerate(headers):
        ws.cell(row=5, column=2 + idx, value=title)
    sample_cids = [
        "IMP-51", "IMP-125", "IMP-69", "IMP-289", "IMP-130",
        "IMP-290", "IMP-200", "IMP-126", "IMP-260", "IMP-314",
    ]
    for offset, cid in enumerate(sample_cids, start=6):
        ws.cell(row=offset, column=2, value=cid)
        ws.cell(row=offset, column=3, value=f"430000{7000 + offset}")
        ws.cell(row=offset, column=4, value=f"410000{2000 + offset}")
        ws.cell(row=offset, column=5, value=12)
        ws.cell(row=offset, column=6, value="Pendiente Arribo")
        ws.cell(row=offset, column=7, value="2026-11-25")
        ws.cell(row=offset, column=8, value="Sintetico Parte B")
    table = Table(displayName="tblParteB", ref="B5:H15")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)


def _populate_comunicados(ws: Any) -> None:
    """Populate Comunicados y Guia SAP with tblNotificaciones (B13:G19)."""
    ws["B8"] = "Asunto Sugerido:"
    ws["C8"] = "Ingreso Listo a Sistema - Contenedor {N}"
    ws["B9"] = "Cuerpo:"
    ws["C9"] = "{Saludo}\nSe informa que el ingreso a nivel de sistema del Contenedor {N}, quedo listo."
    headers = ["Prefijo", "Documento SAP", "Transaccion", "Modulo", "Etapa", "Notas"]
    for idx, title in enumerate(headers):
        ws.cell(row=13, column=2 + idx, value=title)
    rows = [
        ("43", "Orden de Compra", "ME80FN", "MM", "Etapa 1", "PO Principal"),
        ("41", "Pedido Nacionalizacion", "ME23N", "MM", "Etapa 1", "Invoicing / Nac"),
        ("18", "Entrega Entrante", "VL31N", "LE-SHP", "Etapa 2", "Inbound Delivery"),
        ("414", "Doc Entrega EWM", "/SCWM/MON", "EWM", "Etapa 2", "Warehouse Task"),
        ("80/112", "HU 18 Digitos", "RFUI", "EWM", "Etapa 2", "Handling Unit"),
        ("40", "Nota de Credito", "MIRO/NCR", "MM", "Etapa 5", "Reclamo Faltante"),
    ]
    for r_offset, row_vals in enumerate(rows, start=14):
        for c_offset, val in enumerate(row_vals, start=2):
            ws.cell(row=r_offset, column=c_offset, value=val)
    table = Table(displayName="tblNotificaciones", ref="B13:G19")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)


def _add_maestro_validations(ws: Any) -> None:
    """Attach Excel list DataValidations covering rows 6..57 from Catalogos."""
    dv_specs = [
        ("I6:I57", "=Catalogos!$A$2:$A$4"),
        ("J6:J57", "=Catalogos!$B$2:$B$5"),
        ("K6:K57", "=Catalogos!$C$2:$C$5"),
        ("L6:L57", "=Catalogos!$D$2:$D$4"),
        ("M6:M57", "=Catalogos!$E$2:$E$7"),
        ("N6:N57", "=Catalogos!$F$2:$F$5"),
        ("Q6:Q57", "=Catalogos!$G$2:$G$4"),
    ]
    for sqref, formula in dv_specs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(sqref)


def _build_synthetic_22_rows() -> list[dict[str, Any]]:
    """Build 22 representative container rows including IMP-202 duplicate & IMP-320/213."""
    presets = [
        ("IMP-202", "2026-09-26", "4300006692", "4100002331", "180060834", "41461", "112345678005808817",
         "DUA y Docs Listos", "Listo y Notificado", "Listo (Completo)", "CC Aprobado (ZDescarga)",
         "Cuadrado (0 Diferencias)", "Cierre Despacho Listo", "Buenos días", "Cerrado completo"),
        ("IMP-217", "2026-10-15", "4300007011", "4100002310", "-", "41462", "-",
         "DUA y Docs Listos", "Listo y Notificado", "Listo (Completo)", "CC Aprobado (ZDescarga)",
         "Cuadrado (0 Diferencias)", "Pre-Cableado 80% (Listo Mov 641)", "Buenos días", "Retroactivo sin 18/HU"),
        ("IMP-259", "2026-10-20", "4300007050", "4100002315", "180060850", "41465", "112345678005808850",
         "DUA y Docs Listos", "Listo y Notificado", "Listo (Solo Doc. Proveedor)", "Pendiente Revisión Final CC",
         "Pendiente Revisión CC", "Bloqueado por Recepción", "Buenas tardes", "En revision CC"),
    ]
    rows: list[dict[str, Any]] = []
    for idx in range(22):
        r_num = idx + 1
        if idx < len(presets):
            cid, dt, po, p41, ent, ewm, hu, s1, s2, s3, s4, s5, s6, sal, obs = presets[idx]
        elif idx == 15:
            cid, dt, po, p41, ent, ewm, hu = "IMP-202", "2026-11-10", "4300006692", "4100002332", "-", "41470", "-"
            s1, s2, s3, s4, s5, s6, sal, obs = (
                "DUA y Docs Listos", "Listo y Notificado", "Listo (Solo Doc. Proveedor)",
                "Sin CC (Espera Muelle)", "Pausa sin Control Calidad", "Bloqueado por Recepción",
                "Buenos días", "Duplicado IMP-202 trampa",
            )
        elif idx in (17, 18):
            cid = "IMP-320" if idx == 17 else "IMP-213"
            dt, po, p41, ent, ewm, hu = "2026-11-11", "4300007748", "4100002329", f"1800609{idx}", f"414{70 + idx}", f"1123456780058090{idx}"
            s1, s2, s3, s4, s5, s6, sal, obs = (
                "DUA y Docs Listos", "Listo y Notificado", "Listo (Completo)",
                "CC Aprobado (ZDescarga)", "Cuadrado (0 Diferencias)", "En Distribución Tiendas",
                "Buenas tardes", "PO compartido 4300007748",
            )
        else:
            cid = f"IMP-{800 + idx}"
            dt, po, p41, ent, ewm, hu = "2026-11-12", f"4300008{idx:03d}", f"4100002{400 + idx}", f"1800609{idx:02d}", f"414{60 + idx}", f"112345678005809{idx:03d}"
            s1, s2, s3, s4, s5, s6, sal, obs = (
                "DUA y Docs Listos", "Pendiente Ingreso SAP", "En Cola A Ciegas",
                "Sin CC (Espera Muelle)", "Pendiente Cierre", "Bloqueado por Recepción",
                "Buenos días", f"Fila limpia {r_num}",
            )
        rows.append({
            "n": r_num, "cid": cid, "date": dt, "po": po, "p41": p41, "ent": ent, "ewm": ewm, "hu": hu,
            "s1": s1, "s2": s2, "s3": s3, "s4": s4, "s5": s5, "s6": s6, "sal": sal, "obs": obs,
        })
    return rows


def _build_52_rows_for_57_boundary() -> list[dict[str, Any]]:
    """Build 52 data rows (Rows 6..57) for Row-58 table/validation resize test."""
    rows: list[dict[str, Any]] = []
    for idx in range(52):
        rows.append({
            "n": idx + 1, "cid": f"IMP-{700 + idx}", "date": "2026-11-15",
            "po": f"4300009{idx:03d}", "p41": f"4100003{idx:03d}",
            "ent": f"180070{idx:03d}", "ewm": f"415{idx:02d}",
            "hu": f"1123456780059{idx:05d}",
            "s1": "DUA y Docs Listos", "s2": "Listo y Notificado",
            "s3": "Listo (Completo)", "s4": "CC Aprobado (ZDescarga)",
            "s5": "Cuadrado (0 Diferencias)", "s6": "Cierre Despacho Listo",
            "sal": "Buenos días", "obs": f"Boundary row {6 + idx}",
        })
    return rows


def _populate_maestro(ws: Any, rows: list[dict[str, Any]]) -> None:
    """Write Maestro Contenedores headers, rows, formulas, table, and validations."""
    for c_idx, title in enumerate(MAESTRO_HEADERS, start=1):
        ws.cell(row=5, column=c_idx, value=title)
    for offset, row in enumerate(rows):
        r = 6 + offset
        vals = [
            row["n"], row["cid"], row["date"], row["po"], row["p41"], row["ent"],
            row["ewm"], row["hu"], row["s1"], row["s2"], row["s3"], row["s4"],
            row["s5"], row["s6"], _col_o_formula(r), row["obs"], row["sal"], _col_r_formula(r),
        ]
        for c_idx, val in enumerate(vals, start=1):
            ws.cell(row=r, column=c_idx, value=val)
    end_row = 5 + len(rows)
    table = Table(displayName="tblMaestroContenedores", ref=f"A5:R{end_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    _add_maestro_validations(ws)


def _create_daily_log_file(output_path: Path, rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Assemble and finalize a 5-sheet Daily Log .xlsx workbook."""
    wb = openpyxl.Workbook()
    ws_dash = wb.active
    ws_dash.title = "Dashboard General"
    _populate_dashboard_general(ws_dash)
    ws_maestro = wb.create_sheet("Maestro Contenedores")
    _populate_maestro(ws_maestro, rows)
    ws_pb = wb.create_sheet("Control Parte B y Bultos")
    _populate_parte_b(ws_pb)
    ws_com = wb.create_sheet("Comunicados y Guia SAP")
    _populate_comunicados(ws_com)
    ws_cat = wb.create_sheet("Catalogos")
    _populate_catalogos(ws_cat)
    return finalize_workbook(
        wb,
        output_path,
        macro_enabled=False,
        use_shared_strings=True,
        core_creator="Angie Montero Rojas",
        core_last_modified_by="Angie Montero Rojas",
        core_created="2026-11-14T18:04:02Z",
    )


def build_all_daily_log_fixtures(out_dir: Path) -> dict[str, dict[str, Any]]:
    """Build Daily Log synthetic workbooks and lockfile under out_dir/daily_log/."""
    dl_dir = out_dir / "daily_log"
    dl_dir.mkdir(parents=True, exist_ok=True)
    results: dict[str, dict[str, Any]] = {}
    results["daily_log/Control_Diario_Synthetic.xlsx"] = _create_daily_log_file(
        dl_dir / "Control_Diario_Synthetic.xlsx",
        _build_synthetic_22_rows(),
    )
    results["daily_log/Control_Diario_57Rows.xlsx"] = _create_daily_log_file(
        dl_dir / "Control_Diario_57Rows.xlsx",
        _build_52_rows_for_57_boundary(),
    )
    lock_path = write_lockfile(dl_dir / "~$Control_Diario_Synthetic.xlsx", "Jostin OneDrive")
    lock_bytes = lock_path.read_bytes()
    results["daily_log/~$Control_Diario_Synthetic.xlsx"] = {
        "path": str(lock_path),
        "file_sha256": sha256_bytes(lock_bytes),
        "vba_sha256": None,
        "size_bytes": len(lock_bytes),
    }
    return results
