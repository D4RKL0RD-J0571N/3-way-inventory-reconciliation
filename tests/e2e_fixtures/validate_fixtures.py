"""Independent structural and semantic validator for generated ACR E2E fixtures."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
import zipfile

from bs4 import BeautifulSoup
from docx import Document
from lxml import etree
import openpyxl
from pypdf import PdfReader

FIXTURES_ROOT = Path(__file__).resolve().parent
REPO_ROOT = FIXTURES_ROOT.parent.parent
BUILDERS_DIR = FIXTURES_ROOT / "builders"
for path_entry in (str(REPO_ROOT), str(BUILDERS_DIR)):
    if path_entry not in sys.path:
        sys.path.insert(0, path_entry)

from closure_builder import CIERRE_RECEP_HEADERS  # noqa: E402
from daily_log_builder import CATALOGOS_COLUMNS  # noqa: E402
from openxml_utils import MAIN_NS, OLE2_MAGIC, map_sheet_paths, sha256_bytes  # noqa: E402
from src.discrepancy_classifier import (  # noqa: E402
    DeclaredItem,
    DiscrepancyType,
    PhysicalItem,
    compare_merchandise,
)
from src.qc_reader import read_qc_inspection_file  # noqa: E402


def _validate_manifest_presence(
    base_dir: Path, manifest: dict[str, dict]
) -> list[str]:
    """Check that all manifest entries exist on disk."""
    errors: list[str] = []
    for rel_path in manifest:
        if not (base_dir / rel_path).exists():
            errors.append(f"Missing fixture file: {rel_path}")
    return errors


def _validate_zip_and_vba(
    file_path: Path, rel_path: str, spec: dict
) -> list[str]:
    """Validate OpenXML ZIP integrity, sharedStrings encoding, and vbaProject.bin SHA-256."""
    errors: list[str] = []
    if not zipfile.is_zipfile(file_path):
        return [f"Expected valid ZIP OpenXML package, got non-ZIP: {rel_path}"]
    with zipfile.ZipFile(file_path, "r") as zf:
        bad_member = zf.testzip()
        if bad_member is not None:
            errors.append(f"Corrupt ZIP entry {bad_member} in {rel_path}")
        names = set(zf.namelist())
        expected_vba = spec.get("vba_sha256")
        if expected_vba:
            if "xl/vbaProject.bin" not in names:
                errors.append(f"Missing xl/vbaProject.bin in {rel_path}")
            else:
                vba_bytes = zf.read("xl/vbaProject.bin")
                if not vba_bytes.startswith(OLE2_MAGIC):
                    errors.append(f"vbaProject.bin lacks OLE2 header in {rel_path}")
                actual_sha = sha256_bytes(vba_bytes)
                if actual_sha != expected_vba:
                    errors.append(
                        f"vbaProject.bin SHA mismatch in {rel_path}: {actual_sha} != {expected_vba}"
                    )
        if spec.get("shared_strings") is True and "xl/sharedStrings.xml" not in names:
            errors.append(f"Missing xl/sharedStrings.xml in {rel_path}")
        if spec.get("shared_strings") is False and "xl/sharedStrings.xml" in names:
            errors.append(f"Unexpected xl/sharedStrings.xml in inlineStr variant {rel_path}")
    return errors


def _validate_qc_fixtures(base_dir: Path, manifest: dict[str, dict]) -> list[str]:
    """Validate all QC fixtures via OpenXML inspection and Python QcReader."""
    errors: list[str] = []
    for rel_path, spec in manifest.items():
        if spec.get("category") != "qc":
            continue
        fpath = base_dir / rel_path
        errors.extend(_validate_zip_and_vba(fpath, rel_path, spec))
        report = read_qc_inspection_file(str(fpath))
        if len(report.items) != spec["expected_items"]:
            errors.append(f"{rel_path}: items={len(report.items)} != {spec['expected_items']}")
        if report.total_calculated_units != spec["expected_calculated_units"]:
            errors.append(
                f"{rel_path}: calc_u={report.total_calculated_units} != {spec['expected_calculated_units']}"
            )
        if report.total_calculated_cartons != spec["expected_calculated_cartons"]:
            errors.append(
                f"{rel_path}: calc_c={report.total_calculated_cartons} != {spec['expected_calculated_cartons']}"
            )
        found_flags = {f for it in report.flagged_items for f in it.calculation_result.flags}
        for exp_flag in spec.get("expected_flags", []):
            if exp_flag not in found_flags:
                errors.append(f"{rel_path}: missing expected flag {exp_flag}")
    errors.extend(_validate_qc_special_xml(base_dir))
    return errors


def _validate_qc_special_xml(base_dir: Path) -> list[str]:
    """Check bloat row count and inlineStr encoding on specialized QC fixtures."""
    errors: list[str] = []
    bloat_path = base_dir / "qc/qc_bloat_streaming.xlsm"
    with zipfile.ZipFile(bloat_path, "r") as zf:
        sheet_xml = etree.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        row_count = len(sheet_xml.findall(f".//{{{MAIN_NS}}}row"))
        if row_count < 1000:
            errors.append(f"qc_bloat_streaming.xlsm has only {row_count} <row> elements.")
    inline_path = base_dir / "qc/qc_inlinestr.xlsm"
    with zipfile.ZipFile(inline_path, "r") as zf:
        sheet_xml = etree.fromstring(zf.read("xl/worksheets/sheet1.xml"))
        inline_cells = [
            c for c in sheet_xml.findall(f".//{{{MAIN_NS}}}c") if c.get("t") == "inlineStr"
        ]
        if not inline_cells:
            errors.append("qc_inlinestr.xlsm has zero t='inlineStr' cells.")
    return errors


def _validate_cierre_headers(ws: openpyxl.worksheet.worksheet.Worksheet, rel_path: str) -> list[str]:
    """Verify exact 58-column Row-12 CIERRE RECEP schema."""
    errors: list[str] = []
    for col_idx, expected_title in CIERRE_RECEP_HEADERS.items():
        actual = ws.cell(row=12, column=col_idx).value
        if actual != expected_title:
            errors.append(f"{rel_path}: Row 12 col {col_idx} = {actual!r} != {expected_title!r}")
    return errors


def _validate_closure_fixtures(base_dir: Path, manifest: dict[str, dict]) -> list[str]:
    """Validate master closure workbooks, over-500 rows, pre-existing errors, and traps."""
    errors: list[str] = []
    for rel_path, spec in manifest.items():
        if spec.get("category") != "closures":
            continue
        fpath = base_dir / rel_path
        if spec.get("is_zip") is False:
            errors.extend(_validate_non_zip_trap(fpath, rel_path, spec))
            continue
        errors.extend(_validate_zip_and_vba(fpath, rel_path, spec))
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=False)
        try:
            if spec.get("trap") == "missing_cierre_recep":
                errors.extend(_validate_missing_cierre(wb.sheetnames, rel_path, spec))
            elif "CIERRE RECEP" not in wb.sheetnames:
                errors.append(f"{rel_path}: missing CIERRE RECEP sheet.")
            else:
                wb_ro = openpyxl.load_workbook(fpath, read_only=False, data_only=False)
                errors.extend(_validate_cierre_headers(wb_ro["CIERRE RECEP"], rel_path))
                wb_ro.close()
        finally:
            wb.close()
    errors.extend(_validate_closure_special_cases(base_dir))
    return errors


def _validate_missing_cierre(sheetnames: list[str], rel_path: str, spec: dict) -> list[str]:
    """Validate closure_missing_cierre.xlsm omits CIERRE RECEP but retains support tabs."""
    errors: list[str] = []
    if "CIERRE RECEP" in sheetnames:
        errors.append(f"{rel_path}: must omit CIERRE RECEP sheet.")
    for expected in spec.get("expected_sheets", []):
        if expected not in sheetnames:
            errors.append(f"{rel_path}: missing expected sheet {expected!r}.")
    return errors


def _validate_closure_special_cases(base_dir: Path) -> list[str]:
    """Verify over500 omitted dimension/duplicates, preexisting errors, sin_control, and discrepancy pair."""
    errors: list[str] = []
    over500 = base_dir / "closures/closure_over500_rows.xlsm"
    with zipfile.ZipFile(over500, "r") as zf:
        entries = {n: zf.read(n) for n in zf.namelist()}
        smap = map_sheet_paths(entries)
        cr_xml = etree.fromstring(entries[smap["CIERRE RECEP"]])
        if cr_xml.find(f"{{{MAIN_NS}}}dimension") is not None:
            errors.append("closure_over500_rows.xlsm should omit <dimension> on CIERRE RECEP.")
    wb500 = openpyxl.load_workbook(over500, data_only=False)
    try:
        ws = wb500["CIERRE RECEP"]
        if str(ws["E15"].value) != "1455103" or str(ws["E350"].value) != "1455103":
            errors.append("closure_over500_rows.xlsm missing duplicate SKU 1455103 at rows 15 & 350.")
        if ws["E560"].value is None:
            errors.append("closure_over500_rows.xlsm does not extend to row 560.")
    finally:
        wb500.close()
    errors.extend(_validate_preexisting_and_sin_control(base_dir))
    errors.extend(_validate_discrepancy_pair(base_dir))
    return errors


def _validate_preexisting_and_sin_control(base_dir: Path) -> list[str]:
    """Verify preexisting formula error cells and sin_control sheet absence."""
    errors: list[str] = []
    err_wb = openpyxl.load_workbook(base_dir / "closures/closure_preexisting_errors.xlsm", data_only=True)
    try:
        ws_e = err_wb["CIERRE RECEP"]
        checks = {"B5": "#REF!", "V14": "#VALUE!", "AY29": "#N/A", "AT16": "#N/A", "AS17": "#REF!"}
        for coord, expected in checks.items():
            if ws_e[coord].value != expected:
                errors.append(f"closure_preexisting_errors.xlsm {coord}={ws_e[coord].value!r} != {expected!r}")
    finally:
        err_wb.close()
    wb_184 = openpyxl.load_workbook(base_dir / "closures/IMP-184_sin_control_calidad.xlsm", read_only=True)
    try:
        if "CONTROL DATA" in wb_184.sheetnames:
            errors.append("IMP-184_sin_control_calidad.xlsm must not contain CONTROL DATA.")
    finally:
        wb_184.close()
    return errors


def _validate_discrepancy_pair(base_dir: Path) -> list[str]:
    """Verify matched Col E reconciliation between qc_reconciliation_discrepancies and closure_discrepancy_template."""
    errors: list[str] = []
    wb_c = openpyxl.load_workbook(base_dir / "closures/closure_discrepancy_template.xlsm", data_only=False)
    try:
        ws = wb_c["CIERRE RECEP"]
        declared = [
            DeclaredItem(
                material_code=str(ws.cell(row=r, column=5).value),
                vendor_material=str(ws.cell(row=r, column=4).value or ""),
                category=str(ws.cell(row=r, column=40).value or "")[:2],
                declared_units=int(ws.cell(row=r, column=50).value or 0),
                declared_cartons=int(ws.cell(row=r, column=49).value or 0),
                uom=str(ws.cell(row=r, column=55).value or "UN"),
            )
            for r in range(13, 18)
        ]
    finally:
        wb_c.close()
    qc_rep = read_qc_inspection_file(str(base_dir / "qc/qc_reconciliation_discrepancies.xlsm"))
    physical = [
        PhysicalItem(
            material_code=it.material_code,
            vendor_material=it.vendor_material,
            category=it.category,
            physical_units=it.calculation_result.calculated_units,
            physical_cartons=it.calculation_result.calculated_cartons,
            uom=it.uom,
        )
        for it in qc_rep.items if not it.is_child_sku and it.calculation_result is not None
    ]
    summary = compare_merchandise(declared, physical)
    if (summary.total_shortage_units, summary.total_overage_units, summary.revuelto_pairs_count) != (60, 48, 1):
        errors.append(
            f"Discrepancy pair mismatch: shortage={summary.total_shortage_units}, "
            f"overage={summary.total_overage_units}, revuelto={summary.revuelto_pairs_count}"
        )
    by_code = {d.material_code: d.discrepancy_type for d in summary.items}
    expected_types = {
        "1455103": DiscrepancyType.EXACT_MATCH,
        "5033284": DiscrepancyType.SHORTAGE,
        "1455104": DiscrepancyType.OVERAGE,
        "1455105": DiscrepancyType.REVUELTO_CANDIDATE,
        "1455106": DiscrepancyType.REVUELTO_CANDIDATE,
    }
    if by_code != expected_types:
        errors.append(f"Discrepancy pair classifications mismatch: {by_code!r}")
    return errors


def _validate_non_zip_trap(fpath: Path, rel_path: str, spec: dict) -> list[str]:
    """Validate OLE2 .xls and 165-byte ~$ lockfile non-ZIP traps."""
    errors: list[str] = []
    if zipfile.is_zipfile(fpath):
        errors.append(f"Trap file must NOT be a ZIP archive: {rel_path}")
    raw = fpath.read_bytes()
    exp_size = spec.get("expected_size_bytes")
    if exp_size is not None and len(raw) != exp_size:
        errors.append(f"{rel_path}: size {len(raw)} != {exp_size}")
    if rel_path.endswith(".xls") and not raw.startswith(OLE2_MAGIC):
        errors.append(f"{rel_path}: missing OLE2 BIFF magic header.")
    return errors


def _validate_historical_fixtures(base_dir: Path, manifest: dict[str, dict]) -> list[str]:
    """Validate historical sync authorship, IMP-209 sharedStrings trap, and shared POs."""
    errors: list[str] = []
    for rel_path, spec in manifest.items():
        if spec.get("category") != "historical_sync":
            continue
        fpath = base_dir / rel_path
        if spec.get("is_zip") is False:
            errors.extend(_validate_non_zip_trap(fpath, rel_path, spec))
            continue
        errors.extend(_validate_zip_and_vba(fpath, rel_path, spec))
        wb = openpyxl.load_workbook(fpath, data_only=True)
        try:
            dg = wb["DATOS GENERALES"]
            b33 = str(dg["B33"].value or "")
            exp_author = spec.get("author_b33")
            if exp_author and exp_author not in b33:
                errors.append(f"{rel_path}: DG!B33={b33!r} missing {exp_author!r}")
            if str(dg["C7"].value or "") != spec.get("container_id"):
                errors.append(f"{rel_path}: DG!C7={dg['C7'].value!r} != {spec.get('container_id')!r}")
            if str(dg["C9"].value or "") != spec.get("pedido_41"):
                errors.append(f"{rel_path}: DG!C9={dg['C9'].value!r} != {spec.get('pedido_41')!r}")
            if rel_path.endswith("IMP218_Hybrid.xlsm"):
                if "CONTROL DATA" not in wb.sheetnames or dg["C18"].value != 0:
                    errors.append("IMP218_Hybrid.xlsm missing CONTROL DATA or DG!C18=0.")
        finally:
            wb.close()
    imp209 = base_dir / "historical_sync/IMP209_SharedStringsTrap.xlsm"
    with zipfile.ZipFile(imp209, "r") as zf:
        ss_text = zf.read("xl/sharedStrings.xml").decode("utf-8")
        if "Jostin Lopez Sobalbarro" not in ss_text:
            errors.append("IMP209_SharedStringsTrap.xlsm missing planted Jostin decoy in sharedStrings.xml.")
    return errors


def _validate_daily_log_fixtures(base_dir: Path, manifest: dict[str, dict]) -> list[str]:
    """Validate 5-sheet Daily Log schema, Catalogos!A1:H7, tables, and formulas."""
    errors: list[str] = []
    for rel_path, spec in manifest.items():
        if spec.get("category") != "daily_log":
            continue
        fpath = base_dir / rel_path
        if spec.get("is_zip") is False:
            errors.extend(_validate_non_zip_trap(fpath, rel_path, spec))
            continue
        errors.extend(_validate_zip_and_vba(fpath, rel_path, spec))
        wb = openpyxl.load_workbook(fpath, data_only=False)
        try:
            exp_sheets = manifest["daily_log/Control_Diario_Synthetic.xlsx"]["expected_sheets"]
            if wb.sheetnames != exp_sheets:
                errors.append(f"{rel_path}: sheets={wb.sheetnames} != {exp_sheets}")
            ws_cat = wb["Catalogos"]
            for col_idx, (hdr, vals) in enumerate(CATALOGOS_COLUMNS, start=1):
                if ws_cat.cell(row=1, column=col_idx).value != hdr:
                    errors.append(f"{rel_path}: Catalogos col {col_idx} header mismatch.")
                for off, expected_v in enumerate(vals, start=2):
                    if ws_cat.cell(row=off, column=col_idx).value != expected_v:
                        errors.append(f"{rel_path}: Catalogos!R{off}C{col_idx} mismatch.")
            ws_m = wb["Maestro Contenedores"]
            if "tblMaestroContenedores" not in ws_m.tables:
                errors.append(f"{rel_path}: missing tblMaestroContenedores table.")
            if not str(ws_m["O6"].value or "").startswith("=IF("):
                errors.append(f"{rel_path}: missing Col O rollup formula at O6.")
            if not str(ws_m["R6"].value or "").startswith("=IF("):
                errors.append(f"{rel_path}: missing Col R notification formula at R6.")
            if rel_path.endswith("Control_Diario_57Rows.xlsx") and ws_m["B57"].value is None:
                errors.append("Control_Diario_57Rows.xlsx missing row 57 data.")
            if rel_path.endswith("Control_Diario_Synthetic.xlsx"):
                cids = [ws_m.cell(row=r, column=2).value for r in range(6, 28)]
                if cids.count("IMP-202") != 2:
                    errors.append(f"Control_Diario_Synthetic.xlsx expected 2 IMP-202 rows, got {cids.count('IMP-202')}.")
                b45 = wb["Dashboard General"]["B45"].value
                if b45 != '(AC22*AD22) + SI(AE22=""; AF22; AE22*AF22)':
                    errors.append(f"Control_Diario_Synthetic.xlsx Dashboard General!B45 mismatch: {b45!r}")
        finally:
            wb.close()
    return errors


def _validate_sap_docs_fixtures(base_dir: Path) -> list[str]:
    """Validate PDF, scan PDF, DOCX, XLSX (Hoja1 PII isolation), TXT, MD, HTML, and corrupt files."""
    errors: list[str] = []
    guide = PdfReader(str(base_dir / "sap_docs/SYNTH_SAP_GUIDE.pdf"))
    if len(guide.pages) != 2 or "ME80FN" not in guide.pages[0].extract_text():
        errors.append("SYNTH_SAP_GUIDE.pdf missing Page 1 ME80FN content.")
    if "MB51" not in guide.pages[1].extract_text() or "SAP GUI 7.00" not in guide.pages[1].extract_text():
        errors.append("SYNTH_SAP_GUIDE.pdf missing Page 2 MB51 / SAP GUI 7.00 caveat.")
    scan = PdfReader(str(base_dir / "sap_docs/SYNTH_SCAN_NO_TEXT.pdf"))
    if len(scan.pages) != 1 or scan.pages[0].extract_text().strip() != "":
        errors.append("SYNTH_SCAN_NO_TEXT.pdf must have 1 page and 0 extracted text.")
    doc = Document(str(base_dir / "sap_docs/SYNTH_SOP.docx"))
    doc_text = "\n".join(p.text for p in doc.paragraphs)
    for heading in ("1. Recepción DUA", "2.2 Prepacks (MAST y STPO)", "3. ZDescarga y Control Calidad", "4. VKP2 Precio Venta (No Costo)"):
        if heading not in doc_text:
            errors.append(f"SYNTH_SOP.docx missing heading {heading!r}")
    errors.extend(_validate_sap_xlsx_and_text(base_dir))
    return errors


def _validate_sap_xlsx_and_text(base_dir: Path) -> list[str]:
    """Validate Hoja1 PII isolation, TXT line numbers, MD headings, HTML sections, and corrupt/."""
    errors: list[str] = []
    wb_pii = openpyxl.load_workbook(base_dir / "sap_docs/SYNTH_CLOSURE_WITH_HOJA1_PII.xlsx", data_only=True)
    try:
        ws_h1 = wb_pii["Hoja1"]
        h1_dump = " ".join(str(c) for row in ws_h1.iter_rows(values_only=True) for c in row if c is not None)
        for token in ("Autorización de Deducción", "MV :", "998", "Juan Pérez Test", "1-0999-0888", "SYNTHETIC_PII_MV998_SENTINEL_DO_NOT_INDEX"):
            if token not in h1_dump:
                errors.append(f"SYNTH_CLOSURE_WITH_HOJA1_PII.xlsx Hoja1 missing {token!r}")
        cr_dump = " ".join(str(c) for row in wb_pii["CIERRE RECEP"].iter_rows(values_only=True) for c in row if c is not None)
        if "SYNTHETIC_PII_MV998_SENTINEL_DO_NOT_INDEX" in cr_dump:
            errors.append("PII sentinel leaked into CIERRE RECEP tab.")
    finally:
        wb_pii.close()
    notes_lines = (base_dir / "sap_docs/SYNTH_NOTES.txt").read_text(encoding="utf-8").splitlines()
    if len(notes_lines) < 50 or "RFUI" not in notes_lines[24] or "VL3IN" not in notes_lines[41]:
        errors.append("SYNTH_NOTES.txt failed line count or line 25/42 content check.")
    soup = BeautifulSoup((base_dir / "sap_docs/SYNTH_EXPORT_GUIDE.html").read_text(encoding="utf-8"), "html.parser")
    for sec_id in ("vl31n", "scwm-mon", "mb52-mb5t"):
        if soup.find("section", id=sec_id) is None:
            errors.append(f"SYNTH_EXPORT_GUIDE.html missing <section id='{sec_id}'>.")
    if (base_dir / "sap_docs/corrupt/empty.txt").stat().st_size != 0:
        errors.append("corrupt/empty.txt must be 0 bytes.")
    return errors


def validate_directory(base_dir: Path, manifest_path: Path | None = None) -> list[str]:
    """Run all structural and semantic checks against materialized fixtures in base_dir."""
    m_path = manifest_path or (FIXTURES_ROOT / "fixture_manifest.json")
    manifest_data = json.loads(m_path.read_text(encoding="utf-8"))
    fixtures: dict[str, dict] = manifest_data["fixtures"]
    errors = _validate_manifest_presence(base_dir, fixtures)
    if errors:
        return errors
    errors.extend(_validate_qc_fixtures(base_dir, fixtures))
    errors.extend(_validate_closure_fixtures(base_dir, fixtures))
    errors.extend(_validate_historical_fixtures(base_dir, fixtures))
    errors.extend(_validate_daily_log_fixtures(base_dir, fixtures))
    errors.extend(_validate_sap_docs_fixtures(base_dir))
    return errors


def main(argv: list[str] | None = None) -> int:
    """CLI entry point for independent E2E fixture validation."""
    parser = argparse.ArgumentParser(description="Validate generated ACR E2E fixtures.")
    parser.add_argument("--dir", type=Path, required=True, help="Directory with generated fixtures.")
    parser.add_argument(
        "--manifest",
        type=Path,
        default=FIXTURES_ROOT / "fixture_manifest.json",
        help="Path to fixture_manifest.json.",
    )
    args = parser.parse_args(argv)
    failures = validate_directory(args.dir, manifest_path=args.manifest)
    if failures:
        for msg in failures:
            print(f"[FAIL] {msg}")
        return 1
    print(f"All fixtures in {args.dir} passed independent structural and semantic validation.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
