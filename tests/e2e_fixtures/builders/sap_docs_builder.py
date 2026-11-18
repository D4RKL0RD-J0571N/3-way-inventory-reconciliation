"""Synthetic SAP documentation and corrupt-file boundary fixture builder."""

from __future__ import annotations

import io
from pathlib import Path
from typing import Any

from docx import Document
import fitz
import openpyxl
from PIL import Image, ImageDraw

from openxml_utils import finalize_workbook, normalize_docx_file, sha256_bytes

PDF_METADATA: dict[str, str] = {
    "creationDate": "D:20261118120000Z",
    "modDate": "D:20261118120000Z",
    "producer": "ACR Synthetic Fixture Generator",
    "creator": "ACR E2E",
    "title": "SAP Operational Reference Guide (Synthetic)",
}


def _build_sap_guide_pdf(path: Path) -> dict[str, Any]:
    """Create multi-page digital PDF (Page 1 ME80FN, Page 2 MB51 + GUI 7.00 caveat)."""
    doc = fitz.open()
    page1 = doc.new_page()
    page1.insert_text(
        (72, 72),
        (
            "Guia Operativa SAP C001 - Pagina 1\n\n"
            "1. ME80FN - Analisis General de Pedidos (PO 43 y Pedido 41)\n"
            "- Seleccionar vista de repartos / posiciones con Shift+F4.\n"
            "- Verificar UMP (PRK vs UN) y precio neto EKPO-NETPR.\n"
            "- Centro receptor: C001 (IMPORTADORA ALMACEMA, S.A.).\n"
        ),
        fontsize=11,
    )
    page2 = doc.new_page()
    page2.insert_text(
        (72, 72),
        (
            "Guia Operativa SAP C001 - Pagina 2\n\n"
            "2. MB51 - Listado de Documentos de Material\n"
            "- Filtrar centro C001 y clases de movimiento 101-102 (Ingreso/Anulacion).\n"
            "- Verificar conciliacion fisica contra CIERRE RECEP.\n\n"
            "Advertencia de Entorno SAP GUI 7.00:\n"
            "SAP GUI 7.00 es solo el cliente de presentacion y NO identifica la version\n"
            "del backend SAP (ECC 6.0 vs S/4HANA). Confirmar en Sistema -> Estatus.\n"
        ),
        fontsize=11,
    )
    doc.set_metadata(PDF_METADATA)
    pdf_bytes = doc.tobytes(garbage=4, deflate=True, no_new_id=True)
    doc.close()
    path.write_bytes(pdf_bytes)
    return {"path": str(path), "file_sha256": sha256_bytes(pdf_bytes), "size_bytes": len(pdf_bytes)}


def _build_scan_no_text_pdf(path: Path) -> dict[str, Any]:
    """Create 1-page image-only PDF with zero extractable text characters."""
    img = Image.new("RGB", (400, 200), (245, 245, 245))
    draw = ImageDraw.Draw(img)
    draw.rectangle([20, 20, 380, 180], outline=(80, 80, 80), width=2)
    draw.line([(40, 60), (360, 60)], fill=(120, 120, 120), width=3)
    draw.line([(40, 100), (300, 100)], fill=(120, 120, 120), width=3)
    png_buf = io.BytesIO()
    img.save(png_buf, format="PNG")
    doc = fitz.open()
    page = doc.new_page()
    page.insert_image(fitz.Rect(72, 72, 472, 272), stream=png_buf.getvalue())
    doc.set_metadata(PDF_METADATA)
    pdf_bytes = doc.tobytes(garbage=4, deflate=True, no_new_id=True)
    doc.close()
    path.write_bytes(pdf_bytes)
    return {"path": str(path), "file_sha256": sha256_bytes(pdf_bytes), "size_bytes": len(pdf_bytes)}


def _build_sop_docx(path: Path) -> dict[str, Any]:
    """Create structured DOCX SOP normalized deterministically."""
    doc = Document()
    doc.add_heading("SOP-LOG-C001-SYNTH: Procedimiento Recepcion 3-Way", level=0)
    doc.add_heading("1. Recepción DUA", level=1)
    doc.add_paragraph("Validar Radar semanal, DUA sellado y correo __PEDIDO NACIONALIZACION__.")
    doc.add_heading("2.2 Prepacks (MAST y STPO)", level=2)
    doc.add_paragraph(
        "Consultar las tablas de base de datos MAST y STPO mediante SE16N (solo lectura ACTVT=03) "
        "para obtener STLNR, IDNRK y MENGE de la curva de tallas."
    )
    doc.add_heading("3. ZDescarga y Control Calidad", level=1)
    doc.add_paragraph(
        "ZDescarga es una transaccion Z especifica de la empresa para registrar "
        "bultos enteros, complementos y observaciones en muelle."
    )
    doc.add_heading("4. VKP2 Precio Venta (No Costo)", level=1)
    doc.add_paragraph(
        "ADVERTENCIA CRITICA: VKP2 muestra unicamente precio de venta al detalle en CRC. "
        "NUNCA utilizar VKP2 como costo de importacion (usar EKPO-NETPR en ME80FN)."
    )
    buf = io.BytesIO()
    doc.save(buf)
    return normalize_docx_file(buf.getvalue(), path, creator="ACR Synthetic SOP")


def _build_data_dict_xlsx(path: Path) -> dict[str, Any]:
    """Create SAP_Campos and Prefijos multi-sheet XLSX reference."""
    wb = openpyxl.Workbook()
    ws_campos = wb.active
    ws_campos.title = "SAP_Campos"
    ws_campos.append(["Objeto", "Tipo", "CampoClave", "Uso Operativo", "Centro"])
    rows = [
        ("EINA", "Tabla", "MATNR / LIFNR / IDNLF / LMEIN", "Info record proveedor", "C001"),
        ("MAST", "Tabla", "MATNR / STLNR", "Enlace material a lista de materiales", "C001"),
        ("STPO", "Tabla", "STLNR / IDNRK / MENGE", "Componentes y curva prepack", "C001"),
        ("EKPO-NETPR", "Campo", "EBELN / EBELP / NETPR", "Precio neto real orden compra", "C001"),
    ]
    for r in rows:
        ws_campos.append(list(r))
    ws_pref = wb.create_sheet("Prefijos")
    ws_pref.append(["Prefijo", "Documento", "Transaccion", "Etapa"])
    for p in [("43", "PO Orden Compra", "ME80FN", "Etapa 1"),
              ("41", "Pedido Nacionalizacion", "ME23N", "Etapa 1"),
              ("45", "Pedido Tiendas", "ME23N", "Etapa 6"),
              ("40", "Nota de Credito", "MIRO", "Etapa 5"),
              ("18", "Entrega Entrante", "VL31N", "Etapa 2"),
              ("80/112", "HU 18 Digitos", "RFUI", "Etapa 2")]:
        ws_pref.append(list(p))
    return finalize_workbook(wb, path, macro_enabled=False, use_shared_strings=True)


def _build_hoja1_pii_xlsx(path: Path) -> dict[str, Any]:
    """Create benign CIERRE RECEP + Hoja1 MV 998 synthetic PII exclusion trap."""
    wb = openpyxl.Workbook()
    ws_cr = wb.active
    ws_cr.title = "CIERRE RECEP"
    ws_cr["A12"] = "ORDEN DE COMPRA"
    ws_cr["E12"] = "MATERIAL"
    ws_cr["A13"] = "4300008100"
    ws_cr["E13"] = "1455103"
    ws_cr["AD13"] = "Camisa Referencia Segura"
    ws_pii = wb.create_sheet("Hoja1")
    ws_pii["B2"] = "Autorización de Deducción"
    ws_pii["G18"] = "MV :"
    ws_pii["H18"] = 998
    ws_pii["C10"] = "Juan Pérez Test"
    ws_pii["C11"] = "1-0999-0888"
    ws_pii["C12"] = "₡45,000"
    ws_pii["C13"] = "SYNTHETIC_PII_MV998_SENTINEL_DO_NOT_INDEX"
    return finalize_workbook(wb, path, macro_enabled=False, use_shared_strings=True)


def _build_text_markdown_html(s_dir: Path) -> dict[str, dict[str, Any]]:
    """Create SYNTH_NOTES.txt, SYNTH_TCODE_REF.md, and SYNTH_EXPORT_GUIDE.html."""
    results: dict[str, dict[str, Any]] = {}
    notes_lines = [f"L{idx:02d}: Nota operativa recepcion C001." for idx in range(1, 56)]
    notes_lines[24] = (
        "L25: En RFUI usar movimiento 522 con CI proc. Y999 hacia destino CLARIFICACION."
    )
    notes_lines[41] = (
        "L42: Borrador usa VL3IN (error tipografico de VL31N) al crear entrega 1800...."
    )
    notes_bytes = ("\n".join(notes_lines) + "\n").encode("utf-8")
    notes_path = s_dir / "SYNTH_NOTES.txt"
    notes_path.write_bytes(notes_bytes)
    results["sap_docs/SYNTH_NOTES.txt"] = {
        "path": str(notes_path), "file_sha256": sha256_bytes(notes_bytes), "size_bytes": len(notes_bytes),
    }
    md_text = (
        "# Referencia Rapida de Transacciones SAP (Sintetica)\n\n"
        "## ME80FN\nConsulta masiva de pedidos 43 y 41. Usar Shift+F4 y UMP.\n\n"
        "## MB51\nMovimientos 101-102 y 641-642 en centro C001.\n\n"
        "## SE16N (EINA / MAST / STPO)\n"
        "Tablas SAP de solo lectura (ACTVT=03). Nunca habilitar &SAP_EDIT.\n\n"
        "## MIGO / ME23N\nVisualizacion de pedidos (preferir ME23N read-only sobre ME21N/ME22N).\n\n"
        "## VKP2\nSolo precio de venta al detalle en CRC; nunca usar como costo.\n\n"
        "[Pending Live SAP GUI 7.00 Verification]\n"
    ).encode("utf-8")
    md_path = s_dir / "SYNTH_TCODE_REF.md"
    md_path.write_bytes(md_text)
    results["sap_docs/SYNTH_TCODE_REF.md"] = {
        "path": str(md_path), "file_sha256": sha256_bytes(md_text), "size_bytes": len(md_text),
    }
    html_text = (
        "<!DOCTYPE html>\n<html lang=\"es\">\n<head><meta charset=\"utf-8\">"
        "<title>Guia SAP EWM C001</title></head>\n<body>\n"
        "<section id=\"vl31n\"><h2>VL31N Entrega Entrante</h2>"
        "<p>Pestana Gestion 2 Ident. ext. = IMP-X genera documento 1800....</p></section>\n"
        "<section id=\"scwm-mon\"><h2>/SCWM/MON Monitor EWM</h2>"
        "<p>Almacen C001, indicador Z1, seccion ZALB, PICK, muelle Z001, ubicacion Z970.</p></section>\n"
        "<section id=\"mb52-mb5t\"><h2>MB52 / MB5T Stock y Transito</h2>"
        "<p>Verificar stock libre y en transito tras Mov 641.</p></section>\n"
        "</body>\n</html>\n"
    ).encode("utf-8")
    html_path = s_dir / "SYNTH_EXPORT_GUIDE.html"
    html_path.write_bytes(html_text)
    results["sap_docs/SYNTH_EXPORT_GUIDE.html"] = {
        "path": str(html_path), "file_sha256": sha256_bytes(html_text), "size_bytes": len(html_text),
    }
    return results


def _build_corrupt_files(s_dir: Path) -> dict[str, dict[str, Any]]:
    """Create corrupt/ boundary fixtures for resilient importer error handling."""
    c_dir = s_dir / "corrupt"
    c_dir.mkdir(parents=True, exist_ok=True)
    payloads: dict[str, bytes] = {
        "empty.txt": b"",
        "malformed.pdf": b"%PDF-1.4\n%CORRUPT-SYNTHETIC-PAYLOAD\n",
        "corrupt.docx": b"PK\x03\x04CORRUPT_ZIP_PAYLOAD_SYNTHETIC",
        "unsupported.bin": b"\x00\x01\x02\x03ACR_UNSUPPORTED_BIN\xff\xfe",
    }
    results: dict[str, dict[str, Any]] = {}
    for fname, raw in payloads.items():
        fpath = c_dir / fname
        fpath.write_bytes(raw)
        results[f"sap_docs/corrupt/{fname}"] = {
            "path": str(fpath), "file_sha256": sha256_bytes(raw), "size_bytes": len(raw),
        }
    return results


def build_all_sap_docs_fixtures(out_dir: Path) -> dict[str, dict[str, Any]]:
    """Build all SAP documentation and corrupt boundary fixtures under out_dir/sap_docs/."""
    s_dir = out_dir / "sap_docs"
    s_dir.mkdir(parents=True, exist_ok=True)
    results: dict[str, dict[str, Any]] = {
        "sap_docs/SYNTH_SAP_GUIDE.pdf": _build_sap_guide_pdf(s_dir / "SYNTH_SAP_GUIDE.pdf"),
        "sap_docs/SYNTH_SCAN_NO_TEXT.pdf": _build_scan_no_text_pdf(s_dir / "SYNTH_SCAN_NO_TEXT.pdf"),
        "sap_docs/SYNTH_SOP.docx": _build_sop_docx(s_dir / "SYNTH_SOP.docx"),
        "sap_docs/SYNTH_DATA_DICT.xlsx": _build_data_dict_xlsx(s_dir / "SYNTH_DATA_DICT.xlsx"),
        "sap_docs/SYNTH_CLOSURE_WITH_HOJA1_PII.xlsx": _build_hoja1_pii_xlsx(
            s_dir / "SYNTH_CLOSURE_WITH_HOJA1_PII.xlsx"
        ),
    }
    results.update(_build_text_markdown_html(s_dir))
    results.update(_build_corrupt_files(s_dir))
    return results
