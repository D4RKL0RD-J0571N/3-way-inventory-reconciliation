"""Deterministic OpenXML, ZIP, VBA, and Office lockfile utilities."""

from __future__ import annotations

import hashlib
import io
from pathlib import Path
import re
from typing import Any
import zipfile

from lxml import etree
import openpyxl

FIXED_ZIP_TIME: tuple[int, int, int, int, int, int] = (2026, 11, 18, 12, 0, 0)
FIXED_ISO_TIME: str = "2026-11-18T12:00:00Z"
OLE2_MAGIC: bytes = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

MAIN_NS: str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS: str = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS: str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS: str = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_SPACE: str = "{http://www.w3.org/XML/1998/namespace}space"

VBA_CONTENT_TYPE: str = "application/vnd.ms-office.vbaProject"
XLSM_CONTENT_TYPE: str = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
SHARED_STRINGS_CT: str = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
)
VBA_REL_TYPE: str = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
SHARED_STRINGS_REL_TYPE: str = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
)


def sha256_bytes(data: bytes) -> str:
    """Return lowercase hexadecimal SHA-256 digest of bytes."""
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    """Return lowercase hexadecimal SHA-256 digest of a file."""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_vba_project_bin(variant_tag: str = "default", size_bytes: int = 4096) -> bytes:
    """Build deterministic OLE2-header synthetic vbaProject.bin payload."""
    if size_bytes < 512:
        raise ValueError("vbaProject.bin size must be at least 512 bytes.")
    header = bytearray(512)
    header[0:8] = OLE2_MAGIC
    header[24:26] = (0x003E).to_bytes(2, "little")
    header[26:28] = (0x0003).to_bytes(2, "little")
    header[28:30] = (0xFFFE).to_bytes(2, "little")
    header[30:32] = (0x0009).to_bytes(2, "little")
    body = bytearray()
    counter = 0
    while len(body) < size_bytes - len(header):
        seed = f"ACR_VBA_BIN|{variant_tag}|{counter}".encode("utf-8")
        body.extend(hashlib.sha256(seed).digest())
        counter += 1
    return bytes(header) + bytes(body[: size_bytes - len(header)])


def build_legacy_xls_bytes(marker: str = "SYNTH_LEGACY_BIFF8") -> bytes:
    """Build deterministic non-ZIP OLE2/BIFF8 legacy .xls trap bytes."""
    payload = bytearray(4096)
    payload[0:8] = OLE2_MAGIC
    tag_bytes = marker.encode("ascii", errors="ignore")[:64]
    payload[512 : 512 + len(tag_bytes)] = tag_bytes
    return bytes(payload)


def write_lockfile(path: Path, owner: str = "ACR Synthetic Operator") -> Path:
    """Write a deterministic 165-byte non-ZIP Office/OneDrive lockfile."""
    path.parent.mkdir(parents=True, exist_ok=True)
    name_bytes = owner.encode("ascii", errors="ignore")[:50]
    buf = bytearray(165)
    buf[0] = len(name_bytes)
    buf[1 : 1 + len(name_bytes)] = name_bytes
    buf[55] = len(name_bytes)
    buf[56 : 56 + len(name_bytes)] = name_bytes
    path.write_bytes(bytes(buf))
    return path


def normalize_core_xml(
    raw_xml: bytes,
    creator: str = "ACR Fixture Generator",
    last_modified_by: str = "ACR Fixture Generator",
    created: str = FIXED_ISO_TIME,
    modified: str = FIXED_ISO_TIME,
) -> bytes:
    """Normalize docProps/core.xml timestamps and author tags deterministically."""
    root = etree.fromstring(raw_xml)
    ns = {
        "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        "dc": "http://purl.org/dc/elements/1.1/",
        "dcterms": "http://purl.org/dc/terms/",
        "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    }
    _set_or_create_text(root, f"{{{ns['dc']}}}creator", creator)
    _set_or_create_text(root, f"{{{ns['cp']}}}lastModifiedBy", last_modified_by)
    c_el = _set_or_create_text(root, f"{{{ns['dcterms']}}}created", created)
    c_el.set(f"{{{ns['xsi']}}}type", "dcterms:W3CDTF")
    m_el = _set_or_create_text(root, f"{{{ns['dcterms']}}}modified", modified)
    m_el.set(f"{{{ns['xsi']}}}type", "dcterms:W3CDTF")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def _set_or_create_text(parent: etree._Element, tag: str, value: str) -> etree._Element:
    """Find child element by tag or create it, then assign text."""
    elem = parent.find(tag)
    if elem is None:
        elem = etree.SubElement(parent, tag)
    elem.text = value
    return elem


def map_sheet_paths(entries: dict[str, bytes]) -> dict[str, str]:
    """Map workbook sheet names to zip-relative worksheet XML paths."""
    wb_root = etree.fromstring(entries["xl/workbook.xml"])
    rels_root = etree.fromstring(entries["xl/_rels/workbook.xml.rels"])
    rid_to_target: dict[str, str] = {}
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.get("Id", "")
        target = rel.get("Target", "")
        norm = target.lstrip("/") if target.startswith("/") else f"xl/{target}"
        rid_to_target[rid] = norm
    sheet_map: dict[str, str] = {}
    for sheet in wb_root.findall(f".//{{{MAIN_NS}}}sheet"):
        name = sheet.get("name", "")
        rid = sheet.get(f"{{{DOC_REL_NS}}}id", "")
        if name and rid in rid_to_target:
            sheet_map[name] = rid_to_target[rid]
    return sheet_map


EXCEL_ERRORS: tuple[str, ...] = (
    "#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"
)


def inject_cached_values_in_sheet(sheet_xml: bytes, cached: dict[str, Any]) -> bytes:
    """Inject cached <v> values onto formula cells in a worksheet XML."""
    root = etree.fromstring(sheet_xml)
    for cell in root.findall(f".//{{{MAIN_NS}}}c"):
        ref = cell.get("r", "")
        if ref not in cached:
            continue
        val = cached[ref]
        v_el = cell.find(f"{{{MAIN_NS}}}v")
        if v_el is None:
            v_el = etree.SubElement(cell, f"{{{MAIN_NS}}}v")
        if isinstance(val, str) and val in EXCEL_ERRORS:
            cell.set("t", "e")
            v_el.text = val
        elif isinstance(val, str):
            cell.set("t", "str")
            v_el.text = val
        elif isinstance(val, bool):
            cell.set("t", "b")
            v_el.text = "1" if val else "0"
        else:
            if "t" in cell.attrib:
                del cell.attrib["t"]
            v_el.text = str(val)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def omit_dimension_in_sheet(sheet_xml: bytes) -> bytes:
    """Remove <dimension> element from worksheet XML to reproduce IMP-81 pattern."""
    root = etree.fromstring(sheet_xml)
    dim = root.find(f"{{{MAIN_NS}}}dimension")
    if dim is not None:
        root.remove(dim)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def append_bloat_rows(sheet_xml: bytes, count: int) -> bytes:
    """Append trailing empty <row r='...'/> elements for streaming bloat tests."""
    if count <= 0:
        return sheet_xml
    root = etree.fromstring(sheet_xml)
    sheet_data = root.find(f"{{{MAIN_NS}}}sheetData")
    if sheet_data is None:
        return sheet_xml
    existing_rows = sheet_data.findall(f"{{{MAIN_NS}}}row")
    max_r = max((int(r.get("r", "0")) for r in existing_rows), default=0)
    start_r = max_r + 1
    for offset in range(count):
        etree.SubElement(sheet_data, f"{{{MAIN_NS}}}row", {"r": str(start_r + offset)})
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def convert_inline_strings_to_shared(
    sheet_xml: bytes,
    shared_list: list[str],
    shared_index: dict[str, int],
) -> bytes:
    """Convert inlineStr cells to sharedStrings t='s' cells."""
    root = etree.fromstring(sheet_xml)
    for cell in root.findall(f".//{{{MAIN_NS}}}c"):
        if cell.get("t") != "inlineStr":
            continue
        is_el = cell.find(f"{{{MAIN_NS}}}is")
        text = "".join(is_el.itertext()) if is_el is not None else ""
        if text not in shared_index:
            shared_index[text] = len(shared_list)
            shared_list.append(text)
        idx = shared_index[text]
        if is_el is not None:
            cell.remove(is_el)
        cell.set("t", "s")
        v_el = etree.SubElement(cell, f"{{{MAIN_NS}}}v")
        v_el.text = str(idx)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8")


def build_shared_strings_xml(strings: list[str]) -> bytes:
    """Serialize sharedStrings table to canonical OpenXML bytes."""
    total = str(len(strings))
    sst = etree.Element(
        f"{{{MAIN_NS}}}sst",
        nsmap={None: MAIN_NS},
        attrib={"count": total, "uniqueCount": total},
    )
    for text in strings:
        si = etree.SubElement(sst, f"{{{MAIN_NS}}}si")
        t_el = etree.SubElement(si, f"{{{MAIN_NS}}}t")
        if text != text.strip() or "\n" in text or "\r" in text:
            t_el.set(XML_SPACE, "preserve")
        t_el.text = text
    return etree.tostring(sst, xml_declaration=True, encoding="UTF-8")


def _next_rid(rels_root: etree._Element) -> str:
    """Compute next unused rIdN relationship identifier."""
    used = set()
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        match = re.fullmatch(r"rId(\d+)", rel.get("Id", ""))
        if match:
            used.add(int(match.group(1)))
    idx = 1
    while idx in used:
        idx += 1
    return f"rId{idx}"


def ensure_workbook_macro_parts(
    entries: dict[str, bytes],
    vba_bytes: bytes | None,
    has_shared_strings: bool,
) -> None:
    """Update [Content_Types].xml and workbook.xml.rels for VBA and sharedStrings."""
    ct_root = etree.fromstring(entries["[Content_Types].xml"])
    rels_root = etree.fromstring(entries["xl/_rels/workbook.xml.rels"])
    if vba_bytes is not None:
        entries["xl/vbaProject.bin"] = vba_bytes
        for ov in ct_root.findall(f"{{{CT_NS}}}Override"):
            if ov.get("PartName") == "/xl/workbook.xml":
                ov.set("ContentType", XLSM_CONTENT_TYPE)
        _ensure_override(ct_root, "/xl/vbaProject.bin", VBA_CONTENT_TYPE)
        _ensure_rel(rels_root, VBA_REL_TYPE, "vbaProject.bin")
    if has_shared_strings:
        _ensure_override(ct_root, "/xl/sharedStrings.xml", SHARED_STRINGS_CT)
        _ensure_rel(rels_root, SHARED_STRINGS_REL_TYPE, "sharedStrings.xml")
    entries["[Content_Types].xml"] = etree.tostring(
        ct_root, xml_declaration=True, encoding="UTF-8"
    )
    entries["xl/_rels/workbook.xml.rels"] = etree.tostring(
        rels_root, xml_declaration=True, encoding="UTF-8"
    )


def _ensure_override(ct_root: etree._Element, part_name: str, ctype: str) -> None:
    """Add or update an Override entry in [Content_Types].xml."""
    for ov in ct_root.findall(f"{{{CT_NS}}}Override"):
        if ov.get("PartName") == part_name:
            ov.set("ContentType", ctype)
            return
    etree.SubElement(
        ct_root, f"{{{CT_NS}}}Override", {"PartName": part_name, "ContentType": ctype}
    )


def _ensure_rel(rels_root: etree._Element, rel_type: str, target: str) -> None:
    """Add a Relationship entry if not already present."""
    for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
        if rel.get("Type") == rel_type and rel.get("Target") == target:
            return
    etree.SubElement(
        rels_root,
        f"{{{REL_NS}}}Relationship",
        {"Id": _next_rid(rels_root), "Type": rel_type, "Target": target},
    )


def write_deterministic_zip(entries: dict[str, bytes], output_path: Path) -> None:
    """Write zip archive with sorted entry names and fixed timestamps."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(entries.keys()):
            info = zipfile.ZipInfo(filename=name, date_time=FIXED_ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            zf.writestr(info, entries[name])


def _prepare_entries_and_sheets(
    wb: openpyxl.Workbook,
    *,
    use_shared_strings: bool,
    extra_shared_strings: list[str] | None,
    cached_values: dict[str, dict[str, Any]] | None,
    bloat_sheets: dict[str, int] | None,
    omit_dimension_sheets: set[str] | None,
    core_creator: str,
    core_last_modified_by: str,
    core_created: str,
    core_modified: str,
) -> tuple[dict[str, bytes], list[str]]:
    """Serialize workbook and transform worksheet XML and core metadata."""
    buf = io.BytesIO()
    wb.save(buf)
    wb.close()
    buf.seek(0)
    with zipfile.ZipFile(buf, "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    if "docProps/core.xml" in entries:
        entries["docProps/core.xml"] = normalize_core_xml(
            entries["docProps/core.xml"],
            creator=core_creator,
            last_modified_by=core_last_modified_by,
            created=core_created,
            modified=core_modified,
        )
    shared_list: list[str] = []
    shared_index: dict[str, int] = {}
    _transform_sheets(
        entries, map_sheet_paths(entries),
        use_shared_strings=use_shared_strings, shared_list=shared_list,
        shared_index=shared_index, cached_values=cached_values or {},
        bloat_sheets=bloat_sheets or {}, omit_dimension_sheets=omit_dimension_sheets or set(),
    )
    if use_shared_strings and extra_shared_strings:
        for s in extra_shared_strings:
            if s not in shared_index:
                shared_index[s] = len(shared_list)
                shared_list.append(s)
    return entries, shared_list


def finalize_workbook(
    wb: openpyxl.Workbook,
    output_path: Path,
    *,
    macro_enabled: bool = True,
    use_shared_strings: bool = True,
    vba_variant: str = "default",
    extra_shared_strings: list[str] | None = None,
    cached_values: dict[str, dict[str, Any]] | None = None,
    bloat_sheets: dict[str, int] | None = None,
    omit_dimension_sheets: set[str] | None = None,
    core_creator: str = "ACR Fixture Generator",
    core_last_modified_by: str = "ACR Fixture Generator",
    core_created: str = FIXED_ISO_TIME,
    core_modified: str = FIXED_ISO_TIME,
) -> dict[str, Any]:
    """Post-process openpyxl workbook into deterministic OpenXML .xlsm/.xlsx."""
    entries, shared_list = _prepare_entries_and_sheets(
        wb, use_shared_strings=use_shared_strings,
        extra_shared_strings=extra_shared_strings, cached_values=cached_values,
        bloat_sheets=bloat_sheets, omit_dimension_sheets=omit_dimension_sheets,
        core_creator=core_creator, core_last_modified_by=core_last_modified_by,
        core_created=core_created, core_modified=core_modified,
    )
    if use_shared_strings:
        entries["xl/sharedStrings.xml"] = build_shared_strings_xml(shared_list)
    vba_bytes = build_vba_project_bin(vba_variant) if macro_enabled else None
    ensure_workbook_macro_parts(entries, vba_bytes, use_shared_strings)
    write_deterministic_zip(entries, output_path)
    return {
        "path": str(output_path),
        "file_sha256": sha256_file(output_path),
        "vba_sha256": sha256_bytes(vba_bytes) if vba_bytes else None,
        "size_bytes": output_path.stat().st_size,
        "shared_strings_count": len(shared_list) if use_shared_strings else 0,
    }


def _transform_sheets(
    entries: dict[str, bytes],
    sheet_map: dict[str, str],
    *,
    use_shared_strings: bool,
    shared_list: list[str],
    shared_index: dict[str, int],
    cached_values: dict[str, dict[str, Any]],
    bloat_sheets: dict[str, int],
    omit_dimension_sheets: set[str],
) -> None:
    """Apply cached values, shared strings, bloat rows, and dimension omission."""
    for sheet_name, rel_path in sheet_map.items():
        if rel_path not in entries:
            continue
        xml_bytes = entries[rel_path]
        if sheet_name in cached_values:
            xml_bytes = inject_cached_values_in_sheet(xml_bytes, cached_values[sheet_name])
        if use_shared_strings:
            xml_bytes = convert_inline_strings_to_shared(
                xml_bytes, shared_list, shared_index
            )
        if sheet_name in bloat_sheets:
            xml_bytes = append_bloat_rows(xml_bytes, bloat_sheets[sheet_name])
        if sheet_name in omit_dimension_sheets:
            xml_bytes = omit_dimension_in_sheet(xml_bytes)
        entries[rel_path] = xml_bytes


def normalize_docx_file(
    docx_bytes: bytes,
    output_path: Path,
    creator: str = "ACR Fixture Generator",
) -> dict[str, Any]:
    """Normalize python-docx archive timestamps and core.xml for determinism."""
    with zipfile.ZipFile(io.BytesIO(docx_bytes), "r") as zin:
        entries = {name: zin.read(name) for name in zin.namelist()}
    if "docProps/core.xml" in entries:
        entries["docProps/core.xml"] = normalize_core_xml(
            entries["docProps/core.xml"],
            creator=creator,
            last_modified_by=creator,
        )
    write_deterministic_zip(entries, output_path)
    return {
        "path": str(output_path),
        "file_sha256": sha256_file(output_path),
        "size_bytes": output_path.stat().st_size,
    }
