"""Synthetic Historical Sync (.xlsm) corpus builder under historical_sync/."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from closure_builder import build_closure_workbook
from openxml_utils import sha256_bytes, write_lockfile

JOSTIN_CANONICAL = "ELABORADO POR: \nJostin Lopez Sobalbarro"
JOSTIN_TYPO = "ELABORADO POR: \nJostin Lopez Sobalbaro"
ANGIE_AUTHOR = "ELABORADO POR: \nAngie Montero Rojas"


def _hist_items(po_43: str, count: int = 2, cartons_each: int = 5) -> list[dict[str, Any]]:
    """Build deterministic item list for historical closure fixtures."""
    items: list[dict[str, Any]] = []
    for idx in range(count):
        items.append({
            "po_43": po_43,
            "vendor_material": f"HST-{idx:02d}",
            "material_code": f"6001{idx:03d}",
            "description": f"Historico Material {idx + 1}",
            "cod_prov": 1100000577,
            "cat_code": "G40403001",
            "as_cartons": cartons_each,
            "at_units": cartons_each * 24,
            "aw_cartons": cartons_each,
            "ax_units": cartons_each * 24,
            "ay_empaque": "24",
            "curve": 12,
            "uom": "PRK",
            "qty_43": cartons_each * 2,
            "qty_41": cartons_each * 2,
            "qty_101": cartons_each * 2,
        })
    return items


def _build_primary_historical(h_dir: Path) -> dict[str, dict[str, Any]]:
    """Build IMP104, IMP326, IMP209, and IMP182 family historical fixtures."""
    results: dict[str, dict[str, Any]] = {}
    results["historical_sync/IMP104_Clean.xlsm"] = build_closure_workbook(
        h_dir / "IMP104_Clean.xlsm",
        container_id="IMP-104", po_43="4300008021", pedido_41="4100002248",
        author_b33=JOSTIN_CANONICAL, items=_hist_items("4300008021", 3, 10),
        vba_variant="hist_imp104", core_creator="Hilario Lopez Rivera",
    )
    results["historical_sync/IMP326_Typo.xlsm"] = build_closure_workbook(
        h_dir / "IMP326_Typo.xlsm",
        container_id="IMP-326", po_43="4300007309-8144", pedido_41="4100002302",
        author_b33=JOSTIN_TYPO, items=_hist_items("4300007309", 3, 8),
        vba_variant="hist_imp326", core_creator="Hilario Lopez Rivera",
    )
    results["historical_sync/IMP209_SharedStringsTrap.xlsm"] = build_closure_workbook(
        h_dir / "IMP209_SharedStringsTrap.xlsm",
        container_id="IMP-209", po_43="4300007999", pedido_41="4100002290",
        author_b33=ANGIE_AUTHOR, items=_hist_items("4300007999", 2, 6),
        vba_variant="hist_imp209", core_creator="Hilario Lopez Rivera",
        extra_shared_strings=["Jostin Lopez Sobalbarro"],
    )
    results["historical_sync/IMP182_Final.xlsm"] = build_closure_workbook(
        h_dir / "IMP182_Final.xlsm",
        container_id="IMP-182", po_43="4300006764", pedido_41="4100002283",
        author_b33=JOSTIN_CANONICAL, items=_hist_items("4300006764", 3, 12),
        vba_variant="hist_imp182_final", core_creator="Hilario Lopez Rivera",
    )
    results["historical_sync/Copia de IMP182 inprogress.xlsm"] = build_closure_workbook(
        h_dir / "Copia de IMP182 inprogress.xlsm",
        container_id="IMP-182", po_43="4300006764", pedido_41="4100002283",
        author_b33=JOSTIN_CANONICAL, items=_hist_items("", 3, 4),
        vba_variant="hist_imp182_copia", dg_bultos_teo=17172, dg_bultos_rec=14,
        core_creator="Hilario Lopez Rivera",
    )
    results["historical_sync/drafts/IMP182_inprogress.xlsm"] = build_closure_workbook(
        h_dir / "drafts" / "IMP182_inprogress.xlsm",
        container_id="IMP-182", po_43="4300006764", pedido_41="4100002283",
        author_b33=JOSTIN_CANONICAL, items=_hist_items("", 3, 4),
        vba_variant="hist_imp182_draft", dg_bultos_teo=1133, dg_bultos_rec=14,
        core_creator="Hilario Lopez Rivera",
    )
    return results


def _build_shared_po_and_anomaly_historical(h_dir: Path) -> dict[str, dict[str, Any]]:
    """Build shared-PO (224/225/226), zero-carton, hybrid, non-imp, and lockfile fixtures."""
    results: dict[str, dict[str, Any]] = {}
    shared_specs = [
        ("IMP224_SharedPO.xlsm", "IMP-224", "4100002293", "PROVEEDOR TEXTIL SINTETICO S.A."),
        ("IMP225_SharedPO.xlsm", "IMP-225", "4100002296", "#REF!"),
        ("IMP226_SharedPO.xlsm", "IMP-226", "4100002297", ""),
    ]
    for fname, cid, ped41, prov in shared_specs:
        results[f"historical_sync/{fname}"] = build_closure_workbook(
            h_dir / fname,
            container_id=cid, po_43="4300007562", pedido_41=ped41,
            author_b33=JOSTIN_CANONICAL, items=_hist_items("4300007562", 3, 10),
            vba_variant=f"hist_{cid}", dg_provider=prov,
            core_creator="Hilario Lopez Rivera",
        )
    results["historical_sync/IMP195_ZeroCartons.xlsm"] = build_closure_workbook(
        h_dir / "IMP195_ZeroCartons.xlsm",
        container_id="IMP-195", po_43="4300006505", pedido_41="4100002284",
        author_b33=JOSTIN_CANONICAL, items=_hist_items("4300006505", 3, 0),
        vba_variant="hist_imp195", dg_bultos_teo=0, dg_bultos_rec=0,
        core_creator="Hilario Lopez Rivera",
    )
    results["historical_sync/IMP218_Hybrid.xlsm"] = build_closure_workbook(
        h_dir / "IMP218_Hybrid.xlsm",
        container_id="IMP-218", po_43="4300007216", pedido_41="4100002274",
        author_b33=JOSTIN_CANONICAL, items=_hist_items("4300007216", 3, 9),
        vba_variant="hist_imp218", include_control_data=True, dg_lineas_c18=0,
        core_creator="Hilario Lopez Rivera",
    )
    results["historical_sync/Macro para cierre hannia.xlsm"] = build_closure_workbook(
        h_dir / "Macro para cierre hannia.xlsm",
        container_id="IMP-950", po_43="4300008950", pedido_41="4100002950",
        author_b33=JOSTIN_CANONICAL, items=_hist_items("4300008950", 2, 5),
        vba_variant="hist_hannia", core_creator="Hilario Lopez Rivera",
    )
    lock_path = write_lockfile(h_dir / "~$IMP349_Lockfile.xlsm", "IMP349 Lock")
    lock_bytes = lock_path.read_bytes()
    results["historical_sync/~$IMP349_Lockfile.xlsm"] = {
        "path": str(lock_path),
        "file_sha256": sha256_bytes(lock_bytes),
        "vba_sha256": None,
        "size_bytes": len(lock_bytes),
    }
    return results


def build_all_historical_fixtures(out_dir: Path) -> dict[str, dict[str, Any]]:
    """Build all historical sync synthetic fixtures under out_dir/historical_sync/."""
    h_dir = out_dir / "historical_sync"
    h_dir.mkdir(parents=True, exist_ok=True)
    results = _build_primary_historical(h_dir)
    results.update(_build_shared_po_and_anomaly_historical(h_dir))
    return results
