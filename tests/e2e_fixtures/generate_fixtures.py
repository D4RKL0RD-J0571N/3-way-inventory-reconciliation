"""Deterministic CLI generator for ACR Testing-mode E2E synthetic fixtures."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
import sys
import tempfile
from typing import Any

FIXTURES_ROOT = Path(__file__).resolve().parent
BUILDERS_DIR = FIXTURES_ROOT / "builders"
if str(BUILDERS_DIR) not in sys.path:
    sys.path.insert(0, str(BUILDERS_DIR))
if str(FIXTURES_ROOT) not in sys.path:
    sys.path.insert(0, str(FIXTURES_ROOT))

from closure_builder import build_all_closure_fixtures  # noqa: E402
from daily_log_builder import build_all_daily_log_fixtures  # noqa: E402
from historical_builder import build_all_historical_fixtures  # noqa: E402
from qc_builder import build_all_qc_fixtures  # noqa: E402
from sap_docs_builder import build_all_sap_docs_fixtures  # noqa: E402

FORBIDDEN_PATH_FRAGMENTS: tuple[str, ...] = (
    "3 way inventory reconciliation\\data",
    "control_diario_operaciones_jostin.xlsx",
)


def default_output_dir() -> Path:
    """Return isolated temporary output directory outside git-tracked paths."""
    return Path(tempfile.gettempdir()) / "acr_e2e_fixtures"


def _assert_safe_output_dir(out_dir: Path) -> None:
    """Reject output paths targeting production data or repo source trees."""
    resolved = str(out_dir.resolve()).lower()
    for fragment in FORBIDDEN_PATH_FRAGMENTS:
        if fragment in resolved:
            raise ValueError(f"Refusing to write fixtures to protected path: {out_dir}")
    repo_root = FIXTURES_ROOT.parent.parent.resolve()
    try:
        out_dir.resolve().relative_to(repo_root)
        raise ValueError(
            f"Refusing to write generated fixtures inside git repository tree: {out_dir}"
        )
    except ValueError as exc:
        if "Refusing to write" in str(exc):
            raise


def generate_all_fixtures(out_dir: Path, *, clean: bool = False, bloat_rows: int = 25000) -> dict[str, Any]:
    """Materialize all synthetic E2E fixtures into out_dir and return metadata."""
    _assert_safe_output_dir(out_dir)
    if clean and out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    artifacts: dict[str, dict[str, Any]] = {}
    artifacts.update(build_all_qc_fixtures(out_dir, bloat_rows=bloat_rows))
    artifacts.update(build_all_closure_fixtures(out_dir))
    artifacts.update(build_all_historical_fixtures(out_dir))
    artifacts.update(build_all_daily_log_fixtures(out_dir))
    artifacts.update(build_all_sap_docs_fixtures(out_dir))
    return {
        "out_dir": str(out_dir),
        "bloat_rows": bloat_rows,
        "fixture_count": len(artifacts),
        "fixtures": artifacts,
    }


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse CLI arguments for fixture generation."""
    parser = argparse.ArgumentParser(description="Generate synthetic ACR E2E fixtures.")
    parser.add_argument("--out", type=Path, default=default_output_dir(), help="Target directory.")
    parser.add_argument("--clean", action="store_true", help="Clean target directory before generating.")
    parser.add_argument("--bloat-rows", type=int, default=25000, help="Trailing empty XML rows in bloat QC.")
    parser.add_argument("--verify", action="store_true", help="Run validate_fixtures after generation.")
    parser.add_argument("--summary-json", type=Path, default=None, help="Optional summary JSON output path.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """CLI entry point for generating and optionally validating E2E fixtures."""
    args = _parse_args(argv)
    summary = generate_all_fixtures(args.out, clean=args.clean, bloat_rows=args.bloat_rows)
    if args.summary_json is not None:
        _assert_safe_output_dir(args.summary_json.parent)
        args.summary_json.parent.mkdir(parents=True, exist_ok=True)
        args.summary_json.write_text(json.dumps(summary, indent=2, sort_keys=True), encoding="utf-8")
    print(f"Generated {summary['fixture_count']} synthetic fixtures in {summary['out_dir']}")
    if args.verify:
        from validate_fixtures import validate_directory  # noqa: E402

        manifest_path = FIXTURES_ROOT / "fixture_manifest.json"
        failures = validate_directory(args.out, manifest_path=manifest_path)
        if failures:
            for msg in failures:
                print(f"[FAIL] {msg}")
            return 1
        print(f"Verified {summary['fixture_count']} fixtures with 0 failures.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
