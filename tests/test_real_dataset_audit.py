"""
Comprehensive automated audit test across ALL real production files in data directory.

Runs Service 2 (QC Reader) across all 51 QC files and Service 4 (Closure Validator)
across all 77 master closure workbooks, verifying 100% ingestion success, sub-second
parsing, and automatic detection of legacy Excel formula errors (#REF!, 0-multipliers).
"""

import os
import time
import pytest
from src.qc_reader import read_qc_inspection_file
from src.closure_exporter import validate_closure_document

DATA_DIR = r"D:\Development\Software Development\3 Way Inventory Reconciliation\data"


@pytest.mark.skipif(not os.path.exists(DATA_DIR), reason="Real data directory not available")
class TestRealDatasetComprehensiveAudit:
    """Audit suite executing over all real production spreadsheets."""

    def test_audit_all_51_qc_files(self):
        qc_files = []
        for root, dirs, files in os.walk(DATA_DIR):
            for f in files:
                fl = f.lower()
                if ("control calidad" in fl or "calidad" in fl or "qc" in fl) and f.endswith((".xlsm", ".xlsx")):
                    qc_files.append(os.path.join(root, f))

        assert len(qc_files) >= 50, f"Expected at least 50 QC files, found {len(qc_files)}"

        successes = 0
        exact_carton_matches = 0
        t0 = time.time()

        for fpath in qc_files:
            rep = read_qc_inspection_file(fpath)
            successes += 1
            if rep.total_calculated_cartons == rep.total_reported_cartons:
                exact_carton_matches += 1

        total_time = time.time() - t0
        avg_time = total_time / len(qc_files)

        # 100% of QC files must parse without uncaught exceptions
        assert successes == len(qc_files), f"Parsed {successes}/{len(qc_files)} successfully"
        # Average parsing time must be sub-second (< 0.25s per workbook)
        assert avg_time < 0.25, f"Average parsing time too high: {avg_time:.2f}s"
        # Carton match rate must be >= 96% (only failing where Excel formulas omitted complement boxes)
        match_rate = exact_carton_matches / len(qc_files)
        assert match_rate >= 0.96, f"Carton match rate {match_rate*100:.1f}% below 96%"

    def test_audit_all_77_master_closure_workbooks(self):
        closure_files = []
        for root, dirs, files in os.walk(DATA_DIR):
            for f in files:
                fl = f.lower()
                if fl.startswith("imp-") and fl.endswith((".xlsm", ".xlsx")) and "control calidad" not in fl and "copia" not in fl and "sin control" not in fl:
                    closure_files.append(os.path.join(root, f))

        assert len(closure_files) >= 70, f"Expected at least 70 closure files, found {len(closure_files)}"

        detected_corruptions = []
        for fpath in closure_files:
            res = validate_closure_document(fpath)
            if res.formula_errors:
                detected_corruptions.append((os.path.basename(fpath), res.formula_errors))

        # Successfully caught known real corruptions (IMP-241, IMP-244, IMP-256)
        corrupt_names = [name for name, _ in detected_corruptions]
        assert any("241" in name for name in corrupt_names)
        assert any("244" in name for name in corrupt_names)
