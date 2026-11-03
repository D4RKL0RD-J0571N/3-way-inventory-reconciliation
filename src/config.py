"""
Centralized configuration and project path resolution.
Agnostic architecture for multi-vendor and multi-ERP reconciliation.
"""
from pathlib import Path
from typing import List, Tuple, Optional

# Base directories
SRC_DIR = Path(__file__).resolve().parent
BASE_DIR = SRC_DIR.parent
DATA_DIR = BASE_DIR / "data"
DOCS_DIR = BASE_DIR / "docs"

def list_available_local_imports() -> List[str]:
    """
    Scans the local data/ directory for import subfolders (e.g., IMP-182, IMP-183).
    Returns a sorted list of folder names found on disk.
    """
    if not DATA_DIR.exists():
        return []
    return sorted([
        d.name for d in DATA_DIR.iterdir()
        if d.is_dir() and not d.name.startswith(".")
    ])

def resolve_local_import_files(import_id: str) -> Tuple[Optional[Path], Optional[Path]]:
    """
    Resolves the Quality Control and Closure template files for a given import ID.
    Looks for common naming conventions without hardcoding rigid paths.
    """
    imp_dir = DATA_DIR / import_id
    if not imp_dir.exists():
        return None, None
    
    # Heuristic matching for Quality Control file
    qc_candidates = (
        list(imp_dir.glob("*Control*calidad*.xlsm")) +
        list(imp_dir.glob("*Control*Calidad*.xlsx")) +
        list(imp_dir.glob("*QC*.xlsm")) +
        list(imp_dir.glob("*QC*.xlsx"))
    )
    
    # Heuristic matching for Closure / Macro template file
    cierre_candidates = (
        list(imp_dir.glob("*inprogress*.xlsm")) +
        list(imp_dir.glob("*Cierre*.xlsm")) +
        list(imp_dir.glob("*Plantilla*.xlsm")) +
        list(imp_dir.glob("*Cierre*.xlsx"))
    )
    
    qc_file = qc_candidates[0] if qc_candidates else None
    cierre_file = cierre_candidates[0] if cierre_candidates else None
    return qc_file, cierre_file
