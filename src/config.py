"""
Centralized configuration and project path resolution.
Agnostic architecture for multi-vendor and multi-ERP reconciliation.
"""
from pathlib import Path
from typing import List, Tuple, Optional

import sys
from pathlib import Path
from typing import List, Tuple, Optional

# Base directories
SRC_DIR = Path(__file__).resolve().parent

if getattr(sys, "frozen", False):
    # PyInstaller executable root folder
    BASE_DIR = Path(sys.executable).parent
    INTERNAL_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
else:
    BASE_DIR = SRC_DIR.parent
    INTERNAL_DIR = BASE_DIR

DATA_DIR = BASE_DIR / "data"
if not DATA_DIR.exists() and (INTERNAL_DIR / "data").exists():
    DATA_DIR = INTERNAL_DIR / "data"

DOCS_DIR = BASE_DIR / "docs"
if not DOCS_DIR.exists() and (INTERNAL_DIR / "docs").exists():
    DOCS_DIR = INTERNAL_DIR / "docs"

def list_available_local_imports() -> List[str]:
    """
    Scans the local data/ directory for import subfolders (e.g., CONT-001, CONT-002).
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
    Uses case-insensitive and pattern-tolerant matching across all Excel files.
    """
    imp_dir = DATA_DIR / import_id
    if not imp_dir.exists():
        return None, None
    
    files = [f for f in imp_dir.iterdir() if f.suffix.lower() in [".xlsm", ".xlsx", ".xls"]]
    qc_file = None
    cierre_file = None
    
    for f in sorted(files, key=lambda x: len(x.name)):
        f_low = f.name.lower()
        if "final" in f_low or "resultado" in f_low:
            continue
        if "control" in f_low and "calidad" in f_low:
            qc_file = f
        elif "qc" in f_low and not qc_file:
            qc_file = f
        elif any(k in f_low for k in ["progress", "cierre", "plantilla", "recep"]):
            cierre_file = f
            
    return qc_file, cierre_file
