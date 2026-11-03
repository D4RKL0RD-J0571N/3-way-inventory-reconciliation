"""
Data sanitization and defensive conversion utilities.
Handles dirty inputs, messy Excel formatting, bad strings, scientific notation,
and locale-specific number representations (commas vs dots, currencies).
"""
import re
from typing import Any, List, Optional

def clean_string(val: Any) -> str:
    """
    Sanitizes string inputs:
    - Strips leading/trailing whitespace and non-breaking spaces (\xa0).
    - Removes internal repeated whitespace/tabs.
    - Handles None, float('nan'), and boolean values cleanly.
    """
    if val is None:
        return ""
    # Convert to string and replace non-breaking spaces
    s = str(val).replace("\xa0", " ").strip()
    if s.lower() in ["nan", "none", "null", "#n/a", "#value!", "#ref!", "#div/0!"]:
        return ""
    # Normalize multiple whitespace characters into a single space
    return re.sub(r"\s+", " ", s)

def clean_material_code(val: Any) -> str:
    """
    Sanitizes SAP material codes:
    - Converts scientific notation (e.g. 1.4551E+09) into integer string.
    - Strips trailing '.0' produced when Excel reads integers as floats.
    - Strips leading SAP zero padding (e.g. '0000000001455103' -> '1455103').
    - Removes internal spaces, hyphens or special formatting characters.
    """
    if val is None:
        return ""
        
    s = str(val).strip().replace("\xa0", "")
    if not s or s.lower() in ["nan", "none", "#n/a"]:
        return ""
        
    # Handle scientific notation like 1.455102001E+09
    if "e" in s.lower():
        try:
            f = float(s)
            s = f"{int(round(f))}"
        except (ValueError, OverflowError):
            pass

    # Strip .0 if float representation
    if s.endswith(".0"):
        s = s[:-2]

    # Clean characters keeping alphanumeric and basic separators
    s = re.sub(r"[^\w\-]", "", s)

    # If numeric and starts with redundant leading zeros (SAP 18-char padding)
    if s.isdigit() and len(s) > 10 and s.startswith("000"):
        s = s.lstrip("0")

    return s

def clean_number(val: Any, default: float = 0.0) -> float:
    """
    Defensively converts any Excel/user input into a valid float:
    - Handles currency symbols ($, ₡, €).
    - Handles commas as thousands or decimal separators.
    - Handles negative numbers in parentheses: (148) -> -148.0.
    - Handles Excel error strings (#VALUE!, #N/A) -> default.
    - Handles empty strings, None, NaN -> default.
    """
    if val is None:
        return default
    if isinstance(val, (int, float)):
        import math
        return default if math.isnan(val) else float(val)

    s = str(val).strip().replace("\xa0", "")
    if not s or s.lower() in ["nan", "none", "null", "-", "#n/a", "#value!", "#ref!"]:
        return default

    # Check for negative in parentheses: (150.00) -> -150.00
    is_negative = False
    if s.startswith("(") and s.endswith(")"):
        is_negative = True
        s = s[1:-1].strip()
    elif s.startswith("-"):
        is_negative = True
        s = s[1:].strip()

    # Remove currency symbols and letters
    s = re.sub(r"[^\d.,]", "", s)
    if not s:
        return default

    # Handle comma/dot decimal ambiguity
    # e.g., 1,044.50 (US format) vs 1.044,50 (Latin/European format)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            # European: 1.044,50
            s = s.replace(".", "").replace(",", ".")
        else:
            # US: 1,044.50
            s = s.replace(",", "")
    elif "," in s:
        # Only comma: if 2 decimals after comma -> likely decimal separator
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) in [1, 2]:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")

    try:
        num = float(s)
        return -num if is_negative else num
    except (ValueError, TypeError):
        return default

def clean_int(val: Any, default: int = 0) -> int:
    """
    Safely converts input into integer via clean_number.
    """
    f = clean_number(val, default=float(default))
    try:
        return int(round(f))
    except (ValueError, OverflowError):
        return default

def parse_multi_po(po_str: Any) -> List[str]:
    """
    Splits multi-PO strings like '4300006764-6804-6805' or '4300007063, 4300007064'
    into clean, normalized PO identifiers.
    """
    cleaned = clean_string(po_str)
    if not cleaned:
        return []
    # Split by commas, slashes, or hyphens if multiple POs are chained
    tokens = re.split(r"[,;/]+", cleaned)
    results = [t.strip() for t in tokens if t.strip()]
    return results
