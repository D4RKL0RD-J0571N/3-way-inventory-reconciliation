"""
Heuristic and fuzzy column detector for messy, human-made Excel sheets.
Handles warehouse operators changing column names, inserting rows,
or using different terminology across suppliers and departments.
"""
from typing import Dict, List, Optional, Tuple, Any
import openpyxl

# Synonyms for key warehouse and logistics fields
COLUMN_SYNONYMS = {
    "material": ["material", "código", "codigo", "articulo", "artículo", "producto", "sku", "mat", "cod. sap", "cod sap"],
    "vendor_material": ["material proveedor", "mat proveedor", "mat. proveedor", "estilo", "style", "ref proveedor", "referencia"],
    "description": ["descripción", "descripcion", "desc", "detalle", "texto breve", "nombre", "item"],
    "quantity": ["cantidad", "cant", "cant.", "unidades", "unid", "total unidades", "qty", "recibido", "recibido real", "fisico", "físico"],
    "curve": ["curva", "ratio", "curva sistema", "ratios", "curv"],
    "uom": ["um", "u.m.", "unidad de medida", "medida", "tipo"],
    "po": ["orden de compra", "orden compra", "po", "pedido", "pedido compra", "doc. compras"],
    "unit_cost": ["costo", "valor unitario", "precio", "precio unitario", "costo unitario", "cost", "vkp2"],
    "bom_id": ["lista", "lista de materiales", "bom", "bom id", "grupo lista"]
}

def normalize_text(text: Any) -> str:
    """Normalizes string for fuzzy matching (lowercase, no accents, trimmed)."""
    if not text:
        return ""
    s = str(text).lower().strip()
    s = s.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    s = s.replace("\n", " ").replace("\r", " ").replace("\t", " ")
    return " ".join(s.split())

def detect_headers_in_sheet(
    ws,
    max_search_rows: int = 20,
    required_fields: Optional[List[str]] = None
) -> Tuple[int, Dict[str, int]]:
    """
    Scans the first `max_search_rows` of an openpyxl sheet to heuristically
    find the header row and map field names to column indices (1-based).
    """
    if required_fields is None:
        required_fields = ["material", "quantity"]

    best_row = -1
    best_mapping: Dict[str, int] = {}
    best_match_count = -1

    for r in range(1, min(max_search_rows + 1, ws.max_row + 1)):
        row_values = [normalize_text(ws.cell(r, c).value) for c in range(1, min(ws.max_column + 1, 100))]
        
        # Try matching this row against synonyms
        current_mapping: Dict[str, int] = {}
        for col_idx, cell_text in enumerate(row_values, start=1):
            if not cell_text:
                continue
            for field, synonyms in COLUMN_SYNONYMS.items():
                if field not in current_mapping:
                    for syn in synonyms:
                        # Exact match or substring match
                        if syn == cell_text or syn in cell_text.split() or cell_text.startswith(syn):
                            current_mapping[field] = col_idx
                            break

        # Check if required fields exist
        has_required = all(rf in current_mapping for rf in required_fields)
        match_count = len(current_mapping)

        if has_required and match_count > best_match_count:
            best_match_count = match_count
            best_row = r
            best_mapping = current_mapping

    return best_row, best_mapping
