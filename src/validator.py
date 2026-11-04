"""
Data validator and business rule integrity checker.
Detects curve mismatches, unit of measure confusions, missing prices,
and out-of-bounds discrepancies before financial sign-off.
"""
from dataclasses import dataclass
from typing import List, Dict
from src.models import PrepackItem

@dataclass
class ValidationIssue:
    severity: str # "ERROR", "WARNING", "INFO"
    material_code: str
    issue_type: str
    message: str
    recommendation: str

def validate_item_integrity(item: PrepackItem) -> List[ValidationIssue]:
    """
    Validates a single PrepackItem against operational and mathematical rules.
    """
    issues: List[ValidationIssue] = []

    # 1. Curve Sum Integrity Check
    if item.children:
        sum_ratios = sum(c.ratio for c in item.children)
        if sum_ratios != item.curve_total:
            issues.append(ValidationIssue(
                severity="ERROR",
                material_code=item.material_code,
                issue_type="CURVA_DESCUADRADA",
                message=f"La suma de ratios de tallas ({sum_ratios}) no coincide con la curva oficial del Prepack ({item.curve_total}).",
                recommendation="Verificar la lista de materiales en SAP transacción STPO/MAST."
            ))

    # 2. Missing BOM Check (Prepack without children)
    if item.curve_total > 1 and not item.children and (item.received_real_units > 0 or item.container_part in ["", "A"]):
        issues.append(ValidationIssue(
            severity="WARNING",
            material_code=item.material_code,
            issue_type="BOM_FALTANTE",
            message=f"El material es un Prepack de {item.curve_total} unidades pero no se encontró desglose de tallas hijas.",
            recommendation="Confirmar si la tabla STPO fue consultada para este código de material."
        ))

    # 3. Unit of Measure Confusion (Outlier / Extreme Value)
    if item.quantity_inv_41_units > 0 and item.received_real_units > (item.quantity_inv_41_units * 2.5):
        issues.append(ValidationIssue(
            severity="ERROR",
            material_code=item.material_code,
            issue_type="POSIBLE_CONFUSION_UDM",
            message=f"Unidades recibidas ({item.received_real_units:,.0f}) superan en más del 250% lo facturado ({item.quantity_inv_41_units:,.0f}).",
            recommendation="Verificar si en el piso de descarga se digitaron piezas individuales en vez de bultos/prepacks."
        ))

    # 4. Zero Unit Cost with Shortage/Damage
    if (item.faltante > 0 or item.danado > 0) and item.unit_cost <= 0:
        issues.append(ValidationIssue(
            severity="WARNING",
            material_code=item.material_code,
            issue_type="PRECIO_CERO_EN_RECLAMO",
            message="El material tiene faltante o daño registrado pero su precio unitario es $0.00.",
            recommendation="Verificar el costo en la orden de compra ME80FN o factura para calcular la Nota de Crédito."
        ))

    # 5. Missing Vendor Style
    if not item.vendor_material:
        issues.append(ValidationIssue(
            severity="INFO",
            material_code=item.material_code,
            issue_type="ESTILO_PROVEEDOR_VACIO",
            message="No se encontró el código de estilo del proveedor.",
            recommendation="Revisar el registro info en tabla EINA o en el Packing List."
        ))

    return issues

def validate_all_items(items: List[PrepackItem]) -> Dict[str, List[ValidationIssue]]:
    """
    Validates the entire shipment and detects cross-material issues (like duplicates).
    Returns a dict with 'errors', 'warnings', and 'info'.
    """
    all_issues: List[ValidationIssue] = []
    
    # Track duplicates
    seen_codes: Dict[str, List[str]] = {}
    for item in items:
        seen_codes.setdefault(item.material_code, []).append(item.order_po_43)
        # Validate individual item
        all_issues.extend(validate_item_integrity(item))

    # Check for repeated materials in multiple POs
    for code, pos in seen_codes.items():
        if len(pos) > 1:
            all_issues.append(ValidationIssue(
                severity="INFO",
                material_code=code,
                issue_type="MATERIAL_EN_MULTIPLES_PO",
                message=f"El material {code} aparece en {len(pos)} órdenes de compra distintas: {', '.join(pos)}.",
                recommendation="Asegurarse de conciliar las cantidades sumadas entre ambas POs."
            ))

    categorized = {
        "ERROR": [i for i in all_issues if i.severity == "ERROR"],
        "WARNING": [i for i in all_issues if i.severity == "WARNING"],
        "INFO": [i for i in all_issues if i.severity == "INFO"]
    }
    return categorized
