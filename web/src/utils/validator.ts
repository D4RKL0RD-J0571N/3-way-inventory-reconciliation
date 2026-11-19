import type { PrepackItem, ValidationIssue, ValidationReport } from '../types/reconciliation';

/**
 * Data validator and business rule integrity checker.
 * Ported faithfully from src/validator.py.
 */
export function validateItemIntegrity(item: PrepackItem): ValidationIssue[] {
  const issues: ValidationIssue[] = [];

  // 1. Curve Sum Integrity Check
  if (item.children.length > 0) {
    const sumRatios = item.children.reduce((acc, c) => acc + c.ratio, 0);
    if (sumRatios !== item.curve_total) {
      issues.push({
        severity: 'ERROR',
        material_code: item.material_code,
        issue_type: 'CURVA_DESCUADRADA',
        message: `La suma de ratios de tallas (${sumRatios}) no coincide con la curva oficial del Prepack (${item.curve_total}).`,
        recommendation: 'Verificar la lista de materiales en SAP transacción STPO/MAST.',
      });
    }
  }

  // 2. Missing BOM Check
  if (
    item.curve_total > 1 &&
    item.children.length === 0 &&
    (item.received_real_units > 0 || item.container_part === '' || item.container_part === 'A')
  ) {
    issues.push({
      severity: 'WARNING',
      material_code: item.material_code,
      issue_type: 'BOM_FALTANTE',
      message: `El material es un Prepack de ${item.curve_total} unidades pero no se encontró desglose de tallas hijas.`,
      recommendation: 'Confirmar si la tabla STPO fue consultada para este código de material.',
    });
  }

  // 3. Unit of Measure Confusion
  if (item.quantity_inv_41_units > 0 && item.received_real_units > item.quantity_inv_41_units * 2.5) {
    issues.push({
      severity: 'ERROR',
      material_code: item.material_code,
      issue_type: 'POSIBLE_CONFUSION_UDM',
      message: `Unidades recibidas (${item.received_real_units.toLocaleString()}) superan en más del 250% lo facturado (${item.quantity_inv_41_units.toLocaleString()}).`,
      recommendation: 'Verificar si en el piso de descarga se digitaron piezas individuales en vez de bultos/prepacks.',
    });
  }

  // 4. Zero Unit Cost with Shortage/Damage
  if ((item.faltante > 0 || item.danado > 0) && item.unit_cost <= 0) {
    issues.push({
      severity: 'WARNING',
      material_code: item.material_code,
      issue_type: 'PRECIO_CERO_EN_RECLAMO',
      message: 'El material tiene faltante o daño registrado pero su precio unitario es $0.00.',
      recommendation: 'Verificar el costo en la orden de compra ME80FN o factura para calcular la Nota de Crédito.',
    });
  }

  // 5. Missing Vendor Style
  if (!item.vendor_material) {
    issues.push({
      severity: 'INFO',
      material_code: item.material_code,
      issue_type: 'ESTILO_PROVEEDOR_VACIO',
      message: 'No se encontró el código de estilo del proveedor.',
      recommendation: 'Revisar el registro info en tabla EINA o en el Packing List.',
    });
  }

  return issues;
}

export function validateAllItems(items: PrepackItem[]): ValidationReport {
  const allIssues: ValidationIssue[] = [];
  const seenCodes: Record<string, string[]> = {};

  for (const item of items) {
    if (!seenCodes[item.material_code]) seenCodes[item.material_code] = [];
    seenCodes[item.material_code].push(item.order_po_43);
    allIssues.push(...validateItemIntegrity(item));
  }

  for (const [code, pos] of Object.entries(seenCodes)) {
    if (pos.length > 1) {
      allIssues.push({
        severity: 'INFO',
        material_code: code,
        issue_type: 'MATERIAL_EN_MULTIPLES_PO',
        message: `El material ${code} aparece en ${pos.length} órdenes de compra distintas: ${pos.join(', ')}.`,
        recommendation: 'Asegurarse de conciliar las cantidades sumadas entre ambas POs.',
      });
    }
  }

  return {
    ERROR: allIssues.filter((i) => i.severity === 'ERROR'),
    WARNING: allIssues.filter((i) => i.severity === 'WARNING'),
    INFO: allIssues.filter((i) => i.severity === 'INFO'),
  };
}
