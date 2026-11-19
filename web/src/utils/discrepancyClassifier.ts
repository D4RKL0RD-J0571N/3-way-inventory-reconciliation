import type { DiscrepancyType, PrepackItem } from '../types/reconciliation';

/**
 * Deterministic 3-Way Discrepancy Classifier ported from Acr.Core/Services/DiscrepancyClassifier.cs.
 * Enforces:
 * 1. IsZeroTotal -> SIN MOVIMIENTO (0 TOTAL) (never misclassified as green CUADRADO)
 * 2. Cardinal Discrepancy Rule: Shortages and overages NEVER offset across distinct materials.
 * 3. Greedy 1:1 exact-delta RevueltoCandidate pairing without altering aggregate shortage/overage totals.
 */

export function isItemZeroTotal(item: Pick<
  PrepackItem,
  'quantity_inv_41_units' | 'received_real_units' | 'bultos_packing_list' | 'bultos_descarga' | 'danado'
>): boolean {
  return (
    item.quantity_inv_41_units === 0 &&
    item.received_real_units === 0 &&
    item.bultos_packing_list === 0 &&
    item.bultos_descarga === 0 &&
    item.danado === 0
  );
}

export function getDiscrepancyLabel(type: DiscrepancyType, containerPart?: string): string {
  switch (type) {
    case 'ZeroTotal':
      return 'SIN MOVIMIENTO (0 TOTAL)';
    case 'ExactMatch':
      return 'CUADRADO (0 DIF)';
    case 'Shortage':
      return 'FALTANTE';
    case 'Overage':
      return 'SOBRANTE';
    case 'RevueltoCandidate':
      return 'REVUELTO (CANDIDATO)';
    case 'DamageOnly':
      return 'DAÑADO';
    case 'CartonMismatchOnly':
      return 'DIF. BULTOS (UNID. OK)';
    case 'UnexpectedItem':
      return 'SOBRANTE NO FACTURADO';
    case 'MissingItem':
      return 'FALTANTE TOTAL';
    case 'PendingPart':
      return containerPart ? `PENDIENTE PARTE ${containerPart}` : 'PENDIENTE PARTE';
    case 'BlindSetup':
      return 'PAUSA (SIN CONTROL CALIDAD)';
    default:
      return 'REVISAR';
  }
}

export function classifyAllDiscrepancies(items: PrepackItem[], isBlindSetup: boolean): {
  revueltoPairsCount: number;
  totalShortageUnits: number;
  totalOverageUnits: number;
} {
  // Pass 1: individual classification
  for (const item of items) {
    const zeroTotal = isItemZeroTotal(item);
    item.is_zero_total = zeroTotal;

    if (isBlindSetup || item.is_blind_setup) {
      item.discrepancy_type = zeroTotal ? 'ZeroTotal' : 'BlindSetup';
      item.discrepancy_label = getDiscrepancyLabel(item.discrepancy_type);
      item.status = zeroTotal ? 'SIN_MOVIMIENTO' : 'PENDIENTE_CONTROL_CALIDAD';
      continue;
    }

    if (zeroTotal) {
      item.discrepancy_type = 'ZeroTotal';
      item.discrepancy_label = getDiscrepancyLabel('ZeroTotal');
      item.status = 'SIN_MOVIMIENTO';
      item.observations = 'Material sin movimiento (0 facturado / 0 recibido / 0 bultos).';
      item.ncr_amount = 0;
      continue;
    }

    // Pending subsequent container part (only when explicitly assigned to non-A part with 0 received)
    if (item.container_part && item.container_part !== 'A' && item.received_real_units === 0) {
      item.discrepancy_type = 'PendingPart';
      item.discrepancy_label = getDiscrepancyLabel('PendingPart', item.container_part);
      item.status = `PENDIENTE_PARTE_${item.container_part}`;
      item.observations = `Mercadería en Contenedor Parte ${item.container_part} o pendiente de descarga.`;
      item.ncr_amount = 0;
      continue;
    }

    const decl = item.quantity_inv_41_units;
    const phys = item.received_real_units;
    const diffU = phys - decl;
    const diffC = item.bultos_descarga - item.bultos_packing_list;

    // Explicit Revuelto recorded in CIERRE RECEP Col J / L
    if (item.revuelto_saliendo > 0 || item.revuelto_entrando > 0) {
      item.discrepancy_type = 'RevueltoCandidate';
      item.discrepancy_label = 'REVUELTO REGISTRADO';
      item.status = 'REVUELTO';
      item.revuelto = Math.max(item.revuelto_saliendo, item.revuelto_entrando);
      item.revuelto_partner = item.revuelto_destino || item.revuelto_origen || item.revuelto_partner;
      item.observations =
        item.revuelto_saliendo > 0
          ? `Revuelto saliendo: ${item.revuelto_saliendo.toLocaleString()} UN${item.revuelto_destino ? ` -> Destino ${item.revuelto_destino}` : ''}.`
          : `Revuelto entrando: ${item.revuelto_entrando.toLocaleString()} UN${item.revuelto_origen ? ` <- Origen ${item.revuelto_origen}` : ''}.`;
    } else if (decl > 0 && phys === 0) {
      item.discrepancy_type = 'MissingItem';
      item.discrepancy_label = getDiscrepancyLabel('MissingItem');
      item.status = 'FALTANTE';
      if (item.faltante === 0) item.faltante = decl;
      item.observations = `No recibido en muelle (faltante total de ${decl.toLocaleString()} unidades).`;
    } else if (decl === 0 && phys > 0) {
      item.discrepancy_type = 'UnexpectedItem';
      item.discrepancy_label = getDiscrepancyLabel('UnexpectedItem');
      item.status = 'SOBRANTE';
      if (item.sobrante === 0) item.sobrante = phys;
      item.observations = `Material recibido en muelle sin factura 41 (+${phys.toLocaleString()} unidades).`;
    } else if (diffU < 0 || item.faltante > 0) {
      const faltQty = item.faltante > 0 ? item.faltante : Math.abs(diffU);
      item.faltante = faltQty;
      item.discrepancy_type = 'Shortage';
      item.discrepancy_label = getDiscrepancyLabel('Shortage');
      item.status = 'FALTANTE';
      item.observations = `Faltante físico de ${faltQty.toLocaleString()} unidades respecto a Pedido 41.`;
    } else if (diffU > 0 || item.sobrante > 0) {
      const sobQty = item.sobrante > 0 ? item.sobrante : diffU;
      item.sobrante = sobQty;
      item.discrepancy_type = 'Overage';
      item.discrepancy_label = getDiscrepancyLabel('Overage');
      item.status = 'SOBRANTE';
      item.observations = `Sobrante físico de +${sobQty.toLocaleString()} unidades respecto a Pedido 41.`;
    } else if (item.danado > 0) {
      item.discrepancy_type = 'DamageOnly';
      item.discrepancy_label = getDiscrepancyLabel('DamageOnly');
      item.status = 'DAÑADO';
      item.observations = `Avería / daño físico registrado: ${item.danado.toLocaleString()} unidades.`;
    } else if (diffC !== 0 && item.bultos_packing_list > 0 && item.bultos_descarga > 0) {
      item.discrepancy_type = 'CartonMismatchOnly';
      item.discrepancy_label = getDiscrepancyLabel('CartonMismatchOnly');
      item.status = 'DIF_BULTOS';
      item.observations = `Unidades cuadradas (${phys.toLocaleString()} UN), diferencia solo en bultos (${diffC > 0 ? '+' : ''}${diffC}).`;
    } else {
      item.discrepancy_type = 'ExactMatch';
      item.discrepancy_label = getDiscrepancyLabel('ExactMatch');
      item.status = 'LIMPIO';
      item.observations = 'Conciliación exacta sin discrepancias.';
    }

    // Calculate NCR amount accurately (shortfall + damaged + revuelto saliendo if uncompensated)
    const claimableUnits =
      item.faltante + item.danado + (item.revuelto_saliendo > 0 && !item.revuelto_entrando ? item.revuelto_saliendo : 0);
    item.ncr_amount = Math.round(claimableUnits * item.unit_cost * 100) / 100;
  }

  // Pass 2: Greedy 1:1 exact-delta RevueltoCandidate pairing (never modifies totals!)
  let revueltoPairsCount = 0;
  const shortages = items.filter(
    (i) => !i.is_zero_total && (i.discrepancy_type === 'Shortage' || i.discrepancy_type === 'MissingItem')
  );
  const overages = items.filter(
    (i) => !i.is_zero_total && (i.discrepancy_type === 'Overage' || i.discrepancy_type === 'UnexpectedItem')
  );
  const usedOverages = new Set<number>();

  for (const s of shortages) {
    const sDelta = Math.abs(s.received_real_units - s.quantity_inv_41_units) || s.faltante;
    if (sDelta <= 0) continue;

    const matchIdx = overages.findIndex((o, idx) => {
      if (usedOverages.has(idx)) return false;
      const oDelta = Math.abs(o.received_real_units - o.quantity_inv_41_units) || o.sobrante;
      return oDelta === sDelta && o.material_code !== s.material_code;
    });

    if (matchIdx !== -1) {
      usedOverages.add(matchIdx);
      const o = overages[matchIdx];
      revueltoPairsCount++;

      s.discrepancy_type = 'RevueltoCandidate';
      s.discrepancy_label = 'REVUELTO (CANDIDATO)';
      s.status = 'REVUELTO';
      s.revuelto_partner = o.material_code;
      s.observations = `Posible Revuelto (-${sDelta} UN) compensado en magnitud por material ${o.material_code} (+${sDelta} UN). Verificar físicamente antes de ajustar.`;

      o.discrepancy_type = 'RevueltoCandidate';
      o.discrepancy_label = 'REVUELTO (CANDIDATO)';
      o.status = 'REVUELTO';
      o.revuelto_partner = s.material_code;
      o.observations = `Posible Revuelto (+${sDelta} UN) compensado en magnitud por material ${s.material_code} (-${sDelta} UN). Verificar físicamente antes de ajustar.`;
    }
  }

  // Pass 3: Cardinal Discrepancy Totals (shortages and overages NEVER cancel each other!)
  let totalShortageUnits = 0;
  let totalOverageUnits = 0;
  for (const item of items) {
    if (item.is_zero_total || item.discrepancy_type === 'PendingPart' || item.discrepancy_type === 'BlindSetup') {
      continue;
    }
    const diffU = item.received_real_units - item.quantity_inv_41_units;
    if (diffU < 0) totalShortageUnits += Math.abs(diffU);
    else if (item.faltante > 0) totalShortageUnits += item.faltante;

    if (diffU > 0) totalOverageUnits += diffU;
    else if (item.sobrante > 0) totalOverageUnits += item.sobrante;
  }

  return {
    revueltoPairsCount,
    totalShortageUnits,
    totalOverageUnits,
  };
}
