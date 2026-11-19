import type { CierreMetadata, PrepackItem, ReconciliationSummary } from '../types/reconciliation';

export interface SummaryExtraOptions {
  revueltoPairsCount?: number;
  totalShortageUnits?: number;
  totalOverageUnits?: number;
  vbaPresent?: boolean;
  vbaSha256?: string;
}

/**
 * Reconciliation engine for 3-way matching and analytics.
 * Enforces Acr.Core rules:
 * - IsZeroTotal items excluded from ActiveStylesCount and CleanMaterialsCount.
 * - Cardinal Discrepancy Rule: shortages and overages NEVER cancel each other across materials.
 */
export function computeSummary(
  items: PrepackItem[],
  metadata?: Partial<CierreMetadata>,
  extra?: SummaryExtraOptions
): ReconciliationSummary {
  const meta = metadata ?? {};

  const finalImportId = meta.import_id || 'IMPORT-ACTIVA';
  const finalVendor = meta.vendor_name || 'PROVEEDOR INTERNACIONAL';
  const finalInvoice = meta.invoice_number || 'PENDIENTE';

  let theorPkg = meta.theor_packages ?? 0;
  let recPkg = meta.received_packages ?? 0;
  let difPkg = meta.difference_packages ?? 0;

  const activeItems = items.filter((i) => !i.is_zero_total);
  const zeroTotalCount = items.length - activeItems.length;
  const cleanMaterialsCount = activeItems.filter((i) => i.discrepancy_type === 'ExactMatch').length;

  const containerPartsBreakdown: Record<string, number> = {};
  for (const item of activeItems) {
    const p = (item.container_part || '').trim().toUpperCase();
    if (p) {
      containerPartsBreakdown[p] = (containerPartsBreakdown[p] ?? 0) + 1;
    }
  }

  const totalPed43 = items.reduce((s, i) => s + i.quantity_po_43_units, 0);
  const totalPed41 = items.reduce((s, i) => s + i.quantity_inv_41_units, 0);
  const totalRec = items.reduce((s, i) => s + i.received_real_units, 0);

  const totalFalt = items.reduce((s, i) => s + i.faltante, 0);
  const totalDan = items.reduce((s, i) => s + i.danado, 0);
  const totalSob = items.reduce((s, i) => s + i.sobrante, 0);
  const totalNcr = items.reduce((s, i) => s + i.ncr_amount, 0);

  // Cardinal Discrepancy Totals (shortages and overages NEVER offset across distinct materials)
  let totalShortageUnits = extra?.totalShortageUnits;
  let totalOverageUnits = extra?.totalOverageUnits;

  if (totalShortageUnits === undefined || totalOverageUnits === undefined) {
    let calcShort = 0;
    let calcOver = 0;
    for (const item of activeItems) {
      if (item.discrepancy_type === 'PendingPart' || item.discrepancy_type === 'BlindSetup') continue;
      const diffU = item.received_real_units - item.quantity_inv_41_units;
      if (diffU < 0) calcShort += Math.abs(diffU);
      else if (item.faltante > 0) calcShort += item.faltante;

      if (diffU > 0) calcOver += diffU;
      else if (item.sobrante > 0) calcOver += item.sobrante;
    }
    totalShortageUnits = totalShortageUnits ?? calcShort;
    totalOverageUnits = totalOverageUnits ?? calcOver;
  }

  const revueltoPairsCount =
    extra?.revueltoPairsCount ??
    Math.floor(activeItems.filter((i) => i.discrepancy_type === 'RevueltoCandidate').length / 2);

  const totalBultosDesc = items.reduce((s, i) => s + i.bultos_descarga, 0);
  const totalBultosPl = items.reduce((s, i) => s + i.bultos_packing_list, 0);
  const difTotalB = totalBultosDesc - totalBultosPl;
  const totalPrk41 = items.reduce((s, i) => s + i.quantity_inv_41_prk, 0);
  const totalPrk43 = items.reduce((s, i) => s + i.quantity_po_43_prk, 0);

  if (theorPkg === 0 && totalBultosPl > 0) theorPkg = totalBultosPl;
  if (recPkg === 0 && totalBultosDesc > 0) recPkg = totalBultosDesc;
  if (difPkg === 0 && (recPkg !== 0 || theorPkg !== 0)) difPkg = recPkg - theorPkg;

  const isBlind = items.some((i) => i.is_blind_setup || i.status === 'PENDIENTE_CONTROL_CALIDAD');
  const hasAnyDiscrepancy =
    totalShortageUnits > 0 ||
    totalOverageUnits > 0 ||
    totalDan > 0 ||
    revueltoPairsCount > 0 ||
    activeItems.some(
      (i) =>
        i.discrepancy_type === 'Shortage' ||
        i.discrepancy_type === 'Overage' ||
        i.discrepancy_type === 'MissingItem' ||
        i.discrepancy_type === 'UnexpectedItem' ||
        i.discrepancy_type === 'DamageOnly' ||
        i.discrepancy_type === 'RevueltoCandidate'
    );

  const status = isBlind
    ? 'PAUSA_SIN_CONTROL_CALIDAD'
    : !hasAnyDiscrepancy
      ? 'LIMPIO'
      : 'CON_DISCREPANCIAS';

  return {
    import_id: finalImportId,
    vendor_name: finalVendor,
    invoice_number: finalInvoice,
    total_materials: items.length,
    active_styles_count: activeItems.length,
    zero_total_count: zeroTotalCount,
    clean_materials_count: cleanMaterialsCount,
    part_a_materials: containerPartsBreakdown['A'] ?? 0,
    part_b_materials: containerPartsBreakdown['B'] ?? 0,
    total_ped_43_units: totalPed43,
    total_ped_41_units: totalPed41,
    total_received_units: totalRec,
    total_faltante_units: totalFalt,
    total_danado_units: totalDan,
    total_sobrante_units: totalSob,
    total_shortage_units: totalShortageUnits,
    total_overage_units: totalOverageUnits,
    revuelto_pairs_count: revueltoPairsCount,
    total_ncr_amount: totalNcr,
    status,
    theor_packages: theorPkg,
    received_packages: recPkg,
    difference_packages: difPkg,
    container_parts_breakdown: containerPartsBreakdown,
    total_bultos_descarga: totalBultosDesc,
    total_bultos_pl: totalBultosPl,
    dif_total_bultos: difTotalB,
    total_prk_41: totalPrk41,
    total_prk_43: totalPrk43,
    is_blind_setup: isBlind,
    vba_present: extra?.vbaPresent,
    vba_sha256: extra?.vbaSha256,
  };
}

export function filterItems(
  items: PrepackItem[],
  partFilter = 'ALL',
  statusFilter = 'ALL',
  categoryFilter = 'ALL',
  searchQuery = '',
  hideZeroTotal = false
): PrepackItem[] {
  let filtered = hideZeroTotal ? items.filter((i) => !i.is_zero_total) : items;

  if (partFilter !== 'ALL') {
    const pClean = partFilter.trim().toUpperCase();
    if (['NONE', 'SIN_PARTE', 'SINGLE', 'UNICO', 'ÚNICO'].includes(pClean)) {
      filtered = filtered.filter((i) => !(i.container_part || '').trim());
    } else {
      filtered = filtered.filter((i) => (i.container_part || '').trim().toUpperCase() === pClean);
    }
  }

  if (statusFilter !== 'ALL') {
    const sClean = statusFilter.trim().toUpperCase();
    filtered = filtered.filter(
      (i) =>
        (i.status || '').trim().toUpperCase() === sClean ||
        (i.discrepancy_label || '').trim().toUpperCase() === sClean
    );
  }

  if (categoryFilter !== 'ALL') {
    const cClean = categoryFilter.trim().toUpperCase();
    filtered = filtered.filter((i) => (i.category || '').trim().toUpperCase() === cClean);
  }

  if (searchQuery.trim()) {
    const q = searchQuery.toLowerCase().trim();
    filtered = filtered.filter((i) => {
      const parentMatch =
        i.material_code.toLowerCase().includes(q) ||
        i.vendor_material.toLowerCase().includes(q) ||
        i.description.toLowerCase().includes(q) ||
        i.order_po_43.toLowerCase().includes(q) ||
        i.category.toLowerCase().includes(q) ||
        (i.container_part || '').toLowerCase().includes(q) ||
        (i.status || '').toLowerCase().includes(q) ||
        (i.discrepancy_label || '').toLowerCase().includes(q) ||
        (i.canonical_packaging || '').toLowerCase().includes(q) ||
        (i.revuelto_partner || '').toLowerCase().includes(q);
      const childMatch = i.children.some(
        (c) => c.material_code.toLowerCase().includes(q) || c.description.toLowerCase().includes(q)
      );
      return parentMatch || childMatch;
    });
  }

  return filtered;
}
