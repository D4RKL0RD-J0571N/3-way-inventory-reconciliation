export type DiscrepancyType =
  | 'ExactMatch'
  | 'Shortage'
  | 'Overage'
  | 'RevueltoCandidate'
  | 'CartonMismatchOnly'
  | 'DamageOnly'
  | 'UnexpectedItem'
  | 'MissingItem'
  | 'ZeroTotal'
  | 'PendingPart'
  | 'BlindSetup';

export interface PackagingBreakdown {
  total_boxes: number;
  total_units: number;
  boxes_full: number;
  units_per_box_full: number;
  boxes_comp: number;
  units_per_box_comp: number;
  loose_units: number;
  raw_str: string;
  is_compound: boolean;
  has_loose: boolean;
}

export interface ChildSKU {
  material_code: string;
  description: string;
  ratio: number;
  theoretical_units: number;
  received_units: number;
  difference: number;
  unit_cost: number;
}

export interface PrepackItem {
  source_row: number;
  material_code: string;
  vendor_material: string;
  description: string;
  category: string;
  order_po_43: string;
  quantity_po_43_prk: number;
  quantity_po_43_units: number;
  quantity_inv_41_prk: number;
  quantity_inv_41_units: number;
  received_real_units: number;
  curve_total: number;
  unit_cost: number;
  container_part: string;
  sap_mov_101: number;
  sap_mov_641: number;
  children: ChildSKU[];
  faltante: number;
  danado: number;
  sobrante: number;
  revuelto: number;
  revuelto_saliendo: number;
  revuelto_destino: string;
  revuelto_entrando: number;
  revuelto_origen: string;
  revuelto_partner?: string;
  is_zero_total: boolean;
  discrepancy_type: DiscrepancyType;
  discrepancy_label: string;
  status: string;
  ncr_amount: number;
  observations: string;
  bultos_descarga: number;
  unidades_descarga: number;
  bultos_packing_list: number;
  unidades_packing_list: number;
  empaque_str: string;
  canonical_packaging: string;
  bultos_str: string;
  um: string;
  bultos_enteros: number;
  unidades_por_bulto_entero: number;
  bultos_complementos: number;
  unidades_por_bulto_comp: number;
  loose_units: number;
  unidades_por_bulto: number;
  dif_bultos: number;
  quantity_flags: string[];
  is_blind_setup: boolean;
}

export interface ReconciliationSummary {
  import_id: string;
  vendor_name: string;
  invoice_number: string;
  total_materials: number;
  active_styles_count: number;
  zero_total_count: number;
  clean_materials_count: number;
  part_a_materials: number;
  part_b_materials: number;
  total_ped_43_units: number;
  total_ped_41_units: number;
  total_received_units: number;
  total_faltante_units: number;
  total_danado_units: number;
  total_sobrante_units: number;
  total_shortage_units: number;
  total_overage_units: number;
  revuelto_pairs_count: number;
  total_ncr_amount: number;
  status: string;
  theor_packages: number;
  received_packages: number;
  difference_packages: number;
  container_parts_breakdown: Record<string, number>;
  total_bultos_descarga: number;
  total_bultos_pl: number;
  dif_total_bultos: number;
  total_prk_41: number;
  total_prk_43: number;
  is_blind_setup: boolean;
  vba_present?: boolean;
  vba_sha256?: string;
}

export interface ValidationIssue {
  severity: 'ERROR' | 'WARNING' | 'INFO';
  material_code: string;
  issue_type: string;
  message: string;
  recommendation: string;
}

export interface ValidationReport {
  ERROR: ValidationIssue[];
  WARNING: ValidationIssue[];
  INFO: ValidationIssue[];
}

export interface BomEntry {
  child: string;
  ratio: number;
}

export interface QcDetail {
  desc: string;
  um: string;
  qty: number;
  curve: number;
  bultos: number;
  unidades: number;
  bultos_enteros: number;
  unidades_por_bulto_entero: number;
  bultos_complementos: number;
  unidades_por_bulto_comp: number;
  loose_units: number;
  total_prk: number;
  part: string;
  observations: string;
  flags: string[];
}

export interface CierreMetadata {
  import_id: string;
  vendor_name: string;
  invoice_number: string;
  theor_packages: number;
  received_packages: number;
  difference_packages: number;
}

export interface AuditData {
  items: PrepackItem[];
  summary: ReconciliationSummary;
  boms: Record<string, BomEntry[]>;
  qcDetails: Record<string, QcDetail>;
  validation: ValidationReport;
  warnings: string[];
  cierreRawBytes: ArrayBuffer | null;
  cierreFileName: string;
  qcFileName: string;
  isBlindSetup: boolean;
  elapsedSeconds: number;
}

export type Language = 'es' | 'en';
