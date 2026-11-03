"""
Command-line interface for fast reconciliation audits.
"""
import sys
import os

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from src.config import DEFAULT_QC_FILE, DEFAULT_CIERRE_FILE, DEFAULT_IMPORT_ID
from src.parsers import parse_quality_control_boms, parse_cierre_recepcion
from src.engine import compute_summary

def main():
    print("=" * 70)
    print("  3-WAY INVENTORY RECONCILIATION ENGINE — EKONO / ALMACEMA")
    print("=" * 70)
    print(f"Cargando y parseando datos de importación {DEFAULT_IMPORT_ID}...")
    
    boms, qc_details = parse_quality_control_boms(str(DEFAULT_QC_FILE))
    items = parse_cierre_recepcion(str(DEFAULT_CIERRE_FILE), boms, qc_details)
    summary = compute_summary(items)
    
    print("\nRESUMEN EJECUTIVO:")
    print(f"  Contenedor / Importación : {summary.import_id}")
    print(f"  Proveedor Internacional  : {summary.vendor_name}")
    print(f"  Factura / Ped. Nac (41)  : {summary.invoice_number}")
    print(f"  Total Materiales         : {summary.total_materials}")
    print(f"    - Parte A (Cont. 182-A): {summary.part_a_materials} materiales")
    print(f"    - Parte B (Cont. 182-B): {summary.part_b_materials} materiales")
    print("-" * 70)
    print(f"  Unidades Pedido PO (43)  : {summary.total_ped_43_units:,.0f} UN")
    print(f"  Unidades Facturadas (41) : {summary.total_ped_41_units:,.0f} UN")
    print(f"  Unidades Recibidas       : {summary.total_received_units:,.0f} UN")
    print(f"  Total Faltante           : {summary.total_faltante_units:,.0f} UN")
    print(f"  Total Dañado             : {summary.total_danado_units:,.0f} UN")
    print(f"  Total Sobrante           : {summary.total_sobrante_units:,.0f} UN")
    print(f"  Monto Total NCR ($)      : ${summary.total_ncr_amount:,.2f} USD")
    print(f"  Estado Global            : {summary.status}")
    print("=" * 70)
    
    print("\nMostrando primeros 5 materiales con desglose de tallas (curva):\n")
    for item in items[:5]:
        part_badge = f"[PARTE {item.container_part}]"
        print(f"{part_badge} Material: {item.material_code} | Estilo: {item.vendor_material} | {item.description}")
        print(f"       Pedido 41: {item.quantity_inv_41_units} UN ({item.quantity_inv_41_prk} PRK) | Recibido: {item.received_real_units} UN | Curva: {item.curve_total}")
        if item.children:
            print("       Desglose de Curva por Talla:")
            for c in item.children:
                print(f"         |-- Talla SKU: {c.material_code} | Ratio: {c.ratio} | Esperado: {c.theoretical_units} UN | Recibido: {c.received_units} UN | {c.description}")
        print()

if __name__ == "__main__":
    main()
