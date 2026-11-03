"""
Command-line interface for fast reconciliation audits.
Agnostic supply chain auditor supporting multi-container manual selection.
"""
import sys
import os
import argparse
from pathlib import Path

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from src.config import list_available_local_imports, resolve_local_import_files
from src.parsers import parse_quality_control_boms, parse_cierre_recepcion
from src.engine import compute_summary

def main():
    parser = argparse.ArgumentParser(description="3-Way Inventory Reconciliation & Audit Engine (CLI)")
    parser.add_argument("--import-id", "-i", type=str, help="ID of local import folder inside data/ (e.g. IMP-182)")
    parser.add_argument("--qc", type=str, help="Direct path to Quality Control spreadsheet (.xlsm / .xlsx)")
    parser.add_argument("--cierre", type=str, help="Direct path to Closure spreadsheet (.xlsm / .xlsx)")
    args = parser.parse_args()

    print("=" * 70)
    print("  3-WAY INVENTORY RECONCILIATION & AUDIT ENGINE")
    print("=" * 70)

    qc_path = None
    cierre_path = None
    import_label = "IMPORT"

    if args.qc and args.cierre:
        qc_path = Path(args.qc)
        cierre_path = Path(args.cierre)
        import_label = qc_path.parent.name
    elif args.import_id:
        qc_path, cierre_path = resolve_local_import_files(args.import_id)
        import_label = args.import_id
    else:
        # Check local folders available in data/
        available = list_available_local_imports()
        if available:
            print(f"Contenedores locales detectados en data/: {', '.join(available)}")
            selected = available[0]
            print(f"Cargando automáticamente contenedor local: '{selected}'")
            print("(Para seleccionar otro: python cli.py -i <ID> o pasar --qc y --cierre)\n")
            qc_path, cierre_path = resolve_local_import_files(selected)
            import_label = selected
        else:
            print("No se encontraron carpetas de importación en el directorio data/.")
            print("\nUso manual:")
            print("  python cli.py -i IMP-182")
            print("  python cli.py --qc <ruta_qc.xlsm> --cierre <ruta_cierre.xlsm>")
            sys.exit(0)

    if not qc_path or not qc_path.exists() or not cierre_path or not cierre_path.exists():
        print(f"Error: No se localizaron los archivos requeridos para '{import_label}'.")
        sys.exit(1)

    print(f"Cargando y parseando datos de importación: {import_label}...")
    print(f"  • Control de Calidad : {qc_path.name}")
    print(f"  • Plantilla Cierre   : {cierre_path.name}")
    
    boms, qc_details = parse_quality_control_boms(str(qc_path))
    items = parse_cierre_recepcion(str(cierre_path), boms, qc_details)
    summary = compute_summary(items)
    
    print("\nRESUMEN EJECUTIVO:")
    print(f"  Contenedor / Importación : {summary.import_id}")
    print(f"  Proveedor Internacional  : {summary.vendor_name}")
    print(f"  Factura / Ped. Nac (41)  : {summary.invoice_number}")
    print(f"  Total Materiales         : {summary.total_materials}")
    print(f"    - Parte A (Muelle 1)   : {summary.part_a_materials} materiales")
    print(f"    - Parte B (Muelle 2)   : {summary.part_b_materials} materiales")
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
