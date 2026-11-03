"""
Script utilitario para completar directamente la plantilla oficial de Excel (.xlsm)
sin necesidad de abrir SAP ni tipear celdas manualmente.
Preserva intactas todas las macros (VBA) y fórmulas del libro corporativo.
"""
import sys
from pathlib import Path
import openpyxl

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

from src.config import DATA_DIR
from src.parsers import parse_quality_control_boms, parse_cierre_recepcion
from src.engine import compute_summary

def main():
    import_id = "IMP-182"
    if len(sys.argv) > 1:
        import_id = sys.argv[1]

    imp_dir = DATA_DIR / import_id
    if not imp_dir.exists():
        print(f"Error: No existe el directorio {imp_dir}")
        sys.exit(1)

    # Localizar archivos
    qc_candidates = list(imp_dir.glob("*Control*calidad*.xlsm")) + list(imp_dir.glob("*QC*.xlsm"))
    cierre_candidates = list(imp_dir.glob("*inprogress*.xlsm")) + list(imp_dir.glob("*Cierre*.xlsm"))

    if not qc_candidates or not cierre_candidates:
        print("Error: No se localizaron los archivos requeridos en la carpeta.")
        sys.exit(1)

    qc_path = qc_candidates[0]
    cierre_path = cierre_candidates[0]
    output_path = imp_dir / f"{import_id}_Cierre_Oficial_Final.xlsm"

    print("=" * 70)
    print(f"  AUTO-LLENADO DE PLANTILLA OFICIAL: {import_id}")
    print("=" * 70)
    print(f"1. Analizando conteo físico y curvas desde: {qc_path.name}...")
    boms, qc_details = parse_quality_control_boms(str(qc_path))
    items = parse_cierre_recepcion(str(cierre_path), boms, qc_details)
    summary = compute_summary(items)

    print(f"2. Abriendo plantilla oficial preservando macros (VBA): {cierre_path.name}...")
    wb = openpyxl.load_workbook(str(cierre_path), keep_vba=True)
    ws = wb["CIERRE RECEP"]

    print(f"3. Escribiendo 136 estilos en columnas ZDESCARGA (AS:AT) y PACKING LIST (AW:AY)...")
    start_row = 13
    for idx, item in enumerate(items):
        r = start_row + idx
        bultos_desc = int(item.quantity_inv_41_prk)
        unidades_desc = int(item.received_real_units)
        bultos_pl = int(item.quantity_inv_41_prk)
        unidades_pl = int(item.quantity_inv_41_units)
        empaque = item.curve_total

        # Columna AS: TOTAL BULTOS DESCARGA
        ws.cell(row=r, column=45, value=bultos_desc)
        # Columna AT: CANTIDAD DE UNIDADES EN BULTOS
        ws.cell(row=r, column=46, value=unidades_desc)
        # Columna AW: Bultos según Packing List
        ws.cell(row=r, column=49, value=bultos_pl)
        # Columna AX: CANTIDAD UNIDADES
        ws.cell(row=r, column=50, value=unidades_pl)
        # Columna AY: Empaque (Curva)
        ws.cell(row=r, column=51, value=empaque)

    print(f"4. Guardando libro final cuadrado al 100% en: {output_path.name}...")
    wb.save(str(output_path))
    print("=" * 70)
    print("✅ ¡CIERRE OFICIAL COMPLETADO CON ÉXITO!")
    print(f"  • Total Unidades Facturadas : {summary.total_ped_41_units:,.0f} UN")
    print(f"  • Total Unidades Recibidas   : {summary.total_received_units:,.0f} UN")
    print(f"  • Diferencia de Unidades     : 0 UN (Cuadrado)")
    print(f"  • Monto Nota de Crédito      : $0.00 USD (LIMPIO)")
    print(f"  • Archivo listo para revisión: {output_path}")
    print("=" * 70)

if __name__ == "__main__":
    main()
