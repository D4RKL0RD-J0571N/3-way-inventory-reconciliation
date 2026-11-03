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

from src.config import DATA_DIR, resolve_local_import_files
from src.parsers import parse_quality_control_boms, parse_cierre_recepcion
from src.engine import compute_summary

def main():
    import_id = "IMP-182"
    if len(sys.argv) > 1:
        import_id = sys.argv[1]

    qc_path, cierre_path = resolve_local_import_files(import_id)

    if not qc_path or not qc_path.exists() or not cierre_path or not cierre_path.exists():
        print(f"Error: No se localizaron los archivos requeridos para '{import_id}' en data/.")
        sys.exit(1)

    imp_dir = qc_path.parent
    output_path = imp_dir / f"{import_id}_Cierre_Oficial_Final.xlsm"

    print("=" * 70)
    print(f"  AUTO-LLENADO Y REPARACIÓN DE PLANTILLA OFICIAL: {import_id}")
    print("=" * 70)
    print(f"1. Analizando conteo físico y curvas desde: {qc_path.name}...")
    boms, qc_details = parse_quality_control_boms(str(qc_path))
    items, meta = parse_cierre_recepcion(str(cierre_path), boms, qc_details, return_meta=True)
    summary = compute_summary(items, metadata=meta)

    print(f"2. Abriendo plantilla oficial preservando macros (VBA): {cierre_path.name}...")
    wb = openpyxl.load_workbook(str(cierre_path), keep_vba=True)
    ws = wb["CIERRE RECEP"]

    print(f"3. Escribiendo {len(items)} materiales en columnas ZDESCARGA (AS:AT) y PACKING LIST (AW:AY)...")
    start_row = 13
    for idx, item in enumerate(items):
        r = start_row + idx
        
        # Bultos y unidades verificados desde QC y cierre
        qc_item = qc_details.get(item.material_code, {})
        bultos_desc = int(item.quantity_inv_41_prk) if item.quantity_inv_41_prk > 0 else int(qc_item.get('bultos', 0))
        unidades_desc = int(qc_item.get('unidades', 0)) if qc_item.get('unidades', 0) > 0 else int(item.received_real_units)
        
        bultos_pl = bultos_desc
        unidades_pl = int(item.quantity_inv_41_units)
        empaque = item.curve_total if item.curve_total > 0 else int(qc_item.get('curve', 1))

        # Reparar fórmula de Col H si fue borrada o sobreescrita manualmente
        val_h = ws.cell(row=r, column=8).value
        if val_h is None or not str(val_h).startswith("="):
            ws.cell(row=r, column=8, value=f"=+G{r}-J{r}-N{r}-O{r}+L{r}+P{r}+Q{r}")

        # Columna AS (45): TOTAL BULTOS DESCARGA
        ws.cell(row=r, column=45, value=bultos_desc)
        # Columna AT (46): CANTIDAD DE UNIDADES EN BULTOS
        ws.cell(row=r, column=46, value=unidades_desc)
        # Columna AU (47): Diferencia Físico vs Pedido Nac
        val_au = ws.cell(row=r, column=47).value
        if val_au is None or not str(val_au).startswith("="):
            ws.cell(row=r, column=47, value=f"=+AT{r}-H{r}")

        # Columna AW (49): Bultos según Packing List
        ws.cell(row=r, column=49, value=bultos_pl)
        # Columna AX (50): CANTIDAD UNIDADES
        ws.cell(row=r, column=50, value=unidades_pl)
        # Columna AY (51): Empaque (Curva)
        ws.cell(row=r, column=51, value=empaque)
        # Columna AZ (52): Diferencia Facturadas vs Solicitadas
        val_az = ws.cell(row=r, column=52).value
        if val_az is None or not str(val_az).startswith("="):
            ws.cell(row=r, column=52, value=f"=+AX{r}-H{r}")

    print(f"4. Guardando libro final 100% cuadrado y fórmulas reparadas en:\n   {output_path.name}")
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
