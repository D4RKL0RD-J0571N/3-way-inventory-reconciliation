# 🛡️ Sistema de Conciliación de Inventarios 3-Way & Auditoría de Recepción

**Español** | [**English**](README.md) | [**Wiki Operativa y Tutorial**](docs/WIKI.md)

Un motor de cálculo y copiloto interactivo de nivel empresarial para la **Conciliación 3-Way de Inventarios y Cierres de Importación**. Diseñado para eliminar el cruce manual entre Órdenes de Compra (PO), Facturas Comerciales/DUA, Conteos Físicos de Calidad en Muelle y Movimientos Contables ERP (`Mov 101`).

---

## 🎯 Resumen Ejecutivo

En la logística de distribución internacional y centros de acopio, el cierre de un contenedor de importación requiere contrastar datos de fuentes heterogéneas:
1. **Capa Comercial / ERP:** Órdenes de Compra (`PO 43`) y Factura de Nacionalización (`Ped. 41`).
2. **Capa Física de Muelle:** Descarga de bultos/cartones y reporte de Control de Calidad con curvas de tallas.
3. **Capa Contable:** Ingresos a almacén en ERP (`SAP Mov 101`).

### La Problemática Tradicional
- **Desfase de Granularidad:** La factura y las macros corporativas operan al nivel de **Prepack Padre (`PRK`)**, mientras que el muelle audita **SKUs Hijos individuales (`UN`)**.
- **La Trampa de los `-51,926` en Excel:** Las macros iniciales de SAP no extraen los bultos ni unidades de muelle, dejando las columnas en cero y calculando diferencias artificiales masivas que obligan a los operadores a llenar cientos de celdas a mano celda por celda.
- **Fatiga Matemática:** Cruzar curvas complejas de 6 a 12 tallas por estilo provoca fácilmente errores de cálculo, notas de crédito improcedentes o atrasos en la liquidación ante Presidencia.

### La Solución: Copiloto ReconcileAI
- **Cruce Heurístico Determinista:** Agrupa tallas hijas bajo su lista de materiales (BOM) respetando la curva de compra ($100\%$ de precisión sin discrepancias).
- **Tiras de Pegado Rápido (`ZDESCARGA` & `PACKING LIST`):** Genera los bloques de texto listos para pegar en las columnas **`AS:AT`** y **`AW:AY`**, cuadrando la macro oficial a **cero diferencias en 10 segundos** sin modificar sus fórmulas.
- **Ficha 360° del Material (Cero Desplazamiento):** Muestra compras, factura, físico, curvas y costos en una sola pantalla visual.
- **Validador de Portapapeles Pre-Macro:** Audita y desactiva la notación científica (`1.45E+06`) y limpia textos copiados desde el GUI de SAP (`ME80FN`, `MB51`) antes de llevarlos a Excel.

---

## 🚀 Inicio Rápido (En 1 Clic)

### Opción A: Aplicación Web Interactiva
Haz doble clic en `iniciar_asistente.bat` o ejecuta:
```bash
streamlit run app.py
```
- Se abrirá en tu navegador en `http://localhost:8501`.
- **Cero datos corporativos por defecto:** La aplicación inicia con un explorador limpio. Puedes seleccionar una carpeta local o arrastrar los archivos de Calidad y Cierre directamente.

### Opción B: Auditoría por Consola (1 Segundo)
Haz doble clic en `ver_resumen_rapido.bat` o ejecuta:
```bash
python cli.py
```
O especifica archivos puntuales:
```bash
python cli.py -i IMP-182
python cli.py --qc ruta/a/QC.xlsm --cierre ruta/a/Cierre.xlsm
```

---

## 📁 Arquitectura del Repositorio

```
3-way-inventory-reconciliation/
│
├── 🚀 iniciar_asistente.bat            # Lanzador rápido en 1 clic para la App Web
├── ⚡ ver_resumen_rapido.bat            # Lanzador rápido para auditoría por consola
├── 🖥️ app.py                           # Interfaz Web en Streamlit con modo oscuro/claro
├── ⚡ cli.py                            # Motor CLI rápido para auditorías instantáneas
├── 📖 README.md                         # Documentación en Inglés
├── 📖 README.es.md                      # Documentación en Español
│
├── 📁 src/                              # Motor Central Agnóstico
│   ├── __init__.py
│   ├── config.py                       # Resolución dinámica de rutas y carpetas
│   ├── models.py                       # Clases de datos (PrepackItem, ChildSKU)
│   ├── sanitizer.py                    # Limpieza defensiva de notación científica y texto
│   ├── validator.py                    # Validación de curvas matemáticas y estados NCR
│   ├── parsers.py                      # Extractores multilibro de Excel (openpyxl / xlrd)
│   └── engine.py                       # Motor de cruce 3-Way y cálculo de Notas de Crédito
│
├── 📁 docs/                             # Documentación y Manuales
│   ├── WIKI.md                         # Guía de Inicio Paso a Paso y Tutorial Operativo
│   └── SOP_Conciliacion_3Way.md        # Procedimiento Operativo Estándar (SOP)
│
└── 📁 tests/                            # Batería de Pruebas Automatizadas
    ├── __init__.py
    └── test_sanitizer.py               # Tests unitarios (100% éxito)
```

---

## 🧪 Pruebas Unitarias

Para ejecutar la batería de pruebas de robustez:
```bash
python -m pytest tests/test_sanitizer.py
```

---

## 📄 Licencia

Desarrollado como un software independiente de automatización y auditoría logística. Libre para uso profesional y educativo.
