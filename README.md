# 📦 Sistema de Conciliación de Inventarios 3-Way (Recepción UPC)
### Cadena de Tiendas Ekono / Almacema S.A. — Centro de Distribución C001

Este proyecto proporciona un asistente automatizado de cálculo, explosión de curvas y auditoría de inventario para la fase de **RECEPCIÓN Y CIERRES** de importaciones (contenedores `IMP-XXX`), contrastando las órdenes de compra (PO `43...`), facturas de nacionalización (`41...`), descargas de muelle/calidad y movimientos contables de SAP S/4HANA & EWM (`Mov 101`).

---

## 🚀 Inicio Rápido (En 1 Clic)

* **Abrir Aplicación Web Interactiva:**  
  Haz doble clic en `iniciar_asistente.bat` (o ejecuta `streamlit run app.py` en la terminal).
* **Auditoría Rápida por Consola (1 Segundo):**  
  Haz doble clic en `ver_resumen_rapido.bat` (o ejecuta `python cli.py` en la terminal).

---

## 📁 Estructura del Proyecto

```
3 Way Inventory Reconciliation/
│
├── 🚀 iniciar_asistente.bat            # Lanzador rápido de la aplicación web
├── ⚡ ver_resumen_rapido.bat            # Lanzador rápido del resumen por consola
├── 🖥️ app.py                           # Punto de entrada de la interfaz gráfica (Streamlit)
├── ⚡ cli.py                            # Punto de entrada para auditoría por terminal
├── 📖 README.md                         # Documentación general del repositorio
│
├── 📁 data/                             # Archivos de datos organizados por importación
│   └── IMP-182/
│       ├── Copia de IMP-182 inprogress.xlsm        (Plantilla oficial de Cierre)
│       ├── IMP-182 Control del calidad.xlsm        (Conteo físico y BOMs/Curvas)
│       ├── CZC7933A Packing list 16JULY26.xls      (Packing List Proveedor Parte A)
│       └── CZC7933B Packing list 16JULY26.xls      (Packing List Proveedor Parte B)
│
├── 📁 docs/                             # Documentación operativa y manuales
│   ├── SOP - Procedimientos Operativos Conciliacion 3-Way.docx   (Manual Oficial)
│   └── drafts/                         # Apuntes y borradores de trabajo iniciales
│       ├── Borrador de procedimientos operativos.txt
│       └── BorradorRecepcion.txt
│
├── 📁 src/                              # Código fuente modular
│   ├── __init__.py
│   ├── config.py                       # Rutas centralizadas y constantes
│   ├── models.py                       # Modelos de datos (PrepackItem, ChildSKU)
│   ├── sanitizer.py                    # Limpieza defensiva de datos y strings sucios
│   ├── validator.py                    # Reglas de negocio y auditoría de integridad
│   ├── parsers.py                      # Extractores de datos desde Excel
│   └── engine.py                       # Motor de reconciliación 3-Way y cálculo NCR
│
└── 📁 tests/                            # Batería de pruebas automatizadas
    ├── __init__.py
    └── test_sanitizer.py               # Tests unitarios para sanitización y casos límite
```

---

## 📋 Módulos Principales de la Aplicación

1. **🔎 Ficha 360° del Material (Cero Scroll):**  
   Permite buscar cualquier código SAP y muestra en una sola pantalla:
   * Datos de compras (PO 43).
   * Datos de factura de nacionalización (41).
   * Conteo físico de muelle y bultos.
   * Movimientos contables en SAP (Mov 101).
   * Desglose de tallas hijas con sus curvas y estado en color verde/rojo.
2. **📊 Conciliación de Recepción:**  
   Tabla completa con los 136 materiales, filtros por Contenedor (Parte A vs Parte B) y estado de discrepancia.
3. **🔍 Explosión de Curvas y Tallas:**  
   Inspección matemática de los Prepacks (`PRK`) y el reparto a unidades individuales (`UN`).
4. **🖨️ Checklist de Impresión (Expediente Presidencia):**  
   Orden estricto de los 7 documentos físicos que se imprimen para entregar a Presidencia y a los dueños.
5. **📑 Generador para Excel Oficial:**  
   Genera los valores limpios y verificados en el orden exacto de columnas de `CIERRE RECEP` para pegarlos en el libro corporativo sin macros pesadas.
6. **📋 Pegar desde Portapapeles:**  
   Permite pegar directamente filas copiadas de SAP con `Ctrl + C` / `Ctrl + V`.

---

## 🧪 Pruebas Unitarias

Para ejecutar la batería de pruebas de robustez:
```bash
python -m pytest tests/test_sanitizer.py
```
