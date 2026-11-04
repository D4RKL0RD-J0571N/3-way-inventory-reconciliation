# 🛡️ ReconcileAI — Conciliación de Inventarios 3-Way & Auditoría de Recepción
### Copiloto Inteligente para Logística, Control de Calidad y Cierres de Importación

**Español** | [**English**](README.md) | [**Wiki Operativa & Tutorial**](docs/WIKI.md) | [**Notas de la Versión**](docs/RELEASE_NOTES_DEMO.md) | [**SOP Recepción SAP**](docs/SOP_RECEPCION_SAP.md) | [**Auditoría de Bultos & VBA**](docs/CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md) | [**Arquitectura de Despacho**](docs/AUDITORIA_DESPACHO_MACRO.md)

---

## 🌟 Edición Portable Independiente (v1.1.0)

> [!TIP]
> **Cero Instalación:** ReconcileAI se distribuye como un **paquete ejecutable portable** (`ReconcileAI_Portable.zip`). Funciona de inmediato en Windows 10/11 **sin requerir instalación previa de Python ni permisos de Administrador**.

Un motor de cálculo y copiloto interactivo de nivel empresarial para la **Conciliación 3-Way de Inventarios y Cierres de Importación**. Diseñado para eliminar el cruce manual entre Órdenes de Compra (PO), Facturas Comerciales/DUA, Conteos Físicos de Calidad en Muelle y Movimientos Contables ERP (`Mov 101` de Recepción y `Mov 641` de Despacho/Traslado).

```mermaid
flowchart LR
    A["Compras ERP (PO 43)"] --> D["Copiloto ReconcileAI"]
    B["Factura Comercial (Ped 41)"] --> D
    C["Descarga Física y Curvas QC"] --> D
    
    D --> E["📦 Cuadre & Desglose de Bultos"]
    D --> F["🔍 Ficha 360° del Material"]
    D --> G["🚦 Grilla Semáforo por Estado"]
    D --> H["⚡ Tiras de Pegado Rápido (AS:AZ)"]
    D --> I["💾 Descarga Libro Oficial .xlsm"]
    
    H -->|"Pegado en 10 seg"| J["Libro Oficial (.xlsm)"]
    I -->|"Descarga Directa"| J
    J --> K["Cierre 100% Cuadrado a Cero Diferencias"]
```

---

## 🎯 Resumen Ejecutivo & Propuesta de Valor

En la logística de distribución internacional y centros de distribución, el cierre de un contenedor de importación requiere contrastar datos de fuentes heterogéneas:
1. **Capa Comercial / ERP:** Órdenes de Compra (`PO 43`) y Factura de Nacionalización (`Ped. 41`).
2. **Capa Física de Muelle:** Descarga de bultos/cartones y reporte de Control de Calidad con curvas de tallas.
3. **Capa Contable:** Ingresos a almacén en ERP (`SAP Mov 101` para Recepción y `SAP Mov 641` para Despacho a Tiendas).

### Los 3 Niveles Numéricos Independientes en la Cadena de Suministro
ReconcileAI impone una estricta segregación matemática entre los tres niveles de inventario:
* **Nivel 1 — Bultos Físicos (Cartones/Cajas):** Cajas físicas descargadas en muelle (Col `AS`), bultos del Packing List del proveedor (Col `AW`), y conteo de Control de Calidad (`Bultos Enteros` + `Complementos`).
* **Nivel 2 — Prepacks / Curvas Comerciales (PRK / Curva):** Unidad de compra en pedidos mayoristas (Cols `BD` y `BE` en PRKs, Col `BB` proporción de curva, Col `BC` unidad de medida).
* **Nivel 3 — Unidades Totales (Prendas/Piezas):** Unidades facturadas, nacionalizadas y recibidas prenda por prenda (Cols `F`, `G`, `H`, `AT`, `AX`).

### Capacidades Operativas Clave
- **Parseo de Empaque Compuesto:** Interpreta números escalares (`42`) y desglosa automáticamente cadenas de bultos heterogéneos con picos (ej. `'39*60 1*78'` $\rightarrow$ 39 cajas de 60 + 1 caja de 78 = 40 bultos, 2,418 unidades).
- **Pestaña Dedicada de Cuadre de Bultos (`📦 Cuadre & Desglose de Bultos`):** Top 5 tarjetas KPI de bultos, calculadora interactiva de empaque, y réplica exacta 1-a-1 de la hoja `'Bultos'` del libro corporativo.
- **Arquitectura Dual (Recepción y Despacho):** Estructura sincronizada para Cierres de Recepción (`CIERRE RECEP`) y Cierres de Despacho (`CIERRE DESP`, `Check list DESPACHO`, `STOCK CONTROL DE INVENTARIO`, `Mov 641`).
- **Cruce Heurístico Determinista:** Agrupa tallas hijas bajo su lista de materiales (BOM) respetando la curva de compra ($100\%$ de precisión sin discrepancias).
- **Auto-Reparación de Fórmulas Corruptas:** Detecta fórmulas borradas o sobreescritas por error en la plantilla oficial (ej. columna H, AU y AZ) y las restaura automáticamente.
- **Tiras de Pegado Rápido (`ZDESCARGA` & `PACKING LIST`):** Genera los bloques de texto listos para pegar en las columnas **`AS:AT`** y **`AW:AY`**, cuadrando la macro oficial a **cero diferencias en 10 segundos** sin modificar fórmulas.
- **Exportación en 1 Clic del Libro Oficial (.xlsm):** Genera directamente en memoria el archivo Excel oficial cuadrado, preservando intactas todas las macros VBA y el formato original.
- **Interfaz Bilingüe (Español por Defecto / Inglés Opcional):** Traducción completa de UI y traducción dinámica en el cliente de los tooltips de barra de herramientas en tablas.
- **Ficha 360° del Material (Cero Desplazamiento):** Muestra compras, factura, físico, curvas y costos en una sola pantalla visual.
- **Validador de Portapapeles Pre-Macro:** Audita y desactiva la notación científica (`1.45E+06`) y limpia textos copiados desde el GUI de SAP (`ME80FN`, `MB51`) antes de llevarlos a Excel.

---

## 🚀 Inicio Rápido (En 1 Clic)

### Opción A: Paquete Portable Independiente (Recomendado para Demo)
*Ideal para computadoras de oficina o bodega sin Python ni permisos de Administrador.*
1. Descomprime **`ReconcileAI_Portable.zip`** (o utiliza la carpeta `dist/ReconcileAI`).
2. Haz doble clic en **`ReconcileAI.exe`** (o en `INICIAR_RECONCILEAI.bat`).
3. Tu navegador web se abrirá automáticamente en `http://localhost:8501`.

### Opción B: Ejecución con Python Local
*Para desarrolladores o personalización del código.*
```bash
# 1. Configurar entorno y dependencias automáticamente
preparar_equipo.bat

# 2. Iniciar el copiloto web
iniciar_asistente.bat
# o directamente: streamlit run app.py
```

### Opción C: Auditoría Ultrarrápida por Consola (CLI)
*Para auditorías directas en 1 segundo por terminal.*
```bash
python cli.py -i CONT-001
# o especificando archivos directos:
python cli.py --qc ruta/a/Control_Calidad_Muelle.xlsm --cierre ruta/a/Libro_Cierre_Oficial.xlsm
```

---

## 📋 Guía de Demostración en 2 Minutos

1. **Abrir la Aplicación:** Abre `http://localhost:8501`.
2. **Seleccionar Contenedor:** En el explorador de archivos, elige una carpeta local (ej. `data/CONT-001` o tu carpeta de importación) o arrastra tus dos libros de Excel:
   - Archivo 1: Control de Calidad / Recepción de Muelle (`Control_Calidad_Muelle.xlsm`)
   - Archivo 2: Plantilla Oficial de Cierre (`Libro_Cierre_Oficial.xlsm`)
3. **Ejecutar Auditoría:** Haz clic en **`🚀 Iniciar Auditoría de estos Archivos`**.
4. **Auditoría de Bultos y Empaque:**
   - Dirígete a **`📦 Cuadre & Desglose de Bultos`** para comparar bultos de muelle vs packing list, usar la calculadora interactiva de empaque y verificar la réplica de la hoja `'Bultos'`.
5. **Inspección de Materiales y Semáforo:**
   - Usa la **`🔍 Ficha 360° Material`** para inspeccionar curvas padre-hijo y costos unitarios sin desplazamiento.
   - Revisa la **`🚦 Tabla Semáforo`**: Verde = 100% Cuadrado; Azul = Parte B (en tránsito o descargado en segundo lote); Rojo = Faltante real en muelle.
6. **Cuadrar el Cierre Oficial:**
   - Dirígete a la pestaña **`📋 Copiar a Excel Oficial`**.
   - Haz clic en **`⚡ Descargar Libro Oficial Cuadrado (.xlsm con Macros VBA)`** para obtener el archivo listo generado en memoria, o copia las tiras de texto para **`ZDESCARGA` (Cols AS:AT)** y **`PACKING LIST` (Cols AW:AY)** en tu libro oficial.
   - Resultado: La columna de diferencias en Excel pasa inmediatamente a **`0`**.

---

## 📁 Arquitectura del Repositorio

```
3-way-inventory-reconciliation/
│
├── 🚀 iniciar_asistente.bat            # Lanzador rápido Windows (fallback inteligente)
├── ⚡ ver_resumen_rapido.bat            # Lanzador para auditoría por consola (1 seg)
├── 📦 preparar_equipo.bat               # Instalador automático de entorno virtual y dependencias
├── 🖥️ app.py                           # Interfaz Web en Streamlit con modo fluido Widescreen
├── ⚡ cli.py                            # Motor CLI rápido para auditorías instantáneas
├── 📄 generar_cierre_oficial.py        # Generador de .xlsm con macros VBA intactas
├── 🏃 run_standalone.py                # Lanzador Python independiente para PyInstaller
├── 📖 README.md                         # Documentación en Inglés
├── 📖 README.es.md                      # Documentación en Español
│
├── 📁 dist/                             # Entregables Compilados Portables (Ignorados por Git)
│   ├── ReconcileAI/                    # Carpeta portable independiente
│   │   ├── ReconcileAI.exe             # Ejecutable nativo para Windows
│   │   ├── INICIAR_RECONCILEAI.bat     # Lanzador amigable en lote
│   │   └── data/                       # Carpeta local para colocar importaciones
│   └── ReconcileAI_Portable.zip        # Archivo ZIP de distribución independiente
│
├── 📁 src/                              # Motor Central Agnóstico
│   ├── __init__.py
│   ├── config.py                       # Resolución dinámica de rutas y entorno PyInstaller
│   ├── models.py                       # Clases de datos (PrepackItem, PackagingBreakdown, SummaryStats)
│   ├── sanitizer.py                    # Limpieza defensiva de notación científica y texto
│   ├── validator.py                    # Validación de curvas matemáticas y estados NCR
│   ├── parsers.py                      # Escaneo dinámico de columnas y parser de empaque compuesto
│   ├── engine.py                       # Motor de cruce 3-Way, balance de bultos y cálculo NCR
│   ├── i18n.py                         # Internacionalización centralizada (ES/EN)
│   └── ui_helpers.py                   # Shimmers CSS responsivos y paginación virtual
│
├── 📁 docs/                             # Manuales Operativos y Arquitectura Técnica
│   ├── CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md # Base de conocimiento sobre bultos y macros VBA
│   ├── AUDITORIA_DESPACHO_MACRO.md     # Arquitectura de cierre de despacho y sincronización Mov 641
│   ├── WIKI.md                         # Guía de Inicio Paso a Paso y Tutorial Operativo
│   ├── RELEASE_NOTES_DEMO.md           # Notas de la Versión
│   └── SOP_RECEPCION_SAP.md            # Procedimiento Operativo Estándar de Recepción SAP
│
├── 📁 tools/                            # Herramientas de Compilación
│   └── build_exe.py                    # Script de empaquetado PyInstaller y generación ZIP
│
└── 📁 tests/                            # Batería de Pruebas Automatizadas (22/22 Pasando)
    ├── __init__.py
    ├── test_sanitizer.py               # Tests unitarios de sanitización
    ├── test_parsers.py                 # Tests de extracción de curvas y resistencia a columnas
    ├── test_i18n.py                    # Tests de traducción y paginación
    └── test_bultos_empaque.py          # Tests de desglose de bultos y empaque compuesto
```

---

## 🔒 Privacidad, Seguridad y Confidencialidad

- **Diseño 100% Agnóstico:** Funciona con cualquier proveedor, cualquier ERP (SAP, Oracle, NetSuite) y cualquier centro logístico.
- **Sin Datos Confidenciales Embebidos:** El código y los repositorios no contienen nombres corporativos, credenciales ni información privada.
- **Operación 100% en Memoria Local:** Los cálculos se ejecutan exclusivamente en la memoria RAM de tu computadora. Ningún dato viaja a internet ni se almacena en servidores externos.

---

## 📄 Licencia y Créditos

Desarrollado como una iniciativa independiente de automatización y auditoría logística de cadena de suministro. Libre para uso operativo, profesional y académico bajo Licencia MIT.
