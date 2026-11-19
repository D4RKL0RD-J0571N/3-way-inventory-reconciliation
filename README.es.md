# 🛡️ ACR — Asistente de Cierre de Recepción & Conciliador de Inventarios
### Aplicación de Escritorio Empresarial (.NET 10 LTS Win-x64) para Cadena de Suministro y Cierres de Importación

**Español** | [**English**](README.md) | [**Wiki Operativa & Tutorial**](docs/WIKI.md) | [**Notas de la Versión**](docs/RELEASE_NOTES_DEMO.md) | [**SOP Recepción SAP**](docs/SOP_RECEPCION_SAP.md) | [**Auditoría de Bultos & VBA**](docs/CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md) | [**Arquitectura de Despacho**](docs/AUDITORIA_DESPACHO_MACRO.md)

---

## 🌟 Edición Portable de Escritorio (.NET 10 LTS)

> [!TIP]
> **Cero Instalación:** ACR se distribuye como un **único archivo ejecutable nativo de Windows** (`dist/ACR_Release/Acr.Desktop.exe`, Win-x64). Funciona de inmediato **sin requerir instalación de .NET Runtime ni Python, y sin permisos de Administrador**.

Un motor de cálculo y asistente de auditoría interactivo de nivel empresarial para la **Conciliación 3-Way de Inventarios y Cierres de Importación**. Diseñado para eliminar el cruce manual entre Órdenes de Compra (PO), Facturas Comerciales/DUA, Conteos Físicos de Calidad en Muelle y Movimientos Contables ERP (`Mov 101` de Recepción y `Mov 641` de Despacho/Traslado).

```mermaid
flowchart LR
    A["Compras ERP (PO 43)"] --> D["Asistente ACR (.NET 10)"]
    B["Factura Comercial (Ped 41)"] --> D
    C["Descarga Física y Empaque QC"] --> D
    
    D --> E["📦 Franja Ejecutiva de KPIs & Grilla"]
    D --> F["📋 Proyección Oficial Hoja Bultos (1 a 1)"]
    D --> G["⚡ Tiras de Pegado Rápido (AS:AT / AW:AY)"]
    D --> H["📖 Enciclopedia SAP & Diagnóstico Operativo"]
    D --> I["💾 Generación de Cierre .XLSM con Macros"]
    
    G -->|"Pegado en 10 seg"| J["Libro Oficial (.xlsm)"]
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
ACR impone una estricta segregación matemática entre los tres niveles de inventario:
* **Nivel 1 — Bultos Físicos (Cartones/Cajas):** Cajas físicas descargadas en muelle (Col `AS`), bultos del Packing List del proveedor (Col `AW`), y conteo de Control de Calidad (`Bultos Enteros` + `Complementos`).
* **Nivel 2 — Prepacks / Curvas Comerciales (PRK / Curva):** Unidad de compra en pedidos mayoristas (Cols `BD` y `BE` en PRKs, Col `BB` proporción de curva, Col `BC` unidad de medida).
* **Nivel 3 — Unidades Totales (Prendas/Piezas):** Unidades facturadas, nacionalizadas y recibidas prenda por prenda (Cols `F`, `G`, `H`, `AT`, `AX`).

### Capacidades Operativas Clave
- **Arquitectura Nativa en .NET 10 LTS Win-x64:** Algoritmos deterministas y procesamiento OpenXML de alto rendimiento con 181 pruebas automatizadas.
- **Franja Ejecutiva de KPIs:** Visualización inmediata de cabecera con contenedor, PO, estilos auditados, unidades y bultos (facturados vs físicos) y revueltos.
- **Pestaña Oficial de Bultos y Empaque:** Réplica exacta de 10 columnas enlazada a `CIERRE RECEP`, con cálculo de cajas maestras (enteros), picos/complementos y unidades por bulto.
- **Tiras de Pegado Rápido (`ZDESCARGA` & `PACKING LIST`):** Bloques de texto TSV listos para copiar con 1 clic y pegar con `Ctrl+V` en Excel (`AS:AT` y `AW:AY`), cuadrando la macro oficial en **10 segundos**.
- **Enciclopedia de Transacciones SAP y Diagnóstico:** Catálogo integrado de transacciones (`ME80FN`, `SE16N`, `MB51`, `VL31N`, `RFUI`) y guías de triage vinculadas directamente desde la grilla de discrepancias.
- **Exportación en 1 Clic del Libro Oficial (.xlsm):** Genera directamente el archivo Excel oficial cuadrado, preservando intactas todas las macros VBA (`vbaProject.bin`), estilos y fórmulas.

---

## 🚀 Inicio Rápido

### Opción A: Ejecutable Portable Autónomo (Recomendado)
1. Ejecuta `dist/ACR_Release/Acr.Desktop.exe`.
2. Selecciona tu archivo de Control de Calidad (`CONTROL CALIDAD.xlsm`) y la Plantilla Maestra (`CIERRE_TEMPLATE.xlsm`).
3. Haz clic en **"Ejecutar Conciliación y Generar Cierre .XLSM"**.

### Opción B: Compilación desde Código Fuente (.NET 10 LTS)
```bash
# 1. Restaurar y compilar
dotnet build src/dotnet/Acr.sln -c Release

# 2. Ejecutar suite de pruebas (181 tests)
dotnet test src/dotnet/Acr.sln

# 3. Publicar ejecutable único autosuficiente
dotnet publish src/dotnet/Acr.Desktop/Acr.Desktop.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o dist/ACR_Release
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
