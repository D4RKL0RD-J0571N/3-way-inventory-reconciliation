# 📖 Wiki Operativa & Guía de Inicio (Tutorial Paso a Paso)
### ReconcileAI — Copiloto Universal de Conciliación 3-Way & Recepción

[**English Version Below / Versión en Inglés al Final**](#-english-operational-wiki--tutorial) | [**Notas de la Versión (Demo v1.0.0)**](RELEASE_NOTES_DEMO.md) | [**SOP Recepción SAP**](SOP_RECEPCION_SAP.md)

---

## 🌟 1. Versión Demo Portable v1.0.0 (Distribución Independiente)

Para facilitar la evaluación de la herramienta en cualquier computadora de oficina o bodega:
1. **Archivo de Distribución:** `dist/ReconcileAI_Portable.zip` (134 MB).
2. **Cero Instalación:** No requiere instalar Python ni componentes adicionales.
3. **Cero Permisos Especiales:** Funciona como usuario estándar en Windows 10 y Windows 11 sin pedir credenciales de administrador.
4. **Instrucciones de Uso Rápido:**
   - Descomprime el archivo `.zip` en tu equipo o memoria extraíble.
   - Haz doble clic en **`ReconcileAI.exe`** (o en **`INICIAR_RECONCILEAI.bat`**).
   - Se abrirá automáticamente tu navegador web en `http://localhost:8501`.
   - Coloca tus carpetas de importaciones (ej. `CONT-001` o `CONT-002`) dentro de la carpeta `data/` o búscalas con el explorador integrado.
5. **Soporte Bilingüe & Tooltips Traducidos:**
   - Por defecto, la aplicación inicia en **Español**.
   - Los botones y tooltips de las tablas de datos (`Pantalla completa`, `Descargar como CSV`, `Buscar`) se traducen en tiempo real.
   - Puedes cambiar a **Inglés** en cualquier momento desde el selector en la barra lateral (`🌐 Idioma / Language`).

---

## 📌 2. Filosofía Operativa: El Copiloto de Precisión
Esta herramienta fue diseñada con un principio fundamental: **no busca eliminar el uso de Excel ni automatizar tareas que resten visibilidad al operador**. Su objetivo es actuar como un **asistente de pre-validación y cálculo matemático instantáneo**, eliminando el estrés del cruce manual, previniendo errores de notación científica y asegurando que los números presentados a Gerencia y Auditoría estén 100% cuadrados.

```mermaid
flowchart LR
    A["Extracción SAP GUI"] --> B["Validador Pre-Macro (App)"]
    B --> C["Copia a Excel Oficial (.xlsm)"]
    C --> D["Cruce con Muelle (Calidad)"]
    D --> E["Tiras Rápidas AS:AZ (App)"]
    E --> F["Cuadre a 0 y Cierre Listo"]
```

---

## 💼 3. Guía para Presentar una Demostración (Demo de 5 Minutos a Jefatura)

Si deseas presentar ReconcileAI a tu jefatura, supervisores o compañeros de equipo, sigue este guión práctico de 5 minutos:

| Minuto | Acción en Pantalla | Qué Explicar al Evaluador / Jefatura |
| :---: | :--- | :--- |
| **01** | Mostrar el Excel oficial con diferencias artificiales en rojo. | *"Nuestra macro oficial deja vacías las columnas de muelle y packing list, obligando a digitar cientos de números a mano celda por celda y arriesgando errores humanos."* |
| **02** | Abrir ReconcileAI y seleccionar la carpeta del contenedor. | *"Con ReconcileAI, simplemente seleccionamos la carpeta de la importación o arrastramos los archivos. El sistema lee en 1 segundo ambos libros sin modificar el disco."* |
| **03** | Mostrar la **Barra de 6 KPIs** y la **Tabla Semáforo**. | *"El sistema identifica inmediatamente los bultos teóricos vs físicos, clasifica los materiales cuadrados en verde y detecta si hay entregas diferidas (Parte B) o faltantes reales."* |
| **04** | Abrir la **Ficha 360° del Material** y buscar un estilo. | *"En una sola tarjeta vemos compras, factura, físico, costos y la explosión matemática de tallas hijas sin tener que armar fórmulas de BUSCARV."* |
| **05** | Pestaña **`Copiar a Excel Oficial`** y botón de descarga. | *"Para cerrar el expediente: o descargamos el libro .xlsm ya cuadrado con 1 clic en memoria, o copiamos las dos tiras preformateadas en 10 segundos. Las diferencias bajan a cero de inmediato."* |

---

## 🚀 4. Tutorial Paso a Paso: Flujo de Trabajo Completo

### Paso 1: Extracción Segura desde SAP GUI
1. **Pedidos de Compra (PO 43):**
   - Transacción: `ME80FN` (o reporte de compras).
   - Filtro: Número de pedido o proveedor.
   - *Tip Pro:* En lugar de exportar a HTML, usa la opción de exportar a texto o copiar con `Ctrl + C` directo de la grilla ALV.
2. **Factura de Nacionalización (Pedido 41):**
   - Transacción: `ME80FN` filtrando por el documento `41...`.
3. **Ingresos Físicos a Almacén:**
   - Transacción: `MB51`.
   - Filtros: Centro Logístico / Almacén de Recepción, Clases de movimiento `101` y `102` (para netear devoluciones).

---

### Paso 2: Validador de Portapapeles (Pre-Macro)
Antes de pegar cualquier tabla de SAP en la plantilla de Excel:
1. Abre **ReconcileAI** y dirígete a la pestaña **`⚡ Validador SAP (Pre-Macro)`**.
2. Pega la tabla copiada con `Ctrl + V` en el recuadro.
3. Haz clic en **`⚡ Auditar y Limpiar Portapapeles`**.
4. **¿Qué hace el sistema?**
   - Detecta si las columnas vienen separadas por tabulación, comas o barras.
   - Desactiva cualquier notación científica que Excel suele interpretar como fecha o fórmula (ej. `1.0010E+06` vuelve a `1001000`).
   - Elimina saltos de línea huérfanos que puedan desfasar las filas en la macro.

---

### Paso 3: Carga y Explorador de Archivos (Sin Suposiciones)
1. Al abrir la app, verás la pantalla limpia de carga con dos modalidades:
   - **📁 Explorador de Disco (Carpetas Locales):** Escribe la ruta de la carpeta (ej. `data/CONT-001` o `D:/Descargas`) y verás dos desplegables con todos los archivos Excel encontrados. Eliges cuál es Calidad y cuál es Cierre.
   - **📤 Subir Archivos (Arrastrar y Soltar):** Puedes hacer clic en *"Browse files"* para que se abra el Explorador nativo de Windows y selecciones los archivos directamente.
2. Haz clic en **`🚀 Iniciar Auditoría`**.

---

### Paso 4: Auditoría de Bultos y Conciliación Visual
1. **Barra Superior de Métricas:**
   - Verifica que el total de unidades facturadas coincida con las recibidas.
   - Si el contenedor viene dividido (como cuando tiene Parte A y Parte B), el sistema te indica cuántos materiales corresponden a cada lote.
2. **📦 Pestaña Cuadre & Desglose de Bultos:**
   - **Tarjetas de Balance:** Bultos Teóricos vs Bultos de Muelle vs Packing List de Proveedor.
   - **Calculadora de Empaque:** Introduce cualquier desglose (ej. `39*60 1*78`) para calcular al instante total de bultos y prendas sin errores manuales.
   - **Réplica Oficial de Hoja 'Bultos':** Visualiza la matriz exacta de bultos por material y contenedor para validar la documentación de transporte.
3. **🚦 Tabla Semáforo:**
   - 🟢 **Verde suave:** Ítems completamente cuadrados (cero diferencias).
   - 🔵 **Azul suave:** Ítems de la Parte B (en tránsito o descargados en segundo lote).
   - 🔴 **Rojo:** Materiales con faltante o avería física.
4. **🔍 Ficha 360° del Material:**
   - Escribe cualquier código SAP o estilo de proveedor. En una sola tarjeta visual verás compras, factura, físico, bultos y el desglose de tallas hijas con su curva individual.

---

### Paso 5: El Cuadre de Muelle en Excel (Solución a las Diferencias Negativas y a Cierres sin SAP)
En la plantilla oficial de Cierre Administrativo (`CIERRE RECEP`), la macro inicial deja vacías las columnas de muelle y packing list, provocando que las fórmulas de diferencia calculen `0 - 888 = -888`, acumulando miles de prendas negativas.

**¿Cómo terminar de llenar la plantilla sin volver a usar SAP?**
Una vez que pegaste los datos iniciales de SAP en la plantilla, **no necesitas abrir SAP nunca más**:
- **Opción 1 (1 Clic — Automática):** En la pestaña **`📋 Copiar a Excel Oficial`**, haz clic en el botón azul **`⚡ Descargar Libro Oficial Cuadrado (.xlsm con Macros VBA)`**. La app genera el libro corporativo directamente en memoria, escribe las columnas `AS:AT` y `AW:AY`, repara cualquier fórmula dañada y preserva intactas todas las macros VBA.
- **Opción 2 (Copiado Rápido en 1 Clic — 5 segundos):**
  1. Ve a la sub-pestaña **`📦 Bloque 2: ZDESCARGA y PACKING LIST`**.
  2. Haz clic en el botón **`📋 Copiar Tira 1 (AS13)`**. El bloque completo se copia automáticamente al portapapeles con confirmación visual verde. Pégalo directamente en la celda **`AS13`** de tu Excel. *(La columna AU de diferencia calcula 0)*.
  3. Haz clic en el botón **`📋 Copiar Tira 2 (AW13)`**. Pégalo en la celda **`AW13`** de tu Excel. *(La columna AZ de diferencia calcula 0)*.
  *(También tienes los recuadros de texto debajo si prefieres la copia manual con `Ctrl + A` y `Ctrl + C`).*
- **Resultado:** La fila totalizadora de Excel pasa inmediatamente a **`0 DIFERENCIAS (100% Cuadrado)`**.

#### 📦 Regla Matemática del "Bulto Cero":
Para entregas diferidas (`Parte B`, `Parte C`...) o materiales con `0` unidades recibidas:
- **Bultos Descarga (`AS`):** `0.0`
- **Unidades Descarga (`AT`):** `0.0`
- **Bultos Packing List (`AW`):** `0.0`
- **Unidades Packing List (`AX`):** `0.0`
- **Empaque (`AY`):** `0`
Esto elimina cualquier bulto fantasma (`1.0`) y garantiza un **100% de coincidencia exacta** con el llenado oficial humano de las hojas `CIERRE RECEP` y `Bultos`.

#### ⏸️ Modo Pausa para Montaje a Ciegas (Sin Control de Calidad):
Si una importación aún no tiene el informe de muelle de Control de Calidad, **no se puede emitir el cierre oficial**. ReconcileAI permite realizar el montaje preliminar desde SAP, dejando el estado en **`PAUSA_SIN_CONTROL_CALIDAD`** con advertencia visual hasta que se incorpore el conteo físico de descarga.

---

### Paso 6: Checklist de Impresión Física (Expediente Oficial)
Para armar la carpeta física en el orden reglamentario exigido por auditoría:
1. [x] **Portada de Proveedor** (Hoja `PORTADA`).
2. [x] **Datos Generales** (Hoja `DATOS GENERALES` con totales de PO 43, Factura 41 y bultos).
3. [x] **Carátula del Cierre Oficial** (Hoja `CIERRE RECEP` con diferencias en 0).
4. [x] **Formulario de Nota de Crédito** (Hoja `PARA N.C RECEP.` — $0.00 si está limpio).
5. [x] **Respaldo de Ingreso ERP** (Impresión de `MB51` Mov 101/102).
6. [x] **Lista de Empaque de Bultos** (Hoja `Bultos`).
7. [x] **DUA** (Documentación Aduanera sellada con peso y valor CIF).

---

## 🔎 5. Caso de Estudio: Trampas Operativas Frecuentes en Muelle
En los cierres logísticos se presentan habitualmente trampas típicas del trabajo manual en bodega:
1. **La Trampa de las Columnas Desplazadas:** En el archivo de Control de Calidad, si un operador inserta una columna `"Tarima"` o `"Ubicación"` en `CONTROL DATA`, esto desplaza todas las columnas una posición a la derecha. Cualquier fórmula manual de `BUSCARV` en Excel apunta a la columna equivocada, causando que el 100% de los códigos den `#N/A`.
2. **ID de Material Duplicado en `BASE DE DATOS`:** En la hoja de curvas, si una columna duplica un ID de catálogo o lista en lugar del código de material principal.
3. **Fórmulas Borradas en Columna H:** En `CIERRE RECEP`, si alguien borra las fórmulas oficiales de sumatoria y tipea cantidades parciales o en bultos (como ingresar `81` en vez de `1,728`).
4. **La Trampa del Alisto Multi-Almacén (Almacén Central vs Resurtido):** Cuando un cargamento se divide entre Almacén Central y Almacén de Resurtido, si el operador sólo copia la columna de Almacén Central, quedan cientos de unidades aparentemente "desaparecidas" cuando en realidad están en reserva.

**Cómo lo resuelve el motor:**
- **Mapeo Heurístico:** Detecta los encabezados dinámicamente sin importar si se insertan columnas de tarimas o ubicaciones.
- **Detección Universal:** Lee las listas y matrices de curvas para asociar los Prepacks con sus curvas reales (ej. 1-2-2-1).
- **Consolidación Total:** Suma todos los destinos (Central + Resurtido + Especial) para cuadrar exactamente el 100% de las prendas con 0 faltantes artificiales.

---

## ❓ 6. Preguntas Frecuentes & Casos Borde (FAQ)

### ¿Qué pasa si el contenedor viene dividido (Parte A vs Parte B)?
En importaciones con múltiples contenedores o entregas diferidas, la factura comercial suele declarar el 100% de los materiales, pero el muelle solo descarga la primera parte.
- ReconcileAI clasifica automáticamente los ítems descargados (`Parte A`) y marca los restantes como `PENDIENTE_PARTE_B`.
- De esta forma, evitas levantar reclamos falsos de notas de crédito por prendas que simplemente vienen en el segundo viaje.

### ¿Cómo se concilian los Bultos y Cartones para Auditoría y Gerencia?
Auditoría y Gerencia revisan las filas 5 y 6 de `CIERRE RECEP` antes de mirar las unidades.
- ReconcileAI compara los **Bultos Teóricos del Packing List** contra los **Bultos Recibidos en Muelle**.
- Si un contenedor de 750 bultos se consolidó en 700 bultos físicos (diferencia justificada por capacidad de empaque), el sistema lo refleja con claridad para justificar la diferencia de volumen sin inconsistencias de prendas.

### ¿Cómo se conectan los Cierres de Recepción con los Cierres de Despacho?
El libro oficial de cierre es un sistema dual de dos fases (ver [**Arquitectura de Despacho**](AUDITORIA_DESPACHO_MACRO.md)):
1. **Fase 1 (Recepción - `CIERRE RECEP`):** Se concilia aduana e ingreso a bodega.
2. **Fase 2 (Despacho - `CIERRE DESP`):** Se audita la distribución a tiendas vía movimiento ERP `641`.
- **Ventaja de ReconcileAI:** Al completar el cierre de recepción, el 80% de la hoja de despacho queda pre-cableada y sincronizada automáticamente mediante fórmulas nativas de Excel vinculadas a `CIERRE RECEP`.

---
---

# 📖 English Operational Wiki & Tutorial

## 1. Standalone Portable Edition (v1.1.0)
- **Zero Python Installation:** Runs standalone from `ReconcileAI_Portable.zip` on any Windows 10/11 system.
- **USB & Offline Friendly:** Unpack anywhere, double-click `ReconcileAI.exe`, and audit containers immediately.
- **Bilingual & Localized Tables:** Starts in Spanish by default; toggle to English anytime in the sidebar. Table toolbars and hover tooltips (`Fullscreen`, `Download as CSV`, `Search`) dynamically adapt to the chosen language.
- **Dedicated Carton & Packaging Tab (`📦 Cuadre & Desglose de Bultos`):** Top 5 carton KPI summary, interactive packaging calculator, and exact replica of the official `'Bultos'` sheet.
- **Dual-Phase Readiness (Reception & Dispatch):** Pre-wires and reconciles both import goods receipts (`CIERRE RECEP`, `Mov 101`) and retail store transfers (`CIERRE DESP`, `Mov 641`).
- **Enterprise Privacy:** Zero sensitive corporate data. Operates 100% in local workstation memory.

## 2. 5-Minute Evaluation Walkthrough for Leadership
1. **Show the Legacy Pain:** Point out the artificial negative difference in standard Excel closure sheets.
2. **Select Container:** Load Quality Control and Closure files with drag-and-drop or local folder browser.
3. **Inspect KPIs:** Immediate 6-metric summary of Cartons/Bultos, Invoiced Units, and Physical Dock Counts.
4. **Audit Cartons & Packaging:** Review the **`📦 Cuadre & Desglose de Bultos`** tab to verify physical cartons vs packing list cartons.
5. **Drill Down 360°:** Instant size-curve breakdown without manual VLOOKUP formulas.
6. **Square Closure:** Click 1-Click `.xlsm` Export or copy pre-formatted TSV strips (`Cols AS:AT` & `Cols AW:AY`) to square differences to `0`.

## 3. Case Study: Solving Difficult Shipments
- **Column Shift Immunity:** Dynamically maps headers even when operators insert custom columns (e.g. Pallet/Tarima IDs).
- **Multi-Warehouse Allocation:** Accurately consolidates multi-destination shipments (Central Warehouse + Replenishment Warehouse).
- **Compound Packaging Strings:** Automatically decomposes odd packaging formats (e.g. `'39*60 1*78'`).
- **Self-Healing Formulas:** Automatically restores deleted or overwritten formulas in official macro templates.
