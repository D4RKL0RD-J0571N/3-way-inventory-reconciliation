# 📖 Wiki Operativa & Guía de Inicio (Tutorial Paso a Paso)
### ReconcileAI — Copiloto Universal de Conciliación 3-Way & Recepción

[**English Version Below / Versión en Inglés al Final**](#-english-operational-wiki--tutorial)

---

## 📌 1. Filosofía Operativa: El Copiloto de Precisión
Esta herramienta fue diseñada con un principio fundamental: **no busca eliminar el uso de Excel ni automatizar tareas que resten visibilidad al operador**. Su objetivo es actuar como un **asistente de pre-validación y cálculo matemático instantáneo**, eliminando el estrés del cruce manual, previniendo errores de notación científica y asegurando que los números presentados a Presidencia y a los dueños estén 100% cuadrados.

```mermaid
flowchart LR
    A["Extracción SAP GUI"] --> B["Validador Pre-Macro (App)"]
    B --> C["Copia a Excel Oficial (.xlsm)"]
    C --> D["Cruce con Muelle (Calidad)"]
    D --> E["Tiras Rápidas AS:AZ (App)"]
    E --> F["Cuadre a 0 y Cierre Listo"]
```

---

## 🚀 2. Tutorial Paso a Paso: Flujo de Trabajo Completo

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
   - Desactiva cualquier notación científica que Excel suele interpretar como fecha o fórmula (ej. `1.4551E+06` vuelve a `1455100`).
   - Elimina saltos de línea huérfanos que puedan desfasar las filas en la macro.

---

### Paso 3: Carga y Explorador de Archivos (Sin Suposiciones)
1. Al abrir la app, verás la pantalla limpia de carga con dos modalidades:
   - **📁 Explorador de Disco (Carpetas Locales):** Escribe la ruta de la carpeta (ej. `data/IMP-182` o `D:/Descargas`) y verás dos desplegables con todos los archivos Excel encontrados. Eliges cuál es Calidad y cuál es Cierre.
   - **📤 Subir Archivos (Arrastrar y Soltar):** Puedes hacer clic en *"Browse files"* para que se abra el Explorador nativo de Windows y selecciones los archivos directamente.
2. Haz clic en **`🚀 Iniciar Auditoría`**.

---

### Paso 4: Auditoría Visual y Detección de Discrepancias
1. **Barra Superior de Métricas:**
   - Verifica que el total de unidades facturadas coincida con las recibidas.
   - Si el contenedor viene dividido (como `IMP-182` que tiene Parte A y Parte B), el sistema te indica cuántos materiales corresponden a cada lote.
2. **🚦 Tabla Semáforo:**
   - 🟢 **Verde suave:** Ítems completamente cuadrados (cero diferencias).
   - 🔵 **Azul suave:** Ítems de la Parte B (en tránsito o descargados en segundo lote).
   - 🔴 **Rojo:** Materiales con faltante o avería física.
3. **🔍 Ficha 360° del Material:**
   - Escribe cualquier código SAP o estilo de proveedor. En una sola tarjeta visual verás compras, factura, físico, bultos y el desglose de tallas hijas con su curva individual.

---

### Paso 5: El Cuadre de Muelle en Excel (Solución a los `-51,926` y a Cierres sin SAP)
En la plantilla oficial de Presidencia (`CIERRE RECEP`), la macro inicial deja vacías las columnas de muelle y packing list, provocando que las fórmulas de diferencia calculen `0 - 888 = -888`, acumulando miles de prendas negativas.

**¿Cómo terminar de llenar la plantilla sin volver a usar SAP?**
Una vez que pegaste los datos iniciales de SAP en la plantilla, **no necesitas abrir SAP nunca más**:
- **Opción 1 (1 Clic — Automática):** En la pestaña **`📋 Copiar a Excel Oficial`**, haz clic en el botón azul **`⚡ Descargar Libro Oficial Cuadrado (.xlsm con Macros VBA)`**. La app genera el libro corporativo directamente en memoria, escribe las columnas `AS:AT` y `AW:AY`, repara cualquier fórmula dañada y preserva intactas todas las macros VBA.
- **Opción 2 (Manual Controlada — 10 segundos):**
  1. Ve a la pestaña **`📋 Copiar a Excel Oficial`** y selecciona la sub-pestaña **`📦 Bloque 2: ZDESCARGA y PACKING LIST`**.
  2. **Tira 1 (ZDESCARGA - Cols AS y AT):** Haz clic en el recuadro de la Tira 1, presiona `Ctrl + A` y luego `Ctrl + C`. Pégalo en la celda **`AS13`** de tu Excel. *(La columna AU de diferencia calcula 0)*.
  3. **Tira 2 (PACKING LIST - Cols AW, AX y AY):** Haz clic en el recuadro de la Tira 2, presiona `Ctrl + A` y luego `Ctrl + C`. Pégalo en la celda **`AW13`** de tu Excel. *(La columna AZ de diferencia calcula 0)*.
- **Resultado:** La fila totalizadora de Excel pasa inmediatamente a **`0 DIFERENCIAS (100% Cuadrado)`**.

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

## 🔎 3. Caso de Estudio: ¿Por qué era tan difícil cerrar IMP-174?
El contenedor **`IMP-174`** (61,607 unidades) representaba uno de los casos más frustrantes para los supervisores en bodega por una combinación de 4 trampas humanas:
1. **La Trampa de las Columnas Desplazadas:** En `Control Calidad.xlsm`, un operador insertó una columna `"Tarima"` en la columna B de `CONTROL DATA`. Esto desplazó todas las columnas una posición a la derecha. Cualquier fórmula manual de `BUSCARV` en Excel apuntó a la columna equivocada (`Material Proveedor` en vez de `Material SAP`), causando que el 100% de los códigos dieran `#N/A`.
2. **ID de Material Duplicado en `BASE DE DATOS`:** En la hoja de curvas, la columna F tenía duplicado el ID de lista (`24987`) en lugar del material SAP (`1456051`), el cual estaba en la columna I.
3. **Fórmulas Borradas en Columna H:** En `CIERRE RECEP`, alguien borró las fórmulas oficiales (`=+G - J - N - O + L + P + Q`) y tipeó cantidades parciales o en bultos (como ingresar `81` en vez de `1,728`).
4. **La Trampa del Alisto Multi-Almacén (A001 vs A004):** En materiales como el `1455686002`, 444 prendas iban para `A001` (Almacén Central) y 372 para `A004` (Resurtido), totalizando 816 UN. El operador sólo copió la columna de A001, dejando 372 unidades aparentemente "desaparecidas" cuando en realidad estaban en resurtido.

**Cómo lo resolvió el motor:**
- **Mapeo Heurístico:** Detecta los encabezados dinámicamente sin importar si se insertan columnas de tarimas o ubicaciones.
- **Detección Universal:** Lee la columna I en `BASE DE DATOS` para asociar los 44 Prepacks con sus curvas reales (ej. 1-2-2-1).
- **Consolidación Total:** Suma todos los destinos (`A001 + A004 + Especial`) para cuadrar exactamente las 61,607 unidades con 0 faltantes.

---

## ❓ 4. Preguntas Frecuentes & Casos Borde (FAQ)

### ¿Qué pasa si el contenedor viene dividido (Parte A vs Parte B)?
En importaciones con múltiples contenedores o entregas diferidas, la factura comercial suele declarar el 100% de los materiales, pero el muelle solo descarga la primera parte.
- ReconcileAI clasifica automáticamente los ítems descargados (`Parte A`) y marca los restantes como `PENDIENTE_PARTE_B`.
- De esta forma, evitas levantar reclamos falsos de notas de crédito por prendas que simplemente vienen en el segundo viaje.

### ¿Cómo se concilian los Bultos y Cartones para Presidencia?
Presidencia revisa las filas 5 y 6 de `CIERRE RECEP` antes de mirar las unidades.
- ReconcileAI compara los **Bultos Teóricos del Packing List** contra los **Bultos Recibidos en Muelle**.
- Si un contenedor de 767 bultos se consolidó en 713 bultos físicos (diferencia de -54 bultos por empaque), el sistema lo refleja con claridad para justificar la diferencia de volumen sin inconsistencias de prendas.

---
---

# 📖 English Operational Wiki & Tutorial

## 1. Operating Concept
ReconcileAI is a safety copilot designed to eliminate human calculation fatigue during 3-Way inventory reconciliation without disrupting established corporate spreadsheet workflows.

## 2. Fast Workflow Steps
1. **SAP GUI Data Export:** Copy ALV grids from `ME80FN` (Purchasing/Invoice) and `MB51` (Goods movements 101/102).
2. **Pre-Macro Clipboard Validation:** Paste raw SAP data into the Pre-Macro validator to neutralize scientific notation (`1.45E+06`) and verify column delimiters.
3. **File Selection:** Open the web app, browse local directories or drag-and-drop the Quality Control and Closure spreadsheets.
4. **Audit Dashboard:** Inspect discrepancies via the color-coded semáforo table, Carton/Package cards, or the 360° Material Inspector card.
5. **Macro Squaring Strips (`AS:AZ`) or 1-Click `.xlsm` Export:**
   - **Method A:** Click "Download Official Closure (.xlsm with VBA)" to get a healed official file with 0 variances.
   - **Method B:** Copy pre-formatted TSV strips (`AS:AT` and `AW:AY`) and paste into cells `AS13` and `AW13` of Excel.
6. **Physical Binder Assembly:** Follow the 7-step checklist to prepare the official physical closure binder.

## 3. Case Study: Solving Difficult Shipments (IMP-174)
- **Column Shift Immunity:** Automatically locates SAP codes, descriptions, and totals regardless of inserted warehouse columns (e.g. pallet IDs).
- **Multi-Warehouse Allocation:** Accurately reconciles split receipts (Distribution Center A001 vs Replenishment A004).
- **Self-Healing Formulas:** Automatically restores deleted or overwritten formulas in official macro templates.
