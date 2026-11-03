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

### Paso 5: El Cuadre de Muelle en Excel (Solución a los `-51,926`)
En la plantilla oficial de Presidencia (`CIERRE RECEP`), la macro inicial deja vacías las columnas de muelle y packing list, provocando que las fórmulas de diferencia calculen `0 - 888 = -888`, acumulando miles de prendas negativas.

**Para solucionarlo en 10 segundos:**
1. Ve a la pestaña **`📋 Copiar a Excel Oficial`** y selecciona la sub-pestaña **`📦 Bloque 2: ZDESCARGA y PACKING LIST`**.
2. **Tira 1 (ZDESCARGA - Cols AS y AT):**
   - Haz clic en el recuadro de la Tira 1, presiona `Ctrl + A` y luego `Ctrl + C`.
   - Ve a tu Excel oficial, haz clic en la celda **`AS13`** y presiona `Ctrl + V`.
   - *La columna AU calcula automáticamente diferencia cero.*
3. **Tira 2 (PACKING LIST - Cols AW, AX y AY):**
   - Haz clic en el recuadro de la Tira 2, presiona `Ctrl + A` y luego `Ctrl + C`.
   - Ve a tu Excel oficial, haz clic en la celda **`AW13`** y presiona `Ctrl + V`.
   - *La columna AZ calcula automáticamente diferencia cero.*
4. **Resultado:** La fila totalizadora de Excel pasa inmediatamente de `-51,926` a **`0 DIFERENCIAS (100% Cuadrado)`**.

---

### Paso 6: Checklist de Impresión Física (Expediente Oficial)
Para armar la carpeta física en el orden reglamentario exigido por auditoría:
1. [x] **Portada de Proveedor** (Hoja `PORTADA`).
2. [x] **Datos Generales** (Hoja `DATOS GENERALES` con totales de PO 43 y Factura 41).
3. [x] **Carátula del Cierre Oficial** (Hoja `CIERRE RECEP`).
4. [x] **Formulario de Nota de Crédito** (Hoja `PARA N.C RECEP.`).
5. [x] **Respaldo de Ingreso ERP** (Impresión de `MB51` Mov 101/102).
6. [x] **Lista de Empaque de Bultos** (Hoja `Bultos`).
7. [x] **DUA** (Documentación Aduanera sellada con peso y valor CIF).

---

## ❓ 3. Preguntas Frecuentes & Casos Borde (FAQ)

### ¿Qué pasa si el contenedor viene dividido (Parte A vs Parte B)?
En importaciones con múltiples contenedores o entregas diferidas, la factura comercial suele declarar el 100% de los materiales, pero el muelle solo descarga la primera parte.
- ReconcileAI clasifica automáticamente los ítems descargados (`Parte A`) y marca los restantes como `PENDIENTE_PARTE_B`.
- De esta forma, evitas levantar reclamos falsos de notas de crédito por prendas que simplemente vienen en el segundo viaje.

### ¿Qué pasa si los proveedores usan formatos diferentes de Packing List?
Los proveedores de China, EE. UU. o Panamá suelen enviar formatos de lista de empaque completamente diferentes.  
**Solución arquitectónica:** ReconcileAI no intenta descifrar plantillas extranjeras impredecibles; se ancla en los registros oficiales estandarizados del ERP (`ME80FN`) y en el reporte interno del equipo de muelle (`Control de Calidad`), garantizando compatibilidad universal sin importar el origen de la mercancía.

---
---

# 📖 English Operational Wiki & Tutorial

## 1. Operating Concept
ReconcileAI is a safety copilot designed to eliminate human calculation fatigue during 3-Way inventory reconciliation without disrupting established corporate spreadsheet workflows.

## 2. Fast Workflow Steps
1. **SAP GUI Data Export:** Copy ALV grids from `ME80FN` (Purchasing/Invoice) and `MB51` (Goods movements 101/102).
2. **Pre-Macro Clipboard Validation:** Paste raw SAP data into the Pre-Macro validator to neutralize scientific notation (`1.45E+06`) and verify column delimiters.
3. **File Selection:** Open the web app, browse local directories or drag-and-drop the Quality Control and Closure spreadsheets.
4. **Audit Dashboard:** Inspect discrepancies via the color-coded semáforo table or the 360° Material Inspector card.
5. **Macro Squaring Strips (`AS:AZ`):** Copy the pre-formatted TSV strips for Dock Receiving (`AS:AT`) and Packing List (`AW:AY`) and paste directly into cells `AS13` and `AW13` of the Excel closure sheet to resolve artificial negative variances in seconds.
6. **Physical Binder Assembly:** Follow the 7-step checklist to prepare the official physical closure binder.
