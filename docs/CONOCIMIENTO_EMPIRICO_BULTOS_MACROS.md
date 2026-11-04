# 📘 Base de Conocimiento Empírico: Auditoría de Bultos, Empaque y Macros VBA

Este documento recopila la verdad operativa y matemática extraída directamente del análisis forense de 71 importaciones reales (`IMP-98` a `IMP-244`), la descompresión binaria de macros VBA (`vbaProject.bin`), las fórmulas nativas de hojas de cierre (`CIERRE RECEP`, `Bultos`, `DATOS GENERALES`) y las hojas de muelle (`Control Calidad`).

---

## 1. La Arquitectura de Tres Niveles Numéricos

El análisis forense de las importaciones reales reveló que en la cadena de suministro existen **3 niveles numéricos independientes** que jamás deben confundirse:

```
┌────────────────────────────────────────────────────────────────────────┐
│ Nivel 1: Bultos Físicos (Cartons / Boxes)                              │
│ • Lo que descarga el montacargas en muelle.                            │
│ • Col 45 (AS): TOTAL BULTOS DESCARGA (Z-Carga)                         │
│ • Col 49 (AW): Bultos según Packing List                               │
│ • Hoja 'Bultos': Col C (CANTIDAD DE BULTOS Z-Carga)                    │
│ • Hoja 'CONTROL DATA': Col 32 (TOTAL DE BULTOS = Enteros + Complementos)│
└────────────────────────────────────────────────────────────────────────┘
                                   │
                                   ▼
┌────────────────────────────────────────────────────────────────────────┐
│ Nivel 2: Prepacks / Pedido (PRK / Curva)                               │
│ • Unidad de empaque comercial / surtido de tallas y colores.           │
│ • Col 56 (BD): CANTIDAD PEDIDO 43 (en PRK)                             │
│ • Col 57 (BE): CANTIDAD PEDIDO 41 (en PRK)                             │
│ • Col 54 (BB): Curva (unidades por prepack, ej. 6, 12)                 │
│ • Col 55 (BC): Unidad de Medida (UM): 'PRK' o 'UN'                     │
└────────────────────────────────────────────────────────────────────────┘
                                   │
                                   ▼
┌────────────────────────────────────────────────────────────────────────┐
│ Nivel 3: Unidades Totales Nacionalizadas (Pieces / Individual Units)   │
│ • Prendas o artículos individuales nacionales.                         │
│ • Col 6 (F): CANTIDAD PEDIDO 43 (= BD * BB = PRK * Curva)              │
│ • Col 7 (G): CANTIDAD PEDIDO 41 (= BE * BB = PRK * Curva)              │
│ • Col 8 (H): RECIBIDO REAL & FISICO (= G - J - N - O + L + P + Q)       │
│ • Col 46 (AT): CANTIDAD DE UNIDADES EN BULTOS (Muelle)                 │
│ • Col 50 (AX): CANTIDAD UNIDADES (Packing List)                        │
└────────────────────────────────────────────────────────────────────────┘
```

> [!CAUTION]
> **El Error Histórico de la Versión Anterior:**
> En importaciones como `IMP-98`, un material con 420 Prepacks (`BE=420`) y curva 6 contiene 2,520 unidades (`G=2520`). Los bultos físicos reales son **42 cajas** (cada caja contiene 10 prepacks = 60 unidades por caja). La versión anterior asignaba `bultos = 420`, multiplicando por 10 los bultos reales y desfasando el inventario de muelle.

---

## 2. Los 4 Patrones Reales de Empaque

Del muestreo empírico de las columnas `AS` (Bultos Muelle), `AW` (Bultos PL) y `AY` (Empaque) se catalogaron 4 patrones exactos:

### Patrón A: Caja Master con Multi-Prepacks (Ej. `IMP-98`)
* **Col 45 (`AS`):** `42`
* **Col 51 (`AY` - Empaque):** `60`
* **Col 54 (`BB` - Curva):** `6`
* **Col 57 (`BE` - PRK):** `420`
* **Col 7 (`G` - Unidades):** `2,520`
* **Interpretación:** 42 cajas * 60 unidades/caja = 2,520 unidades. Cada caja contiene 10 prepacks de curva 6.

### Patrón B: Cajas Maestras con Bulto Complemento / Pico (Ej. `IMP-98` Mat 1459288)
* **Col 45 (`AS`):** `'39*60 1*78'`
* **Col 46 (`AT`):** `2,418`
* **Col 51 (`AY`):** `60`
* **Col 54 (`BB`):** `6`
* **Col 57 (`BE`):** `403`
* **Interpretación:** 
  * 39 cajas completas de 60 unidades = 2,340 unidades.
  * 1 caja complemento/pico de 78 unidades = 78 unidades.
  * **Total Bultos:** 39 + 1 = **40 bultos**.
  * **Total Unidades:** 2,340 + 78 = **2,418 unidades**.
  * Coincide exactamente con `403 PRK * 6 = 2,418 unidades`.

### Patrón C: Prepack Directo (1 Bulto = 1 Prepack, Ej. `IMP-196` Calzado / Prendas Voluminosas)
* **Col 45 (`AS`):** `50`
* **Col 51 (`AY`):** `'50*12'` o `12`
* **Col 54 (`BB`):** `12`
* **Col 57 (`BE`):** `50`
* **Col 7 (`G`):** `600`
* **Interpretación:** Cada bulto físico contiene exactamente 1 prepack de 12 pares/unidades.

### Patrón D: Unitario / No Textil (Ej. `IMP-100` Juguetes / Bicicletas, `IMP-224`)
* **Col 45 (`AS`):** `65`
* **Col 51 (`AY`):** `1`
* **Col 54 (`BB`):** `1`
* **Col 55 (`BC`):** `'UN'`
* **Col 57 (`BE`):** `65`
* **Col 7 (`G`):** `65`
* **Interpretación:** Mercancía unitaria no asortada. Bultos = Prepacks = Unidades.

---

## 3. Gramática del Analizador de Cadenas de Empaque (`parse_packaging_string`)

Para procesar tanto valores escalares como expresiones compuestas sin pérdida de precisión, el analizador implementa la siguiente gramática formal:

$$\text{Expresión} \rightarrow \sum_{i} (B_i \times U_i)$$

Donde:
* $B_i$: Cantidad de bultos del lote $i$.
* $U_i$: Unidades por bulto del lote $i$.
* Si la cadena no contiene multiplicador `*` o `x`, $B = \text{float}(\text{cadena})$ y $U = 0$ (o heredado de la curva/empaque).

### Ejemplos Validados:
| Entrada (`raw`) | Bultos Totales ($B$) | Unidades Calculadas ($U$) | Bultos Enteros | Unid/Bulto Entero | Bultos Comp. | Unid/Bulto Comp. |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| `'42'` | 42.0 | 0.0 (según curva) | 42.0 | - | 0.0 | - |
| `'39*60 1*78'` | 40.0 | 2,418.0 | 39.0 | 60.0 | 1.0 | 78.0 |
| `'40*60 1*72'` | 41.0 | 2,472.0 | 40.0 | 60.0 | 1.0 | 72.0 |
| `'50*12'` | 50.0 | 600.0 | 50.0 | 12.0 | 0.0 | 0.0 |
| `'11*72 1*96'` | 12.0 | 888.0 | 11.0 | 72.0 | 1.0 | 96.0 |

---

## 4. Hallazgos Forenses en Macros VBA (`vbaProject.bin`)

Al descomprimir el archivo binario `xl/vbaProject.bin` en `IMP-182` e `IMP-100`, se auditaron las siguientes macros:

* **`Format43` / `Mueve43`:** Formatea la extracción SAP ME80FN y copia columnas básicas de compras.
* **`Mueve41`:** Copia el pedido nacionalizado a columnas A hasta AC.
* **`MueveCurva`:** Extrae las curvas de prepacks a las columnas BB, BC, BD, BE.
* **Hallazgo Crítico:** Las macros **nunca escriben en las columnas `AS:AT` (ZDESCARGA: Bultos y Unidades de muelle) ni `AW:AY` (PACKING LIST)**.
  * Los operadores debían transcribir estas tiras manualmente desde el reporte de muelle o packing list del proveedor.
  * Al haber expresiones compuestas como `'39*60 1*78'`, la función nativa `=SUM(AS13:AS...)` fallaba o daba cero, obligando a ingresar el total a mano (`B6=1380`).
  * **ReconcileAI automatiza la generación de ambas tiras (Tira 1 y Tira 2) y el llenado de la hoja `Bultos` sin intervención manual.**

---

## 5. Especificación de la Hoja `'Bultos'`

La hoja `'Bultos'` en el libro de cierre oficial tiene una estructura estandarizada de 10 columnas enlazadas 1 a 1 con `'CIERRE RECEP'`:

| Columna | Encabezado Oficial | Fórmula en Fila 3 | Origen en `'CIERRE RECEP'` | Descripción |
| :---: | :--- | :--- | :--- | :--- |
| **B** (2) | `CATEGORIA` | `='CIERRE RECEP'!C13` | Col `C` (3) | Categoría del producto (ej. DAMAS, CABALLEROS) |
| **C** (3) | `CANTIDAD DE BULTOS Z-Carga` | `='CIERRE RECEP'!AS13` | Col `AS` (45) | Bultos físicos descargados en muelle |
| **D** (4) | `CANTIDAD DE UNIDADES Z-Carga` | `='CIERRE RECEP'!AT13` | Col `AT` (46) | Unidades físicas descargadas en muelle |
| **E** (5) | `MATERIAL` | `='CIERRE RECEP'!E13` | Col `E` (5) | Código de material SAP (Padre) |
| **F** (6) | `RECIBIDO REAL & FISICO` | `='CIERRE RECEP'!H13` | Col `H` (8) | Unidades finales recibidas |
| **G** (7) | `Curva` | `='CIERRE RECEP'!BB13` | Col `BB` (54) | Ratio de unidades por prepack |
| **H** (8) | `UM` | `='CIERRE RECEP'!BC13` | Col `BC` (55) | Unidad de medida (`PRK` o `UN`) |
| **I** (9) | `DESCRIPCION (TEXTO BREVE)` | `='CIERRE RECEP'!AD13` | Col `AD` (30) | Descripción completa del material |
| **J** (10) | `EMPAQUE UNIDADES Z-Carga` | `='CIERRE RECEP'!AY13` | Col `AY` (51) | Unidades por bulto según empaque |

---

## 6. Fórmulas de Encabezado en `'DATOS GENERALES'` y `'CIERRE RECEP'`

* **Bultos Teóricos (Packing List):**
  $$B_5 = \text{Col } AW \text{ (Fila Total, ej. } =AW119)$$
* **Bultos Recibidos (Muelle Descarga):**
  $$B_6 = \text{Col } AS \text{ (Fila Total, ej. } =AS119)$$
* **Diferencia de Bultos:**
  $$B_7 = +B_6 - B_5 \quad (\text{Recibidos} - \text{Teóricos})$$
* **Unidades por Bulto Promedio:**
  $$D_5 = \text{IFERROR}(H_{\text{total}} / B_5, \text{""})$$
