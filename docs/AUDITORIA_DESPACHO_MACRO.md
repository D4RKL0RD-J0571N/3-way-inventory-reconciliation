# 🚚 Arquitectura y Auditoría de Cierres de Despacho (Salidas a Tiendas / Mov-641)

Este documento recopila la investigación técnica y la especificación de diseño de la **Fase 2 de Cierre: Cierre de Despacho (`CIERRE DESP`)**, extraída del análisis forense de la plantilla maestra base de macros corporativas.

---

## 1. La Arquitectura Dual del Libro Maestro de Cierre

El libro de cálculo maestro corporativo está diseñado para soportar dos cierres operativos complementarios dentro del Centro de Distribución:

```
┌────────────────────────────────────────────────────────┐
│ FASE 1: CIERRE DE RECEPCIÓN (Muelle / Aduanas)         │
│ • Entrada de mercadería desde proveedor extranjero     │
│ • Hoja 'CIERRE RECEP'                                  │
│ • Hoja 'Bultos'                                        │
│ • Hoja 'PARA N.C RECEP.'                               │
│ • Movimiento SAP 101 (Entrada de mercancías)           │
└────────────────────────────────────────────────────────┘
                           │
                           ▼ (Sincronización Automática 80%)
┌────────────────────────────────────────────────────────┐
│ FASE 2: CIERRE DE DESPACHO (Alisto / Tiendas)          │
│ • Salida de mercadería hacia red de sucursales         │
│ • Hoja 'CIERRE DESP'                                   │
│ • Hoja 'PARA N.C DESP'                                 │
│ • Hoja 'STOCK CONTROL DE INVENTARIO'                   │
│ • Movimiento SAP 641 (Traslado hacia tiendas)          │
└────────────────────────────────────────────────────────┘
```

---

## 2. Mecanismo de Sincronización en las Macros VBA

En la ejecución de las macros nativas de Excel:
1. **`Mueve43`:** Cuando la macro limpia las filas sobrantes en `CIERRE RECEP`, ejecuta en paralelo el borrado de filas no utilizadas en `CIERRE DESP`:
   ```vba
   wsTotal.Rows(primeraFilaBorrar & ":1011").Delete
   ThisWorkbook.Sheets("CIERRE DESP").Rows(primeraFilaBorrar & ":1011").Delete
   ```
2. **`Mueve51`:** Formatea la hoja `ERASER` en dos bloques paralelos para el reporte SAP `MB51`:
   * **Bloque 1 (`A1:B2`):** `MB51 101-102` (Recepción física / Ingreso) $\rightarrow$ Columna `BF` (58) en `CIERRE RECEP`.
   * **Bloque 2 (`D1:E2`):** `MB51 641-642` (Salidas a tiendas / Despacho) $\rightarrow$ Columna `U` (21) en `CIERRE RECEP`.

Al completarse el Cierre de Recepción, **la hoja `CIERRE DESP` queda automáticamente poblada en un 80% mediante fórmulas directas** hacia `CIERRE RECEP`.

---

## 3. Estructura de Columnas en `'CIERRE DESP'`

La hoja `'CIERRE DESP'` contiene 72 columnas divididas en dos zonas funcionales:

### Zona A: Datos Heredados de Recepción (Columnas 1 a 10 y 37 a 60)
* **Col A (1):** `ORDEN DE COMPRA` (`=+'CIERRE RECEP'!A13`)
* **Col B (2):** `Proveedor` (`=+'CIERRE RECEP'!B13`)
* **Col C (3):** `Categoría` (`=+'CIERRE RECEP'!C13`)
* **Col D (4):** `CANTIDAD DE BULTOS` (`=+'CIERRE RECEP'!AS13` — Bultos Muelle)
* **Col E (5):** `CANTIDAD DE UNIDADES` (`=+'CIERRE RECEP'!AT13` — Unidades Muelle)
* **Col F (6):** `CANTIDAD DE BULTOS PL` (`=+'CIERRE RECEP'!AW13` — Bultos Packing List)
* **Col G (7):** `CANTIDAD UNIDADES PL` (`=+'CIERRE RECEP'!AX13` — Unidades Packing List)
* **Col H (8):** `MAT. PROVEEDOR` (`=+'CIERRE RECEP'!D13`)
* **Col I (9):** `MATERIAL` (`=+'CIERRE RECEP'!E13` — Código SAP)
* **Col J (10):** `RECIBIDO REAL & FISICO PRK` (`=P13/O13` — Prepacks calculados: Unidades / Curva)
* **Col O (15):** `Curva` (`='CIERRE RECEP'!BB13`)
* **Col P (16):** `Qty Variante` (`=+'CIERRE RECEP'!H13` — Unidades físicas reales)
* **Col AB (28):** `CANTIDAD FACTURADA MOV-641` (`='CIERRE RECEP'!U13`)
* **Col AC (29):** `DIFERENCIA MOV 101 VRS MOV 641` (`=+AB13-J13`)

### Zona B: Variables Operativas Exclusivas de Despacho (Columnas 11 a 36)
* **Col K (11):** `CANTIDAD DISTRIBUIDA`: Unidades distribuidas/asignadas a sucursales.
* **Col L (12):** `PARTICIONES`: Particiones o entregas parciales.
* **Col M (13):** `FECHA DE DESPACHO`: Fecha en que salió el camión de bodega.
* **Col N (14):** `INVENTARIO DE LA MB 52`: Stock actual en sistema (Transacción SAP MB52).
* **Cols Q a W (17-23): Inconsistencias de Alisto:**
  * `Q`: Revuelto Saliendo / `R`: Destino de la tienda.
  * `S`: Revuelto Entrando / `T`: Origen de la tienda.
  * `U`: Faltante registrado en picking/alisto.
  * `V`: Dañado registrado durante el empaque o traslado.
  * `W`: Sobrante registrado en alisto.
* **Col Z (26):** `INDICADOR DE INCONSISTENCIAS`:
  $$Z_{13} = +Q_{13} + S_{13} + U_{13} + V_{13} + W_{13}$$
* **Cols AE a AG (31-33): Conciliación de Saldos:**
  * `AE`: Cantidad enviada a contenedores posteriores (Partes B, C, D).
  * `AF`: Stock en almacén de reserva/resurtido.
  * `AG`: Unidades en proceso de facturación (Movimiento SAP 643).
* **Col AH (34):** `UNIDADES PARA NOTA DE CRÉDITO POR DESPACHO`:
  $$AH_{13} = +U_{13} + V_{13} \quad (\text{Faltante} + \text{Dañado en despacho})$$
* **Col AI (35):** `REVISIÓN FINAL DE STOCK`:
  $$AI_{13} = AC_{13} + AE_{13} + AF_{13} + AG_{13}$$

---

## 4. Control de Inventarios y Hojas de Soporte

1. **`STOCK CONTROL DE INVENTARIO`:**
   * Valida la concordancia entre las ubicaciones y almacenes de libre utilización (Alisto Principal, Muelle de Descarga y Racks de Resurtido).
   * Monitorea:
     * Unidades que aplican para Nota de Crédito por faltante: `=+'CIERRE RECEP'!N1012`.
     * Unidades pendientes de contabilizar en facturación: `=+'CIERRE RECEP'!Z1012`.
     * Unidades que se enviarán en partes posteriores (B, C, D): `=+'CIERRE RECEP'!X1012`.
2. **`PARA N.C DESP`:**
   * Calcula la Nota de Crédito final por inconsistencias de despacho:
     $$\text{Total NCR Despacho} = \text{Gastos Logísticos} + \text{Recuperación de Mermas} + \text{Gastos Generales}$$
3. **`Check list DESPACHO`:**
   * 8 verificaciones de auditoría antes del archivo definitivo del expediente de distribución.

---

## 5. Preparación en ReconcileAI para Soporte Futuro

El modelo de datos y la arquitectura de ReconcileAI ya cuentan con los campos necesarios para activar la conciliación de despacho sin modificaciones estructurales:
* `PrepackItem.sap_mov_101` (Recepción).
* `PrepackItem.sap_mov_641` (Salida a tiendas).
* `PrepackItem.container_part` (Partición A, B, C, D...).
* `PrepackItem.revuelto`, `faltante`, `danado`, `sobrante`.

Cuando el área requiera formalizar el Cierre de Despacho, se podrá habilitar la pestaña correspondiente en la interfaz web reutilizando exactamente los mismos datos ya procesados durante la recepción.
