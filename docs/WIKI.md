# Wiki Operativa — Asistente de Cierre de Recepción (ACR)

Guía para el operador de la edición de escritorio. Resumen técnico en [../README.es.md](../README.es.md); procedimiento SAP en [SOP_RECEPCION_SAP.md](SOP_RECEPCION_SAP.md).

## 1. Instalación

- **Portable:** copie `Acr.Desktop.exe` (carpeta `publish/desktop-standalone/`, generada con `empaquetar_standalone.bat`) a cualquier PC Windows x64 y haga doble clic. No requiere .NET, Python ni permisos de administrador.
- **Desde código:** `dotnet run --project src/dotnet/Acr.Desktop`.

## 2. Flujo de trabajo

1. **Control Calidad (QC):** botón *Examinar...* junto a la primera ruta; seleccione el libro con la hoja `CONTROL DATA`.
2. **Plantilla Cierre (Master):** segundo *Examinar...*; seleccione el libro con la hoja `CIERRE RECEP`. Si ambos libros se eligen al revés, la aplicación los intercambia por la firma de sus hojas.
3. **Ejecutar Conciliación 3-Way.** La barra superior muestra contenedor, PO, estilos (cuadrados/discrepantes), unidades y bultos facturados vs físicos, y revueltos.
4. Revise las pestañas, copie las tiras al libro oficial o exporte el `.xlsm`.

Pase el cursor sobre cualquier control: todos tienen tooltip en español.

## 3. Pestañas

| Pestaña | Para qué sirve |
|---|---|
| **Cuadre & Bultos** | Compara bultos de descarga (AS) con packing list (AW). *Modo Estricto Aduanero* exige coincidencia exacta de enteros y complementos. *Ocultar en cero* oculta líneas sin movimiento. La calculadora simula el empaque canónico: `NxM +K` (sueltas) o `NxM 1xK` (con bulto complemento). |
| **Ficha 360° Material** | Cuatro tarjetas por material: identificación y fila Excel; empaque y PO 43; factura 41 vs físico; averías, revuelto y monto NCR. Abajo, tallas hijas del BOM. |
| **Tabla Semáforo** | Matriz de discrepancias. Verde cuadrado, rojo faltante, ámbar sobrante, morado revuelto, azul diferencia de bultos, rosa dañado, gris `SIN MOVIMIENTO (0 TOTAL)`. El botón *Guía SAP* abre el protocolo de investigación. |
| **Explosión Curvas (BOM)** | Elija un prepack y vea el reparto proporcional por talla (método de mayor residuo; la suma siempre coincide con el total físico). |
| **Validador SAP (Portapapeles)** | Pegue tablas crudas de SAP (ME80FN, MB51, SE16N). *Sanitizar* quita ceros a la izquierda, conserva notación científica (`1E2`), convierte `(48)` en `-48` y `$-50.00` en `-50.00`. *Copiar TSV limpio* deja el resultado listo para Excel. |
| **Copiar a Excel & Exportar** | Tira 1 (`AS:AT`) y Tira 2 (`AW:AY`) para pegar desde la fila 13 de `CIERRE RECEP`; vista previa con diferencias `AU` y `AZ`; *Exportar .XLSM Macro-Safe* y hash SHA-256 verificado del proyecto VBA. |
| **Checklist de Expediente Físico** | Siete documentos del expediente con totales reales y barra de progreso. |
| **SAP & Wiki** | Buscador de T-Codes y tablas, guías diagnósticas y manual SOP. |

## 4. Reglas que la aplicación aplica

- **Regla Cardinal:** un faltante en un material y un sobrante en otro nunca se restan entre sí. Si coinciden en magnitud se marcan como *revuelto (candidato)* con su pareja, pero ambos siguen reportándose.
- **Sin movimiento (0 total):** se excluyen por defecto de las tiras y de la hoja `Bultos` para que el pegado no deje filas vacías.
- **ID de importación:** se toma de `DATOS GENERALES`/`CIERRE RECEP`; valores de resumen como `Faltante Importación = 288` se ignoran. Si no hay ID válido se usa el nombre del archivo (`IMP-207...`).

## 5. Montaje a ciegas

Cuando el cierre contiene solo valores de factura (sin QC, sin ZDESCARGA ni packing list, o con todo en 0), la aplicación:

- muestra el banner ámbar **«MONTAJE A CIEGAS DETECTADO»** y el estado *PENDIENTE DE DESCARGA FÍSICA*;
- agrega la advertencia `[MONTAJE A CIEGAS]` al resultado;
- suspende el ítem de Nota de Crédito y marca pendiente el de `Bultos` en el checklist;
- no debe usarse para tramitar reclamos: las diferencias son mercancía aún no contada.

Para probarlo, haga un respaldo de una importación existente y deje en 0 las columnas `H`, `AS`, `AT`, `AW` de `CIERRE RECEP` y las cantidades/bultos de `CONTROL DATA`.

## 6. Problemas frecuentes

| Síntoma | Causa / acción |
|---|---|
| El contenedor aparece como número (p. ej. 288) | Versión anterior; actualice. Verifique que el archivo se llame `IMP-###...`. |
| 21 partes B pendientes en IMP-207 | Versión anterior que contaba hojas `ESQUEMA ALISTO`; actualice. |
| Aviso de Layout #4 | La columna de bultos complemento no existe; se recalcula matemáticamente. Revise el resultado. |
| Texto ilegible en celda seleccionada | Corregido: la selección usa fondo celeste claro con texto oscuro. |
| Error al abrir `.xls` antiguo | Solo se soportan `.xlsx`/`.xlsm`; guarde el archivo en ese formato. |
