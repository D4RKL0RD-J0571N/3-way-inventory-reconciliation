# Handoff & Contexto Completo para Continuación con otro LLM

> **Instrucción:** Copia y pega este documento completo como prompt inicial en tu nueva sesión con otro LLM para darle contexto exacto del proyecto, sus reglas de negocio, su arquitectura y el estado actual del código.

---

## 1. Contexto General y Repositorio

- **Proyecto:** **ACR — Asistente de Cierre de Recepción (3-Way Inventory Reconciliation)**
- **Repositorio GitHub:** `https://github.com/D4RKL0RD-J0571N/3-way-inventory-reconciliation.git`
- **Rama Activa:** `feat/web-react-and-csharp-parity` (sincronizada con `origin`)
- **Rutas Locales de Trabajo:**
  - Worktree actual: `C:\Users\Jostin\.gemini\antigravity\worktrees\3 Way Inventory Reconciliation\web_edition_csharp_parity`
  - Repositorio principal / Datos locales (NO versionados en Git): `D:\Development\Software Development\3 Way Inventory Reconciliation\data`
  - Muestra local de prueba para Montaje a Ciegas: `D:\Development\Software Development\3 Way Inventory Reconciliation\data\IMP-100_MONTAJE_A_CIEGAS\`

---

## 2. Arquitectura de las 3 Ediciones en el Repositorio

El repositorio mantiene tres ediciones bajo las mismas reglas de negocio:

1. **Edición de Escritorio C# / WPF (Principal — `src/dotnet/Acr.sln`):**
   - Target Framework actual: `net9.0` / `net9.0-windows` (el equipo tiene instalado .NET SDK `9.0.312`; `CoreIsolationTests` acepta `net9.0` o `net10.0`).
   - `Acr.Core`: Dominio puro sin I/O (`DiscrepancyClassifier`, `QuantityEngine`, `BomAllocationService`, `SapClipboardSanitizer`, `DailyLogLifecycleValidator`, `HistoricalIdentityMatcher`, `OperationalSafetyPolicy`, modelos y `SapKnowledge`).
   - `Acr.Infrastructure`: Lectores/escritores OpenXML en streaming (`QcReader`, `ClosureTemplateReader`, `WorkbookSignatureScanner`, `ClosureExporter`, `FormulaIntegrityAuditor`, `ReceivingPipeline`, `DailyLogRepository/Service`, `OperationalSafetyGuard`, `AtomicFileTransaction`, `DefaultSapCatalog`).
   - `Acr.Desktop`: Aplicación WPF con `CommunityToolkit.Mvvm` (`MainWindow.xaml` + `MainViewModel.cs`), 8 pestañas operativas, iconos vectoriales `StreamGeometry` (cero emojis) y tooltips en español en todos los controles.
   - `Acr.Tests`: Suite xUnit con **222 pruebas automatizadas pasando (0 fallos)**.

2. **Edición Web React + TypeScript (`web/`):**
   - Stack: React 19, TypeScript, Vite 6, Tailwind CSS 4, SheetJS (`xlsx`), JSZip, Lucide Icons.
   - Procesamiento 100% en navegador (`web/src/utils/excelParser.ts`, `discrepancyClassifier.ts`, `quantityEngine.ts`, `sanitizer.ts`, `macroSafeXlsmExporter.ts`) y 8 componentes de pestañas en `web/src/components/`.

3. **Edición Python / Streamlit (Legado — raíz y `src/*.py`):**
   - `app.py`, `cli.py`, `generar_cierre_oficial.py`. Mantenida únicamente para uso rápido por CLI o lotes.

---

## 3. Reglas de Oro y Comportamientos Críticos Implementados

Cualquier modificación futura **DEBE** preservar estas reglas verificadas por la suite de pruebas:

1. **Regla Cardinal de Discrepancias:**
   - **Jamás** compensar algebraicamente faltantes (`Shortage`) y sobrantes (`Overage`) entre materiales distintos. Cada discrepancia se reporta íntegra por separado. Si dos materiales del mismo contenedor tienen magnitudes opuestas equivalentes, se emparejan como `RevueltoCandidate` indicando la pareja (`REVUELTO ↔ código`), sin borrar las diferencias individuales.
2. **Materiales sin Movimiento (`IsZeroTotal`):**
   - Cuando `DeclaredUnits == 0 && PhysicalUnits == 0 && DeclaredCartons == 0 && PhysicalCartons == 0 && DamagedUnits == 0`, el ítem se clasifica como `SIN MOVIMIENTO (0 TOTAL)` con badge gris neutro (`#F1F5F9`).
   - Por defecto (`ExcludeZeroTotalFromStripsAndBultos = true`), se excluyen de las tiras de pegado (`AS:AT`, `AW:AY`) y de la hoja `Bultos` para no ensuciar el libro oficial.
3. **Empaque Canónico:**
   - Diferencia estrictamente entre cajas complementarias (`8*48 1*36`, cuando hay bulto complemento físico) y unidades sueltas sin bulto adicional (`8*48 +36`).
4. **Exportación `.xlsm` Macro-Safe:**
   - `ClosureExporter` modifica el libro de cierre preservando intacto `xl/vbaProject.bin` y verifica que el hash **SHA-256** del proyecto VBA de entrada y de salida sean idénticos.
5. **Blindaje `IMP-207` (Bug colisión `288`, Layout #4 y `ESQUEMA ALISTO`):**
   - *ID de Importación:* Al extraer metadatos de `DATOS GENERALES` y `CIERRE RECEP`, se excluyen etiquetas de resumen (`importadora`, `faltante`, `dañad`, `sobrant`, `revuelt`, `total`, `bulto`, `valor`, `costo`, `unidad`, `fecha`, `diferencia`, `ncr`, `reclamo`, `monto`) para que `"Faltante Importación" = 288` nunca sobrescriba `"IMP-207"`. Se bloquea el primer ID con prefijo `IMP-` y se respalda con regex `\b(IMP)[\s\-_]*([0-9]+[A-Z0-9\-]*)(?=[_\s.\-]|$|\b)` sobre el nombre del archivo.
   - *Layout #4:* En `QcReader.cs`, la columna de bultos complementos ignora encabezados que contengan `"UNIDADES"` o `"UNID"`, activando el rescate matemático de bultos.
   - *Hojas `ESQUEMA ALISTO`:* Se ignoran estrictamente al detectar partes (`PARTE A/B/C`), evitando 21 falsos `PENDIENTE_PARTE_B` en `IMP-207`.
   - *Sin límite de 500 filas:* `ClosureTemplateReader` lee en streaming sin tope artificial (soporta `IMP-187` fila 612+).
   - *Detección por firma OpenXML:* `WorkbookSignatureScanner` distingue `CONTROL DATA` (QC) vs `CIERRE RECEP` (Cierre). Si el usuario selecciona los archivos al revés, `ReceivingPipeline` los intercambia automáticamente; también soporta libros híbridos (`IMP-218`).
6. **Reparto Proporcional BOM (*Largest-Remainder Method*):**
   - `BomAllocationService.AllocateChildSkus` distribuye las unidades físicas recibidas del prepack padre entre las tallas hijas (`BASE DE DATOS`) proporcionalmente al ratio de curva, asignando los residuos por mayor fracción decimal con desempate determinista para que la suma de hijos sea exactamente igual al total físico del padre.
7. **Sanitizador SAP (`SapClipboardSanitizer`):**
   - Limpia tablas crudas pegadas desde SAP (`\t`, `|`, `;`, `,`), quita ceros a la izquierda en códigos de material sin corromper códigos con forma de notación científica (`1E2`), convierte negativos contables `(48)` → `-48` y monedas `$-50.00` → `-50.00`.
8. **Detección de Montaje a Ciegas (`IsBlindReceiving`):**
   - Cuando `TotalDeclaredUnits > 0 && TotalPhysicalUnits == 0` (ya sea porque se cargó un archivo QC con 0 unidades o porque se dejó vacío el campo QC y solo se cargó la Plantilla de Cierre), `ComparisonSummary.IsBlindReceiving` es `true`.
   - El pipeline emite la advertencia `[MONTAJE A CIEGAS]`, la UI WPF despliega un banner ámbar (`AVISO: MONTAJE A CIEGAS DETECTADO — PENDIENTE DE DESCARGA FÍSICA`), suspende el reclamo NCR en el checklist y marca pendiente la hoja de Bultos.
9. **Estética y Ergonomía UI en WPF (`MainWindow.xaml`):**
   - **Cero emojis:** Todos los iconos son geometrías vectoriales `StreamGeometry` (`GeoFolderOpen`, `GeoPlayReconcile`, `GeoPackage`, `GeoSearch`, `GeoTraffic`, `GeoBom`, `GeoClipboard`, `GeoExcelExport`, `GeoChecklist`, `GeoWiki`, `GeoCalculator`, `GeoCopy`, `GeoShieldCheck`).
   - **Barra de archivos en 9 columnas independientes:** Los dos botones `Examinar...` y el botón `Ejecutar Conciliación 3-Way` tienen columnas propias y jamás se solapan.
   - **Contraste en selección de tablas:** Las filas seleccionadas en `DataGrid` usan fondo celeste suave (`#E0F2FE`) y texto oscuro (`#0F172A` / `#0C4A6E` en *Empaque Canónico*), evitando texto azul ilegible sobre fondo azul.
   - **Tooltips en español** en absolutamente todos los botones, KPIs, filtros, entradas y pestañas.

---

## 4. Las 8 Pestañas de `Acr.Desktop` (`MainWindow.xaml` / `MainViewModel.cs`)

1. **Cuadre & Bultos:** Modos Conteo Muelle vs Estricto Aduanero, filtro *Ocultar en cero*, Calculadora Interactiva de Bultos/Sueltas, copia rápida de Tira 1 (`AS:AT`) y tabla de la hoja `Bultos`.
2. **Ficha 360° Material:** Buscador y selector de estilo padre, banner de estado y monto NCR, 4 tarjetas (Identificación + Fila Excel, Empaque & PO 43, Factura 41 & Físico, Balances & Revuelto) y tabla de tallas hijas BOM.
3. **Tabla Semáforo:** Filtros por texto y clasificación (`CUADRADO`, `FALTANTE`, `SOBRANTE`, `REVUELTO`, `DIF. BULTOS`, `DAÑADO`), badges semánticos, pareja revuelto y botón *Guía SAP* por fila.
4. **Explosión Curvas (BOM):** Selector de prepack con 4 KPIs y tabla del reparto proporcional *largest-remainder* por SKU hijo.
5. **Validador SAP (Portapapeles):** Área de pegado crudo, botón *Cargar Muestra SAP*, *Sanitizar Tabla SAP*, detector de delimitador y copia de TSV limpio.
6. **Copiar a Excel & Exportar:** Exportación `.XLSM Macro-Safe` con hash SHA-256 verificado de `xl/vbaProject.bin`, Tira 1 (`AS:AT`), Tira 2 (`AW:AY`) y vista previa de columnas `AS..AZ`.
7. **Checklist de Expediente Físico:** 7 documentos reglamentarios dinámicos con los totales reales de la importación (o estado adaptado en montaje a ciegas) y barra de progreso.
8. **SAP & Wiki:** Buscador de transacciones/tablas SAP conectado a `DefaultSapCatalog`, guías diagnósticas paso a paso y manual operativo SOP.

---

## 5. Comandos Clave de Compilación, Pruebas y Empaquetado

```powershell
# 1. Compilar solución C# en Release (0 errores, 0 advertencias)
dotnet build src/dotnet/Acr.sln -c Release

# 2. Ejecutar las 222 pruebas unitarias y de integración
dotnet test src/dotnet/Acr.sln

# 3. Ejecutar la aplicación de escritorio en modo desarrollo
dotnet run --project src/dotnet/Acr.Desktop

# 4. Compilar la edición Web (React + TypeScript)
cd web
npm install
npm run build
cd ..

# 5. Generar paquetes portables (publish/desktop-standalone/Acr.Desktop.exe y publish/web-edition.zip)
.\empaquetar_standalone.bat
```

---

## 6. Estado de Entregables y Siguientes Pasos Posibles

- **Entregables listos localmente (ignorados en Git):**
  - `publish/desktop-standalone/Acr.Desktop.exe` (~65.1 MB, ejecutable único win-x64 autocontenido sin requerir .NET Runtime).
  - `publish/web-edition.zip` (~301 KB, bundle estático de producción de `web/dist`).
- **Documentación actualizada:**
  - `README.md`, `README.es.md`, `docs/README.md` (índice maestro), `docs/WIKI.md`, `docs/architecture/CURRENT_ARCHITECTURE.md`, `docs/RELEASE_NOTES_DEMO.md` y `docs/quality/KNOWN_LIMITATIONS.md`.
- **Posibles tareas futuras si el usuario lo solicita:**
  1. Abrir un Pull Request o hacer merge de `feat/web-react-and-csharp-parity` hacia la rama principal (`main` o `repository_audit_planning`).
  2. Migrar `<TargetFramework>` de `net9.0` a `net10.0` cuando se instale el SDK de .NET 10 en la estación de trabajo.
  3. Añadir conversión automática o lector para archivos `.xls` binarios antiguos (BIFF8) si aparecen facturas de proveedores en ese formato.
