# ACR — Asistente de Cierre de Recepción (Conciliación 3-Way de Inventarios)

[English version](README.md) · [Índice de documentación](docs/README.md) · [Edición Web](web/README.es.md)

ACR concilia las tres fuentes de una recepción de contenedor de importación: **PO 43 / factura Pedido 41** (declarado), **Control de Calidad (QC) en muelle** y **packing list / bultos de ZDESCARGA**, y genera el libro oficial de cierre `.xlsm` cuadrado, con sus macros VBA intactas.

El repositorio contiene **tres ediciones** con las mismas reglas de negocio:

| Edición | Ubicación | Tecnología | Estado |
|---|---|---|---|
| **Escritorio (principal)** | `src/dotnet/` | C#, WPF/MVVM, OpenXML, `net9.0` (SDK 9.0.x) | Activa, 221 pruebas automáticas |
| **Web** | `web/` | React 19, TypeScript, Vite 6, Tailwind 4, SheetJS | Activa, referencia visual de las pestañas |
| **Legado** | `app.py`, `cli.py`, `src/*.py` | Python, Streamlit, openpyxl | Solo para CLI / lotes |

## Reglas de oro (en todas las ediciones)

- **Regla Cardinal:** faltantes y sobrantes jamás se compensan algebraicamente; cada uno se reporta por separado.
- **Líneas en cero** (`IsZeroTotal`: 0 facturado, 0 físico, 0 bultos) se muestran como `SIN MOVIMIENTO (0 TOTAL)` en gris neutro y se excluyen de las tiras de pegado y de la hoja `Bultos` por defecto (`ExcludeZeroTotalFromStripsAndBultos = true`).
- **Empaque canónico:** `8*48 1*36` (bulto complemento) vs `8*48 +36` (piezas sueltas).
- **Exportación Macro-Safe:** el `.xlsm` conserva `xl/vbaProject.bin`; se compara el SHA-256 de origen y salida.
- **Terminología operativa en español** en toda cadena visible.

## Edición de escritorio (`src/dotnet/Acr.sln`)

Proyectos: `Acr.Core` (dominio, sin E/S), `Acr.Infrastructure` (lectores/escritores OpenXML, pipeline, seguridad), `Acr.Desktop` (WPF), `Acr.Tests` (xUnit).

Ocho pestañas (iconos vectoriales, tooltips en español en cada control):

1. **Cuadre & Bultos** — modos Conteo Muelle / Estricto Aduanero, calculadora interactiva de bultos y sueltas, filtro «Ocultar en cero», hoja `Bultos`.
2. **Ficha 360° Material** — identificación, empaque/PO 43, factura 41 vs físico, balances/revuelto/NCR, tallas hijas BOM.
3. **Tabla Semáforo** — colores por estado, pareja de revuelto, empaque canónico, acceso a guía SAP.
4. **Explosión Curvas (BOM)** — selector de prepack, KPIs y reparto proporcional por talla (mayor residuo).
5. **Validador SAP (Portapapeles)** — limpia tablas crudas de SAP (`\t`, `|`, `;`, `,`), ceros a la izquierda, respeta `1E2`, `(48)` → `-48`, `$-50.00`.
6. **Copiar a Excel & Exportar** — Tira 1 `AS:AT`, Tira 2 `AW:AY`, diferencias reales `AU`/`AZ`, exportación `.xlsm` Macro-Safe con SHA-256.
7. **Checklist de Expediente Físico** — 7 ítems dinámicos con totales reales y barra de progreso.
8. **SAP & Wiki** — catálogo de T-Codes (`DefaultSapCatalog`), guías diagnósticas y manual SOP.

### Comportamientos relevantes

- **Blindaje del ID (bug IMP-207 → 288):** etiquetas de resumen (`Faltante`, `Dañado`, `Sobrante`, `Total`, `Bulto`, `Valor`, `NCR`, …) nunca sobrescriben el ID; se bloquea el primer ID válido y el nombre de archivo (`IMP-207_...`) es el respaldo.
- **Detección por firma de libro:** `CONTROL DATA` ⇒ QC, `CIERRE RECEP` ⇒ cierre. Se corrigen archivos intercambiados y se soportan libros híbridos (IMP-218).
- **Sin tope de filas** (IMP-187 lee más allá de la fila 600); las hojas `ESQUEMA ALISTO` se ignoran al detectar `PARTE A/B/C`.
- **Layout #4:** los encabezados `UNIDADES BULTOS COMPLEMENTOS` no se confunden con bultos complemento (se aplica rescate matemático).
- **Montaje a ciegas:** si hay unidades facturadas pero el físico es 0, se muestra un aviso ámbar, la advertencia `[MONTAJE A CIEGAS]`, se suspende el ítem de NCR del checklist y la diferencia no se trata como reclamo al proveedor.

### Compilar, probar, ejecutar

```powershell
dotnet build src/dotnet/Acr.sln -c Release
dotnet test  src/dotnet/Acr.sln            # 221 pruebas
dotnet run --project src/dotnet/Acr.Desktop
```

Ejecutable portable de un solo archivo (sin instalar .NET): ejecute `empaquetar_standalone.bat` → `publish/desktop-standalone/Acr.Desktop.exe` y `publish/web-edition.zip`.

> Los proyectos apuntan a `net9.0` porque el equipo de compilación tiene el SDK 9.0.x. `CoreIsolationTests` acepta `net9.0` o `net10.0`; al instalar el SDK de .NET 10 basta con cambiar los cuatro `.csproj`.

## Edición Web (`web/`)

```bash
cd web && npm install && npm run dev      # http://localhost:3000
npm run build                              # bundle en web/dist
```

Se ejecuta completamente en el navegador (sin subir archivos). Ver [web/README.es.md](web/README.es.md).

## Herramientas Python (legado)

```bash
pip install -r requirements.txt
python cli.py --qc ruta/QC.xlsm --cierre ruta/Cierre.xlsm
streamlit run app.py
python -m pytest tests
```

Lanzadores: `iniciar_asistente.bat`, `preparar_equipo.bat`, `ver_resumen_rapido.bat`.

## Estructura del repositorio

```
src/dotnet/          Solución C# (Acr.Core, Acr.Infrastructure, Acr.Desktop, Acr.Tests)
web/                 Edición React + TypeScript
src/*.py, app.py     Motor Python legado, UI Streamlit, cli.py
tests/               Pruebas Python y fixtures e2e
tools/               Utilidades Python de empaquetado/estrés
docs/                Documentación (iniciar en docs/README.md)
empaquetar_standalone.bat   Genera exe portable de escritorio + zip web
```

## Privacidad

Todo el procesamiento es local (solo memoria). Ningún dato sale del equipo; `data/` nunca se versiona.

## Licencia

MIT — uso operativo, profesional y educativo.
