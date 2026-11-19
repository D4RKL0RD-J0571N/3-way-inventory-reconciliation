# ACR — Receiving Closure Assistant (3-Way Inventory Reconciliation)

[Versión en español](README.es.md) · [Documentation index](docs/README.md) · [Web edition](web/README.md)

ACR reconciles the three sources of an import container receiving: **PO 43 / Pedido 41 invoice** (declared), **dock Quality Control (QC)** and **packing list / ZDESCARGA cartons**, and produces the squared official closure workbook (`.xlsm`) with its VBA macros intact.

The repository contains **three editions** that share the same business rules:

| Edition | Location | Stack | Status |
|---|---|---|---|
| **Desktop (primary)** | `src/dotnet/` | C#, WPF/MVVM, OpenXML, `net9.0` (SDK 9.0.x) | Active, 221 automated tests |
| **Web** | `web/` | React 19, TypeScript, Vite 6, Tailwind 4, SheetJS | Active, reference UI for the desktop tabs |
| **Legacy** | `app.py`, `cli.py`, `src/*.py` | Python, Streamlit, openpyxl | Maintained for CLI/batch use only |

## Golden rules (enforced in all editions)

- **Cardinal Rule:** shortages and overages are never offset algebraically. Each is reported on its own.
- **Zero-total lines** (`IsZeroTotal`: 0 declared, 0 physical, 0 cartons) are shown as `SIN MOVIMIENTO (0 TOTAL)` in neutral gray and excluded from the paste strips and the `Bultos` sheet by default (`ExcludeZeroTotalFromStripsAndBultos = true`).
- **Canonical packaging:** `8*48 1*36` (complement carton) vs `8*48 +36` (loose pieces).
- **Macro-safe export:** `.xlsm` output preserves `xl/vbaProject.bin`; SHA-256 of source and output is compared.
- **Operational Spanish** terminology in every user-facing string.

## Desktop edition (`src/dotnet/Acr.sln`)

Projects: `Acr.Core` (domain, no I/O), `Acr.Infrastructure` (OpenXML readers/writers, pipeline, safety), `Acr.Desktop` (WPF), `Acr.Tests` (xUnit).

Eight tabs (vector icons, Spanish tooltips on every control):

1. **Cuadre & Bultos** — dock-count vs strict-customs modes, interactive carton/loose calculator, hide-zero filter, `Bultos` sheet projection.
2. **Ficha 360° Material** — identification, packaging/PO 43, invoice 41 vs physical, balances/revuelto/NCR, BOM child sizes.
3. **Tabla Semáforo** — status colors, revuelto pairing, canonical packaging, SAP guide shortcut.
4. **Explosión Curvas (BOM)** — prepack selector, KPIs, proportional per-size split (largest-remainder).
5. **Validador SAP (Portapapeles)** — sanitizes raw SAP tables (`\t`, `|`, `;`, `,`), strips leading zeros, preserves `1E2`, `(48)` → `-48`, `$-50.00`.
6. **Copiar a Excel & Exportar** — Strip 1 `AS:AT`, Strip 2 `AW:AY`, real differences `AU`/`AZ`, macro-safe `.xlsm` export with SHA-256 check.
7. **Checklist de Expediente Físico** — 7 dynamic items with real totals and progress bar.
8. **SAP & Wiki** — T-code catalog (`DefaultSapCatalog`), diagnostic guides, SOP manual.

### Notable behaviors

- **Import ID hardening (IMP-207 → 288 bug):** summary labels (`Faltante`, `Dañado`, `Sobrante`, `Total`, `Bulto`, `Valor`, `NCR`, …) can never overwrite the container ID; the first valid ID is locked and the file name (`IMP-207_...`) is the fallback.
- **Workbook detection by signature:** `CONTROL DATA` ⇒ QC, `CIERRE RECEP` ⇒ closure. Swapped inputs are auto-corrected; hybrid workbooks (IMP-218) are supported.
- **No row cap** (IMP-187 reads past row 600); `ESQUEMA ALISTO` sheets are ignored when detecting `PARTE A/B/C`.
- **Layout #4:** `UNIDADES BULTOS COMPLEMENTOS` headers are not mistaken for complement-carton counts (mathematical rescue is applied).
- **Blind receiving (montaje a ciegas):** when invoice units are present but physical units are 0, the app shows an amber banner, a `[MONTAJE A CIEGAS]` warning, suspends the NCR checklist item and does not treat the gap as a vendor claim.

### Build, test, run

```powershell
dotnet build src/dotnet/Acr.sln -c Release
dotnet test  src/dotnet/Acr.sln            # 221 tests
dotnet run --project src/dotnet/Acr.Desktop
```

Portable single-file executable (no .NET install needed): run `empaquetar_standalone.bat` → `publish/desktop-standalone/Acr.Desktop.exe` and `publish/web-edition.zip`.

> The projects target `net9.0` because the build machine has SDK 9.0.x. `CoreIsolationTests` accepts `net9.0` or `net10.0`; retarget the four `.csproj` files once the .NET 10 SDK is installed.

## Web edition (`web/`)

```bash
cd web && npm install && npm run dev      # http://localhost:3000
npm run build                              # production bundle in web/dist
```

Runs entirely in the browser (no upload). See [web/README.md](web/README.md).

## Legacy Python tools

```bash
pip install -r requirements.txt
python cli.py --qc path/QC.xlsm --cierre path/Closure.xlsm
streamlit run app.py
python -m pytest tests
```

Helper launchers: `iniciar_asistente.bat`, `preparar_equipo.bat`, `ver_resumen_rapido.bat`.

## Repository layout

```
src/dotnet/          C# solution (Acr.Core, Acr.Infrastructure, Acr.Desktop, Acr.Tests)
web/                 React + TypeScript edition
src/*.py, app.py     Legacy Python engine, Streamlit UI, cli.py
tests/               Python tests and e2e fixtures
tools/               Python build/stress utilities
docs/                Documentation (start at docs/README.md)
empaquetar_standalone.bat   Builds portable desktop exe + web zip
```

## Privacy

All processing is local (memory only). No data leaves the workstation; `data/` is never committed.

## License

MIT — free for operational, professional and educational use.
