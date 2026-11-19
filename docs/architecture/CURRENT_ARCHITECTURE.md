# Current Architecture — ACR

Replaces the former Python-only description (audit of revision `589f6aa`, kept in git history).

## 1. Editions

```
              ┌──────────────────────────── shared business rules ─────────────────────────────┐
  QC .xlsm ──►│ read → classify (Cardinal Rule) → BOM split → packaging → strips/Bultos → export │──► closure .xlsm
  Closure  ──►│                                                                                  │    (VBA preserved)
              └──────────────────────────────────────────────────────────────────────────────────┘
  src/dotnet  (C# WPF, primary)     web/ (React/TS, browser only)     src/*.py (legacy CLI/Streamlit)
```

## 2. Desktop solution (`src/dotnet/Acr.sln`, `net9.0` / `net9.0-windows`)

| Project | Responsibility | Key types |
|---|---|---|
| `Acr.Core` | Pure domain, no file I/O (guarded by `CoreIsolationTests`) | `DiscrepancyClassifier`, `QuantityEngine`, `BomAllocationService`, `SapClipboardSanitizer`, `DailyLogLifecycleValidator`, `HistoricalIdentityMatcher`, `OperationalSafetyPolicy`, models (`DiscrepancyItem`, `ComparisonSummary`, `ChildSkuItem`, `BultosPresentationItem`), `SapKnowledge/*` |
| `Acr.Infrastructure` | OpenXML streaming I/O and orchestration | `QcReader`, `ClosureTemplateReader`, `WorkbookSignatureScanner`, `ReceivingPipeline`, `ClosureExporter`, `FormulaIntegrityAuditor`, `DailyLogRepository/Service`, `HistoricalClosureScanner`, `OperationalSafetyGuard`, `AtomicFileTransaction`, `JsonProcessingReportStore`, `OperationHistoryStore`, `DefaultSapCatalog` |
| `Acr.Desktop` | WPF + `CommunityToolkit.Mvvm`; one `MainViewModel`, `MainWindow.xaml` with 8 tabs | commands: browse, execute reconciliation, copy strips, sanitize SAP, export macro-safe `.xlsm`, search SAP |
| `Acr.Tests` | xUnit, 221 tests | Core, Infrastructure (incl. `Imp207HardeningTests`, `BlindReceivingTests`), Desktop view-model tests |

### Pipeline (`ReceivingPipeline.Execute`)

1. **Signature check** — `WorkbookSignatureScanner` classifies each input (`QcInspection`, `ReceivingClosure`, `HybridQcAndClosure`); swapped paths are exchanged.
2. **Read** — `QcReader` (sheet `CONTROL DATA`, header row 6; `BASE DE DATOS` for BOM; `PARTE A/B/C` detection ignoring `ESQUEMA`/`ALISTO` sheets) and `ClosureTemplateReader` (`DATOS GENERALES` + `CIERRE RECEP`, header row 12, data from row 13, declared units from `G`, cartons from `AW`, 20-empty-row early exit, no row cap).
3. **Classify** — `DiscrepancyClassifier.CompareMerchandise` yields `ComparisonSummary` (Shortage, Overage, RevueltoCandidate, CartonMismatchOnly, DamageOnly, Unexpected/Missing, ExactMatch) without offsetting.
4. **Enrich** — proportional BOM split of child SKUs (largest-remainder), curve, canonical packaging, NCR amount.
5. **Resolve container ID** — template ID (non-numeric) → QC ID → file-name regex `\b(IMP)[\s\-_]*([0-9]+[A-Z0-9\-]*)` → file name.
6. **Outputs** — `BultosItems`, Strip 1 (`AS:AT`), Strip 2 (`AW:AY`), optional macro-safe export (SHA-256 of `xl/vbaProject.bin` compared), JSON report and history when configured. Warnings include `[MONTAJE A CIEGAS]` when `ComparisonSummary.IsBlindReceiving` (declared > 0 and physical = 0).

### Safety

`OperationalSafetyGuard` canonicalizes paths and blocks writes to production data except explicit user save/export; writes go through `AtomicFileTransaction`.

## 3. Web edition (`web/src`)

`utils/excelParser.ts` (parsing, `extractCierreMetadata`, `allocateChildSkus`), `discrepancyClassifier.ts`, `quantityEngine.ts`, `sanitizer.ts`, `validator.ts`, `macroSafeXlsmExporter.ts`, `i18n.ts`; UI in `components/*Tab.tsx` (Bultos, Inspector, Semáforo, Curves, SapValidator, ExcelExport, Checklist, SapWiki) plus `KpiHeaderBar` and `FileExplorerView`. Everything runs client-side.

## 4. Legacy Python (`src/*.py`, `app.py`, `cli.py`)

Streamlit UI and CLI over openpyxl/pandas. Not extended with the IMP-207/BOM/blind-receiving fixes; use the C# or web edition for current behavior.

## 5. Desktop UI structure

Header + KPI strip → file loader (QC path, closure path, two independent Browse buttons, Run) → optional blind-receiving banner → 8 tabs → status bar. Global styles (tooltips, buttons, DataGrid selection contrast) live in `Window.Resources` of `MainWindow.xaml`; icons are `StreamGeometry` resources (no emojis).

## 6. Packaging

`empaquetar_standalone.bat` runs `dotnet publish -r win-x64 --self-contained -p:PublishSingleFile=true` (compressed, ~62 MB) and zips `web/dist`. Output folder `publish/` is git-ignored.
