> **Estado: HISTÓRICO / Historical snapshot.** Describe el sistema Python-only (revisión 589f6aa) o un plan previo. Para el estado actual vea [docs/README.md](../README.md) y [CURRENT_ARCHITECTURE](../architecture/CURRENT_ARCHITECTURE.md).

# Target System Architecture — ReconcileAI (Track B)

**Assessment & Design Date:** 2026-11-17  
**Repository:** `https://github.com/D4RKL0RD-J0571N/3-way-inventory-reconciliation`  
**Baseline Audited Revision:** `589f6aa` (`repository_audit_planning` / `operational-benchmark`)  
**Planning Mode:** Strictly non-modifying target architecture specification (**R1**, **R3**, **R8**)  
**Aligned Companion Specs:** `docs/ui/REDESIGN_SPECIFICATION.md` (Track C), `docs/sap-knowledge/KNOWLEDGE_EXPANSION_PLAN.md` (Track D)

---

## 1. Executive Architectural Thesis: Evolutionary Refactoring, Not a Big-Bang Rewrite

ReconcileAI **must not undergo a ground-up rewrite**. Direct inspection of the codebase (`docs/architecture/CURRENT_ARCHITECTURE.md` and `docs/quality/CODE_AUDIT.md`) confirms that the five core domain modules are already cleanly decoupled from Streamlit and encode substantial, verified operational logic:

- **Existing Streamlit-Free Domain Core (`1,157` lines across 5 modules):**
  - `src/models.py` (`94` lines): `PackagingBreakdown`, `ChildSKU`, `PrepackItem` (already carrying Phase 2 fields `sap_mov_101`, `sap_mov_641`, `container_part`, `revuelto`), and `ReconciliationSummary`.
  - `src/sanitizer.py` (`138` lines): `clean_string`, `clean_material_code`, `clean_number`, `clean_int`, `parse_multi_po`.
  - `src/parsers.py` (`679` lines): `parse_packaging_string`, `extract_cierre_metadata`, `parse_quality_control_boms` (already using `read_only=True` streaming and Row 6 dynamic header scanning), `ItemList`, and `parse_cierre_recepcion` (Row 12 header refinement, QC cross-referencing, Blind Setup pause mode `PENDIENTE_CONTROL_CALIDAD`).
  - `src/engine.py` (`138` lines): `compute_summary`, `filter_items` (multi-part `A..Z+` and dynamic status filtering).
  - `src/validator.py` (`108` lines): `ValidationIssue`, `validate_item_integrity`, `validate_all_items` (6 domain integrity checks).
- **Verified Baseline Execution:** All `27` existing test functions in `tests/` (`827` lines) pass under `python -B -m pytest -p no:cacheprovider -v` (`27 passed in 2.86s`, `CODE_AUDIT.md:166–173`).

### Why a Big-Bang Rewrite Is Rejected (R3)
1. **Unnecessary Regression Risk:** Rewriting `src/parsers.py` and `src/engine.py` from scratch would discard hard-won domain logic—multi-container `PARTE <X>` resolution, Blind Setup gating (`PAUSA_SIN_CONTROL_CALIDAD`), `CIERRE RECEP` formula healing (`H`, `AU`, `AZ`), `'Bultos'` cross-sheet formulas, and corporate workbook quirks (`Tarima` column shifts)—without solving any problem that targeted repairs cannot fix faster and safer.
2. **Root Cause Is Missing Application Service Boundaries, Not Failed Domain Abstractions:** Nearly all architectural sprawl (`1,773` lines in `app.py`), cross-script behavioral drift (`cli.py`, `generar_cierre_oficial.py`, `tools/operational_stress_test.py`), and UI fragility (`src/i18n.py`, `src/ui_helpers.py`) result from mixing file discovery, formatting, `.xlsm` patching, and validation orchestration directly inside presentation blocks.

### Adopted Strategy: 2-Stage Service-Extraction-First / Strangler Refactoring
1. **Stage 1 — Repair Domain Core & Extract 8 Shared Services (`src/services/`):** Fix localized correctness defects in `src/` (`CODE_AUDIT.md` `A1-1`..`A10-10`), standardize Python packaging (`pyproject.toml`, `src/__init__.py`, `tests/__init__.py`), decouple `src/i18n.py` from Streamlit, and extract the **8 modular application services** (`Section 2.3`) so UI, CLI, MCP, AI pipelines, and benchmark scripts execute identical deterministic logic.
2. **Stage 2 — Strangle the Presentation Shell:** Replace the monolithic `app.py` presentation layer with a native **PySide6 (Qt for Python)** Windows Desktop ERP shell (`Section 3`) consuming `src/services/` directly—while maintaining a clean fallback path to a **Refactored Neutral-Gray Streamlit** UI over the exact same `src/services/` interfaces if browser/intranet deployment or zero-new-dependency constraints are confirmed by the user.

---

## 2. Keep / Improve / Isolate Matrix

### 2.1 What Remains (Preserve Core Responsibility & Pure Interfaces)

| Module | Current Lines | Retained Responsibilities & Rationale |
|---|---:|---|
| `src/models.py` | 94 | Retain `PackagingBreakdown`, `ChildSKU`, `PrepackItem`, and `ReconciliationSummary` as pure Python dataclasses with zero UI dependencies, keeping pre-wired Phase 2 fields (`sap_mov_101`, `sap_mov_641`, `container_part`, `revuelto`). Apply additive extensions only (`source_row`, `Decimal`/`int` typing, `sap_mov_102`, `loose_units` / `segments`). |
| `src/sanitizer.py` | 138 | Retain pure-function deterministic normalization API (`clean_string`, `clean_material_code`, `clean_number`, `clean_int`, `parse_multi_po`) while fixing localized parsing defects. |
| `src/parsers.py` | 679 | Retain `openpyxl` `read_only=True` streaming pattern (`parse_quality_control_boms`), `BASE DE DATOS` BOM extraction, `ItemList`/metadata contract, dynamic header scanning, and Blind Setup gating (`PENDIENTE_CONTROL_CALIDAD`). |
| `src/engine.py` | 138 | Retain pure functional aggregation (`compute_summary`) and dynamic filtering (`filter_items`) across arbitrary container parts (`A..Z+`) and statuses. |
| `src/validator.py` | 108 | Retain rule-based `ValidationIssue` dataclass (`severity`, `material_code`, `issue_type`, `message`, `recommendation`) and all 6 existing checks as the foundation of `ValidationService`. |

### 2.2 What Improves (Targeted Repairs Tied to `CODE_AUDIT.md` Defect IDs)

Every defect identified in `docs/quality/CODE_AUDIT.md` (`A1-1`..`A10-10`) maps to a bounded, test-verified repair:

| Improvement Area | Target Modules | Exact `CODE_AUDIT.md` Defect IDs Addressed | Concrete Repair Specification |
|---|---|---|---|
| **1. `Decimal` vs. `int` Typing & Precision** | `src/models.py`, `src/sanitizer.py`, `src/parsers.py`, `src/services/formatting_service.py` | `A5-1` (`75.0` vs `75`), `A5-2` (`int()` floor truncation), `A5-3` (IEEE-754 `3.1666666666666665`), `A5-4` (`:.2f` sub-cent loss), `A5-5` (calculator `,.0f`), `A5-6` (`'$-50.00'` $\rightarrow$ `50.0`, `'1,234'` ambiguity, silent `0.0` fallback) | Type discrete counts (`quantity_*_units`, `received_real_units`, `faltante`, `danado`, `sobrante`, `bultos_*`, `c_theo`, `c_rec`) as validated `int` (or `Decimal` quantized only after explicit integrality check—never silent `int()` floor truncation). Type monetary/prorated fields (`unit_cost`, `ncr_amount`, `total_ncr_amount`) as `Decimal` (`ROUND_HALF_UP` at calculation boundary) so `$3.1667` / `$550.08` are exact and `unit_cost * qty == extended_total`. Fix `clean_number` sign detection **before** currency symbol stripping (`'$-50.00'` $\rightarrow$ `Decimal('-50.00')`), preserve locale-explicit separator disambiguation, and emit structured warnings on malformed numbers instead of silent `0.0`. |
| **2. Packaging Grammar (`+<units>` & `raw_ay`) & Carton Fallback** | `src/models.py`, `src/parsers.py`, `src/services/reconciliation_service.py`, `src/services/export_service.py` | `A2-1` (`+<units>` dropped; `8*72+24` vs `56*12+672`), `A2-2` (`parts[0]`/`parts[1]` truncation on 3+ segments), `A2-3` (`raw_ay` ignored when `raw_as` scalar), `A2-4` (`PRK` conflated with `Bultos`), `A2-5` (duplicated Zero-Carton rule & `0` vs `"0"` drift), `A4-1` (`PrepackItem.um` ignored in `PRK`/`UN` conversions & carton fallback) | Extend `PackagingBreakdown` to store full arbitrary `segments: List[Tuple[int, int]]` plus `loose_units: int`. Update `parse_packaging_string` to parse trailing `+<units>` and disambiguate `A*B+<loose>` (`IMP-200`) vs `A*B+<total>` (`IMP-192`) against `AT`/`H` (`AT == A*B + C` vs `AT == C`). Parse `raw_ay` (`Empaque`) even when `raw_as` is scalar (`IMP-187`, `IMP-200`). Stop substituting `quantity_inv_41_prk` as physical cartons without explicit `PRK`-to-carton metadata, and unify the Zero-Carton rule inside `ExportService`. Normalize `PrepackItem.um` in `IngestionPipeline` (falling back from blank `CIERRE RECEP` Col `BC` to `CONTROL DATA` `'um'` before defaulting to `"PRK"`, `src/parsers.py:314,477`). Drive `PRK`/`UN` conversion and `curve_total` checks from `um` in `ReconciliationService` (treating `um == "UN"` as single units with `curve_total = 1` unless explicit `raw_ay`/`empaque_str` carton packing exists, rather than defaulting `curve = 12`). Guard carton fallback in `ExportService.apply_zero_carton_rule` (`app.py:1527,1631`, `generar_cierre_oficial.py:67`) so `um == "UN"` **never** divides `unidades_desc` by `curve_total` or substitutes `quantity_inv_41_prk` (`A2-4`), requiring explicit `bultos_descarga`, `qc_details['bultos']`, or parsed carton packing instead. |
| **3. Metadata Overwrite, Row Caps & Duplicate Records** | `src/parsers.py`, `src/services/ingestion_pipeline.py`, `src/services/reconciliation_service.py` | `A6-1` (`CONTROL DATA` duplicate material overwrite), `A6-2` (`BASE DE DATOS` duplicate parent appends), `A6-3` (multi-PO `qc_map` dock-count double-counting & duplicate UI/lookup keys), `A6-4` (`'importaci'` Row 22 `'288'` overwrite on `IMP-207`), `A8-4` (Row `500` fallback cap on `IMP-187` `612` rows) | Anchor `extract_cierre_metadata` to exact label matches (`'importación'` / `'importacion'`, explicitly excluding `'faltante'`, `'dañado'`, `'sobrante'`), preventing `'Faltante Importación' = 288` from overwriting `'IMP-207'`. Remove the `min(..., 500)` cap in metadata carton fallback summation. Accumulate split-pallet rows (`qty`, `bultos`, `unidades`, `total_prk`) per material in `CONTROL DATA` instead of overwriting, and replace-or-deduplicate reappearing parent BOMs in `BASE DE DATOS`. Prevent multiple PO rows sharing a `material_code` in `parse_cierre_recepcion` (`src/parsers.py:482-533`) from each independently cloning `qc_map[mat_code]` (`bultos`, `unidades`, `bultos_enteros`/`complementos`, fallback `rec_u`); instead allocate `qc_map[mat_code]` dock totals across matching PO rows (proportionally by invoiced `ped41_u` or once-per-pool with an explicit `IngestionWarning` for manual split verification). Disambiguate multi-PO items across `ReconciliationService`, `ValidationService`, `Inspector 360`, and `Curves` (`app.py:1080,1368`) using composite key `(material_code, order_po_43, source_row)` and surface `MATERIAL_EN_MULTIPLES_PO`. |
| **4. NCR Math, Col `J`/`revuelto`, `Mov 101` & Status Completeness** | `src/parsers.py`, `src/engine.py`, `src/services/export_service.py` | `A3-1` (`IMP-207` `$550.08` vs `$869.76`), `A3-2` (`sap_mov_101`, `sap_mov_641`, `revuelto` unused), `A3-3` (`app.py:1543,1547` hardcoded `AU=0`/`AZ=0`), `A3-4` (`compute_summary` ignores `sobrante`/`diff`) | Reconcile `ncr_amount` shortfall calculation (supporting `max(ped41 - rec, 0) * unit_cost` / Col `J` `revuelto` pending Open Question 1 confirmation). Parse `revuelto` from `CIERRE RECEP` / `DATOS GENERALES`, reconcile `sap_mov_101` (`Mov 101 - Mov 102`) vs `received_real_units` in `ReconciliationService`, and include `total_sobrante_units` and item-level `rec_u - ped41_u` variances in `ReconciliationSummary.status` so `"LIMPIO"` is never reported when surpluses or unannotated shortfalls exist. Replace hardcoded `"Dif Físico vs 41 (AU)": 0` and `"Dif Fact vs Sol (AZ)": 0` (`app.py:1543,1547`) with calculated `AU = AT - H` and `AZ = AX - H` inside `ExportService.build_paste_strips`. |
| **5. Partial Child Receipts (`c_rec`), Rounding & `PARTE B` False Positives** | `src/parsers.py` | `A1-1` (`c_rec = 0` on non-`LIMPIO` rows in `IMP-103`), `A1-4` (banker's rounding on `c_theo`), `A1-5` (silent `ratio=1`, `curve=0`/`12` defaults), `A3-5` (`ESQUEMA ALISTO PARTE B/C` false `PENDIENTE_PARTE_B` on `IMP-207`) | Allocate `c_rec` proportionally across `ChildSKU` entries using `received_real_units` (with largest-remainder rounding so `sum(c_rec) == received_real_units` and `sum(c_theo) == ped41_u`, never zeroing all children on `FALTANTE`/`SOBRANTE`). Restrict `PARTE <X>` detection to genuine part-manifest sheets containing valid material rows (ignoring `ESQUEMA ALISTO` layout sheets unless they hold unreceived container SKUs), and fall back from missing `CIERRE RECEP` `curve` to `qc_details[mat]['curve']` before raising a warning rather than silently assuming `12`. |
| **6. Material IDs & Multi-PO Splitting** | `src/sanitizer.py`, `src/parsers.py`, `src/validator.py` | `A7-1` (`'1E2'` $\rightarrow$ `'100'`, $>2^{53}$ precision loss), `A7-2` (`000` prefix rule), `A7-3` (`PARTE <X>` `isdigit() and len >= 7`), `A7-4` (`parse_multi_po` omits `-`) | Restrict scientific-notation decoding in `clean_material_code` strictly to true numeric exponential notation (`^\d+(?:\.\d+)?[eE][+-]?\d+$`) via `Decimal` (never `float`, leaving `'1E2'` alphanumeric codes intact). Normalize SAP `CHAR(18)` zero-padding consistently without stripping valid >10-digit EAN prefixes. Allow alphanumeric SKUs and `<7`-digit codes in `PARTE <X>` extraction. Update `parse_multi_po` regex to `r"[,;/\-]+"` (or smart PO-suffix expansion for `'4300006764-6804-6805'`) and wire it into `ValidationService`. |
| **7. Streaming `CIERRE RECEP`, Dynamic Headers, `source_row` & `.xlsm` Integrity** | `src/parsers.py`, `src/services/export_service.py` | `A8-1`, `A8-2`, `A8-3` (hardcoded rows `5`/`6`/`12` & missing Row 12 matchers), `A8-5` (`13 + idx` `.xlsm` row misalignment), `A8-6` (trailing `'Bultos'` rows not cleared), non-streaming `3.1s–4.5s` load | Switch `parse_cierre_recepcion` to `openpyxl.load_workbook(..., data_only=True, read_only=True)` for ingestion (bringing parse times from `3.1s–4.5s` in line with `0.2s–0.6s` QC reads), reserving `keep_vba=True` full workbook loading strictly for `ExportService.patch_official_xlsm`. Add Row 12 matchers for `cat`, `mov641`, `desc`, `unit_cost`, and `mov_101` at `Col BF` (`58`) alongside `Col I` (`9`); scan headers dynamically in `BASE DE DATOS` and `CONTROL DATA` (not only Row 6); wire `col_part` (`src/parsers.py:227,247-248,310-325`) as an explicit secondary container-part signal (cross-checked against `PARTE <X>` manifest sheets) or remove the dead `col_part` assignment; record `source_row: int` on every `PrepackItem` so `.xlsm` patching anchors directly to `item.source_row` instead of `13 + idx`; and clear stale trailing rows in `'Bultos'`. |
| **8. Discovery Unification & Legacy `.xls` Handling** | `src/config.py`, `src/services/discovery_service.py`, `src/services/ingestion_pipeline.py` | `A8-7` (62 BIFF `.xls` files unsupported), `A10-10` (1-file QC/Cierre collision, duplicate imports in `src/config.py:5-10`), `0/71` resolution in `data/`, `IMP-81` false blind setup | Replace all 3 divergent discovery implementations (`src/config.py:31–67`, `app.py:389–429`, `tools/operational_stress_test.py:50–82`) with `DiscoveryService`. Use recursive scanning (`rglob`), `~$` lock exclusion, and lightweight `read_only=True` sheet-signature inspection (`CONTROL DATA`/`BASE DE DATOS` vs `CIERRE RECEP`) so `IMP 81.xlsm` is recognized without `"calidad"` in the filename, single-file folders never auto-collide, and BIFF `.xls` files (`d0cf11e0...`) route to `LegacyBiffXlsAdapter` (optional `[xls]` `xlrd` extra) or raise an explicit, actionable unsupported-format error. Remove duplicate imports in `src/config.py:5–10`. |
| **9. Validation Surfacing, UI Decoupling, Packaging, Offline & Python 3.10+ Fixes** | `app.py`, `src/i18n.py`, `src/ui_helpers.py`, `src/validator.py`, `src/heuristic_parser.py`, `generar_cierre_oficial.py`, `tools/`, `ver_resumen_rapido.bat` | `A1-2`, `A1-3`, `A4-2`, `A10-1` (discarded `validation`), `A10-2` (`folder_input` reset), `A10-3` (`paginate_dataframe`), `A10-4` (`tab_clipboard`), `A10-5` (`icons8.com`), `A10-6` (`'c'` key hijack), `A10-7` (`tab_checklist`), `A10-8` (Python 3.10/3.11 `SyntaxError` at `app.py:1090`), `A10-9` (`generar_cierre_oficial.py` hardcoded print), `A8-8` (dead `src/heuristic_parser.py` `VKP2` hazard), Area 9 silent `try/except` catches | Surface all `ValidationService` issues in UI/CLI; expand `POSIBLE_CONFUSION_UDM` to bidirectional `PRK`/`UN` under-entry (`81` vs `1,728`); fix `validator.py:33` (`"tabla STPO/MAST"`, not `"transacción"`); eliminate `except Exception: pass` silent swallows (Area 9); replace `app.py:313` `icons8.com` URL with a bundled local monochrome SVG; fix `app.py:1090` f-string quotes for Python 3.10+; decouple `src/i18n.py` pure catalog from `streamlit` (`MutationObserver` & `'c'` hijack removed); retire/remove dead `src/heuristic_parser.py` (preventing its dangerous `"vkp2"` $\rightarrow$ `"unit_cost"` mapping); add `pyproject.toml`, `src/__init__.py`, `tests/__init__.py` so bare `pytest` passes; update `ver_resumen_rapido.bat` to prefer `.venv` & forward `%*`; bundle `cli.py` + `generar_cierre_oficial.py` in `tools/build_exe.py`; and fix `generar_cierre_oficial.py:124–125` to print actual `summary` metrics. |

---

### 2.3 What Isolates into 8 Shared Modular Application Services (`src/services/`)

All presentation shells (PySide6 Desktop UI, Refactored Streamlit fallback, CLI, future MCP server, future Track E AI/document pipelines, and automated stress/benchmark runners) **must** call only these 8 application services. Business math, workbook patching, formatting, and discovery are strictly prohibited inside UI widgets or adapter handlers.

```text
+---------------------------------------------------------------------------------------------------+
|                                   ADAPTERS & ENTRY POINTS                                         |
|  +----------------------+  +---------------+  +------------+  +----------------+  +------------+  |
|  | PySide6 Desktop Shell|  | Refactored    |  | Headless   |  | Future MCP     |  | Future AI/ |  |
|  | (Primary ERP UI)     |  | Streamlit     |  | CLI & .exe |  | Server (R6)    |  | OCR Engine |  |
|  | ui_desktop/          |  | (Fallback UI) |  | cli.py     |  | (Read-Only)    |  | (Track E)  |  |
|  +----------+-----------+  +-------+-------+  +-----+------+  +-------+--------+  +-----+------+  |
+-------------|----------------------|----------------|------------------|-----------------|---------+
              |                      |                |                  |                 |
              v                      v                v                  v                 v
+---------------------------------------------------------------------------------------------------+
|                            SHARED APPLICATION SERVICES  (src/services/)                           |
|                                                                                                   |
|  [1] DiscoveryService      [2] IngestionPipeline + DocumentAdapters   [3] ReconciliationService   |
|  [4] ValidationService     [5] ExportService                          [6] FormattingService       |
|  [7] AuditTrailService     [8] SapKnowledgeService (read-only over docs/sap-knowledge/index.json) |
+---------------------------------------------------------------------------------------------------+
              |                      |                |                  |                 |
              v                      v                v                  v                 v
+---------------------------------------------------------------------------------------------------+
|                                  DETERMINISTIC DOMAIN CORE  (src/)                                |
|  src/models.py  |  src/sanitizer.py  |  src/parsers.py  |  src/engine.py  |  src/validator.py     |
|  src/i18n.py (pure translation catalog, zero UI framework imports)                                |
+---------------------------------------------------------------------------------------------------+
              |                                                          |
              v                                                          v
+------------------------------------------+        +-----------------------------------------------+
|  LOCAL FILESYSTEM & WORKBOOKS (R/O + Out)|        |  LOCAL KNOWLEDGE & AUDIT ARTIFACTS            |
|  data/IMP-*/  |  *_Cierre_Oficial.xlsm   |        |  docs/sap-knowledge/index.json | run_manifests/ |
+------------------------------------------+        +-----------------------------------------------+
```

#### Concrete Contracts for the 8 Shared Services

1. **`DiscoveryService` (`src/services/discovery_service.py`)**
   - **Absorbs:** `src/config.py:31–67`, `app.py:389–429`, `tools/operational_stress_test.py:50–82`.
   - **Interface:**
     - `list_import_folders(root_dir: Optional[Path] = None) -> List[ImportFolderEntry]` (defaults to `DATA_DIR`)
     - `discover_case_files(folder_path: Path, inspect_sheets: bool = True) -> CaseDiscoveryResult`
   - **Contract & Guarantees:** Performs recursive `rglob`, excludes `~$` Excel lock files and `"final"`/`"resultado"` outputs, classifies `.xlsm`/`.xlsx` vs `.xls` (BIFF) vs `.pdf`/`.msg`/`.eml`/`.jpeg` inventory counts, inspects workbook sheet signatures (`CONTROL DATA`/`BASE DE DATOS` vs `CIERRE RECEP`) via `read_only=True` so `IMP 81.xlsm` is recognized even without `"calidad"` in the filename (`KNOWN_LIMITATIONS.md:31–34`), prevents single-file QC/Cierre auto-collision (`A10-10`), and returns structured `DiscoveryWarning` records (including `PermissionError`s, never `except Exception: pass`).
   - **Consumers:** View 1 Case Loader (Desktop/Fallback UI), `cli.py`, `generar_cierre_oficial.py`, `operational_stress_test.py`, MCP `list_cases` tool.

2. **`IngestionPipeline` & `DocumentAdapter`s (`src/services/ingestion_pipeline.py`)**
   - **Absorbs:** Direct `openpyxl` reads in `src/parsers.py`, `app.py:275–296, 511–552, 1394–1431`.
   - **Interface:**
     - `Protocol DocumentAdapter`: `supports(source: InputSource) -> bool`, `read_qc(...)`, `read_cierre(...)`, `read_raw_table(...)`.
     - Concrete & Planned Adapters:
       - `OpenXmlAdapter` (Core, `.xlsm`/`.xlsx` via `openpyxl` `read_only=True` streaming + `source_row` provenance tracking)
       - `SapClipboardAdapter` (Core, replaces dead `tab_clipboard`, parsing SAP GUI ALV grids with `|---|` borders, preserving leading-zero/alphanumeric `MATNR`s, outputting clean rows for export or `MB51`/`ME80FN` cross-check)
       - `LegacyBiffXlsAdapter` (Optional `[xls]` extra via `xlrd` for the 62 `.xls` files in `data/` — emits explicit actionable notice when extra is absent)
       - `PdfTextOcrAdapter` & `MsgEmlAdapter` (Future Track E `[ocr]`/`[mail]` adapters emitting `CandidateExtractedDocument` proposals requiring human review before entering `ReconciliationService`)
     - `IngestionPipeline.ingest_case(qc_source: Optional[InputSource], cierre_source: InputSource, is_blind_setup: bool = False) -> IngestionResult`
     - `IngestionPipeline.parse_sap_clipboard(raw_text: str, preset: Optional[str] = None) -> SapClipboardResult` (delegates to `SapClipboardAdapter.read_raw_table(raw_text, preset=preset)`, supports presets `None` / `"auto"` / `"mb51"` / `"me80fn"` / `"mb52"` / `"se16n"`, strips SAP GUI ALV `|---|` borders, preserves leading-zero and alphanumeric `MATNR`s via `clean_material_code`, parses numeric columns without float corruption, and returns `SapClipboardResult(detected_format, detected_preset, headers, rows, row_count, warnings)`)
   - **Contract & Guarantees:** Returns `IngestionResult(items, metadata, boms, qc_details, warnings, input_fingerprints)`. Every `PrepackItem` carries composite identity `(material_code, order_po_43, source_row)` (`source_row: int` 1-based Excel row in `'CIERRE RECEP'`), preserves `PrepackItem.um` from `CIERRE RECEP` / `CONTROL DATA` (`A4-1`), prevents multi-PO `qc_map[mat_code]` dock-count double-counting (`A6-3`), and retains raw/normalized packaging evidence.

3. **`ReconciliationService` (`src/services/reconciliation_service.py`)**
   - **Absorbs:** `app.py:268–310, 835–890, 1052–1069`, `src/engine.py`, `cli.py:67–69`, `generar_cierre_oficial.py:38–40`.
   - **Interface:**
     - `run_reconciliation(ingestion: IngestionResult, blind_mode: bool = False) -> ReconciliationRun`
     - `filter_items(items: List[PrepackItem], part_filter: Optional[str] = None, status_filter: Optional[str] = None, category_filter: Optional[str] = None, search_query: Optional[str] = None) -> List[PrepackItem]`
     - `project_bultos(items, mode: CartonProjectionMode) -> List[BultosRow]` (`DOCK_COUNT` vs `STRICT_CUSTOMS`)
     - `calculate_packaging(units: int, curve: int, standard_per_box: int, odd_units: int = 0) -> PackagingCalcResult` (replaces `app.py:796–816` inline math)
   - **Contract & Guarantees:** Pure deterministic Python + `Decimal`/`int` arithmetic. Drives `PRK`/`UN` conversion and `curve_total` semantics from `PrepackItem.um` (`A4-1`), disambiguates multi-PO items via composite key `(material_code, order_po_43, source_row)` (`A6-3`), handles Blind Setup (`PAUSA_SIN_CONTROL_CALIDAD`), reconciles `Mov 101 - Mov 102` (`sap_mov_101`), `sobrante`/`revuelto` statuses, proportional `ChildSKU` `c_rec`, and provides extension hooks for Phase 2 (`CIERRE DESP`: `Mov 641-642`, `MB52`, `Mov 643`) once macro `Col BF` vs `Col I` and `AB13-J13` unit parity are resolved.

4. **`ValidationService` (`src/services/validation_service.py`)**
   - **Absorbs:** `src/validator.py`, discarded `app.py:298, 576`.
   - **Interface:**
     - `validate_run(items: List[PrepackItem], summary: ReconciliationSummary) -> ValidationReport` (`errors`, `warnings`, `infos`, `issues_by_item_key`)
     - `issues_for_item(report: ValidationReport, item: PrepackItem, include_cross_po: bool = True) -> List[ValidationIssue]`
   - **Contract & Guarantees:** Executes all 6 existing `src/validator.py` rules (`CURVA_DESCUADRADA`, `BOM_FALTANTE` excluding `um == "UN"`, bidirectional `POSIBLE_CONFUSION_UDM`, `PRECIO_CERO_EN_RECLAMO`, `ESTILO_PROVEEDOR_VACIO`, `MATERIAL_EN_MULTIPLES_PO` via `parse_multi_po`) plus planned Phase 1/2 rules (`SAP_MOV_101_DESCUADRE`, `DESCARGA_VS_PACKING_DESCUADRE` `AU`/`AZ`, `EMPAQUE_INVALIDO` `AY` vs `AT`/`H`, `DISPATCH_MOV_641_DESCUADRE`). Indexes `ValidationReport.issues_by_item_key` by `(item.material_code, item.order_po_43, item.source_row)` so multi-PO rows never collide (`A6-3`), while cross-PO `MATERIAL_EN_MULTIPLES_PO` alerts are keyed by `item.material_code` and merged by `issues_for_item(...)` when `include_cross_po=True` (sorted `ERROR` -> `WARNING` -> `INFO`). Extends `ValidationIssue` with additive fields `order_po_43: str = ""`, `source_row: int = 0`, and `kb_entry_id: Optional[str] = None` so UI & AI companions link directly to `SapKnowledgeService`.

5. **`ExportService` (`src/services/export_service.py`)**
   - **Absorbs:** `app.py:835–967, 1455–1688`, `generar_cierre_oficial.py:43–127`, `tools/operational_stress_test.py:240–292`.
   - **Interface:**
     - `apply_zero_carton_rule(item: PrepackItem) -> CanonicalCartonValues` (**single canonical implementation**; guards against dividing `unidades_desc` by `curve_total` or substituting `quantity_inv_41_prk` when `um == "UN"`, `A4-1`, `A2-4`)
     - `build_core_table(items: List[PrepackItem]) -> CoreExportTable` (18 columns, clean `Decimal`-normalized values)
     - `build_paste_strips(items: List[PrepackItem]) -> PasteStripsResult` (Tira 1 `AS:AT`, Tira 2 `AW:AY` with integer counts `"65\t65"`, calculated `AU = AT - H` and `AZ = AX - H`, and `ZDESCARGA_PACKING_LIST_AS_AZ.xlsx`)
     - `export_bultos_xlsx(items: List[PrepackItem], mode: CartonProjectionMode, output_path: Optional[Path] = None) -> ExportArtifact`
     - `patch_official_xlsm(cierre_source: Path | bytes, items: List[PrepackItem], output_path: Optional[Path] = None) -> ExportArtifact`
   - **Contract & Guarantees:** Opens source `.xlsm` with `keep_vba=True` strictly at export time (never mutating the input file in place), anchors every `'CIERRE RECEP'` cell write (`AS`, `AT`, `AW`, `AX`, `AY`) to **`item.source_row`** (eliminating `13 + idx` `A8-5` misalignment), heals missing/non-formula `H`, `AU`, and `AZ` cells, links `'Bultos'` rows `3 + idx` and clears trailing template/footer rows (`A8-6`), and raises explicit `ExportError`s on failure (never `except Exception: pass` or muted captions).

6. **`FormattingService` (`src/services/formatting_service.py`)**
   - **Absorbs:** Scattered `int(...)`, `:.0f`, `:.2f`, `$%.2f`, and `f"{val}"` formatting across `app.py`, `cli.py`, `generar_cierre_oficial.py`.
   - **Interface:**
     - `format_count(value: int | Decimal | float, grouping: bool = True) -> str` (UI `"1,152"`, `"75"`, `"0"`; TSV `grouping=False` $\rightarrow$ `"65"`, `"1152"` — **never `.0`**; flags/formats genuinely non-integer quantities up to 4 trimmed decimals rather than silently `int()`-flooring)
     - `format_unit_cost(value: Decimal | float, currency: str = "$") -> str` (minimum 2 decimals `$2.13`; preserves up to 4–6 non-zero decimals `$3.1667` / `$1.380833` — eliminates `3.1666666666666665` `A5-3` and lossy `:.2f` `A5-4`)
     - `format_money_total(value: Decimal | float, currency: str = "$") -> str` (exact 2 currency decimals `$550.08`, `$869.76`, `$0.00`)
     - `format_tsv_strip_row(values: Sequence[Union[int, Decimal, str]]) -> str`
   - **Consumers:** All Desktop UI `QAbstractTableModel` delegates, Fallback UI tables, TSV clipboard generators, Block 1 `.xlsx` builder, `cli.py`, `generar_cierre_oficial.py`, MCP JSON/text formatters.

7. **`AuditTrailService` (`src/services/audit_trail_service.py`)**
   - **Absorbs:** Stale path-keyed `@st.cache_data` (`app.py:267–310, 341–343`), unkeyed `tab_checklist` state (`app.py:1695–1701`).
   - **Interface:**
     - `fingerprint_input(source: Path | bytes) -> InputFingerprint` (`path`, `sha256`, `size_bytes`, `mtime_ns`)
     - `get_cached_or_compute(fingerprints, blind_mode, app_version, parser_version, compute_fn) -> ReconciliationRun` *(cache invalidates automatically whenever file bytes, size, `mtime_ns`, `app_version`, or `parser_version` change)*
     - `record_run_manifest(run: ReconciliationRun, exports: List[ExportArtifact]) -> Path` (immutable local JSON manifest in `run_manifests/` recording input SHA-256s, app/parser version, Blind vs Certified state, validation summary counts, export hashes, and timestamps)
     - `load_checklist_state(import_id: str) -> BinderChecklistState` / `save_checklist_state(...)` (defaults all 7 steps to `False` per `import_id`, persisted locally)

8. **`SapKnowledgeService` (`src/services/sap_knowledge_service.py`)**
   - **Absorbs:** Hardcoded SAP strings in `src/validator.py` and static/contradictory prose in `app.py` `tab_wiki` (`1714–1773`).
   - **Interface:**
     - `get_index_status() -> KnowledgeIndexStatus` (`state: Literal["ready", "not_built", "empty", "corrupt"]`, `index_path: Path`, `entry_count: int`, `banner_message: str`, `error_detail: Optional[str] = None`)
     - `lookup_by_entry_id(entry_id: str) -> Optional[SapKnowledgeEntry]`
     - `lookup_for_issue(issue_type: str, context: Optional[Dict[str, Any]] = None) -> List[SapKnowledgeEntry]`
     - `search(query: str, source_type: Optional[str] = None, sap_release: Optional[str] = None) -> List[SapKnowledgeEntry]`
   - **Contract & Guarantees:** Strictly **read-only** local interface over `docs/sap-knowledge/index.json` (`KNOWLEDGE_EXPANSION_PLAN.md:338–427`). Anchored to **SAP ECC 6.0 (`version=6.00.29`)** (`https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29`), enforces `source_type` (`company_specific` Ekono vs `generic_sap`), surfaces `verification_status` badges (`confirmed`, `needs_env_check`, `conflicting`, `unverified`), emits `source_citation` + `ai_usage_policy` metadata, and **never** returns executable SAP write automation.
   - **Pre-Priority-#4 Bootstrap & Missing/Empty-Index Guarantee:** Because Roadmap Priority #3 (UI redesign) executes before Priority #4 (`docs/sap-knowledge/index.json` creation), `SapKnowledgeService` **must never raise `FileNotFoundError`, `JSONDecodeError`, or `KeyError`** on startup or lookup when `docs/sap-knowledge/index.json` is absent (`state="not_built"`), contains `0` entries (`state="empty"`), or fails parsing (`state="corrupt"`). When `state != "ready"`, `get_index_status().banner_message` returns **`"Knowledge Index Not Built (Roadmap #4 Pending)"`** (with `error_detail` populated if `"corrupt"`). `lookup_by_entry_id(...)` returns `None` safely; `lookup_for_issue(...)` and `search(...)` degrade gracefully (non-fatal) to **built-in `ValidationService` remediation metadata** (rule explanations, corrected `"tabla MAST/STPO"` text, `ME80FN` vs `VKP2` exclusion warning, `EINA` lookup) plus a **verified `SOP_RECEPCION_SAP` summary notice** (7-document binder checklist, `ME80FN`/`SE16N`/`MB51` queries, multi-warehouse `TOTAL DE UNIDADES` rule — explicitly flagging `docs/SOP_RECEPCION_SAP.md:1,30-31` `S/4HANA`/`EWM`/`VL31N`/`RFUI` claims as `conflicting` / unverified under the `ECC 6.0` `6.00.29` baseline).

---

## 3. UI Framework Trade-Off Analysis & Recommendation (R3 & R4)

### 3.1 Candidate Evaluation Matrix

Baseline verified on the target workstation (`Explorer 2`): `requirements.txt` contains 5 packages (`streamlit`, `openpyxl`, `pandas`, `pyinstaller`, `pytest`), Python `3.13.5` is active, `tkinter` (`Tk 8.6`) is in stdlib, and `PySide6`/`PyQt6`/`customtkinter`/`pywebview`/Node/Rust are not currently installed.

| Evaluation Dimension | Option A: Keep & Refactor Streamlit | Option B: PySide6 (Qt for Python) Desktop ERP | Option C: Tkinter / CustomTkinter | Option D: Local Web Stack (FastAPI + React / pywebview / Tauri / Electron) |
|---|---|---|---|---|
| **ERP Ergonomics (R4)** | **Fair / Constrained.** Can adopt light neutral-gray theme, compact CSS, and monochrome SVGs, **but** lacks native `QMenuBar`, `QStatusBar`, docked split panes, rerun-free row selection, and multi-pane master-detail (`Semáforo` + `Ficha 360°` + `Companion Drawer`). | **Excellent.** Native Windows `QMainWindow` with `QMenuBar`, `QToolBar` (16×16 monochrome SVGs), persistent `QStatusBar`, `QSplitter`/`QDockWidget`, native `QPalette`/QSS neutral-gray + restrained slate-blue styling, and zero full-script reruns. | **Poor for Dense Grids.** Supports native menus, but `ttk.Treeview` lacks scalable 18-column horizontal scrolling, frozen columns, per-cell numeric alignment, and HiDPI polish; `CTkTable` degrades past ~100 rows. | **High UI Fidelity, High Stack Overhead.** React + TanStack/AG Grid renders dense tables well, but pure `FastAPI + React` still lives in a browser tab unless wrapped in Edge WebView2 (`pywebview`), Chromium (`Electron`), or Rust (`Tauri`). |
| **Dense Data Tables & SAP/Excel TSV Clipboard** | `st.dataframe` uses Glide canvas; TSV copy relies on `components.html` sandboxed iframe + `execCommand` (`src/ui_helpers.py:191–282`) and `window.parent.document` JS observer + global `'c'` key suppression (`src/i18n.py:520–620`). | `QTableView` + `QAbstractTableModel` + `QSortFilterProxyModel` handles 5,000+ rows smoothly, aligns integers (`75`) vs prorated costs (`$3.1667`) via delegates, and copies/pastes native Windows TSV directly via `QApplication.clipboard()` (`Ctrl+C`/`Ctrl+V`). | Requires brittle manual `Treeview` key bindings and clipboard plumbing; poor multi-cell selection UX. | Browser clipboard still faces `isSecureContext` or focus restrictions unless bridged via Tauri/Electron/pywebview IPC. |
| **Windows Offline Portable `.exe` (`PyInstaller`)** | **Heavy & Fragile.** Bundles Tornado + PyArrow + Altair + Streamlit assets, launches a local HTTP server (`localhost:8501`) + browser tab (`run_standalone.py`), risks port collisions/Windows firewall prompts, and orphans background server on tab close. | **Clean Native `.exe`.** `PyInstaller --onedir --windowed` produces a self-contained native Windows binary (Qt6 ~80–130 MB, largely offset by dropping Streamlit/Altair/PyArrow/Tornado). Zero ports, zero browser popups, zero firewall prompts, ~1–2s cold start, **100% air-gapped offline by construction**. | **Smallest `.exe`**, 100% offline, zero ports/browsers (Tcl/Tk ~10–15 MB). | **Complex Dual-Toolchain Build.** Requires Node/Vite + PyInstaller (`pywebview`), ~150+ MB Chromium (`Electron`), or Rust toolchain (`Tauri`), plus port/IPC orchestration. |
| **Maintenance Burden (Python-Only Shop)** | **Lowest Initial Effort.** Zero new dependencies (`5` packages), but ongoing maintenance tax from Streamlit BaseWeb/DOM selector churn and rerun state workarounds. | **Moderate Initial, Low Long-Term.** 100% Python, zero JS/TS/Rust. Requires adding `PySide6` (respecting `LGPLv3` dynamic linking), typed Qt table models, and `QThreadPool` workers for workbook I/O. | **Low Dependency, High Boilerplate.** Emulating sortable/filterable ERP grids in `ttk.Treeview` produces brittle custom Tcl/Tk code. | **Highest.** Introduces TypeScript/React/npm (and Rust for Tauri) into a pure-Python repository with zero JS/TS/Rust tooling. |
| **Incremental Strangler Safety (R3)** | **In-place.** Extract `src/services/` first, simplify `app.py` to thin Streamlit shell. | **Clean Strangler Migration.** Extract `src/services/` first (preserving 100% of `src/` domain core); build thin `ui_desktop/` shell over `src/services/` while CLI/existing shell remain operational until cutover. | Compatible with `src/services/`, but widget limitations constrain Validation & Companion UX. | Forces HTTP/JSON serialization (`FastAPI`) over `src/services/` purely to feed a local single-user UI. |

### 3.2 Primary Recommendation & Documented Fallback

- **Primary Recommendation: Option B — PySide6 (Qt for Python) Desktop ERP Shell via 2-Stage Strangler Migration**
  1. **Evidence Alignment:** Repository evidence confirms that Streamlit forces brittle browser workarounds (`render_clipboard_copy_button` 45px iframe JS, `window.parent.document` `MutationObserver`, global `'c'` key suppression, `@st.cache_data` path-key hazard, `folder_input` rerun reset, localhost server overhead in `run_standalone.py`), whereas **R4** explicitly specifies traditional menus, toolbars, status bars, compact neutral-gray surfaces, crisp dense tables, and practical ERP ergonomics for Windows workstations.
  2. **Non-Blocking Parsing (`QThreadPool`):** `parse_cierre_recepcion` takes `3.1s–4.5s` synchronously (`CODE_AUDIT` / `KNOWN_LIMITATIONS.md:83`); even after switching to `read_only=True` streaming (`~0.2s–0.6s`), all `DiscoveryService`, `IngestionPipeline`, `ReconciliationService`, and `ExportService` calls **must** execute on `QThreadPool` `QRunnable` workers emitting Qt signals to progress indicators so the UI thread never freezes.
  3. **`LGPLv3` Packaging & CLI Inclusion:** `PySide6` is licensed under `LGPLv3`. Standard `PyInstaller --onedir` packaging keeps Qt shared DLLs dynamically linked in `_internal/` (rather than statically linked), preserving `LGPLv3` compliance for internal/proprietary enterprise distribution. Update `tools/build_exe.py` to bundle `cli.py` and `generar_cierre_oficial.py` (or expose `--cli` / `export` subcommands) alongside the desktop `.exe` so portable installations retain headless automation parity.
- **Documented Fallback: Option A — Refactored Neutral-Gray Streamlit over `src/services/`**
  - If user responses to Open Questions (`Section 6.4`) require **zero new UI dependencies** (`requirements.txt` locked to Streamlit) or **centralized browser/intranet multi-user hosting**, the exact same `src/services/` architecture supports Option A cleanly: replace `.streamlit/config.toml` dark theme and `app.py:43–258` CSS gradients with neutral-gray light styling, delete `icons8.com` and all emojis, remove `get_table_toolbar_js` (`MutationObserver` + `'c'` hijack), and wire `ValidationService` + `SapKnowledgeService` + `AuditTrailService` into thin Streamlit views—while explicitly accepting Streamlit's menu, rerun, and clipboard limitations.
- **100% Alignment Guarantee:** Both `TARGET_ARCHITECTURE.md` and `docs/ui/REDESIGN_SPECIFICATION.md` target **PySide6** as Primary and **Refactored Neutral-Gray Streamlit** as Fallback over the identical 8 `src/services/` contracts.

---

## 4. ReconcileAI Development Rules (R1) as Architectural Guardrails

Every rule in `ORIGINAL_REQUEST.md` (`R1`) is enforced as a concrete structural boundary rather than an honor-system convention:

| R1 Guardrail | Structural Boundary in Target Architecture | Concrete Enforcement & Verification Mechanism |
|---|---|---|
| **1. Deterministic Calculation Boundary** (*Never let AI replace deterministic business calculations*) | `ReconciliationService`, `ValidationService`, `ExportService`, and `FormattingService` own 100% of unit arithmetic, BOM curves, `Mov 101-102`, Zero-Carton rules, `AS:AT`/`AW:AY` strips, and `ncr_amount`. | Future AI/OCR/LLM modules (Track E) may **only** output structured `CandidateExtractedDocument` proposals into `IngestionPipeline` with field-level confidence scores; proposals cannot enter `ReconciliationService` without human confirmation and deterministic `ValidationService` checks. |
| **2. Epistemic & Citation Contract** (*Require AI recommendations to explain evidence and limitations; mark unverified SAP knowledge*) | `SapKnowledgeService` reads strictly from schema-validated `docs/sap-knowledge/index.json` (`KNOWLEDGE_EXPANSION_PLAN.md:361–376`). | Every response/card returns `{entry_id, title, source_type, sap_release, sap_doc_version, verification_status, source_citation[], ai_usage_policy, limitations}`. Entries tagged `needs_env_check`, `unverified`, or `conflicting` render a mandatory caution banner in the UI Companion Drawer and are never presented as confirmed Ekono SOP truth. |
| **3. No Silent Data Discard or Mutation** (*Never silently modify or discard user data*) | Original `.xlsm`/`.xlsx`/`.xls` files are opened read-only (`read_only=True` during ingestion); zero `except Exception: pass` blocks (`CODE_AUDIT.md` Area 9). | `IngestionPipeline` emits structured `IngestionWarning`s for malformed cells/unparsed rows; `ExportService` writes strictly to new output paths/buffers (`*_Cierre_Oficial_Final.xlsm`) anchored by `item.source_row`; `AuditTrailService` records SHA-256 input and output hashes. |
| **4. Read-Only SAP Air-Gap** (*Do not allow SAP write operations without explicit authorization and separate safety design*) | Zero SAP GUI write scripting, write RFCs, or write BAPIs (`MIGO`, `MIRO`, `VL31N`, `MR11`, `SE16N &SAP_EDIT`) exist anywhere in `src/services/`. | `SapKnowledgeService` is read-only (`ACTVT=03` guidance only). Any future live SAP connector must live in a separate isolated module requiring explicit security sign-off, read-only default roles, and dual authorization for any write extension. |
| **5. Minimal Offline Dependency Policy** (*Do not introduce dependencies without purpose & maintenance justification*) | Core package remains lightweight, 100% offline, and `PyInstaller`-packable with zero runtime CDN requests (`app.py:313` `icons8.com` removed; local monochrome SVGs bundled). | Optional capabilities are isolated behind optional extras (`[xls]` `xlrd`, `[ocr]` `pymupdf`/OCR, `[ai]` vector/LLM clients) plugged into `DocumentAdapter` / Track E interfaces without bloating core offline installs. |
| **6. Bounded Verification & Git Discipline** (*Inspect before modifying; verify tests executed; bounded subsystems*) | Standardized package layout (`pyproject.toml`, `src/__init__.py`, `tests/__init__.py`) ensures bare `pytest` executes cleanly. | Every defect fix or service extraction requires targeted unit + golden-fixture regression tests across all `15` `sandbox_benchmark` cases before committing or rebuilding `tools/build_exe.py`. |

---

## 5. Planned Repository Module Layout (Target State)

```text
3-way-inventory-reconciliation/
├── pyproject.toml                   # Package metadata + pytest pythonpath=["."] + extras ([web-fallback], [xls], [ocr], [ai])
├── requirements.txt                 # Core desktop runtime deps (openpyxl, pandas, PySide6, pyinstaller, pytest); streamlit isolated in optional [web-fallback] extra
├── src/
│   ├── __init__.py                  # Explicit package marker (fixes PEP 420 bare-pytest collection)
│   ├── config.py                    # Path constants (BASE_DIR, INTERNAL_DIR, DATA_DIR, DOCS_DIR)
│   ├── models.py                    # Core dataclasses (Decimal/int typed + source_row provenance)
│   ├── sanitizer.py                 # Deterministic normalization (Decimal sci-notation, multi-PO '-')
│   ├── parsers.py                   # Streaming read_only=True QC/BOM/Cierre parsers + raw_ay/+units
│   ├── engine.py                    # Pure 3-way summary, Mov 101-102 & Phase 2 readiness, filter_items
│   ├── validator.py                 # Rule definitions backing ValidationService
│   ├── i18n.py                      # Pure-Python ES/EN catalog (zero emojis, zero Streamlit/DOM JS)
│   └── services/                    # 8 shared modular application services
│       ├── discovery_service.py
│       ├── ingestion_pipeline.py    # DocumentAdapter protocol + OpenXml/Clipboard/(opt .xls/OCR)
│       ├── reconciliation_service.py
│       ├── validation_service.py
│       ├── export_service.py        # Canonical Zero-Carton, TSV strips, source_row .xlsm patcher
│       ├── formatting_service.py    # Whole-number (75) vs Decimal unit cost ($3.1667) & total ($550.08)
│       ├── audit_trail_service.py   # SHA-256+mtime+version cache, JSON manifests, persistent checklist
│       └── sap_knowledge_service.py # Read-only query over docs/sap-knowledge/index.json (graceful pre-#4 fallback)
├── ui_desktop/                      # Primary PySide6 ERP Desktop Shell (QMainWindow, views, delegates)
│   └── assets/icons/                # Bundled 16x16 monochrome SVGs (zero external CDN calls)
├── ui_web_fallback/                 # Optional Refactored Neutral-Gray Streamlit adapter (requires [web-fallback] extra, if selected)
├── cli.py                           # Headless CLI (Discovery + Ingestion + Reconciliation + Validation)
├── generar_cierre_oficial.py        # Headless batch .xlsm patcher (calls ExportService + prints real KPIs)
├── tools/
│   ├── build_exe.py                 # Bundles desktop .exe + cli.py + generar_cierre_oficial.py + assets
│   └── operational_stress_test.py   # Benchmark harness calling shared services
└── tests/
    ├── __init__.py
    └── ...                          # Unit, golden-fixture (15 benchmark cases), service & UI model tests
```

---

## 6. Epistemic Separation (R8)

### 6.1 Confirmed Findings
- The domain core (`src/models.py`, `src/sanitizer.py`, `src/parsers.py`, `src/engine.py`, `src/validator.py` = `1,157` lines) has zero Streamlit imports and passes `27/27` tests via `python -B -m pytest -p no:cacheprovider -v`.
- All 49 numbered defects (`A1-1`..`A10-10`), the 16-entry Area 9 `try/except` catalog, `0/71` discovery failure in `src/config.py`, `IMP-81` false blind setup, `62` BIFF `.xls` files in `data/`, `3.1s–4.5s` non-streaming `parse_cierre_recepcion`, 4-way `.xlsm`/TSV patching drift, `app.py:298,576` unrendered `validation`, `app.py:313` external `icons8.com` fetch, `app.py:1090` Python 3.12+ quote syntax, `src/i18n.py:520–620` DOM/shortcut injection, dead `src/heuristic_parser.py` (`VKP2` hazard), and bare `pytest` collection error (`ModuleNotFoundError: No module named 'src'`) are directly verified in the repository and read-only external data.

### 6.2 Architectural Decisions & Assumptions
- **Decision:** Evolutionary **Service-Extraction-First / Strangler** refactoring into 8 shared services under `src/services/`, retiring `src/heuristic_parser.py` and decoupling `src/i18n.py` / `src/ui_helpers.py` from Streamlit.
- **Decision:** **PySide6 (Qt for Python)** as the Primary Target UI (`QMainWindow` ERP ergonomics, native `QTableView` + clipboard TSV, air-gapped `--onedir` `.exe`, `QThreadPool` workers), with **Refactored Neutral-Gray Streamlit** as a documented zero-dependency/browser-hosted fallback over the identical `src/services/` contracts.
- **Assumption:** ReconcileAI primarily runs locally on Windows analyst workstations (corroborated by `iniciar_asistente.bat`, `preparar_equipo.bat`, and `tools/build_exe.py`), and `PySide6` `LGPLv3` dynamic DLL bundling via `PyInstaller --onedir` complies with organizational IT policy.
- **Assumption:** Phase 2 (`CIERRE DESP`) will consume the same `PrepackItem` stream once `Mov 101-102`, `Mov 641-642`, `MB52`, and `revuelto` parsers are implemented and `Col BF` vs `Col I` / `AB13-J13` ambiguities are verified.

### 6.3 Unverified Domain & SAP Items
- Exact accounting composition behind `IMP-207 (NCR)` `PARA N.C RECEP.!G6 = 869.76` vs. shortfall-only `max(ped41 - rec, 0) * unit_cost` (`$890.82`) and net `(ped41 - rec) * unit_cost` (`$750.90`).
- Canonical disambiguation rule for `+` inside packaging strings (`IMP-200` `'8*72+24'` loose remainder vs `IMP-192` `'56*12+672'` repeated total).
- Whether downstream corporate VBA macros (`Módulo1`) require string `"0"` or integer `0` in Col `AY` on zero-carton rows, and whether `AW:AY` should preserve or zero Packing List values when `received == 0 and ped41 > 0`.
- Live SAP ECC 6.0 verification checklist items (`KNOWLEDGE_EXPANSION_PLAN.md §3.6`: `System -> Status`, `41...` `BSART`, `643` cross-company vs intra-company, `LE-WM` vs decentralized `EWM`, Z-objects).

### 6.4 Open Questions for the User
1. **Desktop vs. Browser Deployment & `PySide6` Approval:** Should ReconcileAI target native Windows workstations with **PySide6** (`~80–130 MB` `LGPLv3` Qt6 runtime in `--onedir` `.exe`), or must it remain browser-hosted / strictly locked to the existing 5 `requirements.txt` packages (selecting the **Refactored Neutral-Gray Streamlit** fallback)?
2. **Bundled Headless CLI:** Should `cli.py` and `generar_cierre_oficial.py` be bundled inside the portable `dist/ReconcileAI/` distribution (`tools/build_exe.py`) for batch/automation use?
3. **NCR & Zero-Carton Business Rules (`IMP-207`, `AW:AY`, `+` Grammar):** Which exact deterministic formulas should `ReconciliationService` and `ExportService` enforce for (a) `ncr_amount` (Col `J`/unannotated shortfalls), (b) `AW:AY` when `received == 0 and ped41 > 0`, (c) Col `AY` `0` vs `"0"`, and (d) `+` packaging notation?
4. **Legacy `.xls` Priority:** Should `LegacyBiffXlsAdapter` (`xlrd` optional `[xls]` extra for the 62 `.xls` files in `data/`) be prioritized in Roadmap Phase 1 (Correctness) or Phase 5 (Document Intelligence)?
