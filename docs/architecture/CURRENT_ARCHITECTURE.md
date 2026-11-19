# Current System Architecture — ReconcileAI (Track A)

**Assessment Date:** 2026-11-17  
**Repository:** `https://github.com/D4RKL0RD-J0571N/3-way-inventory-reconciliation`  
**Audited Revision:** `589f6aa` (`repository_audit_planning` / `operational-benchmark`, 8 commits ahead of `main` `c706868`)  
**Audit Mode:** Strictly read-only / non-modifying

---

## 1. Executive System Overview & Runtime Topologies

**ReconcileAI** is a Python 3.13 / Streamlit / OpenPyXL / Pandas desktop-oriented web and CLI tool that reconciles three operational import-receiving inputs in memory:
1. **ERP Purchase Orders (`PO 43`)** and **Commercial/Customs Invoices (`Pedido Nacionalización 41` / DUA)** embedded inside corporate macro-enabled Excel workbooks (`CIERRE RECEP` / `DATOS GENERALES` sheets).
2. **Warehouse Dock Quality Control & Size-Curve Workbooks** (`CONTROL DATA`, `BASE DE DATOS`, and `PARTE <X>` sheets).
3. **Physical Carton & Packaging Breakdowns** (`AS:AT` dock discharge vs. `AW:AY` vendor packing list).

The application has **no database, no HTTP backend API, no authentication layer, no background job queue, and zero AI, OCR, PDF, email, RAG, vector store, or MCP dependencies**. `requirements.txt` (`lines 1–10`) declares exactly five Python packages:
- `streamlit>=1.35.0`
- `openpyxl>=3.1.2`
- `pandas>=2.0.0`
- `pyinstaller>=6.0.0`
- `pytest>=7.4.0`

### 1.1 Runtime Topologies

ReconcileAI executes in three distinct topologies over a shared `src/` directory:

| Topology | Entry Chain | Bundled / Available Scope | Configuration & Runtime Behavior |
|---|---|---|---|
| **1. Portable Standalone (`PyInstaller --onedir`)** | `tools/build_exe.py` $\rightarrow$ `dist/ReconcileAI/ReconcileAI.exe` (`run_standalone.py`) $\rightarrow$ `app.py` | Bundles `app.py`, `src/`, `.streamlit/`, `docs/`, plus metadata/data for `streamlit`, `openpyxl`, `pandas`, `altair`, `pyarrow`, `tornado` (`tools/build_exe.py:31-53`). **Excludes** `cli.py` and `generar_cierre_oficial.py`. | `run_standalone.py:11-30` resolves `sys._MEIPASS`, `os.chdir(base_dir)`, sets `sys.argv` with `--global.developmentMode=false --server.headless=false --browser.gatherUsageStats=false`, and invokes `streamlit.web.cli.main()`. Post-build generates `dist/ReconcileAI/INICIAR_RECONCILEAI.bat`, `data/LEEME.txt`, `LEEME_PORTABLE.txt`, and `dist/ReconcileAI_Portable.zip`. |
| **2. Local Python Streamlit Server** | `preparar_equipo.bat` (installs `.venv`), `iniciar_asistente.bat` $\rightarrow$ `run_standalone.py` $\rightarrow$ `app.py` (or `streamlit run app.py`) | Full source tree. `iniciar_asistente.bat:11-45` tries: (1) `dist\ReconcileAI\ReconcileAI.exe`, (2) `.venv\Scripts\python.exe run_standalone.py`, (3) system `python run_standalone.py`, (4) exit `1`. | `.streamlit/config.toml:1-13` configures `[server]` (`headless = true`, `enableCORS = false`, `enableXsrfProtection = false`, `maxUploadSize = 50`) and `[theme]` (`base = "dark"`, `primaryColor = "#0284C7"`, `backgroundColor = "#0F172A"`, `secondaryBackgroundColor = "#1E293B"`, `textColor = "#F8FAFC"`, `font = "sans serif"`). Note that `run_standalone.py:26` overrides `headless = true` via CLI `--server.headless=false`. In addition, `app.py:43–258` injects a 215-line inline `<style>` block hardcoding dark-mode surfaces (`#0F172A`/`#1E293B`), gradients (`button[kind="primary"]`, `.app-header`, `.welcome-card`), `.detail-card`, and `.status-banner-*` classes (complemented by `SKELETON_CSS` in `src/ui_helpers.py:13–91`). |
| **3. Headless CLI & Batch Scripts** | `cli.py` (via `ver_resumen_rapido.bat` or `python cli.py`), `generar_cierre_oficial.py`, `tools/operational_stress_test.py` | Uses `src/config.py`, `src/parsers.py`, `src/engine.py`; only `tools/operational_stress_test.py` calls `src/validator.py`. | `ver_resumen_rapido.bat:1-5` invokes system `python cli.py` (ignores `.venv` and forwards no `%*` flags). `cli.py` prints KPIs + `items[:5]`. `generar_cierre_oficial.py` patches `{import_id}_Cierre_Oficial_Final.xlsm` on disk. `tools/operational_stress_test.py` benchmarks 15 hardcoded cases from `I:\My Drive\IMP` into `sandbox_benchmark/`. |

---

## 2. Exact Module Inventory & Dependency Graph

Verified across the repository on disk:
- **Implicit PEP 420 Namespace Packages:** Neither `src/__init__.py` nor `tests/__init__.py` exists on disk (contradicting `README.md:133,154` and `README.es.md:132,153`).
- **Python Environment:** `win32` / Python `3.13.5` / `pytest-9.0.2`.

### 2.1 Source, Entry-Point, Tooling & Test Inventory

| File | Lines | Internal Imports | External / Stdlib Imports | Key Public Symbols (Line Ranges) & Notes |
|---|---:|---|---|---|
| `src/models.py` | 94 | None | `dataclasses`, `typing` | `PackagingBreakdown` (`7–16`), `ChildSKU` (`18–26`), `PrepackItem` (`28–67`), `ReconciliationSummary` (`69–94`) |
| `src/config.py` | 67 | None | `pathlib`, `typing`, `sys` *(duplicate `Path`/`typing` imports at `5–6` & `9–10`)* | `SRC_DIR`, `BASE_DIR`, `INTERNAL_DIR`, `DATA_DIR`, `DOCS_DIR` (`13–29`), `list_available_local_imports` (`31–41`), `resolve_local_import_files` (`43–67`) |
| `src/sanitizer.py` | 138 | None | `re`, `typing`, `math` *(lazy inside `clean_number:73`)* | `clean_string` (`9–23`), `clean_material_code` (`25–59`), `clean_number` (`61–115`), `clean_int` (`117–125`), `parse_multi_po` (`127–138`, unused in production) |
| `src/parsers.py` | 679 | `src.models`, `src.sanitizer` | `re`, `openpyxl`, `typing` | `parse_packaging_string` (`17–72`), `extract_cierre_metadata` (`74–167`), `parse_quality_control_boms` (`169–335`), `ItemList` (`337–343`), `parse_cierre_recepcion` (`345–679`) |
| `src/engine.py` | 138 | `src.models` | `typing` | `compute_summary` (`7–91`), `filter_items` (`93–138`) |
| `src/validator.py` | 108 | `src.models` | `dataclasses`, `typing` | `ValidationIssue` (`10–16`), `validate_item_integrity` (`18–76`), `validate_all_items` (`78–108`) |
| `src/heuristic_parser.py` | 72 | None | `typing`, `openpyxl` | `COLUMN_SYNONYMS` (`10–20`), `normalize_text` (`22–29`), `detect_headers_in_sheet` (`31–72`) — **DEAD MODULE** (zero importers) |
| `src/i18n.py` | 620 | None | `typing`, `streamlit` | `DEFAULT_LANGUAGE` (`12`), `LANGUAGES` (`13–16`), `TRANSLATIONS` (`19–456`), `get_current_language` (`459–466`), `set_language` (`469–472`), `t` (`475–493`), `render_language_selector` (`496–517`), `get_table_toolbar_js` (`520–620`) |
| `src/ui_helpers.py` | 283 | None | `streamlit`, `streamlit.components.v1`, `json`, `pandas`, `typing` | `SKELETON_CSS` (`13–91`), `inject_skeleton_css` (`94–96`), `render_skeleton_kpis` (`99–110`), `render_skeleton_table` (`113–124`), `render_skeleton_dashboard` (`127–137`), `paginate_dataframe` (`140–188`), `render_clipboard_copy_button` (`191–282`) |
| `app.py` | 1,773 | `src.config`, `src.parsers`, `src.engine`, `src.validator`, `src.sanitizer`, `src.models`, `src.i18n`, `src.ui_helpers` | `streamlit`, `pandas`, `openpyxl`, `os` *(unused)*, `io`, `time`, `hashlib` *(unused)*, `collections`, `pathlib`, `streamlit.components.v1` | `run_audit` (`267–310`), `make_f360_label` (`1073–1078`), `highlight_status` (`1328–1339`); Vista 1 Explorer/Uploader (`346–568`) gated by `st.stop()`, Vista 2 8-tab Dashboard (`570–1773`) |
| `cli.py` | 101 | `src.config`, `src.parsers`, `src.engine` | `sys`, `os` *(unused)*, `argparse`, `pathlib` | `main` (`20–98`); prints summary KPIs and `items[:5]` |
| `generar_cierre_oficial.py` | 130 | `src.config`, `src.parsers`, `src.engine` | `sys`, `pathlib`, `openpyxl` | `main` (`20–127`); writes `{import_id}_Cierre_Oficial_Final.xlsm` |
| `run_standalone.py` | 32 | *(invokes `app.py` via `stcli`)* | `sys`, `os`, `pathlib`, `streamlit.web.cli` | `main` (`11–29`) |
| `tools/build_exe.py` | 130 | *(bundles `run_standalone.py`, `app.py`, `src/`, `.streamlit/`, `docs/`)* | `sys`, `os`, `shutil`, `pathlib`, `PyInstaller.__main__`, `PyInstaller.utils.hooks` | `build` (`18–127`) |
| `tools/operational_stress_test.py` | 363 | `src.parsers`, `src.engine`, `src.validator`, `src.sanitizer` | `os`, `sys`, `shutil`, `time`, `json`, `pathlib`, `datetime`, `typing`, `openpyxl` | `find_files_for_folder` (`50–82`), `extract_human_cierre_metrics` (`84–171`), `run_stress_test_suite` (`173–360`) |
| `tests/test_bultos_empaque.py` | 238 | `src.models`, `src.parsers`, `src.engine` | `pytest` | 8 test functions (`10–236`) |
| `tests/test_i18n.py` | 64 | `src.i18n`, `src.ui_helpers` | `pytest`, `pandas` | 5 test functions (`16–64`) |
| `tests/test_parsers.py` | 397 | `src.parsers`, `src.engine`, `src.config`, `src.models` | `openpyxl`, `pathlib` | 7 test functions (`10–396`) |
| `tests/test_sanitizer.py` | 128 | `src.sanitizer`, `src.models`, `src.validator` | `pytest` | 7 test functions (`15–128`) |

Non-Python repository files verified: `.gitignore` (`40` lines), `.streamlit/config.toml` (`13` lines), `requirements.txt` (`10` lines), `iniciar_asistente.bat` (`45` lines), `preparar_equipo.bat` (`32` lines), `ver_resumen_rapido.bat` (`5` lines), `README.md` (`173` lines), `README.es.md` (`172` lines), `docs/AUDITORIA_DESPACHO_MACRO.md` (`118` lines), `docs/CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md` (`175` lines), `docs/RELEASE_NOTES_DEMO.md` (`137` lines), `docs/SOP_RECEPCION_SAP.md` (`120` lines), `docs/WIKI.md` (`202` lines).

### 2.2 Directed Module Dependency Graph

```text
+-------------------------------------------------------------------+
|                   Packaging & Launch Wrappers                     |
|  tools/build_exe.py --(bundles)--> run_standalone.py, app.py,     |
|                                    src/, .streamlit/, docs/       |
|  run_standalone.py  --(stcli)----> app.py                         |
+-------------------------------------------------------------------+

+------------------+   +--------------------+   +----------------------------+
|     app.py       |   |       cli.py       |   | generar_cierre_oficial.py  |
| (1,773 lines UI  |   |   (101 lines CLI)  |   |     (130 lines patcher)    |
|  + export + calc)|   +---------+----------+   +-------------+--------------+
+--------+---------+             |                            |
         |                       |      +---------------------+
         |                       v      v
         |               +--------------+        +-------------------------------+
         +-------------> | src/config.py|        | tools/operational_stress_test |
         |               +--------------+        |         (363 lines)           |
         |                                       +---------------+---------------+
         |   +---------------------------------------------------+
         v   v
+------------------------------------------------------------------------+
|                  Domain Core (Zero Streamlit Dependency)               |
|                                                                        |
|  src/validator.py -------> src/models.py <------- src/engine.py        |
|                                 ^                                      |
|                                 |                                      |
|  src/sanitizer.py <------ src/parsers.py                               |
+------------------------------------------------------------------------+

+------------------------------------------------------------------------+
|            Streamlit-Coupled Modules Inside src/ (UI Layer)            |
|  src/i18n.py --------> streamlit (session_state, sidebar, JS)          |
|  src/ui_helpers.py --> streamlit, streamlit.components.v1, pandas      |
+------------------------------------------------------------------------+

+------------------------------------------------------------------------+
|                       Orphaned / Dead Module                           |
|  src/heuristic_parser.py --> openpyxl (ZERO incoming import edges)     |
+------------------------------------------------------------------------+
```

---

## 3. End-to-End Data Flows

### 3.1 (a) Three-Way Divergent Input Discovery

Input file discovery is implemented in **three incompatible places** across the codebase:

1. **CLI & Batch Discovery (`src/config.py:31–67`):**
   - `list_available_local_imports()` lists immediate subdirectories of `DATA_DIR` (`BASE_DIR / "data"` with `INTERNAL_DIR / "data"` fallback).
   - `resolve_local_import_files(import_id)` iterates **non-recursively** (`imp_dir.iterdir()`) over `[".xlsm", ".xlsx", ".xls"]`, sorted ascending by filename length, skipping `"final"` or `"resultado"` in the name. It matches `"control" and "calidad"` (or `"qc"` if unset) for `qc_file` and `["progress", "cierre", "plantilla", "recep"]` for `cierre_file`. It does **not** filter `~$` Excel lock files.
   - **Verified failure rate:** Across all `71` `IMP-*` directories in `D:\Development\Software Development\3 Way Inventory Reconciliation\data`, `resolve_local_import_files()` resolves **0/71 pairs** (`0` full matches, `3` missing QC, `3` missing Cierre, `65` missing both) because QC workbooks routinely reside in subfolders (e.g., `INFORMACION GENERAL/`) or lack `"calidad"`/`"qc"` in the filename (e.g., `IMP-81/INFORMACION GENERAL/IMP 81.xlsm` and `IMP-81/IMP-81.xlsm`).
2. **Streamlit UI Disk Explorer & Uploader (`app.py:389–429, 511–552`):**
   - Disk Explorer (`app.py:389–429`) takes `folder_input` (defaulting to `f"data/{local_folders[0]}"`), performs a **recursive** `p.rglob("*")` for `[".xlsm", ".xlsx", ".xls"]`, and excludes `~$` lock files.
   - Ranks `qc_candidates_idx` by `["calidad", "qc", "conteo"]` and `cierre_candidates_idx` by `["cierre", "plantilla", "inprogress", "in progress", "recep"]` excluding `"final"`, with fallback `"imp-" in f.lower()` excluding `["calidad", "qc", "conteo", "portada"]`.
   - File Uploader (`tab_uploader`, `app.py:511–552`) restricts uploads to `type=["xlsm", "xlsx"]` (excluding `.xls`).
3. **Benchmark Script Discovery (`tools/operational_stress_test.py:50–82`):**
   - Hardcodes `IMP_DIR = Path(r"I:\My Drive\IMP")` (`line 29`), inspects root `.xlsm`/`.xlsx` files ranked by `("borrador" in name, "sin control" in name, -size)` for Cierre, and scans subdirectories matching `"general"` (or root) for `["control", "calidad", "qc"]`.

### 3.2 (b) Quality Control & BOM Parsing (`src/parsers.py:169–335`)

1. `parse_quality_control_boms(qc_source)` attempts `openpyxl.load_workbook(qc_source, data_only=True, read_only=True)` (`line 178`), falling back to non-`read_only` mode on any `Exception` (`lines 179–180`).
2. **BOM Extraction (`BASE DE DATOS`):** Selects first sheet matching `['BASE DE DATOS', 'BASE_DATOS', 'BOM', 'CURVAS']` (`line 183`), iterates from `min_row=5` until `empty_streak > 30`, reading fixed 0-based columns `4` (`bom_id`), `8` or `5` (`mat_val`), `9` or `6` (`ratio_val`, default `1`), and `10` (`um_val`). When `bom_id != current_bom_id or um_val == 'PRK'`, sets `current_parent = mat_clean`; subsequent rows append `{'child': mat_clean, 'ratio': ratio}` to `boms_by_parent[current_parent]`.
3. **QC Sheet Header Scan (`CONTROL DATA`):** Selects first sheet matching `['CONTROL DATA', 'CONTROL_DATA', 'CONTROL', 'QC']` (or `wb.sheetnames[0]`). Scans **only Row 6** (`lines 234–237`) to remap 13 column indices (`col_mat`, `col_part` [never read later], `col_desc`, `col_um`, `col_qty`, `col_curva`, `col_bultos`, `col_unidades`, `col_bultos_enteros`, `col_unid_enteros`, `col_bultos_comp`, `col_unid_comp`, `col_total_prk`).
4. **Multi-Part Sheet Scan (`PARTE <X>`):** Matches all sheet names against `r'PARTE\s+([A-Z0-9]+)'` (`line 276`), scans rows `1..500`, first 6 columns (`row_p[:6]`), and maps any numeric string of length `>= 7` to `part_materials[clean_m] = p_code` (`lines 282–290`).
5. **QC Row Loop:** Reads `CONTROL DATA` starting at `min_row=7`, storing per-material dicts in `qc_details[m_clean]` (overwriting duplicates) plus `qc_details['_detected_parts']`.

### 3.3 (c) Cierre Parsing & Blind Setup Pause Mode (`src/parsers.py:74–167, 345–679`)

1. `parse_cierre_recepcion` loads the closure workbook **without `read_only=True`** (`openpyxl.load_workbook(cierre_source, data_only=True)` at `line 360`), explaining observed `3.1s–4.5s` parse times vs. `0.2s–0.6s` for streaming QC reads.
2. `extract_cierre_metadata(wb)` scans `'DATOS GENERALES'` rows `1..24` (`cols B–C`) and `'CIERRE RECEP'` rows `1..9` (`cols A–B`), with a fallback summing `parse_packaging_string(AS)` and `parse_packaging_string(AW)` across rows `13..min(max_row, 500)`.
3. Row 12 of `'CIERRE RECEP'` (`cols 1..64`) is scanned to update `col_map` (`lines 392–436`), though `cat` (`3`), `mov641` (`21`), and `unit_cost` (`32`) lack header matchers.
4. Rows `13..ws.max_row` with non-empty `clean_material_code` (`lines 440–443`) resolve:
   - **Packaging & Cartons:** Calls `parse_packaging_string` on `AS` (`raw_as`) and `AW` (`raw_aw`), while reading `AY` (`raw_ay`) only via `clean_string`. Falls back to `qc_map[mat_code]` when `bultos_desc == 0` or `unidades_desc == 0` (`lines 482–509`).
   - **Physical Units (`rec_u`, `lines 527–534`):** Precedence is `Col H` (if non-empty and non-`"0"`) $\rightarrow$ `Col AT` (`unidades_desc`) $\rightarrow$ `qc_map[mat]['unidades']` $\rightarrow$ `0.0`.
   - **Container Part (`lines 541–565`):** `qc_map[mat]['part']` $\rightarrow$ `partes_bcd > 0` $\rightarrow$ regex `PARTE ([A-Z0-9]+)` in `desc`/col 20/col 29 $\rightarrow$ `"A"` if in QC/BOM else first non-`A` in `_detected_parts`.
   - **Blind Setup Mode (`line 568`):** Triggered explicitly (`is_blind_setup=True`) or implicitly when `not qc_map and not boms_map and raw_rec_u is None and unidades_desc == 0`. Forces `status = "PENDIENTE_CONTROL_CALIDAD"` and zeroes `rec_u`, `falt`, `dan`, `sob`, `diff`, and `ncr_amount`.
   - **Status, NCR & Child BOM Explosion (`lines 577–628`):** Computes `ncr_amount = (falt + dan) * unit_cost` and builds `ChildSKU` entries with `c_rec = int(c_theo) if status == "LIMPIO" else 0`. Returns `ItemList(items, metadata=meta)`.

### 3.4 (d) Reconciliation & Validation (`src/engine.py`, `src/validator.py`)

- `compute_summary` (`src/engine.py:7–91`) sums units, PRKs, cartons, NCR, and `container_parts_breakdown`, reconciles missing header carton totals from line-item totals (`lines 53–58`), and sets `status = "PAUSA_SIN_CONTROL_CALIDAD"` if blind, else `"LIMPIO"` if `total_falt == 0 and total_dan == 0`, else `"CON_DISCREPANCIAS"`.
- `filter_items` (`src/engine.py:93–138`) filters dynamically by `part_filter`, `status_filter`, `category_filter`, and `search_query` (matching parent fields and `children`).
- `validate_all_items` (`src/validator.py:78–108`) evaluates `CURVA_DESCUADRADA`, `BOM_FALTANTE`, `POSIBLE_CONFUSION_UDM`, `PRECIO_CERO_EN_RECLAMO`, `ESTILO_PROVEEDOR_VACIO`, and `MATERIAL_EN_MULTIPLES_PO`.
  - **Computed-but-Discarded Defect:** `app.py:298` runs `validate_all_items(items)`, caches the result in `run_audit`, and unpacks `validation = audit_res["validation"]` at `app.py:576`, **but never reads or renders `validation` anywhere in the UI**. Neither `cli.py` nor `generar_cierre_oficial.py` calls `validate_all_items`; only `tools/operational_stress_test.py:233,340–341` logs counts.

### 3.5 (e) Excel Patching, Formula Self-Healing & TSV Strips

ReconcileAI produces four export mechanisms across three files:
1. **Block 1 Core Table (`app.py:1455–1494`):** Exports an 18-column `pd.DataFrame` to sheet `CIERRE_RECEP_VALORES` (`.xlsx`).
2. **Stand-Alone `'Bultos'` Sheet (`app.py:956–967` & `tools/operational_stress_test.py:240–259`):** Exports `{import_id}_Hoja_Bultos_Oficial.xlsx`.
3. **TSV Paste Strips (`app.py:1509–1589`):** Formats Tira 1 (`AS:AT`: `"{bultos_desc}\t{unidades_desc}"`) and Tira 2 (`AW:AY`: `"{bultos_pl}\t{unidades_pl}\t{curva_empaque}"`) for 1-click JS clipboard copy (`render_clipboard_copy_button`) plus `{import_id}_ZDESCARGA_PACKING_LIST_AS_AZ.xlsx`.
4. **Macro-Preserving `.xlsm` Patching (`app.py:1607–1688`, `generar_cierre_oficial.py:43–127`, `tools/operational_stress_test.py:262–292`):**
   - Loads original `.xlsm` via `openpyxl.load_workbook(..., keep_vba=True)`.
   - Patches `CIERRE RECEP` rows `13 + idx` at cols `45` (`AS`), `46` (`AT`), `49` (`AW`), `50` (`AX`), `51` (`AY`).
   - In `app.py` and `generar_cierre_oficial.py` (not in `operational_stress_test.py`), repairs missing/non-formula cells in:
     - Col `H` (`8`): `=+G{r}-J{r}-N{r}-O{r}+L{r}+P{r}+Q{r}`
     - Col `AU` (`47`): `=+AT{r}-H{r}`
     - Col `AZ` (`52`): `=+AX{r}-H{r}`
   - Populates `'Bultos'` sheet rows `3 + idx`, cols `2..10`, with cross-sheet formulas pointing to `'CIERRE RECEP'!C, AS, AT, E, H, BB, BC, AD, AY`.

### 3.6 (f) Persistence, `@st.cache_data` Caching Hazard, `st.session_state` & Offline Contradiction

- **Ephemeral In-Memory Processing:** No database, disk state store, or audit log exists.
- **`st.session_state` Keys:** `audit_data` (`app.py:264`), `app_language` (`src/i18n.py:464`), Vista 1 widget keys (`sel_qc_local`, `sel_cierre_local`, `btn_blind_local`, `btn_local_forced`, `btn_upload_start`, `btn_upload_blind` at `app.py:418,425,435,457,522,546`), and dashboard/component widget keys (`bultos_mode_selector`, `calc_*`, `fb_*`, `f360_*`, `t_*`, `curv_p`, `paste_area_input`, `tsv_*`, `language_selector_widget`, `{key_prefix}_page_size`, `{key_prefix}_current_page`).
- **`@st.cache_data` Stale-Path Hazard (`app.py:267–310, 341–343`):** When Disk Explorer calls `run_audit(qc_path_str, cierre_path_str)` (`app.py:445,468,493`), Streamlit keys the cache on the **path strings**, not file contents or `mtime`. Clicking `Recargar Auditoría` (`app.py:341–343`) only sets `st.session_state.audit_data = None` without calling `st.cache_data.clear()`, so re-auditing a modified Excel file at the same path serves stale cached results.
- **Offline Promise Contradiction:** `README.md:167`, `src/i18n.py:122`, and `app.py:1754–1755,1772` assert 100% local offline execution with zero external requests, yet `app.py:313` unconditionally invokes `st.sidebar.image("https://img.icons8.com/color/96/warehouse.png", width=56)` on every render.

---

## 4. Coupling & Boundary Analysis

### 4.1 Monolithic Presentation + Business Logic Coupling in `app.py` (1,773 lines)

While 6 active modules in `src/` (`models.py`, `config.py`, `sanitizer.py`, `parsers.py`, `engine.py`, `validator.py`; plus dead `heuristic_parser.py`) are Streamlit-free, `app.py` embeds core business logic directly inside UI rendering blocks:
- Recursive disk discovery & candidate classification (`app.py:389–429`)
- Blind-setup gating & warning logic (`app.py:431–455, 544–552`)
- Carton & odd-lot calculator math (`app.py:796–816`)
- Dual-mode `Bultos` projection (Dock Count vs. Strict Customs) & Excel writer (`app.py:835–967`)
- Inline Ficha 360° filtering duplicating `src.engine.filter_items` (`app.py:1053–1069`)
- SAP clipboard delimiter detection & sanitization (`app.py:1413–1425`)
- Block 1 18-column Excel export (`app.py:1455–1494`)
- TSV strip calculation (`app.py:1509–1589`)
- `.xlsm` VBA-preserving patching, formula self-healing (`H`, `AU`, `AZ`), and `'Bultos'` sheet sync (`app.py:1607–1688`)
- Hardcoded dark-mode styling & inline HTML/gradient markup (`app.py:43–258` `<style>` block; `.app-header` at `348–351, 592–595`; `.welcome-card` at `358–363`; sidebar card at `331–335`; blind-setup amber gradient banner at `581–588`; `.status-banner-*` at `1097–1104`; Ficha 360° `.detail-card` grid at `1110–1155`) — primary UI presentation debt targeted by Track C (`REDESIGN_SPECIFICATION.md`).

### 4.2 Quadruplicated `AS:AT / AW:AY` & `.xlsm` Patching Comparison

The Zero-Carton rule and `AS:AT`/`AW:AY` packing-strip computation are duplicated **four times** in production/tooling code (plus a 5th inline copy in `tests/test_bultos_empaque.py:213–231` and a 6th partial projection in `app.py:837–890`), with confirmed behavioral drift:

| Behavior | `app.py` TSV Strips (`1509–1552`) | `app.py` `.xlsm` Writer (`1607–1688`) | `generar_cierre_oficial.py` (`43–127`) | `tools/operational_stress_test.py` (`262–292`) |
|---|---|---|---|---|
| **Output Artifact** | Clipboard TSV + `ZDESCARGA_PACKING_LIST_AS_AZ.xlsx` | In-memory `.xlsm` download (`BytesIO`) | Disk `.xlsm` in `qc_path.parent` | Disk `.xlsm` in `app_output_dir` |
| **Sheet Existence Guard** | N/A | None (`KeyError` if `'CIERRE RECEP'` missing) | None (`KeyError` if `'CIERRE RECEP'` missing) | `if 'CIERRE RECEP' in wb_patch.sheetnames:` (`268`) |
| **Zero-Carton Types (`AS:AT`)** | `int` `0`, `0` (`1518–1519`) | `int` `0`, `0` (`1622–1623`) | `float` `0.0`, `0.0` (`58–59`) | `float` `0.0`, `0.0` (`273–274`) |
| **Zero-Carton Condition (`AW:AY`)** | `is_other_part or (rec == 0 and ped41_u == 0)` (`1529`) | `is_other_p or (rec == 0 and ped41_u == 0)` (`1633`) | `is_other_p or (rec == 0 and ped41_u == 0)` (`70`) | **Zeroes PL whenever `rec == 0`**, even if `ped41_u > 0` (`272`) |
| **Zero-Carton `Empaque` (`AY`) Type** | **`str` `"0"`** (`1532`) | **`int` `0`** (`1636`) | **`str` `"0"`** (`73`) | **`str` `"0"`** (`277`) |
| **Compound `bultos_str` (`*`)** | Checked in non-zero `bultos_desc > 0` branch (`1523`); `bultos_pl` falls back to `bultos_desc` (`1535`), risking `*` string in `AW` | Checked in non-zero `bultos_desc > 0` branch (`1627`); `b_pl` falls back to `b_desc` (`1639`), risking `*` string in `AW` | Checked unconditionally on `ws.cell(..., 45)` (`86`), risking `*` string even on zero-carton rows; `b_pl` uses numeric `bultos_desc` (`76`) | **Ignored** (`bultos_str` never written to `AS` or `AW`) |
| **`qc_details` Fallback** | Reads `qc_item.get('unidades')` & `'bultos'` (`1521,1524–1525`) | Reads `qc_i.get('unidades')` & `'bultos'` (`1625,1628–1629`) | Reads `qc_item.get('unidades')` & `'bultos'` (`61,64–65`) | **Ignores `qc_details`** (`279–280`) |
| **`ped41_prk` Fallback** | Uses `quantity_inv_41_prk` before `unidades / curve` (`1527`) | Uses `quantity_inv_41_prk` before `u_desc / curve` (`1631`) | Uses `quantity_inv_41_prk` before `unidades_desc / curve` (`67`) | Skips `quantity_inv_41_prk`, uses `round(u_val / max(curve_total, 1), 2)` (`280`) |
| **Formula Healing (`H`, `AU`, `AZ`)** | Hardcodes `0` in preview columns `AU`/`AZ` (`1543,1547`) | Repairs `H`, `AU`, `AZ` if missing or non-formula (`1643–1660`) | Repairs `H`, `AU`, `AZ` if missing or non-formula (`80–103`) | **Not implemented** |
| **`'Bultos'` Sheet Sync** | Separate tab (`app.py:831–967`) | Links `'Bultos'` rows `3 + idx` cols `2..10` (`1663–1673`) | Links `'Bultos'` rows `3 + idx` cols `2..10` (`106–116`) | Writes separate `_Hoja_Bultos_Oficial.xlsx` (`240–259`) |
| **Error Handling** | N/A | `except Exception as e_xlsm:` renders muted `st.caption(...)` (`1687–1688`) | Unhandled exception crashes script | `except Exception: pass` swallows failure silently (`291–292`) |

### 4.3 `src/i18n.py` & `src/ui_helpers.py` Coupling + Global DOM JS Injection

- Both `src/i18n.py` (`line 9`) and `src/ui_helpers.py` (`line 6`) import `streamlit` at module top level inside `src/`, coupling the package to a live Streamlit runtime.
- `src/i18n.py:459–517` stores active language in `st.session_state["app_language"]`, renders `st.sidebar.selectbox`, and triggers `st.rerun()`.
- `src/i18n.py:520–620` (`get_table_toolbar_js`) injects a `<script>` via `st.html(..., unsafe_allow_javascript=True)` (`app.py:323`) that reaches into `window.parent.document` (`line 546`), attaches a permanent `MutationObserver` on `doc.body`, adds a capture-phase `mouseover` listener, and attaches a global capture-phase `keydown` listener (`lines 601–615`) suppressing `'c'`/`'C'` outside `INPUT`/`TEXTAREA`/`isContentEditable` to block Streamlit's cache-clear modal.
- `src/ui_helpers.py:191–282` (`render_clipboard_copy_button`) embeds inline HTML/JS via `components.html(..., height=45)`, interpolating `button_id` and `label` directly into JS/HTML template strings.

### 4.4 Dead & Unused Symbols

- `src/heuristic_parser.py` (`72` lines): zero importers across the entire repository.
- `app.py:11` (`import os`), `app.py:14` (`import hashlib`), `app.py:28–29` (`render_skeleton_dashboard`, `render_skeleton_table`), `cli.py:6` (`import os`): imported, unused.
- `app.py:576` (`validation = audit_res["validation"]`): computed on every run, never rendered or read.
- `src/parsers.py:227,247–248`: `col_part` scanned in `CONTROL DATA` row 6, never read when building `qc_details`.
- `src/sanitizer.py:127–138` (`parse_multi_po`): unused outside `tests/test_sanitizer.py`.
- `src/i18n.py`: Of `124` keys in `TRANSLATIONS` (`lines 19–456`), `55` are unused by `app.py` (`app.py` calls `t(...)` on `69` keys); `5` of those remaining keys are used internally in `src/i18n.py` (`language_label` at `line 506`, and `tooltip_fullscreen`, `tooltip_close_fullscreen`, `tooltip_download_csv`, `tooltip_search` at `lines 526–529`), leaving `50` translation keys unused across the entire repository.

---

## 5. Git / Worktree Topology & Non-Destructive Reconciliation Workflow

### 5.1 Verified Git & Worktree State

Observed directly from `git status --short`, `git branch -a -vv`, `git worktree list --porcelain`, `git remote -v`, and `git rev-list --left-right --count`:

| Element | Verified Value |
|---|---|
| **Remote (`origin`)** | `https://github.com/D4RKL0RD-J0571N/3-way-inventory-reconciliation.git` |
| **Shared `.git` Common Dir** | `D:/Development/Software Development/3 Way Inventory Reconciliation/.git` |
| **Main Worktree** | `D:/Development/Software Development/3 Way Inventory Reconciliation` on `operational-benchmark` (`589f6aa`, tracks `origin/operational-benchmark`). Contains gitignored local folders `data/` (71 `IMP-*` dirs), `sandbox_benchmark/` (15 cases + JSON), `dist/`, `build/`, `scratch/`, `ReconcileAI.spec`. |
| **Audit Linked Worktree** | `C:/Users/Jostin/.gemini/antigravity/worktrees/3 Way Inventory Reconciliation/repository_audit_planning` on `repository_audit_planning` (`589f6aa`, no upstream tracking branch set). |
| **`main` vs `operational-benchmark`** | `0  8` (`operational-benchmark` is 8 commits ahead of `main` `c706868`). |
| **`operational-benchmark` vs `repository_audit_planning`** | `0  0` (identical commit `589f6aa`). |
| **`.gitignore` Gap (`.agents/`)** | `.gitignore` (`40` lines) ignores `data/`, `sandbox_benchmark/`, `dist/`, `build/`, `.venv/`, `.pytest_cache/`, `__pycache__/`, `docs/drafts/`, etc., **but does NOT ignore `.agents/`**. |

### 5.2 Non-Destructive Reconciliation Workflow (For Final Delivery Milestone)

To preserve `main`, `operational-benchmark`, and all existing files/histories while excluding `.agents/`:
1. Verify zero tracked modifications (`git diff` is empty) and verify untracked files (`git status --short` shows only `?? .agents/` plus the new `docs/` files).
2. **Never run `git add .` or `git add -A`** (because `.agents/` is unignored in `.gitignore`). Stage **only the 8 explicit planning paths**:
   ```powershell
   git add `
     docs/architecture/CURRENT_ARCHITECTURE.md `
     docs/architecture/TARGET_ARCHITECTURE.md `
     docs/quality/CODE_AUDIT.md `
     docs/quality/KNOWN_LIMITATIONS.md `
     docs/ui/REDESIGN_SPECIFICATION.md `
     docs/sap-knowledge/KNOWLEDGE_EXPANSION_PLAN.md `
     docs/ai/AI_ARCHITECTURE_PROPOSAL.md `
     docs/roadmap/EVOLUTION_ROADMAP.md
   ```
3. Confirm via `git diff --cached --name-only` that **only** those 8 files are staged and `.agents/` remains untracked.
4. Commit on `repository_audit_planning` (`git commit -m "docs: add non-modifying audit, architecture, and evolution planning package"`).
5. Push without merging into `main` or `operational-benchmark`:
   ```powershell
   git push -u origin repository_audit_planning
   ```

---

## 6. Epistemic Separation

### 6.1 Confirmed Findings
- All runtime topologies, `requirements.txt` (`5` packages), `.streamlit/config.toml`, and PyInstaller exclusion of `cli.py`/`generar_cierre_oficial.py` are directly verified from repository files.
- Exact line counts, imports, symbols, missing `src/__init__.py` & `tests/__init__.py`, dead `src/heuristic_parser.py`, unrendered `validation` (`app.py:298,576`), `@st.cache_data` path-key caching (`app.py:267–310,341–343`), external `icons8` image (`app.py:313`), and the 4-way `.xlsm`/TSV patching divergence are directly verified.
- Git worktree layout (`589f6aa` vs `c706868`, `0 8`, `0 0`) and unignored `.agents/` in `.gitignore` are directly verified via `git`.

### 6.2 Assumptions
- `tools/operational_stress_test.py:29` (`I:\My Drive\IMP`) is assumed to represent the author's original Google Drive mount path prior to/alongside the local `data/` directory.
- `docs/CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md`, `docs/AUDITORIA_DESPACHO_MACRO.md`, and `docs/SOP_RECEPCION_SAP.md` are assumed to reflect intended corporate Excel/VBA/SAP procedures, pending Track D (M2) verification.

### 6.3 Unverified Domain / SAP Items
- Whether corporate VBA macros inside `.xlsm` workbooks (e.g., `Mueve43`, `Mueve51`) depend on specific cell data types (`0` int vs `"0"` str in Col `AY`) or trailing rows in `'Bultos'`.
- Exact SAP environment capabilities (ECC 6.0 vs. S/4HANA/EWM) referenced in `docs/SOP_RECEPCION_SAP.md` (investigated in Track D / M2).

### 6.4 Open Questions for the User
1. Are `cli.py` and `generar_cierre_oficial.py` intended to be packaged inside the portable Windows distribution (`tools/build_exe.py`), or are they strictly developer/operator scripts?
2. Should the 8 planning documents be pushed directly on `repository_audit_planning`, or should a separate branch be cut from `repository_audit_planning` before committing?
3. Is strict air-gapped / offline workstation execution required (making the `img.icons8.com` fetch at `app.py:313` a blocker)?
4. Was `validate_all_items()` (`app.py:298,576`) intentionally omitted from UI rendering or left incomplete during demo development?
