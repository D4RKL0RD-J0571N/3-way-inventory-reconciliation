# UI/UX Redesign Specification — ReconcileAI (Track C)

**Specification Date:** 2026-11-17  
**Repository:** `https://github.com/D4RKL0RD-J0571N/3-way-inventory-reconciliation`  
**Baseline Audited Revision:** `589f6aa` (`repository_audit_planning` / `operational-benchmark`)  
**Planning Mode:** Strictly non-modifying UI/UX specification (**R4**, **R8**)  
**Aligned Companion Specs:** `docs/architecture/TARGET_ARCHITECTURE.md` (Track B), `docs/sap-knowledge/KNOWLEDGE_EXPANSION_PLAN.md` (Track D)

---

## 1. Current UI Audit Traceability (Verified Defects & Debt)

Every design decision in this specification traces directly to confirmed defects observed in `.streamlit/config.toml` (`13` lines), `app.py` (`1,773` lines), `src/i18n.py` (`620` lines), and `src/ui_helpers.py` (`283` lines):

| Category | Exact File & Lines | Observed Defect & Impact | Target Remediation |
|---|---|---|---|
| **Dark Neon & Gradient Styling** | `.streamlit/config.toml:7–13`; `app.py:43–258, 331–335, 348–363, 581–588, 1097–1155`; `src/ui_helpers.py:13–91` | Hardcodes dark-mode surfaces (`#0F172A`, `#1E293B`), linear gradients (`button[kind="primary"]`, `.app-header`, `.welcome-card`, blind amber banner), drop shadows, `border-top: 3px solid #0284C7` `.detail-card`s, and dark shimmer skeletons (`SKELETON_CSS`). | Replace with neutral gray ERP surfaces (`#F3F4F6`, `#FFFFFF`, `#EAECEF`), crisp 1px borders (`#D1D5DB`), restrained slate-blue accents (`#1D4ED8`), zero gradients, zero drop shadows. |
| **Pervasive Emoji Usage** | `app.py` (`83` lines, e.g., `L37`, `L332`, `L432–546`, `L717–730`, `L997–1104`, `L1206–1336`, `L1709–1760`); `src/i18n.py` (`79` lines / `82` strings: `2` in `LANGUAGES` `L14–15` + `77` lines across `40` `TRANSLATIONS` keys / `80` strings `L26–415`); `src/ui_helpers.py:240` | Embeds emojis in window icon, tab labels, status badges, KPI headers, language selector, buttons, and clipboard toast (`U+2705 "¡Copiado!"`), clashing with established corporate ERP ergonomics (**R4**). | **ZERO emojis anywhere** in window chrome, tabs, buttons, status pills, dialogs, or `TRANSLATIONS`. Use crisp text labels and bundled 16×16 monochrome SVGs. |
| **External CDN Dependency** | `app.py:313` | `st.sidebar.image("https://img.icons8.com/color/96/warehouse.png", width=56)` executes an HTTPS fetch on every render, breaking air-gapped warehouse operation and contradicting `README.md:167` / `src/i18n.py:122`. | Bundle local 16×16 monochrome SVG icons inside `ui_desktop/assets/icons/` (0 external network calls). |
| **Unrendered Validation Engine** | `app.py:298, 305, 576` (`A10-1`) | `validate_all_items(items)` computes all 6 domain/SAP integrity checks on every audit and unpacks `validation = audit_res["validation"]`, then **never displays any issue** in the UI. | Add clickable Status Bar / KPI Summary Badge (`Errors: N | Warnings: N | Info: N`), dedicated **Validation Alerts Panel** tab, and inline item alerts in `Inspector 360` & `Reconcile Grid`. |
| **Stale Path-Key Cache & False Streaming Claim** | `app.py:267–310, 332–335, 341–343` | Disk Explorer keys `@st.cache_data` on path strings (`str(path)`), not bytes/`mtime`; `reload_audit_btn` clears only `session_state`, serving stale data after Excel edits. Sidebar badge claims active streaming while `parse_cierre_recepcion` loads non-`read_only` (`3.1s–4.5s`). | Replace with `AuditTrailService` SHA-256 + `mtime_ns` + size fingerprinting and `QThreadPool` async execution over `read_only=True` streaming (`IngestionPipeline`). |
| **`folder_input` Reset & Discovery Collisions** | `app.py:377–429, 511–552` (`A10-2`, `A8-7`, `A10-10`) | `folder_input` lacks state binding and resets on `st.rerun()`; single-file folders auto-select the same file for QC and Cierre (`409–410`); missing `"calidad"` keyword (`IMP-81`) forces false Blind Setup; `.xls` is accepted in Explorer (`395`) yet fails `openpyxl` and is blocked in Uploader (`514, 518`). | Route all discovery through `DiscoveryService` (sheet-signature detection, single-file disambiguation, explicit `.xls` BIFF status notice, persistent bound path state). |
| **Pagination & Clipboard JS Hacks** | `src/ui_helpers.py:140–282` (`A10-3`); `src/i18n.py:520–620` (`A10-6`) | `paginate_dataframe` ignores `default_page_size=100` (`index=1` hardcode) and throws `StreamlitValueAboveMaxError` when filter shrinks `total_pages`. `render_clipboard_copy_button` injects a 45px sandboxed iframe using deprecated `execCommand`. `get_table_toolbar_js` attaches a `MutationObserver` on `window.parent.document` and blocks `'c'`/`'C'` keystrokes globally. | Replace with native `QTableView` + `QSortFilterProxyModel` virtualization (no pagination crash), native `QApplication.clipboard()` TSV copy/paste, and zero DOM/key hijacking. |
| **Numeric & TSV Formatting Corruption** | `app.py:815, 824, 852–890, 1139, 1168, 1316–1322, 1349, 1387, 1463–1483, 1543–1551` (`A5-1`..`A5-5`, `A3-3`) | TSV strips emit `.0` on counts (`65.0\t65.0`, `12.0\t888.0`); `df_exp` outputs IEEE-754 floats (`3.1666666666666665`); detail cards/tables truncate unit costs via `:.2f`; `int()` floor-truncates floats; `subtab_descarga` hardcodes `AU=0`/`AZ=0`. | Route 100% of formatting through `FormattingService`: whole counts `75` / TSV `"65\t65"`, adaptive unit costs `$2.13` / `$3.1667`, 2-decimal totals `$550.08`, real `AU = AT - H` / `AZ = AX - H`. |
| **Tab-Specific & Compatibility Bugs** | `app.py:1080, 1090, 1368, 1394–1431, 1608–1688, 1695–1773` (`A6-3`, `A8-5`, `A8-6`, `A10-4`, `A10-7`, `A10-8`) | `opt_labels` collides on multi-PO SKUs; `app.py:1090` fails with `SyntaxError` on Python `<=3.11`; `tab_curves` uses `next(...)` without Part filter; `tab_clipboard` is a 15-row dead end breaking on SAP `\|---\|` pipes; `.xlsm` patcher writes `13 + idx` without clearing trailing `'Bultos'` rows and hides errors in `st.caption`; `tab_checklist` defaults 7 unkeyed checkboxes to `True`; `tab_wiki` contradicts code reality. | Disambiguate items by `(material_code, order_po_43, source_row)`; remove PEP 701 quotes; upgrade `tab_clipboard` to `SapClipboardAdapter`; anchor `.xlsm` writes to `source_row`; persist `False`-default checklist via `AuditTrailService`; replace `Wiki` with `SapKnowledgeService`. |

---

## 2. Target UI Stack Alignment (100% Aligned with `TARGET_ARCHITECTURE.md`)

Per `docs/architecture/TARGET_ARCHITECTURE.md §3`:
- **Primary Target Stack:** **PySide6 (Qt for Python)** native Windows desktop shell (`ui_desktop/`) using:
  - `QMainWindow` with native `QMenuBar`, `QToolBar`, `QStatusBar`, `QSplitter`, and collapsible `QDockWidget` (right-docked **SAP Knowledge / Analyst Companion Drawer**).
  - `QTableView` + `QAbstractTableModel` + `QSortFilterProxyModel` for high-density sorting, filtering, numeric right-alignment, and `Ctrl+C` / `Ctrl+V` TSV ergonomics via `QApplication.clipboard()`.
  - `QThreadPool` + `QRunnable` background workers for `DiscoveryService`, `IngestionPipeline`, `ReconciliationService`, and `ExportService` so Excel parsing never blocks the UI thread.
  - Packaged offline via `PyInstaller --onedir --windowed` (dynamically linking `LGPLv3` Qt DLLs, zero localhost server, zero browser popup, zero CDN calls).
- **Strict Service-Layer Boundary:** UI widgets contain **zero** business math, Excel parsing, `.xlsm` patching, or file discovery logic. Every screen calls only the 8 shared application services in `src/services/` (`DiscoveryService`, `IngestionPipeline`, `ReconciliationService`, `ValidationService`, `ExportService`, `FormattingService`, `AuditTrailService`, `SapKnowledgeService`).
- **Documented Zero-Dependency / Browser Fallback:** If user confirmation (`Open Questions`) requires keeping `requirements.txt` strictly within the existing 5 packages or hosting via intranet browser, the exact same visual design system (`Section 3`), screen layout (`Section 4`), `FormattingService` rules, and `src/services/` calls apply identically to a **Refactored Neutral-Gray Streamlit** adapter (replacing `QMenuBar`/`QDockWidget` with compact horizontal radio/nav strips and a collapsible right sidebar/expander, while stripping all emojis, gradients, `icons8.com`, and `MutationObserver`/`'c'` JS hacks). For TSV clipboard strips on corporate intranet HTTP (`window.isSecureContext === false`, where `navigator.clipboard.writeText` is blocked), replace hidden-textarea `execCommand` iframe hacks with a read-only monospaced `st.code(tsv, language=None)` block (native Streamlit copy icon) paired with a visible auto-select read-only `<textarea readonly onclick="this.select()">` / `st.text_area` fallback and `"HTTP LAN: click strip and press Ctrl+C"` caption.

---

## 3. Visual & Ergonomic Design System (Strict R4 Compliance)

### 3.1 Hard Prohibitions
1. **ZERO emojis** in tabs, status pills, menus, toolbar buttons, KPI cards, tooltips, dialogs, or `src/i18n.py` strings.
2. **ZERO futuristic or generic AI dashboard aesthetics** (no glowing neon cyan borders, no dark `#0F172A` surfaces, no conversational chat bubbles, no animated pulse indicators).
3. **ZERO CSS/QSS linear gradients** (`linear-gradient` / `qlineargradient` prohibited on buttons, headers, banners, and cards).
4. **ZERO oversized KPI cards or oversized icons** (compact single-row summary strip replaces stacked hero boxes; all toolbar/menu/tab icons fixed at `16×16` px monochrome SVG).

### 3.2 Design Tokens (`QPalette` + QSS / Fallback CSS Variables)

| Token Name | Hex / Value | Usage |
|---|---|---|
| `bg.canvas` | `#F3F4F6` | Main window background, tab pane canvas, splitter handle |
| `bg.surface` | `#FFFFFF` | Data tables (`QTableView`), property grids, input fields, Companion Drawer body |
| `bg.chrome` | `#EAECEF` | `QMenuBar`, `QToolBar`, table header row, `QStatusBar`, group headers |
| `bg.row_alt` | `#F8FAFC` | Subtle alternating zebra stripe in dense grids (`#E5E7EB` hover highlight) |
| `border.hairline` | `#D1D5DB` | Crisp `1px solid` borders around all panels, tables, toolbar separators, pills |
| `border.strong` | `#9CA3AF` | Active splitter / table header bottom divider |
| `border.focus` | `#1D4ED8` | Explicit keyboard `:focus` ring (`border: 1px solid #1D4ED8`) on inputs, tabs, buttons, and table views |
| `accent.primary` | `#1D4ED8` | Primary CTA buttons (`Run Audit`, `Export .xlsm`), active tab underline (`2px`) |
| `accent.primary_hover` | `#1E40AF` | Hover state for primary buttons (flat color shift, zero shadow/transform) |
| `accent.selection` | `#EFF6FF` | Selected table row / active list item background (`#1E3A8A` text) |
| `text.primary` | `#1E293B` | Primary labels, table cell values, headings |
| `text.secondary` | `#4B5563` | Field captions, secondary metadata, inactive tab text |
| `text.muted` | `#6B7280` | Placeholders, disabled controls, timestamp footers |

### 3.3 Restrained 1px-Bordered Semantic Status Pills (Zero Emojis)

All status badges render as compact `20px`-high pills (`border-radius: 3px`, `1px solid` border, `11px` semi-bold uppercase/title text):

| Domain / Epistemic Status | Background | 1px Border | Text Color | Exact Label (`EN` / `ES`) |
|---|---|---|---|---|
| `LIMPIO` / `Confirmed` | `#ECFDF5` | `#A7F3D0` | `#065F46` | `CLEAN` / `LIMPIO` (`Confirmed` / `Confirmado`) |
| `PENDIENTE_PARTE_X` / `Generic SAP` | `#EFF6FF` | `#BFDBFE` | `#1E40AF` | `PENDING PART B` / `PENDIENTE PARTE B` (`Generic SAP ECC 6.0`) |
| `PAUSA_SIN_CONTROL_CALIDAD` / `WARNING` / `Validate in Env` | `#FFFBEB` | `#FDE68A` | `#92400E` | `PAUSED NO QC` / `PAUSA SIN QC` (`Validate in SAP ECC 6.0`) |
| `FALTANTE` / `DAÑADO` / `ERROR` / `Conflicting` | `#FEF2F2` | `#FECACA` | `#991B1B` | `SHORTAGE` / `FALTANTE` (`Error` / `Conflicto`) |
| `SOBRANTE` / `REVUELTO` / `Company-Specific` | `#F5F3FF` | `#DDD6FE` | `#5B21B6` | `EXCESS` / `SOBRANTE` (`Company-Specific: Ekono`) |
| `INFO` / `Neutral` | `#F3F4F6` | `#D1D5DB` | `#374151` | `INFO` / `INFORMATIVO` |

### 3.4 Typography, Iconography & Spacing Scale
- **UI Font:** `Segoe UI` (Windows native), `12px` body, `11px` table headers & status bar, `13px` section headers (`600` weight), `16px` window/case title (`600` weight).
- **Tabular Numerals (`tnum`) & Codes:** `Consolas` / `Cascadia Mono` `12px` for `MATNR`, `PO 43`, `Doc 41`, TSV strips, and right-aligned numeric table columns so decimals and unit digits align vertically.
- **Iconography:** Bundled `16×16` monochrome stroke SVGs (`#374151` default, `#1D4ED8` active, `#9CA3AF` disabled) stored locally in `ui_desktop/assets/icons/` (e.g., `folder-open.svg`, `play-audit.svg`, `refresh.svg`, `alert-triangle.svg`, `book-open.svg`, `copy.svg`, `download.svg`, `check-square.svg`). Zero external URLs.
- **Spacing Scale & High-DPI Row Sizing:** `4px` (cell/badge padding), `8px` (toolbar/grid gap), `12px` (pane/section padding). Minimum table row height is derived dynamically via `QFontMetrics` as `max(24px, lineHeight + 8px)` (yielding ~`24px–26px` compact ERP density at 100% scale without clipping `Segoe UI` `12px` at Windows 125%–150% DPI scaling). All interactive controls and grid views enforce an explicit keyboard focus ring (`border: 1px solid #1D4ED8`, `border.focus`).

### 3.5 Numeric & Precision Formatting Contract (`FormattingService`)
Every table cell, detail property, KPI metric, `.xlsx` export, and TSV strip **must** format values via `FormattingService`:
1. **Whole-Number Counts (`UN`, `PRK`, `Bultos`, `Curve`, `Child Ratios`):**
   - **UI Tables & KPIs:** Exact/near-integers (`abs(v - round(v)) < 1e-6`) render with thousands grouping and **zero decimals** (`75`, `0`, `1,152` — **NEVER `75.0`**). Genuinely fractional quantities (e.g., calculator `10 / 6 = 1.67 PRK`, fixing `A5-5`) display up to `2–4` trimmed decimals rather than silently `int()`-flooring (`A5-2`), flagging an integrality warning when discrete units are expected.
   - **TSV Clipboard Strips (`AS:AT` & `AW:AY`):** Plain ungrouped integer strings (`"65\t65"`, `"12\t888"`, `"0\t0\t0"` — **NEVER `"65.0\t65.0"`**, fixing `A5-1`), preserving compound packaging expressions (`"39*60 1*78"`) verbatim after normalization.
2. **Prorated / Monetary Unit Costs (`unit_cost`, `child.unit_cost`):**
   - Backed by `Decimal` (`ROUND_HALF_UP`). Displays a minimum of `2` decimals (`$2.13`) and preserves up to `4–6` non-zero decimals whenever sub-cent precision exists (`$3.1667`, `$1.380833`, `$4.79902`), eliminating both IEEE-754 noise (`3.1666666666666665`, `A5-3`) and lossy `:.2f` truncation (`A5-4`).
3. **Extended Monetary Totals (`ncr_amount`, `total_invoice_value`):**
   - Computed from full-precision `Decimal` `unit_cost * qty`, then formatted to exact **2 currency decimals** (`$550.08`, `$869.76`, `$0.00`).

---

## 4. Screen-by-Screen & Component Specification

### 4.1 Persistent Application Shell (`QMainWindow`)

```text
+---------------------------------------------------------------------------------------------------+
| File   Audit   Export   Validation   Knowledge   Help                                    [_][O][X]|  <- QMenuBar
+---------------------------------------------------------------------------------------------------+
| [Open Case] [Run Audit] [Reload] | [Export .xlsm] [Copy AS:AT] [Copy AW:AY] | [Alerts(3)] [KB] [ES|EN] |  <- QToolBar (16px icons)
+---------------------------------------------------------------------------------------------------+
| CASE: IMP-207 | DOC 41: 4100001234 | SUPPLIER: ... | STATUS: [CON_DISCREPANCIAS]                  |  <- Compact Case Header
| Styles: 35 (A:35) | Invoiced: 12,480 | Received: 12,156 | Cartons: 182/184 | Mov101-102: 12,156 | NCR: $869.76 | Alerts: E:1 W:2 I:3 |
+-----------------------------------------------------------------------------------+---------------+
| [View 1: Case Loader (Ctrl+O)] | [1 Bultos] [2 Inspector 360] [3 Reconcile Grid]  | SAP COMPANION |
| [4 Curves] [5 SAP Clipboard] [6 Official Export] [7 Checklist] [8 Alerts(6)] [9 SOP]| [x]           |
+-----------------------------------------------------------------------------------+---------------+
|                                                                                   | Context:      |
|                                                                                   | MAT 5033284   |
|                          ACTIVE WORKSPACE TAB CONTENT                             | [Generic ECC] |
|                    (QSplitter / QTableView / Master-Detail)                       | [Validate Env]|
|                                                                                   |               |
|                                                                                   | T-Code / Tbl  |
|                                                                                   | Movements     |
|                                                                                   | Triage Steps  |
|                                                                                   | Citations     |
+-----------------------------------------------------------------------------------+---------------+
| Ready | Local Offline (0 network calls) | Parse: 0.42s | Cache: SHA256 Verified | E:1 W:2 I:3     |  <- QStatusBar
+---------------------------------------------------------------------------------------------------+
```

- **`QMenuBar` Structure:**
  - `File`: `Open Import Folder... (Ctrl+O)`, `Load Individual Workbooks...`, `Recent Cases`, `Exit (Alt+F4)`.
  - `Audit`: `Run / Re-run Audit (F5)`, `Clear Hash Cache & Reload (Ctrl+F5)`, `Toggle Blind Setup Mode`.
  - `Export`: `Export Block 1 (.xlsx)`, `Export Bultos (.xlsx)`, `Patch Official Closure (.xlsm) (Ctrl+E)`, `Copy Tira 1 AS:AT`, `Copy Tira 2 AW:AY`.
  - `Validation`: `Open Validation Alerts Panel (F8)`, `Jump to Next Error (F4)`, `Export Alerts (.xlsx)`.
  - `Knowledge`: `Toggle SAP Companion Drawer (F1)`, `Open SOP & Reference Tab (Alt+9)`, `SAP ECC 6.0 Verification Checklist`.
  - `Help`: `Keyboard Shortcuts`, `About ReconcileAI (Offline / Version / Manifest)`.
- **Compact Single-Row KPI Summary Strip (replaces 6 oversized `st.metric` cards):**
  - Displays 7 dense inline readouts: (1) Import ID & Doc `41`, (2) Audited Styles + `container_parts_breakdown`, (3) Invoiced Units (`FormattingService.format_count`), (4) Received Physical Units (`0 (Pending QC)` if Blind Setup), (5) Cartons `Dock / PL` (`dif_total_bultos`), (6) **ERP Balance (`Mov 101 - Mov 102`)** (wiring previously unused `kpi_erp_balance`), and (7) Credit Note Total (`$0.00 (Locked)` if Blind Setup). Plus clickable **Validation Summary Pill** (`E: N | W: N | I: N`).
  - When `summary.is_blind_setup == True`, displays a flat 1px-bordered amber notice bar (`#FFFBEB` / `#FDE68A`): `"PAUSA_SIN_CONTROL_CALIDAD: QC missing or bypassed. Received quantities and NCR totals are locked at 0 until Control de Calidad is loaded."`

---

### 4.2 View 1 — Case Loader / Start Workspace

Replaces `app.py:346–568` (`welcome-card` hero + `tab_local_explorer` + `tab_uploader`):

```text
+---------------------------------------------------------------------------------------------------+
| CASE LOADER  (Ctrl+O)                                                                             |
+-----------------------------------------------+---------------------------------------------------+
| 1. LOCAL CASE EXPLORER (data/)                | 2. MANUAL FILE SELECTION / DRAG-AND-DROP          |
| Root Folder: [ D:\...\data\IMP-81  ] [Browse] |                                                   |
| Quick Case:  [ IMP-81              v] [Load]  | QC Workbook (.xlsm/.xlsx):                        |
|                                               |  [ INFORMACION GENERAL\IMP 81.xlsm    ] [Browse]  |
| Detected Workbooks (Sheet-Signature Scan):    |  Signature: CONTROL DATA, BASE DE DATOS (Valid)   |
| +-------------------------------------------+ |                                                   |
| | File                        | Role        | | Closure Workbook (.xlsm/.xlsx):                   |
| | IMP-81.xlsm                 | CIERRE [OK] | |  [ IMP-81.xlsm                        ] [Browse]  |
| | INFORMACION GENERAL/IMP 81  | QC BOM [OK] | |  Signature: CIERRE RECEP, Bultos (Valid)          |
| | Factura/INV.xls             | BIFF .xls(!)| |                                                   |
| +-------------------------------------------+ | [x] Blind Setup Pause Mode (No QC Available)      |
|                                               |     Warning: NCR & physical units locked to 0.    |
| Non-Excel Corpus Summary:                     |                                                   |
| 2 PDF | 1 .eml | 1 .xls (Requires [xls] extra)| [ Run 3-Way Audit (F5) ]   [ Run Blind Audit ]    |
+-----------------------------------------------+---------------------------------------------------+
```

- **Service Calls:** `DiscoveryService.list_import_folders()`, `DiscoveryService.discover_case_files()`, `IngestionPipeline.ingest_case()`, `AuditTrailService.get_cached_or_compute()`, `ReconciliationService.run_reconciliation()`, `ValidationService.validate_run()`.
- **Key Fixes Enforced:**
  - `Root Folder` / `Quick Case` binds persistently to controller state (`A10-2` fixed).
  - Sheet-signature inspection recognizes `INFORMACION GENERAL/IMP 81.xlsm` as QC even without `"calidad"` in the filename (`IMP-81` false Blind Setup fixed).
  - Single-file folders never auto-assign the same workbook to both QC and Closure (`A10-10` fixed).
  - Legacy BIFF `.xls` files (`62` in `data/`) display an explicit status row (`BIFF .xls — requires [xls] xlrd adapter or Save As .xlsx`, fixing `A8-7`) rather than silently failing.
  - Blind Setup requires explicit confirmation when a valid Closure exists without QC.

---

### 4.3 View 2 — Reconciliation Workspace Tabs & Panels (All 8 Existing + 2 Mandatory New Surfaces)

#### Tab 1: `Bultos` (Cartons & Packaging)
- **Layout:** Top bar with **Projection Mode** radio (`Dock Count (Physical)` vs `Strict Customs (Full + Odd Lots)`), 5 compact KPI pills (`cb1`..`cb5`, fixing `cb4` `app.py:770` so it never divides `total_received_units / total_prk_41`), collapsible **Carton & Odd-Lot Calculator** (`units`, `curve`, `std_per_box`, `has_odds`, `odd_units` outputting exact decimal PRKs `1.67 PRK` instead of `,.0f`, fixing `A5-5`), Category & Search filters, dense `QTableView`, and `[Export Bultos .xlsx]` button.
- **Service Calls:** `ReconciliationService.project_bultos(mode)`, `ReconciliationService.calculate_packaging(...)`, `FormattingService`, `ExportService.export_bultos_xlsx()`.
- **Key Fixes:** Single canonical Zero-Carton rule shared with `ExportService` (`A2-5`), compound `raw_ay` (`"9*72 1*90"`) & `+<units>` (`"8*72+24"`) display (`A2-1`, `A2-3`), whole-number carton/unit formatting (`75`, not `75.0`).

#### Tab 2: `Inspector 360` (Material Master-Detail)
```text
+---------------------------------------------------------------------------------------------------+
| Filter: Part [ALL v] Status [ALL v] Cat [ALL v] Search [          ]  Item: [5033284 | PO 43..6709 v]|
+---------------------------------------------------------------------------------------------------+
| [SHORTAGE]  MAT: 5033284 | Vendor Style: ST-991 | PO 43: 4300006709 | Source Row: 18 | Alerts: W:1|
+-----------------------------------------------+---------------------------------------------------+
| 1. Identification & Logistics                 | 2. Physical & ERP Quantities (UN / PRK)           |
|   Material (MATNR)    | 5033284               |   PO 43 Ordered       | 90 UN  (15 PRK)           |
|   Vendor Style (IDNLF)| ST-991                |   Invoice 41 Billed   | 90 UN  (15 PRK)           |
|   Category / Part     | Damas / Part A        |   Dock Received (H)   | 54 UN                     |
|   SAP Mov 101-102     | 54 UN  [Match]        |   Short/Dmg/Exc/Rev   | 36 / 0 / 0 / 0            |
+-----------------------------------------------+---------------------------------------------------+
| 3. Valuation & NCR Impact                     | 4. Packaging & Cartons (AS:AT vs AW:AY)           |
|   Unit Cost (Doc 41)  | $2.13                 |   Dock Cartons (AS:AT)| 3 Bultos / 54 UN          |
|   Invoiced Total      | $191.70               |   Packing List (AW:AY)| 5 Bultos / 90 UN          |
|   Claimable NCR       | $76.68                |   Empaque Breakdown   | 3*18                      |
+-----------------------------------------------+---------------------------------------------------+
| Inline Validation Alerts: [WARNING] PRECIO_CERO_EN_RECLAMO / ...  [View in Alerts] [Open SAP KB]  |
+---------------------------------------------------------------------------------------------------+
| Child SKU Breakdown (Curve Sum: 6 == 6 [OK])                                                      |
| +-----------------------------------------------------------------------------------------------+ |
| | Child MATNR | Description        | Ratio | Theo UN | Rec UN | Diff UN | Unit Cost | Claim $   | |
| | 5033284001  | Talla S            |     2 |      30 |     18 |     -12 |     $2.13 |    $25.56 | |
| +-----------------------------------------------------------------------------------------------+ |
+---------------------------------------------------------------------------------------------------+
```
- **Service Calls:** `ReconciliationService.filter_items()`, `ValidationService.issues_for_item()`, `FormattingService`, `SapKnowledgeService.lookup_for_issue()`.
- **Key Fixes:** Selector keyed by unique tuple `(material_code, order_po_43, source_row)` (`A6-3`); zero PEP 701 f-string quotes (`A10-8`); adaptive unit cost (`$3.1667` / `$2.13`) + exact 2-decimal totals (`A5-4`); non-zeroed partial `ChildSKU.received_units` (`A1-1`); inline validation alerts + direct Companion jump.

#### Tab 3: `Reconcile Grid` (`Semáforo`)
- **Layout:** Filter toolbar (`Part`, `Status`, `Category`, `Search`, `Only With Alerts` toggle) + **16-column virtualized `QTableView`** (the existing 15 columns of `df_t` plus a compact **`Alerts`** pill column `E/W/I` — note `df_exp` remains 18 columns in `Official Export`).
- **Interactions:** Double-clicking any row (or pressing `Enter` / clicking `[Inspect Material]`) switches directly to `Inspector 360` bound to that exact `(material_code, order_po_43, source_row)` and updates the **SAP Companion Drawer** context. `Ctrl+C` copies selected cells/rows as clean TSV (`75`, `$3.1667`).
- **Service Calls:** `ReconciliationService.filter_items()`, `ValidationService`, `FormattingService`.
- **Key Fixes:** Eliminates `paginate_dataframe` crashes (`A10-3`), `int()` truncation (`A5-2`), and `$%.2f` cost rounding (`A5-4`).

#### Tab 4: `Curves` (BOM & Size-Curve Breakdown)
- **Layout:** Added **Container Part** & **Status** filters alongside a PO-disambiguated selector `(material_code, order_po_43, source_row)` (replacing `next(...)` at `app.py:1368`, fixing `A6-3`). Displays 4 compact KPI readouts (`Curve Total`, `Ratio Sum`, `Theo UN`, `Rec UN`), an inline `CURVA_DESCUADRADA` / `BOM_FALTANTE` banner if flagged by `ValidationService`, and the Child SKU `QTableView` formatted via `FormattingService`.
- **Service Calls:** `ReconciliationService.filter_items()`, `ValidationService`, `FormattingService`.

#### Tab 5: `SAP Clipboard Sanitizer`
- **Layout:** Multi-line paste input (`Ctrl+V` directly from SAP GUI ALV lists `ME80FN`, `MB51`, `MB52`, `SE16N`), target preset selector (`Auto-Detect`, `MB51 Movements`, `ME80FN POs`, `SE16N Table`), full preview `QTableView` (removing the 15-row `head(15)` limit of `A10-4`), and action bar: `[Clean & Parse]`, `[Copy Clean TSV]`, `[Export .xlsx]`, `[Cross-Check vs Loaded Audit]`.
- **Service Calls:** `IngestionPipeline.parse_sap_clipboard()` (`SapClipboardAdapter`), `FormattingService`, `ExportService`.
- **Key Fixes:** Strips SAP ALV `|---|` borders and leading/trailing `|` columns, preserves alphanumeric & leading-zero `MATNR`s (`A7-1`), turns dead-end `tab_clipboard` into a working analyst utility.

#### Tab 6: `Official Export & Paste Strips`
- **Subtab A — Core Table (`Block 1`, 18 columns):** Fixes the `"17 columnas"` caption drift (`app.py:1453`), virtualizes all rows in `QTableView`, formats unit costs cleanly via `FormattingService` (`$3.1667` / `Decimal`-normalized `.xlsx` export, eliminating `3.1666666666666665` `A5-3`), and exports `{import_id}_Cierre_Oficial_Datos.xlsx`.
- **Subtab B — Official `.xlsm` Patcher & `ZDESCARGA` Paste Strips (`Block 2`):**
  - Displays full `descarga_rows` preview table with **real calculated variances** `AU = AT - H` and `AZ = AX - H` (fixing hardcoded `0`s `A3-3`).
  - **Tira 1 (`AS:AT`)** & **Tira 2 (`AW:AY`)** monospaced read-only strip boxes with native `[Copy Tira 1 TSV]` / `[Copy Tira 2 TSV]` buttons using `QApplication.clipboard()` (zero iframe JS hacks; Refactored Streamlit fallback uses `navigator.clipboard.writeText` when `isSecureContext=true`, or read-only `st.code` / auto-select `<textarea readonly>` on HTTP LAN) and **exact integer counts** (`"65\t65"`, `"12\t888"` — NEVER `"65.0\t65.0"`, fixing `A5-1`).
  - `[Download ZDESCARGA AS:AZ .xlsx]` and `[Patch Official .xlsm (Save As...)]` (anchored strictly to `item.source_row`, healing `H`/`AU`/`AZ`, clearing trailing `'Bultos'` template rows, and surfacing any I/O/sheet error in a modal `QMessageBox.critical` + `AuditTrailService` log rather than muted `st.caption`, fixing `A8-5`, `A8-6`).
- **Service Calls:** `ExportService.build_core_table()`, `ExportService.build_paste_strips()`, `ExportService.patch_official_xlsm()`, `FormattingService`, `AuditTrailService`.

#### Tab 7: `Binder Checklist` (7-Document Physical Closing Sign-Off)
- **Layout:** 7-row structured table/checklist matching the physical binder (`SOP_RECEPCION_SAP.md:112–120`): step description, automated pre-check status badge from `ReconciliationService`/`ValidationService` (e.g., Step 4 `PARA N.C RECEP.` links to `total_ncr_amount` & `PRECIO_CERO_EN_RECLAMO`), manual analyst sign-off checkbox (**default `False`**, never pre-checked `True`, fixing `A10-7`), operator note, and timestamp.
- **Service Calls:** `AuditTrailService.load_checklist_state(import_id)` / `save_checklist_state(...)` (persisted per `import_id` across restarts and language switches).

#### Tab 8 (NEW): `Validation Alerts Panel` (`Alt+8` / `F8` / Clickable Status Badge)
```text
+---------------------------------------------------------------------------------------------------+
| VALIDATION ALERTS  |  Summary: [ERROR: 1] [WARNING: 2] [INFO: 3]                                  |
| Severity: [ALL v]  Code: [ALL v]  Search: [              ]   [Export Alerts .xlsx]                |
+---------------------------------------------------------------------------------------------------+
| Sev     | Code                   | Material | PO 43      | Message & Remediation                  |
|---------+------------------------+----------+------------+----------------------------------------|
| [ERROR] | CURVA_DESCUADRADA      | 1455103  | 4300006709 | Ratio sum (11) != Curve (12). Check    |
|         |                        |          |            | tables MAST/STPO via SE16N.            |
|         |                        |          |            | [Jump: Inspector] [Ask Companion (KB)] |
| [WARN]  | PRECIO_CERO_EN_RECLAMO | 5033299  | 4300006710 | Shortage 12 UN with $0.00 cost. Check  |
|         |                        |          |            | ME80FN (never VKP2). [Jump] [KB]       |
+---------------------------------------------------------------------------------------------------+
```
- **Surfaces 100% of `ValidationService` (`src/validator.py`) Issues (fixing `A10-1`):**
  1. `CURVA_DESCUADRADA` (`ERROR`) $\rightarrow$ jumps to `Curves` / `Inspector 360`; links to `sap.table.mast_stpo` (correcting `validator.py:33`'s erroneous `"transacción STPO/MAST"` text to `"tabla MAST/STPO"`).
  2. `POSIBLE_CONFUSION_UDM` (`ERROR`) $\rightarrow$ bidirectional `PRK`/`UN` check (`81` vs `1,728` & `>2.5x`); jumps to `Inspector 360`; links to `sap.troubleshooting.prk_vs_un`.
  3. `BOM_FALTANTE` (`WARNING`) $\rightarrow$ excludes `um == "UN"` false positives; jumps to `Inspector 360`; links to `sap.table.stpo`.
  4. `PRECIO_CERO_EN_RECLAMO` (`WARNING`) $\rightarrow$ checks `falt`/`dan`/unannotated `ped41 > rec`; jumps to `Inspector 360` Card 3; links to `sap.tcode.me80fn` (explicit `VKP2` exclusion warning).
  5. `ESTILO_PROVEEDOR_VACIO` (`INFO`) $\rightarrow$ jumps to `Inspector 360` Card 1; links to `sap.table.eina`.
  6. `MATERIAL_EN_MULTIPLES_PO` (`INFO`) $\rightarrow$ filters `Reconcile Grid` by `material_code`; links to `sap.troubleshooting.multi_po`.
  7. **Planned Phase 1/2 Checks (`[Planned]`):** `SAP_MOV_101_DESCUADRE`, `DESCARGA_VS_PACKING_DESCUADRE` (`AU`/`AZ`), `EMPAQUE_INVALIDO` (`AY` vs `AT`/`H`), `DISPATCH_MOV_641_DESCUADRE`.

#### Tab 9: `SOP & Reference` (`Alt+9`, Replaces Contradictory `tab_wiki`)
- Backed exclusively by `SapKnowledgeService` (`docs/sap-knowledge/index.json`) rather than hardcoded static markdown (`app.py:1714–1773`), eliminating version drift (`v1.0.0`/`v1.1.0`/`v1.2.0`) and S/4HANA/EWM vs ECC 6.0 contradictions while documenting exact workbook contracts (`CIERRE RECEP`, `CONTROL DATA`, `Bultos`).
- **Pre-Priority-#4 Missing/Empty/Corrupt Index Degraded State:** On tab activation, inspects `SapKnowledgeService.get_index_status()`. Whenever `state != "ready"` (`not_built`, `empty`, or `corrupt`), renders a persistent, non-dismissible 1px-bordered amber banner (`#FFFBEB` / `#FDE68A` / `#92400E` per `§3.3`): **`"Knowledge Index Not Built (Roadmap #4 Pending)"`** (localized via `kb_index_not_built` / `kb_index_corrupt`). Populates the reference pane non-fatally from **built-in `ValidationService` remediation metadata** and the **verified `SOP_RECEPCION_SAP` operational summary notice** (7-document binder checklist, `ME80FN`/`SE16N`/`MB51` queries, multi-warehouse `TOTAL DE UNIDADES` rule — explicitly flagging `docs/SOP_RECEPCION_SAP.md:1,30-31` `S/4HANA`/`EWM`/`VL31N`/`RFUI` claims as unverified against `ECC 6.0` `6.00.29`; **never** reverting to `app.py:1714–1773` `tab_wiki` prose). All reconciliation, validation, clipboard, and export workflows remain 100% functional.

#### Dock Panel (`F1`, NEW): `SAP Knowledge / Analyst Companion Drawer` (`QDockWidget`, Right Dock, `360px`)
A **non-chatbot, deterministic operational lookup drawer** (`KNOWLEDGE_EXPANSION_PLAN.md §5.2–5.3`) pinned to the analyst's active tab, selected `PrepackItem`, or clicked `ValidationIssue`:
- **Pre-Priority-#4 Bootstrap Behavior:** Checks `SapKnowledgeService.get_index_status()` on activation. When `state != "ready"`, displays the same non-dismissible amber banner **`"Knowledge Index Not Built (Roadmap #4 Pending)"`** (`#FFFBEB` / `#FDE68A` / `#92400E`) and gracefully populates lookup cards from deterministic **built-in `ValidationService` remediation metadata** (rule descriptions, corrected `"tabla MAST/STPO"`, `VKP2` exclusion warning, `EINA` lookup) plus the **verified `SOP_RECEPCION_SAP` summary notice** (carrying the `S/4HANA`/`EWM`/`VL31N`/`RFUI` conflict warning, never `tab_wiki` prose), keeping all audit/export workflows unblocked.
- **Context & Epistemic Header (Mandatory on Every Entry):**
  - **Context Pill:** `Issue: CURVA_DESCUADRADA | MAT: 1455103 | Row: 18` + manual lookup selector (`T-Code`, `Table`, `Movement`, `Doc Prefix`, `Troubleshooting`).
  - **Provenance Badge (`source_type`):** `[Company-Specific: Ekono / IM 01]` vs `[Generic SAP ECC 6.0]`.
  - **Verification Badge (`verification_status`):** `[Confirmed]` | `[Validate in SAP ECC 6.0]` (`needs_env_check`) | `[Conflicting]` | `[Unverified]`. Any non-`confirmed` entry displays a non-dismissible amber/red 1px-bordered banner listing the exact read-only SAP check (`SE16N` on `T161`/`T156`/`T320`, `System -> Status`, `KNOWLEDGE_EXPANSION_PLAN.md §3.6`).
  - **Release Tag (`sap_release` / `sap_doc_version`):** `ECC 6.0 (6.00.29)` or `S/4HANA / EWM Conflict Flagged`.
- **4 Structured Operational Cards:**
  1. **T-Code & Table Quick Reference:** Exact read-only GUI paths/fields (`ME80FN`, `MB51`, `MB52`, `MB5T`, `MM60`, `SE16N` `EINA`/`MAST`/`STPO` `ACTVT=03` only — never `&SAP_EDIT`), plus explicit `VKP2` warning (`"Retail selling price only — NEVER map to import unit_cost"`).
  2. **Movement Type & Stock Bucket Impact:** Reversal pairs (`101/102`, `641/642`, `501/502`, `643/644`), `MARD` buckets (`LABST`, `INSME`, `SPEME`, `TRAME`), FI-MM `WRX`/`BSX` note, and Ekono caveats (`643`, `998`).
  3. **Discrepancy Triage Checklist:** Deterministic step-by-step checks for the active row (`FALTANTE`, `SOBRANTE`, `PENDIENTE_PARTE_B`, `POSIBLE_CONFUSION_UDM`, `PRECIO_CERO_EN_RECLAMO`).
  4. **Citations & Read-Only Boundary Footer:** Lists `source_citation` (`repo_file:line`, `vba_macro`, `corpus_file`, `https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29`), `last_verified_date`, and permanent notice: *"Read-only operational guidance. Calculations remain deterministic in ReconcileAI; zero SAP write operations (`MIGO`/`MIRO`/`VL31N`/`MR11`) are permitted."*

---

## 5. Keyboard Navigation, TSV Ergonomics, `i18n` Parity & Error States

### 5.1 Keyboard Shortcuts & Grid Navigation
- Global `'c'`/`'C'` key hijacking (`src/i18n.py:601–615`, `A10-6`) is **deleted**. Standard grid typing, screen readers, and shortcuts work natively:
  - `Ctrl+O`: Open Case Loader / Folder
  - `F5` / `Ctrl+F5`: Run Audit / Clear Hash Cache & Reload
  - `F1`: Toggle SAP Companion Drawer
  - `F4`: Jump to Next Validation Error
  - `F8`: Focus Validation Alerts Panel
  - `Ctrl+F`: Focus active tab search box
  - `Ctrl+C`: Copy selected table cells/rows as clean TSV (`75`, `$3.1667`) via `QApplication.clipboard()` (Refactored Streamlit fallback uses `navigator.clipboard.writeText` when `isSecureContext=true`, or read-only `st.code` / auto-select `<textarea readonly>` on HTTP LAN)
  - `Ctrl+E`: Export Official `.xlsm`
  - `Ctrl+Tab` / `Alt+1..9`: Cycle workspace tabs or jump directly to Workspace Tabs `1..9` (`Alt+1` `Bultos` through `Alt+9` `SOP & Reference`; `View 1: Case Loader` uses `Ctrl+O` and `Dock Panel` uses `F1`)

### 5.2 `i18n` Parity & Cleanup (`src/i18n.py`)
Directly audited in `src/i18n.py:19–456`:
- **Existing Catalog (`124` keys, 100% ES/EN parity):** `69` currently used in `app.py`, `5` used internally in `src/i18n.py`, and **`50` currently unused**.
- **Required Actions:**
  1. Decouple `src/i18n.py` from `streamlit` (`render_language_selector`, `get_table_toolbar_js` removed from `src/i18n.py`).
  2. Strip all emojis from `LANGUAGES` (`"ES Español"`, `"EN English"`) and all `80` emoji-containing strings across `40` `TRANSLATIONS` keys (`79` lines / `82` total strings in `src/i18n.py`).
  3. Wire relevant unused keys (`kpi_erp_balance`, `kpi_erp_balance_sub`, `card_*`, `field_*`, `strip1_*`, `strip2_*`, `col_*`) and replace all inline `if current_lang == "es"` branches in UI code with `t(...)`.
  4. Add parity-tested ES/EN keys for: Menu/Toolbar/Status Bar (`menu_*`, `toolbar_*`, `status_*`), Validation Alerts Panel (`tab_validation`, `val_*`), SAP Companion Drawer (`companion_*`, `kb_index_not_built`, `kb_index_corrupt`, `kb_fallback_sop_notice`), and Error/Empty states (`err_xls_unsupported`, `err_xlsm_export_failed`, `err_clipboard_failed`, `err_folder_permission`, `warn_stale_cache_cleared`).

### 5.3 Explicit Error, Loading & Empty States
- **Loading:** Non-blocking thin progress bar in `QStatusBar` + subtle table placeholder driven by `QThreadPool` signals (zero dark shimmer HTML).
- **`.xls` BIFF File Detected:** Inline warning banner in Case Loader (`"Legacy .xls (BIFF) detected. Enable [xls] xlrd adapter or convert to .xlsx before auditing."`).
- **`.xlsm` Export / Permission Failure:** Modal `QMessageBox.critical` showing exact file path, error reason (e.g., workbook locked in Excel), and retry/Save-As option (never muted `st.caption`, `app.py:1687`).
- **Zero Validation Issues:** Explicit `"0 errors, 0 warnings, 0 info issues detected"` confirmation badge in Validation Alerts Panel.
- **SAP Knowledge Index Missing / Empty / Corrupt (Pre-Roadmap #4 State):** When `SapKnowledgeService.get_index_status().state != "ready"` (`not_built`, `empty`, or `corrupt`), display the non-dismissible amber banner **`"Knowledge Index Not Built (Roadmap #4 Pending)"`** (`#FFFBEB` / `#FDE68A` / `#92400E`; plus `error_detail` diagnostic caption when `"corrupt"`), degrade non-fatally to built-in `ValidationService` remediation metadata + verified `SOP_RECEPCION_SAP` summary notice (flagging `S/4HANA`/`EWM`/`VL31N`/`RFUI` conflict, never falling back to `tab_wiki`), and keep all reconciliation, validation, clipboard, and export workflows 100% operational.

---

## 6. Epistemic Separation (R8)

### 6.1 Confirmed Current-UI Findings
- `.streamlit/config.toml:7–13`, `app.py:43–258` CSS gradients/dark theme, emoji usage across `app.py` (`83` lines), `src/i18n.py` (`79` lines / `82` strings across `LANGUAGES` and `40` `TRANSLATIONS` keys), and `src/ui_helpers.py:240`, `app.py:313` external `icons8.com` fetch, `app.py:298,576` discarded `validation`, `app.py:377–387` `folder_input` reset, `app.py:267–343` `@st.cache_data` path cache, `app.py:1090` Python 3.12+ quote syntax, `app.py:1695–1701` `value=True` checkboxes, `src/ui_helpers.py:140–282` pagination/clipboard iframe JS, `src/i18n.py:520–620` `MutationObserver` + `'c'` hijack, and exact `124` (`69` app / `5` internal / `50` unused) `i18n` key counts are directly verified.
- Exact tab boundaries (`tab_reconcile` 15-col `df_t`, `tab_curves` lacking Part filter & using `next(...)`, `tab_excel_export` 18-col `df_exp` with IEEE-754 `unit_cost`, TSV `.0` strips, hardcoded `AU=0`/`AZ=0`, `tab_clipboard` 15-row dead end) are directly verified.

### 6.2 Design Decisions & Assumptions
- **Decision:** Target **PySide6** (`QMainWindow`, `QTableView`, `QDockWidget`, `QApplication.clipboard()`, `QThreadPool`) as Primary ERP UI over `src/services/`, with **Refactored Neutral-Gray Streamlit** as documented fallback sharing the identical visual tokens, layouts, formatting contract, and service calls.
- **Decision:** Add two mandatory surfaces (`Validation Alerts Panel` and read-only `SAP Knowledge / Analyst Companion Drawer`) while replacing `tab_wiki` with `SapKnowledgeService`-backed `SOP & Reference`.
- **Assumption:** Analysts operate primarily on Windows desktop workstations with Excel and SAP GUI open side-by-side (`Ctrl+C`/`Ctrl+V` TSV transfer workflow).

### 6.3 Open Questions for the User
1. **Desktop (`PySide6`) vs. Refactored Browser (`Streamlit`):** Confirm whether native Windows desktop ergonomics (**PySide6** primary recommendation) is preferred over retaining browser-hosted Streamlit with a neutral-gray stylesheet.
2. **Col `AY` Zero-Carton Type (`0` vs `"0"`):** Do downstream corporate `.xlsm` VBA macros expect numeric `0` (`app.py:1636`) or string `"0"` (`app.py:1532`, `generar_cierre_oficial.py:73`) in `AY` on zero-carton / `Parte B` rows?
3. **`SAP Clipboard Sanitizer` Scope:** Should the SAP Clipboard Sanitizer live as a workspace tab (`Tab 5`) or also open as a quick modal dialog from the `Tools`/`Audit` menu feeding directly into `MB51`/`ME80FN` cross-checks?
4. **Checklist & Run History Persistence:** Should `AuditTrailService` store the 7-step Binder Checklist and audit manifests as local JSON files beside the application / case folder, or in an embedded local SQLite store?
