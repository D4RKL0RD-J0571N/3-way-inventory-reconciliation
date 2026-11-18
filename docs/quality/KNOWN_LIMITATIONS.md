# Known Limitations & Operational Constraints — ReconcileAI (Track A)

**Assessment Date:** 2026-11-17  
**Audited Revision:** `589f6aa` (`repository_audit_planning`)  
**Audit Mode:** Strictly read-only / non-modifying  

---

## 1. Real-World `data/` Corpus Coverage & Ingestion Limitations

Read-only filesystem, header, and magic-byte inspection across `D:\Development\Software Development\3 Way Inventory Reconciliation\data` (`71` `IMP-*` subdirectories + `2` root `.xlsm` files, totaling **`593` files**) confirms severe gaps between real import folders and ReconcileAI's supported inputs:

### 1.1 File Format & Extension Breakdown (`593` Files Total)

| Extension | Count | Current ReconcileAI Support | Verified Findings & Operational Impact |
|---|---:|---|---|
| **`.xlsm`** | 148 | Supported (via `openpyxl`) | Supported when sheets match hardcoded `'CIERRE RECEP'`, `'CONTROL DATA'`, `'BASE DE DATOS'` assumptions; fails if template structure shifts or file is legacy `.xls`. |
| **`.xlsx`** | 97 | Supported (via `openpyxl`) | Same structural constraints as `.xlsm`. |
| **`.pdf`** | 160 | **Unsupported (0%)** | `130` contain `/Font` or `/ToUnicode` markers (digital/hybrid); **`30` lack `/Font`/`/ToUnicode` and contain `/Image` markers** (scanned/image-only, e.g. `IMP-105  NCR\NCR\da064a8_48511.pdf`, `IMP-205  NCR\NCR-RECEPCION\NCR IMP-205    10 BULTOS FALTANTES.pdf`, `IMP-207 (NCR)\Factura\20260818165021642_0003.pdf`). No PDF parser or OCR dependency exists in `requirements.txt`. |
| **`.msg`** | 71 | **Unsupported (0%)** | Outlook email messages containing vendor notifications, approvals, and packing updates cannot be read. |
| **`.xls`** | 62 | **False-Accepted / Crashes** | **62/62** verified with OLE2/BIFF binary compound document header (`d0cf11e0a1b11ae1`, e.g. `IMP-99\Factura\EN26E01002-98 INV.PL.xls`). Both `src/config.py:52` and `app.py:395` accept `.xls` in disk discovery, yet `openpyxl` fails to load BIFF `.xls` (`xlrd` is absent from `requirements.txt`), and `tab_uploader` (`app.py:514,518`) blocks `.xls` uploads. |
| **`.eml`** | 35 | **Unsupported (0%)** | Standard MIME email files (e.g. `IMP-81\INFORMACION GENERAL\__PEDIDO NACIONALIZACION 81__.eml`) cannot be parsed. |
| **`.jpeg`** | 12 | **Unsupported (0%)** | Dock/damage/packing photographs cannot be ingested or attached to NCR evidence. |
| **`.db` / `.zip` / `.a` / `.pptx`** | 4 / 2 / 1 / 1 | **Unsupported (0%)** | Auxiliary/system/archive files ignored or unsupported. |

### 1.2 Input File Discovery Failure (`0/71` Resolution Rate in `src/config.py`)

- **`cli.py` and `generar_cierre_oficial.py` (`src/config.py:43–67`):** Running `resolve_local_import_files()` across all `71` `IMP-*` folders in `data/` yields **`0/71` usable pairs** (`ok_both=0`, `miss_qc=3`, `miss_cierre=3`, `miss_both=65`). Root causes:
  1. `imp_dir.iterdir()` is **non-recursive**, missing QC/closure files stored inside `INFORMACION GENERAL/` or `FACTURAS/` subdirectories.
  2. Strict filename matching (`"control" and "calidad"` / `"qc"`, and `["progress", "cierre", "plantilla", "recep"]`) misses real corporate naming conventions.
- **`IMP-81` Confirmed False Blind Setup (`app.py:404` & `src/config.py:60`):**
  - `IMP-81\IMP-81.xlsm` contains `CIERRE RECEP`, `Bultos`, and `CIERRE DESP`.
  - `IMP-81\INFORMACION GENERAL\IMP 81.xlsm` contains `CONTROL DATA`, `BASE DE DATOS`, `DIST. PARTE A`, `ESQUEMA ALISTO PARTE B`, and `ESQUEMA ALISTO PARTE C `.
  - Because `IMP 81.xlsm` lacks `"calidad"`, `"qc"`, or `"conteo"` in its filename, `app.py:404` returns `qc_candidates_idx = []`, triggering `has_qc_candidate == False` and warning the operator to enter **Blind Setup Pause Mode** even though a complete QC workbook is present.

---

## 2. Unimplemented ERP, SAP & Phase 2 Dispatch Capabilities

Despite claims in `README.md:50`, `docs/AUDITORIA_DESPACHO_MACRO.md:110–118`, `docs/SOP_RECEPCION_SAP.md:14,75–78`, and `docs/RELEASE_NOTES_DEMO.md:58–60`:

1. **Zero Phase 2 Dispatch Support:**
   - No parser, engine calculation, validator rule, UI tab, or export exists for **`CIERRE DESP`**, **`PARA N.C DESP`**, **`STOCK CONTROL DE INVENTARIO`**, **`Check list DESPACHO`**, **`MB52`** stock checks, **`Mov 641`** reconciliation, or **`Mov 643`** pending billing.
   - `PrepackItem.sap_mov_641` (`src/models.py:44`, parsed at `src/parsers.py:454`) is display-only, and `PrepackItem.revuelto` (`src/models.py:49`) is never populated by any parser (`0.0` permanently).
2. **Unimplemented SAP `MB51 Mov 101 - Mov 102` Net Balance Verification:**
   - `docs/SOP_RECEPCION_SAP.md:75–78` specifies `Total Unidades Ingresadas = Mov 101 - Mov 102` reconciled against `received_real_units`.
   - `src/parsers.py` never parses `Mov 102` reversals, `src/engine.py` and `src/validator.py` never compare `sap_mov_101` against `received_real_units` or `quantity_inv_41_units`, and `kpi_erp_balance` in `src/i18n.py:160` is among **50 unused i18n keys**.
3. **Zero Live SAP Integration (Read or Write):**
   - No RFC/BAPI/OData/SAP GUI scripting connector exists; all ERP inputs arrive via manual copy-paste into Excel templates or the dead-end `tab_clipboard` text box (`app.py:1394–1431`).
4. **S/4HANA & EWM Documentation Conflict:**
   - `docs/SOP_RECEPCION_SAP.md:1,30–31` labels the process `SAP S/4HANA` and `SAP EWM (VL31N / RFUI)`, conflicting with the organization's `SAP ECC 6.0` baseline expectation (`ORIGINAL_REQUEST.md:28`) and classic ECC transactions/tables (`ME80FN`, `SE16N`, `EINA`, `MAST`, `STPO`, `MB51`) inside the same SOP.

---

## 3. Calculation, Packaging, Precision & Template Structural Limits

Summarized from `docs/quality/CODE_AUDIT.md`:
1. **Unrendered Validation (`app.py:298,576`):** `validate_all_items` detects `CURVA_DESCUADRADA`, `BOM_FALTANTE`, `POSIBLE_CONFUSION_UDM`, `PRECIO_CERO_EN_RECLAMO`, `ESTILO_PROVEEDOR_VACIO`, and `MATERIAL_EN_MULTIPLES_PO`, **but none are shown in the Streamlit UI, `cli.py`, or `generar_cierre_oficial.py`**.
2. **Metadata `import_id` Overwrite (`src/parsers.py:104–107`):** Substring `'importaci'` causes Row 22 `'Faltante Importación'` (`288`) to overwrite `IMP-207` with `'288'`; Row 21 `'Dañado Importación'` and Row 23 `'Sobrantes Importación'` can trigger the same overwrite if non-zero.
3. **NCR Undercount (`src/parsers.py:603`):** `ncr_amount = (falt + dan) * unit_cost` ignores Col `J` (`revuelto`/shortage) and unannotated shortfall-only `max(ped41 - rec, 0)` variances (`$550.08` vs. `PARA N.C RECEP.!G6`/`G13` `$869.76` on `IMP-207`).
4. **Child BOM Destruction on Non-`LIMPIO` Rows (`src/parsers.py:619`):** `c_rec = int(c_theo) if status == "LIMPIO" else 0` forces `0` received units on every child SKU whenever any parent discrepancy exists (`IMP-103`).
5. **False `PENDIENTE_PARTE_B` on `ESQUEMA ALISTO` Sheets (`src/parsers.py:276,561`):** Single-container imports whose QC workbook has `ESQUEMA ALISTO PARTE B/C` sheets have unreceived/zeroed items fabricated into `PENDIENTE_PARTE_B` (`21` items on `IMP-207`).
6. **Packaging Grammar & `raw_ay` Gaps (`src/parsers.py:41–59,474–519`):**
   - `parse_cierre_recepcion` ignores compound strings in `raw_ay` (`Empaque`) when `raw_as` is scalar (`IMP-187`, `IMP-200`).
   - `parse_packaging_string` drops trailing `+<units>` (`'8*72+24'`, `'9*48+12'`), truncates `boxes_full`/`boxes_comp` at 2 segments on 3+ segment strings, and cannot disambiguate `A*B+<loose>` (`IMP-200`) from `A*B+<total>` (`IMP-192`: `'56*12+672'`).
7. **Numeric Formatting & Precision Corruption:**
   - Discrete quantities format with `.0` in TSV copy strips (`'65.0\t65.0'`, `'12.0\t888.0'`) and mix `int`/`float` (`0` vs `0.0` vs `"0"`).
   - `df_exp` (`app.py:1471`) exports raw IEEE-754 floats (`3.1666666666666665`), while UI detail cards truncate sub-cent prorated prices (`$1.380833` $\rightarrow$ `$1.38`).
   - `clean_material_code` (`src/sanitizer.py:40–46`) loses precision on >15-digit scientific notation and corrupts alphanumeric codes like `'1E2'` $\rightarrow$ `'100'`.
   - `clean_number` (`src/sanitizer.py:85–115`) cannot separate Latin 3-decimal values from US thousands (`'1,234'`) and drops negative signs after currency symbols (`'$-50.00'` $\rightarrow$ `+50.0`).
8. **Excel Structural & Patching Fragility:**
   - `parse_quality_control_boms` hardcodes `BASE DE DATOS` cols `4,5,6,8,9,10` from Row 5 and `CONTROL DATA` headers at Row 6.
   - `parse_cierre_recepcion` hardcodes `CIERRE RECEP` headers at Row 12 (`cols 1..64`), lacks header matchers for `cat`/`mov641`/`unit_cost`, and caps `extract_cierre_metadata` fallback sums at Row `500` (`IMP-187` reaches Row `612`).
   - `.xlsm` patchers write to `row = 13 + idx` without row-ID anchoring, misaligning whenever blank/subtotal rows are skipped, and do not clear trailing `'Bultos'` template rows.
   - `generar_cierre_oficial.py:124–125` prints hardcoded `$0.00 USD (LIMPIO)` even when discrepancies exist.

---

## 4. Scalability, Performance, Statefulness & Offline Contradictions

1. **Non-Streaming `parse_cierre_recepcion` Performance (`src/parsers.py:360`):**
   - `parse_quality_control_boms` uses `read_only=True` (`0.219s–0.595s` across tested `sandbox_benchmark` cases).
   - `parse_cierre_recepcion` loads full XML + VBA/styles without `read_only=True`, taking **`3.108s–4.457s`** per workbook (measured on `IMP-100` through `IMP-182`).
2. **Zero Persistence, Multi-User Concurrency, or Audit Trail:**
   - All results exist ephemerally in Streamlit `st.session_state` and `@st.cache_data` RAM. There is no database, case history, user attribution, approval workflow, or immutable export audit log.
   - Because `app.py:267` caches `run_audit` by file path string (when invoked via Disk Explorer) without hashing bytes or checking `mtime`, and `Recargar Auditoría` (`app.py:341–343`) does not clear `@st.cache_data`, modified Excel files at the same path return stale cached audits.
3. **Offline / Air-Gapped Contradiction:**
   - `app.py:313` requests `https://img.icons8.com/color/96/warehouse.png` from an external CDN on every page load, failing in offline warehouse environments despite documentation guaranteeing 100% local execution.
4. **Packaging & Python Version Constraints:**
   - `tools/build_exe.py` excludes `cli.py` and `generar_cierre_oficial.py` from `dist/ReconcileAI/`, so standalone users have no headless CLI.
   - `app.py:1090` uses PEP 701 nested double quotes inside an f-string (`f"{'... ' if current_lang == "es" ...}"`), raising `SyntaxError` on Python `<= 3.11` despite `preparar_equipo.bat:13` instructing users to install Python `3.10+`.
   - Missing `src/__init__.py`, `tests/__init__.py`, and `pyproject.toml`/`pytest.ini` causes bare `pytest` to fail collection (`ModuleNotFoundError: No module named 'src'`).
5. **Global Browser Shortcut Interference:**
   - `src/i18n.py:601–615` blocks `'c'`/`'C'` keydown events globally on `window.parent` (outside inputs/textareas), interfering with table cell copy/search shortcuts and accessibility tools.

---

## 5. Epistemic Separation

### 5.1 Confirmed Findings
- All `data/` file counts (`593` files, `160` PDFs [`~130` digital / `~30` image-only], `62` BIFF `.xls`, `71` `.msg`, `35` `.eml`, `12` `.jpeg`), `0/71` `resolve_local_import_files()` resolution, `IMP-81` false blind-setup trigger, non-streaming `parse_cierre_recepcion` timings (`3.1s–4.5s`), unimplemented Phase 2 / `Mov 102` checks, `app.py:313` external CDN call, and `app.py:1090` Python 3.12+ f-string syntax are directly verified.

### 5.2 Assumptions
- The `130` digital vs. `30` image-only PDF classification used raw byte marker inspection (`/Font` or `/ToUnicode` vs. `/Image`) without adding a PDF library dependency; PDFs storing `/Font` inside compressed `/ObjStm` streams would require a dedicated PDF library (`pymupdf`/`pdfplumber`) for exact confirmation.
- `I:\My Drive\IMP` (`tools/operational_stress_test.py:29`) is assumed to have been the original shared drive path prior to local `data/` synchronization.

### 5.3 Unverified Domain / SAP Items
- Whether `PARA N.C RECEP.!G6` (`869.76` literal $\rightarrow$ `G12 = =+G6` $\rightarrow$ `G13 = =SUM(G10:G12)` = `$869.76` on `IMP-207`) reflects manual freight/expense allocations or partial Col `L` sobrante offsets relative to shortfall-only `max(ped41 - rec, 0) * unit_cost` = Col `J` (`$340.74`) + Col `N` (`$550.08`) = **`$890.82`** (vs. **`$750.90`** net across all 35 rows after subtracting `-$139.92` from Col `L` sobrantes).
- Actual corporate SAP version (ECC 6.0 vs. S/4HANA) and active warehouse management module (LE-WM vs. EWM vs. MM-IM only).

### 5.4 Prioritized Open Questions for the User
1. **P1 — `IMP-207` NCR Calculation Rule:** How was `$869.76` derived on `IMP-207`, and should ReconcileAI compute `ncr_amount` from shortfall-only `max(ped41_u - rec_u, 0) * unit_cost`, net `(ped41_u - rec_u) * unit_cost`, `(falt + dan + revuelto) * unit_cost`, or strictly manual `N + O`?
2. **P2 — Packaging Grammar (`+` in Col `AY`):** In `IMP-200` (`'9*48+12'`), `+12` represents loose remainder units (`9*48 + 12 = 444`), whereas in `IMP-192` (`'56*12+672'`), `+672` repeats the total. Which convention should be enforced, and should parser validation cross-check `AY` against `AT`/`H`?
3. **P3 — Scope of `.xls`, `.pdf`, `.msg`, and `.eml` Ingestion:** Which of the unsupported formats in `data/` (`62` BIFF `.xls` vendor packing lists, `160` PDFs including `~30` scanned NCRs/invoices, `71` `.msg`, `35` `.eml`) are high-priority automated inputs for Phase 1 / Phase 5 document intelligence?
4. **P4 — SAP ECC 6.0 vs. S/4HANA EWM Confirmation:** Please confirm the exact production SAP release and whether `VL31N`/`RFUI` (EWM) are used or should be corrected to ECC 6.0 transactions (`MIGO`, `MB51`, `ME80FN`, `LT01`/`LT12` if LE-WM).
5. **P5 — Air-Gapped Offline & Portable CLI Requirements:** Is strict offline execution required (requiring replacement of `img.icons8.com` in `app.py:313` and `data:`/`local` font/assets), and should `cli.py` / `generar_cierre_oficial.py` be bundled into the portable Windows build?
