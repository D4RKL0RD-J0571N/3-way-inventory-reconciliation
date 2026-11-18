# Line-by-Line Code & Test Suite Audit — ReconcileAI (Track A)

**Assessment Date:** 2026-11-17  
**Audited Revision:** `589f6aa` (`repository_audit_planning`)  
**Audit Mode:** Strictly read-only / non-modifying  
**Verification Resources (Read-Only):**
- `D:\Development\Software Development\3 Way Inventory Reconciliation\sandbox_benchmark` (`operational_stress_test_results.json` & per-case `<CASE>/unmodified_originals/`)
- `D:\Development\Software Development\3 Way Inventory Reconciliation\data` (`71` `IMP-*` folders, `593` files)

---

## 1. Executive Defect Summary

Full line-by-line inspection across `src/` (`9` modules, `2,199` lines; `8` active modules / `2,127` lines excluding dead `src/heuristic_parser.py`), `app.py` (`1,773` lines), `cli.py` (`101` lines), `generar_cierre_oficial.py` (`130` lines), `run_standalone.py` (`32` lines), `tools/` (`493` lines), and `tests/` (`827` lines, `27` test functions), corroborated by read-only execution against `sandbox_benchmark/` and `data/`, identifies **49 numbered defects** (`A1-1`..`A8-8`, `A10-1`..`A10-10`) across Areas 1–8 and 10, plus a **16-entry `try/except` catalog** in Area 9:

| Severity | Count | Representative Numbered Defects (`A1-1`..`A10-10`) |
|---|---:|---|
| **Critical** | 5 | Unrendered `validation` engine (`app.py:298,576`, `A10-1`); `IMP-207` `ncr_amount` undercount (`$550.08` vs `$869.76`, `src/parsers.py:603`, `A3-1`); `extract_cierre_metadata` `'importaci'` overwrite (`'IMP-207'` $\rightarrow$ `'288'`, `src/parsers.py:104-107`, `A6-4`); `c_rec = 0` child receipt destruction on non-`LIMPIO` rows (`IMP-103`, `src/parsers.py:619`, `A1-1`); `.xlsm` row `13 + idx` positional misalignment (`app.py:1616`, `generar_cierre_oficial.py:49`, `tools/operational_stress_test.py:270`, `A8-5`). |
| **High** | 27 | Missing `+<units>` grammar (`IMP-192`/`IMP-200`/`IMP-103`/`IMP-196`, `A2-1`); ignored `raw_ay` (`IMP-187`/`IMP-200`, `A2-3`); `PRK`-to-`Bultos` conflation (`A2-4`); duplicated Zero-Carton rule with `0` vs `"0"` drift (`A2-5`); unwired `sap_mov_101`/`sap_mov_641`/`revuelto` (`A3-2`); hardcoded `"0"` in `AU`/`AZ` (`app.py:1543,1547`, `A3-3`); false `PENDIENTE_PARTE_B` via `ESQUEMA ALISTO PARTE B/C` (`IMP-207`, `A3-5`); `75.0` vs `75` TSV decimals (`A5-1`); raw IEEE-754 `unit_cost` (`3.1666666666666665`, `A5-3`); `clean_number`/`clean_material_code` corruption (`A5-6`, `A7-1`); `qc_details`/`boms_by_parent` duplicate overwrites (`A6-1`, `A6-2`); unsupported 62 BIFF `.xls` files (`A8-7`); `paginate_dataframe` bugs (`A10-3`); `folder_input` reset (`A10-2`); `icons8` CDN call (`A10-5`); `generar_cierre_oficial.py` hardcoded clean print (`A10-9`). |
| **Medium** | 13 | Banker's rounding on `c_theo` (`A1-4`); `parts[0]`/`parts[1]` 3+ truncation (`A2-2`); silent defaults (`ratio=1`, `curve=0`/`12`, `A1-5`); `compute_summary` ignoring `sobrante`/`diff` (`A3-4`); `int()` truncation (`A5-2`); `:.2f` sub-cent truncation (`A5-4`); `000` prefix rule (`A7-2`); Row 500 cap (`A8-4`); trailing `'Bultos'` rows not cleared (`A8-6`); `tab_clipboard` dead-end (`A10-4`); global `'c'` key suppression (`A10-6`); `tab_checklist` unkeyed `True` checkboxes (`A10-7`); Python `<= 3.11` `SyntaxError` (`app.py:1090`, `A10-8`). |
| **Low** | 4 | Calculator `,.0f` rounding (`app.py:815,824`, `A5-5`); dead `parse_multi_po` hyphen mismatch (`src/sanitizer.py:127-138`, `A7-4`); dead `src/heuristic_parser.py` (`A8-8`); single-file QC/Cierre auto-select collision (`app.py:409-410`), duplicate imports (`src/config.py:5-10`), and `operational_stress_test.py` `c - 1` NCR extraction bug (`A10-10`). |

> **Area 9 (`try/except` Catalog) Note:** Section 2 Area 9 separately catalogs all `16` `try/except` blocks across the repository, flagging `2` `HIGH` silent swallows (`src/parsers.py:280-292`, `tools/operational_stress_test.py:266-292`) and `3` `MEDIUM` silent catches (`app.py:392-399`, `app.py:1687-1688`, plus `src/sanitizer.py:111-115` overlapping `A5-6`) without double-counting them in the `49` numbered rows above.

---

## 2. Detailed Line-by-Line Audit (R2 Focus Areas 1–10)

### Area 1: Product Curves & BOM Explosion

| ID | Severity | File & Lines | Symbol | Root Cause & Concrete Evidence |
|---|---|---|---|---|
| **A1-1** | **CRITICAL** | `src/parsers.py:619` | `parse_cierre_recepcion` (`c_rec`) | **Partial receipt child destruction:** `c_rec = int(c_theo) if status == "LIMPIO" else 0` zeroes out **all** child SKU `received_units` whenever a parent item has any discrepancy (`FALTANTE`, `SOBRANTE`, `DAÑADO`, `PENDIENTE_PARTE_X`), even when physical units were received. Verified on `IMP-103`: `1431895` (`SOBRANTE`, `rec=1164.0`, `ped41=1152.0`) reports `theoretical_units=[192, 384, 384, 192]`, `received_units=[0, 0, 0, 0]`, and `ChildSKU.difference=[-192, -384, -384, -192]` (`sum(c.difference for c in item.children) == -1152`, since `difference` lives on `ChildSKU`, not `PrepackItem`); `1460786` (`FALTANTE`, `rec=2088.0`, `ped41=2100.0`) reports `theoretical_units=[350, 700, 700, 350]`, `received_units=[0, 0, 0, 0]`, and `ChildSKU.difference=[-350, -700, -700, -350]` (`sum(c.difference for c in item.children) == -2100`). |
| **A1-2** | **HIGH** | `src/validator.py:25-34`; `app.py:298,576` | `CURVA_DESCUADRADA` | `validate_item_integrity` flags `sum(c.ratio for c in item.children) != item.curve_total` as `ERROR`, **but** (a) `parse_cierre_recepcion` still explodes child quantities using the mismatched ratios (`src/parsers.py:607-628`), and (b) `app.py:576` unpacks `validation = audit_res["validation"]` and never renders it. |
| **A1-3** | **HIGH** | `src/validator.py:37-44`; `app.py:298,576` | `BOM_FALTANTE` | Flags `item.curve_total > 1 and not item.children and (item.received_real_units > 0 or item.container_part in ["", "A"])` as `WARNING`, **but** output is never rendered in `app.py`. Additionally, non-prepack items (`um == "UN"`) whose missing Col `BB` defaults to `curve = 12` (`src/parsers.py:461`) trigger false-positive `BOM_FALTANTE` warnings. |
| **A1-4** | **MEDIUM** | `src/parsers.py:611-616` | `parse_cierre_recepcion` (`c_theo`) | **Banker's rounding & PRK/UN mismatch:** `c_theo = int(round(ped41_prk * c_ratio))` when `ped41_prk > 0`, else `int(round((ped41_u / curve) * c_ratio))`. Python 3 round-half-to-even on fractional ratios causes `sum(c_theo) != ped41_u`; when `ped41_prk > 0` disagrees with `ped41_u / curve`, child `theoretical_units` follow `ped41_prk` and silently diverge from `ped41_u`. |
| **A1-5** | **MEDIUM** | `src/parsers.py:217,316,461` | `parse_quality_control_boms`, `parse_cierre_recepcion` | **Silent defaults (`ratio=1`, `curve=0`/`12`):** Missing/invalid BOM child ratios default to `1` (`line 217`); in `CONTROL DATA`, `line 316` defaults `curve` to `12` when `col_curva >= len(row)` and to `clean_int(row[col_curva])` (`0`) when the cell exists but is blank/invalid; in `CIERRE RECEP`, missing/invalid `curve` defaults to `12` (`line 461`), and `parse_cierre_recepcion` never falls back to `qc_details[mat]['curve']`. |

### Area 2: Prepacks & Packaging Grammar

| ID | Severity | File & Lines | Symbol | Root Cause & Concrete Evidence |
|---|---|---|---|---|
| **A2-1** | **HIGH** | `src/parsers.py:41-59` | `parse_packaging_string` | **Missing `+<units>` loose-piece grammar:** Regex `r'(\d+(?:\.\d+)?)\s*[*xX]\s*(\d+(?:\.\d+)?)'` ignores trailing `+<units>`. Verified read-only: `parse_packaging_string('372*1+372')` (`IMP-192`) returns `total_boxes=372.0, total_units=372.0, is_compound=False` (drops `+372`); `parse_packaging_string('8*72+24')` (`IMP-200`) returns `total_boxes=8.0, total_units=576.0, is_compound=False` (drops `+24`); also confirmed in `IMP-103` (`'6*192 +48'`) and `IMP-196` (`'11*54+6'`). Meanwhile `IMP-192` contains conflicting notation `B*U+TOTAL` (`'56*12+672'`) and `B*U*TOTAL` (`'12*32*384'`). |
| **A2-2** | **MEDIUM** | `src/parsers.py:45-48` | `parse_packaging_string` | **`parts[0]`/`parts[1]` truncation for 3+ segments:** `total_boxes` and `total_units` sum all matches, while `boxes_full`/`units_per_box_full` store `parts[0]` and `boxes_comp`/`units_per_box_comp` store `parts[1]`. Verified read-only on `'10*12 5*6 1*3'`: `total_boxes=16.0, total_units=153.0`, but `boxes_full*units_per_box_full + boxes_comp*units_per_box_comp = 150.0` (drops `1*3`). |
| **A2-3** | **HIGH** | `src/parsers.py:474-519` | `parse_cierre_recepcion` (`raw_ay`) | **`raw_ay` (`Empaque`) ignored when `raw_as` is scalar:** `parse_packaging_string` runs only on `raw_as` and `raw_aw`. Verified in `IMP-187` (83 rows, e.g. `1456087`: `AS=10, AT=738, AY='9*72 1*90'`) and `IMP-200` (e.g. `1425861`: `AS=15, AY='14*72 1*54'`): `bultos_enteros=10.0, unidades_por_bulto_entero=0.0, bultos_complementos=0.0`. `line 516` (`empaque_str.replace('.', '', 1).isdigit()`) also fails on compound `Empaque`. |
| **A2-4** | **HIGH** | `app.py:770,885,1527,1631`; `generar_cierre_oficial.py:67` | `PRK` vs `Bultos` fallback | When `bultos_descarga == 0` and QC lacks `bultos`, all four sites fall back to **`quantity_inv_41_prk`** (Prepacks, not physical cartons—one carton often holds multiple prepacks). `app.py:770` also divides `total_received_units / max(total_prk_41, 1)` (mixing physical received units with invoiced PRKs). |
| **A2-5** | **HIGH** | `app.py:837-890,1515-1536,1619-1640`; `generar_cierre_oficial.py:54-77`; `tools/operational_stress_test.py:271-283` | Duplicated Zero-Carton rule | Absent from `src/engine.py`/`src/parsers.py` and copy-pasted 5 times (4 prod/tool + 1 test). Type mismatch in Col `AY`: `app.py:1636` writes `int` `0`, whereas `app.py:1532`, `generar_cierre_oficial.py:73`, and `operational_stress_test.py:277` write `str` `"0"`. Condition mismatch: `app.py:837-838` (`tab_bultos`) and `operational_stress_test.py:272` zero Packing List whenever `received_real_units == 0` (even when `quantity_inv_41_units > 0` on `FALTANTE`), whereas `app.py:1529,1633` and `generar_cierre_oficial.py:70` preserve Packing List unless `received == 0 and ped41 == 0`. |

### Area 3: Quantity Calculations & NCR Math

| ID | Severity | File & Lines | Symbol | Root Cause & Concrete Evidence |
|---|---|---|---|---|
| **A3-1** | **CRITICAL** | `src/parsers.py:603` | `parse_cierre_recepcion` (`ncr_amount`) | **`ncr_amount = (falt + dan) * unit_cost` zeroes unannotated `ped41 > rec` shortfalls:** When Col `N` (`falt`) and Col `O` (`dan`) are blank (`None` $\rightarrow$ `0.0`), `FALTANTE` lines with `rec_u < ped41_u` get `ncr_amount = 0.0`. Verified on `IMP-207 (NCR)`: 4 `FALTANTE` rows (`5033284`: `G=90, H=54, N=None, AF=2.13`; `5033285`: `G=72, H=18, N=None, AF=2.13`; `5033293`: `G=600, H=576, N=None, AF=2.23`; `5033296`: `G=600, H=588, N=None, AF=2.23`) yield `ncr_amount = 0.0`, and Row 38 (`5033307`: `G=1080, H=756, J=36, N=288, AF=1.91`) yields `$550.08`, totaling **`$550.08`** vs `PARA N.C RECEP.!G6` (`869.76` literal $\rightarrow$ `G12 = =+G6` $\rightarrow$ `G13 = =SUM(G10:G12)` = **`$869.76`**). |
| **A3-2** | **HIGH** | `src/models.py:43,44,49`; `src/parsers.py:453-454`; `src/engine.py:37-65` | `sap_mov_101`, `sap_mov_641`, `revuelto` | `sap_mov_101` and `sap_mov_641` are parsed (`src/parsers.py:453-454`) and never reconciled in `engine.py` or `validator.py`. `PrepackItem.revuelto` (`src/models.py:49`) is **never parsed at all** (always `0.0`) despite `IMP-207` `DATOS GENERALES` Row 24 having `162` revueltos and `app.py:1002,1211,1302` rendering `"REVUELTO"` badges. |
| **A3-3** | **HIGH** | `app.py:1543,1547` | `descarga_rows` (`AU`, `AZ`) | Preview DataFrame and `ZDESCARGA_PACKING_LIST_AS_AZ.xlsx` unconditionally hardcode `"Dif Físico vs 41 (AU)": 0` and `"Dif Fact vs Sol (AZ)": 0` regardless of actual discrepancies between `unidades_desc`/`unidades_pl` and `received_real_units`. |
| **A3-4** | **MEDIUM** | `src/engine.py:64` | `compute_summary` (`status`) | `status = "LIMPIO" if (total_falt == 0 and total_dan == 0) else "CON_DISCREPANCIAS"` ignores `total_sob` (`sobrante`) and item-level `diff = rec_u - ped41_u`, reporting global `"LIMPIO"` when surpluses or unannotated shortfalls exist. |
| **A3-5** | **HIGH** | `src/parsers.py:276-280,559-564,578-580` | `parse_quality_control_boms`, `parse_cierre_recepcion` | **False `PENDIENTE_PARTE_B` fabrication:** Regex `r'PARTE\s+([A-Z0-9]+)'` matches `'ESQUEMA ALISTO PARTE B'` and `'ESQUEMA ALISTO PARTE C '` in `IMP-207`, populating `_detected_parts = ['A', 'B', 'C']`. At `line 561`, `other_parts[0]` classifies **21 items** with `ped41=0.0, rec=0.0` as `container_part='B'`, `status='PENDIENTE_PARTE_B'`. |

### Area 4: Unit Conversions & `POSIBLE_CONFUSION_UDM`

| ID | Severity | File & Lines | Symbol | Root Cause & Concrete Evidence |
|---|---|---|---|---|
| **A4-1** | **HIGH** | `src/models.py:60`; `src/parsers.py:314,477`; `app.py:1527,1631` | `PrepackItem.um` | `um` (`"PRK"`, `"UN"`, etc.) is parsed and displayed in `tab_bultos` (`app.py:869`), **never driving unit conversion** between `quantity_inv_41_prk` and `quantity_inv_41_units` or guarding carton fallbacks (`app.py:1527,1631` still divide by `curve_total` even when `um == "UN"`). |
| **A4-2** | **HIGH** | `src/validator.py:47-54` | `POSIBLE_CONFUSION_UDM` | **One-sided `2.5x` check & unrendered output:** Only checks `received_real_units > (quantity_inv_41_units * 2.5)`. Misses inverse confusion (`received << invoiced / 2.5`), misses `curve=2` confusions (`2.0x <= 2.5x`), false-positives on tiny overages (e.g. `1` invoiced vs `3` received), ignores `item.um`/`curve_total`, and is never displayed in `app.py`. |

### Area 5: Floating-Point Precision & Whole-Number vs. Monetary Decimals

| ID | Severity | File & Lines | Symbol | Root Cause & Concrete Evidence |
|---|---|---|---|---|
| **A5-1** | **HIGH** | `src/sanitizer.py:61-115`; `app.py:1550-1551`; `generar_cierre_oficial.py:58-72` | `clean_number`, TSV strips | **`75.0` vs `75` decimal pollution:** `clean_number` returns `float` for integer counts (`75.0`). `app.py:1550-1551` (`f"{bultos_desc}\t{unidades_desc}"` and `f"{bultos_pl}\t{unidades_pl}\t{curva_empaque}"`) emits **`'65.0\t65.0'`** (`IMP-100`) and **`'12.0\t888.0'`** (`IMP-182`) into clipboard TSV strips instead of `'65\t65'` / `'12\t888'`. |
| **A5-2** | **MEDIUM** | `app.py:854-858,885-890,1141,1316-1322,1463-1474` | `int(...)` casts | Discrete-count columns silently floor-truncate `float` via `int(...)` (e.g., `10.9` $\rightarrow$ `10`) instead of rounding or validating integrality. |
| **A5-3** | **HIGH** | `app.py:1471,1478-1483` | `df_exp` (`Valor unitario PEDIDO NAC $`) | Passes unrounded IEEE-754 `i.unit_cost` directly into `df_exp` (`st.dataframe` + `.xlsx` export). Verified on real workbooks: **`3.1666666666666665`** (`IMP-182` `1455314` `AF14`, formula `=IFERROR(AL14/G14,0)`), **`2.1300000000000003`** (`IMP-207` `5033285`), **`1.3808333333333334`** (`IMP-103`), **`4.799019607843137`** (`IMP-225`), **`4.969767441860465`** (`IMP-364`). |
| **A5-4** | **MEDIUM** | `app.py:1139,1168,1349,1387` | `:.2f` / `$%.2f` formatting | Truncates 3–6 decimal prorated prepack unit costs to 2 decimals (`$1.380833` $\rightarrow$ `$1.38`, `$4.79902` $\rightarrow$ `$4.80`), causing `displayed_unit_cost * qty != displayed_total_invoice` (`app.py:1140`). |
| **A5-5** | **LOW** | `app.py:815,824` | Carton Calculator (`res_prk`) | `res_prk = cal_units / max(cal_curve, 1)` formatted as `f"{res_prk:,.0f} PRK"`, rounding fractional prepacks (e.g., `10 / 6 = 1.67` displays `"2 PRK"`). |
| **A5-6** | **HIGH** | `src/sanitizer.py:85-115` | `clean_number` | Verified read-only: `clean_number('1,234')` $\rightarrow$ `1234.0` (cannot disambiguate US thousands vs Latin 3-decimal unit cost); `clean_number('1,2,3')` $\rightarrow$ `123.0`; `clean_number('1.234.567')` $\rightarrow$ `0.0`; `clean_number('$-50.00')` $\rightarrow$ **`50.0`** (loses negative sign because `s.startswith("-")` runs before stripping `$`). |

### Area 6: Duplicate Records & Overwrites

| ID | Severity | File & Lines | Symbol | Root Cause & Concrete Evidence |
|---|---|---|---|---|
| **A6-1** | **HIGH** | `src/parsers.py:312-325` | `parse_quality_control_boms` (`qc_details`) | Repeated material codes in `CONTROL DATA` (split pallets/tarimas) **overwrite** `qc_details[m_clean]` with the last row instead of accumulating `qty`, `bultos`, `unidades`, `bultos_enteros`, `bultos_complementos`, `total_prk`. |
| **A6-2** | **HIGH** | `src/parsers.py:210-222` | `parse_quality_control_boms` (`boms_by_parent`) | If a `current_parent` reappears in `BASE DE DATOS`, `if current_parent not in boms_by_parent:` preserves the existing list and appends duplicate children, doubling `sum_ratios`. Also, a non-`PRK` row with a new `bom_id` sets `current_parent = mat_clean`. |
| **A6-3** | **HIGH** | `src/validator.py:86-101`; `src/parsers.py:482-533`; `app.py:1080,1368` | `MATERIAL_EN_MULTIPLES_PO` | (1) `validation` is unrendered in `app.py`; (2) `parse_cierre_recepcion` lets multiple PO rows with the same `material_code` each fall back independently to `qc_map[mat_code]`, double-counting dock counts; (3) `app.py:1368` (`next((i for i in items if ...), None)`) displays only the first PO row in `tab_curves`; (4) `app.py:1080` (`opt_labels`) overwrites dict keys on duplicate formatted labels. |
| **A6-4** | **CRITICAL** | `src/parsers.py:104-107` | `extract_cierre_metadata` (`'importaci'`) | **Confirmed `import_id` overwrite bug:** Loops `DATOS GENERALES` rows `1..24` matching `'importaci' in lbl_lower and 'importadora' not in lbl_lower`. On `IMP-207 (NCR)`, Row 7 (`'Importación'`) sets `'IMP-207'`, then **Row 22 (`'Faltante Importación'`, `C22=288`) overwrites `meta['import_id']` with `'288'`** (confirmed in `operational_stress_test_results.json:428-429`, `"case": "IMP-207 (NCR)"` and `"import_id": "288"`). |

### Area 7: Material Identification

| ID | Severity | File & Lines | Symbol | Root Cause & Concrete Evidence |
|---|---|---|---|---|
| **A7-1** | **HIGH** | `src/sanitizer.py:40-46` | `clean_material_code` | **Scientific notation & alphanumeric corruption:** `f = float(s); s = f"{int(round(f))}"` loses precision above `2^53` (verified: `clean_material_code(123456789012345678.0)` $\rightarrow$ `'123456789012345680'`), corrupts alphanumeric codes containing `'e'`/`'E'` (verified: `clean_material_code('1E2')` $\rightarrow$ `'100'`), and silently ignores `OverflowError` on `'1e500'`. |
| **A7-2** | **MEDIUM** | `src/sanitizer.py:56-57` | `clean_material_code` | `if s.isdigit() and len(s) > 10 and s.startswith("000"):` misses leading zeros on 8–10 digit codes (`'0001001001'`) or 1–2 leading zeros (`'001001001001'`), while stripping legitimate `'000'` prefixes from 11+ digit EAN/GTIN barcodes. |
| **A7-3** | **HIGH** | `src/parsers.py:282-290` | `parse_quality_control_boms` (`PARTE <X>`) | Scans only rows `1..500`, cols `1..6` (`row_p[:6]`) for `val_s.isdigit() and len(val_s) >= 7`: misses `<7`-digit codes, alphanumeric SKUs, and rows `>500`/cols `>6`, while matching 7+ digit PO/invoice numbers as material codes. |
| **A7-4** | **LOW** | `src/sanitizer.py:127-138` | `parse_multi_po` | Docstring promises splitting `'4300006764-6804-6805'`, but `re.split(r"[,;/]+", cleaned)` omits `-` (verified: returns `['4300006764-6804-6805']`). Confirmed in `IMP-207` Col `A` (`'4300006709-7847'`). Function is unused outside tests. |

### Area 8: Excel Parsing & Patching Fragility

| ID | Severity | File & Lines | Symbol | Root Cause & Concrete Evidence |
|---|---|---|---|---|
| **A8-1** | **HIGH** | `src/parsers.py:191-222` | `parse_quality_control_boms` (`BASE DE DATOS`) | Hardcodes 0-based indices `4, 5, 6, 8, 9, 10` starting at `min_row=5` with zero header detection. |
| **A8-2** | **HIGH** | `src/parsers.py:234-270` | `parse_quality_control_boms` (`CONTROL DATA`) | Hardcodes header row `6` (`if r_idx == 6:`) and data start `min_row=7`; maps `col_part` (`line 248`) and never reads it. |
| **A8-3** | **HIGH** | `src/parsers.py:392-440` | `parse_cierre_recepcion` (`CIERRE RECEP`) | Hardcodes header row `12` (`cols 1..64`) and data start `row 13`; omits header matchers for `cat` (`3`), `mov641` (`21`), and `unit_cost` (`32`), causing shifted `unit_cost` columns to stay `0.0`. |
| **A8-4** | **MEDIUM** | `src/parsers.py:152-159` | `extract_cierre_metadata` | Fallback carton sum caps at `min(ws_cr.max_row + 1, 500)` and hardcodes cols `45`/`49`, truncating sheets `>500` rows (`IMP-187` `max_row=612`). |
| **A8-5** | **CRITICAL** | `app.py:1616-1617`; `generar_cierre_oficial.py:48-50`; `tools/operational_stress_test.py:270` | `.xlsm` patchers (`13 + idx`) | `parse_cierre_recepcion` skips blank/subtotal rows lacking a `material_code` (`src/parsers.py:442-443`), while all three `.xlsm` patchers write sequentially to `row = 13 + idx` without tracking source row numbers. Any blank/subtotal row shifts all subsequent writes onto wrong material rows. |
| **A8-6** | **MEDIUM** | `app.py:1663-1673`; `generar_cierre_oficial.py:106-116` | `'Bultos'` sheet sync | Writes rows `3..3+len(items)-1` without clearing trailing template/footer rows (`IMP-187`: `Bultos max_row=269` vs `234`; `IMP-200`: `max_row=165` vs `130`). |
| **A8-7** | **HIGH** | `app.py:395,507`; `src/config.py:52`; `src/parsers.py:80,178,360` | `.xls` glob vs `openpyxl` | Disk globbing accepts `.xls`, and `data/` contains **62 OLE2/BIFF `.xls` files** (`d0cf11e0...`), yet `openpyxl` fails on `.xls` (`xlrd` absent) and `tab_uploader` (`app.py:514,518`) restricts to `["xlsm", "xlsx"]`. |
| **A8-8** | **LOW** | `src/heuristic_parser.py:1-72` | `detect_headers_in_sheet` | Dead module with zero callers anywhere in the repository. |

### Area 9: Complete Silent Errors (`try/except`) Catalog

| File & Lines | Caught Exception | Behavior & Impact |
|---|---|---|
| `src/parsers.py:28-38` | `except ValueError: pass` | Intended fallback from scalar `float(s)` to regex parsing. |
| `src/parsers.py:177-180` | `except Exception:` | Masks `read_only=True` failure (including `FileNotFoundError`/`BadZipFile`) and retries full-load. |
| `src/parsers.py:280-292` | `except Exception: pass` | **HIGH:** Silently swallows errors scanning `PARTE <X>` sheets; misclassifies multi-part items as single-container/Part A. |
| `src/parsers.py:330-333,670-673` | `except Exception: pass` | Silently ignores `wb.close()` exceptions. |
| `src/sanitizer.py:42-46` | `except (ValueError, OverflowError): pass` | Silently leaves non-finite/overflowing `'e'` strings (`'1e500'`) unparsed. |
| `src/sanitizer.py:111-115` | `except (ValueError, TypeError): return default` | **MEDIUM:** Silently coerces malformed strings (`'1.234.567'`) to `0.0` (overlaps `A5-6`). |
| `src/sanitizer.py:122-125` | `except (ValueError, OverflowError): return default` | Silently coerces `inf`/`NaN` to `0` in `clean_int`. |
| `src/i18n.py:489-492` | `except Exception: return template` | Silently swallows format `KeyError`/`ValueError`, exposing raw `{placeholder}` strings. |
| `src/i18n.py:544-548,577,617` | JS `catch(e) {}` | Silently suppresses DOM/iframe/`MutationObserver` JS errors. |
| `src/ui_helpers.py:275-277` | JS `catch(e) { console.error(...) }` | Logs clipboard copy failure only to DevTools console; user receives no error message. |
| `app.py:392-399` | `except Exception: pass` | **MEDIUM:** Silently swallows `PermissionError`/`OSError` during `p.rglob("*")`, showing empty/partial file list. |
| `app.py:453-455,478-480,503-505,541-543,551-552,1428-1429` | `except Exception as ...:` | Displays `st.error(...)` without Python `logging` or traceback capture. |
| `app.py:1687-1688` | `except Exception as e_xlsm:` | **MEDIUM:** Downgrades `.xlsm` export failure to muted `st.caption(...)` without alert or logging. |
| `cli.py:11-14`; `generar_cierre_oficial.py:11-14` | `except Exception: pass` | Silently ignores `sys.stdout.reconfigure(encoding='utf-8')` failure. |
| `tools/operational_stress_test.py:168-169` | `except Exception as e:` | Prints note and returns zeroed `human_metrics`, skewing benchmark comparisons. |
| `tools/operational_stress_test.py:266-292` | `except Exception as e_pt: pass` | **HIGH:** Completely swallows `.xlsm` patching/save failures while printing `-> OK!`. |

### Area 10: UI, Data Integrity & Script Bugs

| ID | Severity | File & Lines | Symbol | Root Cause & Concrete Evidence |
|---|---|---|---|---|
| **A10-1** | **CRITICAL** | `app.py:298,576` | `validation = audit_res["validation"]` | `validate_all_items(items)` runs and caches results on every audit, then `app.py:576` unpacks `validation` and never renders any `ERROR`, `WARNING`, or `INFO` issues in the UI. |
| **A10-2** | **HIGH** | `app.py:377-387` | `folder_input`, `use_folder_btn` | `st.text_input` lacks a `key` / `session_state` binding. Clicking `use_folder_btn` assigns local `folder_input` and triggers `st.rerun()`, immediately resetting the text input to `default_folder`. |
| **A10-3** | **HIGH** | `src/ui_helpers.py:140-188` | `paginate_dataframe` | (1) `default_page_size: int = 100` (`line 143`) is unused—`st.selectbox` hardcodes `index=1` (`line 162`, selecting `50` when `[25, 50, 100, 200]`); (2) if `current_page` in `session_state` exceeds new `total_pages` after filtering or resizing page size, `st.number_input(..., max_value=total_pages)` raises `StreamlitValueAboveMaxError`; (3) `min(page_size_options)` raises `ValueError` if empty. |
| **A10-4** | **MEDIUM** | `app.py:1394-1431` | `tab_clipboard` | Dead-end feature: renders `df_raw.head(15)` with no download button, clipboard copy, or handoff to `run_audit`. `pd.read_csv(..., sep=sep)` also breaks on SAP GUI leading/trailing `|` delimiters. |
| **A10-5** | **HIGH** | `app.py:313` | `st.sidebar.image(...)` | Fetches `https://img.icons8.com/color/96/warehouse.png` over HTTPS on every render, breaking air-gapped operation and contradicting `README.md:167` / `src/i18n.py:122`. |
| **A10-6** | **MEDIUM** | `src/i18n.py:601-615` | `blockClearCacheShortcut` | Attaches global capture `keydown` listener blocking `'c'`/`'C'` outside `INPUT`/`TEXTAREA`/`isContentEditable`, interfering with table search/copy and keyboard accessibility. |
| **A10-7** | **MEDIUM** | `app.py:1695-1701` | `tab_checklist` (`c_p1..c_p7`) | Seven checkboxes default to `True` before any printing occurs, lack `key=`, and reset on reload or language switch. |
| **A10-8** | **MEDIUM** | `app.py:1090` | `st.info(f"...")` | Nested double quotes `"es"` inside `f"{'...' if current_lang == "es" ...}"` require Python 3.12+ (PEP 701); causes `SyntaxError` on Python 3.10/3.11 (even though `preparar_equipo.bat:13` tells users to install Python 3.10+). |
| **A10-9** | **HIGH** | `generar_cierre_oficial.py:124-125` | `main()` stdout | Prints hardcoded `Diferencia de Unidades : 0 UN (Cuadrado)` and `Monto Nota de Crédito : $0.00 USD (LIMPIO)` regardless of actual `summary.total_ncr_amount` or discrepancies. |
| **A10-10** | **LOW** | `app.py:409-410`; `src/config.py:5-10`; `tools/operational_stress_test.py:155-160` | Auto-select / imports / NCR extraction | When a folder has 1 Excel file, `qc_default_idx=0` and `cierre_default_idx=0` select the same file for both inputs. `src/config.py:5-10` duplicates `Path`/`typing` imports. `operational_stress_test.py:158` reads `c - 1` instead of `c - 3` in `PARA N.C RECEP.`, recording `"human_ncr_amount": 0.0` on `IMP-207`. |

---

## 3. Documentation vs. Code Drift

| Documented Claim | Source Location | Verified Reality | Evidence Location |
|---|---|---|---|
| `src/__init__.py` & `tests/__init__.py` exist | `README.md:133,154`; `README.es.md:132,153` | Both files absent; `src/` & `tests/` are implicit PEP 420 namespace dirs | `src/`, `tests/` |
| Test suite is `22/22 Passing` | `README.md:153`; `README.es.md:152` | `27` `def test_*` functions exist and pass via `python -B -m pytest` | `tests/*.py` |
| Project trees list all modules | `README.md:132-151`; `README.es.md:131-150` | Omit `src/heuristic_parser.py` and `tools/operational_stress_test.py` | `src/`, `tools/` |
| Consistent release version | `README.md:8` (`v1.2.0`), `docs/WIKI.md:8` (`v1.0.0`), `docs/WIKI.md:182` (`v1.1.0`), `docs/SOP_RECEPCION_SAP.md:6` (`v1.0.0`) | Mixed `v1.0.0` / `v1.1.0` / `v1.2.0` across docs | `README.md`, `docs/*.md` |
| Phase 2 Dispatch (`CIERRE DESP`, `MB52`, `Mov 641`, `Mov 643`) pre-wired / supported | `README.md:50`; `docs/AUDITORIA_DESPACHO_MACRO.md:110-118`; `docs/RELEASE_NOTES_DEMO.md:58-60` | `sap_mov_641` is parsed (`src/parsers.py:454`) and `revuelto` exists (`src/models.py:49`), but neither is used in `engine.py`/`validator.py`/`app.py`. Zero `CIERRE DESP`, `MB52`, or `Mov 643` logic exists | `src/engine.py`, `src/validator.py`, `app.py` |
| `MB51 Mov 101 - Mov 102` net balance check | `docs/SOP_RECEPCION_SAP.md:14,75-78`; `src/i18n.py:160` (`kpi_erp_balance`) | `sap_mov_101` is display-only; `Mov 102` is unparsed; `kpi_erp_balance` is among **50 unused keys** in `src/i18n.py` | `app.py`, `src/parsers.py`, `src/i18n.py` |
| `100%` offline / zero external servers | `README.md:167`; `docs/RELEASE_NOTES_DEMO.md:51`; `src/i18n.py:122` | `app.py:313` fetches `https://img.icons8.com/color/96/warehouse.png` | `app.py:313` |
| SAP `S/4HANA` & `SAP EWM` (`VL31N / RFUI`) | `docs/SOP_RECEPCION_SAP.md:1,30-31` | Conflicts with ECC 6.0 expectation (`ORIGINAL_REQUEST.md:28`) and classic ECC transactions/tables (`ME80FN`, `SE16N`, `EINA`, `MAST`, `STPO`, `MB51`) inside the same file | `docs/SOP_RECEPCION_SAP.md` |
| `parse_multi_po` splits hyphens (`4300006764-6804-6805`) | `src/sanitizer.py:129` | `re.split(r"[,;/]+", cleaned)` does not split `-` | `src/sanitizer.py:136` |

---

## 4. Read-Only Test Suite Audit (`tests/`)

### 4.1 Verbatim Execution Outputs

1. **Module invocation (`$env:PYTHONDONTWRITEBYTECODE = "1"; python -B -m pytest -p no:cacheprovider -v`):**
   ```text
   platform win32 -- Python 3.13.5, pytest-9.0.2, pluggy-1.6.0 -- C:\Python313\python.exe
   rootdir: C:\Users\Jostin\.gemini\antigravity\worktrees\3 Way Inventory Reconciliation\repository_audit_planning
   collecting ... collected 27 items
   ...
   ============================= 27 passed in 2.86s ==============================
   ```
2. **Bare executable (`$env:PYTHONDONTWRITEBYTECODE = "1"; pytest -p no:cacheprovider -v`):**
   Exits with code `2` (`0 items / 4 errors` during collection):
   ```text
   tests\test_bultos_empaque.py:6: in <module>
       from src.models import PrepackItem, PackagingBreakdown, ReconciliationSummary
   E   ModuleNotFoundError: No module named 'src'
   !!!!!!!!!!!!!!!!!!! Interrupted: 4 errors during collection !!!!!!!!!!!!!!!!!!!
   ============================== 4 errors in 1.08s ==============================
   ```
   *(Caused by missing `src/__init__.py`, `tests/__init__.py`, `pyproject.toml`, `pytest.ini`, or `conftest.py`.)*

### 4.2 Per-Test Coverage & Quality Assessment (`27` Tests, `0` `@pytest.mark.parametrize`)

| File | Test Function (Lines) | Target | Assessment / Defect |
|---|---|---|---|
| `tests/test_bultos_empaque.py` | `test_parse_packaging_scalar` (`10–18`) | `parse_packaging_string` | Basic scalar check; no negative/comma/`+<units>` cases |
| `tests/test_bultos_empaque.py` | `test_parse_packaging_single_pattern` (`20–30`) | `parse_packaging_string` | Basic `50*12`/`10x24`; no `raw_ay` integration |
| `tests/test_bultos_empaque.py` | `test_parse_packaging_compound` (`32–47`) | `parse_packaging_string` | Misses `A*B+C` (`8*72+24`), `A*B+TOTAL` (`56*12+672`), and 3+ segments |
| `tests/test_bultos_empaque.py` | `test_parse_packaging_empty_or_none` (`49–53`) | `parse_packaging_string` | Adequate null/nan check |
| `tests/test_bultos_empaque.py` | `test_bultos_math_in_prepack_item` (`55–88`) | `PrepackItem` | **Tautological:** calls zero production functions; asserts manually supplied constructor values |
| `tests/test_bultos_empaque.py` | `test_compute_summary_bultos` (`90–130`) | `compute_summary` | Omits metadata package fallback and NCR aggregation |
| `tests/test_bultos_empaque.py` | `test_dual_carton_modes_dock_vs_strict_customs` (`133–178`) | `PrepackItem` | **Tautological:** never invokes `app.py` dual-mode logic or `parse_cierre_recepcion` |
| `tests/test_bultos_empaque.py` | `test_zero_units_produce_zero_cartons` (`180–237`) | None (inline copy) | **Inline re-implementation:** duplicates Zero-Carton `if/else` (`213–231`) inside the test rather than calling production code |
| `tests/test_i18n.py` | `test_default_language` (`16–20`) | `DEFAULT_LANGUAGE` | Adequate constant check |
| `tests/test_i18n.py` | `test_translation_fallback` (`22–34`) | `t()` | Does not check ES/EN parity or unused keys |
| `tests/test_i18n.py` | `test_translation_interpolation` (`36–40`) | `t()` | Adequate interpolation check |
| `tests/test_i18n.py` | `test_toolbar_js_generation` (`42–52`) | `get_table_toolbar_js` | Does not verify `render_language_selector` or `'c'` shortcut behavior |
| `tests/test_i18n.py` | `test_paginate_dataframe_returns_dataframe` (`54–64`) | `paginate_dataframe` | **Misses active branch:** `len(df)=10 <= 50` hits early return (`src/ui_helpers.py:152`), never testing pagination |
| `tests/test_parsers.py` | `test_dynamic_header_detection_with_inserted_column` (`10–66`) | `parse_quality_control_boms` | Adequate Row 6 + `Tarima` column shift test |
| `tests/test_parsers.py` | `test_metadata_extractor` (`68–95`) | `extract_cierre_metadata` | Stops at Row 15; misses Row 22 `'Faltante Importación'` overwrite (`'288'`) |
| `tests/test_parsers.py` | `test_synthetic_full_reconciliation_clean` (`97–164`) | `parse_quality_control_boms`, `parse_cierre_recepcion`, `compute_summary` | **Silent Row-1 QC bypass:** `CONTROL DATA` places headers at Row 1 (`line 105`) with 3 total rows; `src/parsers.py:234-237,296` only reads Row 6 / `min_row=7`, so `qc_details == {}` throughout the test |
| `tests/test_parsers.py` | `test_dynamic_multi_part_c_d_extraction` (`166–200`) | `parse_quality_control_boms` | Misses `ESQUEMA ALISTO PARTE B` false-positive sheet names |
| `tests/test_parsers.py` | `test_single_container_fallback_never_fabricates_part_b` (`202–253`) | `parse_quality_control_boms`, `parse_cierre_recepcion` | Same Row-1 QC bypass (`qc_details == {}`); relies on `boms={}` |
| `tests/test_parsers.py` | `test_dynamic_filter_items_arbitrary_parts_and_statuses` (`255–336`) | `filter_items` | Misses `"NONE"` / `"SIN_PARTE"` filter branch |
| `tests/test_parsers.py` | `test_blind_setup_mode_without_qc` (`338–396`) | `parse_cierre_recepcion`, `compute_summary` | Puts `Cantidad facturada (UN)` in Col 13, which fails `col_map` matching (`cantidad pedido 41`), relying on fallback `ped41_u=0.0` |
| `tests/test_sanitizer.py` | `test_clean_string` (`15–21`) | `clean_string` | Adequate |
| `tests/test_sanitizer.py` | `test_clean_material_code` (`23–34`) | `clean_material_code` | Misses `>2^53` scientific notation precision loss & `'1E2'` corruption |
| `tests/test_sanitizer.py` | `test_clean_number` (`36–56`) | `clean_number` | Misses `'1,234'`, `'1,2,3'`, `'1.234.567'`, `'$-50.00'` |
| `tests/test_sanitizer.py` | `test_parse_multi_po` (`58–61`) | `parse_multi_po` | Misses hyphenated POs (`'4300006764-6804-6805'`) |
| `tests/test_sanitizer.py` | `test_validator_curve_mismatch` (`63–87`) | `validate_item_integrity` | Adequate |
| `tests/test_sanitizer.py` | `test_validator_uom_confusion` (`89–107`) | `validate_item_integrity` | Misses inverse `< 0.4x` UOM confusion |
| `tests/test_sanitizer.py` | `test_validator_missing_price_on_claim` (`109–128`) | `validate_item_integrity` | Misses `validate_all_items`, `BOM_FALTANTE`, `ESTILO_PROVEEDOR_VACIO`, `MATERIAL_EN_MULTIPLES_PO` |

**Modules with 0% Test Coverage:** `app.py` (exports, TSV strips, `.xlsm` patching, calculator, clipboard parser, disk explorer), `cli.py`, `generar_cierre_oficial.py`, `src/config.py`, `src/heuristic_parser.py`, `tools/build_exe.py`, `tools/operational_stress_test.py`, and `render_clipboard_copy_button` / active `paginate_dataframe` in `src/ui_helpers.py`.

---

## 5. Epistemic Separation

### 5.1 Confirmed Findings
- All 49 numbered code/UI/test defects (`A1-1`..`A8-8`, `A10-1`..`A10-10`) plus the 16-entry Area 9 `try/except` catalog, line numbers, symbols, verbatim `pytest` outputs (`27 passed` vs `4 errors`, exit `0` vs `2`), doc-vs-code drift items, and external read-only reproductions (`IMP-207`, `IMP-182`, `IMP-100`, `IMP-103`, `IMP-187`, `IMP-192`, `IMP-200`, `data/`) were directly verified during this audit.

### 5.2 Assumptions
- `tools/operational_stress_test.py` was not re-executed because it writes files into `sandbox_benchmark/`; findings on that script come from static code inspection plus reading the existing `operational_stress_test_results.json`.

### 5.3 Unverified Domain / SAP Items
- Exact accounting composition of `PARA N.C RECEP.!G6 = 869.76` (`G6` literal $\rightarrow$ `G12 = =+G6` $\rightarrow$ `G13 = =SUM(G10:G12)` = `$869.76`) in `IMP-207 (NCR)`: `CIERRE RECEP` Col `J` (`$340.74`) + Col `N` (`$550.08`) = **`$890.82`** (matching shortfall-only `max(ped41 - rec, 0) * unit_cost` across the 5 shortfall rows, a `$21.06` delta above `$869.76`), whereas net `(ped41 - rec) * unit_cost` across all 35 rows after subtracting `-$139.92` from Col `L` sobrantes is **`$750.90`**.
- Whether `IMP-192` (`'56*12+672'`) vs. `IMP-200` (`'8*72+24'`) reflects vendor-specific or operator-specific notation differences.

### 5.4 Open Questions for the User
1. On `IMP-207`, how should `ncr_amount` treat Col `J` (`revuelto`/shortage) and unannotated `ped41 - rec` variances relative to Col `N`/`O`, and how was `$869.76` reconciled?
2. For packaging strings containing `+` (`'8*72+24'` vs. `'56*12+672'`), what is the canonical business rule when `+` denotes loose pieces vs. appended total units?
3. When `received_real_units == 0` and `quantity_inv_41_units > 0` (`FALTANTE`), should Packing List columns (`AW:AY`) preserve invoiced Packing List values (as in `app.py:1529,1633` and `generar_cierre_oficial.py:70`) or zero them out (as in `tab_bultos` `app.py:837-838` and `operational_stress_test.py:272`)?
