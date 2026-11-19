# SAP Knowledge Expansion Plan (`KNOWLEDGE_EXPANSION_PLAN.md`)

**Milestone:** M2 — Track D: SAP Knowledge Engineering  
**Target System:** ReconcileAI (3-Way Inventory Reconciliation Copilot)  
**Expected ERP Baseline:** SAP ECC 6.0 (SAP ERP 6.0 `version=6.00.29` Help Portal baseline, with IS-Retail / LE-WM / decentralized-or-embedded EWM touchpoints pending live environment validation)  
**Status:** Planning & Gap Assessment (Strictly Non-Modifying; Read-Only Audit)

---

## 1. Executive Summary & Scope (R5, R8, R1 Alignment)

### 1.1 Purpose
This document audits all SAP ERP knowledge currently embedded across the ReconcileAI repository, Python source code, corporate Excel/VBA macro templates, and operational import corpus, classifies every claim by epistemic certainty, identifies missing SAP functional and technical knowledge, and defines a bounded, source-backed, version-aware knowledge base (KB) architecture for Milestone M2 (Track D, Requirements **R5** and **R8**).

### 1.2 Verified Evidence Sources & Authoritative Documentation Anchor
All empirical findings in this plan derive exclusively from direct read-only inspection of four evidence tiers, supplemented by a user-provided official SAP Help Portal baseline anchor:
1. **Repository Markdown Documentation (7 tracked files):** `docs/SOP_RECEPCION_SAP.md` (120 lines), `docs/CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md` (175 lines), `docs/AUDITORIA_DESPACHO_MACRO.md` (118 lines), `docs/WIKI.md` (202 lines), `docs/RELEASE_NOTES_DEMO.md` (137 lines), `README.md` (173 lines), and `README.es.md` (172 lines).
2. **Python Application Codebase:** `src/parsers.py` (679 lines), `src/validator.py` (108 lines), `src/heuristic_parser.py` (72 lines), `src/sanitizer.py` (138 lines), `src/models.py` (94 lines), `src/engine.py` (138 lines), `src/i18n.py` (620 lines), `app.py` (1,773 lines), `cli.py` (101 lines), `generar_cierre_oficial.py` (130 lines), and `tools/operational_stress_test.py` (363 lines).
3. **External Benchmark Workbooks & VBA Macros (`D:\Development\Software Development\3 Way Inventory Reconciliation\sandbox_benchmark`):** 15 benchmarked import cases (QC `.xlsm` and Cierre `.xlsm` templates) plus decompressed MS-OVBA source (`xl/vbaProject.bin`, `Módulo1` `480` raw CRLF lines / `22,273` bytes, raw MD5 `6c2f475b58e1`; `477` non-`Attribute` lines, LF-normalized UTF-8 MD5 `90a614f461f2`, or `479` lines / `bca9b31838d5` retaining the 2 `ArrayRowSel` attributes; and `ThisWorkbook` `26` non-`Attribute` lines / `34` raw) across representative closure workbooks and `data\Macro para cierre hannia.xlsm`.
4. **External Real Data Corpus (`D:\Development\Software Development\3 Way Inventory Reconciliation\data`):** 71 `IMP-*` folders + 2 root macro workbooks totaling 593 files (`.pdf`, `.xlsm`, `.xlsx`, `.xls`, `.msg`, `.eml`, `.jpeg`, `.pptx`, etc.), including `@ekonocr.com` operational emails, `Adobe LiveCycle Designer ES 9.8` credit note PDFs, Costa Rica DUA declarations, and store distribution workbooks.
5. **User-Provided Official SAP Documentation Anchor (Follow-Up Reference, 2026-11-17):** `https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29` (SAP ERP 6.0 / ECC 6.0 Help Portal baseline). Recorded here strictly as the authoritative official documentation portal anchor for Track D reference verification; individual sub-page titles, sub-URLs, and topic contents have **not** been fetched or fabricated in this planning pass and must be retrieved and cited per topic during Track D KB execution.

### 1.3 Core Finding
The external operational corpus and VBA macros reveal an active **Importadora Almacema S.A. (`SOCIEDAD IM 01`) / Cadena de Tiendas Ekono** retail import and distribution workflow anchored in classic **SAP ECC / IS-Retail-style GUI transactions and tables** (`ME80FN`, `MM60`, `VKP2`, `MB51`, `MB52`, `MB5T`, `SE16N` on `EINA`/`MAST`/`STPO`), custom Z-objects (`Z-Carga`, `ZDESCARGA`/`ZDECARGA`, `Z51 -Z52`, `"7.UBICACIÓN MONITOR"`), distribution center plant **`C001`** (`A001`, `A002`, `A004`, `03`–`99`), and store plants **`C101+`**.

However, existing repository documentation suffers from severe epistemic and structural gaps:
- **Version & Architecture Conflict:** `docs/SOP_RECEPCION_SAP.md:1,30-31` claims `SAP S/4HANA` and `SAP EWM` (`VL31N / RFUI`) while the rest of the docs, VBA macros, and expected organization baseline point to **SAP ECC 6.0**, and `README.md:165` paradoxically claims the architecture is `"100% Agnostic"` of any ERP.
- **Unverified / Orphaned Claims:** `VL31N`, `18...` inbound deliveries, and `80...` handling units appear **zero times** across all 15 benchmark cases, VBA macros, emails, and extractable PDFs.
- **Code-vs-Template & Semantic Bugs:** `src/validator.py`'s SAP integrity checks are computed in `app.py:298,576` but **never rendered** in the UI; `src/heuristic_parser.py` is **dead code** that dangerously equates IS-Retail sales pricing (`VKP2`) with import `unit_cost`; `PrepackItem.sap_mov_101` (`Col I` in `src/parsers.py:373` vs `Col BF` in `Mueve51` / `docs/AUDITORIA_DESPACHO_MACRO.md:43`) and `sap_mov_641` are **never reconciled**.
- **Missing Functional Domains:** Existing docs omit FI-MM valuation and GR/IR clearing impact, SAP authorization boundaries (`S_TCODE`, `M_MSEG_BWA`), LE-WM vs EWM architecture, reversal pairing and stock types (`LABST`/`INSME`/`SPEME`/`TRAME`), and systematic troubleshooting procedures.

### 1.4 Alignment with R1 (ReconcileAI Development Rules) & R6 (Track E AI/RAG)
In compliance with **R1** and **R6**:
- **Zero Fabrication:** No SAP procedures, transactions, tables, or SAP Note numbers are fabricated. Unverified standard ECC 6.0 candidates are explicitly tagged `[Generic ECC 6.0 Candidate — Validate in Env]`.
- **Epistemic Hygiene (R8):** Confirmed facts, company-specific customizations, generic SAP knowledge, conflicting/unverified claims, items needing official SAP documentation, and items requiring live SAP ECC 6.0 validation are strictly separated.
- **Deterministic & Read-Only Guardrails:** SAP knowledge in the KB informs validation warnings and future Track E RAG explanations with explicit citations and verification status, **never** replacing deterministic reconciliation math (`PRK`/`UN`, NCR amounts, net `101 - 102`) and **never** permitting SAP write operations without separate authorization and safety design.

---

## 2. Inventory of What Is Already Documented (Exact Citations)

### 2.1 Existing Repository Markdown Documents (7 Files)

| File | Lines | SAP Concepts, Transactions, Prefixes, Movements & Rules Documented | Key Omissions / Drift |
|---|---:|---|---|
| `docs/SOP_RECEPCION_SAP.md` | 1–120 | Claims `SAP S/4HANA` (line 1), `"Agnóstica"` (line 5), `SAP EWM` + `VL31N / RFUI` (lines 30–31). Mermaid flow (`S1`–`S4`, lines 10–37). Prefix/Movement dictionary (`43...`, `41...`, `45...`, `40...`, `18...`, `80...`, `101`, `102`, `641`, `501`, lines 41–54). `ME80FN` (`Shift + F4`, `UMP` `PRK`/`UN`, lines 60–66). `SE16N` on `EINA` (`IDNLF`/`MATNR`) and `MAST`/`STPO` (`MENGE`) (lines 68–73). `MB51` `101 - 102` balance (lines 75–78). Dock `ZDescarga` + multi-warehouse sum (`Central` + `Resurtido` + `Especiales`, lines 82–97). Error table incl. `81` vs `1,728` PRK/UN confusion & scientific notation (lines 101–108). 7-document physical binder checklist (lines 112–120). | Zero procedural steps for `VL31N`, `RFUI`, `45...`, `40...`, `18...`, `80...`, `501`; omits `642`, `643`, `MB52`, `MM60`, `VKP2`, `998`, authorizations, and FI-MM valuation. |
| `docs/CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md` | 1–175 | 3-level segregation: Cartons (`AS`, `AW`, `Bultos!C`, `CONTROL DATA!32`) vs Prepacks (`PRK`, `BD`, `BE`, `BB` curve, `BC` UM) vs Units (`F`, `G`, `H`, `AT`, `AX`) (lines 7–45). 4 Packaging Patterns A–D & compound grammar `'39*60 1*78'` (lines 49–111). VBA macros `Format43`/`Mueve43`, `Mueve41`, `MueveCurva` & unpopulated `AS:AT`/`AW:AY` gap (lines 115–125). `Bultos` sheet 10-col mapping `B..J` (lines 129–143). `DATOS GENERALES` formulas `B5`, `B6`, `B7`, `D5` (lines 147–156). `Bulto Cero` rule for `U_recibidas == 0` or `Parte B` (`IMP-81`, lines 160–174). | Omits `Format51`, `Mueve51`, `NoDuplicados`, `CalcDesg`, `MueveMatProv`, `SE16N-EINA`. |
| `docs/AUDITORIA_DESPACHO_MACRO.md` | 1–118 | Dual Phase 1 (`CIERRE RECEP`, `Bultos`, `PARA N.C RECEP.`, `Mov 101`) & Phase 2 (`CIERRE DESP`, `PARA N.C DESP`, `STOCK CONTROL DE INVENTARIO`, `Mov 641`) architecture (lines 7–30). `Mueve43` row deletion `1011` across both sheets (lines 37–41). `Mueve51` dual `ERASER` blocks: `MB51 101-102` -> `CIERRE RECEP!BF` (58) and `MB51 641-642` -> `CIERRE RECEP!U` (21) (lines 42–44). `CIERRE DESP` 72-col spec (`A..BT`: Zone A `1-10, 37-60`, Zone B `11-36`, incl. `J` PRKs, `N` `MB52`, `AB` `Mov-641`, `AC` `=+AB13-J13`, `AG` `Mov 643`, `AI` final stock, lines 50–90). `STOCK CONTROL DE INVENTARIO`, `PARA N.C DESP`, `Check list DESPACHO` (lines 94–106). `PrepackItem` Phase-2 fields (`sap_mov_101`, `sap_mov_641`, `container_part`, `revuelto`, lines 110–118). | Never mentions `45...` store orders, `MB5T`, `Z51 -Z52`, or `998`; leaves `BF` vs `I` and `AB13-J13` unit mismatch unresolved. |
| `docs/WIKI.md` | 1–202 | SAP GUI extraction summary (`43...` & `41...` in `ME80FN`, `MB51` `101 - 102`, lines 55–66). QC column shift (`Tarima`), duplicate BOM IDs, deleted Col `H` recovery, `Bulto Cero`, `Parte B`, multi-warehouse sum (lines 80–167). 7-document binder checklist (lines 133–141). Phase 2 dispatch roadmap (`Mov 641`, `CIERRE DESP`, lines 171–202). | Version header conflict (`v1.0.0` at line 8 vs `v1.1.0` at line 182 vs `v1.2.0` elsewhere). |
| `docs/RELEASE_NOTES_DEMO.md` | 1–137 | Documents `v1.2.0` (lines 1, 8): `Bulto Cero` rule, dual carton modes (`Conteo de Muelle` vs `Estricto Aduanero`), `PAUSA_SIN_CONTROL_CALIDAD`, `Mov 101`/`641`, `ZDescarga` TSV strips, zero proprietary corporate names/codes claim (line 84), and roadmap items for NCR PDF export, DUA OCR scanner, and multi-user sharing (lines 133–137). | Asserts `"Agnostic Compliance"` / zero corporate references (lines 83–84) and leaves the `S/4HANA`/`EWM` vs `ECC 6.0` vs `ERP-agnostic` conflict unaddressed; omits `ME80FN`, `SE16N`, `MB52`, `MM60`, `VKP2`, authorizations, and valuation. |
| `README.md` | 1–173 | Overview of 3-way reconciliation (`PO 43`, `Doc 41`, `Mov 101`/`641`, lines 13–52), `ZDescarga` strips, `v1.2.0`, claims `22/22 Passing` tests (line 153) and `"100% Agnostic Architecture... any ERP (SAP, Oracle, NetSuite)"` (line 165), zero company branding (line 166). | Contradicts SAP-coupled code/docs (`43`/`41`/SAP GUI); outdated test count (22 vs 27 in `PROJECT.md:17`). |
| `README.es.md` | 1–172 | Spanish mirror of `README.md`, explicitly naming `ME80FN, MB51` at line 57, `ZDescarga` at lines 53, 104, `22/22 Pasando` at line 152, `"Diseño 100% Agnóstico"` at line 164. | Same ERP-agnostic contradiction and test count drift. |

### 2.2 Python Codebase SAP References & Implementation Gaps

| File & Symbol | Exact Lines | SAP Reference / Behavior Implemented | Verified Gap / Risk |
|---|---|---|---|
| `src/parsers.py` — `extract_cierre_metadata` | 74–167 | Reads `'DATOS GENERALES'` (`'nacionalizaci'`/`'pedido nac'` -> `invoice_number` `41...`) and falls back to `'CIERRE RECEP'` rows 1–10 & cols `45` (`AS`), `49` (`AW`). | Reads pre-exported Excel cells; zero direct SAP integration. |
| `src/parsers.py` — `parse_quality_control_boms` | 169–335 | Extracts BOM parent/child ratios from `'BASE DE DATOS'`/`'BOM'`/`'CURVAS'` (`row[4]` BOM ID, `row[8]`/`row[5]` material, `row[9]`/`row[6]` ratio, `row[10]` UM `'PRK'`). Scans `'CONTROL DATA'` row 6 only; `PARTE X` parser (`line 287`) requires `val_s.isdigit() and len(val_s) >= 7`. | `col_part` matched (`lines 247-248`) but never read (`lines 310-325`); alphanumeric SAP `MATNR`s silently skipped in `PARTE X`; `um` defaults to `'PRK'` (`line 314`), `curve` to `12` (`line 316`). |
| `src/parsers.py` — `parse_cierre_recepcion` | 345–679 | Hardcodes `col_map` (`lines 365-389`): `po_43=1 (A)`, `mat=5 (E)`, `ped43_u=6 (F)`, `ped41_u=7 (G)`, `rec_u=8 (H)`, `mov101=9 (I)`, `falt=14 (N)`, `dan=15 (O)`, `sob=16 (P)`, `mov641=21 (U)`, `desc=30 (AD)`, `unit_cost=32 (AF)`, `as_bultos=45`, `at_unidades=46`, `aw_bultos=49`, `ax_unidades=50`, `ay_empaque=51`, `curve=54 (BB)`, `um=55 (BC)`, `ped43_prk=56 (BD)`, `ped41_prk=57 (BE)`. | Row 12 scan (`lines 392-436`) restricts `'mov-101'` to `c < 20` (conflicts with macro `Col BF = 58`) and lacks header entries for `cat`, `mov641`, `desc`, `unit_cost`. Empty `UM` defaults to `"PRK"` (`line 477`); empty/invalid `curve` defaults to `12` (`line 461`). Child `c_rec` is synthesized (`c_theo` if `"LIMPIO"` else `0`, `line 619`). |
| `src/validator.py` — `validate_item_integrity`, `validate_all_items` | 10–108 | Checks: `CURVA_DESCUADRADA` (cites `"SAP transacción STPO/MAST"`, `line 33`), `BOM_FALTANTE` (cites `"tabla STPO"`, `line 43`), `POSIBLE_CONFUSION_UDM` (`rec > inv * 2.5`, `lines 47-54`), `PRECIO_CERO_EN_RECLAMO` (cites `ME80FN`, `line 63`), `ESTILO_PROVEEDOR_VACIO` (cites `EINA`, `line 73`), `MATERIAL_EN_MULTIPLES_PO` (`lines 86-101`). | **Never rendered in UI or CLI** (`app.py:298,305,576` stores/unpacks `validation` but never displays it; `cli.py` never imports it). Calls `STPO/MAST` a "transacción" at `line 33`; misses `PRK`-as-`UN` under-entry (`81` vs `1,728`); `MATERIAL_EN_MULTIPLES_PO` triggers on duplicate identical/blank POs without using `parse_multi_po`. |
| `src/heuristic_parser.py` — `COLUMN_SYNONYMS`, `detect_headers_in_sheet` | 1–72 | Maps `"unit_cost"` to `["costo", ..., "vkp2"]` (`line 18`), `"po"` to `"doc. compras"` (`line 17`), `"description"` to `"texto breve"` (`line 13`), `"bom_id"` to `"lista de materiales"` (`line 19`). | **Dead code** (zero imports anywhere in repo). Semantic hazard: `VKP2` is IS-Retail sales price calculation, not import unit cost. Lacks `"idnlf"` and `"prk"`. |
| `src/sanitizer.py` — `clean_material_code`, `parse_multi_po` | 25–59, 127–138 | Converts scientific notation (`lines 41-46`) and strips SAP 18-char leading zeros if `s.isdigit() and len(s) > 10 and s.startswith("000")` (`lines 55-57`). | `clean_material_code` preserves `-`/`_` (`line 53`), fails on 8–10 digit zero-padded codes or `00...` prefixes, strips valid >10-digit leading-zero EANs, and corrupts alphanumeric `"10E5"` SKUs as float (`lines 41-46`). `parse_multi_po` splits on `[,;/]+` (`line 136`), omitting `-` despite docstring example `'4300006764-6804-6805'` (`line 129`), and is dead code outside tests. |
| `src/models.py` & `src/engine.py` | `models.py:29-94`, `engine.py:7-91` | Defines `PrepackItem.sap_mov_101` (`line 43`), `sap_mov_641` (`line 44`), `revuelto` (`line 49`), and `ReconciliationSummary`. | `sap_mov_101` is never reconciled against `received_real_units`; `sap_mov_641` and `revuelto` are completely unused; `ReconciliationSummary` omits SAP movement totals and validation counts. |
| `src/i18n.py` | 159–166, 224, 362–363, 396–405 | Defines `kpi_erp_balance` (`"Balance Neto (101 - 102)"`, `lines 160-161`) and references `ME80FN, MB51, SE16N`, `Z-Carga`, `ZDESCARGA`, `PO 43`, `PED 41`, `DUA`. | `kpi_erp_balance` is **never called** in `app.py`. |
| `app.py`, `cli.py`, `generar_cierre_oficial.py`, `tools/operational_stress_test.py` | `app.py:298,576,1394-1773`; `cli.py:1-101`; `generar_cierre_oficial.py:124-125`; `tools/operational_stress_test.py:29,233` | `app.py` implements `blind_setup` (`PAUSA_SIN_CONTROL_CALIDAD`), `tab_clipboard` (SAP GUI paste cleaning), `tab_excel_export` (`.xlsm` patching, `Bultos` sync, self-healing formulas `H`/`AU`/`AZ`, `ZDESCARGA` strips), `tab_checklist`, `tab_wiki`. | `tab_clipboard` has no export/copy-back button and doesn't strip SAP ALV pipe borders `\|---\|`. `generar_cierre_oficial.py:124-125` hardcodes `"0 UN (Cuadrado)"` / `"$0.00 USD (LIMPIO)"` output text. |

### 2.3 External Template, VBA Macro & Data Corpus Evidence

1. **QC Workbooks (`sandbox_benchmark\*\unmodified_originals\*Control*Calidad*.xlsm`):**
   - 14 of 15 share the 9-sheet layout: `['CONTROL DATA', 'BASE DE DATOS', 'DIST. PARTE A', 'DIST PART A004', 'DIST- ESPECIALES', 'ESQUEMA A001', 'ESQUEMA ALISTO PARTE B', 'ESQUEMA ALISTO PARTE C ', 'ESQUEMA A004']`; `imp 103 control calidad.xlsm` has 11 sheets (adds `Hoja3`, `Hoja1`).
   - `BASE DE DATOS` Row 1 documents the exact 7-step analyst SAP extraction pipeline:
     - Step 1 (`A1`): `'1.ME80FN (Consolidar cantidad dePRK y UN)'`
     - Step 2 (`E1`): `' 2.Transacciones MAST y STPO (Ordenar por lista Primero...)'`
     - Step 3 (`I1`): `'3. Consolidar Materiales y variantes sueltas...'` (`UN` / `PRK` + curve)
     - Step 4 (`N1`): `'4.Busqueda de Descripción (... se coloca en la MM60)'` *(absent from all repo docs/code)*
     - Step 5 (`Q1`): `'5.Busqueda de Mat. Proveedor. ... SE16N, en la Tabla "EINA"...'`
     - Step 6 (`T1`): `'6. Busqueda de Precios ... en la transacción VKP2'`
     - Step 7 (`W1`): `'7.UBICACIÓN MONITOR '` *(present across all 15 QC workbooks)*
   - `CONTROL DATA` exhibits `Tarima ` column shift in Col B (`11/15` cases incl. `IMP-100`, `207` vs `CATEGORIA ` in `4/15` cases incl. `IMP-182`), carries `NACIONALIZACIÓN: 410000xxxx`, `Precios VKP2` (Col `P` vs `O`), `ALISTO / A001 / A004 / Especial` warehouse columns, and a UEN category list (`MULTIPLE` + up to 16 categories, rows `6–22`) that shifts across Col `AO` (`IMP-102`, `182`, `207`), Col `AP` (`10/15` cases incl. `IMP-100`), and Col `AQ` (`IMP-103`, `200`, shifted right alongside `TOTAL DE UNIDADES` `AI` / `TOTAL PRK` `AJ` shifts).
   - `ESQUEMA A001` identifies `'CENTRO DE DISTRIBUCIÓN CADENA DE TIENDAS EKONO'` and `'Almacén A001'`; `ESQUEMA A004` references `'Almacén A004'`; `6/15` cases (`IMP-102`, `103`, `180` [spelled `'ACOMNDICIONAMIENTO'`], `187`, `207 (NCR)`, `225`) note splits with `ACONDICIONAMIENTO`.

2. **Cierre Workbooks (`sandbox_benchmark\*\unmodified_originals\*.xlsm`) & VBA (`xl/vbaProject.bin`):**
   - 14 of 15 share 13 sheets (`Check list RECEPCION`, `Check list DESPACHO`, `PORTADA`, `DATOS GENERALES`, `CIERRE RECEP`, `Bultos`, `PARA N.C RECEP.`, `ERASER`, `CIERRE DESP`, `Hoja1`, `PARA N.C DESP`, `STOCK CONTROL DE INVENTARIO`, `PROV Y CATEG`); `IMP-200.xlsm` is a 14-sheet `-A` split variant (drops `ERASER`, adds `641-A` and `PED-45-A`, and renames the 3 dispatch sheets `CIERRE DESP-A`, `PARA N.C DESP-A`, `STOCK CONTROL DE INVENTARIO-A`).
   - `Check list DESPACHO` (`IMP-100.xlsm`) explicitly names: `MB51` (Mov `641`, `A12`), `MB5T Transitos` (`A16`), `Respaldo SAP EWM` (`A17`), `MB52` (`A18`), and `Lista de pedidos 45 de distribucion` (`A19`).
   - `Hoja1` (confirmed across all 15 benchmark Cierre workbooks, e.g., `IMP-182.xlsm`) is a payroll deduction authorization (`'Autorización de Deducción'`) containing `G18 = 'MV :'`, `H18 = 998`, `G19 = 'DSG:'`, `G20 = 'DOC SAP:'`.
   - `STOCK CONTROL DE INVENTARIO` (`IMP-182.xlsm`) tracks `C001` storage locations (`B22:K22`): `A001`, `A002`, `A004`, `03`, `04`, `05-06`, `07`, `8`, `09`, `99`.
   - `PED-45-A` (`IMP-200.xlsm`) maps 43 store/channel destinations (`C101` `EKONO_AVENIDA4` through `C155` `EKONO_ESCAZU`, plus `EC02` `EKONO_ECOMMERCE`) to `Pedido Almacén` numbers `4500283933`–`4500283975`.
   - `xl/vbaProject.bin` (`Módulo1`, `477` non-`Attribute` lines [`480` raw], identical LF-normalized UTF-8 MD5 `90a614f461f2` across `IMP-100`, `IMP-182`, `IMP-207 NCR`, `IMP-200`, and `data\Macro para cierre hannia.xlsm`): implements `Format43` (`ME80FN` headers incl. `Centro`, `Valor Neto`, `Grupo Art.`, `UMP`), `Mueve43` (enforces `Centro = "C001"`, deletes rows `..1011`), `Mueve41`, `MueveMatProv` (`SE16N-EINA`: `MATNR`, `IDNLF`, `LMEIN`), `MueveCurva` (`SE16N-MAST` + `SE16N-STPO` join on `STLNR`, sums `MENGE`), `Format51`/`Mueve51` (`MB51 101-102` -> `CIERRE RECEP!BF`; `MB51 641-642` -> `CIERRE RECEP!U`), plus `NoDuplicados` and `CalcDesg` (`CIERRE RECEP!AY`).

3. **Real Data Corpus (`data/`, `591` files across `71` `IMP-*` folders + `2` root `.xlsm` files = `593` total):**
   - `@ekonocr.com` emails confirm PO `430000xxxx` + contract -> `Pedido Nacio 410000xxxx` nationalization workflow, receiving closures for `almacén a001`, and zero mentions of `VL31N`, `EWM`, `ECC`, `S/4HANA`, `18...`, or `80...` in email text.
   - `Portada 100.xlsx` confirms `CADENA DE TIENDAS EKONO`, `SOCIEDAD: IM 01`, `ALMACEN FISCAL CLADISA`.
   - 4 SAP Credit Note PDFs (`IMP-105`, `IMP-82`, `IMP-205`, `IMP-207`) generated by `Adobe LiveCycle Designer ES 9.8` confirm `IMPORTADORA ALMACEMA S.A`, NCR doc numbers `4000001153/54/57/1210`, 10-digit SAP vendor IDs `1100000003` / `1100000577`, 7-digit `Cod.SAP` (`1455918`, `5033283`), and 16-digit barcodes (`2000000000156838`).

### 2.4 Consolidated Catalog of Currently Observed SAP Objects

- **T-Codes Observed:** `ME80FN`, `SE16N`, `MB51`, `MB52`, `MB5T` *(corpus/template only)*, `MM60` *(corpus/template only)*, `VKP2` *(corpus/template + dead code)*, `VL31N` *(diagram-only in SOP)*, `RFUI` *(diagram-only in SOP)*.
- **Tables & Fields Observed:** `EINA` (`MATNR`, `IDNLF`, `LMEIN`), `MAST` (`MATNR`, `STLNR`), `STPO` (`STLNR`, `MENGE`).
- **Prefixes & IDs Observed:** `430000xxxx` (Import PO), `410000xxxx` (Nationalization Order), `4500xxxxxx` (Store Distribution Order, e.g., `C101`–`C155` + `EC02`), `400000xxxx` (Credit Note / NCR), `1100000xxx` (SAP Vendor `LIFNR`), 7-digit/10-digit `MATNR` (`1455103`, `1455102001`), `18...` / `80...` *(SOP dictionary only — 0 occurrences in corpus)*.
- **Movement Types Observed:** `101`, `102`, `641`, `642`, `643`, `501`, `998` *(deduction form `Hoja1` only)*.
- **Organizational & Z-Objects Observed:** Company Code `IM 01`, DC Plant `C001` (`A001`, `A002`, `A004`, `03`–`99`), Store/Channel Plants `C101`–`C155` & `EC02` (`C101+`), `Z-Carga`, `ZDESCARGA`/`ZDECARGA`, `Z51 -Z52`, `"7.UBICACIÓN MONITOR"`.

---

## 3. Explicit Epistemic Classification (Strict 6-Way Separation)

Every SAP-related concept found across the repository, templates, macros, and corpus is classified below into six mutually exclusive categories per **R5** and **R8**.

### 3.1 Confirmed in Repo / Templates / Corpus

| Item | Verified Scope | Primary Evidence |
|---|---|---|
| Classic SAP GUI copy-paste into Excel (`ME80FN`, `SE16N`, `MB51`, `MB52`) | Actively executed by analysts | `docs/SOP_RECEPCION_SAP.md:60-78`; `Módulo1` VBA (`lines 10-348`); `Check list RECEPCION/DESPACHO` |
| `MM60` (Material descriptions), `VKP2` (Price lookup in QC), `MB5T` (In-transit check) | Actively used in QC/Dispatch templates | `BASE DE DATOS!N1,T1`; `CONTROL DATA` (`Precios VKP2`); `Check list DESPACHO!A16` |
| Tables `EINA` (`MATNR`, `IDNLF`, `LMEIN`), `MAST` (`MATNR`, `STLNR`), `STPO` (`STLNR`, `MENGE`) | Queried via `SE16N` and joined in `Módulo1` | `Módulo1` (`Mueve41`, `MueveMatProv`, `MueveCurva`); `BASE DE DATOS!E1,Q1` |
| Prefixes `430000xxxx` (Import PO), `410000xxxx` (Nationalization), `4500xxxxxx` (Store Order), `400000xxxx` (NCR), `1100000xxx` (Vendor) | Observed in real workbooks, emails, and PDFs | `DATOS GENERALES!C8:C9`; `PED-45-A`; `PARA N.C RECEP.!D6`; 4 NCR PDFs (`IMP-82/105/205/207`); `PROV Y CATEG` |
| Movements `101`, `102`, `641`, `642` in `MB51` extracts; `501`, `643`, `998` template references | `101-102` & `641-642` parsed by `Mueve51`; `501`/`643`/`998` in sheet headers | `Módulo1` (`Format51`/`Mueve51`); `CIERRE RECEP` (`Q`, `Z`); `CIERRE DESP` (`AG`); `Hoja1!G18:H18` (`'MV :' / 998`) |
| `Adobe LiveCycle Designer ES 9.8` PDF forms for `400000xxxx` Credit Notes | Confirmed PDF producer metadata & layout | `IMP-105`, `IMP-82`, `IMP-205`, `IMP-207` NCR PDFs |
| Python codebase only parses pre-exported Phase 1 Excel (`CIERRE RECEP` + QC) | `validator.py` unrendered; `heuristic_parser.py` dead; `sap_mov_101`/`641` unreconciled | `src/parsers.py:345-679`; `app.py:298,576`; `src/models.py:43-44` |

### 3.2 Company-Specific Workflows & Customizations

These conventions belong to **Importadora Almacema S.A. (`IM 01`) / Cadena de Tiendas Ekono** and **must NOT** be generalized as standard SAP behavior:

| Item | Company-Specific Meaning / Rule | Evidence Citation |
|---|---|---|
| **Org Structure (`IM 01`, `C001`, `C101–C155` / `EC02`, `A001/A002/A004/03-99`)** | `IM 01` = Importadora Almacema; `C001` = DC; `A001` = main/central/picking; `A004` = replenishment (`resurtido`); `03..99` = sub-stores/zones; `C101–C155` & `EC02` (`C101+`) = retail store & e-commerce destinations (`EKONO_AVENIDA4`..`EKONO_ESCAZU`, `EKONO_ECOMMERCE`). | `Portada 100.xlsx`; `Mueve43`/`Mueve41`; `STOCK CONTROL DE INVENTARIO!B22:K22`; `PED-45-A`; `ESQUEMA A001`/`A004` |
| **Local Document Type Ranges (`43`, `41`, `45`, `40`)** | `43...` = International supplier PO; `41...` = Nationalization / DUA-backed commercial order queried in `ME80FN`; `45...` = DC-to-store distribution order (`Pedido Almacén`); `40...` = Supplier Credit Note (NCR). | `SOP_RECEPCION_SAP.md:45-48`; `Módulo1` (`Pedidos_43`, `Pedidos_41`); `PED-45-A`; NCR PDFs |
| **Custom Z-Designations (`Z-Carga`, `ZDESCARGA`/`ZDECARGA`, `Z51 -Z52`, `7.UBICACIÓN MONITOR`)** | Dock tally/bin upload conventions, macro strips, and monitor location check (`BASE DE DATOS` step 7 across all 15 QC workbooks). Exact SAP transaction/report status unverified. | `Bultos!C2:D2,J2`; `CIERRE RECEP!AS11`; `CIERRE DESP!D11,AX11,BQ12`; `BASE DE DATOS!W1` |
| **Custom UoM `PRK` & 3-Level Segregation** | `PRK` = assorted prepack bag/bundle; strict separation of Level 1 Cartons (`AS`/`AW`), Level 2 Prepacks (`PRK`/`BD`/`BE`), and Level 3 Units (`F`/`G`/`H`/`AT`/`AX`). | `SOP_RECEPCION_SAP.md:64`; `CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md:7-45`; `MueveMatProv` |
| **`Bulto Cero` & Split Containers (`Parte A/B/C/D`)** | Zero cartons/units (`AS=AT=AW=AX=0.0`) when `U_recibidas == 0` or `Parte != A`; `PAUSA_SIN_CONTROL_CALIDAD` mode locks NCR at `$0.00` without QC. | `CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md:160-174`; `app.py:431-455` |
| **Official `.xlsm` Workbooks, `Módulo1` VBA & 7-Doc Binder** | 9/11-sheet QC + 13/14-sheet Cierre workbooks, `Módulo1` 477 non-`Attribute`-line (480 raw) ETL macro, dual carton view modes (`Muelle` vs `Estricto Aduanero`), `MV 998` payroll deduction form (`Hoja1`), and 7-document physical binder. | `sandbox_benchmark/*`; `xl/vbaProject.bin`; `SOP_RECEPCION_SAP.md:112-120` |
| **Local `643` Usage (`"En proceso de facturar"`)** | Uses `643` in `CIERRE RECEP!Z` and `CIERRE DESP!AG` to represent units pending store billing/invoicing. | `AUDITORIA_DESPACHO_MACRO.md:86`; `CIERRE RECEP!Z12` |

### 3.3 Generic SAP Knowledge (ECC 6.0 / Standard MM-LE-WM-Retail Baseline)

These concepts apply broadly across SAP ECC 6.0 (where version differences vs S/4HANA are secondary or clearly noted):
- **Standard Reporting & Query T-Codes:** `ME80FN` (General Evaluation of Purchasing Documents via `EKKO`/`EKPO`), `MB51` (Material Document List via `MKPF`/`MSEG`), `MB52` (Warehouse Stocks of Material via `MARD`/`MARC`/`MBEW`), `MB5T` (Stock in Transit CC via `MARC-TRAME`/`EKPO`), `MM60` (Materials List via `MARA`/`MAKT`/`MBEW`), `SE16N` (General Table Display), `VKP2` (IS-Retail Display Sales Price Calculation — **retail selling price, not purchasing cost**).
- **Standard Purchasing & Master Data Tables:** `EINA` (`MATNR`, `IDNLF` vendor material, `LMEIN` base UoM), `MAST` (`MATNR`, `WERKS`, `STLAN`, `STLNR`, `STLAL`), `STPO` (`STLNR`, `IDNRK` component, `MENGE` component qty), `MARA` (`MATNR` 18-char `CHAR(18)` with `CONVERSION_EXIT_MATN1_INPUT` leading zeros).
- **Standard Movement Types:** `101` (Goods Receipt for PO into unrestricted/QI/blocked stock), `102` (Reversal of `101` GR), `641` (Goods Issue for intra-company two-step STO into stock in transit), `642` (Reversal of `641`), `501` (Receipt without PO into unrestricted stock), `643` (Goods Issue for **cross-company** two-step STO).

### 3.4 Conflicting & Technically Unverified Items

Every contradiction identified across the repository docs, Python code, and external corpus is enumerated below:

| # | Conflicting / Unverified Claim | Evidence A (Claim) | Evidence B (Counter-Evidence / SAP Conflict) | Risk if Unresolved |
|---|---|---|---|---|
| 1 | **`SAP S/4HANA` + `SAP EWM` vs `SAP ECC 6.0` + `"100% ERP-Agnostic"`** | `docs/SOP_RECEPCION_SAP.md:1,30-31` (`S/4HANA`, `SAP EWM`), `Check list DESPACHO!A17` (`"Respaldo SAP EWM"`), `README.md:165` (`"100% Agnostic"`). | `ORIGINAL_REQUEST.md:28` expects **SAP ECC 6.0**; all 7 docs and macros use classic GUI (`ME80FN`, `SE16N`, `MB51`, `MB52`, `MM60`), storage locations `A001`/`A004` (`C001`), and zero Fiori or `/SCWM/*` t-codes; codebase hardcodes `43`/`41`/`101`/`641`. | Authoring wrong KB procedures (e.g., S/4HANA `ACDOCA`/`MATDOC` or EWM `/SCWM/PRDI` instead of ECC 6.0 `MKPF`/`MSEG` & `LE-WM`). |
| 2 | **`VL31N / RFUI` Flowchart Anomaly & `18...`/`80...` Zero-Corpus Presence** | `docs/SOP_RECEPCION_SAP.md:30-31,49-50` shows `VL31N / RFUI` in `S4` (post-audit), `18...` Inbound Delivery, `80...` Handling Unit. | `VL31N` (ECC LE) vs `RFUI` (`/SCWM/RFUI` EWM) conflate two modules; `S4` occurs *after* `MB51 101-102` (`S1`) and 3-way audit (`S3`); **0 occurrences** of `VL31N`, `RFUI`, `18...`, or `80...` in 15 benchmark cases, VBA, emails, or PDFs. Moreover, `80...` (`0080000000+`) is SAP's default **Outbound Delivery (`LIKP`)** range, not HU (`VEKP`). | Analysts/AI giving non-existent inbound delivery or misclassified HU instructions. |
| 3 | **`41...` "Factura de Proveedor / Pedido Nac." Queried in `ME80FN`** | `SOP_RECEPCION_SAP.md:46,62`, `Módulo1` (`Pedidos_41`), `DATOS GENERALES!C9` query `41...` in `ME80FN` (`Documento de compras`). | In standard SAP MM, `ME80FN` queries purchasing docs (`EKKO`/`EKPO`), whereas FI/MM invoices live in `MIRO`/`RBKP` (`51...`). Likely a custom PO document type (`EKKO-BSART`) or customs-transfer PO in ECC. | Confusion between customs nationalization PO (`41...`) and vendor logistic invoice (`RBKP`), breaking FI-MM/GR-IR guidance. |
| 4 | **`VKP2` Mapped to Import `unit_cost`** | `src/heuristic_parser.py:18` includes `"vkp2"` as a synonym for `"unit_cost"`; `QC!BASE DE DATOS!T1` & `CONTROL DATA` fetch `VKP2`. | In SAP IS-Retail, `VKP2` calculates **retail selling price** (POS/store price), NOT purchase order net cost (`EKPO-NETPR`), landed DUA CIF cost, or moving average cost (`MBEW-VERPR`). | Catastrophic overstatement of NCR vendor claims if `heuristic_parser.py` is ever wired in using `VKP2` prices. |
| 5 | **`Mov 643` ("Unidades en proceso de facturación")** | `AUDITORIA_DESPACHO_MACRO.md:86`, `CIERRE RECEP!Z12`, `CIERRE DESP!AG12` label `643` as units in billing/invoicing process. | Standard SAP `643` is **Cross-Company Two-Step STO Goods Issue** (between distinct company codes, e.g., `IM 01` -> another legal entity), whereas `641` is intra-company. Unverified whether Ekono stores sit in a separate company code (`BUKRS`) or if `643` is customized. | Incorrect accounting/stock-in-transit interpretation during dispatch reconciliation. |
| 6 | **`CIERRE RECEP` `Mov 101` Column Conflict (`Col BF` vs `Col I`)** | `AUDITORIA_DESPACHO_MACRO.md:43` and `Mueve51` VBA (`Módulo1:289-348`) write `MB51 101-102` into **Col `BF` (58)** (PRK/base UM), while `Col I (9)` holds `=BF*BB` (Units). | `src/parsers.py:373,409-410` restricts `mov-101` scan to `c < 20` and defaults to **Col `I` (9)**, matching `app.py:1466` unit-level display (`CANTIDAD INGRESO MOV-101`), while `Col BF` is ignored. | Misalignment if template columns shift or if `BF` (PRK) vs `I` (UN) semantics are confused. |
| 7 | **`CIERRE DESP!AC13 = +AB13-J13` (Units vs PRKs & Label Conflict)** | `AUDITORIA_DESPACHO_MACRO.md:64,67-68` defines `J` as `RECIBIDO REAL & FISICO PRK (=P13/O13)`, `AB` as `MOV-641` (`='CIERRE RECEP'!U13`, units), and `AC` as `DIFERENCIA MOV 101 VRS MOV 641` (`=+AB13-J13`). | Subtracts `J13` (PRKs) from `AB13` (Units) and labels `J13` as `MOV 101` instead of `I13` or `BF13`. Requires verification in actual `.xlsm` `CIERRE DESP` formulas. | Corrupted Phase 2 dispatch balance if copied blindly into Python `engine.py`. |
| 8 | **`STPO/MAST` Called a "Transacción" in `validator.py:33`** | `src/validator.py:33` prints `"transacción STPO/MAST"`. | `MAST` and `STPO` are **database tables** queried via `SE16N` (or displayed via `CS03`), not transactions. (`validator.py:43` correctly calls `STPO` a table). | Misleads junior analysts and pollutes AI/RAG terminology. |
| 9 | **Doc Version & Test Count Drift** | `WIKI.md:8` (`v1.0.0`), `WIKI.md:182` (`v1.1.0`), `RELEASE_NOTES_DEMO.md:1`/`README.md:8` (`v1.2.0`), `README.md:153` (`22/22 Passing`). | `PROJECT.md:17` confirms 27 passing test items across 4 test files. | Minor docs drift requiring KB versioning discipline. |

### 3.5 Items Requiring Official SAP References

The user-provided official SAP Help Portal anchor for SAP ERP 6.0 / ECC 6.0 (`version=6.00.29`):
- **`https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29`**

This URL serves as the **primary official SAP ERP 6.0 / ECC 6.0 documentation anchor** for verifying standard MM, LE-WM, SD, FI-MM, and IS-Retail semantics during Track D execution. Consistent with **R1** and **R8**, **no specific sub-page titles, sub-URLs, or unverified page contents are fabricated here**; individual topic links and exact section citations under `version=6.00.29` must be retrieved, read, and recorded per KB entry during Track D execution alongside ABAP Data Dictionary (`SE11`) specifications for:
1. **IS-Retail Article & Prepack Modeling:** `MARA-ATTYP` (Single `00`, Generic `01`, Variant `02`, Sales Set `10`, Prepack `11`, Display `12`), generic/variant linkage (`MARA-SATNR`), and prepack BOM storage (`MAST`/`STPO`, `STLAL`, `WERKS`) vs Retail allocation/prepack transactions (`MM43`, `WSOH`, `WAK3`).
2. **IS-Retail Pricing (`VKP2` / `VKP5`):** Condition tables, sales price calculation schema, and explicit proof separating `VKP2` retail price from `EKKO`/`EKPO` purchasing net price and `MBEW` valuation price.
3. **Inventory Management Movements & Stock Determination (`OMJJ` / `MSEG` / `MKPF`):** Exact posting keys, stock-type effects (`LABST`, `INSME`, `SPEME`, `TRAME`), and reversal behavior for `101/102`, `103/105`, `122/123`, `161/162`, `311/312`, `501/502`, `641/642`, `643/644`, `701/702`.
4. **Material Number Conversion Exit & UoM Conversions:** `CONVERSION_EXIT_MATN1_INPUT` / `OUTPUT` (`CHAR(18)` zero-padding behavior for pure numeric vs alphanumeric `MATNR`) and `MARM` numerator/denominator (`UMREZ`/`UMREN`) conversion between `PRK` (`EINA-LMEIN` / `EKPO-MEINS`) and base unit `UN`/`ST`.
5. **FI-MM Valuation & GR/IR Clearing (`OBYC` / `WRX` / `BSX`):** Moving Average Price (`MBEW-VERPR`, price control `V`) vs Standard Price (`MBEW-STPRS`, `S`), customs/freight planned delivery costs, and `MB5S`/`MR11` GR/IR clearing mechanics.
6. **SAP Standard Authorization Objects:** `S_TCODE`, `M_MSEG_BWA`, `M_MSEG_WMB`, `M_BEST_WRK`, `M_BEST_EKG`, `S_TABU_DIS`, `S_TABU_NAM`, `S_TABU_CLI` (display `ACTVT=03` vs change/post activities).

### 3.6 Items Requiring Validation in the User's Actual SAP ECC 6.0 Environment

The following concrete, **strictly read-only** checks must be run by an authorized analyst, functional consultant, or Basis/MM lead in the live Ekono/Almacema SAP environment—and cross-checked against the official SAP ERP 6.0 (`version=6.00.29`) Help Portal anchor (`https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29`)—to separate standard ECC 6.0 behavior from `S/4HANA`/`EWM` claims (`SOP_RECEPCION_SAP.md:1,30-31`) and company-specific customizations before unverified KB entries are promoted to `confirmed`:

| # | Validation Target | Exact Read-Only SAP GUI Check / Table | Question Answered |
|---|---|---|---|
| 1 | **SAP Release, Enhancement Pack & IS-Retail Switch** | `System -> Status` (`SAP System data` -> Component version, e.g. `SAP_APPL 60x`, `IS-RETAIL` / `EA-RETAIL`); `SFW5` (display only) | Confirms exact SAP ECC 6.0 (`version=6.00.29` baseline) vs S/4HANA release and whether IS-Retail is active. |
| 2 | **Document Types `43`, `41`, `45`** | `SE16N` (Display) on `T161` / `T161T` and `EKKO` (`EBELN` in `43*`, `41*`, `45*` -> `BSART`, `BSTYP`, `BUKRS`, `EKORG`) | Confirms whether `41...` is a custom PO type (`BSART`) under `IM 01` and how `45...` STOs (`UB`/`NB`) deviate from standard ECC 6.0 (`6.00.29`). |
| 3 | **Number Ranges `18...`, `80...`, `40...`** | `SNRO` / `NRIV` (Display) for objects `RV_BELEG` (`LIKP` deliveries), `HU_VEKP` (Handling Units), `RE_BELEG` / `BELNR` (`RBKP` / `BKPF` / `40...` NCRs) | Confirms whether `18...` and `80...` are used at all, and whether `40...` NCRs are MM credit memos (`RBKP`), FI docs (`BKPF`), or SD/custom forms. |
| 4 | **Movement Types `101`, `102`, `501`, `641`, `642`, `643`, `998`** | `SE16N` on `T156` / `T156T` and `OMJJ` (Display only) | Confirms standard ECC 6.0 (`6.00.29`) definitions vs company-specific customization of `643` ("en proceso de facturar") and `998` (`Hoja1` payroll/inventory deduction). |
| 5 | **Warehouse Management (`LE-WM` vs `EWM`)** | `SE16N` on `T001L` (`WERKS='C001'`), `T320` (Storage Loc to WM Warehouse `LGNUM`), and `/SCWM/TMAPSTLOC` (if EWM exists) | Resolves `SOP_RECEPCION_SAP.md` / `Check list DESPACHO!A17` EWM claim vs ECC 6.0 `LE-WM` (`C001`: `A001`, `A002`, `A004`, `03`–`99`). |
| 6 | **Company Code Topology (`IM 01` & Stores `C101+`)** | `SE16N` on `T001W` (`WERKS` = `C001`, `C101..`) and `T001K` (`BWKEY` -> `BUKRS`) | Determines if stores `C101+` are in `IM 01` (intra-company `641`) or separate company codes (cross-company `643`). |
| 7 | **Custom Z-Objects (`Z-Carga`, `ZDESCARGA`, `Z51-Z52`, Monitor)** | `SE93` (Display) / `TSTC` (`Z*`) and `SE11`/`SE38` (Display only) | Identifies exact Z-transaction/program names behind `"7.UBICACIÓN MONITOR"`, `Z-Carga`, `ZDESCARGA`, and `Z51 -Z52`. |
| 8 | **Prepack BOMs & `MARM` UoM** | `MM03` / `CS03` and `SE16N` on `MARA` (`ATTYP`, `SATNR`), `MARM` (`MEINH='PRK'`, `UMREZ`, `UMREN`), `MAST`, `STPO` for sample `1455103` | Confirms whether prepack curves are strictly `MAST`/`STPO` (`MENGE`) and how `PRK`-to-`UN` conversion is maintained in `MARM`. |
| 9 | **Analyst Authorizations** | `SU53` (after test read), `PFCG` / `AGR_1251` (Display by Basis/Security) | Inventories exact read-only roles granted to Inventory Analysts vs write roles (`MIGO`, `MIRO`, `VL31N`). |

---

## 4. Missing SAP Knowledge & Expansion Scope (Bounded, Non-Speculative)

All transactions and tables introduced below that were **not** directly observed in the repository/corpus are explicitly tagged **`[Generic ECC 6.0 Candidate — Validate in Env]`** vs **`[Confirmed in Use]`**. Scope is strictly bounded to Ekono/Almacema's 3-way import reception, distribution, credit note, and reconciliation workflows.

### 4.1 Relevant SAP Modules & Architectural Boundaries

1. **MM-PUR & MM-IM (Materials Management — Purchasing & Inventory Management):** Core module confirmed in active use (`ME80FN`, `EINA`, `MB51`, `MB52`, `MB5T`, `101/102/501/641/642/643`). Governs POs (`43`, `41`), STOs (`45`), material documents (`MKPF`/`MSEG`), and stock balances (`MARD`/`MARC`).
2. **IS-Retail / Master Data (`LO-MD` / `W*`):** Partially confirmed via `VKP2`, `PRK` prepacks, 7-digit generic/prepack vs 10-digit variant `MATNR`s (`1455103` vs `1455102001`), and `MAST`/`STPO` BOMs. Requires validation whether full IS-Retail (`MM43`) or classic `MM03`/`CS03` is active.
3. **LE-WM vs EWM (Warehouse Management Boundary in ECC 6.0):** In ECC 6.0, standard warehouse management is **LE-WM** (`LT03`, `LT12`, `LS24`, `LX02`, `LM00`/`LM01` RF), whereas **SAP EWM** is normally a **separate decentralized SCM/EWM server** (communicating via qRFC deliveries/HU queues) unless running S/4HANA embedded EWM. Because `Check list DESPACHO!A17` says `"Respaldo SAP EWM"` while `C001` uses `A001`/`A004` storage locations and `MB52`, the exact `LE-WM` vs decentralized `EWM` boundary must be validated via `T320` before documenting putaway/picking procedures.
4. **SD / LE-SHP (Logistics Execution — Shipping & Deliveries) & MM-IV (Logistics Invoice Verification):** Required to bridge the gap between physical dock/dispatch discrepancies and financial settlement (`VL10B`/`VL02N` STO deliveries, `MIRO`/`MIR4` invoice/credit memo posting, `40...` NCR generation).
5. **FI-MM (Valuation & Account Determination):** Required to explain how `43`/`41`/`101`/`102`/`40` affect `MBEW` inventory value and the `WRX` GR/IR clearing account.

### 4.2 Transactions & Tables Expansion Scope

| Domain | Object | Kind | Status Tag | Role in Expanded Knowledge Base |
|---|---|---|---|---|
| **Purchasing / PO / STO** | `ME80FN` | T-Code | `[Confirmed in Use]` | Mass ALV extraction of `43...` / `41...` / `45...` orders (`UMP`, qty, net value). |
| | `ME23N`, `ME13` | T-Code | `[Generic ECC 6.0 Candidate — Validate in Env]` | Single PO/STO display (`PO History` tab) & Purchasing Info Record display. |
| | `EINA` | Table | `[Confirmed in Use]` | Vendor material (`IDNLF`), base/order UoM (`LMEIN`), `MATNR`. |
| | `EKKO`, `EKPO`, `EKET`, `EKBE`, `EINE` | Table | `[Generic ECC 6.0 Candidate — Validate in Env]` | PO header (`BSART`, `LIFNR`, `WAERS`), items (`NETPR`, `MENGE`, `MEINS`, `WERKS`, `LGORT`), schedule lines, PO history (GR/IR links), info record org data. |
| **GR & Reversals** | `MB51` | T-Code | `[Confirmed in Use]` | Material document list (`101-102`, `641-642`, `501`, `643`). |
| | `MIGO`, `MB03` | T-Code | `[Generic ECC 6.0 Candidate — Validate in Env]` | Post/Display Goods Receipt, Goods Issue, Cancellation (`102`), Return Delivery (`122`), and material document inspection. |
| | `MI07` | T-Code | `[Generic ECC 6.0 Candidate — Validate in Env]` | Post physical inventory count differences (`701`/`702`) against `LABST`. |
| | `VL31N`/`VL32N`/`VL33N` | T-Code | `[Conflicting / Unverified — Validate in Env]` | Inbound delivery create/change/display (0 corpus evidence; verify if used). |
| | `MKPF`, `MSEG`, `LIKP`, `LIPS` | Table | `[Generic ECC 6.0 Candidate — Validate in Env]` | Material doc header/items (`BWART`, `BUDAT`, `MBLNR`, `MJAHR`, `EBELN`, `MENGE`, `SHKZG`) & delivery header/items. |
| **Material / UoM / Prepack** | `MM60`, `SE16N`, `VKP2` | T-Code | `[Confirmed in Use]` | Material list/descriptions (`MM60`), table queries (`SE16N`), IS-Retail sales price (`VKP2` — **sales price only**). |
| | `MAST`, `STPO` | Table | `[Confirmed in Use]` | BOM-to-material link (`STLNR`) and component ratios (`IDNRK`, `MENGE`). |
| | `MM03`, `MM43`, `CS03` | T-Code | `[Generic ECC 6.0 Candidate — Validate in Env]` | Display material master / Retail article / BOM structure. |
| | `MARA`, `MAKT`, `MARC`, `MARD`, `MARM`, `MEAN` | Table | `[Generic ECC 6.0 Candidate — Validate in Env]` | Material general (`ATTYP`, `SATNR`), descriptions, plant data, storage location stocks (`LABST`/`INSME`/`SPEME`), UoM conversions (`PRK`<->`UN`), EANs. |
| **Stock, Transit & WM** | `MB52`, `MB5T` | T-Code | `[Confirmed in Use]` | Warehouse stock (`MB52`) and stock in transit between plants (`MB5T`). |
| | `MMBE`, `VL10B`, `VL02N`, `VL06O` | T-Code | `[Generic ECC 6.0 Candidate — Validate in Env]` | Stock overview by status/storage loc (`MMBE`), STO delivery creation/outbound monitor (`VL10B`/`VL02N`/`VL06O`). |
| | `LT03`, `LT12`, `LS24`, `LS26`, `LX02`, `LM00`/`LM01` | T-Code | `[Generic ECC 6.0 Candidate — Validate in Env]` | ECC `LE-WM` transfer orders, bin stock, and RF menu (if `LE-WM` is active instead of EWM). |
| | `/SCWM/MON`, `/SCWM/PRDI`, `/SCWM/RFUI` | T-Code | `[Conflicting / Unverified — Validate in Env]` | SAP EWM equivalents (only applicable if EWM is confirmed live). |
| | `LQUA`, `LTBK`, `LTBP`, `VEKP`, `VEPO` | Table | `[Generic ECC 6.0 Candidate — Validate in Env]` | WM quants, transfer requirements, Handling Unit header/contents. |
| | `T320` | Table | `[Generic ECC 6.0 Candidate — Validate in Env]` | Assignment of Plant / Storage Location (`WERKS`/`LGORT`, e.g., `C001`) to `LE-WM` Warehouse Number (`LGNUM`). |
| | `902` | WM Storage Type | `[Generic ECC 6.0 Candidate — Validate in Env]` | Standard `LE-WM` interim Goods Receipt area holding quants post-`MIGO 101` before putaway (`LT03`/`LT12`). |
| **Invoice, NCR & Valuation** | `MIRO`, `MIR4`, `FB03`, `MB5S`, `MR11` | T-Code | `[Generic ECC 6.0 Candidate — Validate in Env]` | Logistics invoice/credit memo entry & display, FI accounting doc display, GR/IR balance list (`MB5S`), GR/IR clearing (`MR11`). |
| | `MBEW`, `RBKP`, `RSEG`, `BKPF`, `BSEG` | Table | `[Generic ECC 6.0 Candidate — Validate in Env]` | Material valuation (`VERPR`, `STPRS`, `VPRSV`), MM invoice header/items, FI accounting header/line items. |
| | `BSX`, `WRX`, `GBB` | Account Key (`OBYC`) | `[Generic ECC 6.0 Candidate — Validate in Env]` | FI-MM automatic account determination keys: Inventory (`BSX`), GR/IR Clearing (`WRX`), and Offsetting Entry (`GBB`, e.g. `501`). |
| **Audit & Auth** | `SU53`, `CDHDR`, `CDPOS`, `AUT10` | T-Code / Table | `[Generic ECC 6.0 Candidate — Validate in Env]` | Authorization check failure diagnostics and master/document change logs. |

### 4.3 End-to-End Document Flows

#### Flow A: Import Procurement, Nationalization, 3-Way Reception & NCR (`C001`)
```text
[1. Import PO 430000xxxx] (ME80FN / EKPO; Supplier 1100000xxx; Price EKPO-NETPR)
        |
        v  (Customs Nationalization + Costa Rica DUA; email @ekonocr.com)
[2. Nationalization Order 410000xxxx] (ME80FN / Ped. Nac.; CIF + DUA backing)
        |
        +--> [Optional / Unverified: Inbound Delivery 18... (VL31N) / HU 80...]
        |
        v  (Physical Dock Unload -> QC CONTROL DATA + Z-Carga/ZDESCARGA -> Posting)
[3. Goods Receipt in SAP: MIGO Mov 101 / Reversal 102] (MB51 -> MKPF/MSEG)
        |   Stock lands in C001 (A001 Main, A002, A004 Resurtido, 03..99)
        v
[4. ReconcileAI 3-Way Audit] (PO 43 vs Doc 41 vs QC Physical vs MB51 101-102)
        |
        +--> Match (0 Diff, Parte A complete) --> [CIERRE RECEP + Bultos + 7-Doc Binder]
        |
        +--> Shortage / Damage (Cols N, O) ----> [PARA N.C RECEP. -> SAP NCR 400000xxxx]
        |                                         (MIRO Credit Memo / 122-161 Return? Validate)
        +--> Partial Container (Parte B/C/D) --> [Pending Balance -> Col X / EKBE Open Qty]
```

#### Flow B: Distribution Center Dispatch & Store Transfer (`C001` -> `C101+`)
```text
[1. Store Distribution Order 4500xxxxxx] (PED-45-A; Receiving Plant C101..C155, EC02)
        |
        v  (Delivery VL10B/VL02N [Candidate] + Warehouse Picking A001/A004)
[2. Goods Issue from C001: Mov 641 (or 642 reversal) / Mov 643 (if cross-company)]
        |   Stock enters Stock in Transit (MB5T / MARC-TRAME)
        v
[3. Store Receipt (C101+) / Direct 501 / Remaining DC Stock Check (MB52)]
        |
        v
[4. CIERRE DESP + STOCK CONTROL DE INVENTARIO + PARA N.C DESP + Hoja1 MV 998]
```

### 4.4 Business Processes & Authorization Boundaries (R1 & R5)

To uphold **R1** (*"do not allow SAP write operations without explicit authorization and a separate safety design"*) and standard segregation of duties (SoD):

| Role Tier | Permitted Activities | Key SAP Authorization Objects `[Generic ECC 6.0 Candidate — Validate in Env]` | Hard Boundary |
|---|---|---|---|
| **Tier 1: Inventory Reconciliation Analyst / ReconcileAI Read-Only User** | Display POs (`ME80FN`, `ME23N`), material lists (`MM60`, `MM03`, `CS03`), movements/stock (`MB51`, `MB52`, `MB5T`, `MMBE`), table display (`SE16N` read-only), and run ReconcileAI offline on exported workbooks. | `S_TCODE` (display t-codes only); `S_TABU_DIS` / `S_TABU_NAM` with `ACTVT = 03` (Display **only**; **prohibit** `&SAP_EDIT` / `ACTVT = 02`); `M_MSEG_WMB` (`ACTVT = 03`); `M_BEST_WRK` (`ACTVT = 03`). | **Zero posting or master-data modification rights.** ReconcileAI must never execute RFCs, BAPIs, or GUI scripts that post `MIGO`, `MIRO`, `VL31N`, or `MR11`. |
| **Tier 2: Warehouse Receiving / Dispatch Clerk** | Post physical Goods Receipts (`MIGO` `101`/`102`), warehouse putaway/transfer orders (`LT03`/`LT12` or EWM), and store GI (`641`). | `M_MSEG_BWA` (restricted by allowed `BWART` e.g. `101`, `102`, `311`, `641` and `WERKS = C001`). | Cannot approve financial Credit Notes (`40...`), override PO prices, or adjust GR/IR (`MR11`). |
| **Tier 3: Finance / AP & Import Nationalization** | Post customs/vendor invoices and Credit Memos (`MIRO` / `40...`), perform GR/IR clearing (`MR11`), and approve payroll deductions (`MV 998`). | `F_BKPF_BUK` (`BUKRS = IM 01`), `M_RECH_WRK`, `F_BKPF_KOA`. | Segregated from physical dock counting to prevent concealing shrink. |

### 4.5 Inventory Movements, Reversal Pairs & Stock Types

| Movement / Reversal Pair | Standard ECC 6.0 Meaning `[Generic ECC 6.0 Candidate]` (or `[Confirmed]`) | Stock / Accounting Impact | Relevance to ReconcileAI |
|---|---|---|---|
| **`101` / `102`** `[Confirmed]` | PO Goods Receipt / Reversal (`MB51`, `MIGO`). | `+` / `-` unrestricted (or QI/blocked) stock; credits/debits `WRX` GR/IR clearing vs `BSX` inventory. | Net ERP received quantity = `SUM(101) - SUM(102)`. Must reconcile against `received_real_units`. |
| **`103` / `105`** `[Candidate]` | GR into Blocked Stock (non-valuated) / Release to valuated stock. | `103` does not post FI inventory value until `105`. | Explains why dock-received goods may not yet appear in `MB52` unrestricted or `MB51 101`. |
| **`122` / `123` & `161` / `162`** `[Generic ECC 6.0 Candidate]` | Return delivery to vendor against GR (`122`, reversed by `123`) vs Returns PO goods issue (`161`, reversed by `162`). | Two distinct return mechanisms: `122` reverses/references the original GR and re-opens `WRX`; `161` posts GI on a Returns PO item expecting a vendor credit memo. | Relevant when damaged garments (`Col O`) are physically returned vs financial-only NCR (`40...`). |
| **`311` / `312`** `[Candidate]` | Plant-internal Storage Location transfer (`A001` <-> `A004`). | No FI posting; updates `MARD-LABST` between storage locations in `C001`. | Explains stock moving from dock/A001 to `A004` resurtido without changing total `C001` stock. |
| **`501` / `502`** `[501 Confirmed; 502 Candidate]` | Receipt without PO / Reversal (`CIERRE RECEP!Q`). | Increases stock without PO history (`EKBE`); offsets `GBB` account. | Must be tracked in `Col Q` (`UNIDADES CON MOV 501 EN TIENDA`) so direct store receipts are not double-counted as missing. |
| **`641` / `642`** `[Confirmed]` | Intra-company 2-step STO Goods Issue / Reversal (`C001` -> `C101+`). | Moves stock from `C001` `LABST` to receiving plant `TRAME` (Stock in Transit, `MB5T`). | Core `CIERRE DESP!AB` & `CIERRE RECEP!U` metric (`Mov 641 - Mov 642`). |
| **`643` / `644`** `[643 Confirmed in Template; 644 Candidate]` | Cross-company 2-step STO Goods Issue / Reversal. | Triggers intercompany billing (`SD`) if plants belong to different company codes. | Requires live SAP validation against local `"proceso de facturar"` label (`CIERRE DESP!AG`). |
| **`701` / `702`** `[Candidate]` | Physical inventory count gain / loss (`MI07`). | Adjusts `LABST` against inventory difference account (`INV`). | Relevant for post-audit `STOCK CONTROL DE INVENTARIO` write-offs (`Hoja1 MV 998` check). |

**Stock Status Types (`MARD` / `MMBE` / `MB52`):**
- **`LABST` (Unrestricted-Use / Libre Utilización):** Available for store distribution (`45...` / `641`).
- **`INSME` (Quality Inspection Stock):** Held pending QC release; **excluded** from unrestricted `MB52` column totals unless selected.
- **`SPEME` (Blocked Stock):** Damaged/quarantined stock; **excluded** from unrestricted `MB52`.
- **`TRAME` (Stock in Transit):** Dispatched via `641` from `C001`, not yet received (`101`) at store `C101+`; monitored via `MB5T`.

### 4.6 Valuation & GR/IR Clearing Impact

1. **Cost Basis Separation (`43` vs `41` vs `MBEW` vs `VKP2`):**
   - **`EKPO-NETPR` (`PO 43`):** Negotiated FOB/supplier unit price (`CIERRE RECEP!AE = IFERROR(AK/F,0)`, populated by `Módulo1!Mueve43` from `ME80FN` `"Valor Neto"`).
   - **`Doc 41` (`CIERRE RECEP!AF = IFERROR(AL/G,0)`, header `'Valor unitario PEDIDO NAC $'`):** `[Unverified / Needs Env Check]` per `§3.4 #3` (`line 151`). In `Módulo1!Mueve41` (`line 153`), `Col AL` is populated directly from `ME80FN` `"Valor Neto"` (`Pedidos_41`), and across inspected workbooks (`IMP-100` 18/18 rows, `IMP-182` 135/136 rows, `IMP-207 NCR` 13/13 populated rows) `AF == AE` with **no** observed customs DUA/freight uplift (`AL=None`/`G=0` yields `AF=0`). Any separate DUA/freight landed-cost component or `MIRO` role for `41...` requires live SAP validation (`§3.6 #2`).
   - **`MBEW-VERPR` / `STPRS` `[Generic ECC 6.0 Candidate — Validate in Env]`:** SAP accounting valuation price (Moving Average `V` or Standard `S`).
   - **`VKP2` (Retail Sales Price):** Store selling price in Colones/local currency (`CONTROL DATA`, e.g. `18,000.00` or `35,674.39`). **Hard Rule:** `VKP2` must **never** be mapped to `unit_cost` or used in NCR valuation (`src/heuristic_parser.py:18` must be corrected).
2. **GR/IR Clearing Account (`WRX`) Impact `[Generic ECC 6.0 Candidate — Validate in Env]`:**
   - In standard ECC 6.0, when `MIRO` (`RBKP`, typically `51...`) posts an invoice before all physical containers arrive (`Parte B/C`) or when shortages occur (`Col N`), `WRX` carries an open debit/credit imbalance (`Invoice Qty > GR 101 Qty`). Whether `41...` triggers `MIRO`/`WRX` postings or is purely a custom purchasing document (`EKKO-BSART`) is `[Unverified / Needs Env Check]` (`§3.4 #3`).
   - **Clearing Mechanism Distinction:** `MR11` performs **internal GR/IR account maintenance** (writing off `WRX` balances internally to inventory/price-difference accounts `BSX`/`PRD` without debiting the vendor), whereas a `MIRO` **vendor credit memo** debits vendor AP (`KBS`, candidate mechanism for `40...` NCRs via `PARA N.C RECEP.`, pending `§3.6 #3`). ReconcileAI must distinguish **unreceived split containers (`PENDIENTE_PARTE_B`)** from **claimable supplier shortages (`FALTANTE`)** before any financial clearing recommendation.

### 4.7 Warehouse Management (`LE-WM` vs `EWM` in `C001`)

- In `C001`, physical operations split stock across `A001` (Central/Picking), `A002`, `A004` (Resurtido high racks), `03`–`99`, and `Especiales` / `ACONDICIONAMIENTO`.
- **Critical Timing Gap:** `MB51` (`Mov 101`) records IM-level receipt at the storage location, whereas WM putaway (`LT03`/`LT12` in `LE-WM`, or Warehouse Tasks in `EWM` via `Check list DESPACHO!A17`) moves quants from the interim dock bin (`902` GR area) to final high-rack bins (`7.UBICACIÓN MONITOR`). Analysts checking bin stock before putaway confirmation see stock stuck in interim zones even though `MB51 101` is complete.

### 4.8 Real-World Troubleshooting Playbook (10 Evidence-Grounded Scenarios)

| # | Scenario | Observable Symptom in Repo / Corpus | SAP Root Cause & Diagnostic Check | Deterministic / Operational Resolution |
|---|---|---|---|---|
| 1 | **`PRK` vs `UN` Bidirectional Entry Error** | `81` PRKs entered instead of `1,728` units (`SOP_RECEPCION_SAP.md:107`), or units entered as PRKs (`rec > inv * 2.5`). | `ME80FN` UMP (`PRK` vs `UN`) mixed with physical piece counts; `MARM` ratio (`UMREZ`/`UMREN`) or `STPO-MENGE` curve (`12`/`24`) misapplied; `validator.py:47` only catches over-entry, missing under-entry. | Check `EINA-LMEIN` and `MAST`/`STPO` (`MENGE`) via `SE16N`; add guarded lower-bound check in `validator.py` (`um == 'PRK' and curve > 2 and rec > 0 and rec <= (inv / curve) * 1.25`—and skip `PENDIENTE_PARTE_*`) so `curve <= 2` or `rec == 0` (`Parte B`) never false-positives. |
| 2 | **18-Char `MATNR` Leading Zeros, Scientific Notation & Alphanumeric SKUs** | Join miss between `CIERRE RECEP`, `CONTROL DATA`, and `MB51`; `"1.00E+06"` in Excel; `PARTE X` skips alphanumeric SKUs. | SAP `CHAR(18)` zero-padding (`CONVERSION_EXIT_MATN1_INPUT`); Excel float coercion; `sanitizer.py:41-57` corrupts `"10E5"` and misses 8–10 char or `00*` codes; `parsers.py:287` requires `.isdigit()`. | Always format `MATNR` as text before pasting from SAP GUI; normalize consistently in `sanitizer.py` preserving alphanumeric SKUs and exact EANs. |
| 3 | **Split Containers (`Parte A/B/C/D`) & `Bulto Cero` Rule** | Apparent massive shortage (`-500` units) in `CIERRE RECEP` even though goods are on another container (`IMP-81`). | `41...` covers total nationalized shipment, while `MIGO 101` and QC count only `Parte A` arrived at dock. | Tag `PENDIENTE_PARTE_X`, enforce `Bulto Cero` (`AS=AT=AW=AX=0.0`), track open balance in `Col X` (`CIERRE RECEP`) / `Col AE` (`CIERRE DESP`) and `EKBE`. |
| 4 | **Blocked (`SPEME`) or QI (`INSME`) Stock Invisible in `MB52`** | `MB51 101` shows receipt posted, but `CIERRE DESP!N` (`MB52`) shows `0` available for store `641` dispatch. | Goods posted to `INSME` (Quality Inspection) or `SPEME` (Blocked) or interim WM bin (`902`), whereas default `MB52` layout highlights `LABST`. | Re-run `MB52` with "Also select blocked/QI stock" or check `MMBE` / `MARD` (`LABST`, `INSME`, `SPEME`) and `LX02`/`LS24` bin location. |
| 5 | **Cross-Period `102` / `642` Reversals** | `MB51` balance for `IMP-XXX` mismatches dock count depending on date filter. | Original `101` posted in Month M and `102` reversal (or `642`) posted in Month M+1 (`MKPF-BUDAT`). | Always query `MB51` by Material + Plant `C001` with wide `BUDAT` window (or by PO `EBELN` in `EKBE`) and compute net `SUM(101) - SUM(102)`. |
| 6 | **`Mov 101` Column Conflict (`Col I` vs `Col BF`)** | ReconcileAI reads `0` or wrong `Mov 101` when macro output differs from expectation. | `Mueve51` writes `MB51 101-102` to `Col BF` (base/PRK), `Col I` computes `=BF*BB` (units), while `parsers.py:409` scans `c < 20` (`Col I`). | Document both columns explicitly; verify cached formula values (`data_only=True`) exist in `Col I` and reconcile `sap_mov_101` against `received_real_units` in `engine.py`. |
| 7 | **Multi-PO & Hyphenated PO Strings** | Material appears across multiple `43...` POs (`4300006709-7847`) causing false unmatched rows. | `Mueve43` concatenates POs with hyphens/spaces; `parse_multi_po` (`sanitizer.py:136`) fails to split on `-` and `validator.py:86` flags duplicate identical POs. | Aggregate by `MATNR` across POs using `EKBE` / `ME80FN` totals and fix `parse_multi_po` delimiter regex in Track A/Roadmap #1. |
| 8 | **Zero Unit Cost / `ncr_amount = 0` on Discrepancy** | Row status is `FALTANTE`, yet `PARA N.C RECEP.` / `ncr_amount` is `$0.00`. | `parsers.py:379` hardcodes `unit_cost` at `Col AF (32)` without header scan, and `parsers.py:603` computes `(falt + dan) * unit_cost` ignoring `diff`. | Cross-check `Col AE`/`AF` against `ME80FN` (`EKPO-NETPR`) for `43...` / `41...` (noting `AF` defaults to `0` when `AL`/`G` is unpopulated); surface `PRECIO_CERO_EN_RECLAMO` in UI (`validator.py:57`). |
| 9 | **Template Drift (`Tarima` Shift & `-A` Split Sheets)** | QC `CONTROL DATA` columns shift by +1 (`IMP-100` vs `IMP-182`); `IMP-200.xlsm` uses `CIERRE DESP-A` & `PED-45-A`. | Operators insert `Tarima` or split dispatches across multiple containers without updating hardcoded column indices. | Rely on header-based detection (not hardcoded indices like `Col 32` or row 6 only) and maintain template variant profiles in KB. |
| 10 | **`VKP2` Retail Price Mistaken for Import Cost** | Discrepancy claim value exceeds PO value by 2x–4x. | `VKP2` price (Col `U` in `BASE DE DATOS` / `Precios VKP2` in `CONTROL DATA`) confused with import cost; `heuristic_parser.py:18` lists `"vkp2"` under `"unit_cost"`. | Ban `VKP2` from cost/NCR mapping; restrict `VKP2` strictly to retail price verification (`ESQUEMA A001` / `CONTROL DATA`). |

---

## 5. Strategy for a Reliable, Source-Backed, Version-Aware Knowledge Base

### 5.1 Proposed Repository Layout under `docs/sap-knowledge/`

When Milestone #4 of the Roadmap (**R7**: *SAP knowledge expansion*) is executed, store curated SAP knowledge as atomic, versioned Markdown files with strict YAML frontmatter plus a machine-readable index under `docs/sap-knowledge/`:

```text
docs/sap-knowledge/
├── KNOWLEDGE_EXPANSION_PLAN.md      # This strategic plan (Milestone M2)
├── SCHEMA.md                        # Formal YAML frontmatter & validation rules
├── index.json                       # Generated catalog for validator.py & Track E RAG
├── generic-ecc6/                    # Broadly applicable SAP ERP 6.0 (ECC_6.0 / 6.00.29 baseline) MM/LE/WM/FI-MM entries
│   ├── tcodes/                      # e.g., me80fn.md, mb51.md, mb52.md, migo.md, miro.md
│   ├── tables/                      # e.g., ekko-ekpo.md, eina.md, mast-stpo.md, mkpf-mseg.md, mbew.md
│   ├── movements/                   # e.g., mov-101-102.md, mov-641-642.md, stock-types.md
│   └── processes/                   # e.g., gr-ir-clearing.md, authorizations-readonly.md
├── company-ekono/                   # Company-specific Ekono / Almacema (IM 01) workflows
│   ├── doc-types-43-41-45-40.md
│   ├── org-c001-stores-c101.md
│   ├── z-objects-and-macros.md
│   └── workbook-contracts.md        # CIERRE RECEP, CIERRE DESP, Bultos, CONTROL DATA
├── troubleshooting/                 # Playbook entries linked to validator.py issue_type IDs
│   ├── prk-vs-un-confusion.md
│   ├── matnr-18char-padding.md
│   └── split-containers-parte-b.md
└── validation-log/                  # Signed evidence from live SAP ECC 6.0 checks (§3.6)
```

### 5.2 Mandatory Metadata Schema

Every KB entry file MUST begin with validated YAML frontmatter containing these exact fields:

| Field | Allowed Values / Format | Purpose |
|---|---|---|
| `entry_id` | `sap.<domain>.<slug>` (e.g. `sap.tcode.me80fn`) | Stable identifier referenced by `validator.py` and RAG citations. |
| `title` | Concise descriptive string | Human-readable title (no emojis). |
| `object_kind` | `tcode` \| `table` \| `movement_type` \| `doc_prefix` \| `z_program` \| `process` \| `troubleshooting` \| `workbook_contract` | Enables exact typed lookup. |
| `source_type` | `company_specific` \| `generic_sap` | Enforces R1/R5 separation between Ekono workflows and standard SAP. |
| `sap_release` | `ECC_6.0` (`6.00.29` baseline) \| `S4HANA` \| `IS_RETAIL` \| `release_agnostic` (array supported) | Prevents mixing S/4HANA/EWM instructions into ECC 6.0 workflows. |
| `sap_doc_version` | Optional string (e.g. `"6.00.29"`, or omit / `null`) | Records the exact SAP Help Portal version anchor when applicable (`"6.00.29"` for ECC 6.0). |
| `verification_status` | `confirmed` \| `unverified` \| `needs_env_check` \| `conflicting` | Gates whether `validator.py` or Track E AI may treat the entry as authoritative. |
| `source_citation` | Array of `{type, ref[, note]}` where `type` is `repo_file` \| `vba_macro` \| `corpus_file` \| `corpus_audit` \| `sap_help_url` \| `sap_note` \| `live_env_check`, and `note` is an optional string | Enforces R8 traceability and forbids unsourced claims (`repo_file` uses `path:line` in `ref`; `corpus_audit` cites aggregate scan results). For ECC 6.0 `sap_help_url` entries, anchor to `https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29` and record the exact retrieved sub-page/topic path once fetched (never fabricated). |
| `owner_role` | `inventory_analyst` \| `sap_mm_functional` \| `sap_basis_security` \| `engineering` | Accountability for review and updates. |
| `effective_date` | `YYYY-MM-DD` | Date the rule/procedure became operationally effective. |
| `last_verified_date` | `YYYY-MM-DD` (or `null` if unverified) | Tracks staleness and triggers periodic re-verification. |
| `ai_usage_policy` | `informational_only` \| `rag_with_status_banner` \| `prohibited_for_math` \| `read_only_guidance` (single enum or array supported) | Machine-enforced guardrail for Track E RAG/MCP/LLM prompts (array allows combining status banners with deterministic-math and read-only enforcement). |

**Example Encodings (Demonstrating `ECC_6.0` `6.00.29` Anchor vs. `conflicting` Handling):**
```yaml
---
entry_id: sap.tcode.me80fn
title: "General Evaluation of Purchasing Documents (ME80FN)"
object_kind: tcode
source_type: generic_sap
sap_release: "ECC_6.0"
sap_doc_version: "6.00.29"
verification_status: needs_env_check
source_citation:
  - { type: "repo_file", ref: "docs/SOP_RECEPCION_SAP.md:60-66" }
  - { type: "vba_macro", ref: "xl/vbaProject.bin!Módulo1 (Format43 / Mueve43)" }
  - { type: "sap_help_url", ref: "https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29", note: "Official ECC 6.0 Help Portal anchor; append exact retrieved sub-page/topic URL during Track D execution" }
owner_role: sap_mm_functional
effective_date: 2026-11-17
last_verified_date: null
ai_usage_policy: rag_with_status_banner
---
```
```yaml
---
entry_id: sap.tcode.vl31n_rfui
title: "Inbound Delivery & Warehouse RF (VL31N / RFUI)"
object_kind: tcode
source_type: generic_sap
sap_release: ["ECC_6.0", "S4HANA"]
verification_status: conflicting
source_citation:
  - { type: "repo_file", ref: "docs/SOP_RECEPCION_SAP.md:30-31" }
  - { type: "corpus_audit", ref: "0 occurrences across 15 sandbox_benchmark cases and 593 data/ files" }
owner_role: sap_mm_functional
effective_date: 2026-11-17
last_verified_date: null
ai_usage_policy: ["rag_with_status_banner", "prohibited_for_math", "read_only_guidance"]
---
```

### 5.3 Governance Rules & ReconcileAI Development Rules Enforcement (R1, R5, R6)

1. **Zero Fabrication Rule:** Contributors and AI agents MUST NOT invent SAP transaction codes, table/field names, IMG paths, SAP Note numbers, or **SAP Help Portal sub-URLs, page titles, or unverified page contents**. Until an individual topic page under `https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29` is actually retrieved and verified, entries may cite only that root `version=6.00.29` portal anchor and any generic ECC 6.0 candidate added before live verification MUST carry `verification_status: needs_env_check` and an explicit banner.
2. **Epistemic Gating in `validator.py` and Track E RAG (Qdrant):**
   - `index.json` and Track E vector ingestion MUST store `verification_status`, `source_type`, and `sap_release` as payload metadata filters in Qdrant.
   - Retrieval pipelines MUST filter or visibly annotate results: entries marked `unverified`, `needs_env_check`, or `conflicting` MUST **never** be presented as confirmed Ekono SOP truth; any AI response citing them MUST surface the exact verification caveat and required live SAP check (§3.6).
3. **Deterministic Math Sovereignty (R1 & R6):**
   - Knowledge base entries may explain *why* a discrepancy occurs (`PRK` vs `UN`, `Mov 101 - 102`, `Bulto Cero`, `Parte B`), **never** compute or override authoritative quantities, carton breakdowns, unit costs, or NCR dollar totals. All calculations remain strictly deterministic in Python (`src/engine.py`, `src/parsers.py`).
4. **Strict Read-Only SAP Authorization Boundary (R1):**
   - No KB procedure, CLI command, MCP tool, or AI prompt may instruct or automate SAP write operations (`MIGO`, `MIRO`, `VL31N`, `MR11`, `SE16N &SAP_EDIT`, write RFC/BAPI calls) without explicit user authorization and a separate, reviewed safety design.
5. **Curated Ingestion & Change Control:**
   - Every KB addition or status promotion (`needs_env_check` -> `confirmed`) requires: (a) for generic ECC 6.0 (`6.00.29`) semantics, an exact retrieved and verified sub-page/topic citation under `https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29` (never fabricated from memory) plus a cited live-environment evidence artifact in `validation-log/` (e.g., anonymized `T161`/`T156`/`T320` screenshot or `System -> Status` export), (b) schema validation of YAML frontmatter in CI, and (c) re-indexing of `docs/sap-knowledge/index.json` / Qdrant chunks per atomic `entry_id` (no arbitrary fixed-token splitting across entries).

### 5.4 Open Questions for the User & Sequencing

To unblock promotion of `needs_env_check` and `conflicting` items prior to Roadmap Step 4 (SAP Knowledge Expansion):
1. What exact SAP release, Enhancement Pack, and IS-Retail switch status appear under `System -> Status` in the organization's SAP GUI (confirming alignment with the SAP ERP 6.0 / `version=6.00.29` Help Portal baseline vs. S/4HANA claims)?
2. Are `VL31N`, `/SCWM/RFUI` (decentralized EWM), `18...` inbound deliveries, and `80...` handling units actually used in warehouse operations, or is `"Respaldo SAP EWM"` in `Check list DESPACHO!A17` legacy/unused or tied to a separate EWM instance?
3. In `T161`/`EKKO`, what SAP document type (`BSART`) is `41...` (`Pedido de Nacionalización`), and why is it queried via `ME80FN`?
4. Do retail stores (`C101+`) belong to the same company code `IM 01` as `C001` or separate company codes (clarifying local `Mov 643` usage)?
5. What are the exact ABAP/transaction definitions behind `Z-Carga`, `ZDESCARGA`, `Z51 -Z52`, `"7.UBICACIÓN MONITOR"`, and `MV 998`?

**Recommended Next Execution Sequence (Aligned with R7):** Resolve code correctness and regression tests first (Roadmap Priorities #1–#2, including surfacing `validator.py`, fixing `clean_material_code`/`parse_multi_po`, removing the `VKP2` unit-cost synonym, and reconciling `sap_mov_101`), execute the read-only SAP validation checklist (§3.6) with the organization's functional/Basis team while retrieving exact topic citations from `https://help.sap.com/docs/SAP_ERP_Additional_Documentation?version=6.00.29`, and then populate the versioned `docs/sap-knowledge/` registry (Roadmap Priority #4) ahead of Track E RAG indexing (Priority #6).
