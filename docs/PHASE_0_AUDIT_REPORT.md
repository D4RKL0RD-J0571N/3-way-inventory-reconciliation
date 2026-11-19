> **Estado: HISTÓRICO / Historical snapshot.** Describe el sistema Python-only (revisión 589f6aa) o un plan previo. Para el estado actual vea [docs/README.md](README.md) y [CURRENT_ARCHITECTURE](architecture/CURRENT_ARCHITECTURE.md).

# Phase 0 Audit & Baselines Report — Asistente de Cierre de Recepción (ACR) / Receiving Closure Assistant (RCA)

**Document ID:** `PHASE-0-AUDIT-ACR-2026-11-18`  
**Audit Date (UTC):** `2026-11-18T05:44:00Z`  
**Target System:** Asistente de Cierre de Recepción (ACR) / Receiving Closure Assistant (RCA) — C# / WPF on `.NET 10 LTS`  
**Operating Organization:** `IMPORTADORA ALMACEMA, S.A.` / Distribution Center `C001`  
**Repository & Branch:** `3-way-inventory-reconciliation` (`repository_audit_planning`, HEAD `14240a81edfc0298a65b3c644053e13321e13b8d`)  
**Audit Mode:** Strictly non-destructive, read-only against production data (`data/` and `Control_Diario_Operaciones_Jostin.xlsx`)

---

## Executive Summary & Phase 0 Gate Readiness

Phase 0 establishes the verified technical, operational, and dataset baselines required before writing any C# code for the `.NET 10 LTS` ACR desktop workstation (`Acr.Core`, `Acr.Infrastructure`, `Acr.Desktop`, `Acr.Tests`).

### Verified Baselines Summary

| Audit Domain | Verified Finding | Phase 0 Status |
|---|---|---|
| **Repository & Git Baseline** | Branch `repository_audit_planning` at commit `14240a8` contains 4 authoritative Python services (`quantity_engine.py`, `qc_reader.py`, `discrepancy_classifier.py`, `closure_exporter.py`), 9 legacy/UI/support modules, and **0** `.cs`/`.csproj`/`.sln` files. | Verified |
| **Python Regression Baseline** | `python -m pytest -v --tb=short` passes **58/58 tests** (`0` failures, `0` skipped) in read-only mode across 51 Python-matched QC files and 77 master closure workbooks. Bare `pytest` fails collection (`ModuleNotFoundError: No module named 'src'`). | Verified |
| **Production Dataset Corpus** | `data/` holds **830** entries (including **147** readable `.xlsm` workbooks plus **1** `~$` `.xlsm` lockfile = **148** `.xlsm`-extension entries total, **97** `.xlsx`, **62** BIFF `.xls`, **160** `.pdf`, **71** `.msg`, **35** `.eml`). Sheet-signature inspection across the **147** readable `.xlsm` files reconciles **63** standalone QC (`CONTROL DATA`) + **1** hybrid (`IMP-218`) = **64** QC workbooks, and **77** non-variant + **4** `copia`/`sin control` = **81** `CIERRE RECEP` workbooks with **80** unique `vbaProject.bin` SHA-256 hashes. | Verified |
| **`CIERRE RECEP` Column Schema** | **81/81** closures share an identical **58-column (`A`–`BF`)** Row-12 schema (53 labeled headers, 2 formula-echo helper columns `AQ`/`AR`, 3 spacer columns `AP`/`AV`/`BA`), reconciling the `"45 columns"` reference (`AS` = Col 45; `58 - 13` optional/spacer/notes cols = `45`). Exactly **5/81** workbooks carry pre-existing formula/literal errors (`IMP-241`, `IMP-244`, `IMP-256`, `IMP-360`, `IMP-90`). | Verified |
| **Daily Log (`Control_Diario_Operaciones_Jostin.xlsx`)** | **5** sheets, **3** Excel tables (`tblMaestroContenedores` `A5:R27`, `tblParteB`, `tblNotificaciones`), **22** active container rows, exact **6** lifecycle stage names and `Catalogos!A1:F7` dropdown values, 7-state `Col O` auto-rollup formula, SAP intake notification template (`Col R`), and **5** live data integrity exceptions (`IMP-202` duplicate, `IMP-320`/`IMP-213` shared PO/Pedido, stale `J8`, stray `B45`, table/validation row range mismatch). | Verified |
| **Historical Sync Baseline** | Author attribution lives in `'DATOS GENERALES'!B33` (not `docProps/core.xml`, which carries 2016 template author `Hilario Lopez Rivera` in 19/20 files). **20** files reference Jostin (`IMP-326` spells `"Jostin Lopez Sobalbaro"`), resolving to **18** unique containers + **2** `IMP-182` drafts and **3** containers (`IMP-224/225/226`) sharing PO `4300007562`. | Verified |
| **SAP Docs & Processes** | Main workspace contains 4 authoritative operational artifacts absent from the worktree (`SOP-LOG-C001-01.docx`, `drafts/Borrador de procedimientos operativos.txt`, `drafts/BorradorRecepcion.txt`, `Control_Diario_Operaciones_Jostin.xlsx`), reconciling `VL31N`, `RFUI`, and `/SCWM/MON` as active Stage 2 intake transactions. Live SAP GUI 7.00 backend release (`ECC 6.0` vs `S/4HANA`) and Z-object definitions remain pending live GUI verification. | Verified (live GUI items flagged) |
| **Host `.NET 10 SDK` Setup** | Machine-wide `winget install Microsoft.DotNet.SDK.10` (`10.0.404`) failed at UAC elevation (`0x80070032`) on `C:\Program Files\dotnet` (which retains `.NET 8.0.428` and `9.0.312`). With explicit user authorization (`"ADELANTE"`), `.NET 10 SDK` (`10.0.404`), `Microsoft.NETCore.App 10.0.15`, `Microsoft.AspNetCore.App 10.0.15`, and `Microsoft.WindowsDesktop.App 10.0.15` were installed user-locally via Microsoft's `dotnet-install.ps1` into `%LocalAppData%\Microsoft\dotnet`, with User `DOTNET_ROOT` and User `PATH` configured. | **Verified (User-Local)** |

### Phase 0 Gate Readiness Assessment
- **Audit & Baseline Documentation Gate:** Ready. All code, dataset, schema, lifecycle, historical sync, exception, and SAP documentation requirements are audited and documented herein before any C# code exists.
- **Host SDK Gate:** **PASS.** `.NET 10 SDK` (`10.0.404`) and `Microsoft.WindowsDesktop.App 10.0.15` are installed and verified at `C:\Users\Jostin\AppData\Local\Microsoft\dotnet`. Because Windows System `PATH` (`C:\Program Files\dotnet\`) precedes User `PATH`, build/test commands in non-elevated shells must prepend `$env:LOCALAPPDATA\Microsoft\dotnet` to `PATH` (or invoke `& "$env:LOCALAPPDATA\Microsoft\dotnet\dotnet.exe"`). Phase 1 (`Acr.Core`) may begin immediately.

---

## 1. Host & Tooling Environment

### 1.1 Operating System, Tooling & Python Environment
- **OS:** Windows (`10.0.26200`, `win-x64`)
- **Visual Studio / Desktop Tooling (via `vswhere.exe`):**
  - Visual Studio Community 2022 (`17.14.37111.16` at `C:\Program Files\Microsoft Visual Studio\2022\Community`) with `Microsoft.VisualStudio.Workload.ManagedDesktop` (WPF / .NET Desktop workload installed).
  - Visual Studio Build Tools 2022 (`17.14.36127.28`) and Visual Studio Build Tools 2019 (`16.11.36227.6`).
- **Python Environment:** Python `3.13.5` (`C:\Python313\python.exe`), `pytest-9.0.2`, `pluggy-1.6.0`. No `.venv` directory exists in the worktree (`Test-Path ".venv"` = `False`).
- **Package Managers:** `winget.exe` (`C:\Users\Jostin\AppData\Local\Microsoft\WindowsApps\winget.exe`), `choco.exe` (`0.12.1.0`), `pwsh.exe` (`7.6.10.0`).

### 1.2 Machine-Wide Installed .NET SDKs & Runtimes (`C:\Program Files\dotnet`)
Commands executed:
```powershell
dotnet --list-sdks
dotnet --list-runtimes
```
Observed outputs:
```text
8.0.428 [C:\Program Files\dotnet\sdk]
9.0.312 [C:\Program Files\dotnet\sdk]
---
Microsoft.AspNetCore.App 6.0.36 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 8.0.25 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 8.0.34 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.AspNetCore.App 9.0.14 [C:\Program Files\dotnet\shared\Microsoft.AspNetCore.App]
Microsoft.NETCore.App 6.0.36 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 7.0.20 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 8.0.25 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 8.0.34 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.NETCore.App 9.0.14 [C:\Program Files\dotnet\shared\Microsoft.NETCore.App]
Microsoft.WindowsDesktop.App 6.0.36 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 7.0.20 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 8.0.25 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 8.0.34 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
Microsoft.WindowsDesktop.App 9.0.14 [C:\Program Files\dotnet\shared\Microsoft.WindowsDesktop.App]
```
- `global.json`: Not present in repository root.

### 1.3 `.NET 10 SDK` Elevation Blocker & Authorized User-Local Installation
1. **Machine-Wide Install Attempt (`0x80070032` Blocker):**
   Command executed:
   ```powershell
   winget install --id Microsoft.DotNet.SDK.10 --exact --accept-package-agreements --accept-source-agreements
   ```
   Observed result:
   - Package resolved: `Microsoft.DotNet.SDK.10` version `10.0.404` (`dotnet-sdk-10.0.404-win-x64.exe`, SHA-256 `f7fdaf56ff976c3450e6e984ddea3f01790b45ecec848bcfa5dc32f62b6aa93a`), cached under `%LocalAppData%\Temp\WinGet\Microsoft.DotNet.SDK.10.10.0.404\`.
   - Installation **failed with exit code 50** (`0x80070032: Failed to launch elevated child process`) because the agent shell token has `BUILTIN\Administrators` set to `Group used for deny only` and UAC elevation is blocked in non-interactive/restricted tool execution (`Start-Process -Verb RunAs` returned `ERROR_NOT_SUPPORTED`; `C:\Program Files\dotnet` ACL grants write only to `TrustedInstaller`, `SYSTEM`, and `BUILTIN\Administrators`).

2. **Authorized User-Local Installation (`"ADELANTE"` — Completed & Verified):**
   With explicit user authorization (`"ADELANTE"`), `.NET 10 SDK` (channel `10.0` LTS) was installed user-locally via Microsoft's official `dotnet-install.ps1` into `C:\Users\Jostin\AppData\Local\Microsoft\dotnet` (`%LocalAppData%\Microsoft\dotnet`), and the temporary installer script was deleted afterward:
   ```powershell
   Invoke-WebRequest -Uri "https://dot.net/v1/dotnet-install.ps1" -OutFile "$env:TEMP\dotnet-install.ps1" -UseBasicParsing
   & "$env:TEMP\dotnet-install.ps1" -Channel 10.0 -InstallDir "$env:LOCALAPPDATA\Microsoft\dotnet"
   [Environment]::SetEnvironmentVariable("DOTNET_ROOT", "$env:LOCALAPPDATA\Microsoft\dotnet", "User")
   # Prepended C:\Users\Jostin\AppData\Local\Microsoft\dotnet to User PATH (without duplicates)
   Remove-Item "$env:TEMP\dotnet-install.ps1" -Force
   ```
   Verified outputs (`& "$env:LOCALAPPDATA\Microsoft\dotnet\dotnet.exe" --list-sdks` / `--list-runtimes`, and `dotnet --info` with `$env:DOTNET_ROOT` and prepended `$env:PATH`):
   ```text
   10.0.404 [C:\Users\Jostin\AppData\Local\Microsoft\dotnet\sdk]
   ---
   Microsoft.AspNetCore.App 10.0.15 [C:\Users\Jostin\AppData\Local\Microsoft\dotnet\shared\Microsoft.AspNetCore.App]
   Microsoft.NETCore.App 10.0.15 [C:\Users\Jostin\AppData\Local\Microsoft\dotnet\shared\Microsoft.NETCore.App]
   Microsoft.WindowsDesktop.App 10.0.15 [C:\Users\Jostin\AppData\Local\Microsoft\dotnet\shared\Microsoft.WindowsDesktop.App]
   ```
   - `.NET SDK`: `10.0.404` (`Commit: 2ab2854867`, `MSBuild version: 18.9.15+2ab285486`, Host `10.0.15`, RID `win-x64`).
   - **Environment Variables & Windows `PATH` Precedence Note:** User-level `DOTNET_ROOT` is set to `C:\Users\Jostin\AppData\Local\Microsoft\dotnet` and `C:\Users\Jostin\AppData\Local\Microsoft\dotnet` is prepended to User `PATH`. However, on Windows, **Machine/System `PATH` (`C:\Program Files\dotnet\`) precedes User `PATH`**, so a bare `dotnet` invocation in a new non-elevated shell resolves to `C:\Program Files\dotnet\dotnet.exe` (`.NET 8.0.428` / `9.0.312`). All Phase 1–6 build, test, and publish commands in non-elevated shells must either prepend `$env:LOCALAPPDATA\Microsoft\dotnet` to `$env:PATH` (with `$env:DOTNET_ROOT = "$env:LOCALAPPDATA\Microsoft\dotnet"`) or invoke `& "$env:LOCALAPPDATA\Microsoft\dotnet\dotnet.exe"` explicitly. *(Optional future convenience: if the host administrator later installs `Microsoft.DotNet.SDK.10` system-wide into `C:\Program Files\dotnet`, bare `dotnet` will also resolve `10.0.404` without prepending `PATH`.)*

---

## 2. Repository & Python Reference Baseline

### 2.1 Git State & File Inventory
- **Active branch:** `repository_audit_planning`
- **HEAD commit:** `14240a81edfc0298a65b3c644053e13321e13b8d` (`fix(qc_reader): prevent comp_cartons from matching units_comp column headers`)
- **Other branches:** `main`, `operational-benchmark`
- **Recent authoritative service commits:**
  - `a8dd035`: `src/quantity_engine.py` + `tests/test_quantity_engine.py`
  - `823515e`: `src/qc_reader.py` + `src/discrepancy_classifier.py` + unit tests
  - `e620903`: `src/closure_exporter.py` + `tests/test_closure_exporter.py`
  - `b64e31c`: `tests/test_real_dataset_audit.py`
  - `14240a8`: `src/qc_reader.py` exclusion guard (`if field_name == "comp_cartons" and "UNIDADES" in c_clean: continue`)
- **Existing C#/.NET code:** `0` files (`Test-Path "src/dotnet"` = `False`; zero `*.cs`, `*.csproj`, `*.sln`, `*.xaml` files).

**Module Inventory:**

| Category | Files | Purpose & Migration Role |
|---|---|---|
| **Authoritative Reference Services (`src/`)** | `quantity_engine.py` (334 lines), `qc_reader.py` (308 lines), `discrepancy_classifier.py` (260 lines), `closure_exporter.py` (217 lines) | Authoritative behavioral reference for `Acr.Core` and `Acr.Infrastructure` until Phase 3 parity sign-off. |
| **Legacy / Streamlit / Support (`src/`)** | `app.py` (1,773 lines), `cli.py` (101), `generar_cierre_oficial.py` (130), `run_standalone.py` (32), `config.py` (67), `models.py` (94), `sanitizer.py` (138), `parsers.py` (679), `engine.py` (138), `validator.py` (108), `heuristic_parser.py` (72, dead), `i18n.py` (620), `ui_helpers.py` (283) | Legacy Streamlit app and parsers. Superseded by the 4 services above, while supplying domain context for `'Bultos'` sheet sync, `Bulto Cero`, `validate_all_items`, and known defect regressions (`CODE_AUDIT.md`). |
| **Test Suite (`tests/`)** | Service tests (31): `test_quantity_engine.py` (17), `test_qc_reader.py` (3), `test_discrepancy_classifier.py` (6), `test_closure_exporter.py` (3), `test_real_dataset_audit.py` (2). Legacy/UI tests (27): `test_bultos_empaque.py` (8), `test_i18n.py` (5), `test_parsers.py` (7), `test_sanitizer.py` (7). | Total **58** test functions across **9** files. |
| **Tools (`tools/`)** | `build_exe.py` (130), `operational_stress_test.py` (363) | PyInstaller bundler and historical benchmark harness. |

### 2.2 Behavioral Contracts & Edge Cases of the 4 Authoritative Services

#### A. `src/quantity_engine.py` (Service 1 — 0 External Dependencies)
- **Models (`lines 16–61`):** Frozen `ParsedPackaging`, `CartonCalculationInput`, `CartonCalculationResult` using integer/optional-integer quantities exclusively.
- **Flag Constants (`lines 64–72`):** `QC_COMPLEMENT_WITHOUT_CARTON_COUNT`, `QC_FORMULA_MISMATCH`, `SHARED_CARTON_PRIMARY`, `SHARED_CARTON_SECONDARY`, `VARIANTE_SUELTA_ROW`, `UNEVEN_PREPACK_CURVE`, `NEGATIVE_REPORTED_DIFF`, `QC_OBSERVATION_ADJUSTMENT`, `EMPTY_COMPLEMENT_CARTONS`.
- **Packaging Expression Grammar (`_clean_str`, `_parse_compound_terms`, `_parse_loose_addition`, `parse_packaging_expression`, `lines 75–148`):**
  1. `_clean_str` (`75–80`): Converts `None` and case-insensitive `"none"`, `"nan"`, `"null"` to `""`; trims whitespace.
  2. `_parse_compound_terms` (`83–90`): Regex `r'(\d+)\s*[*xX]\s*(\d+)'` extracts `(cartons, units_per_carton)` integer pairs.
  3. `_parse_loose_addition` (`93–100`): Removes all `N*M`/`NxM` terms via `re.sub(r'\d+\s*[*xX]\s*\d+', '', expr)`, then extracts the first `r'\+\s*(\d+)'` integer (`0` if absent). Thus `'40*60 + 1*72'` returns `loose=0` (two terms, no loose remainder), whereas `'6*192 +48'` returns `loose=48`.
  4. `parse_packaging_expression` (`103–148`):
     - When `terms` is non-empty: `total_cartons = sum(t[0])`, `total_units = sum(t[0] * t[1]) + loose`. `full_cartons, units_per_full = terms[0]`; `complement_cartons, units_per_complement = terms[1]` if `len(terms) > 1` else `(0, 0)`. `is_compound = len(terms) > 1 or loose > 0`; `has_loose = loose > 0`.
     - When `terms` is empty and `raw.isdigit()` (`137–146`): `val = int(raw)` -> `total_cartons = 1`, `total_units = val`, `full_cartons = 1`, `units_per_full = val`, `is_compound = False`. *(Note contrast with legacy `parsers.parse_packaging_string`, where scalar `'42'` meant `total_boxes=42, total_units=0`.)*
     - Otherwise (`148`): returns zeroed `ParsedPackaging(raw_string=raw)`.
- **Complement Rescue (`_calculate_complements`, `lines 151–180`):**
  - `comp_c = c or 0`, `comp_d = d or 0`.
  - `comp_c > 0 and comp_d > 0`: `(comp_c, comp_c * comp_d, [])`.
  - `(c is None or comp_c == 0) and comp_d > 0`: emits `QC_COMPLEMENT_WITHOUT_CARTON_COUNT` and rescues `comp_d` units:
    - `reported_total_cartons == a` -> `(0, comp_d, flags)` (loose units packed inside existing full cartons).
    - `reported_total_cartons == a + 1` -> `(1, comp_d, flags)` (1 uncounted complement carton).
    - Otherwise -> `(0 if a > 0 else 1, comp_d, flags)`.
  - `comp_c > 0 and comp_d == 0`: emits `EMPTY_COMPLEMENT_CARTONS`, returns `(comp_c, 0, flags)`.
  - Otherwise: `(0, 0, [])`.
- **Curve Consistency (`_check_curve_consistency`, `lines 183–202`):**
  - `effective_curve = curve or (1 if uom and uom.upper() == "UN" else None)`.
  - If `effective_curve > 1`: `packs = units // effective_curve`, `loose = units % effective_curve`; non-zero `loose` emits `UNEVEN_PREPACK_CURVE`.
  - Else: `(units if effective_curve == 1 else None, 0, flags)`.
- **Observation Adjustments (`_check_observation_flags`, `lines 205–213`):**
  - Case-insensitive substring match against `("MAS 1PK", "MAS 1 PK", "+1PK", "MENOS 1PK", "MENOS 1 PK", "-1PK")` emits `QC_OBSERVATION_ADJUSTMENT`.
- **Precedence & Entry (`_handle_variante_suelta`, `_handle_shared_carton`, `calculate_carton_quantities`, `lines 216–334`):**
  1. `VARIANTE SUELTA` (`is_variante_suelta` or `"VARIANTE SUELTA"` in `obs.upper()`): `units = reported_total_units or units_in_complement or 0`, `calculated_cartons = 0`, `calculated_packs = units`, `loose_units = 0`, `formula_discrepancy = 0`, `flags = [VARIANTE_SUELTA_ROW]`.
  2. `DOS MATERIALES` (`"DOS MATERIALES"` in `obs.upper()`): secondary when `full_cartons is None or full_cartons == 0` (`calculated_cartons = 0`, `SHARED_CARTON_SECONDARY`); primary uses `full_cartons * units_per_full` (`SHARED_CARTON_PRIMARY`).
  3. Standard path: `full_units = a * b`, rescues complements, `tot_units = full_units + comp_u`, `tot_cartons = a + comp_c`. If `tot_units == 0` and `packaging_string` is set, falls back to `parse_packaging_expression(packaging_string)`. Computes `formula_discrepancy = (tot_units - rep_u)` when `rep_u is not None`, emitting `QC_FORMULA_MISMATCH` if non-zero and `NEGATIVE_REPORTED_DIFF` if `reported_diff < 0`.
  - **C# Migration Notes:** Replace Python arbitrary-precision `int` with `checked` 32/64-bit integer arithmetic (`OverflowException` / explicit validation). Note `reported_formula_units` (`line 40`) is defined on `CartonCalculationInput` but currently unused.

#### B. `src/qc_reader.py` (Service 2)
- **Sheet Resolution (`_find_qc_sheet_name`, `lines 221–229`):** Prefers `"CONTROL DATA"` substring, then `"CONTROL"` or `"CALIDAD"`, else `sheetnames[0]`. *(Survey 2 showed this fallback returns `0` items on `SIN CONTROL CALIDAD` closure workbooks when filename filters misclassify them.)*
- **Converters (`_to_int`, `_to_clean_str`, `lines 66–86`):** `_to_int` parses floats/ints via `int(round(val))` (round-half-to-even), treats `("", "none", "nan", "null", "-", "c", "x")` as `None`, and returns `None` on non-numeric strings.
- **Metadata & Header Resolution (`_extract_row_metadata`, `_build_column_mapping`, `lines 89–144`):**
  - Metadata scanned in rows `1..5` for `"CONTENEDOR"`, `"RESPONSABLE"`, `"NACIONALIZACION"`/`"NACIONALIZACIÓN"`, `"FECHA"` taking `idx + 1`.
  - Header row is the first row with `r_idx > 5` containing `"MATERIAL"` and (`"BULTOS"` or `"UNIDADES"`). Maps 18 fields by first-match substring search on newline-normalized uppercase headers, enforcing two critical exclusions: `field_name == "material" and "PROVEEDOR" in c_clean` (`138–139`) and `field_name == "comp_cartons" and "UNIDADES" in c_clean` (`140–141`, commit `14240a8`).
- **Child SKUs, Footer Filtering & Streaming (`lines 147–308`):**
  - Footer rows matched by `"TOTAL"`, `"SUMA"`, `"RESUMEN"`, `"FIRMA"`, or `(no digits in mat_str and not desc_str)`.
  - `is_variante_suelta = "VARIANTE SUELTA" in tarima.upper()`; `is_child_sku = (len(mat_raw) >= 10 and not tarima and not is_variante_suelta)`.
  - `uom = _to_clean_str(get_val("uom")) or ("PRK" if get_val("total_prk") else "UN")`.
  - Reads with `openpyxl.load_workbook(..., read_only=True, data_only=True)` in a single forward pass, stopping after `15` consecutive empty rows post-header.
  - `_sum_report_totals` aggregates strictly over primary rows (`not it.is_child_sku`).

#### C. `src/discrepancy_classifier.py` (Service 3 — 0 External Dependencies)
- **Enum (`DiscrepancyType`, `lines 15–24`):** `EXACT_MATCH`, `SHORTAGE`, `OVERAGE`, `REVUELTO_CANDIDATE`, `CARTON_MISMATCH_ONLY`, `DAMAGE_ONLY`, `UNEXPECTED_ITEM`, `MISSING_ITEM`.
- **Classification Precedence (`lines 92–112`):** `MISSING_ITEM` (`decl > 0, phys == 0`) -> `UNEXPECTED_ITEM` (`decl == 0, phys > 0`) -> `SHORTAGE` (`diff_u < 0`) -> `OVERAGE` (`diff_u > 0`) -> `DAMAGE_ONLY` (`diff_u == 0, dam > 0`) -> `CARTON_MISMATCH_ONLY` (`diff_u == 0, diff_c != 0`) -> `EXACT_MATCH`.
- **Cardinal Discrepancy Rule (`compare_merchandise`, `lines 214–260`):** Sums `total_shortage_units = sum(abs(diff_u) for diff_u < 0)` and `total_overage_units = sum(diff_u for diff_u > 0)` independently—shortages on one material never offset overages on another.
- **Revuelto Detection (`_detect_revuelto_pairs`, `lines 180–211`):** Greedy 1:1 exact-delta pairing (`o.difference_units == abs(s.difference_units)`) linking `revuelto_partner` and assigning `REVUELTO_CANDIDATE` without altering `difference_units` or aggregate shortage/overage totals.
- **C# Target Enhancements (`R2`):** Add explicit `RevueltoConfidence` (`Confirmed`, `Probable`, `Possible`) with evidence rationale rather than silent resolution, and handle duplicate `material_code` entries across multiple PO lines without silent dictionary key overwrite (`lines 222–223`).

#### D. `src/closure_exporter.py` (Service 4)
- **Validation (`validate_closure_document`, `lines 38–128`):** Opens `read_only=True, data_only=False`, verifies `"CIERRE RECEP"` presence, scans all cells for `("#REF!", "#VALUE!", "#N/A", "#DIV/0!", "#NAME?")` (up to 5 samples), and checks physical sums in `AS` (col 45) / `AT` (col 46) from row `13+` for numeric literals.
- **Export (`export_receiving_closure`, `lines 131–217`):** Copies template via `shutil.copy2`, opens `keep_vba=True`, writes `B5` (`total_declared_cartons`), `B6` (`total_physical_cartons`), `B7` (`=+B6-B5`), scans `range(13, 300)` matching Col `E` (`material_code`), and writes `H`, `J:M`, `AS:AT`, `AU`, `AW:AY`.
- **Identified Parity Gaps for C# `ClosureExporter` (`R3`):**
  1. Hardcodes `range(13, 300)` (insufficient when blank/subtotal rows or large containers push rows past `299`; `max_row` reaches `612` in `IMP-187`).
  2. Does not populate `'Bultos'` sheet or `AZ` / Zero-Carton `"0"` rules.
  3. Does not distinguish pre-existing template errors, newly introduced exporter errors, and legitimate optional `VLOOKUP` `#N/A` outcomes, nor check both formula strings and cached evaluation values.
  4. `openpyxl` `keep_vba=True` repackages the `.xlsm` archive; C# must explicitly verify `vbaProject.bin` SHA-256 equality before and after export.

### 2.3 Python Test Suite Execution Baseline (`tests/`)
- **Bare `pytest -v --tb=short`:** Fails collection (exit code `1` / `2`, `9 errors`: `ModuleNotFoundError: No module named 'src'`), due to absent `src/__init__.py`, `tests/__init__.py`, `pyproject.toml`, or `pytest.ini`.
- **Module invocation (`python -m pytest -v --tb=short`):** Passes **`58 passed in 43.14s`** (`0 failed`, `0 skipped`, `0 errors`) in strictly read-only mode against `data/`:
  - **51/51** real QC files parsed (`5.223s`, `0.1024s` avg/file, `23/51` flagged, `50/51` = `98.04%` exact carton match; sole mismatch `imp298 control calidad ver2.xlsm`: `1036` calculated vs `1032` reported; `IMP-103` verified `267` items / `1038` cartons).
  - **77/77** real closure workbooks opened and validated (`34.874s`, `0.4529s` avg/file), flagging **3/77** under `data_only=False` (`IMP-241.xlsm`, `IMP-244.xlsm`, `IMP-256.xlsm`).

---

## 3. Production Dataset Corpus Audit

Read-only filesystem and OpenXML (`xl/workbook.xml`, `xl/vbaProject.bin`, `xl/styles.xml`, `xl/_rels/workbook.xml.rels`, `xl/sharedStrings.xml`) traversal across `D:\Development\Software Development\3 Way Inventory Reconciliation\data` inventoried **830** directory/file entries (`71+` `IMP-*` folders + root templates):
- `.xlsm`: `147` readable `.xlsm` workbooks plus `1` Office/OneDrive `~$` lockfile (`148` `.xlsm`-extension entries total; lockfile `IMP-349 PEND CONTROL\~$IMP-349 (BORRADOR)  AQUA FASHION S.A..xlsm` is not a valid ZIP and **must always be skipped via `~$` prefix filtering**).
- `.xlsx`: `97` (invoices, packing lists, `Portada *.xlsx`).
- `.xls`: `62` (legacy OLE2/BIFF8 `d0cf11e0a1b11ae1` vendor invoices/packing lists inside `Factura*` directories; 0 QC/closure files).
- Other formats: `.pdf` (`160`: ~130 digital/hybrid, ~30 image-only scans), `.msg` (`71`), `.eml` (`35`), `.jpeg` (`12`), `.db` (`4`), `.zip` (`2`), `.pptx` (`1`), `.a` (`1`).

### 3.1 Reconciling QC & Closure Workbook Counts (`51` vs `64`; `77` vs `81`)

| Corpus Category | `test_real_dataset_audit.py` Filename Filter | Sheet-Signature Reality on Disk | Exact Reconciliation |
|---|---|---|---|
| **QC Workbooks (`CONTROL DATA`)** | **51** matched (`"control calidad"`, `"calidad"`, or `"qc"` in filename) | **63** standalone QC (`CONTROL DATA`, no `CIERRE RECEP`) + **1** hybrid (`IMP-218\IMP-218.xlsm`, containing both `CIERRE RECEP` and `CONTROL DATA`) = **64** total | **3 false positives** in the 51-list (`IMP-184 sin control calidad.xlsm`, `IMP-185 (sin control calidad).xlsm`, `IMP 211 (SIN CONTROL CALIDAD).xlsm`) actually contain `CIERRE RECEP` (not `CONTROL DATA`) and yield `0` QC items via `Check list RECEPCION` fallback. **15 genuine QC workbooks** were missed because filenames omit `"calidad"`/`"qc"` (e.g., `IMP 81.xlsm`, `imp 207.xlsm`, `IMP244.xlsm`, `IMP 233.xlsm`, `IMP 406.xlsm`, `imp 84.xlsm`). |
| **Master Closures (`CIERRE RECEP`)** | **73** matched (`imp-` prefix, excluding `control calidad`/`copia`/`sin control`) | **81** workbooks contain `CIERRE RECEP` (**77** excluding `4` `copia`/`sin control` variants) | Excluding `4` variant files (`Copia de IMP-182 inprogress.xlsm`, `IMP-184 sin control calidad.xlsm`, `IMP-185 (sin control calidad).xlsm`, `IMP 211 (SIN CONTROL CALIDAD).xlsm`) leaves **77** closures. The 4 valid non-`imp-` closures missed by `startswith("imp-")` are `Macro para cierre hannia.xlsm`, `IMP-184\SDF.xlsm`, `IMP-295\Macro para cierre hannia.xlsm`, and `IMP-83\CIERRE 83.xlsm`. |
| **Other Workbooks** | — | **100** workbooks (`3` `.xlsm` EKONO distribution files + `97` `.xlsx`) | Contain neither `CONTROL DATA` nor `CIERRE RECEP`. |

**Design Invariant for `Acr.Infrastructure`:** File classification MUST inspect worksheet signatures (`CONTROL DATA` vs `CIERRE RECEP`), never filename substrings alone, and MUST filter out `~$` lockfiles before opening ZIP/OpenXML streams.

### 3.2 QC Workbook Characteristics (63 Standalone + `IMP-218`)
- **Format & Macros:** `63/63` `.xlsm` with `xl/vbaProject.bin` (`62` distinct SHA-256 hashes) and `xl/sharedStrings.xml` (`uniqueCount` `112`–`623`).
- **81 MB Trailing XML Bloat:** `60/63` `CONTROL DATA` worksheets have **69.6 MB – 81.9 MB** of uncompressed `sheet1.xml` (e.g., `IMP-100`: `83,294,222` bytes for `19` data rows, where all actual `<f>` and `t="s"` cells end by byte `47,009`, followed by ~83 MB of empty pre-formatted grid XML). Only `3` files are `< 1 MB` (`IMP-104`: `443 KB`, `imp - 209`: `348 KB`, `IMP 91`: `197 KB`). **Mandate:** `QcReader` must use forward-only streaming `OpenXmlReader` with early empty-row termination (`15` consecutive blank rows) rather than DOM loading.
- **Merged Cells & Formulas:** `52` files have `11` merged ranges and `11` files have `10` merged ranges—all confined to Rows `1–5` metadata/group banners, never in item rows (`>= 7`). All `63/63` contain formulas (`17` to `9,330`), and some workbooks store `BULTOS ENTEROS` as a formula (e.g., `IMP 188-189`: `=W7/AD7`; `IMP-88`: `=X7/AD7`).
- **Pre-Existing QC Formula Errors:** `IMP-327` (`cierre imp 327 Control de calidad.xlsm`) has `20` cached `#VALUE!` errors in `AD13..AD32` (`=F13/AC13` on summary/footer rows); `IMP-348` (`imp 348.xlsm`) has `=IFERROR(VLOOKUP(E7,#REF!,3,0),0)` (evaluates to `0`).
- **15 Row-6 Header Layouts:** All `63/63` place headers on Row 6, but operators inserted/deleted columns (`Tarima`, `MASTER PACK`, `INER PACK`, `cantidad a 48`, blank spacers before `Material` or `TOTAL DE BULTOS`, or omitted `BULTOS COMPLEMENTOS`), yielding **15 distinct signatures**:
  - Signature #1 (`39/63` files, canonical): Col `A` `Numero de Linea`, `B` `Tarima`, `C` `CATEGORIA`, `D` `Material proveedor`, `E` `Material`, `F` `CANTIDAD DE UNIDADES`, `I` `CURVA SISTEMA`, `K` `CURVA FISICA`, `L` `UM`, `M` `DESCRIPCIÓN`, `AC` `BULTOS ENTEROS`, `AD` `UNIDADES POR BULTOS COMPLETOS`, `AE` `BULTOS COMPLEMENTOS`, `AF` `UNIDADES BULTOS COMPLEMENTOS`, `AG` `TOTAL DE BULTOS`, `AH` `TOTAL DE UNIDADES`, `AI` `TOTAL PRK`, `AJ` `CURVA`, `AK` `DIFERENCIA EN UNIDADES`, `AL` `OBSERVACIONES`.
  - Signature #2 (`7/63`): Omits `Tarima` (shifts cols left by 1).
  - Signature #3 (`4/63`): Inserts blank spacer before `TOTAL DE BULTOS` (shifts `AG..` to col 34+).
  - Signature #4 (`2/63`, `IMP-207`): Deletes `BULTOS COMPLEMENTOS` (shifts `AF..` left).
  - Signatures #5–#15 (`1/63` each): `imp 103` (`NUMERO DE TARIMA`), `imp 184` (2 versions), `IMP 196` (Col 32 header `None`, Col 31 `UNIDADES BULTOS COMPLEMENTOS`), `IMP244` (`MASTER PACK`/`INER PACK`), `IMP 257`, `IMP 406` (`cantidad a 48`), `imp 84` (`Material` shifted to `G`), `IMP 91`, `imp98`.
- **Duplicate/Versioned QC Files:** `IMP-103`, `IMP-184`, `IMP-207`, and `IMP-298` each contain two QC workbook versions in the same `Informacion General` folder.

### 3.3 Master Closure Workbook Characteristics (81 `CIERRE RECEP` Workbooks)
- **Format & VBA Preservation:** `81/81` are `.xlsm` containing `xl/vbaProject.bin`, producing **80 distinct `vbaProject.bin` SHA-256 hashes** (only the two unedited `1,464 KB` templates `Macro para cierre hannia.xlsm` and `IMP-349 (BORRADOR)  AQUA FASHION S.A..xlsm` share `792e5613f5bfa325b5069a5e0739ae34696595b9d8869a8f0b81c7ec4565b7ad`). Every saved closure carries its own compiled VBA state, confirming `ClosureExporter` must preserve `vbaProject.bin` verbatim (SHA-256 before == SHA-256 after).
- **Styles, Relationships & Named Ranges:** `51` distinct `xl/styles.xml` SHA-256 hashes and `9` distinct `xl/_rels/workbook.xml.rels` hashes across `81` files. `definedNames` breakdown: `74/81` have `19` (`_xlnm._FilterDatabase`, `_xlnm.Print_Area`); `3/81` (`IMP-81`, `IMP-182 inprogress`, `Copia de IMP-182 inprogress`) have `7` (including `JK`, `Z_*.wvu.FilterData`); `3/81` (`IMP-200`, `IMP-207 NCR`, `IMP-216`) have `20` (`_xleta.VLOOKUP`); `1/81` (`IMP-218`) has `25` (`_xleta.DATE`/`IF`/`TODAY` plus embedded QC sheet filter databases).
- **Row Counts (`max_row` vs `>299` / `>500` limits):** `IMP-187` has `231` item rows (`max_row = 612`), `IMP-185` has `202` (`583`), `IMP-298 (borrador)` has `168` (`549`), `IMP-105` & `IMP-97` have `167` (`548`), and **10** populated closures exceed `max_row = 500` (`505..612`) — plus **2** unpopulated `1,380`-row blank templates (`Macro para cierre hannia.xlsm` and `IMP-349 (BORRADOR)  AQUA FASHION S.A..xlsm`, making **12** in Survey 2 JSON), and **13** total on disk including `IMP-81\IMP-81.xlsm` (`1,000` rows, where `read_only=True` yields `max_row = None` due to an omitted `<dimension>` tag).

---

## 4. Actual `CIERRE RECEP` 58-Column (`A`–`BF`) Schema & Applicability

All **81/81** `CIERRE RECEP` sheets share an identical **58-column (`A`–`BF`, 1-based indices `1..58`)** structure on **Row 12** (`53` labeled headers, `2` unlabeled formula-echo columns `AQ`/`AR` where `AR` carries `'AP'` in Row 11, and `3` structural blank spacer columns `AP`/`AV`/`BA`).

**Reconciliation with `"45 columns"` Reference:** Column `AS` (`TOTAL BULTOS DESCARGA`, the primary physical unload column where exporter strips begin) is **Column 45** (`A..AS`), and `58` total columns minus `13` optional/discrepancy/notes/spacer columns equals `45` core populated columns.

### 4.1 Complete Verified 58-Column `CIERRE RECEP` Schema (Sampled Rows 13–22 Across All 81 Closures)

| Col | Ltr | Row 12 Header | Applicability | Dominant Formula / Value Pattern & Role |
|---:|:---:|---|---:|---|
| 1 | `A` | `ORDEN DE COMPRA` | 79/81 | Literal PO (`4300006910`); empty in `IMP-182` drafts |
| 2 | `B` | `Proveedor` | 81/81 | `=IF(ISNA(VLOOKUP(AM{r},'PROV Y CATEG'!$A$2:$B$194,2,0))=TRUE,"",VLOOKUP(...))` or literal |
| 3 | `C` | `Categoría` | 81/81 | `=IF(ISNA(VLOOKUP(AO{r},'PROV Y CATEG'!$F$2:$G$76,2,0))=TRUE,"",VLOOKUP(...))` or literal |
| 4 | `D` | `MAT. PROVEEDOR` | 79/81 | Literal vendor style |
| 5 | `E` | `MATERIAL` | 79/81 | Literal SAP Material code (**Primary row-anchoring key**) |
| 6 | `F` | `CANTIDAD PEDIDO 43` | 81/81 | Formula `=BD{r}*BB{r}` (ordered units on PO 43) |
| 7 | `G` | `CANTIDAD PEDIDO 41` | 81/81 | Formula `=BE{r}*BB{r}` (invoiced units on Pedido 41) |
| 8 | `H` | `RECIBIDO REAL & FISICO` | 81/81 | **Exporter guard/repair target:** Formula `=+G{r}-J{r}-N{r}-O{r}+L{r}+P{r}+Q{r}` |
| 9 | `I` | `CANTIDAD INGRESO MOV-101` | 81/81 | Formula `=BF{r}*BB{r}` (SAP Mov 101 net receipt units) |
| 10 | `J` | `REVUELTO SALIENDO` | 28/81 | **Exporter target:** Optional numeric (units swapped out) |
| 11 | `K` | `DESTINO` | 0/81 (in sample)* | **Exporter target:** Optional destination material code |
| 12 | `L` | `REVUELTO ENTRANDO` | 31/81 | **Exporter target:** Optional numeric (units swapped in) |
| 13 | `M` | `ORIGEN` | 0/81 (in sample)* | **Exporter target:** Optional origin material code |
| 14 | `N` | `FALTANTE` | 28/81 | Optional numeric (shortage units) |
| 15 | `O` | `DAÑADO` | 27/81 | Optional numeric (damaged units) |
| 16 | `P` | `SOBRANTE` | 28/81 | Optional numeric (excess units) |
| 17 | `Q` | `UNIDADES CON MOV 501 EN TIENDA` | 26/81 | Optional numeric |
| 18 | `R` | `# DE MOVIMIENTO` | 0/81 (in sample)* | Optional string/number |
| 19 | `S` | `INDICADOR  DE  INCONSIST` | 81/81 | Formula `=+J{r}+L{r}+N{r}+O{r}+P{r}` (or `=SUM(J{r}:Q{r})`) |
| 20 | `T` | `ACCIONES A SEGUIR U OBSERVACIONES` | 0/81 (in sample)* | Optional text |
| 21 | `U` | `CANTIDAD FACTURADA MOV-641` | 79/81 | Numeric (`0` or Mov 641; `'X'` text error in `IMP-360!U14`) |
| 22 | `V` | `DIFERENC MOV 101 VRS MOV 641` | 81/81 | Formula `=+U{r}-I{r}` |
| 23 | `W` | `MATERIAL NUEVO EN CREACIONES` | 0/81 (in sample)* | Optional text |
| 24 | `X` | `CANTIDAD ENVIADA  A PARTES ( B-C-D )` | 81/81 | Numeric (`0` or deferred units) |
| 25 | `Y` | `STOCK  A004 RESURTIDO` | 81/81 | Numeric (`0` or reserve stock) |
| 26 | `Z` | `UNIDADES EN PROCESO DE FACTURAR (643)` | 81/81 | Numeric (`0` or Mov 643) |
| 27 | `AA` | `UNIDADES QUE APLICAN PARA NOTA DE CREDITO` | 81/81 | Formula `=+N{r}+O{r}` |
| 28 | `AB` | `REVISION FINAL DE STOCK` | 81/81 | Formula `=V{r}+X{r}+Y{r}+Z{r}` |
| 29 | `AC` | `OBSERVACIONES GENERALES` | 0/81 (in sample)* | Optional text |
| 30 | `AD` | `DESCRIPCION (TEXTO BREVE)` | 79/81 | Literal string |
| 31 | `AE` | `VALOR UNITARIO` | 81/81 | Formula `=IFERROR(AK{r}/F{r},0)` |
| 32 | `AF` | `Valor unitario PEDIDO NAC $` | 81/81 | Formula `=IFERROR(AL{r}/G{r},0)` |
| 33 | `AG` | `Valor Revuelto` | 81/81 | Formula `=+J{r}*AE{r}` |
| 34 | `AH` | `Valor faltante` | 81/81 | Formula `=+N{r}*AE{r}` |
| 35 | `AI` | `Valor Dañado` | 81/81 | Formula `=+O{r}*AE{r}` |
| 36 | `AJ` | `Valor Sobrante` | 81/81 | Formula `=+P{r}*AE{r}` |
| 37 | `AK` | `VALOR TOTAL PEDIDO 43` | 79/81 | Numeric currency total |
| 38 | `AL` | `VALOR TOTAL PEDIDO 41` | 74/81 | Numeric currency total |
| 39 | `AM` | `COD PROVEEDOR` | 79/81 | Numeric SAP vendor ID (`1100000577`) |
| 40 | `AN` | `CATEGORIA` | 79/81 | String merchandise hierarchy (`G40403001`) |
| 41 | `AO` | `COD CATEGORIA` | 81/81 | Formula `=MID(AN{r},1,2)` |
| 42 | `AP` | *(None — spacer)* | 0/81 | Always empty |
| 43 | `AQ` | *(None — helper)* | 81/81 | Formula `=+D{r}` |
| 44 | `AR` | *(None — helper; r11=`'AP'`)* | 81/81 | Formula `=+E{r}` (`VLOOKUP` lookup key) |
| 45 | `AS` | `TOTAL BULTOS DESCARGA` | 74/81 | **Exporter target:** Numeric or compound string (`'39*60 1*78'`); `VLOOKUP` in `IMP-244` |
| 46 | `AT` | `CANTIDAD DE UNIDADES EN BULTOS` | 74/81 | **Exporter target:** Numeric or `VLOOKUP` (`IMP-90`, `IMP-244`) |
| 47 | `AU` | `DIFERENCIA UNIDADES FISICO VS PEDIDO NAC` | 81/81 | Formula `=+AT{r}-H{r}` (preserve formula, never overwrite with `"0"`) |
| 48 | `AV` | *(None — spacer)* | 26/81 | Empty on item rows; `=SUM(AV13:AV18)` in totals rows |
| 49 | `AW` | `Bultos según Paking List` | 74/81 | **Exporter target:** Numeric, compound string, or `=+AS{r}` |
| 50 | `AX` | `CANTIDAD UNIDADES` | 74/81 | **Exporter target:** Numeric or `=+AT{r}` |
| 51 | `AY` | `Empaque` | 74/81 | **Exporter target:** Numeric, compound string (`'9*48+12'`), or `=+AX{r}/AW{r}` |
| 52 | `AZ` | `Diferencia unidades facturadas vs solicitadas` | 81/81 | Formula `=+AX{r}-H{r}` (preserve formula, never overwrite with `"0"`) |
| 53 | `BA` | *(None — spacer)* | 0/81 | Always empty |
| 54 | `BB` | `Curva` | 81/81 | Numeric (`1` for `UN`) |
| 55 | `BC` | `UM` | 79/81 | String (`'PRK'` or `'UN'`) |
| 56 | `BD` | `CANTIDAD PEDIDO 43` | 79/81 | Numeric (PRK/UN count for PO 43) |
| 57 | `BE` | `CANTIDAD PEDIDO 41` | 74/81 | Numeric (PRK/UN count for Pedido 41) |
| 58 | `BF` | `CANTIDAD INGRESO MOV-101` | 76/81 | Numeric (PRK/UN count for SAP Mov 101) |

*\*Columns `K`, `M`, `R`, `T`, `W`, `AC` sampled `0/81` inside rows `13..22` because exception/NCR annotations occur sparsely on later rows (e.g., `IMP-207` Row `38`).*

**Header Summary Cells (`CIERRE RECEP!B5:B7` and `'Bultos'` Sheet):**
- `B5`: Total theoretical/Packing List cartons (`=AW{total_row}` or literal).
- `B6`: Total physical/unload cartons (`=AS{total_row}` or literal).
- `B7`: `=+B6-B5`.
- `'Bultos'` (10 columns `B..J`, rows `3+`): Links `B`=`C{r}`, `C`=`AS{r}`, `D`=`AT{r}`, `E`=`E{r}`, `F`=`H{r}`, `G`=`BB{r}`, `H`=`BC{r}`, `I`=`AD{r}`, `J`=`AY{r}`.

### 4.2 3-Way Formula Error Audit (All 81 Workbooks, Rows 1–1500)
Scanning both raw formulas and cached values across all `81` workbooks identifies **5** pre-existing error workbooks (`0` workbooks had `*0` or `#REF!` in Col `H`):

| Workbook | Error Count | Exact Cells & Root Cause | 3-Way Audit Classification (`R3`) |
|---|---:|---|---|
| `IMP-241 Sin Control Calidad\IMP-241.xlsm` | 5 | `B5: =#REF!`, `D5: =IFERROR(#REF!/B5,"")`, `B6: =#REF!`, `D6: =IFERROR(#REF!/B6,"")`, `B7: =+B6-B5` (`#REF!`). Rows `13+` are clean. | **Pre-existing source error** (missed by row-13-only scanners) |
| `IMP-244\IMP-244.xlsm` | 51 | `B5: =AW19`, `B6: =AS19`, `B7: =+B6-B5`, and rows `13..18` `AS`/`AT`/`AY` `=VLOOKUP(AR{r},#REF!,2/3/5,0)` cascading to `AU`/`AW`/`AX`/`AZ` and Row 19 totals. | **Pre-existing source error** (deleted `VLOOKUP` sheet) |
| `IMP-256\IMP-256.xlsm` | 1 | `AY29`: literal string `'#N/A'`. | **Pre-existing source error** (literal pasted `#N/A`) |
| `IMP-360\IMP-360.xlsm` | 4 | `U14 = 'X'` causes `V14 = =+U14-I14` (`#VALUE!`), cascading to `AB14`, `V63`, `AB63`. | **Pre-existing source error** (text `'X'` in numeric column) |
| `IMP-90\IMP-90.xlsm` | 4 | `AT16 = =VLOOKUP(AR16,[2]Hoja1!A$1:C$5,3,0)` (`#N/A` for material `5033815002`), cascading to `AU16`, `AT17`, `AU17`. | **Pre-existing unguarded external VLOOKUP `#N/A`** (by contrast, template Col `B`/`C` wrap `IF(ISNA(VLOOKUP(...))=TRUE,"",VLOOKUP(...))` so missing optional fields cleanly evaluate to `""`) |

**Exporter Audit Contract:** `ClosureExporter` must (a) scan Rows `1..max_row` on both source template and generated output to separate `PreExistingSourceError` from `ExporterIntroducedError` (`0` allowed), (b) classify guarded/optional `VLOOKUP` `#N/A` outcomes (`LegitimateOptionalNa`) separately from broken `#REF!` / `#VALUE!`, and (c) verify both formula string presence (`=`) and cached evaluation state.

---

## 5. Daily Log (`Control_Diario_Operaciones_Jostin.xlsx`) Schema, Lifecycle & Transitions

- **File Metadata:** `47,126` bytes; `dc:creator = 'Angie Montero Rojas'`, `dcterms:created = '2026-11-14T18:04:02Z'`; `0` VBA macros; `3` Excel Tables (`tblMaestroContenedores`, `tblParteB`, `tblNotificaciones`); `0` defined names.
- **5 Visible Sheets:** `Dashboard General`, `Maestro Contenedores` (`A5:R27`, `22` data rows `6..27`, validations `6..57`, conditional formatting `I6:O1000`), `Control Parte B y Bultos` (`tblParteB` `B5:H15`, 10 deferred containers `IMP-51, 125, 69, 289, 130, 290, 200, 126, 260, 314`, plus 3-level calculator `B18:L24`), `Comunicados y Guia SAP` (`tblNotificaciones` `B13:G19` + SAP prefix/8-step guide), and `Catalogos` (`A1:H7`).

### 5.1 `Maestro Contenedores` Column Schema (`A5:R5`)
- `A`: `N°` | `B`: `Contenedor (Ubic. EWM)` | `C`: `Fecha Act.` (`mm-dd-yy`) | `D`: `PO Orden Compra (430000...)` (slash-separated or `'-'`) | `E`: `Pedido Nacionaliz. (410000...)` | `F`: `Entrega VL31N (1800...)` | `G`: `Doc. Entrega EWM (Tarea Alm.)` (`414xx`) | `H`: `Unidad Manipulación HU (18 díg)` | `I–N`: **6 Lifecycle Stage Dropdowns** | `O`: `Estado Maestro Actual (Auto)` | `P`: `Observaciones y Contexto Operativo` | `Q`: `Saludo` (`Buenos días` | `Buenas tardes` | `-`) | `R`: `Notificación Automática SAP`.

### 5.2 Exact 6 Real Lifecycle Stage Names & `Catalogos` Dropdown Values

| Stage | `Dashboard General!B13:B18` Name | `Maestro Contenedores!I5:N5` Header | `Catalogos` Range & Header | Exact Validated Dropdown Values (`Catalogos`) | Completed Target |
|---:|---|---|---|---|---|
| **1** | `1. Aduana, DUA y Radar` | `1. Aduana y DUA` (`I`) | `A2:A4` (`1. Aduana y DUA`) | `DUA y Docs Listos`, `En Radar / Tránsito`, `Pendiente Docs` | `DUA y Docs Listos` |
| **2** | `2. Ingreso SAP / EWM (VL31N & RFUI)` | `2. Ingreso SAP / EWM` (`J`) | `B2:B5` (`2. Ingreso SAP / EWM`) | `Listo y Notificado`, `Pendiente Ingreso SAP`, `En Proceso EWM/RFUI`, `Sin Doc. 18... Aún` | `Listo y Notificado` |
| **3** | `3. Montaje a Ciegas (Doc. Proveedor)` | `3. Montaje a Ciegas` (`K`) | `C2:C5` (`3. Montaje a Ciegas`) | `Listo (Solo Doc. Proveedor)`, `Listo (Completo)`, `En Cola A Ciegas`, `Pendiente` | `Listo*` |
| **4** | `4. Control Calidad Muelle (ZDescarga)` | `4. Control Calidad (Muelle)` (`L`) | `D2:D4` (`4. Control Calidad`) | `CC Aprobado (ZDescarga)`, `Pendiente Revisión Final CC`, `Sin CC (Espera Muelle)` | `CC Aprobado (ZDescarga)` |
| **5** | `5. Cierre Recepción 3-Way (Mov 101)` | `5. Cierre Recepción (Mov 101)` (`M`) | `E2:E7` (`5. Cierre Recepción (101)`) | `Cuadrado (0 Diferencias)`, `Pausa sin Control Calidad`, `Pendiente Revisión CC`, `Parcial / Pendiente Parte B`, `Discrepancia / NCR (40...)`, `Pendiente Cierre` | `Cuadrado (0 Diferencias)` |
| **6** | `6. Cierre Despacho Tiendas (Mov 641)` | `6. Cierre Despacho (Mov 641)` (`N`) | `F2:F5` (`6. Cierre Despacho (641)`) | `Cierre Despacho Listo`, `Pre-Cableado 80% (Listo Mov 641)`, `En Distribución Tiendas`, `Bloqueado por Recepción` | `Cierre Despacho Listo` |

Additional catalogs: `Saludos` (`G2:G4`: `Buenos días`, `Buenas tardes`, `-`) and `Estado Parte B` (`H2:H4`: `Pendiente Arribo`, `En Muelle / Descarga`, `Conciliado 0 Dif`).

### 5.3 `Col O` Auto-State Formula, Allowed Transitions & SAP Notification Template
`Maestro Contenedores!O{r}` computes a deterministic 7-state portfolio rollup from Stage 6 down to Stage 2:
```excel
=IF(N6="Cierre Despacho Listo","7. Cierre Total (Recep + Desp)",
 IF(M6="Cuadrado (0 Diferencias)","6. Recepción Cerrada (Listo Recep.)",
 IF(ISNUMBER(SEARCH("Revisión",L6)),"5. En Auditoría Control Calidad",
 IF(M6="Pausa sin Control Calidad","4. Pausa sin CC (Montado a Ciegas)",
 IF(J6="Listo y Notificado","3. Ingresado SAP (Espera Cierre)",
 IF(J6="Pendiente Ingreso SAP","2. Doc SAP Listo (Pend. Ingreso)",
 "1. En Cola Conteo a Ciegas"))))))
```

**Allowed Lifecycle Transitions & Guardrails (derived from `Dashboard General`, `Col O`, and 22 active rows):**
1. **Entry / Pre-SAP (`1. En Cola Conteo a Ciegas` / `2. Doc SAP Listo`):** Stage 5 = `Pendiente Cierre`; Stage 6 = `Bloqueado por Recepción`.
2. **SAP Inbound Ready (`3. Ingresado SAP (Espera Cierre)`):** Stage 2 -> `Listo y Notificado` normally requires `Entrega VL31N (1800...)`, `Doc. Entrega EWM (414xx)`, and `HU (18 díg)`. *(Historical retroactive rows `IMP-195`, `IMP-217`, `IMP-259`, `IMP-349` store `"-"` for `18...`/`HU`; ACR should allow retroactive persistence with a non-blocking warning.)*
3. **Blind Setup Pause (`4. Pausa sin CC (Montado a Ciegas)`):** When Stage 3 is `Listo (Solo Doc. Proveedor)` (or `Listo (Completo)`) while Stage 4 is `Sin CC (Espera Muelle)`, Stage 5 **must** transition to `Pausa sin Control Calidad` and Stage 6 stays `Bloqueado por Recepción`.
4. **QC Audit (`5. En Auditoría Control Calidad`):** When QC arrives (Stage 4 -> `Pendiente Revisión Final CC`), Stage 5 moves from `Pausa sin Control Calidad` / `Pendiente Cierre` to `Pendiente Revisión CC` (or `Discrepancia / NCR (40...)` / `Parcial / Pendiente Parte B`).
5. **Reception Closed (`6. Recepción Cerrada (Listo Recep.)`):** Stage 5 -> `Cuadrado (0 Diferencias)` requires Stage 1 = `DUA y Docs Listos` and Stage 4 = `CC Aprobado (ZDescarga)`. *(Historical `IMP-217` Row 8 shows an Excel override with `Cuadrado` while Stage 4 still mentions pending review; ACR must flag this inconsistency.)*
6. **Dispatch Unlocked & Closed (`7. Cierre Total (Recep + Desp)`):** Stage 6 cannot leave `Bloqueado por Recepción` (`-> Pre-Cableado 80% (Listo Mov 641) -> En Distribución Tiendas -> Cierre Despacho Listo`) until Stage 5 reaches `Cuadrado (0 Diferencias)` or `Parcial / Pendiente Parte B`.

**SAP Intake Notification Template (`Maestro Contenedores!R{r}` & `Comunicados y Guia SAP!C8:C9`):**
- Formula (`Col R`): `=IF(F{r}<>"-", Q{r} & ", Se informa que el ingreso a nivel de sistema del Contenedor " & SUBSTITUTE(B{r}, "IMP-", "") & ", quedo listo.", "-")`
- Suggested Email Subject (`C8`): `Ingreso Listo a Sistema - Contenedor {N}`
- Clipboard Body (`C9`): `{Buenos días|Buenas tardes}\nSe informa que el ingreso a nivel de sistema del Contenedor {N}, quedo listo.` (verified verbatim against `.eml` files in `data/IMP-181`).

### 5.4 Observed Data Integrity Exceptions in `Control_Diario_Operaciones_Jostin.xlsx`
1. **Duplicate Container (`IMP-202`):** Present in both **Row 6** (`N°=1`, `2026-09-26`, `Entrega=180060834`, `EWM=41461`, `HU=112345678005808817`) and **Row 21** (`N°=16`, `2026-11-10`, `POs=4300006692...`, `Ped41=4100002332`, `Entrega='-'`), contradicting `Dashboard General!B8` (`"0 duplicados en bitácora"`).
2. **Shared PO/Pedido across Different Containers:** Row 23 (`IMP-320`) and Row 24 (`IMP-213`) share `PO = '4300007748'` and `Pedido = '4100002329'`.
3. **Stale Hardcoded Dashboard Subtitle (`J8`):** `Dashboard General!J8` reads `"IMP-217 e IMP-259"`, whereas `J7` (`COUNTIF(...,"*Revisión*")`) counts `IMP-259` (Row 14) and `IMP-201-1` (Row 18); `IMP-217` (Row 8) is already `CC Aprobado (ZDescarga)`.
4. **Stray Formula-Text Artifact (`Dashboard General!B45`):** Contains literal text `"(AC22*AD22) + SI(AE22=\"\"; AF22; AE22*AF22)"`.
5. **Table Range vs Validation Range:** `tblMaestroContenedores` ref ends at `A5:R27`, whereas data validations cover rows `6..57` and conditional formatting covers `I6:O1000`. `DailyLogRepository` must dynamically resize `tblMaestroContenedores.ref` (and extend data validations beyond row `57` as rows grow).

---

## 6. Historical Sync Baseline

Scanning `docProps/core.xml` and worksheet cells (`PORTADA`, `DATOS GENERALES`, `Check list RECEPCION`, `Check list DESPACHO`) across all `81` closures found **20** workbooks referencing Jostin, resolving to **18 unique containers** plus **2 duplicate/draft copies of `IMP-182`**:
- **Authorship Rule:** `docProps/core.xml` `dc:creator` is `'Hilario Lopez Rivera'` (`2016-09-05T16:30:28Z`, corporate template creator) in `19/20` files (only `IMP-364` has `dc:creator = 'Jostin Lopez'`; only `IMP-218` has `cp:lastModifiedBy = 'Jostin López'`). True operational authorship resides strictly in cell **`'DATOS GENERALES'!B33`** (`ELABORADO POR: \nJostin Lopez Sobalbarro` in `19` files, and **`Jostin Lopez Sobalbaro`** — single `r` typo, missing second `r`, two `b`s — in **`IMP-326\IMP-326.xlsm`**). **Scanner Guardrail (`IMP-209` Edge Case):** `IMP-209\IMP-209.xlsm` contains `"Jostin Lopez Sobalbarro"` inside `xl/sharedStrings.xml` while `'DATOS GENERALES'!B33` is `"Angie Montero Rojas"`; therefore Phase 2 `HistoricalClosureScanner` MUST read `'DATOS GENERALES'!B33` directly and NEVER scan `sharedStrings.xml`.

### 6.1 Complete 20-File Historical Sync Baseline (Business Identity vs Provenance)

| # | Business Identity (`DG!C7`) | Pedido 41 (`DG!C9`) | PO(s) `DG!C8` (`CR!A` distinct) | Fecha Rev. (`DG!C30`) | Bultos Teó/Rec (`DG!C13/14`) | Líneas `DG!C18` (`CR` items) | Provenance (Relative `.xlsm` Path) | Exceptions / Review Flags |
|---:|---|---|---|---|---|---|---|---|
| 1 | `IMP-104` | `4100002248` | `4300008021-...` (20) | `2026-11-03` | `660 / 660` | `124 (124)` | `IMP-104\IMP-104.xlsm` | Clean |
| 2 | `IMP-181` | `4100002271` | `4300007257` (1) | `2026-10-27` | `517 / 517` | `4 (4)` | `IMP-181\IMP-181.xlsm` | Clean |
| 3 | `IMP-182` | `4100002283` | `4300006764-...` (9) | `2026-11-03` | `927 / 927` | `136 (136)` | `IMP-182\IMP-182.xlsm` | Clean final |
| *D1* | *`IMP-182` (dup)* | `4100002283` | `4300006764-...` (`0` in CR) | `2026-11-02` | `17172 / 14` | `136 (136)` | `Copia de IMP-182 inprogress.xlsm` | Duplicate draft; empty CR Col A; `-17158` diff |
| *D2* | *`IMP-182` (dup)* | `4100002283` | `4300006764-...` (`0` in CR) | `2026-11-03` | `1133 / 14` | `136 (136)` | `IMP-182\IMP-182 inprogress.xlsm` | In-progress draft; empty CR Col A; `-1119` diff |
| 4 | `IMP-195` | `4100002284` | `4300006505-...` (15) | `2026-10-29` | `0 / 0` | `154 (154)` | `IMP-195 SIN CONTROL\IMP-195.xlsm` | `SIN CONTROL`; `0/0` bultos despite 154 lines |
| 5 | `IMP-196` | `4100002205` | `4300006878-...` (4) | `2026-10-23` | `519 / 519` | `29 (29)` | `IMP-196\IMP-196.xlsm` | Clean |
| 6 | `IMP-218` | `4100002274` | `4300007216-...` (6) | `2026-11-16` | `781 / 781` | `0 (8)` | `IMP-218\IMP-218.xlsm` | `DG!C18=0` vs 8 CR items; hybrid QC+Closure |
| 7 | `IMP-224` | `4100002293` | **`4300007562`** (1) | `2026-11-02` | `4168 / 4168` | `4 (4)` | `IMP-224\IMP-224.xlsm` | Shared PO `4300007562` across `IMP-224/225/226` |
| 8 | `IMP-225` | `4100002296` | **`4300007562`** (1) | `2026-11-03` | `2804 / 2804` | `3 (3)` | `IMP-225\IMP-225.xlsm` | Shared PO; `DG!C10` Proveedor = `#REF!` |
| 9 | `IMP-226` | `4100002297` | **`4300007562`** (1) | `2026-11-03` | `2804 / 2804` | `3 (3)` | `IMP-226\IMP-226.xlsm` | Shared PO; `DG!C10` Proveedor blank |
| 10 | `IMP-241` | `4100002294` | `4300007076-...` (7) | `2026-10-31` | `#REF! / #REF!` | `0 (62)` | `IMP-241 Sin Control Calidad\IMP-241.xlsm` | `Sin Control Calidad`; `5` `#REF!` errors |
| 11 | `IMP-242` | `4100002285` | `4300006902-...` (7) | `2026-10-24` | `0 / 0` | `67 (67)` | `IMP-242  PEND CONTROL\IMP-242.xlsm` | `PEND CONTROL`; `0/0` bultos despite 67 lines |
| 12 | `IMP-325` | `4100002301` | `4300007749` (1) | `2026-11-03` | `664 / 664` | `6 (6)` | `IMP-325 PEND CONTROL\IMP-325.xlsm` | Folder `PEND CONTROL` yet bultos = `664/664` |
| 13 | `IMP-326` | `4100002302` | `4300007309-8144` (1) | `2026-11-03` | `669 / 669` | `6 (6)` | `IMP-326\IMP-326.xlsm` | Author typo `Jostin Lopez Sobalbaro` (single `r`); 2 DG POs vs 1 CR |
| 14 | `IMP-327` | `4100002303` | `4300007308` (1) | `2026-11-03` | `669 / 669` | `6 (6)` | `IMP-327\IMP-327.xlsm` | Clean |
| 15 | `IMP-360` | `4100002286` | `4300007312` (1) | `2026-10-27` | `57 / 57` | `0 (50)` | `IMP-360\IMP-360.xlsm` | `DG!C18=0` vs 50 items; `4` `#VALUE!` errors |
| 16 | `IMP-363` | `4100002299` | `4300007679-8138` (2) | `2026-11-02` | `620 / 620` | `5 (5)` | `IMP-363\IMP-363.xlsm` | Clean |
| 17 | `IMP-364` | `4100002298` | `4300008022` (1) | `2026-11-02` | `2172 / 2172` | `5 (5)` | `IMP-364\IMP-364.xlsm` | `core.xml` `creator = Jostin Lopez` |
| 18 | `IMP-88` | `4100002275` | `4300006855-7729` (1) | `2026-10-29` | `839 / 839` | `8 (8)` | `IMP-88\IMP- 88.xlsm` | DG lists 2 POs (`-7729`) vs 1 CR |

**Business Identity vs Provenance Rules (`R4`):**
- **Primary Business Identity Key:** `(NormalizedContainerId, Pedido41)` (falling back to `NormalizedContainerId` with a mandatory human-review flag if `Pedido41` is missing). Never key on `PO_43` alone (`4300007562` spans `IMP-224/225/226`) and never key on file path, filename, or SHA-256.
- **Provenance Metadata (Strictly Separated):** Source filename, full/relative path, workbook identifier, file SHA-256, `vbaProject.bin` SHA-256, discovery timestamp. Moving or copying a file (`Copia de IMP-182 inprogress.xlsm`) must collapse to the same Business Identity and produce **0** duplicate Daily Log rows.
- **Human Review Queue Triggers:** Multiple file versions for one container (`IMP-182`), `#REF!`/`#VALUE!` summary cells (`IMP-241`, `IMP-360`, `IMP-225`), `0/0` cartons with >0 lines (`IMP-195`, `IMP-242`), or `DG!C18=0` mismatch (`IMP-218`, `IMP-360`).

---

## 7. Real Operational Processes & SAP Documentation Inventory

### 7.1 Complete Documentation Inventory (Worktree & Main Workspace)
Comparing the git worktree (`docs/`) with `D:\Development\Software Development\3 Way Inventory Reconciliation\` uncovered **4 authoritative operational documents absent from the worktree**:
1. `Control_Diario_Operaciones_Jostin.xlsx` (47,126 B, `.xlsx`) — Live daily container control log, 6 lifecycle stages, `Catalogos`, SAP prefix/8-step guide, and intake notification generator.
2. `docs\SOP - Procedimientos Operativos Conciliacion 3-Way.docx` (42,065 B, `.docx`) — Formal SOP `SOP-LOG-C001-01` (v1.2, Nov 2026) for Supervisión UPC Receiving (`C001`).
3. `docs\drafts\Borrador de procedimientos operativos.txt` (21,574 B, `.txt`, 16 sections / 778 lines) — Detailed SAP GUI navigation draft with explicit `"Pendiente de confirmar"` items.
4. `docs\drafts\BorradorRecepcion.txt` (3,470 B, `.txt`, 136 lines) — Raw operator shadowing field notes.

Existing worktree documentation (`docs/`): `SOP_RECEPCION_SAP.md`, `CONOCIMIENTO_EMPIRICO_BULTOS_MACROS.md`, `AUDITORIA_DESPACHO_MACRO.md`, `WIKI.md`, `RELEASE_NOTES_DEMO.md`, `README.md`/`README.es.md`, `sap-knowledge/KNOWLEDGE_EXPANSION_PLAN.md`, `architecture/CURRENT_ARCHITECTURE.md`, `architecture/TARGET_ARCHITECTURE.md`, `quality/CODE_AUDIT.md`, `quality/KNOWN_LIMITATIONS.md`, `ui/REDESIGN_SPECIFICATION.md`, and `ai/AI_ARCHITECTURE_PROPOSAL.md`. External corpus artifacts in `data/`: `BASE DE DATOS!A1:W1` (7-step analyst pipeline), `vbaProject.bin` (`Módulo1`: `Format43`, `Mueve43`, `Mueve41`, `MueveMatProv`, `MueveCurva`, `Format51`/`Mueve51`, `NoDuplicados`, `CalcDesg`), `106` `.msg`/`.eml` operational emails (`__PEDIDO NACIONALIZACION X__`, intake confirmations), and `160` PDFs (including 4 SAP Credit Notes `4000001153/54/57/1210`).

### 7.2 End-to-End Receiving Workflow & Reconciliation of `VL31N` / `RFUI` / `/SCWM/MON`
1. **Stage 1 — Customs, DUA & Radar:** Weekly Radar, sealed DUA, commercial invoice/Packing List (`INV.PL`), and `__PEDIDO NACIONALIZACION X__` email linking PO `430000xxxx`, contract, Nationalization `410000xxxx`, and `IMP-XXX`.
2. **Stage 2 — Pre-Reception SAP GUI 7.00 Extraction & EWM Intake:**
   - **Data Extraction:** `ME80FN` (mass PO `43...`/Pedido `41...` up to `UMP` `PRK`/`UN` and `EKPO-NETPR`), `SE16N` on `EINA` (`MATNR`/`LIFNR` -> `IDNLF`/`LMEIN`), `MAST` & `STPO` (`STLNR`, `IDNRK`, `MENGE` curve ratios), plus `MM60` (descriptions) and `VKP2` (retail price only — never cost).
   - **EWM Monitor (`/SCWM/MON`, Warehouse `C001`):** Mass product attribute update (`Z1`, putaway `C001`, section `ZALB`, removal `PICK`, dock/door `Z001`), bin creation `IMP-X` in Storage Type `Z970`.
   - **Inbound Delivery (`VL31N`, spelled `VL3IN` in drafts/.docx):** Pedido `41...` -> Header `Gestión 2 -> Ident. ext. = IMP-X` -> generates 9-digit Inbound Delivery (`1800...`) and 5-digit EWM doc (`414xx`).
   - **Packing & RFUI:** `Ctrl+F3` (`Embalar`) with packaging material `8000018` -> 18-digit HU (`80...` / `112...`), Warehouse Task to destination `CLARIFICACION`, background confirmation (`Confi.tarea.alma.fondo`), and `RFUI` Movement `522` (`CI proc. Y999`), followed by `MB51` (`C001`, Mov `101 - 102`) verification and intake notification email.
   - **Resolution of Prior Conflict:** Earlier worktree-only planning (`KNOWLEDGE_EXPANSION_PLAN.md §3.4`) flagged `VL31N`/`RFUI`/`18...`/`80...` as `0`-occurrence in `data/` closures. Inspecting `Control_Diario_Operaciones_Jostin.xlsx` (live `180060834`–`180060961`, `41461`–`41487`, `112345678005808817`–`823674`), `SOP-LOG-C001-01.docx`, and `drafts/*.txt` confirms these transactions **are actively used in Stage 2 intake**, even though those identifiers are not written into `CIERRE RECEP`.
3. **Stage 3 — Blind Setup (`Montaje a Ciegas`):** Preliminary closure mounted from supplier docs in `Pausa sin Control Calidad` when QC has not arrived (`$0.00` / `0` physical counts).
4. **Stage 4 — Physical Dock Unload & Quality Control (`ZDescarga`):** Produces `CONTROL DATA` (`Bultos Enteros`, `Complementos`, `A001 + A004 + Especiales`, `Observaciones`).
5. **Stage 5 — 3-Way Reconciliation, NCR & Official Closure (`Mov 101`):** Reconciles PO `43` vs Pedido `41` vs Physical QC vs `MB51 (101 - 102)` under the Cardinal Discrepancy Rule, populates `CIERRE RECEP` + `Bultos` (and `PARA N.C RECEP.` / `400000xxxx` NCR when applicable), and prints the 7-document Presidencia physical binder (`PORTADA`, `DATOS GENERALES`, `CIERRE RECEP`, `PARA N.C RECEP.`, `MB51 101/102`, `Bultos`, sealed `DUA`).
6. **Stage 6 — Dispatch Handoff (`CIERRE DESP` / `Mov 641`):** Separate dispatch role (`4500xxxxxx` store orders, `MB52`, `MB5T`, `Mov 643`); tracked in Daily Log Stage 6 only.

### 7.3 Standard SAP vs Company-Specific Z-Objects, Warnings & Verification Items

| Classification | Objects | Operational Warnings & Epistemic Rules |
|---|---|---|
| **Standard SAP Transactions & Tables (R5-B Core)** | `ME80FN`, `SE16N` (read-only `ACTVT=03`, never `&SAP_EDIT`), `MB51`, `VL31N`, `RFUI` (`/SCWM/RFUI`), `/SCWM/MON`, `MIGO`, `ME21N`/`ME22N`/`ME23N` (prefer read-only `ME23N`), `MM60`, `VKP2`, `MB52`, `MB5T`; Tables `EINA`, `MAST`, `STPO`. | **`VKP2` Retail Price Warning:** `VKP2` returns retail selling price in CRC, **never** import cost (`EKPO-NETPR`). **`STPO`/`MAST` Terminology:** `src/validator.py:33` mislabels `STPO`/`MAST` as a "transacción"—KB must classify `EINA`, `MAST`, `STPO` strictly as SAP **tables** queried via `SE16N`. |
| **Company-Specific Z-Transactions, Parameters & Custom Workflows** | `ZDescarga` (`Z-Carga`/`ZDECARGA`), `Z Facturas`, `Z51 -Z52`, `"7.UBICACIÓN MONITOR"`, EWM parameters (`C001`, `Z1`, `ZALB`, `PICK`, `Z001`, `Z970`, `8000018`, `CLARIFICACION`, `Y999`), document prefixes (`43`, `41`, `45`, `40`, `18`, `414xx`, `80/112`). | Never assume standard SAP behavior applies unchanged; tag all company-specific procedures explicitly vs standard SAP reference material. |
| **Payroll PII Exclusion Boundary (`Hoja1 MV 998`)** | `Hoja1` (`Autorización de Deducción`, `MV 998`) present in `81/147` `.xlsm` files in `data/`. | **Mandatory Exclusion:** `LocalDocImporter` and search indexers MUST strictly exclude `Hoja1` / `MV 998` employee `cédula` and payroll deduction PII. |
| **SAP GUI 7.00 Live Verification Items** | (1) Backend release (`System -> Status`: `ECC 6.0 + decentralized EWM` vs `S/4HANA` embedded EWM—remember GUI `7.00` does not identify backend release); (2) Exact T-Codes in `SE93`/`TSTC` for `VL31N` vs `VL3IN`, `ZDescarga`, `Z Facturas`, `Z51 -Z52`, `"7.UBICACIÓN MONITOR"`; (3) `T161-BSART` for `41...`, `Mov 643`, `MV 998` (`T156`). | Tag unverified items as `Reviewed` / unverified in environment (`R5-H`) until confirmed in live SAP GUI 7.00 (`Environment validated`). |

### 7.4 Traceability to R5 Sub-Requirements (`R5-A` through `R5-H`)
- **R5-A (Local Library):** Import PDF, DOCX, XLSX, TXT, MD, HTML offline without mutating sources; 7 categories; 15 required metadata fields per entry (including read-only vs modifying, permissions, risks/reversals, verification steps, source doc/page/section, validation status).
- **R5-B (Transaction Encyclopedia):** Includes all required initial R5-B entries (11 entries / 13 individual codes: `ME80FN`, `SE16N`, `MB51`, `VL31N`, `RFUI`, `MIGO`, `ME21N`/`22N`/`23N`, `ZDescarga`, `EINA`, `MAST`, `STPO`) plus corpus entries (`MM60`, `VKP2`, `MB52`, `MB5T`, `/SCWM/MON`).
- **R5-C (Guided Troubleshooting):** 8 operational incident workflows presenting all **10** required investigation sections, enforcing **5 epistemic tiers** (verified company procedures, official SAP docs, env-validated procedures, general SAP ref, unverified hypotheses).
- **R5-D & R5-E (Navigation & Integration):** SAP GUI 7.00 discovery guides and 1-click deep links from Reconciliation discrepancies and Daily Log rows.
- **R5-F (Technical Constraints):** Zero Qdrant, zero vector DB, zero LLM, zero external AI API, zero automatic SAP execution; deterministic local indexing and ranking only.
- **R5-G (Context-Aware Diagnostics & History):** Injects available PO, material, container, supplier, delivery, and HU context without inventing missing IDs; separates Observed Discrepancy vs Suspected Cause vs Confirmed Root Cause; allows user-confirmed promotion to personal KB entries.
- **R5-H (Provenance & Lifecycle States):** Tracks `Imported` -> `Extracted` -> `Reviewed` -> `Environment validated` -> `Deprecated` with exact page, section, cell, or line location.

---

## 8. Consolidated Historical Exceptions, Parity Gaps & Phase 1–6 Risk Register

### 8.1 Specific Historical Calculation & Dataset Exceptions
1. **`IMP-298` (`imp298 control calidad ver2.xlsm`) — `1036` vs `1032` Cartons:** `qc_reader.py` calculates `1036` cartons while `CONTROL DATA` reports `1032`. Verified on rows `108`, `110`, `111`, `112` (materials `1465571`, `1465573`, `1465574`, `1465575`), where operators omitted `complement_cartons = 1` from `reported_total_cartons` (also triggering `QC_FORMULA_MISMATCH`). `Acr.Core` must reproduce `1036` identically for Python parity.
2. **`IMP-192` (`'56*12+672'`) vs `IMP-200` (`'8*72+24'`) `+` Ambiguity:** In `IMP-200`, `+24` denotes loose pieces (`8*72 + 24 = 600`), whereas in `IMP-192` `'56*12+672'` repeats the product total (`56*12 = 672`; adding `+672` would double-count to `1344`). `QuantityEngine` preserves Python `_parse_loose_addition` parity for Phase 3 while flagging `loose == sum(terms)` candidates for cross-check against `AT`/`H`.
3. **`IMP-207 (NCR)` — `$869.76` NCR Composition Puzzle:** Legacy `src/parsers.py:603` computes `$550.08` (Row 38 `N=288 * $1.91` only, ignoring 4 unannotated `ped41 > rec` rows). Shortfall-only `max(ped41 - rec, 0) * unit_cost` across all 5 shortfall rows (Col `J` `$340.74` + Col `N` `$550.08`) equals `$890.82`; net across all 35 rows equals `$750.90`; human-entered `PARA N.C RECEP.!G6` is literal `$869.76` (`$21.06` below `$890.82`). Additionally, legacy `extract_cierre_metadata` overwrote `IMP-207` with `'288'` via Row 22 `'Faltante Importación'` (`CODE_AUDIT.md:A6-4`).
4. **`IMP-81` Naming & `Bulto Cero` / `Parte B` Rules:** QC file `IMP-81\INFORMACION GENERAL\IMP 81.xlsm` omits `"calidad"`/`"qc"`, causing false Blind Setup under filename matching (`KNOWN_LIMITATIONS.md`). `IMP-81.xlsm` establishes the **Bulto Cero** rule (`received_real_units == 0` or `container_part != "A"` -> `AS=AT=AW=AX=0`, `AY="0"`/`0`, noting legacy `app.py` vs `tab_bultos` diverged on `ped41 > 0` `FALTANTE` lines and `int 0` vs `str "0"`). Conversely, `IMP-207` demonstrates false `PENDIENTE_PARTE_B` fabrication when `ESQUEMA ALISTO PARTE B` layout tabs are matched by regex `PARTE\s+([A-Z0-9]+)`.
5. **Rows `>299` / `>500` & `'Bultos'` Trailing Row Cleanup:** `closure_exporter.py` caps at `range(13, 300)` and `extract_cierre_metadata` caps at `500`, whereas `10` populated closures exceed `max_row = 500` (`505..612`, `IMP-187` reaching `612` with `231` items) — plus `2` unpopulated `1,380`-row blank templates (`12` in Survey 2 JSON), and `13` total on disk including `IMP-81\IMP-81.xlsm` (`1,000` rows, where `read_only=True` yields `max_row = None`) — and `'Bultos'` exports (`app.py:1663`) leave stale template rows below new items (`IMP-187` `269` vs `234`; `IMP-200` `165` vs `130`). `Acr.Infrastructure` must scan dynamically to actual sheet end anchored on Col `E` (+ `order_po_43` / `source_row` for multi-PO duplicate SKUs) without trusting XML `<dimension>`/`max_row`, and clear trailing `'Bultos'` rows cleanly.
6. **Revuelto Confidence Levels (`R2`):** Python `discrepancy_classifier.py` only performs greedy exact-delta matching (`o.difference_units == abs(s.difference_units)`) without confidence tiers. `Acr.Core` must emit explicit `RevueltoConfidence` (`Confirmed`, `Probable`, `Possible`) with evidence details and never silently offset quantities.

### 8.2 Phase 1–6 Risk Register

| ID | Risk Description | Affected Phase(s) | Severity | Mitigation |
|---|---|---|---|---|
| **RK-01** | Machine-wide `.NET 10 SDK` install blocked from non-elevated shell (`0x80070032`), and Windows System `PATH` (`C:\Program Files\dotnet\`) precedes User `PATH`. | Phase 0 -> Phase 6 | Low (Mitigated) | Mitigated via authorized user-local `.NET 10 SDK` (`10.0.404`) + `Microsoft.WindowsDesktop.App 10.0.15` install at `%LocalAppData%\Microsoft\dotnet` with User `DOTNET_ROOT`/`PATH`; Phase 1–6 scripts/shells prepend `$env:LOCALAPPDATA\Microsoft\dotnet` or call `%LocalAppData%\Microsoft\dotnet\dotnet.exe`. |
| **RK-02** | Modifying `vbaProject.bin`, `styles.xml`, `definedNames`, or introducing `#REF!`/`#VALUE!` in `CIERRE RECEP` during `.xlsm` export. | Phase 2, 3, 6 | Critical | Preserve `vbaProject.bin` verbatim in OpenXML packaging; assert `SHA-256` equality and 3-way formula audit (`0` exporter-introduced errors) across all 81 closures in `Testing` mode. |
| **RK-03** | `81 MB` trailing XML bloat in `60/63` QC workbooks causing OOM or multi-second stalls. | Phase 2, 3, 4 | High | Forward-only streaming `OpenXmlReader` with `15`-empty-row early exit (target `<= 0.15s`/workbook). |
| **RK-04** | Corrupting or colliding on OneDrive/M365-locked `Control_Diario_Operaciones_Jostin.xlsx` (`~$`). | Phase 2, 4, 6 | Critical | Enforce `OperationalSafetyGuard`, `~$` lock detection, exclusive handle check, `.bak` + `.tmp` + validate + atomic replace, and `tblMaestroContenedores` resizing. |
| **RK-05** | Duplicate Daily Log entries from `IMP-202`-style duplicates, `IMP-182` copies, `IMP-326` `"Sobalbaro"` typo (`IMP-209` `sharedStrings.xml` trap), or shared PO `4300007562`. | Phase 1, 2, 4 | High | Key Business Identity on `(NormalizedContainerId, Pedido41)` separately from Provenance; read `'DATOS GENERALES'!B33` strictly (never `sharedStrings.xml`) and normalize author name; route ambiguous/error closures to human review queue. |
| **RK-06** | Indexing `Hoja1 MV 998` employee deduction PII or presenting unverified SAP hypotheses as official SOPs. | Phase 5, 6 | Critical | Hard-exclude `Hoja1`/`MV 998` in `LocalDocImporter`; enforce 5 epistemic tiers, `VKP2` non-cost warning, and 5 fragment lifecycle states. |
| **RK-07** | Multi-PO duplicate `material_code` overwrite in `DiscrepancyClassifier` or `ClosureExporter` row matching. | Phase 1, 2, 3 | Medium | Preserve row/PO identity `(material_code, po_43, source_row)` or explicit aggregation policy verified against Python regression baseline. |

---

## 9. Phase 0 Exit Criteria Checklist & Target `.NET 10 LTS` Architecture Summary

### 9.1 Phase 0 Exit Criteria Checklist

| Exit Criterion | Status | Evidence / Notes |
|---|---|---|
| Host `.NET 10 SDK` checked and installed (`10.0.x` alongside `9.0.312` & `8.0.428`; `Microsoft.WindowsDesktop.App 10.x`) | **PASS** | Machine-wide `winget` hit UAC `0x80070032`; after user authorization (`"ADELANTE"`), `.NET 10 SDK` `10.0.404` + `Microsoft.NETCore.App 10.0.15` + `Microsoft.AspNetCore.App 10.0.15` + `Microsoft.WindowsDesktop.App 10.0.15` were installed and verified in `C:\Users\Jostin\AppData\Local\Microsoft\dotnet` with User `DOTNET_ROOT` and User `PATH` configured (prepend `%LocalAppData%\Microsoft\dotnet` or invoke `dotnet.exe` directly in non-elevated shells because System `PATH` precedes User `PATH`). |
| Repository & Python reference baseline (`14240a8`, 4 services, `58/58` pytest) audited | **PASS** | `python -m pytest -v --tb=short` -> `58 passed in 43.14s`; `0` C# files exist. |
| Production corpus (`63+1` QC, `77+4` closures, `~$` lockfile, XML bloat, `vbaProject.bin`) audited | **PASS** | Verified across all `830` entries in `data/` and `80` `vbaProject.bin` SHA-256 hashes. |
| Actual `CIERRE RECEP` 58-column (`A..BF`) schema, `"45 columns"` reconciliation, and 5 error workbooks audited | **PASS** | `81/81` identical Row-12 headers; 3-way error classification verified on `IMP-241`, `IMP-244`, `IMP-256`, `IMP-360`, `IMP-90`. |
| Real 6 lifecycle stages, `Catalogos` dropdowns, `Col O` rollup, transitions, SAP template, and Daily Log exceptions audited | **PASS** | All `5` sheets and `22` rows of `Control_Diario_Operaciones_Jostin.xlsx` verified verbatim. |
| Historical Sync baseline (`20` files -> `18` containers, `DG!B33`, `IMP-326`, `IMP-224/225/226`, `IMP-182`) audited | **PASS** | Business Identity vs Provenance separation contract defined. |
| Real operational processes, SAP docs inventory, `VL31N`/`RFUI`/`/SCWM/MON` reconciliation, and R5-A..H mapped | **PASS** | Worktree + main workspace docs reconciled; live SAP GUI 7.00 verification items explicitly isolated. |
| Zero C# code written before Phase 0 Audit Report; production data unmodified | **PASS** | Only `docs/PHASE_0_AUDIT_REPORT.md` added; `data/` and Daily Log untouched. |

### 9.2 Target `.NET 10 LTS` Architecture Summary (`src/dotnet/`)
- **`Acr.Core`** (`net10.0`): Zero dependencies on WPF, OpenXML/ClosedXML, `System.IO`, or external NuGet packages. Contains domain models, enums, `QuantityEngine` (`checked` integer arithmetic), `DiscrepancyClassifier` (Cardinal Rule + `Confirmed`/`Probable`/`Possible` Revuelto detection), `DailyLogLifecycleValidator` (6 real stages + transitions), `HistoricalIdentityMatcher` (Business Identity vs Provenance), and SAP KB/Diagnostic domain types & deterministic ranker.
- **`Acr.Infrastructure`** (`net10.0`): OpenXML SDK streaming `QcReader`, macro/formula-preserving `ClosureExporter` + `FormulaIntegrityAuditor`, `OperationalSafetyGuard` (`Development`, `Testing`, `Production`), `AtomicFileTransaction` (`~$` lock check, `.bak`, `.tmp`, validate, atomic replace, verify), `DailyLogRepository` (table resizing), `HistoricalClosureScanner`, `ReceivingPipeline`, local JSON processing reports & operation history, and offline `LocalDocImporter` (PDF/DOCX/XLSX/TXT/MD/HTML excluding `Hoja1 MV 998`) + deterministic search index.
- **`Acr.Desktop`** (`net10.0-windows`, `UseWPF=true`, `CommunityToolkit.Mvvm`): Sober neutral slate/zinc corporate theme (zero emojis, zero neon gradients, zero AI buzzwords), user-switchable ES/EN localization, and 4 integrated views (Reconciliation, Daily Log, Historical Sync, SAP Knowledge & Context-Aware Operational Diagnostics), packaged in Phase 6 as a self-contained portable Windows `.exe`.
- **`Acr.Tests`** (`net10.0-windows`, xUnit): Unit, Infrastructure, Regression Parity (64 QC [63+1, covering 51 Python-matched] & 81 closures [77+4]), SAP KB/Diagnostic, ViewModel/Theme/Localization, and Opaque-Box E2E Tiers 1–4 + Tier 5 Adversarial suites executed strictly in `Testing` mode.
