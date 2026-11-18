# Opaque-Box E2E Test Specifications — Asistente de Cierre de Recepción (ACR)

**Document ID:** `ACR-E2E-SPEC-PHASE-A-2026-11-18`  
**Target System:** Asistente de Cierre de Recepción (ACR) / Receiving Closure Assistant (RCA) (`.NET 10 LTS` / C# WPF)  
**Authoritative Sources:** `ORIGINAL_REQUEST.md` (`R1`–`R6`), `PROJECT.md`, `docs/PHASE_0_AUDIT_REPORT.md`, Python Reference Services 1–4  
**Safety Mode:** Strictly `OperationalSafetyMode.Testing` (Production `data/` and `Control_Diario_Operaciones_Jostin.xlsx` are NEVER mutated)  
**Total Specified Cases:** **254** (`105` Tier 1 + `105` Tier 2 + `28` Tier 3 + `16` Tier 4 across all `21` operational features `F02`–`F22`)

> **Phase 0 Gate Note (`F01`):** Feature `#1` (`Phase 0 Setup & Formal Audit Report`) was completed and verified at M0 Gate (`docs/PHASE_0_AUDIT_REPORT.md`, `.NET 10 SDK 10.0.404` + `Microsoft.WindowsDesktop.App 10.0.15`, `58/58` Python pytest baseline). Therefore, `F02`–`F22` (`N = 21`) form the complete operational feature inventory under E2E test.

---

## Coverage Summary

| Tier | Description | Formula Threshold (`N = 21`) | Minimum | Actual Specified | ID Pattern |
|---|---|---|---:|---:|---|
| **Tier 1** | Feature Coverage (Equivalence Classes) | $\ge 5 \times N$ (`>= 5` per feature) | 105 | **105** | `E2E-F{02..22}-T1-0{1..5}` |
| **Tier 2** | Boundary & Corner Cases | $\ge 5 \times N$ (`>= 5` per feature) | 105 | **105** | `E2E-F{02..22}-T2-0{1..5}` |
| **Tier 3** | Cross-Feature Pairwise Interactions | $\ge 1 \times N$ | 21 | **28** | `E2E-T3-01`..`E2E-T3-28` |
| **Tier 4** | Real-World Application Scenarios | $\ge \max(5, \lceil N / 2 \rceil)$ | 11 | **16** | `E2E-T4-01`..`E2E-T4-16` |
| **Total** | **Full Opaque-Box E2E Suite** | **Sum of Tiers 1–4** | **242** | **254** | **0 ID collisions** |

---

# Part 1 — Tiers 1 & 2 Specifications (`F02`–`F15`: Core, Infrastructure & Parity Gate)

## Feature F02: Solution Structure, `Acr.Core` Isolation & Portable Constraints (`R1, R6`)

### Tier 1 — Feature Coverage (F02)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F02-T1-01** | 4-project `.NET 10 LTS` solution layout & clean compilation | `src/dotnet/Acr.sln` containing `Acr.Core`, `Acr.Infrastructure`, `Acr.Desktop`, `Acr.Tests` built via `.NET 10 SDK` (`10.0.404`) | `dotnet build src/dotnet/Acr.sln -c Release` exits `0` with `0` errors; all 4 projects target `net10.0` or `net10.0-windows`. |
| **E2E-F02-T1-02** | Strict dependency direction & `Acr.Core` isolation | Project & assembly reference metadata of `Acr.Core.csproj` and compiled `Acr.Core.dll` | `Acr.Core.csproj` has `0` `<PackageReference>` and `0` `<ProjectReference>` elements; `Acr.Core.dll` references only core BCL (`System.Runtime`, `System.Collections`, etc.) and **0** references to `PresentationFramework`, `WindowsBase`, `DocumentFormat.OpenXml`, `ClosedXML`, `System.IO.FileSystem`, or `CommunityToolkit.Mvvm`. |
| **E2E-F02-T1-03** | Forbidden infrastructure bloat absence (Lean Workstation Scope) | Full `src/dotnet/` `.csproj` packages and source imports | Zero references anywhere in `src/dotnet/` to `Qdrant`, vector DB clients, `OpenAI`/`Anthropic`/LLM SDKs, `Serilog`/`NLog`/`EnterpriseLibrary`, `EntityFramework`/SQL audit DBs, `ASP.NET` microservices, or update-server frameworks (`ORIGINAL_REQUEST.md:R1`). |
| **E2E-F02-T1-04** | Self-contained portable Windows executable capability | `dotnet publish src/dotnet/Acr.Desktop/Acr.Desktop.csproj -c Release -r win-x64 --self-contained true` into an isolated temp directory | Publish exits `0`; output contains runnable `.exe` without MSI/Setup installer; launches cleanly without requiring machine-wide `.NET 10` runtime installation. |
| **E2E-F02-T1-05** | Code standards compliance on public domain/infrastructure APIs | Public assemblies & source files in `Acr.Core` / `Acr.Infrastructure` | Public types/methods use `PascalCase`; XML documentation comments (`/// <summary>`) exist on public types/methods (producing XML doc file without CS1591 errors if enabled); non-generated methods respect `< 50` lines guideline (`R6`). |

### Tier 2 — Boundary & Corner Cases (F02)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F02-T2-01** | `System.IO` filesystem API leak check inside `Acr.Core` | Inspect `Acr.Core` compiled IL/assembly references and `*.cs` syntax trees | Zero usages of `System.IO.File`, `Directory`, `FileStream`, `StreamReader`, `StreamWriter`, `FileInfo`, or `DirectoryInfo` in `Acr.Core` (test fails immediately if any filesystem call is added to `Acr.Core`). |
| **E2E-F02-T2-02** | Windows `PATH` precedence trap (`C:\Program Files\dotnet` vs `%LocalAppData%\Microsoft\dotnet`) | Execute build/test command in a sub-shell where System `PATH` has `.NET 8/9` first, vs User-prepended `.NET 10` path | Test/build harness detects SDK version explicitly (`10.0.404`) and never silently builds/runs under `.NET 8/9` (`PHASE_0_AUDIT_REPORT.md §1.3`). |
| **E2E-F02-T2-03** | Air-gapped / offline execution boundary | Run `Acr.Core` and `Acr.Infrastructure` test suites with network disabled / proxy blocked | `100%` of offline operations pass with `0` HTTP/HTTPS requests (no external icon/CDN fetches like legacy `app.py:313` `icons8.com`). |
| **E2E-F02-T2-04** | Portable `.exe` launched from non-ProgramFiles path with spaces | Copy published self-contained folder to `%TEMP%\ACR Portable Test (x64)\` and launch with `Testing` mode args | Application starts cleanly without needing UAC elevation or write access to `C:\Program Files\`. |
| **E2E-F02-T2-05** | Existing Python baseline non-interference | Presence of `src/dotnet/` alongside `src/` and `tests/` | Running `python -B -m pytest -p no:cacheprovider -v --tb=short` after `.NET` solution creation still passes `58/58` tests (`0` failures, `0` skipped). |

---

## Feature F03: `QuantityEngine` Integer Grammar & Normalization (`R2`)

### Tier 1 — Feature Coverage (F03)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F03-T1-01** | Single multiplication term (`*` and `x`/`X` delimiters) | `"50*12"`, `"10x24"`, `"8X36"` | `"50*12"` -> `TotalCartons=50, TotalUnits=600, FullCartons=50, UnitsPerFull=12, IsCompound=false, HasLoose=false`; `"10x24"` -> `10, 240`; `"8X36"` -> `8, 288`. |
| **E2E-F03-T1-02** | Scalar integer string | `"192"` | `TotalCartons=1, TotalUnits=192, FullCartons=1, UnitsPerFull=192, IsCompound=false, HasLoose=false` (matches `quantity_engine.py:137-146`, NOT legacy `42` boxes). |
| **E2E-F03-T1-03** | Compound multi-term expression (space-separated & `+`-separated) | `"3*192 3*204"`, `"40*60 + 1*72"` | `"3*192 3*204"` -> `TotalCartons=6, TotalUnits=1188, Full=(3,192), Comp=(3,204), IsCompound=true, HasLoose=false`; `"40*60 + 1*72"` -> `TotalCartons=41, TotalUnits=2472, LooseUnits=0, IsCompound=true, HasLoose=false`. |
| **E2E-F03-T1-04** | Compound + trailing loose addition (`IMP-103` Row 37 / `IMP-200`) | `"6*192 +48"`, `"8*72+24"` | `"6*192 +48"` -> `TotalCartons=6, TotalUnits=1200, Full=(6,192), LooseUnits=48, IsCompound=true, HasLoose=true`; `"8*72+24"` -> `TotalCartons=8, TotalUnits=600, LooseUnits=24, HasLoose=true`. |
| **E2E-F03-T1-05** | Null, empty, and sentinel string normalization | `null`, `""`, `"   "`, `"None"`, `"NAN"`, `"null"` | Returns zeroed `ParsedPackaging` (`TotalCartons=0, TotalUnits=0, FullCartons=0, UnitsPerFull=0, IsCompound=false, HasLoose=false`) with `RawString=""` and zero exceptions. |

### Tier 2 — Boundary & Corner Cases (F03)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F03-T2-01** | 3+ term compound expression (no 2-term truncation) | `"10*12 5*6 1*3"` (`CODE_AUDIT.md:A2-2`) | `TotalCartons=16, TotalUnits=153` (`120 + 30 + 3`, NOT `150` truncated), `IsCompound=true`. |
| **E2E-F03-T2-02** | Whitespace & separator noise normalization | `"  6  *  192   +   48  "`, `"3 * 192\t3 x 204"` | Trims outer/inner spaces/tabs cleanly: `"  6  *  192   +   48  "` -> `TotalCartons=6, TotalUnits=1200, LooseUnits=48`. |
| **E2E-F03-T2-03** | Integer overflow protection (`checked` arithmetic) | `"2147483647*2"`, `"100000*100000"` (32-bit checked) or `"9223372036854775807*2"` (64-bit) | Never wraps silently to a negative or truncated number; raises explicit `OverflowException` or deterministic validation error (`ORIGINAL_REQUEST.md:R2`). |
| **E2E-F03-T2-04** | Zero, negative, and malformed non-numeric tokens | `"0"`, `"0*12"`, `"-5*12"`, `"ABC"`, `"12*34*56"` | `"0"` -> `TotalCartons=1, TotalUnits=0` (`0*12` -> `0, 0`); malformed `"ABC"` -> zeroed `ParsedPackaging` without crash; negative `"-5*12"` never produces negative physical units (`>= 0`). |
| **E2E-F03-T2-05** | `IMP-192` `+` ambiguity (`loose == sum(terms)`) | `"56*12+672"` (`PHASE_0_AUDIT_REPORT.md §8.1(2)`) | Preserves Python `_parse_loose_addition` parity (`TotalCartons=56, TotalUnits=1344, LooseUnits=672`) AND surfaces explicit ambiguity candidate warning for cross-check against `AT`/`H` without silent data loss. |

---

## Feature F04: `QuantityEngine` Complement Rescue, Special Rows & Curve Breakdown (`R2`)

### Tier 1 — Feature Coverage (F04)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F04-T1-01** | Standard full + explicit complement cartons | `FullCartons=3, UnitsPerFull=192, CompCartons=3, UnitsInComp=204, RepCartons=6, RepUnits=1188, Curve=12, Uom="PRK"` | `CalculatedUnits=1188, CalculatedCartons=6, CalculatedPacks=99, LooseUnits=0, FormulaDiscrepancy=0, Flags=[]`. |
| **E2E-F04-T1-02** | Complement Rescue (Excel 0-multiplier bug, `IMP-103` Row 37) | `FullCartons=6, UnitsPerFull=192, CompCartons=null, UnitsInComp=48, RepCartons=6, RepUnits=1200, RepDiff=-48, Curve=12, Uom="PRK"` | `CalculatedUnits=1200, CalculatedCartons=6, CalculatedPacks=100, LooseUnits=0, FormulaDiscrepancy=0`; flags include `QC_COMPLEMENT_WITHOUT_CARTON_COUNT` and `NEGATIVE_REPORTED_DIFF`; explanation records `48` rescued units. |
| **E2E-F04-T1-03** | `VARIANTE SUELTA` row | `FullCartons=null, RepCartons=0, RepUnits=10, Observation="VARIANTE SUELTA", Uom="UN"` | `CalculatedUnits=10, CalculatedCartons=0, CalculatedPacks=10, LooseUnits=0, FormulaDiscrepancy=0, Flags=[VARIANTE_SUELTA_ROW]`. |
| **E2E-F04-T1-04** | `DOS MATERIALES POR BULTO` shared carton pair | Primary: `Full=1, UnitsPerFull=300, Obs="DOS MATERIALES POR BULTO"`; Secondary: `Full=null, UnitsPerFull=300, Obs="DOS MATERIALES POR BULTO"` | Primary -> `CalculatedCartons=1, CalculatedUnits=300, Flags=[SHARED_CARTON_PRIMARY]`; Secondary -> `CalculatedCartons=0, CalculatedUnits=300, Flags=[SHARED_CARTON_SECONDARY]`; combined = `1` carton, `600` units. |
| **E2E-F04-T1-05** | Prepack curve division & remainder (`UNEVEN_PREPACK_CURVE`) | `Full=15, UnitsPerFull=48, CompC=1, CompU=22, Curve=12, Uom="PRK"` | `CalculatedUnits=742, CalculatedCartons=16, CalculatedPacks=61, LooseUnits=10, Flags=[UNEVEN_PREPACK_CURVE]` (integer division `742 / 12 = 61`, `742 % 12 = 10`). |

### Tier 2 — Boundary & Corner Cases (F04)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F04-T2-01** | Complement Rescue carton-count inference boundary (`a` vs `a+1` vs `a=0`) | (a) `Full=6, CompC=0, CompU=48, RepCartons=6`; (b) `Full=6, CompC=null, CompU=48, RepCartons=7`; (c) `Full=0, CompC=null, CompU=24, RepCartons=null` | (a) `CalculatedCartons=6` (`comp_c=0`); (b) `CalculatedCartons=7` (`comp_c=1`); (c) `CalculatedCartons=1` (`comp_c=1` when `a==0`); all emit `QC_COMPLEMENT_WITHOUT_CARTON_COUNT` and rescue units. |
| **E2E-F04-T2-02** | Empty complement cartons (`comp_c > 0, comp_d == 0`) | `Full=10, UnitsPerFull=50, CompC=2, CompU=0` | `CalculatedCartons=12, CalculatedUnits=500, Flags=[EMPTY_COMPLEMENT_CARTONS]`. |
| **E2E-F04-T2-03** | Precedence & fallback boundaries (`VARIANTE SUELTA` vs `packaging_string` fallback) | (a) `IsVarianteSuelta=true, UnitsInComp=14, RepUnits=null, Obs=""`; (b) All numeric fields `null/0`, `PackagingString="6*192 +48", Curve=12, Uom="PRK"` | (a) `CalculatedUnits=14, CalculatedCartons=0, Flags=[VARIANTE_SUELTA_ROW]`; (b) `CalculatedUnits=1200, CalculatedCartons=6, CalculatedPacks=100, LooseUnits=0`. |
| **E2E-F04-T2-04** | Observation adjustment keywords (case-insensitive & spacing variants) | `"MAS 1PK, 5685920"`, `"mas 1 pk"`, `"+1PK"`, `"MENOS 1PK"`, `"menos 1 pk"`, `"-1PK"`, `"SIN NOVEDAD"` | First 6 variants emit `QC_OBSERVATION_ADJUSTMENT`; `"SIN NOVEDAD"` does NOT emit it. |
| **E2E-F04-T2-05** | `UOM="UN"` vs `Curve=null` & `checked` multiplication overflow | (a) `Full=5, UnitsPerFull=100, Curve=null, Uom="UN"`; (b) `Full=5, UnitsPerFull=100, Curve=null, Uom="PRK"`; (c) `Full=100000, UnitsPerFull=100000` (32-bit) | (a) `CalculatedPacks=500, LooseUnits=0`; (b) `CalculatedPacks=null, LooseUnits=0`; (c) throws `OverflowException` / explicit overflow failure instead of wrapping. |

---

## Feature F05: `DiscrepancyClassifier` Cardinal Rule & 8 Statuses (`R2`)

### Tier 1 — Feature Coverage (F05)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F05-T1-01** | Cardinal Discrepancy Rule (never offset `-X` against `+Y`) | `Declared=[("1001", 100u, 2c), ("1002", 100u, 2c)]`; `Physical=[("1001", 80u, 2c), ("1002", 120u, 2c)]` | `NetUnitDifference == 0`, YET `TotalShortageUnits == 20`, `TotalOverageUnits == 20`, `DiscrepantMaterialsCount == 2`, `CleanMaterialsCount == 0`. |
| **E2E-F05-T1-02** | `EXACT_MATCH` & `SHORTAGE` & `OVERAGE` | `Decl=[("A",100,2),("B",100,2),("C",100,2)]`, `Phys=[("A",100,2),("B",90,2),("C",115,2)]` | `A -> EXACT_MATCH (diff=0)`, `B -> SHORTAGE (diff=-10)`, `C -> OVERAGE (diff=+15)`; `TotalShortageUnits=10, TotalOverageUnits=15, NetUnitDifference=+5`. |
| **E2E-F05-T1-03** | `MISSING_ITEM` & `UNEXPECTED_ITEM` | `Decl=[("1001",100,2),("1002",200,4)]`, `Phys=[("1001",100,2),("9999",50,1)]` | `"1002" -> MISSING_ITEM (diff=-200)`, `"9999" -> UNEXPECTED_ITEM (diff=+50)`, `"1001" -> EXACT_MATCH`; `TotalShortageUnits=200, TotalOverageUnits=50`. |
| **E2E-F05-T1-04** | `DAMAGE_ONLY` & `CARTON_MISMATCH_ONLY` | `Decl=[("D1",100,2),("C1",100,5)]`, `Phys=[("D1",100,2,damaged=5),("C1",100,4,damaged=0)]` | `"D1" -> DAMAGE_ONLY (diff_u=0, damaged=5)`, `"C1" -> CARTON_MISMATCH_ONLY (diff_u=0, diff_c=-1)`; `TotalDamagedUnits=5, DiscrepantMaterialsCount=2`. |
| **E2E-F05-T1-05** | `REVUELTO_CANDIDATE` status classification | `Decl=[("M1",144,3),("M2",144,3)]`, `Phys=[("M1",120,3),("M2",168,3)]` | Both `"M1"` (`-24`) and `"M2"` (`+24`) are classified `REVUELTO_CANDIDATE` with `RevueltoPairsCount=1`, `TotalShortageUnits=24`, `TotalOverageUnits=24`. |

### Tier 2 — Boundary & Corner Cases (F05)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F05-T2-01** | Classification precedence when multiple conditions hold simultaneously | (a) `Decl=100, Phys=80, Damaged=10, DiffCartons=-1`; (b) `Decl=100, Phys=100, Damaged=5, DiffCartons=-1`; (c) `Decl=0, Phys=0` | (a) `SHORTAGE` takes precedence over `DAMAGE_ONLY` & `CARTON_MISMATCH_ONLY`; (b) `DAMAGE_ONLY` takes precedence over `CARTON_MISMATCH_ONLY`; (c) `EXACT_MATCH` (`0` vs `0`). |
| **E2E-F05-T2-02** | Empty declared & empty physical lists | `Declared=[]`, `Physical=[]` | Returns valid empty `ComparisonSummary` with all totals `0` and `Items.Count == 0` (no `ArgumentException` or null ref). |
| **E2E-F05-T2-03** | Duplicate `material_code` across multiple PO lines (`RK-07`) | `Declared=[("1455103", PO="4300001", 100u, 2c), ("1455103", PO="4300002", 50u, 1c)]`; `Physical=[("1455103", 150u, 3c)]` | Never silently overwrites/drops the first PO line; total declared is `150` (not `50` or `100`), physical `150`, `NetUnitDifference == 0`. |
| **E2E-F05-T2-04** | Asymmetric Cardinal Rule with non-matching multi-material variances | `Decl=[("A",100,2),("B",100,2),("C",100,2)]`, `Phys=[("A",70,2),("B",120,2),("C",110,2)]` | `NetUnitDifference == 0` (`-30 + 20 + 10`), YET `TotalShortageUnits == 30`, `TotalOverageUnits == 30`, `RevueltoPairsCount == 0`. |
| **E2E-F05-T2-05** | Whitespace/leading-zero material code & alphanumeric `"1E2"` SKU preservation | `" 001455103 "` vs `"1455103"` and alphanumeric `"1E2"` | Alphanumeric `"1E2"` is never corrupted to `"100"`; comparison is deterministic and preserves exact SKU identity without scientific-notation mangling. |

---

## Feature F06: `DiscrepancyClassifier` Revuelto Detection with Confidence (`R2`)

### Tier 1 — Feature Coverage (F06)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F06-T1-01** | `Confirmed` confidence revuelto pair (exact delta + strong shared attributes) | Shortage: `Mat="1455102001", VendorMat="STYLE-A", Category="CABALLERO", Uom="PRK", Curve=12, Diff=-24`; Overage: `Mat="1455102002", VendorMat="STYLE-A", Category="CABALLERO", Uom="PRK", Curve=12, Diff=+24` | Paired as `REVUELTO_CANDIDATE`, `RevueltoConfidence == Confirmed`, `RevueltoPartner` linked bidirectionally, non-empty evidence rationale citing matching delta & shared vendor/category/curve; `TotalShortageUnits=24, TotalOverageUnits=24`. |
| **E2E-F06-T1-02** | `Probable` confidence revuelto pair (exact delta + partial attribute overlap) | Shortage: `Mat="5033284", VendorMat="V-10", Category="DAMA", Diff=-36`; Overage: `Mat="5033999", VendorMat="V-99", Category="DAMA", Diff=+36` | Paired as `REVUELTO_CANDIDATE`, `RevueltoConfidence == Probable` (exact magnitude match + shared category/UOM, different vendor style), with explicit evidence rationale. |
| **E2E-F06-T1-03** | `Possible` confidence revuelto pair (exact delta with differing categories) | Shortage: `Mat="1001", Category="CABALLERO", Diff=-48`; Overage: `Mat="2002", Category="INFANTIL", Diff=+48` | Marked `REVUELTO_CANDIDATE` with `RevueltoConfidence == Possible` and warning rationale noting cross-category mismatch; never auto-resolved. |
| **E2E-F06-T1-04** | Non-matching deltas remain independent `SHORTAGE` & `OVERAGE` | `Mat="1001" Diff=-25`, `Mat="1002" Diff=+24` | `RevueltoPairsCount == 0`; `"1001"` remains `SHORTAGE`, `"1002"` remains `OVERAGE`, `RevueltoPartner == null`. |
| **E2E-F06-T1-05** | Non-offsetting invariant on revuelto pairs | Any paired `Confirmed`/`Probable`/`Possible` revuelto (`-X` and `+X`) | `TotalShortageUnits` still includes `X`, `TotalOverageUnits` still includes `X`, `DiscrepantMaterialsCount` still counts both items (never silently converted to `EXACT_MATCH`). |

### Tier 2 — Boundary & Corner Cases (F06)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F06-T2-01** | Tie-breaking when 2 overages match 1 shortage | Shortage: `S1 (VendorMat="ST1", -24)`; Overages: `O1 (VendorMat="OTHER", +24)` listed first, `O2 (VendorMat="ST1", +24)` listed second | Deterministic pairing prefers the stronger evidence match (`O2`, `Confirmed`) over blind insertion-order `O1`, leaving `O1` unpaired as `OVERAGE`. |
| **E2E-F06-T2-02** | 1-to-1 exclusivity when 1 shortage meets 2 identical overages | `S1 (-24)`, `O1 (+24)`, `O2 (+24)` | Only 1 pair is formed (`RevueltoPairsCount == 1`); one overage remains unpaired `OVERAGE` (no double-pairing of `S1`). |
| **E2E-F06-T2-03** | Zero-unit / `MISSING_ITEM` vs `UNEXPECTED_ITEM` exact swap | `Decl=[("A",100,2)]`, `Phys=[("B",100,2)]` (`A`: `decl=100, phys=0`; `B`: `decl=0, phys=100`) | Detects revuelto swap (`A` `-100` <-> `B` `+100`) with `RevueltoPairsCount == 1`, `TotalShortageUnits == 100`, `TotalOverageUnits == 100`, preserving `DifferenceUnits` (`-100` / `+100`). |
| **E2E-F06-T2-04** | Split 1-to-2 swap (`-48` vs `+24` and `+24`) preserves Cardinal totals | `S1 (-48)`, `O1 (+24)`, `O2 (+24)` | Never alters `TotalShortageUnits (48)` or `TotalOverageUnits (48)`; either emits `RevueltoPairsCount=0` (strict 1:1 Python parity) or flags `Possible` multi-item hint without mutating per-item deltas. |
| **E2E-F06-T2-05** | Multiple disjoint revuelto pairs in a single container | `(-12, +12)`, `(-36, +36)`, `(-72, +72)` | `RevueltoPairsCount == 3`; each pair links bidirectionally (`s.Partner == o.Code && o.Partner == s.Code`) with zero cross-contamination. |

---

## Feature F07: Daily Log 6-Stage Domain Model, `Col O` Rollup & Transition Rules (`R4`)

### Tier 1 — Feature Coverage (F07)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F07-T1-01** | Exact 6-stage catalog values acceptance (`Catalogos!A1:F7`) | All valid dropdown strings for Stages 1..6 from `PHASE_0_AUDIT_REPORT.md §5.2` | Domain validator accepts every exact catalog string across Stages 1–6 and rejects unknown/arbitrary strings. |
| **E2E-F07-T1-02** | `Col O` (`Estado Maestro Actual`) 7-state deterministic rollup | Representative 7 stage combinations matching `Col O` formula precedence (`PHASE_0_AUDIT_REPORT.md §5.3`) | Computed `EstadoMaestroActual` matches all 7 exact corporate rollup strings verbatim (`"1. En Cola Conteo a Ciegas"` .. `"7. Cierre Total (Recep + Desp)"`). |
| **E2E-F07-T1-03** | Blind Setup Pause guardrail (Rule 3) | Stage 3 = `Listo (Solo Doc. Proveedor)`, Stage 4 = `Sin CC (Espera Muelle)`, Stage 5 = `Pausa sin Control Calidad`, Stage 6 = `Bloqueado por Recepción` | Valid transition -> `Col O = "4. Pausa sin CC (Montado a Ciegas)"`. Attempting Stage 5 = `Cuadrado` or Stage 6 = `En Distribución Tiendas` under `Sin CC` fails transition validation. |
| **E2E-F07-T1-04** | Reception closure & Dispatch unlock progression (Rules 5 & 6) | Stage 1 = `DUA y Docs Listos`, Stage 4 = `CC Aprobado (ZDescarga)`, Stage 5 = `Cuadrado (0 Diferencias)` -> Stage 6 transitions `Bloqueado por Recepción` -> `Pre-Cableado 80% (Listo Mov 641)` -> `En Distribución Tiendas` -> `Cierre Despacho Listo` | All progressive transitions succeed; `Col O` updates from `6. Recepción Cerrada (Listo Recep.)` to `7. Cierre Total (Recep + Desp)`. |
| **E2E-F07-T1-05** | SAP intake notification formatting (`Col R` / `C8:C9`) | `ContainerId="IMP-181"`, `Entrega18="180060834"`, `Saludo="Buenos días"` | Subject: `"Ingreso Listo a Sistema - Contenedor 181"`; Body: `"Buenos días\nSe informa que el ingreso a nivel de sistema del Contenedor 181, quedo listo."` (strips `"IMP-"` prefix per `SUBSTITUTE(B{r},"IMP-","")`). |

### Tier 2 — Boundary & Corner Cases (F07)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F07-T2-01** | Stage 6 unlock blocked while Stage 5 is still open/discrepant | Stage 5 = `Pendiente Cierre` or `Discrepancia / NCR (40...)` or `Pausa sin Control Calidad`; attempt Stage 6 = `Pre-Cableado 80% (Listo Mov 641)` | Validation **blocks** Stage 6 exit from `Bloqueado por Recepción` (only `Cuadrado (0 Diferencias)` or `Parcial / Pendiente Parte B` unlocks Stage 6). |
| **E2E-F07-T2-02** | Retroactive `"-"` SAP IDs on Stage 2 `Listo y Notificado` (`IMP-195/217/259/349` rule) | Stage 2 = `Listo y Notificado` with `Entrega18="-"`, `EwmDoc="-"`, `Hu="-"` | Allowed with a **non-blocking warning** (retroactive persistence supported per `PHASE_0_AUDIT_REPORT.md §5.3(2)`), and `Col R` notification yields `"-"`. |
| **E2E-F07-T2-03** | Inconsistent Stage 5 `Cuadrado` override (`IMP-217` Row 8 case) | Stage 5 = `Cuadrado (0 Diferencias)` while Stage 4 = `Pendiente Revisión Final CC` or Stage 1 = `Pendiente Docs` | Transition validator explicitly flags inconsistency (requiring Stage 1 `DUA y Docs Listos` and Stage 4 `CC Aprobado (ZDescarga)`). |
| **E2E-F07-T2-04** | `Col O` multi-match precedence boundary (Stage 6 vs Stage 5 vs Stage 4 `"Revisión"`) | Stage 6 = `Cierre Despacho Listo`, Stage 5 = `Cuadrado (0 Diferencias)`, Stage 4 = `Pendiente Revisión Final CC` | Top-down `IF` precedence selects `"7. Cierre Total (Recep + Desp)"`; when Stage 6 != `Cierre Despacho Listo` and Stage 5 != `Cuadrado` and Stage 4 contains `"Revisión"`, selects `"5. En Auditoría Control Calidad"`. |
| **E2E-F07-T2-05** | `Saludo="-"` and multi-part container ID in SAP notification | `ContainerId="IMP-201-1"`, `Entrega18="-"` vs `Entrega18="180060900"`, `Saludo="-"` | When `Entrega18 == "-"`, `Col R` returns `"-"`; when `Entrega18 != "-"` and `ContainerId="IMP-201-1"`, `SUBSTITUTE` produces `"Contenedor 201-1"`. |

---

## Feature F08: Historical Sync Business Identity vs Provenance Domain (`R4`)

### Tier 1 — Feature Coverage (F08)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F08-T1-01** | Primary Business Identity `(NormalizedContainerId, Pedido41)` | `Container="IMP-104"`, `Pedido41="4100002248"`, `Po43="4300008021-..."`, `Provenance="IMP-104\\IMP-104.xlsm"` | Produces canonical identity key `("IMP-104", "4100002248")` completely independent of path, filename, or file SHA-256. |
| **E2E-F08-T1-02** | Provenance separation on file copy/move (`IMP-182`) | Record A: `Path="IMP-182\\IMP-182.xlsm", SHA="aaa"`; Record B: `Path="Backup\\Copia de IMP-182.xlsm", SHA="bbb"` (both `IMP-182`, `4100002283`) | Both resolve to the **same** Business Identity `("IMP-182", "4100002283")`, preserving both provenance records without creating a duplicate container entry. |
| **E2E-F08-T1-03** | Shared PO across distinct containers (`IMP-224/225/226`) | 3 records sharing `PO="4300007562"` with `(IMP-224, 4100002293)`, `(IMP-225, 4100002296)`, `(IMP-226, 4100002297)` | Resolves to **3 distinct** Business Identities; shared PO never merges or deduplicates separate containers. |
| **E2E-F08-T1-04** | Author normalization (`DATOS GENERALES!B33`) | `"ELABORADO POR: \nJostin Lopez Sobalbarro"`, `"Jostin Lopez Sobalbaro"` (`IMP-326`), `"Jostin López Sobalbarro"`, `"Angie Montero Rojas"` (`IMP-209`) | First 3 normalize to `IsTargetAuthor = true` (`Jostin Lopez Sobalbarro`); `"Angie Montero Rojas"` yields `IsTargetAuthor = false` (`IgnoredNonOwner`). |
| **E2E-F08-T1-05** | Idempotent re-sync against existing Daily Log | Existing Daily Log has `("IMP-181", "4100002271")`; incoming scan yields `("IMP-181", "4100002271")` and `("IMP-363", "4100002299")` | `IMP-181` classified `AlreadySynced` (0 inserts); `IMP-363` classified `NewImportCandidate` (1 insert). |

### Tier 2 — Boundary & Corner Cases (F08)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F08-T2-01** | Missing `Pedido41` fallback + mandatory human review flag | `Container="IMP-182"`, `Pedido41=""` or `null` | Falls back to `NormalizedContainerId` (`"IMP-182"`) and **mandatorily** sets `RequiresHumanReview = true` with reason `MissingPedido41`. |
| **E2E-F08-T2-02** | Container ID format normalization (`"imp 104"`, `"IMP-104"`, `"IMP- 88"`, `"IMP-201-1"`) | `"imp 104"`, `"  IMP-104  "`, `"IMP- 88"`, `"IMP-201-1"` | All single-container variants normalize consistently (`"IMP-104"`, `"IMP-88"`), while split-container suffix `"IMP-201-1"` remains distinct from `"IMP-201"`. |
| **E2E-F08-T2-03** | Conflicting multi-version drafts (`IMP-182` clean vs `inprogress` `-17158` / `-1119` diff) | `IMP-182.xlsm` (`927/927` cartons) + `Copia de IMP-182 inprogress.xlsm` (`17172/14`, empty CR Col A) + `IMP-182 inprogress.xlsm` (`1133/14`) | Collapses to 1 Business Identity (`IMP-182`), flags `RequiresHumanReview = true` (`MultipleVersionsDetected` / `InconsistentDraftTotals`), and proposes the clean complete version without creating 3 rows. |
| **E2E-F08-T2-04** | Human review triggers for data anomalies (`0/0` cartons, `DG!C18=0`, `#REF!`) | `IMP-195`: `Bultos=0/0, Lines=154`; `IMP-218`: `DG!C18=0, CR_Items=8`; `IMP-241`: `Bultos=#REF!` | Each triggers `RequiresHumanReview = true` with specific diagnostic reason (`ZeroCartonsWithItems`, `LineCountMismatch`, `SummaryFormulaError`), blocking blind auto-commit. |
| **E2E-F08-T2-05** | `core.xml` template author ignored when `B33` differs | `core.xml dc:creator = "Hilario Lopez Rivera"`, `DG!B33 = "Jostin Lopez Sobalbarro"` vs `core.xml = "Jostin Lopez"`, `DG!B33 = "Angie Montero Rojas"` | Operational authorship decision depends **strictly** on `DG!B33` (`true` for first, `false` for second). |

---

## Feature F09: SAP Knowledge & Diagnostic Domain Models (`Acr.Core`) (`R5, R6`)

### Tier 1 — Feature Coverage (F09)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F09-T1-01** | 15-field Knowledge Entry & `EINA`/`MAST`/`STPO` table classification | KB entries for `ME80FN`, `ZDescarga`, `EINA`, `MAST`, `STPO` | All 15 metadata fields (`R5-A`) validate; `EINA`, `MAST`, `STPO` carry `ObjectType = Table` (NEVER `Transaction`); `ZDescarga` carries `IsCompanySpecific = true`. |
| **E2E-F09-T1-02** | 5 Epistemic Tiers strict separation | Procedures tagged `VerifiedCompanyProcedure`, `OfficialSapDocumentation`, `EnvironmentValidated`, `GeneralSapReference`, `UnverifiedHypothesis` | Domain validator enforces distinct tier enum and rejects displaying/promoting an `UnverifiedHypothesis` as `EnvironmentValidated` or `VerifiedCompanyProcedure` without explicit confirmation. |
| **E2E-F09-T1-03** | 5 Fragment Lifecycle States (`Imported`..`Deprecated`) & provenance | Fragment transitioning `Imported -> Extracted -> Reviewed -> EnvironmentValidated -> Deprecated` with `SourceDoc`, `PageOrSection`, `OriginalLocation` | Valid forward transitions succeed; `Deprecated` fragments are excluded/demoted from active recommendations while retaining full source provenance. |
| **E2E-F09-T1-04** | 10-Section Guided Troubleshooting & context non-invention (`R5-C, R5-G`) | `DiagnosticContext(Container="IMP-103", Po43="4300006910", Material="1455103", DiscrepancyType=SHORTAGE, ShortageUnits=24)` with `Entrega18=null, Hu=null` | Output troubleshooting model contains all **10 required sections** (`R5-C`), injects `IMP-103`, `4300006910`, `1455103`, and leaves `Entrega18`/`Hu` `null`/unpopulated (never fabricates `"1800..."` or `"80..."`). |
| **E2E-F09-T1-05** | Deterministic ranking & Observed vs Suspected vs Confirmed separation | Query `"MB51 101 faltante"` + `DiagnosticContext(DiscrepancyType=SHORTAGE)` over seeded entries (`MB51`, `ME80FN`, `VKP2`, `ZDescarga`) | Ranking is `100%` deterministic across repeated calls (`MB51`/shortage guides rank above `VKP2`); `InvestigationRecord` preserves separate fields for `ObservedDiscrepancy`, `SuspectedCause`, and `ConfirmedRootCause`. |

### Tier 2 — Boundary & Corner Cases (F09)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F09-T2-01** | `VKP2` retail-price-not-cost domain guardrail | Query/context asking for import `unit_cost` or NCR claim calculation referencing `VKP2` | `VKP2` entry carries mandatory non-cost warning (`VKP2` = retail selling price in CRC, never `EKPO-NETPR` import cost) and is never bound as NCR cost source. |
| **E2E-F09-T2-02** | `Hoja1 MV 998` payroll PII exclusion rule | Source metadata / fragment tagged sheet `"Hoja1"`, `"Autorización de Deducción"`, `"MV 998"`, or `"cédula"` | Domain guard rejects indexing/storing fragment (`IsExcludedPii = true`), ensuring zero PII enters the KB model. |
| **E2E-F09-T2-03** | Empty/whitespace context & missing optional IDs (`PO=null, HU=null`) | `DiagnosticContext` with only `DiscrepancyType = REVUELTO_CANDIDATE` | Produces valid 10-section generic revuelto diagnostic guide requesting missing identifiers in Section 2 (`Information required to investigate`) without throwing or inventing IDs. |
| **E2E-F09-T2-04** | Invalid lifecycle state transition (skipping `Imported -> EnvironmentValidated`) | Attempt direct transition `Imported -> EnvironmentValidated` | Blocked by domain lifecycle validator (must pass `Reviewed` / explicit user/environment validation). |
| **E2E-F09-T2-05** | Promotion of investigation finding to personal KB without user confirmation | `InvestigationRecord` with `ConfirmedRootCause="Supplier mixed boxes"` and `UserConfirmedPromotion = false` | Promotion rule rejects creating a reusable personal KB entry until `UserConfirmedPromotion == true` and source/verification evidence are present. |

---

## Feature F10: Operational Data Safety Modes (`Development`, `Testing`, `Production`) & Atomic File Transactions (`R3`)

### Tier 1 — Feature Coverage (F10)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F10-T1-01** | `Development` mode blocks all writes to production `data/` and Daily Log | `Mode=Development`; attempt write to production `data/` path and production `Control_Diario_Operaciones_Jostin.xlsx` | Guard throws/blocks **before** opening any file stream; target timestamps and SHA-256 hashes remain `100%` unchanged. |
| **E2E-F10-T1-02** | `Testing` mode allows writes ONLY inside isolated temp root while blocking production paths | `Mode=Testing`, `IsolatedTempRoot="%TEMP%\AcrE2E\..."`; (a) write to temp root, (b) attempt write to production `data/` or production Daily Log | (a) succeeds inside isolated temp directory; (b) blocks/throws before modifying production paths. |
| **E2E-F10-T1-03** | `Production` mode protects source QC/templates read-only while allowing controlled Daily Log save | `Mode=Production`; (a) attempt overwrite of source QC/closure template `.xlsm`, (b) controlled save on isolated test-configured Daily Log target | (a) blocked/throws (source QC & closure templates always RO); (b) succeeds via `AtomicFileTransaction`. |
| **E2E-F10-T1-04** | Happy-path `AtomicFileTransaction` commit lifecycle | Valid existing workbook `target.xlsx` in temp dir; transaction writes updated copy | Creates `target.xlsx.bak` matching pre-commit SHA-256, writes `target.xlsx.tmp`, validates `.tmp`, atomically replaces `target.xlsx`, removes `.tmp`, and passes post-commit verification. |
| **E2E-F10-T1-05** | `~$` OneDrive/M365 lockfile detection aborts transaction safely | Temp dir containing `target.xlsx` AND sibling `~$target.xlsx` lockfile (`daily_log/` fixture) | Transaction detects `~$` lockfile and aborts with clear lock error before creating/replacing `target.xlsx`; original `target.xlsx` bytes & SHA-256 remain untouched. |

### Tier 2 — Boundary & Corner Cases (F10)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F10-T2-01** | Exclusive file handle held by another process | Open `target.xlsx` with `FileShare.None` in test harness, then invoke `AtomicFileTransaction` | Fails cleanly with controlled lock/IO exception; does not leave truncated `target.xlsx` or corrupt `.bak`. |
| **E2E-F10-T2-02** | Validation failure on `.tmp` aborts commit and preserves original + `.bak` | Payload or validation delegate fails during `.tmp` verification step | Original `target.xlsx` remains byte-identical to pre-transaction state; `.bak` retains valid backup; `.tmp` is cleaned up or quarantined. |
| **E2E-F10-T2-03** | Path traversal / canonicalization bypass attempt against `Development`/`Testing` guard | Path using `..`, trailing spaces, or mixed casing resolving inside production `data/` | Canonicalized `Path.GetFullPath` comparison detects production root and blocks write. |
| **E2E-F10-T2-04** | First-time file creation when target does not yet exist | Non-existent `new_export.xlsm` in `Testing` temp dir | Writes `.tmp`, validates, commits `new_export.xlsm` cleanly without failing on missing `.bak` source. |
| **E2E-F10-T2-05** | Pre-existing `.bak` / stale `.tmp` leftover from prior crashed run | Temp dir has `target.xlsx`, stale `target.xlsx.tmp`, and previous `target.xlsx.bak` | Safely replaces/refreshes `.bak` from current valid `target.xlsx`, overwrites stale `.tmp`, and completes verified commit. |

---

## Feature F11: `QcReader` Streaming QC `.xlsm` Reader (`R3`)

### Tier 1 — Feature Coverage (F11)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F11-T1-01** | Canonical Row-6 QC layout & metadata extraction | `qc/qc_canonical_clean.xlsm` + `qc/qc_edge_cases.xlsm` | Extracts exact metadata (`container_id`, `responsible`, `nationalization_po`, `date`), parses all primary items, matches `TotalCalculatedUnits` & `TotalCalculatedCartons`, and populates `FlaggedItems`. |
| **E2E-F11-T1-02** | Sheet-signature resolution (not filename-based) | `qc/SYNTH_IMP81_NOCALIDAD.xlsm` (`CONTROL DATA` without `"calidad"`/`"qc"`) vs `closures/IMP-184_sin_control_calidad.xlsm` (`CIERRE RECEP` only) | First resolves `"CONTROL DATA"` and parses items accurately (`IMP-81` pattern); sheet-signature pre-filtering rejects `closures/IMP-184_sin_control_calidad.xlsm` (`CONTROL DATA` absent) before Row-12 fallback can emit zero-unit phantom rows. |
| **E2E-F11-T1-03** | Multiple Row-6 column-shift layouts | `qc/qc_no_tarima_layout.xlsm` (#2), `qc/qc_shifted_spacer_layout.xlsm` (#3), `qc/qc_no_comp_cartons_layout.xlsm` (#4), `qc/qc_numero_tarima_layout.xlsm` (#5 `NUMERO DE TARIMA`), `qc/qc_material_col_g_layout.xlsm` (#12) | Dynamic Row-6 header mapper resolves every column accurately across all shifted layouts with identical `TotalCalculatedUnits` (`1080`) and `TotalCalculatedCartons` (`19`). |
| **E2E-F11-T1-04** | Formula cells, `sharedStrings.xml`, and `inlineStr` in QC columns (`IMP 188-189` `=W7/AD7`) | `qc/qc_canonical_clean.xlsm` (`sharedStrings.xml` `t="s"` + `=W7/AD7` cached values) + `qc/qc_inlinestr.xlsm` (`t="inlineStr"` without `sharedStrings.xml`) | Streaming reader reads cached formula values (`data_only=True` equivalent) and resolves both `t="s"` shared strings and `t="inlineStr"` cells accurately (`1080` units / `19` cartons) without returning formula text as `null`. |
| **E2E-F11-T1-05** | Child SKU exclusion from aggregate primary totals + footer filtering | `qc/qc_canonical_clean.xlsm` (parent 7-digit SKU `1455103` + two 10-digit child SKUs `1455103001`/`1455103002` + `"TOTAL"` / `"FIRMA"` rows) | Child SKUs have `IsChildSku=true, ParentMaterial="1455103"` and are **excluded** from `TotalCalculatedUnits`/`TotalCalculatedCartons`; footer rows ignored. |

### Tier 2 — Boundary & Corner Cases (F11)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F11-T2-01** | Trailing empty XML bloat & `15`-empty-row early exit | `qc/qc_bloat_streaming.xlsm` (`19` data rows + `25,000` trailing empty `<row>` elements = `25,027` total rows, `~401 KB` uncompressed `sheet1.xml`, configurable via `--bloat-rows`; host `data/` `IMP-100` covers `70–82 MB`) | Completes in `< 0.50s` (target `<= 0.15s–0.25s`) with bounded memory (`no OOM`), stopping after `15` consecutive blank rows post-header and yielding exact `19` items (`2280` units / `95` cartons). |
| **E2E-F11-T2-02** | Header exclusion guards (`material` vs `Material proveedor`, `comp_cartons` vs `UNIDADES BULTOS COMPLEMENTOS`) | `qc/qc_no_comp_cartons_layout.xlsm` (commit `14240a8` trap) | `material` maps strictly to `"Material"` (NOT `"Material proveedor"`); `comp_cartons` remains unmapped (`null`) instead of falsely binding `"UNIDADES BULTOS COMPLEMENTOS"`. |
| **E2E-F11-T2-03** | Cell token normalization (`"-"` / `"c"` / `"x"` / `"nan"` / float rounding `12.0`) | `qc/qc_edge_cases.xlsm` containing `"-"`, `"C"`, `"X"`, `"nan"`, `12.0`, `"-48"` | `"-"`, `"C"`, `"X"`, `"nan"` convert cleanly to `null`; `12.0` -> `12`; `"-48"` -> `-48`. |
| **E2E-F11-T2-04** | Pre-existing QC cached `#VALUE!` / `#REF!` in summary/footer cells (`IMP-327`, `IMP-348`) | `qc/qc_edge_cases.xlsm` with `#VALUE!` in `"TOTAL"`/`"RESUMEN"` rows and `0` cached from `=IFERROR(...)` | Does not throw unhandled exception; parses valid merchandise rows cleanly and handles error tokens defensively. |
| **E2E-F11-T2-05** | Missing file & `~$` Office lockfile handling | (a) Non-existent `missing_qc.xlsm`; (b) `historical_sync/~$IMP349_Lockfile.xlsm` & `closures/~$SYNTH_LOCK.xlsm` (165-byte non-ZIP lockfiles) | (a) throws `FileNotFoundException`; (b) directory/batch ingestion skips `~$` files before opening ZIP, or direct call rejects non-ZIP cleanly without crashing process. |

---

## Feature F12: `ClosureExporter` Macro-Safe Export & 3-Way Formula Audit (`R3`)

### Tier 1 — Feature Coverage (F12)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F12-T1-01** | Verbatim VBA macro & workbook artifact preservation | `closures/closure_template_clean.xlsm` (`vbaProject.bin`, `styles.xml`, `xl/_rels/workbook.xml.rels`, `definedNames`) | Exported `.xlsm` has **identical SHA-256** for `xl/vbaProject.bin`, preserves `styles.xml`, `xl/_rels/workbook.xml.rels` (`sharedStrings.xml` rel), and `definedNames` (including `_xlnm._FilterDatabase` and `_xlnm.Print_Area`), and opens cleanly. |
| **E2E-F12-T1-02** | Standard 58-column `CIERRE RECEP` population (`B5:B7`, `H`, `AS:AT`, `AU`, `AW:AY`, `AZ`) | `closures/closure_template_clean.xlsm` + `ComparisonSummary` | `B5 = total_declared_cartons`, `B6 = total_physical_cartons`, `B7 = "=+B6-B5"`; `AS:AT` & `AW:AY` populated; `H`, `AU`, `AZ` retain formulas (never overwritten with literal `"0"`). |
| **E2E-F12-T1-03** | Revuelto column population (`J:M`) & broken `Col H` formula repair | Template row with deleted/missing `Col H` formula + paired `REVUELTO_CANDIDATE` (`-24` partner `"B"`, `+24` partner `"A"`) | `Col H` is repaired to `=+G{r}-J{r}-N{r}-O{r}+L{r}+P{r}+Q{r}`; shortage row writes `J=24, K="B"`; overage row writes `L=24, M="A"`. |
| **E2E-F12-T1-04** | `'Bultos'` sheet population (`B..J`, rows `3+`) & trailing row cleanup | Template whose `'Bultos'` sheet has `15` pre-filled rows (`closures/closure_template_clean.xlsm`), exporting `5` current items | Rows `3..7` of `'Bultos'` reflect the `5` exported items (`B..J` linked/populated); stale rows `8..17` are cleared (`0` stale template rows left). |
| **E2E-F12-T1-05** | 3-Way Formula Integrity Audit classification | (a) `closures/closure_template_clean.xlsm`; (b) `closures/closure_preexisting_errors.xlsm` (`B5:B7 #REF!`); (c) guarded `=IF(ISNA(VLOOKUP(...))=TRUE,"",VLOOKUP(...))` in `Col B/C` | (a) `ExporterIntroducedErrors == 0`, `IsValid == true`; (b) inspects **both** formula text and cached `<v>` error nodes, flagging all 9 pre-existing error cells (`PreExistingSourceError`, even in rows `< 13`), `ExporterIntroducedErrors == 0`; (c) classifies guarded optional `VLOOKUP` as `LegitimateOptionalNa`. |

### Tier 2 — Boundary & Corner Cases (F12)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F12-T2-01** | Rows `> 299` and `> 500` dynamic scanning + omitted `<dimension>` tag | `closures/closure_over500_rows.xlsm` (`548` items ending at `max_item_row = 560`, duplicate `1455103` at rows `15` & `350`, row `413` `7000400`, omitted `<dimension>`) | Every material at rows `350`, `413`, `560` is matched and populated accurately (zero `range(13, 300)` or `500` truncation; host `data/` `IMP-187` covers `max_row = 612` / `231` items). |
| **E2E-F12-T2-02** | `Bulto Cero` rule (`received_real_units == 0` or `Part != "A"`) | Item with `PhysicalUnits=0, PhysicalCartons=0, DeclaredUnits=100, DeclaredCartons=2` (and a `Part B` item) | Sets `AS=0, AT=0` (and `Bulto Cero` rules for `AW:AX:AY`), while **never** overwriting `AU` (`=+AT{r}-H{r}`) or `AZ` (`=+AX{r}-H{r}`) formulas with `"0"`. |
| **E2E-F12-T2-03** | Multi-PO duplicate `material_code` rows in `CIERRE RECEP` (`RK-07`) | `closures/closure_over500_rows.xlsm` (`"1455103"` under `4300001001` at row `15` and `4300001002` at row `350`) | Both rows `15` and `350` are populated accurately according to `(material_code, po_43, source_row)` — neither skipped at `range(13, 300)` nor overwritten with the wrong PO's quantity. |
| **E2E-F12-T2-04** | Exporter-introduced error trap & pre-existing 5-type error differentiation | (a) Malformed formula written by exporter vs (b) `closures/closure_preexisting_errors.xlsm` (`IMP-241` `B5:B7 #REF!`, `IMP-244` `VLOOKUP #REF!`, `IMP-256` literal `'#N/A'`, `IMP-360` `U14='X'` `#VALUE!` in `V14`/`AB14`, `IMP-90` unguarded `#N/A` in `AT16`/`AU16`) | (a) `ExporterIntroducedErrors > 0` fails export validation gate; (b) `FormulaIntegrityAuditor` inspects both formula strings AND cached `<v>` error nodes to detect all 9 error cells across all 5 patterns (`PreExistingSourceError` / unguarded `#N/A`). |
| **E2E-F12-T2-05** | Missing `'CIERRE RECEP'` sheet or non-existent template | `closures/closure_missing_cierre.xlsm` (valid `.xlsm` with `Bultos`, `PROV Y CATEG`, `DATOS GENERALES`, `xl/vbaProject.bin`, omitting `'CIERRE RECEP'`), or missing path | Throws `FileNotFoundException` on missing template; `Validate` on `closures/closure_missing_cierre.xlsm` returns `IsValid=false` with explicit missing-sheet error. |

---

## Feature F13: `DailyLogRepository`, `ReceivingPipeline` & Local JSON Reports (`R1, R3, R4`)

### Tier 1 — Feature Coverage (F13)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F13-T1-01** | Read cloned/synthetic Daily Log (5 sheets, 22 rows) | Isolated copy of `daily_log/Control_Diario_Synthetic.xlsx` in `Testing` temp dir | Parses all `22` rows (`tblMaestroContenedores` `A5:R27`), `6` stages, `Col O`/`Col R` formula text (uncached `<v>` = `None` under `data_only=True`), `tblParteB` (`10` rows), `tblNotificaciones`, and `Catalogos` (`A1:H7`) without unhandled exceptions while detecting pre-seeded `IMP-202` duplicate (rows `6` & `21`) and `Dashboard General!B45`. |
| **E2E-F13-T1-02** | Insert new container row with dynamic `tblMaestroContenedores` & validation resizing | Add Row `28` (`IMP-410`) on isolated copy of `daily_log/Control_Diario_Synthetic.xlsx` | `tblMaestroContenedores` XML `ref` updates to `A5:R28`, data validations (`I..N`) cover Row `28`, `.bak` is created, and workbook re-opens validly. |
| **E2E-F13-T1-03** | Duplicate container detection on insert (`IMP-202` rule) | Attempt to register `IMP-202` (pre-seeded in `daily_log/Control_Diario_Synthetic.xlsx` rows `6` & `21`) without explicit override | Repository/service detects duplicate `NormalizedContainerId` and blocks or flags duplicate conflict before write. |
| **E2E-F13-T1-04** | End-to-end `ReceivingPipeline` execution (`QcReader -> QuantityEngine -> DiscrepancyClassifier -> ClosureExporter`) | `qc/qc_canonical_clean.xlsm` + `closures/closure_template_clean.xlsm` in `Testing` temp dir | Executes full pipeline cleanly, produces valid exported `.xlsm` (`vbaProject.bin` SHA-256 preserved, `0` exporter errors), and returns `ComparisonSummary` + `PrePrintValidationResult`. |
| **E2E-F13-T1-05** | Structured local JSON processing report & operation history persistence (`R1, R4`) | Run `ReceivingPipeline` twice in `Testing` temp dir | Writes valid JSON report containing all required fields (`Timestamp`, `QcSource`, `ClosureTemplate`, `MaterialsProcessed`, `MatchedCount`, `MissingCount`, `ExcessCount`, `RevueltosCount`, `ExportDestination`, `ValidationResult`) and appends 2 entries to local JSON operation history without overwriting entry #1. |

### Tier 2 — Boundary & Corner Cases (F13)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F13-T2-01** | Inserting past Row `57` (template validation limit `6..57`) | Isolated copy of `daily_log/Control_Diario_57Rows.xlsx` (`52` rows `6..57`, `tblMaestroContenedores` `A5:R57`), then insert Row `58` | Data validations (`I..N`) dynamically extend to cover Row `58` (not capped at `57`), and `tblMaestroContenedores.ref` becomes `A5:R58`. |
| **E2E-F13-T2-02** | Existing live anomalies preserved safely (`IMP-202` dup, shared PO, `B45` text) | Read & update unrelated row (`IMP-399`) on `daily_log/Control_Diario_Synthetic.xlsx` containing `IMP-202` duplicate (rows `6` & `21`), shared PO (`IMP-320`/`IMP-213`), and `Dashboard General!B45` stray text `"(AC22*AD22) + SI(AE22=\"\"; AF22; AE22*AF22)"` | Update does not crash on `IMP-202` duplicate, shared PO, or stray text `Dashboard General!B45`, preserves `B45` intact, and surfaces integrity warnings. |
| **E2E-F13-T2-03** | `ReceivingPipeline` in Blind Setup mode (no QC file supplied) | `closures/closure_template_clean.xlsm` without QC file | Pipeline runs cleanly without NRE, records Blind Setup / 0 physical QC items in JSON report, and preserves template integrity. |
| **E2E-F13-T2-04** | Corrupt/truncated operation history JSON recovery | Temp operation history JSON truncated/corrupted before pipeline run | Pipeline handles corrupt history gracefully (backs up corrupt file or initializes clean history without crashing reconciliation/export). |
| **E2E-F13-T2-05** | Daily Log atomic write blocked by `~$` lockfile in `Testing` temp dir | Place `daily_log/~$Control_Diario_Synthetic.xlsx` beside `daily_log/Control_Diario_Synthetic.xlsx` in temp dir, attempt Daily Log save | Aborts cleanly before modifying temp Daily Log; existing rows remain intact. |

---

## Feature F14: `HistoricalClosureScanner` Infrastructure (`R4`)

### Tier 1 — Feature Coverage (F14)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F14-T1-01** | Sheet-signature discovery ignoring filename filters | `historical_sync/` with `historical_sync/IMP104_Clean.xlsm` (`CIERRE RECEP`), `historical_sync/Macro para cierre hannia.xlsm` (`CIERRE RECEP`), `historical_sync/IMP218_Hybrid.xlsm` (`CIERRE RECEP` + `CONTROL DATA`), and `qc/qc_canonical_clean.xlsm` (`CONTROL DATA` only) | Discovers all `3` `CIERRE RECEP` closures (including non-`imp-` filename `Macro para cierre hannia.xlsm` and hybrid `IMP218_Hybrid.xlsm`) and excludes the QC-only file. |
| **E2E-F14-T1-02** | Strict `'DATOS GENERALES'!B33` authorship vs `sharedStrings.xml` (`IMP-209` trap) | `historical_sync/` containing `historical_sync/IMP104_Clean.xlsm`, `historical_sync/IMP326_Typo.xlsm`, and `historical_sync/IMP209_SharedStringsTrap.xlsm` | Includes `IMP-104` and `IMP-326` (`Sobalbaro` typo); **excludes `IMP-209`** (`0` false positives from `sharedStrings.xml`). |
| **E2E-F14-T1-03** | Business Identity + Provenance extraction & shared PO (`IMP-224/225/226`) | `historical_sync/IMP224_SharedPO.xlsm`, `historical_sync/IMP225_SharedPO.xlsm`, `historical_sync/IMP226_SharedPO.xlsm` sharing `PO=4300007562` | Extracts `3` distinct `HistoricalSyncRecord`s with separate `(ContainerId, Pedido41)` identities and full provenance (`RelativePath`, `FileSha256`, `VbaSha256`, `DiscoveredAt`). |
| **E2E-F14-T1-04** | Moved/copied file deduplication (`IMP-182` drafts & copies) | `historical_sync/IMP182_Final.xlsm`, `historical_sync/Copia de IMP182 inprogress.xlsm`, `historical_sync/drafts/IMP182_inprogress.xlsm` | Collapses all `3` files to **1** container identity (`IMP-182`, `4100002283`) with `RequiresHumanReview=true` and `0` duplicate Daily Log inserts upon import. |
| **E2E-F14-T1-05** | Idempotent re-scan & controlled 1-click import into isolated Daily Log | Scan `historical_sync/` -> import new records into temp Daily Log -> move/copy one `.xlsm` to new subfolder -> re-scan | Second scan reports `0` new records (`AlreadySynced` for all imported containers) and produces `0` duplicate rows in Daily Log. |

### Tier 2 — Boundary & Corner Cases (F14)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F14-T2-01** | `~$` lockfile & corrupt/non-ZIP `.xlsm` inside scan tree (`IMP-349` trap) | `historical_sync/` containing `historical_sync/~$IMP349_Lockfile.xlsm` (165 bytes) + valid closures | Scanner skips `~$` files silently without throwing `InvalidDataException` and completes scanning all valid closures. |
| **E2E-F14-T2-02** | `core.xml` `"Hilario Lopez Rivera"` ignored when `B33` differs | `historical_sync/IMP104_Clean.xlsm` (`core.xml dc:creator="Hilario Lopez Rivera"`, `DG!B33="Jostin Lopez Sobalbarro"`) vs `historical_sync/IMP209_SharedStringsTrap.xlsm` (`DG!B33="Angie Montero Rojas"`) | Strictly uses `DG!B33` (first matched, second ignored). |
| **E2E-F14-T2-03** | Human-review anomaly triggers (`#REF!`, `0/0` cartons, `DG!C18=0`) | `closures/closure_preexisting_errors.xlsm` (`#REF!` `IMP-241` profile), `historical_sync/IMP195_ZeroCartons.xlsm` (`0/0` cartons), `historical_sync/IMP218_Hybrid.xlsm` (`DG!C18=0`) | All flagged with `RequiresHumanReview=true` and specific issue tags (`SummaryFormulaError`, `ZeroCartonsWithItems`, `LineCountMismatch`). |
| **E2E-F14-T2-04** | Empty scan directory & directory with nested non-Excel files | Dir with 0 `.xlsm` files, and dir with `.xls`/`.pdf`/`.msg` only | Returns empty `HistoricalSyncPreview` (`0` errors, `0` crashes). |
| **E2E-F14-T2-05** | Multi-PO mismatch between `DG!C8` (`-7309-8144`) and `CR!A` | `historical_sync/IMP326_Typo.xlsm` (`DG!C8="4300007309-8144"`, `CIERRE RECEP!A="4300007309"`) | Scanner extracts cleanly without failure and surfaces multi-PO metadata discrepancy note if applicable. |

---

## Feature F15: Full Regression Parity Gate (64 QC & 81 Closures) (`R6`)

### Tier 1 — Feature Coverage (F15)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F15-T1-01** | Full 51 Python-matched QC corpus parity against `src/qc_reader.py` | Read-only run in `Testing` mode over the `51` Python-matched QC files in `data/` (or synthetic parity fallback when `data/` absent) | `51/51` parse with `0` exceptions; `total_calculated_units`, `total_calculated_cartons`, `total_reported_units`, `total_reported_cartons`, and `flagged_items` match Python baseline `100%` (`IMP-103` = `267` items / `1038` cartons; `IMP-298` = `1036` calc vs `1032` rep). |
| **E2E-F15-T1-02** | Extended `64` QC sheet-signature corpus (`63` standalone + `1` hybrid `IMP-218`) | All `.xlsm` files in `data/` identified by `CONTROL DATA` sheet signature | `64/64` genuine QC files discovered and parsed cleanly (including `IMP 81.xlsm`, `imp 207.xlsm`, `IMP244.xlsm`); average parse time `< 0.25s/workbook`; `3` `SIN CONTROL CALIDAD` false-positive closures excluded from QC set. |
| **E2E-F15-T1-03** | Full `81` `CIERRE RECEP` closure audit (`77` non-variant + `4` variant) | All `81` `CIERRE RECEP` `.xlsm` files in `data/` validated in `Testing` mode | `81/81` open and audit cleanly; exact `5` pre-existing error workbooks flagged (`IMP-241`, `IMP-244`, `IMP-256`, `IMP-360`, `IMP-90`); `76/81` clean. |
| **E2E-F15-T1-04** | Macro-safe roundtrip export parity across representative closures in temp dir | Copy closures (clean, `>500` rows `IMP-187`, error closures, hybrid `IMP-218`) to `Testing` temp dir and run export | `100%` exported closures have `vbaProject.bin` SHA-256 identical to source and `0` `ExporterIntroducedError`s. |
| **E2E-F15-T1-05** | Formal Regression Exception Report & production non-mutation | Full Parity Gate execution in `Testing` mode | Emits structured Exception Report listing `FilesProcessed`, `Passed`, `RequiringReview`, `Unprocessable`; `ProductionImmutabilityGuard` confirms **0 bytes** created/modified/deleted in `data/` or `Control_Diario_Operaciones_Jostin.xlsx`. |

### Tier 2 — Boundary & Corner Cases (F15)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F15-T2-01** | `~$IMP-349...xlsm` lockfile resilience during full corpus walk | Full `data/` walk including `~$IMP-349 (BORRADOR)  AQUA FASHION S.A..xlsm` (or synthetic `~$` lockfile) | Lockfile skipped cleanly (`147` readable `.xlsm` processed, `0` `BadZipFile` crashes). |
| **E2E-F15-T2-02** | Historical `20` Jostin closures -> `18` unique containers parity | Full `HistoricalClosureScanner` pass over `data/` in `Testing` mode | Matches Phase 0 baseline: `20` files attributed to Jostin -> `18` unique containers + `2` `IMP-182` drafts; `IMP-209` excluded; `IMP-326` included; `IMP-224/225/226` distinct. |
| **E2E-F15-T2-03** | Performance & memory bound across bloated `70–82 MB` QC files | Batch parse all `64` QC files sequentially (or bloated synthetic benchmark) | Total QC batch completes in `< 16s` (`< 0.25s/workbook` avg) without `OutOfMemoryException`. |
| **E2E-F15-T2-04** | Graceful handling when `data/` is absent in CI / isolated environment | Run parity test when `data/` path does not exist vs synthetic mini-corpus mode | Uses explicit fallback synthetic parity corpus (`tests/e2e_fixtures/`) or explicit environment trait skip without hanging or modifying workspace files. |
| **E2E-F15-T2-05** | Zero-tolerance gate failure on single-unit or single-hash drift | Synthetic parity-gate check injecting 1-unit mismatch or 1-byte `vbaProject.bin` mutation | Regression Parity Gate **fails** with explicit file & field delta in Exception Report (never silently passes). |

---
# Part 2 — Tiers 1 & 2 Specifications (`F16`–`F22`: Desktop UI, SAP Knowledge, Diagnostics & Portable `.exe`)

## Feature F16: Bilingual Localization (ES/EN Switching) & Sober Enterprise WPF Theme (`R4`)

### Tier 1 — Feature Coverage (F16)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F16-T1-01** | Default Spanish (`es`) Startup Localization | Clean app startup without persisted language override. | Active culture/language defaults to `"es"`; main window title contains `"Asistente de Cierre de Recepción (ACR)"` (never `"ReconcileAI"`); all 4 primary navigation tabs (`Conciliación`, `Bitácora Diaria`, `Sincronización Histórica`, `Conocimiento SAP y Diagnóstico`) render in Spanish. |
| **E2E-F16-T1-02** | Runtime Language Switch `es` -> `en` -> `es` Across All 4 Views | Trigger language toggle `es` -> `en`, navigate all 4 views, then toggle `en` -> `es`. | Every bound label, grid column header, status badge, tooltip, and button text updates immediately to English (`"Receiving Closure Assistant (RCA)"`, etc.) and back to Spanish without requiring app restart or throwing binding errors. |
| **E2E-F16-T1-03** | Localization Resource Catalog Parity (`Strings.es` vs `Strings.en`) | Enumerate all resource keys in `Acr.Desktop` `Strings.es.resx` and `Strings.en.resx`. | Key sets are `100%` identical (`missing_in_en == 0`, `missing_in_es == 0`), no empty values exist, and format-placeholder tokens (`{0}`, `{count}`, `{container}`) match identically per key across ES and EN. |
| **E2E-F16-T1-04** | Sober Enterprise Color Palette & Typography Verification | Inspect `Acr.Desktop` merged XAML theme dictionaries and active styles across all 4 views. | Background/surface/chrome colors conform to neutral enterprise slate/zinc/gray Fluent/SAP tokens (`#F3F4F6`, `#FFFFFF`, `#EAECEF`, `#D1D5DB`, `#1E293B`, restrained `#1D4ED8` accent); zero dark-neon surfaces (`#0F172A`, `#00FFFF`/`#39FF14`), zero `LinearGradientBrush` neon/glow effects. |
| **E2E-F16-T1-05** | Offline-Only Asset & Zero-CDN Compliance | Inspect all XAML `Image`/`Uri`/font/icon references and resource strings in `Acr.Desktop`. | Zero `http://` or `https://` URLs in UI assets/icons (e.g., `icons8.com` absent); all icons/vector glyphs resolve offline from local assembly resources. |

### Tier 2 — Boundary & Corner Cases (F16)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F16-T2-01** | Exhaustive Zero-Emoji Scan Across UI Strings & XAML | Scan `100%` of `Strings.es.resx`, `Strings.en.resx`, `.xaml` files, and ViewModel display strings for Unicode emoji blocks (`U+1F300–U+1FAFF`, `U+2600–U+27BF`, `U+FE0F`, flag pairs `U+1F1E6–U+1F1FF`). | Exact count of emoji characters across all localized strings and XAML UI markup = `0`. |
| **E2E-F16-T2-02** | Zero AI-Buzzword & Legacy-Branding Scan | Scan all ES/EN resource values, XAML labels, and ViewModel user-facing strings against prohibited terms (`ReconcileAI`, `Copilot`, `Copiloto`, `AI`, `IA`, `LLM`, `GPT`, `Vector`, `Neural`, `Chatbot`, `Magic`, `Mágico`), excluding valid acronyms (`SAP`, `EWM`, `DUA`, `NCR`, `HU`). | `0` prohibited buzzword matches in user-facing UI text. |
| **E2E-F16-T2-03** | Missing Localization Key & Placeholder Safety Boundary | Request an intentionally non-existent localization key `"__E2E_MISSING_KEY__"` and a parameterized key with mismatched/extra args. | Lookup returns safe deterministic fallback without throwing `KeyNotFoundException` or `FormatException`; UI renders gracefully. |
| **E2E-F16-T2-04** | State Preservation Across Mid-Workflow Language Toggle | Populate Reconciliation search filter `"1455103"` and unsaved Daily Log container draft `"IMP-950"`, then switch language `es` -> `en`. | Entered search query `"1455103"`, grid selection, and unsaved Daily Log draft fields remain intact after UI labels re-translate. |
| **E2E-F16-T2-05** | Localized Status Badge & Numeric Formatting Consistency | Switch `es` <-> `en` while viewing `SHORTAGE`/`FALTANTE`, `OVERAGE`/`SOBRANTE`, `REVUELTO`, and `OK` rows with counts `1,152` and unit cost `$3.1667`. | Status badges translate/display cleanly per localization contract without `.0` float artifacts (`75`, not `75.0`) or IEEE-754 noise (`3.1666666666666665`). |

---

## Feature F17: Reconciliation View & ViewModel (`R4`)

### Tier 1 — Feature Coverage (F17)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F17-T1-01** | Direct File Picker Selection & Clean Audit Execution | Select `qc/qc_canonical_clean.xlsm` and `closures/closure_template_clean.xlsm` via ViewModel/coordinator; execute audit in `Testing` mode. | Grid populates all expected materials with badge `OK`; KPI summary (`declared`, `physical`, `cartons`, `discrepant = 0`) matches `Acr.Core` calculation without UI error. |
| **E2E-F17-T1-02** | Discrepancy Grid Status Badges (`OK`, `FALTANTE`, `SOBRANTE`, `REVUELTO`) | Load `qc/qc_reconciliation_discrepancies.xlsm` + `closures/closure_discrepancy_template.xlsm` (matched `Col E` keys `1455103` `OK`, `5033284` `-24`, `1455104` `+12`, `1455105`/`1455106` `Confirmed` revuelto `-36` / `+36`). | Grid renders distinct visual badges for `OK`, `FALTANTE`, `SOBRANTE`, and `REVUELTO`; Cardinal Discrepancy totals report `TotalShortageUnits = 60` and `TotalOverageUnits = 48` (zero offsetting). |
| **E2E-F17-T1-03** | Material-Level Filtering & SKU/Description Search | Apply status filter (`FALTANTE` only), then clear and search by partial SKU `"1455"`, vendor style `"VS-99"`, and description `"Camisa"`. | Grid row count updates deterministically to matching subset; clearing filters restores full material count; empty search matches show clean empty-state indicator. |
| **E2E-F17-T1-04** | Expandable Discrepancy Explanation & SAP Diagnostic Deep-Link | Expand a `REVUELTO` (`Confirmed`/`Probable`) or `SHORTAGE` row, inspect explanation, and click the SAP Troubleshooting deep-link action. | Expanded details display packaging breakdown, revuelto partner + confidence rationale, and child curve breakdown; clicking deep-link navigates directly to SAP Diagnostics (`F21`) pre-seeded with `MaterialCode`, `Po43`, `ContainerId`, `DiscrepancyType`, and delta units. |
| **E2E-F17-T1-05** | Local JSON Processing Report Display & Macro-Safe `.xlsm` Export | Run audit with discrepancies, inspect rendered JSON report panel, and trigger export to `%TEMP%\AcrE2E\...\IMP-901_exported.xlsm`. | JSON report file contains `Timestamp`, `QcSource`, `ClosureTemplate`, `MaterialsProcessed`, `Matched`, `Missing`, `Excess`, `Revueltos`, `ExportDestination`, `ValidationResult`; exported `.xlsm` populates `CIERRE RECEP` (`H`, `J:M`, `AS:AT`, `AU`, `AW:AY`, `B5:B7`) and preserves `vbaProject.bin` SHA-256. |

### Tier 2 — Boundary & Corner Cases (F17)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F17-T2-01** | Non-Standard QC Filename (`IMP-81` Pattern) & `~$` / `.xls` Guards | Fixture dir containing `closures/~$SYNTH_LOCK.xlsm`, `closures/SYNTH_LEGACY.xls` (BIFF), `qc/SYNTH_IMP81_NOCALIDAD.xlsm` (`CONTROL DATA` without `"calidad"`/`"qc"`), and `closures/closure_template_clean.xlsm`. | `~$` lockfile skipped; `.xls` displays explicit legacy-BIFF warning without crashing; `qc/SYNTH_IMP81_NOCALIDAD.xlsm` auto-detected as QC via sheet signature (`CONTROL DATA`), preventing false Blind Setup. |
| **E2E-F17-T2-02** | Single-File Disambiguation & Blind Setup Guard | Select folder containing only 1 `.xlsm` (`closures/closure_template_clean.xlsm`, no QC workbook). | Does NOT assign closure workbook as both QC and Closure; prompts/requires explicit Blind Setup confirmation (locking physical counts/NCR to `0`). |
| **E2E-F17-T2-03** | Multi-PO Duplicate Material Disambiguation & Rows `> 299` / `> 500` | Load `closures/closure_over500_rows.xlsm` + matching QC with same `MaterialCode = "1455103"` on two `PO 43` lines (`Row 15` `4300001001` and `Row 350` `4300001002`, rows extending to `max_item_row = 560`). | Both rows appear distinctly keyed by `(MaterialCode, Po43, SourceRow)` with zero silent overwrite; export writes Row `350` and Row `560` accurately and clears trailing `'Bultos'` rows. |
| **E2E-F17-T2-04** | 3-Way Formula Audit on Export (Pre-Existing vs Optional `#N/A` vs Exporter Error) | Export using `closures/closure_preexisting_errors.xlsm` (pre-existing `B5=#REF!`) + guarded `=IF(ISNA(VLOOKUP(...)),"",...)` lookup. | Classifies `B5` as `PreExistingSourceError`, guarded lookup as `LegitimateOptionalNa`, and confirms `ExporterIntroducedError == 0`; any injected `ExporterIntroducedError > 0` fails loudly and rolls back `.tmp`. |
| **E2E-F17-T2-05** | Locked Target File / Export Permission Failure Surface | Lock destination output `.xlsm` with an exclusive `FileShare.None` stream and trigger Export. | Export fails atomically leaving no corrupt partial `.xlsm`, surfaces explicit error message with exact path & lock reason, and records failure in JSON report/history. |

---

## Feature F18: Daily Log & Historical Sync Views & ViewModels (`R4`)

### Tier 1 — Feature Coverage (F18)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F18-T1-01** | Register New Container with All 6 Real Lifecycle Dropdowns & Auto-Rollup | Create entry on isolated temp copy of `daily_log/Control_Diario_Synthetic.xlsx` (`IMP-901`, `PO=4300008100`, `Ped41=4100002400`, `Entrega=180060999`, `EWM=41499`, `HU=112345678005809999`, Stages 1–6 completed through `Cuadrado`, `Saludo="Buenos días"`). | Entry validates, auto-computes `Col O = "6. Recepción Cerrada (Listo Recep.)"`, saves atomically (`.bak` created, `.tmp` validated & replaced), resizes `tblMaestroContenedores`, and reloads identically. |
| **E2E-F18-T1-02** | Daily Log Auto-Fill & SAP Intake Clipboard Notification Formatting | Auto-fill `"IMP-902"` (`Entrega=180061001`, `Saludo="Buenas tardes"`), trigger SAP Notification Copy action. | Notification body equals `"Buenas tardes\nSe informa que el ingreso a nivel de sistema del Contenedor 902, quedo listo."` (or `Col R` inline form) and subject `"Ingreso Listo a Sistema - Contenedor 902"`. |
| **E2E-F18-T1-03** | Historical Sync Authorship Detection & Provenance Separation | Scan `historical_sync/` containing Jostin closures (`historical_sync/IMP104_Clean.xlsm`, `IMP326_Typo.xlsm`, `IMP182_Final.xlsm` + drafts), Angie closure (`IMP209_SharedStringsTrap.xlsm`). | Only Jostin-authored closures appear as sync candidates; each candidate displays separated Business Identity (`NormalizedContainerId`, `Pedido41`) and Provenance (`FileName`, `RelativePath`, `FileSha256`, `DiscoveryTimestamp`). |
| **E2E-F18-T1-04** | Historical Sync Review Queue & 1-Click Idempotent Import | Review clean candidate `historical_sync/IMP104_Clean.xlsm` (`IMP-104`, `4100002248`) and exception candidate `historical_sync/IMP195_ZeroCartons.xlsm` (`0/0` cartons) / `IMP182_Final.xlsm` drafts, approve `IMP-104`, click import, then re-run sync. | First import adds 1 row for `IMP-104`; exception candidates stay in human-review queue; second sync detects `IMP-104` as `AlreadySynced` and produces `0` duplicate rows. |
| **E2E-F18-T1-05** | Exact 6-Stage `Catalogos` Dropdown Fidelity & Allowed Transition Advance | Open all 6 stage dropdowns in Daily Log View and advance container sequentially from Stage 1/2 -> Stage 3/4 pause -> Stage 5 `Cuadrado` -> Stage 6 `Cierre Despacho Listo`. | Dropdown option sets match `Catalogos!A1:F7` verbatim (`3`, `4`, `4`, `3`, `6`, `4` options); `Col O` transitions `1.` -> `4.` -> `6.` -> `7.`. |

### Tier 2 — Boundary & Corner Cases (F18)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F18-T2-01** | Duplicate Container UI Guard (`IMP-202` Case/Spacing Normalization) | With `"IMP-202"` present in `daily_log/Control_Diario_Synthetic.xlsx`, attempt to add `"imp 202"` or `"IMP-202"`. | Duplicate container detector catches collision immediately, blocks duplicate creation, and displays existing row's `N°`, `Fecha`, and `Pedido41`. |
| **E2E-F18-T2-02** | Illegal Lifecycle Transition Rejection vs Retroactive Warning | (A) Stage 6 -> `En Distribución Tiendas` while Stage 5 = `Pendiente Cierre`; (B) Stage 5 -> `Cuadrado` while Stage 4 = `Pendiente Revisión Final CC`; (C) Stage 2 = `Listo y Notificado` with `Entrega="-"`, `HU="-"`. | (A) and (B) rejected with explicit prerequisite error; (C) permits retroactive save with non-blocking warning banner (`IMP-195`/`IMP-217` precedent). |
| **E2E-F18-T2-03** | Historical Sync `IMP-209` `sharedStrings.xml` Trap & `IMP-326` `"Sobalbaro"` Typo | `historical_sync/IMP209_SharedStringsTrap.xlsm` (`DG!B33="Angie"`, `"Jostin Lopez Sobalbarro"` in `sharedStrings.xml`) + `historical_sync/IMP326_Typo.xlsm` (`"Jostin Lopez Sobalbaro"`). | `IMP-209` is **excluded**; `IMP-326` is **included** and attributed to Jostin. |
| **E2E-F18-T2-04** | Shared PO `4300007562` Across Distinct Containers & Duplicate Drafts | `historical_sync/IMP224_SharedPO.xlsm`, `IMP225_SharedPO.xlsm`, `IMP226_SharedPO.xlsm` sharing `PO=4300007562` AND `historical_sync/IMP182_Final.xlsm` + `Copia de IMP182 inprogress.xlsm` + `drafts/IMP182_inprogress.xlsm`. | Shared PO produces **3** distinct Business Identities; `IMP-182` drafts collapse to **1** Business Identity and trigger human-review conflict prompt before import. |
| **E2E-F18-T2-05** | OneDrive `~$` Lockfile Collision During Atomic Daily Log Save | Place `daily_log/~$Control_Diario_Synthetic.xlsx` beside target temp Daily Log and click Save. | Save aborts safely before writing `.tmp` or replacing target; original temp workbook and `.bak` remain intact; UI displays explicit OneDrive/Excel lock warning. |

---

## Feature F19: SAP Local Doc Library & Deterministic Offline Search (`R5`)

### Tier 1 — Feature Coverage (F19)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F19-T1-01** | Multi-Format Offline Document Import (PDF, DOCX, XLSX, TXT, MD, HTML) | Import `sap_docs/SYNTH_SAP_GUIDE.pdf`, `SYNTH_SOP.docx`, `SYNTH_DATA_DICT.xlsx`, `SYNTH_NOTES.txt`, `SYNTH_TCODE_REF.md`, `SYNTH_EXPORT_GUIDE.html` into isolated KB store. | All 6 files import offline; fragments created with `Imported`/`Extracted` state; 15 metadata fields supported; **SHA-256 hash of every original source file is 100% unchanged**. |
| **E2E-F19-T1-02** | Page, Section, Cell & Line Provenance Preservation | Search unique markers seeded in PDF Page 2, DOCX Heading `"2.2 Prepacks"`, XLSX `"SAP_Campos!B4"`, MD `"## MB51"`, TXT Line 42, HTML `<section id="vl31n">`. | Each returned result displays exact source file name + exact location (`Page 2`, `Section 2.2 Prepacks`, `SAP_Campos!B4`, `## MB51`, `Line 42`, `Section vl31n`). |
| **E2E-F19-T1-03** | Deterministic Offline Search & Reproducible Ranking | Execute queries `"ME80FN UMP"`, `"Mov 101 102"`, `"prepack curve"`, `"doc 41"` across two separate offline runs. | Identical result lists, scores, and rank ordering returned on every run; zero external HTTP calls, zero Qdrant/vector DB/LLM dependencies invoked. |
| **E2E-F19-T1-04** | Fragment Lifecycle Progression (`Imported` -> `Environment validated`) | Advance a fragment through `Imported` -> `Extracted` -> `Reviewed` -> `Environment validated` with validation note, then filter KB by lifecycle state. | State persists across reload; status badge updates at each step; filtering by `Environment validated` includes only validated entries; audit trail preserves provenance. |
| **E2E-F19-T1-05** | Category, Module & Synonym/Tag Filtering | Filter library by category (`Transactions`, `SOPs`, `Tables`, `Troubleshooting`, `Personal Notes`), SAP module (`MM`, `EWM`), and synonyms (`"bulto"` <-> `"carton"`, `"curva"` <-> `"BOM"`). | Matching entries return accurately across Spanish/English synonyms and metadata facets. |

### Tier 2 — Boundary & Corner Cases (F19)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F19-T2-01** | Strict `Hoja1 MV 998` Payroll PII Exclusion Boundary | Import `sap_docs/SYNTH_CLOSURE_WITH_HOJA1_PII.xlsx` containing `CIERRE RECEP` AND `Hoja1` (`Autorización de Deducción`, `MV 998`, `"Juan Pérez Test"`, `"1-0999-0888"`, `SYNTHETIC_PII_MV998_SENTINEL_DO_NOT_INDEX`). | **`Hoja1` / `MV 998` is strictly skipped (`0` fragments extracted)**; searching `"998"`, `"1-0999-0888"`, `"Juan Pérez"`, or the sentinel returns **0 results**. |
| **E2E-F19-T2-02** | `Deprecated` Fragment Suppression & Explicit Flagging | Mark `"VL31N old unverified draft"` as `Deprecated`, then search `"VL31N"` with default filter vs `"Include Deprecated"` filter. | Default search hides (or ranks strictly below active and marks non-recommended) `Deprecated` fragments; `"Include Deprecated"` reveals it with a non-dismissible `Deprecated` badge. |
| **E2E-F19-T2-03** | Corrupt, Empty & Unsupported File Import Resilience | Batch-import `sap_docs/corrupt/` (`empty.txt`, `malformed.pdf`, `corrupt.docx`, `unsupported.bin`) + valid `sap_docs/SYNTH_TCODE_REF.md` (and `sap_docs/SYNTH_SCAN_NO_TEXT.pdf` 0-text scan). | Valid `.md` imports cleanly; corrupt/empty/unsupported files produce per-file structured diagnostic errors without crashing importer or corrupting KB JSON index. |
| **E2E-F19-T2-04** | Exact-Match T-Code Boost vs Partial Substring / Multi-Token Tie-Break | Query `"MB51"` when KB contains `MB51`, `MB52`, `MB5T`, and prose mentioning `"51"`. | `MB51` exact T-code match ranks #1 deterministically ahead of `MB52`/`MB5T`; ties break deterministically (`Environment validated` > `Reviewed` > `EntryId`). |
| **E2E-F19-T2-05** | Re-Import Idempotency & Source Update Detection | Import `sap_docs/SYNTH_NOTES.txt`, promote one fragment to `Reviewed`, then re-import unchanged file vs modified file. | Unchanged re-import creates 0 duplicates and preserves `Reviewed` state; modified file flags updated provenance/re-extraction without silently wiping existing validation metadata. |

---

## Feature F20: SAP Transaction Encyclopedia & Discovery Guides (`R5`)

### Tier 1 — Feature Coverage (F20)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F20-T1-01** | Complete Initial R5-B Encyclopedia Coverage | Search and open all 11 initial R5-B entries (`ME80FN`, `SE16N`, `MB51`, `VL31N`, `RFUI`, `MIGO`, `ME21N`/`ME22N`/`ME23N`, `ZDescarga`, `EINA`, `MAST`, `STPO`) plus `/SCWM/MON`, `MM60`, `VKP2`, `MB52`, `MB5T`. | Every entry opens and populates all required fields: purpose, when to use, SAP GUI 7.00 navigation, inputs, outputs, tables/fields, related procedures, read-only vs modifying indicator, auth notes, risks/reversals, verification steps. |
| **E2E-F20-T1-02** | Standard SAP vs Company-Specific `Z` Classification | Open `ME80FN` / `MB51` vs `ZDescarga` / `/SCWM/MON` (`Z1`, `ZALB`, `Z001`, `Z970`, `Y999`). | Standard transactions display `Standard SAP`; `ZDescarga` and Ekono parameters (`C001`, `Z1`, `ZALB`, `Z970`, `43/41/45/40`) carry explicit `Company-Specific` tag and caveat. |
| **E2E-F20-T1-03** | Read-Only vs Modifying Operation & Risk Separation | Compare `ME23N` / `MB51` / `SE16N (ACTVT=03)` against `MIGO` / `ME21N` / `ME22N` / `VL31N`. | Read-only entries marked `Read-Only` (`SE16N` prohibits `&SAP_EDIT`); modifying transactions marked `Modifying`, listing approvals and reversal mechanisms (`101` <-> `102`, `641` <-> `642`). |
| **E2E-F20-T1-04** | SAP GUI 7.00 Data Discovery Guides Navigation | Open Discovery Guides for: (1) PO `43`/`41`, (2) `MB51` `101-102`, (3) `EINA` `IDNLF`, (4) `PO 43 -> Ped 41 -> Entrega 18 -> EWM 414xx -> HU -> Mov 101`. | Each guide renders step-by-step SAP GUI 7.00 field paths (`ME80FN` `Shift+F4`, `UMP`, `SE16N` `EINA`, `MAST`/`STPO-MENGE`, `MB51` `C001` `101-102`). |
| **E2E-F20-T1-05** | Cross-Linked Related Procedures & Table Relationships | Inspect `ME80FN` and `MAST`/`STPO` entries and follow related links. | `MAST` links to `STPO` (`STLNR`) and `SE16N`; `ME80FN` links to `MB51` and `ME23N`; clicking related entry opens it seamlessly. |

### Tier 2 — Boundary & Corner Cases (F20)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F20-T2-01** | `VKP2` Retail-vs-Import-Cost Mandatory Guardrail | Search/open `VKP2` via `"VKP2"`, `"costo unitario"`, or `"NCR price"`. | `VKP2` entry prominently displays warning that `VKP2` shows **retail selling price (CRC)** and **MUST NEVER be used as import cost (`EKPO-NETPR`) or NCR valuation**. |
| **E2E-F20-T2-02** | `EINA` / `MAST` / `STPO` Table (Not Transaction) Classification | Open `EINA`, `MAST`, and `STPO` entries; search `"transacción STPO"` or `"transaction MAST"`. | `ObjectKind` for all three is strictly **`Table`** (queried via `SE16N` or `CS03`), NEVER mislabeled as `"Transacción"` (`validator.py:33` fix). |
| **E2E-F20-T2-03** | SAP GUI 7.00 Backend Version Agnosticism Caveat | Inspect Encyclopedia entries and Discovery Guides header/environment metadata. | Explicitly displays caveat that **SAP GUI 7.00 does NOT identify backend release (`ECC 6.0` vs `S/4HANA`)**, instructing verification via `System -> Status`. |
| **E2E-F20-T2-04** | Conflicting / Unverified Live-Check Items (`VL31N`, `RFUI`, `ZDescarga`) | Open `VL31N`, `RFUI`, and `ZDescarga` entries prior to live SAP GUI verification. | Entries do NOT masquerade as `Environment validated`; each surfaces `Reviewed` / `Conflicting` / `NeedsEnvCheck` status and required read-only check (`TSTC`/`SE93`, `T161`, `T156`). |
| **E2E-F20-T2-05** | Non-Existent / Typo T-Code Search (`VL3IN` vs `VL31N`) | Search `"VL3IN"` (`Borrador` draft typo) and `"ZZZZ999"`. | `"VL3IN"` suggests `VL31N` noting draft spelling variance; `"ZZZZ999"` returns structured empty state without fabricating an entry. |

---

## Feature F21: Context-Aware Guided Troubleshooting & Investigation History (`R4, R5`)

### Tier 1 — Feature Coverage (F21)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F21-T1-01** | Complete 10-Section Rendering Across All 8 Incident Workflows | Execute all 8 R5-C incident workflows (PO vs physical, SAP vs QC, missing prepack, missing movement, revuelto, unlocated doc, SAP/QC/closure mismatch, prior operation). | Every generated diagnostic dossier renders **all 10 required non-empty sections** (`1. Problem Summary & Objects` .. `10. Supporting Documentation References`). |
| **E2E-F21-T1-02** | Context Injection from Reconciliation Discrepancy Row | Launch troubleshooting from Reconciliation `SHORTAGE` row (`IMP-901`, `PO=4300008100`, `Ped41=4100002400`, `Mat=1455103`, `Supplier=1100000577`, `Shortage=24`). | Diagnostic summary and steps reuse all provided IDs; deterministic ranker prioritizes shortage/PO-vs-GR/PRK-vs-UN guides. |
| **E2E-F21-T1-03** | Context Injection from Daily Log Container Row | Launch troubleshooting from Daily Log row (`IMP-902`, `Ped41=4100002402`, `Entrega18=180061001`, `EWM=41488`, `HU=112345678005809001`). | Diagnostic context receives all 5 IDs; ranker prioritizes EWM/VL31N/RFUI/MB51 intake verification workflows. |
| **E2E-F21-T1-04** | Investigation History Recording & Root-Cause State Separation | Record findings, external action (`"Checked ME80FN & ZDescarga"`), verification result, and transition `ObservedDiscrepancy` -> `SuspectedCause` -> `ConfirmedRootCause`. | Investigation record saves locally to JSON history, reloads across sessions, and preserves separate fields/badges for `Observed Discrepancy`, `Suspected Cause`, and `Confirmed Root Cause`. |
| **E2E-F21-T1-05** | User-Confirmed Promotion of Validated Finding to Personal KB | Complete investigation with `ConfirmedRootCause` + source citation, click `"Promote to Personal KB"`, and confirm. | New personal KB entry appears in Local Doc Library (`F19`) with `Personal Operational Notes` category, preserved source reference, and user-confirmed state. |

### Tier 2 — Boundary & Corner Cases (F21)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F21-T2-01** | Missing Context Identifiers Never Invented | Launch diagnostics with `ContainerId="IMP-910"`, `MaterialCode="5033284"`, while `Po43=null`, `Pedido41=null`, `Entrega18=null`, `SupplierCode=null`, `Hu=null`. | Dossier renders `IMP-910` and `5033284`; missing fields remain explicitly marked as missing/not provided and listed in Section 2; **zero synthetic PO/Delivery/HU numbers are fabricated**. |
| **E2E-F21-T2-02** | 5-Tier Epistemic Separation Guardrail | Inspect diagnostic output containing both verified Ekono SOP steps and an unverified hypothesis. | UI and exported dossier visually & structurally separate all 5 epistemic tiers; unverified hypotheses are NEVER labeled as verified/official SOPs. |
| **E2E-F21-T2-03** | Unconfirmed / Cancelled Promotion Guard | Click `"Promote to Personal KB"` then Cancel/Decline; also test promotion when `ConfirmedRootCause` is empty. | Declining confirmation creates **0** new KB entries; attempting promotion with only `SuspectedCause` is blocked/warns root cause must be confirmed first. |
| **E2E-F21-T2-04** | Zero Automatic SAP Execution Verification | Inspect all diagnostic actions, commands, and service interfaces invoked during `F21` and `F19`/`F20`. | Zero SAP GUI COM scripting, RFC/BAPI, or automated `MIGO`/`VL31N`/`&SAP_EDIT` calls exist; every resolution path states human-only execution + approvals (Section 8). |
| **E2E-F21-T2-05** | Revuelto-Specific Context & Multi-Warehouse `ZDescarga` Trap | Launch diagnostics from `REVUELTO` (`Probable`: `-48`/`+48`) and separately from partial warehouse-count discrepancy. | Dossier retains Cardinal Rule (never offsets A vs B), recommends checking physical bin/size-curve tags (`STPO`/`MAST`), and reminds operator to verify `TOTAL DE UNIDADES` (`Central + Resurtido + Especiales`). |

---

## Feature F22: Operational Validation, UI Smoke/Judge Review & Portable `.exe` (`R1, R6`)

### Tier 1 — Feature Coverage (F22)

| Test ID | Equivalence Class | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F22-T1-01** | Clean WPF Desktop Startup in `Testing` Mode | Launch `Acr.Desktop.exe` with `Testing` mode configuration pointing to isolated temp workspace. | Window opens within startup timeout, displays main shell, `OperationalSafetyMode = Testing` guard active, zero unhandled exceptions or crash dialogs. |
| **E2E-F22-T1-02** | 4-View Live UI Smoke Navigation & Screenshot Capture | Navigate all 4 views (`Reconciliation`, `Daily Log`, `Historical Sync`, `SAP Knowledge & Diagnostics`), load synthetic fixtures, capture PNG screenshots. | All 4 views render non-blank, clip-free controls; 4 PNG screenshots saved for judge rubric audit. |
| **E2E-F22-T1-03** | Agent-as-Judge UI Rubric Compliance Audit | Evaluate 4 captured screenshots + XAML/strings against `R4` & Acceptance Criteria rubric. | Rubric passes `100%`: sober enterprise corporate theme (Fluent/SAP neutral grays), zero emojis, zero neon gradients, zero AI buzzwords, all required controls visible & legible. |
| **E2E-F22-T1-04** | Portable Self-Contained `.NET 10 LTS` Publish & Clean-Dir Execution | Run `dotnet publish` (`-c Release -r win-x64 --self-contained true`), copy output to isolated temp dir without installer, and launch. | Executable launches cleanly without installer, admin elevation, or network access; passes offline reconciliation + KB search smoke run. |
| **E2E-F22-T1-05** | Production Data Immutability End-to-End Verification | Record pre-run directory snapshot & SHA-256 of production `data/` + `Control_Diario_Operaciones_Jostin.xlsx`, run full E2E suite in `Testing` mode, re-verify. | Production `data/` and `Control_Diario_Operaciones_Jostin.xlsx` remain `100%` unmodified (`0` added/deleted/modified files, identical SHA-256). |

### Tier 2 — Boundary & Corner Cases (F22)

| Test ID | Boundary / Corner Condition | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|
| **E2E-F22-T2-01** | Startup with Empty/Missing Local JSON History & Unbuilt KB Store | Launch portable `.exe` in fresh temp profile where `operation_history.json`, `investigations.json`, and `sap_kb.json` do not yet exist. | App initializes cleanly with empty history/KB states (or pre-seeded reference entries) without throwing `FileNotFoundException` or `JsonException`. |
| **E2E-F22-T2-02** | Corrupted Local JSON State Recovery | Place malformed JSON (`"{bad_json"`) in local operation history / investigation store path and launch/open view. | App handles corrupt JSON gracefully with explicit non-fatal warning and safe fallback/backup rather than crashing on startup. |
| **E2E-F22-T2-03** | `Testing` Mode Guard Blocking Accidental Production Write Path | Configure `OperationalSafetyMode.Testing` and attempt export or Daily Log save targeting production `Control_Diario_Operaciones_Jostin.xlsx` or `data/` path. | `OperationalSafetyGuard` blocks write deterministically with explicit safety exception/error; production file remains byte-identical. |
| **E2E-F22-T2-04** | Portable Execution Without Global `DOTNET_ROOT` / Elevated Permissions | Run self-contained portable `.exe` in child process with cleared `DOTNET_ROOT` and standard non-elevated token. | Portable `.exe` uses bundled `.NET 10` runtime DLLs and starts cleanly without prompting for `.NET` desktop runtime install or UAC elevation. |
| **E2E-F22-T2-05** | High-DPI / Window Resize Layout Integrity Smoke Check | Resize main WPF window between minimum supported size (`1024x680`) and `1920x1080` across ES and EN modes. | No clipped buttons, overlapping labels, or collapsed discrepancy/lifecycle grids across all 4 views. |

---

# Part 3 — Tier 3: Cross-Feature Pairwise Interactions (`E2E-T3-01`..`E2E-T3-28`)

| Test ID | Interacting Features | Scenario & Data Flow | Concrete Inputs & Fixture | Observable Pass/Fail Criteria |
|---|---|---|---|---|
| **E2E-T3-01** | **F03 + F04** (`QuantityEngine` Grammar + Complement/Curve) | `packaging_string` fallback feeds curve consistency & observation flags | `CartonCalculationInput` with blank carton cols, `PackagingString="6*192 +48"`, `Curve=12`, `Uom="PRK"`, `Observation="MAS 1PK"` | Falls back to `1200` units, `6` cartons; computes `100` PRK, `0` loose; emits `QC_OBSERVATION_ADJUSTMENT`. |
| **E2E-T3-02** | **F03/F04 + F11** (`QcReader` -> `QuantityEngine`) | `QcReader` extracts shifted layout columns & formula cached values into `QuantityEngine` | `qc/qc_no_tarima_layout.xlsm` + `qc_edge_cases.xlsm` (Row 37 blank `comp_cartons`, `comp_units=48`, formula `=W7/AD7`) | `QcReader` maps shifted columns without `comp_cartons`/`UNIDADES` collision; `QuantityEngine` rescues `48` units and emits `QC_COMPLEMENT_WITHOUT_CARTON_COUNT`. |
| **E2E-T3-03** | **F11 + F05** (`QcReader` -> `DiscrepancyClassifier`) | Primary vs child SKUs & duplicate `material_code` aggregation into 3-way comparison | `qc/qc_canonical_clean.xlsm` (parent `1455103` `480u/10c`, two `10`-digit child SKUs `1455103001`/`1455103002`, and `5033284` `600u/9c`, total `1080u/19c`) vs `closures/closure_template_clean.xlsm` (plus multi-PO duplicate `1455103` in `closures/closure_over500_rows.xlsm`) | Child SKUs do not inflate primary totals (`1080u/19c`); `1455103` & `5033284` classify `EXACT_MATCH`; multi-PO `1455103` rows aggregate cleanly without overwriting. |
| **E2E-T3-04** | **F05 + F06** (Cardinal Rule + Revuelto Confidence) | Multi-discrepancy container with exact swap + independent shortage/overage | `Decl`: `A(100), B(100), C(100), D(100)`; `Phys`: `A(76), B(124)` (same style/cat as `A`), `C(85), D(110)` | `A` & `B` link as `REVUELTO_CANDIDATE` (`Confirmed`, `-24`/`+24`); `C` (`-15`) & `D` (`+10`) stay `SHORTAGE`/`OVERAGE`; `TotalShortageUnits=39`, `TotalOverageUnits=34` (never offset). |
| **E2E-T3-05** | **F05/F06 + F12** (`DiscrepancyClassifier` -> `ClosureExporter`) | Revuelto pairs, `Bulto Cero`, and multi-PO duplicate SKUs exported to 58-col `CIERRE RECEP` + `'Bultos'` | `ComparisonSummary` (`Confirmed` revuelto, `Bulto Cero`, 2 PO lines `1455103`) + `closures/closure_over500_rows.xlsm` | `J/K` and `L/M` populated; `Bulto Cero` sets `AS=AT=0` preserving `AU/AZ`; both `1455103` rows (`15` & `350`) and `7000400` (`413`) populated past row `299`; `_xlnm._FilterDatabase` + `rels` preserved; `'Bultos'` trailing rows cleared. |
| **E2E-T3-06** | **F10 + F12** (`SafetyModes` + `ClosureExporter`) | Exporting `.xlsm` in `Testing` vs `Development` mode with `~$` lockfile checks | `closures/closure_template_clean.xlsm`; target temp output in `Testing` vs overwrite of source/production `data/` | `Testing` export succeeds and preserves `vbaProject.bin` SHA-256; source/production write attempts fail before disk mutation. |
| **E2E-T3-07** | **F10 + F13** (`AtomicFileTransaction` + `DailyLogRepository`) | Concurrent `~$` lockfile and validation rollback during Daily Log update | `daily_log/Control_Diario_Synthetic.xlsx` + `~$Control_Diario_Synthetic.xlsx`; (a) with `~$`, (b) without `~$` insert Row `28` | (a) Aborts safely with 0 mutation; (b) creates `.bak`, resizes `tblMaestroContenedores` `A5:R28`, passes post-commit validation. |
| **E2E-T3-08** | **F07 + F13** (`DailyLogLifecycleValidator` + `DailyLogRepository`) | Lifecycle transition enforcement before persisting to Excel workbook | `daily_log/Control_Diario_Synthetic.xlsx`; attempt Stage 5 `Pausa sin Control Calidad` + Stage 6 `En Distribución Tiendas` vs `Bloqueado` | Invalid transition rejected **before** `AtomicFileTransaction` writes Excel; valid transition persists and computes `Col O = "4. Pausa sin CC (Montado a Ciegas)"`. |
| **E2E-T3-09** | **F08 + F14** (`HistoricalIdentityMatcher` + `HistoricalClosureScanner`) | Full historical sync across `IMP-209`, `IMP-326`, `IMP-182` copies, and shared PO `IMP-224/225/226` | `historical_sync/` tree containing all 6 variants | `Scanner` reads `DG!B33` strictly; `IdentityMatcher` yields `IMP-326` (`1`), `IMP-182` (`1` + review), `IMP-224/225/226` (`3`), and `0` for `IMP-209`. |
| **E2E-T3-10** | **F07 + F14 + F13** (Historical Sync -> Lifecycle Domain -> Daily Log Repo) | Importing scanned Jostin historical closures into isolated Daily Log and re-scanning | `historical_sync/` imported into `daily_log/Control_Diario_Synthetic.xlsx`, then scan repeated | Valid records insert with compliant 6-stage values + `Col O` rollup; re-scan sees `AlreadySynced` (`0` duplicates). |
| **E2E-T3-11** | **F05/F06 + F09** (`DiscrepancyClassifier` -> SAP `DiagnosticContext` & Ranker) | Discrepancy output seeds SAP diagnostic model without inventing missing IDs | `MaterialDiscrepancy(Mat="1455103", Type=SHORTAGE, Diff=-24)` + `Po43="4300006910"` (`Entrega18=null`) | `DiagnosticContext` passes `1455103`, `4300006910`, `SHORTAGE`; deterministic ranker surfaces `MB51`/`ZDescarga` 10-section guide with `Entrega18` uninvented. |
| **E2E-T3-12** | **F11 + F12 + F13** (`ReceivingPipeline` + JSON Report + Operation History) | Full pipeline run (`QC -> Engine -> Classifier -> Exporter -> JSON Report -> History`) | `qc/qc_bloat_streaming.xlsm` (`25,000` trailing empty rows) + `closures/closure_template_clean.xlsm` in `Testing` temp dir | Pipeline completes in `< 1s`, writes `.xlsm` with identical `vbaProject.bin` SHA-256, and writes matching JSON report + history entry. |
| **E2E-T3-13** | **F02 + F15** (Solution Isolation + Regression Parity Gate) | Pure `Acr.Core` assembly executes all parity calculations driven by `Acr.Infrastructure` | `Acr.Core.dll` (zero OpenXML/`System.IO` deps) + `Acr.Infrastructure.dll` running parity audit | `Acr.Core` isolation check passes AND full parity gate passes in `Testing` mode. |
| **E2E-T3-14** | **F11 + F09** (`QcReader` / Doc Scanner vs `Hoja1 MV 998` PII Guard) | Hybrid workbook containing `Hoja1` (`MV 998`) alongside `CIERRE RECEP` and `CONTROL DATA` | `sap_docs/SYNTH_CLOSURE_WITH_HOJA1_PII.xlsx` / `historical_sync/IMP218_Hybrid.xlsm` | `QcReader` reads `CONTROL DATA` only; `ClosureExporter` preserves `Hoja1` untouched; `Acr.Core` KB guard rejects any `Hoja1`/`MV 998` fragment. |
| **E2E-T3-15** | **F12 + F15** (3-Way Formula Audit across Error Closures vs Clean Exports) | Distinguishing pre-existing source errors from exporter-introduced errors at corpus scale | `closures/closure_preexisting_errors.xlsm` + `closures/closure_template_clean.xlsm` | Auditor inspects formulas + cached `<v>` nodes, classifies all 9 error cells (`5` patterns) as `PreExistingSourceError`, confirms `0` `ExporterIntroducedError` on clean exports, and passes parity gate. |
| **E2E-T3-16** | **F10 + F11 + F14** (`~$` Lockfile across Safety, QC Reader, and Historical Scanner) | Presence of `~$` OneDrive lockfile during QC discovery, historical scan, and Daily Log write | Temp folder with `historical_sync/~$IMP349_Lockfile.xlsm`, `closures/~$SYNTH_LOCK.xlsm`, `daily_log/~$Control_Diario_Synthetic.xlsx`, and valid `.xlsm` files | `QcReader`/`HistoricalClosureScanner` skip `~$` files cleanly; `AtomicFileTransaction` aborts write on `~$Control_Diario_Synthetic.xlsx`. |
| **E2E-T3-17** | **F17 + F03/F04/F11 + F05/F06** (Reconciliation UI + Core Engine + `QcReader`) | Load shifted QC workbook with Complement Rescue, `VARIANTE SUELTA`, and Revuelto pair into Reconciliation View | `qc/qc_no_tarima_layout.xlsm` + `qc/qc_edge_cases.xlsm` + `qc/qc_reconciliation_discrepancies.xlsm` + `closures/closure_discrepancy_template.xlsm` | Grid displays exact integer quantities matching `QuantityEngine` parity, `REVUELTO` badge with `Confirmed`/`Probable` confidence, and expandable breakdown without float errors. |
| **E2E-T3-18** | **F17 + F12 + F13** (Reconciliation UI + `ClosureExporter` + JSON Report) | Export audited closure from Reconciliation View -> validate `vbaProject.bin` SHA-256, 3-way formula audit, `'Bultos'` cleanup, JSON report | `closures/closure_preexisting_errors.xlsm` / `closure_template_clean.xlsm` | UI export invokes atomic `.xlsm` export, displays `PreExistingSourceError` / `LegitimateOptionalNa` counts, confirms `0` `ExporterIntroducedError`, and renders matching JSON report. |
| **E2E-T3-19** | **F17 -> F21 -> F20** (Reconciliation Row -> Guided Diagnostics -> Encyclopedia) | Click SAP Diagnostic deep-link on `SHORTAGE` row lacking `Entrega18`/`HU` -> inspect 10-section dossier -> click `ME80FN` link | Reconciliation `SHORTAGE` row (`Mat="1455103"`, `Po43="4300006910"`, `Entrega18=null`, `Hu=null`) | `F21` opens with `MaterialCode`, `Po43`, `ContainerId` injected (never inventing `Entrega18`/`HU`); clicking `ME80FN` opens `F20` entry with `Shift+F4`/`UMP` details & `GUI 7.00` caveat. |
| **E2E-T3-20** | **F18 -> F21 + F20** (Daily Log Stage 2 -> Diagnostics + Encyclopedia `VL31N`/`RFUI`/`/SCWM/MON`) | Select Daily Log row at Stage 2 (`Pendiente Ingreso SAP`) -> open SAP Diagnostics/Encyclopedia + generate SAP clipboard notification | `daily_log/Control_Diario_Synthetic.xlsx` row with `Entrega18`, `EWM`, `HU` | `F21` injects `ContainerId`, `Pedido41`, `Entrega18`, `EWM`, `HU`; `F20` surfaces `VL31N`/`RFUI`/`/SCWM/MON` with env-verification tag; clipboard formatter outputs exact notification. |
| **E2E-T3-21** | **F18 + F08/F14 + F07/F13** (Historical Sync UI + Scanner + Daily Log Repo) | Scan `historical_sync/` -> review -> 1-click import -> inspect Daily Log View | `historical_sync/` (`IMP-326`, `IMP-209` trap, `IMP-182` drafts, `IMP-224/225/226` shared PO) + `daily_log/` | `IMP-209` excluded; `IMP-326` + `IMP-224/225/226` imported as distinct rows; `IMP-182` resolved via review without duplicates; `tblMaestroContenedores` resizes cleanly. |
| **E2E-T3-22** | **F19 + F21 + F16** (Local Doc Library + Investigation Promotion + Bilingual/Theme) | Import `sap_docs/` (incl. `Hoja1 MV 998` trap) -> complete investigation in `F21` -> confirm promotion to personal KB -> toggle `es`/`en` and search | `sap_docs/*` + `F21` confirmed investigation | `Hoja1 MV 998` stays excluded; promoted personal note is searchable offline in `F19` with provenance; UI remains emoji-free and sober in both `es` and `en`. |
| **E2E-T3-23** | **F17 + F10 + F22** (Reconciliation UI + `OperationalSafetyGuard` + Portable `.exe`) | Run portable `.exe` in `Testing` mode and attempt export/save to production path vs isolated temp path | Published `Acr.Desktop.exe` + `qc/` + `closures/` | Production path write blocked by `F10`; isolated temp path write succeeds; operation history updates locally. |
| **E2E-T3-24** | **F18 + F07 + F16** (Daily Log Lifecycle + Transition Rules + `es`/`en` Localization) | Switch UI to `en` and test valid/invalid 6-stage transitions & duplicate container guard in Daily Log View | `daily_log/Control_Diario_Synthetic.xlsx` (`IMP-202` dup + illegal Stage 6 jump) | `Catalogos` stage values remain faithful operational strings while validation errors, warnings, and `Col O` UI labels respect localization without emojis or corrupted placeholders. |
| **E2E-T3-25** | **F19 + F20 + F21** (Deterministic Search + Encyclopedia `VKP2`/`STPO`/`MAST` + Epistemic Tiers) | Query `"unit cost VKP2"` and `"STPO curve"` in `F19`/`F20` and inspect Section 5/6/7 in `F21` | `sap_docs/SYNTH_SOP.docx`, `SYNTH_DATA_DICT.xlsx`, `SYNTH_TCODE_REF.md` | All three modules agree deterministically: `VKP2` blocked as import cost, `STPO`/`MAST` classified as `Table`, unverified items carry Tier-5/env-check caveats. |
| **E2E-T3-26** | **F17 + F18** (Reconciliation `Bulto Cero` / `Parte B` + Daily Log Stage 5 `Parcial / Pendiente Parte B`) | Audit a `Parte B` / `Bulto Cero` case in Reconciliation View, then update Daily Log Stage 5 to `Parcial / Pendiente Parte B` | `qc/` + `closures/` (`Bulto Cero` item) + `daily_log/` | Reconciliation computes `0` cartons (`Bulto Cero`) without false shortage confusion; Daily Log reflects `Parcial / Pendiente Parte B` and permits Stage 6 unlock policy. |
| **E2E-T3-27** | **F19 -> F21** (`Deprecated` Fragment Suppression in Diagnostic Ranking) | Mark a troubleshooting guide as `Deprecated` in `F19`, then trigger a matching discrepancy workflow in `F21` | `sap_docs/` with one guide transitioned to `Deprecated` | `F21` excludes or demotes and flags `Deprecated` source, prioritizing `Environment validated` / `Reviewed` procedures. |
| **E2E-T3-28** | **F16 + F19 + F22** (Theme/Zero-CDN + Offline Search + Portable `.exe`) | Run portable `.exe` with all network/AI endpoints unavailable and exercise all 4 views + offline KB search + `.xlsm` export | Published `Acr.Desktop.exe` + all isolated fixture folders | `100%` of features operate offline with zero network timeouts, zero missing icons, and zero AI/vector-DB dependencies. |

---

# Part 4 — Tier 4: Real-World Application Scenarios (`E2E-T4-01`..`E2E-T4-16`)

| Scenario ID | Scenario Name | Features Exercised | End-to-End Workflow & Fixture Set | Observable Pass/Fail Criteria |
|---|---|---|---|---|
| **E2E-T4-01** | **`IMP-103`: Prepack Receiving & Complement Rescue Workflow** | F02, F03, F04, F05, F10, F11, F12, F13 | Synthetic fallback: `qc/qc_edge_cases.xlsm` (`13` items / `123` calc vs `122` rep cartons, row 7 `"6*192 +48"` rescue), `qc/qc_numero_tarima_layout.xlsm` (Layout #5 `4` items / `19` cartons), `closures/closure_template_clean.xlsm`, `daily_log/Control_Diario_Synthetic.xlsx`; optional host `data/` `IMP-103` (`267` items / `1038` cartons). | Synthetic `qc_edge_cases` rescues `48` units (`1200u/6c/100 PRK`) and triggers all 9 flags (`5492u/123c`); Layout #5 parses `1080u/19c`; host `IMP-103` parses `267` items / `1038` cartons; `vbaProject.bin` SHA-256 unchanged; `0` exporter errors; JSON report & history match. |
| **E2E-T4-02** | **`IMP-298`: QC Omitted Complement Carton Mismatch Detection & Audit** | F03, F04, F11, F13, F15 | Synthetic fallback: `qc/qc_edge_cases.xlsm` (row `17`: `252` calc vs `240` rep, `123` calc vs `122` rep cartons); optional host `data/` `IMP-298` (`1036` calc vs `1032` rep cartons). | Synthetic `QcReader` reports `TotalCalculatedCartons == 123`, `TotalReportedCartons == 122`, flags `QC_FORMULA_MISMATCH` (`IMP-298` host: `1036` vs `1032`), and JSON report surfaces discrepancy warning. |
| **E2E-T4-03** | **`IMP-207`: Cardinal Rule, Revuelto Confidence, Multi-PO & False `Parte B` Prevention** | F05, F06, F09, F11, F12 | Synthetic fallback: `qc/qc_no_comp_cartons_layout.xlsm` (omits `BULTOS COMPLEMENTOS`), `qc/qc_reconciliation_discrepancies.xlsm` + `closures/closure_discrepancy_template.xlsm`; optional host `data/` `IMP-207` (`ESQUEMA ALISTO PARTE B/C`). | `comp_cartons` ignores `UNIDADES BULTOS COMPLEMENTOS`; Cardinal Rule preserves `0` silent offsetting (`60` shortage, `48` overage); revueltos get `Confirmed`/`Probable`/`Possible` confidence; does NOT fabricate `PENDIENTE_PARTE_B` from `ESQUEMA ALISTO`; 10-section guide opens with no invented IDs. |
| **E2E-T4-04** | **`IMP-81`: Filename-Blind QC Discovery, `Bulto Cero` & Blind Setup Prevention** | F04, F07, F11, F12, F13 | `qc/SYNTH_IMP81_NOCALIDAD.xlsm` (`IMP-81` pattern omitting `"calidad"`/`"qc"`), `closures/IMP-184_sin_control_calidad.xlsm` (rejected as QC), `closures/closure_over500_rows.xlsm` (omitted `<dimension>`), `daily_log/Control_Diario_Synthetic.xlsx`. | Sheet-signature discovery identifies `SYNTH_IMP81_NOCALIDAD.xlsm` as QC (`1080u/19c`) and excludes `IMP-184_sin_control_calidad.xlsm`; `ClosureExporter` handles high row count without `<dimension>` truncation; `Bulto Cero` sets `AS=AT=0` preserving `AU`/`AZ` formulas. |
| **E2E-T4-05** | **`IMP-187`: High-Row-Count (`max_item_row = 560` / host `612`) Closure & `'Bultos'` Trailing Row Cleanup** | F03, F05, F10, F11, F12, F13 | `closures/closure_over500_rows.xlsm` (`548` items ending at `max_item_row = 560`, duplicate `1455103` at rows `15` & `350`, row `413` `7000400`, omitted `<dimension>`, `'Bultos'` trailing rows); optional host `data/` `IMP-187` (`231` items, `max_row = 612`). | All items past row `299` (`350`, `413`, `560`; host `612`) anchored by `(MaterialCode, Po43)` and populated accurately; `_xlnm._FilterDatabase` + `xl/_rels/workbook.xml.rels` + `vbaProject.bin` SHA-256 preserved; `'Bultos'` stale rows cleared; `0` exporter formula errors. |
| **E2E-T4-06** | **`IMP-182` + `IMP-209` + `IMP-326` + `IMP-224/225/226`: Idempotent Historical Sync into Daily Log** | F07, F08, F10, F13, F14 | Populate `Testing` temp root from `historical_sync/` (`IMP182_Final.xlsm` + 2 drafts, `IMP209_SharedStringsTrap.xlsm`, `IMP326_Typo.xlsm`, `IMP224/225/226_SharedPO.xlsm`, `~$IMP349_Lockfile.xlsm`) + `daily_log/Control_Diario_Synthetic.xlsx`. | Scan 1 skips `~$` and `IMP-209`, includes `IMP-326`, keeps `IMP-224/225/226` separate (`3` rows), collapses `IMP-182` drafts to `1` row + review flag; Scan 2 after move produces `0` new rows (`100%` idempotent); `tblMaestroContenedores` resizes cleanly with `.bak` preserved. |
| **E2E-T4-07** | **`IMP-241` + `IMP-244` + `IMP-360` + `IMP-90`: 3-Way Formula Error Containment & Pre-Print Gate** | F10, F12, F13, F15 | Run `ClosureExporter.Validate` and export audit in `Testing` temp dir on `closures/closure_preexisting_errors.xlsm` (all 5 pre-existing error profiles + guarded `IF(ISNA(VLOOKUP(...)),"",...)`). | Auditor inspects formulas + cached `<v>` nodes, detects all `9` pre-existing error cells (`B5:B7`, `V14`, `AB14`, `AT16`, `AU16`, `AS17`, `AY29`) as `PreExistingSourceError`/unguarded `#N/A`, leaves guarded `ISNA(VLOOKUP)` as `LegitimateOptionalNa`, and enforces `0` `ExporterIntroducedError`. |
| **E2E-T4-08** | **`IMP-100`: Bloated QC (`25,000` trailing rows / host `83.3 MB`) + Blind Setup -> QC Arrival Daily Log Lifecycle** | F07, F10, F11, F13 | Advance `daily_log/Control_Diario_Synthetic.xlsx` container lifecycle `1. En Cola` -> `4. Pausa sin CC` -> ingest `qc/qc_bloat_streaming.xlsm` (`25,000` trailing empty `<row>` elements, `25,027` total rows, `~401 KB` XML, configurable via `--bloat-rows`; host `IMP-100` `83.3 MB`) -> advance to `Cuadrado` -> `Cierre Despacho Listo` + copy SAP notification. | Bloated QC parses in `< 0.50s` (`<= 0.25s` target, `19` items / `2280u` / `95c`); Stage 6 stays blocked until Stage 5 reaches `Cuadrado`; `Col O` transitions `1 -> 4 -> 6 -> 7`; SAP notification formats `"Ingreso Listo a Sistema - Contenedor 100"` cleanly. |
| **E2E-T4-09** | **`IMP-192` & `IMP-200`: Compound `+` Grammar & Packaging String Parity** | F03, F04, F11, F12 | Run `QuantityEngine` + `QcReader` + `ClosureExporter` on `qc/qc_edge_cases.xlsm` (`'6*192 +48'`, `'56*12+672'`) and `closures/closure_template_clean.xlsm` (`'8*72+24'`). | `'6*192 +48'` = `1200`, `'8*72+24'` = `600`; `'56*12+672'` preserves `parse_packaging_expression` parity (`1344`) while `QcReader` row 18 computes `672` (`56*12`) and flags `loose == sum(terms)` ambiguity for analyst cross-check. |
| **E2E-T4-10** | **Full `Testing`-Mode Corpus Parity Gate + Production Non-Mutation Proof** | F02, F10, F11, F12, F14, F15 | Record pre-run snapshot/SHA-256 of `data/` and `Control_Diario_Operaciones_Jostin.xlsx`; execute full F15 Regression Parity Gate (`64` QC + `81` Closures + `20` Historical closures, or `47`-fixture synthetic fallback when `data/` absent) in `Testing` mode. | `64/64` QC & `81/81` closures (or `47/47` synthetic fallback) audited, Python parity matches `100%`, Exception Report generated in temp dir, and post-run production `data/` + Daily Log are `100%` byte-identical. |
| **E2E-T4-11** | **Full Container Receiving Closure UI with Revuelto & Shortage Investigation (`IMP-909` / `IMP-901`)** | F03–F06, F09–F13, F17, F19–F22 | Launch ACR in `Testing` mode -> Reconciliation View (`qc/qc_reconciliation_discrepancies.xlsm` + `closures/closure_discrepancy_template.xlsm`) -> inspect badges (`OK`, `FALTANTE`, `SOBRANTE`, `REVUELTO`) & Cardinal totals (`60` shortage, `48` overage) -> SAP Diagnostic deep-link (`F21`) -> inspect `MAST`/`STPO` & `ME80FN` (`F20`) -> confirm KB promotion (`F19`) -> export `.xlsm` & inspect JSON report. | Exact Cardinal totals (`TotalShortageUnits=60`, `TotalOverageUnits=48`) & `RevueltoConfidence=Confirmed` (`1455105`/`1455106`); `vbaProject.bin` SHA-256 identical; `0` `ExporterIntroducedError`; JSON report written; investigation persisted; personal KB entry created only after explicit confirmation. |
| **E2E-T4-12** | **Stage 1–6 Daily Log Lifecycle UI, Intake Notification & Guardrail Enforcement (`IMP-902`)** | F07, F10, F13, F18, F20 | Register `IMP-902` in Daily Log View (`daily_log/Control_Diario_Synthetic.xlsx`) -> verify `Col O` = `2.` -> attempt illegal jump to Stage 6 (blocked) -> enter `Entrega 180061001`, `EWM 41488`, `HU`, copy SAP notification -> consult `VL31N`/`RFUI`/`/SCWM/MON` -> advance Stages 4->5->6 -> save atomically -> test `imp 902` duplicate guard. | All 6 `Catalogos` stages & `Col O` transitions match Phase 0 rules; illegal transition & duplicate container blocked; exact SAP notification formatted; `.bak` + atomic `.tmp` preserves tables & validations. |
| **E2E-T4-13** | **Historical Archive Onboarding UI, Conflict Triage & Idempotent Re-Sync** | F08, F14, F18 | Point Historical Sync View to `historical_sync/` (`IMP-326`, `IMP-209`, `IMP-224/225/226`, `IMP-182` + copy, `IMP-241`) -> verify `IMP-209` excluded & human-review prompts for `IMP-182`/`IMP-241` -> 1-click import -> move file -> re-sync. | `IMP-209` ignored; `IMP-326` & `IMP-224/225/226` imported cleanly; Business Identity prevents duplicate rows on moved/copied files; re-sync adds `0` duplicates. |
| **E2E-T4-14** | **Offline SAP Knowledge Ingestion with `Hoja1 MV 998` PII Exclusion & Epistemic Hygiene** | F19, F20, F21 | Import `sap_docs/` (`PDF`, `DOCX`, `XLSX` incl. `SYNTH_CLOSURE_WITH_HOJA1_PII.xlsx`, `TXT`, `MD`, `HTML`) -> search `VKP2`, `STPO`, `ZDescarga`, `MV 998`, `"1-0999-0888"` -> promote validated entry / deprecate draft -> launch `F21` troubleshooting. | `Hoja1 MV 998` PII completely absent; source files byte-identical; `VKP2` non-cost warning, `STPO` Table tag, `ZDescarga` Company-Specific tag, `GUI 7.00` caveat enforced; `Deprecated` draft suppressed/flagged. |
| **E2E-T4-15** | **Bilingual Operator Shift Handoff (`ES` <-> `EN`) & Sober Theme Validation** | F16, F17, F18, F21 | Run workflow across all 4 views in `es` (reconcile `IMP-81` pattern, update Daily Log, preview Historical Sync, open 10-section Diagnostic dossier) -> switch live to `en` -> run automated theme/string compliance audit. | Zero lost state or exceptions on ES/EN toggle; zero emojis, zero neon gradients, zero AI buzzwords, zero missing resource keys; `IMP-81` non-`"calidad"` QC recognized via sheet signature. |
| **E2E-T4-16** | **Portable `.exe` Clean-Workstation Deployment, 4-View Smoke/Judge Rubric & Safety Audit** | F02, F10, F17–F19, F22 | Publish self-contained `.NET 10 LTS` portable `.exe` (`win-x64`) -> copy to clean isolated temp dir -> launch, smoke-walk all 4 views, capture PNG screenshots for judge rubric, run export + Daily Log save in `Testing` mode -> verify production snapshot. | Launches without installer or global `.NET 10` PATH dependency; judge rubric passes `100%` on all 4 screenshots; production `data/` & Daily Log remain `100%` untouched. |

---

# Part 5 — Fixture-to-Test Traceability Matrix (`tests/e2e_fixtures/`)

All Phase A/B E2E fixtures are materialized deterministically by `python -B tests/e2e_fixtures/generate_fixtures.py --out <temp_dir> --verify` into isolated temp subdirectories (`qc/`, `closures/`, `historical_sync/`, `daily_log/`, `sap_docs/`), leaving `58/58` pytest untouched and production `data/` immutable.

| Fixture Subdir | Logical Fixture ID & Generated Artifacts | Key Characteristics & Traps | Mapped Test IDs |
|---|---|---|---|
| **`qc/`** | **`FIX-01`**: `qc/qc_canonical_clean.xlsm`, `qc/qc_no_tarima_layout.xlsm`, `qc/qc_shifted_spacer_layout.xlsm`, `qc/qc_no_comp_cartons_layout.xlsm`, `qc/qc_numero_tarima_layout.xlsm`, `qc/qc_material_col_g_layout.xlsm`, `qc/qc_edge_cases.xlsm`, `qc/qc_inlinestr.xlsm`, `qc/qc_reconciliation_discrepancies.xlsm`, `qc/SYNTH_IMP81_NOCALIDAD.xlsm` | Row-6 shifted signatures (#1, #2, #3, #4 commit `14240a8` trap, #5 `NUMERO DE TARIMA`, #12 `Col G`), `sharedStrings.xml` + `inlineStr`, formula `BULTOS ENTEROS`, child SKUs (`len >= 10`), Complement Rescue, `VARIANTE SUELTA`, `DOS MATERIALES`, `IMP-192`/`200`/`298` patterns, matched discrepancy QC (`1455103`..`1455106`), non-`"calidad"` filename (`IMP-81`) | `E2E-F03-T1-*`, `E2E-F04-T1-*`, `E2E-F11-T1-*`, `E2E-F11-T2-02..05`, `E2E-F17-T1-01..05`, `E2E-F17-T2-01..03`, `E2E-T3-01..03`, `E2E-T3-17`, `E2E-T4-01..04`, `E2E-T4-09`, `E2E-T4-11`, `E2E-T4-15` |
| **`qc/`** | **`FIX-02`**: `qc/qc_bloat_streaming.xlsm` | `19` data rows + `25,000` trailing empty `<row>` elements (`25,027` total rows, `~401 KB` uncompressed `sheet1.xml`, configurable via `--bloat-rows`; host `IMP-100` covers `70–82 MB`) | `E2E-F11-T2-01`, `E2E-F15-T2-03`, `E2E-T3-12`, `E2E-T4-08` |
| **`closures/`** | **`FIX-03`**: `closures/closure_template_clean.xlsm`, `closures/closure_discrepancy_template.xlsm`, `closures/closure_over500_rows.xlsm`, `closures/IMP-184_sin_control_calidad.xlsm` | 58-col `CIERRE RECEP` (`A..BF`), `'Bultos'` (`15` template rows + stale trailing cleanup), `'PROV Y CATEG'`, `'DATOS GENERALES'`, synthetic OLE2 `vbaProject.bin` (pinned SHA-256), `definedNames` (`_xlnm._FilterDatabase`, `_xlnm.Print_Area`), guarded `IF(ISNA(VLOOKUP(...)),"",...)`, `548` items / `max_item_row = 560` + duplicate `1455103` (`15`/`350`), omitted `<dimension>`, `Bulto Cero`, `CIERRE RECEP`-only trap | `E2E-F11-T1-02`, `E2E-F12-T1-01..05`, `E2E-F12-T2-01..03`, `E2E-F13-T1-04..05`, `E2E-F17-T1-01..05`, `E2E-F17-T2-02..05`, `E2E-T3-05`, `E2E-T3-06`, `E2E-T3-12`, `E2E-T3-18`, `E2E-T4-01`, `E2E-T4-03..05`, `E2E-T4-11`, `E2E-T4-16` |
| **`closures/`** | **`FIX-04`**: `closures/closure_preexisting_errors.xlsm`, `closures/closure_missing_cierre.xlsm`, `closures/SYNTH_LEGACY.xls`, `closures/~$SYNTH_LOCK.xlsm` | Pre-existing `B5:B7 #REF!` (`IMP-241`), deleted `VLOOKUP #REF!` (`IMP-244`), literal `'#N/A'` (`IMP-256`), `U14='X'` `#VALUE!` (`IMP-360`), unguarded external `VLOOKUP #N/A` (`IMP-90`), missing `'CIERRE RECEP'` sheet trap, legacy OLE2 BIFF `.xls`, `165`-byte `~$` lockfile | `E2E-F10-T1-05`, `E2E-F11-T2-05`, `E2E-F12-T1-05`, `E2E-F12-T2-04..05`, `E2E-F15-T1-03`, `E2E-F17-T2-01`, `E2E-F17-T2-04`, `E2E-T3-15`, `E2E-T3-16`, `E2E-T4-07` |
| **`daily_log/`** | **`FIX-05`**: `daily_log/Control_Diario_Synthetic.xlsx`, `daily_log/Control_Diario_57Rows.xlsx`, `daily_log/~$Control_Diario_Synthetic.xlsx` (+ optional RO clone of host Daily Log) | `5` sheets, `3` tables (`tblMaestroContenedores` `A5:R27` & `A5:R57`, `tblParteB`, `tblNotificaciones`), `Catalogos!A1:H7` exact 6-stage dropdowns, `Col O` 7-state formula, `Col R` notification formula (uncached `<v>` = `None`), pre-seeded `IMP-202` duplicate + `Dashboard General!B45` stray text, Row-57 boundary, `~$` lockfile | `E2E-F07-T1/T2`, `E2E-F10-T1-04..05`, `E2E-F13-T1-01..03`, `E2E-F13-T2-01..02`, `E2E-F13-T2-05`, `E2E-F18-T1-01..02`, `E2E-F18-T1-05`, `E2E-F18-T2-01..02`, `E2E-F18-T2-05`, `E2E-T3-07`, `E2E-T3-08`, `E2E-T3-10`, `E2E-T3-20`, `E2E-T3-24`, `E2E-T3-26`, `E2E-T4-01`, `E2E-T4-04`, `E2E-T4-06`, `E2E-T4-08`, `E2E-T4-12`, `E2E-T4-15`, `E2E-T4-16` |
| **`historical_sync/`** | **`FIX-06`**: `historical_sync/IMP104_Clean.xlsm`, `historical_sync/IMP326_Typo.xlsm`, `historical_sync/IMP209_SharedStringsTrap.xlsm`, `historical_sync/IMP182_Final.xlsm`, `historical_sync/Copia de IMP182 inprogress.xlsm`, `historical_sync/drafts/IMP182_inprogress.xlsm`, `historical_sync/IMP224_SharedPO.xlsm`, `historical_sync/IMP225_SharedPO.xlsm`, `historical_sync/IMP226_SharedPO.xlsm`, `historical_sync/IMP195_ZeroCartons.xlsm`, `historical_sync/IMP218_Hybrid.xlsm`, `historical_sync/Macro para cierre hannia.xlsm`, `historical_sync/~$IMP349_Lockfile.xlsm` | `DG!B33` vs `core.xml` `"Hilario Lopez Rivera"`, `IMP-209` `sharedStrings.xml` decoy, `IMP-326` `"Sobalbaro"` typo, `IMP-182` duplicate/moved/nested drafts, shared `PO=4300007562` (`3` distinct containers), non-`imp-` closure, `IMP-195` zero-carton & `IMP-218` hybrid `C18=0` traps, `~$` 165-byte lockfile | `E2E-F08-T1/T2`, `E2E-F14-T1/T2`, `E2E-F15-T2-02`, `E2E-F18-T1-03..04`, `E2E-F18-T2-03..04`, `E2E-T3-09`, `E2E-T3-10`, `E2E-T3-16`, `E2E-T3-21`, `E2E-T4-06`, `E2E-T4-13` |
| **`sap_docs/`** | **`FIX-07`**: `sap_docs/SYNTH_SAP_GUIDE.pdf`, `sap_docs/SYNTH_SCAN_NO_TEXT.pdf`, `sap_docs/SYNTH_SOP.docx`, `sap_docs/SYNTH_DATA_DICT.xlsx`, `sap_docs/SYNTH_CLOSURE_WITH_HOJA1_PII.xlsx`, `sap_docs/SYNTH_NOTES.txt`, `sap_docs/SYNTH_TCODE_REF.md`, `sap_docs/SYNTH_EXPORT_GUIDE.html`, `sap_docs/corrupt/empty.txt`, `sap_docs/corrupt/malformed.pdf`, `sap_docs/corrupt/corrupt.docx`, `sap_docs/corrupt/unsupported.bin` | All 6 formats (`PDF`, `DOCX`, `XLSX`, `TXT`, `MD`, `HTML`), page/heading/cell/line/section provenance markers, 0-text scan PDF, `Hoja1 MV 998` synthetic PII sentinel (`SYNTHETIC_PII_MV998_SENTINEL_DO_NOT_INDEX`), `VKP2` non-cost warning, `EINA`/`MAST`/`STPO` Table tags, `VL3IN` typo, `SAP GUI 7.00` caveat, `corrupt/` boundary files | `E2E-F09-T1/T2`, `E2E-F19-T1/T2`, `E2E-F20-T1/T2`, `E2E-F21-T1/T2`, `E2E-T3-11`, `E2E-T3-14`, `E2E-T3-19`, `E2E-T3-20`, `E2E-T3-22`, `E2E-T3-25`, `E2E-T3-27`, `E2E-T4-03`, `E2E-T4-11`, `E2E-T4-14`, `E2E-T4-16` |
| **Runtime Sentinel / RO Mirror** | **`FIX-08`**: `ProductionImmutabilityGuard` + optional RO mirror of host `data/` (`51`/`64` QC, `77`/`81` Closures, `Control_Diario_Operaciones_Jostin.xlsx`) | Pre/post SHA-256 & metadata snapshot of production paths; optional host-corpus parity tests with synthetic fallback when `data/` absent | `E2E-F10-T1-01..03`, `E2E-F10-T2-03`, `E2E-F15-T1-01..05`, `E2E-F15-T2-01..04`, `E2E-F22-T1-05`, `E2E-F22-T2-03`, `E2E-T3-06`, `E2E-T3-13`, `E2E-T4-10`, `E2E-T4-16` |

---

# Part 6 — Verification Criteria & Phase B xUnit Trait Mapping Table

## 1. Phase A Verification Criteria (Specification & Fixture Readiness)
1. **Count & Collision Verification:** Exactly **254** distinct test IDs (`105` Tier 1 + `105` Tier 2 + `28` Tier 3 + `16` Tier 4) across `21` operational features (`F02`–`F22`), matching `TEST_INFRA.md` 1-to-1 with `0` duplicate or missing IDs.
2. **Python Baseline Integrity:** `python -B -m pytest -p no:cacheprovider -v --tb=short` passes **58/58** tests (`0` failures, `0` skipped, `0` collection errors).
3. **Fixture Generator Determinism:** `python -B tests/e2e_fixtures/generate_fixtures.py --out "$env:TEMP\acr_e2e_fixtures" --verify` exits `0` with deterministic SHA-256 hashes across runs and `0` untracked binary files in `git status`.

## 2. Phase B xUnit Trait & Test Class Mapping (`src/dotnet/Acr.Tests/E2E/`)

Every Phase B xUnit test method MUST decorate with `[Trait("Category", "E2E")]`, `[Trait("Tier", "Tier1"|"Tier2"|"Tier3"|"Tier4")]`, and `[Trait("Feature", "F02".."F22")]` and execute under `E2eWorkspaceFixture` + `ProductionImmutabilityGuard` (`OperationalSafetyMode.Testing`).

| Feature / Tier | Planned Phase B Test File (`src/dotnet/Acr.Tests/E2E/`) | Required xUnit Traits | Test IDs Covered |
|---|---|---|---|
| **F02** | `Tier1_2/F02_SolutionArchitectureE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F02` | `E2E-F02-T1-01..05`, `E2E-F02-T2-01..05` |
| **F03** | `Tier1_2/F03_QuantityGrammarE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F03` | `E2E-F03-T1-01..05`, `E2E-F03-T2-01..05` |
| **F04** | `Tier1_2/F04_QuantitySpecialRowsAndCurvesE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F04` | `E2E-F04-T1-01..05`, `E2E-F04-T2-01..05` |
| **F05** | `Tier1_2/F05_DiscrepancyCardinalRuleE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F05` | `E2E-F05-T1-01..05`, `E2E-F05-T2-01..05` |
| **F06** | `Tier1_2/F06_RevueltoConfidenceE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F06` | `E2E-F06-T1-01..05`, `E2E-F06-T2-01..05` |
| **F07** | `Tier1_2/F07_DailyLogLifecycleDomainE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F07` | `E2E-F07-T1-01..05`, `E2E-F07-T2-01..05` |
| **F08** | `Tier1_2/F08_HistoricalIdentityDomainE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F08` | `E2E-F08-T1-01..05`, `E2E-F08-T2-01..05` |
| **F09** | `Tier1_2/F09_SapKnowledgeDomainE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F09` | `E2E-F09-T1-01..05`, `E2E-F09-T2-01..05` |
| **F10** | `Tier1_2/F10_OperationalSafetyAndAtomicIoE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F10` | `E2E-F10-T1-01..05`, `E2E-F10-T2-01..05` |
| **F11** | `Tier1_2/F11_QcReaderStreamingE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F11` | `E2E-F11-T1-01..05`, `E2E-F11-T2-01..05` |
| **F12** | `Tier1_2/F12_ClosureExporterAuditE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F12` | `E2E-F12-T1-01..05`, `E2E-F12-T2-01..05` |
| **F13** | `Tier1_2/F13_DailyLogRepoAndPipelineE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F13` | `E2E-F13-T1-01..05`, `E2E-F13-T2-01..05` |
| **F14** | `Tier1_2/F14_HistoricalClosureScannerE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F14` | `E2E-F14-T1-01..05`, `E2E-F14-T2-01..05` |
| **F15** | `Tier1_2/F15_RegressionParityGateE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F15` | `E2E-F15-T1-01..05`, `E2E-F15-T2-01..05` |
| **F16** | `Tier1_2/F16_BilingualLocalizationAndThemeE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F16` | `E2E-F16-T1-01..05`, `E2E-F16-T2-01..05` |
| **F17** | `Tier1_2/F17_ReconciliationViewE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F17` | `E2E-F17-T1-01..05`, `E2E-F17-T2-01..05` |
| **F18** | `Tier1_2/F18_DailyLogAndSyncViewsE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F18` | `E2E-F18-T1-01..05`, `E2E-F18-T2-01..05` |
| **F19** | `Tier1_2/F19_SapDocLibraryAndSearchE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F19` | `E2E-F19-T1-01..05`, `E2E-F19-T2-01..05` |
| **F20** | `Tier1_2/F20_SapEncyclopediaAndDiscoveryE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F20` | `E2E-F20-T1-01..05`, `E2E-F20-T2-01..05` |
| **F21** | `Tier1_2/F21_GuidedTroubleshootingAndHistoryE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F21` | `E2E-F21-T1-01..05`, `E2E-F21-T2-01..05` |
| **F22** | `Tier1_2/F22_OperationalSmokeAndPortableExeE2eTests.cs` | `Category=E2E`, `Tier=Tier1\|Tier2`, `Feature=F22` | `E2E-F22-T1-01..05`, `E2E-F22-T2-01..05` |
| **Tier 3** | `Tier3/PairwiseInteractionE2eTests.cs` | `Category=E2E`, `Tier=Tier3`, `Feature=<each_pair_feature>` | `E2E-T3-01`..`E2E-T3-28` |
| **Tier 4** | `Tier4/RealWorldWorkflowE2eTests.cs` | `Category=E2E`, `Tier=Tier4`, `Feature=<each_workflow_feature>` | `E2E-T4-01`..`E2E-T4-16` |
