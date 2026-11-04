# 🚀 ReconcileAI — Release Notes: Standalone Edition (v1.2.0)
### Universal 3-Way Inventory Reconciliation & Goods Receipt Auditor

---

## 📌 Executive Summary

**ReconcileAI v1.2.0** brings significant operational hardening, sub-second streaming performance, and refined carton mathematics based on direct real-world benchmark stress-testing.

Key advancements in this version:
1. **Zero-Carton Precision Rule (`0.0` Bultos for Zero Units):** Eliminates legacy phantom 1-carton defaults for split containers (`Parte B`, `Parte C`) and unreceived lines.
2. **Dual-Mode Carton Reconciliation (`Conteo de Muelle` vs `Estricto Aduanero`):** Instant switch between physical dock count grouping and mathematical DUA carton/odd decomposition.
3. **50x Streaming Loader Optimization (<0.3s):** Selective `read_only=True` workbook streaming that loads 4.5 MB+ workbooks in ~0.28 seconds instead of 16+ seconds.
4. **Recursive File Discovery:** Automatically locates Quality Control files nested inside subdirectories (`INFORMACION GENERAL/`, `Facturas/`, etc.).
5. **1-Click Clipboard Strips & Hotkey Protection:** One-click copy buttons for `Tira 1 (AS13)` and `Tira 2 (AW13)` with built-in suppression of accidental modal shortcuts.
6. **Blind Setup Mode (`PAUSA_SIN_CONTROL_CALIDAD`):** Safely mounts initial SAP closure macros while placing the import on pause until Quality Control physically confirms receiving counts.

This release delivers an updated **100% portable standalone executable** (`dist/ReconcileAI_Portable.zip`, 134 MB), requiring **zero Python installation, zero external dependencies, and zero administrative rights**.

---

## ✨ Key Capabilities & Highlights

### 1. 🛡️ Zero-Carton Precision Rule (Level 1 Hardening)
- **Elimination of Phantom Defaults:** When items belong to deferred parts (e.g., `PENDIENTE_PARTE_B`) or have zero physical receipt (`received_real_units == 0`), Column `AS` (Bultos Muelle), Column `AT` (Unidades Muelle), Column `AW` (Bultos PL), Column `AX` (Unidades PL), and the official `'Bultos'` sheet are set strictly to **`0.0`**.
- **100% Exact Match with Human Reference:** Validated against real historical closed imports (e.g., `IMP-81.xlsm`), achieving **0 mismatches out of 52 items** across all columns.

### 2. ⚖️ Dual Carton View Modes (Dock vs. Strict Customs)
- **Modo Conteo de Muelle (Consolidado Humano):** Direct physical carton totals matching dock tally sheets and patio receiving tags.
- **Modo Estricto Aduanero (Enteros + Picos):** Rigorous mathematical decomposition into full standard boxes and odds (`Enteros + Picos`) aligned with commercial curves and DUA declarations.

### 3. ⚡ Ultra-Fast Streaming Excel Engine (<0.3s vs 16.1s)
- **Selective Sheet Streaming:** Reads heavy Quality Control files (3.5 MB – 4.5 MB) using streaming row iterators (`iter_rows(values_only=True)`), bypassing drawing and image XML layers.
- **50x Speedup:** Cuts file processing latency from 16.1 seconds down to **0.28 seconds**.

### 4. 📁 Recursive Subfolder Discovery
- Scans target import directories recursively (`p.rglob("*")`). Automatically pairs nested Quality Control workbooks located inside subfolders like `INFORMACION GENERAL/` or `Facturas/` without requiring manual file reorganization.

### 5. 📋 1-Click Clipboard Copying & Modal Hotkey Protection
- **Direct 1-Click Copy Buttons:** `📋 Copiar Tira 1 (AS13)` and `📋 Copiar Tira 2 (AW13)` copy tab-delimited blocks straight to the Windows clipboard with instant visual confirmation (`✅ ¡Copiado!`).
- **Global Hotkey Suppression:** Intercepts Streamlit's default window-level `c` hotkey in capture phase, preventing the intrusive *"Clear caches"* dialog from interrupting copy operations.

### 6. ⏸️ Blind SAP Setup Pause Mode (`PAUSA_SIN_CONTROL_CALIDAD`)
- When a container is awaiting physical unloading or lacks a Quality Control workbook, ReconcileAI enables preliminary SAP mounting while enforcing an explicit pause banner. It disallows certified closure generation until dock receiving counts are supplied.

---

### 1. 🛡️ Portable Standalone Packaging (Zero Configuration)
- **Zero-Install Deployment:** Distributed as a standalone bundle `ReconcileAI_Portable.zip`. Simply unpack and double-click `ReconcileAI.exe` or `INICIAR_RECONCILEAI.bat`.
- **Sub-Second Boot:** Built using PyInstaller's optimized `--onedir` architecture, avoiding `%TEMP%` extraction delays and Windows Defender SmartScreen bottlenecks.
- **Offline & Autonomous:** Bundles an embedded Python runtime, Streamlit web engine, Pandas, and OpenPyXL. Operates 100% offline without accessing external cloud servers.

### 2. 📦 Carton Breakdown & Compound Packaging Engine (Level 1 Segregation)
- **Mathematical Segregation:** Distinctly tracks Level 1 (Physical Cartons), Level 2 (Commercial Prepacks/PRK), and Level 3 (Individual Units/Prendas).
- **Compound Packaging String Parser:** Parses mixed carton allocations such as `'39*60 1*78'` into 40 cartons and 2,418 units without manual calculator conversions.
- **Dedicated Carton Tab (`📦 Cuadre & Desglose de Bultos`):** Top 5 KPI cards for cartons, interactive packaging breakdown calculator, and 1-to-1 digital replica of the official `'Bultos'` sheet.

### 3. 🔄 Dual-Phase Architecture (Reception & Dispatch Readiness)
- **Reception Closure (`CIERRE RECEP`):** Reconciles customs declarations, dock QC counts, and `Mov 101` receipts.
- **Dispatch Closure (`CIERRE DESP`):** Pre-wires and synchronizes material lines, curves, and `Mov 641` store transfers via Excel formulas, saving 80% of repetitive operational effort for future dispatch audits.

### 4. ⚡ Heuristic 3-Way Cross-Referencing Engine
- **Parent Prepack (`PRK`) to Child SKU (`UN`) Decomposition:** Reconciles bulk commercial lines against individual size-curve tallies with 100% mathematical precision.
- **Dynamic Header & Column-Shift Tolerance:** Automatically scans headers dynamically (`Material`, `Descripción`, `Bultos`, `Unidades`, `Curva`). Resists manual Excel alterations (such as operators inserting new columns like `Tarima` in Column B).
- **Universal Material Resolution:** Cross-references SAP codes, vendor style numbers, and internal BOM IDs, avoiding duplicate BOM ID traps.

### 5. 🩹 Self-Healing Spreadsheet Recovery
- **Restoration of Corrupted Formulas:** Automatically restores deleted or overwritten formulas in official closure sheets (e.g. Column H `G-J-N-O+L+P+Q`, Column AU `AT-H`, Column AZ `AX-H`).
- **Elimination of the "Negative Formula Trap":** Default ERP templates leave receiving counts empty, displaying artificial shortages (e.g. massive negative differences). ReconcileAI immediately calculates true receiving values to bring discrepancies to zero.
- **1-Click Official `.xlsm` Export:** Generates an official, fully-squared closure workbook in memory while strictly preserving all VBA macros and original formatting.

### 6. 📋 Dual-Strip Fast Copy-Pasting (10-Second Manual Closure)
- **Pre-Formatted Tab-Delimited Blocks:** Generates ready-to-paste text blocks for Column `AS:AT` (`ZDESCARGA`) and Column `AW:AY` (`PACKING LIST`).
- **Zero ERP Dependency for Closing:** Once initial SAP data is pasted into the workbook, operators can complete the entire closure without reopening SAP GUI.

### 7. 🔍 360° Material Inspector & Visual Audit Grid
- **Color-Coded Status Semáforo:**
  - 🟢 **Verde (Cuadrado):** 100% physically received, zero discrepancy.
  - 🔵 **Azul (Parte B / En Tránsito):** Container split or multi-lot cargo.
  - 🔴 **Rojo (Discrepancia / NCR):** Genuine physical shortage or transit damage requiring Credit Note claim.
- **Zero-Scroll Ficha 360°:** Instant drilldown into any SAP code or vendor style, showing purchasing units, invoiced quantities, dock receipts, bultos, and full size-curve breakdowns.

### 8. 🔒 Enterprise Privacy & Agnostic Compliance
- **Zero Hardcoded Corporate Data:** No proprietary company names, facility codes, or confidential vendor references exist in code, logs, or documentation.
- **Local Memory Guarantee:** All data operations occur exclusively in local workstation RAM. No telemetry, no external API calls, and zero external persistence.

### 9. 🌐 Bilingual Engine & Dynamic Table Tooltip Localization
- **Spanish as Default (`es`):** All headers, KPI cards, table columns, checklist items, and interactive wiki walkthroughs default to Spanish, with instant toggle to English (`en`) via sidebar selector.
- **Client-Side Table Toolbar & Tooltip Translation:** Client-side DOM mutation engine intercepts Streamlit / BaseWeb table toolbars in real time, translating hover tooltips (e.g. `Fullscreen` ➔ `Pantalla completa`, `Download as CSV` ➔ `Descargar como CSV`, `Search` ➔ `Buscar`) to ensure an entirely cohesive, localized user experience.

---

## 📂 Deliverables in this Release

| Component | Path / File | Purpose |
| :--- | :--- | :--- |
| **Portable Bundle (ZIP)** | `dist/ReconcileAI_Portable.zip` | Single-file archive for offline distribution and transfer |
| **Standalone Folder** | `dist/ReconcileAI/` | Unzipped portable application ready to run on any PC |
| **Main Executable** | `dist/ReconcileAI/ReconcileAI.exe` | Compiled standalone Windows binary |
| **Quick Launcher** | `dist/ReconcileAI/INICIAR_RECONCILEAI.bat` | Friendly Windows batch launcher |
| **Data Directory** | `dist/ReconcileAI/data/` | Drop folder for container folders (e.g. `CONT-001`) |
| **Quickstart Guide** | `dist/ReconcileAI/LEEME_PORTABLE.txt` | On-disk instructions for end-users |

---

## 🚀 2-Minute Demo Evaluation Guide

1. **Extract Archive:**
   - Unpack `ReconcileAI_Portable.zip` to your Desktop or a USB drive.
2. **Launch:**
   - Double-click `ReconcileAI.exe`.
   - Your default web browser will automatically open at `http://localhost:8501`.
3. **Load Container Data:**
   - Either drag and drop your Quality Control and Closure spreadsheets into the **Subir Archivos** tab, or select a folder in the **Explorador de Disco**.
   - Click **`🚀 Iniciar Auditoría`**.
4. **Evaluate Results:**
   - Observe the **6-Metric KPI Summary Bar** across the top.
   - Inspect items in the **🚦 Tabla de Conciliación Semáforo**.
   - Navigate to **`📋 Copiar a Excel Oficial`** and click **`⚡ Descargar Libro Oficial Cuadrado (.xlsm)`** or copy the **Tiras de Pegado Rápido**.

---

## 🖥️ System Requirements

- **Operating System:** Windows 10 (64-bit) or Windows 11 (64-bit).
- **RAM:** 4 GB minimum (8 GB recommended for large multi-container workbooks).
- **Disk Space:** ~350 MB uncompressed.
- **Web Browser:** Google Chrome, Microsoft Edge, Mozilla Firefox, or any modern Chromium browser.
- **Software Dependencies:** **None.** (Python, openpyxl, pandas, and Streamlit are fully self-contained).

---

## 📈 Roadmap & Upcoming Enhancements

- [ ] Direct export of official Credit Note (NCR) claim PDFs for vendor billing dispute.
- [ ] Automated OCR scanner for scanned customs DUA documents.
- [ ] Real-time multi-user network sharing mode for warehouse floor terminals.
