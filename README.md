# 🛡️ ReconcileAI — 3-Way Inventory Reconciliation & Goods Receipt Auditor
### Enterprise Copilot for Supply Chain, Quality Control & Import Closures

[**Español**](README.es.md) | **English** | [**Operational Wiki**](docs/WIKI.md) | [**Release Notes (Demo v1.0.0)**](docs/RELEASE_NOTES_DEMO.md) | [**SAP Receiving SOP**](docs/SOP_RECEPCION_SAP.md)

---

## 🌟 Demo Release v1.0.0 (Standalone Portable Edition)

> [!TIP]
> **Zero Installation Required:** ReconcileAI is packaged as a **100% portable standalone distribution** (`ReconcileAI_Portable.zip`). It runs out of the box on Windows 10/11 **without requiring a Python installation or administrator privileges**.

An agnostic, enterprise-grade calculation engine and interactive audit copilot for **3-Way Supply Chain Inventory Reconciliation**. Designed to eliminate manual cross-referencing between Purchase Orders (PO), Commercial Invoices/Customs Declarations, Physical Warehouse Receiving counts, and ERP Material Movements (`Mov 101`).

```mermaid
flowchart LR
    A["ERP Purchasing (PO 43)"] --> D["ReconcileAI Copilot"]
    B["Commercial Invoice (Ped 41)"] --> D
    C["Physical Dock QC & Curves"] --> D
    
    D --> E["360° Material Inspector"]
    D --> F["Color-coded Semáforo Grid"]
    D --> G["Fast Paste Strips (Cols AS:AZ)"]
    D --> H["1-Click Official .xlsm Export"]
    
    G -->|"10-Second Copy/Paste"| I["Official Spreadsheet (.xlsm)"]
    H -->|"Instant Download"| I
    I --> J["100% Squared Zero-Discrepancy Closure"]
```

---

## 🎯 Executive Overview & Value Proposition

In international retail logistics and distribution centers, closing an import container typically requires reconciling data across multiple disconnected sources:
1. **Commercial / ERP Layer:** Purchase Orders (`PO 43`) and Customs Declaration Invoices (`Invoice 41`).
2. **Physical Receiving Layer:** Warehouse dock counts, carton/bulto tallies, and Quality Control size-curve breakdowns.
3. **Accounting / Ledger Layer:** ERP inventory movements (`SAP Mov 101`).

### The Industry Challenge
- **Granularity Mismatch:** Invoices and macros often operate at the **Parent Prepack (`PRK`)** level, while dock quality control counts items at the individual **Child SKU (`UN`)** level.
- **The Negative Formula Trap:** Default ERP spreadsheet macros often omit dock receiving and packing list tallies, causing artificial negative differences (e.g. `-51,926 UN`) that force operators to manually enter hundreds of cells line by line.
- **Human Math Fatigue:** Calculating complex size curves across dozens of styles by hand easily causes misalignments, erroneous credit note claims (NCR), or delayed nationalization closures.
- **Column Shift Traps:** In practice, operators frequently insert columns (e.g. `Tarima` in Column B) or alter sheet layouts, breaking traditional Excel `VLOOKUP` (`BUSCARV`) formulas.

### The ReconcileAI Solution
- **Deterministic Heuristic Cross-Referencing:** Groups child SKUs under their parent BOM using mathematical curve verification ($100\%$ zero-error curve validation).
- **Self-Healing Spreadsheet Recovery:** Detects deleted, overwritten, or corrupted formulas in official closure workbooks and automatically repairs them.
- **Dual Strip Copy-Pasting (`ZDESCARGA` & `PACKING LIST`):** Generates pre-formatted TSV text blocks for Columns `AS:AT` and `AW:AY` so operators can square the official corporate macro in **10 seconds** without modifying formulas.
- **1-Click Official `.xlsm` Export:** Automatically generates the finalized, fully-squared official Excel file in memory with all VBA macros intact.
- **360° Material Inspector (Zero-Scroll):** Instantly displays PO quantities, invoice data, physical counts, size explosions, and unit costs on a single screen.
- **SAP Pre-Macro Clipboard Validator:** Sanitizes raw data copied from SAP GUI (`ME80FN`, `MB51`, `SE16N`), eliminating scientific notation (`1.45E+06`) and column shifts before pasting.

---

## 🚀 Quickstart Guide

### Option A: Portable Standalone Executable (Recommended for Demo)
*Ideal for office or warehouse computers without Python or admin rights.*
1. Unpack **`ReconcileAI_Portable.zip`** (or use the `dist/ReconcileAI` folder).
2. Double-click **`ReconcileAI.exe`** (or `INICIAR_RECONCILEAI.bat`).
3. Your default web browser will open automatically at `http://localhost:8501`.

### Option B: Local Python Development
*For developers and local customization.*
```bash
# 1. Clone repository
git clone https://github.com/your-repo/3-way-inventory-reconciliation.git
cd 3-way-inventory-reconciliation

# 2. Run automated setup (creates virtual environment & installs dependencies)
preparar_equipo.bat

# 3. Launch interactive copilot
iniciar_asistente.bat
# or: streamlit run app.py
```

### Option C: 1-Second Terminal Audit (CLI Mode)
*For headless batch verification and terminal lovers.*
```bash
python cli.py -i IMP-182
# or specify direct paths:
python cli.py --qc path/to/QC.xlsm --cierre path/to/Cierre.xlsm
```

---

## 📋 2-Minute Demo Walkthrough

1. **Launch the App:** Open `http://localhost:8501`.
2. **Select Container:** In the file explorer, choose a local folder (e.g. `data/IMP-182` or your own container folder) or drag and drop your two Excel files:
   - File 1: Quality Control / Dock Receiving (`Control Calidad.xlsm`)
   - File 2: Official Closure Template (`Copia de IMP-XXX inprogress.xlsm`)
3. **Run Audit:** Click **`🚀 Iniciar Auditoría de estos Archivos`**.
4. **View KPIs & Semáforo:**
   - Review the **6-Metric KPI Summary** (Theoretical Bultos, Dock Cartons, Ped 41 Invoiced Units, Dock Received Units, Net Balance, Net Discrepancy).
   - Check the **🚦 Tabla Semáforo**: Green items are 100% matched; Blue items indicate split shipments (Parte B); Red items indicate actual physical shortages.
5. **Square the Closure:**
   - Go to **`📋 Copiar a Excel Oficial`**.
   - Either click **`⚡ Descargar Libro Oficial Cuadrado (.xlsm con Macros VBA)`** to download the completed file, or copy the pre-formatted text blocks for **`ZDESCARGA` (Cols AS:AT)** and **`PACKING LIST` (Cols AW:AY)** into your official workbook.
   - Result: Discrepancy column instantly drops to **`0`**.

---

## 📁 Project Architecture

```
3-way-inventory-reconciliation/
│
├── 🚀 iniciar_asistente.bat            # One-click Windows launcher (smart fallback)
├── ⚡ ver_resumen_rapido.bat            # One-click CLI quick auditor
├── 📦 preparar_equipo.bat               # 1-click venv & dependencies setup
├── 🖥️ app.py                           # Fullscreen Streamlit UI (Widescreen Fluid Layout)
├── ⚡ cli.py                            # Fast command-line reconciliation engine
├── 🏃 run_standalone.py                # Standalone Streamlit launcher for PyInstaller
├── 📖 README.md                         # English Documentation & Demo Guide
├── 📖 README.es.md                      # Spanish Documentation & Demo Guide
│
├── 📁 dist/                             # Compiled Standalone Deliverables (Git-ignored)
│   ├── ReconcileAI/                    # Standalone executable folder
│   │   ├── ReconcileAI.exe             # Native Windows executable
│   │   ├── INICIAR_RECONCILEAI.bat     # Friendly batch launcher
│   │   └── data/                       # Container files directory
│   └── ReconcileAI_Portable.zip        # Portable standalone distribution archive
│
├── 📁 src/                              # Core Agnostic Reconciliation Engine
│   ├── __init__.py
│   ├── config.py                       # Dynamic path discovery & PyInstaller environment resolution
│   ├── models.py                       # Dataclasses: PrepackItem, ChildSKU, SummaryStats
│   ├── sanitizer.py                    # Defensive data cleaning (scientific notation, strings)
│   ├── validator.py                    # Business rules, curve sum checks & discrepancy logic
│   ├── parsers.py                      # Dynamic header scanner & multi-sheet Excel parsers
│   └── engine.py                       # 3-Way matching logic, formula recovery & NCR calculation
│
├── 📁 docs/                             # Operational Manuals & Reference Guides
│   ├── WIKI.md                         # Comprehensive Starting Guide & Demo Tutorial
│   ├── RELEASE_NOTES_DEMO.md           # Demo v1.0.0 Release Notes
│   └── SOP_RECEPCION_SAP.md            # Standard Operating Procedure (SAP Goods Receipt)
│
├── 📁 tools/                            # Build & Packaging Utilities
│   └── build_exe.py                    # Automated PyInstaller onedir builder & packager
│
└── 📁 tests/                            # Automated Regression Test Suite
    ├── __init__.py
    ├── test_sanitizer.py               # Unit tests for text & numeric sanitizer
    └── test_parsers.py                 # Automated tests for Excel extraction & curve math
```

---

## 🔒 Security, Privacy & Confidentiality Guarantee

- **100% Agnostic Architecture:** Designed to work universally with any vendor, any ERP (SAP, Oracle, NetSuite), and any warehouse system.
- **Zero Company Branding or Sensitive Data:** No proprietary corporate credentials, facility codes, or confidential supplier data are baked into the codebase.
- **Local Workstation Memory:** All parsing and mathematical calculations run in temporary local memory (RAM). No data is transmitted over the Internet or persisted to third-party servers.

---

## 📄 License & Attribution

Developed as an independent, enterprise-grade supply chain automation and reconciliation copilot. Free for operational, professional, and educational use under the MIT License.
