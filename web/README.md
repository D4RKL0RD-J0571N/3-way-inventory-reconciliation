# 🛡️ ACR — ReconcileAI (3-Way Inventory Reconciliation Web App)

Universal 3-Way Supply Chain Inventory Reconciliation & Goods Receipt Closure Assistant built with **React, TypeScript, Vite, Tailwind CSS, and SheetJS** for Node.js 22.

## Key Capabilities

- **In-Browser Excel Processing (`.xlsm` / `.xlsx`):** Dynamic header detection on `CONTROL DATA` (Row 6) and `CIERRE RECEP` (Row 12), BOM curve extraction from `BASE DE DATOS`, and multi-part container detection (`PARTE A`, `PARTE B`, etc.).
- **Ready-to-Audit Sample Containers:** Includes canonical clean closure (`IMP-901`), multi-part discrepancy & NCR claim (`IMP-909`), and Blind SAP Setup / Pause Mode without QC (`IMP-184`).
- **Executive KPI Strip & 8 Operational Modules:**
  1. **Carton Reconciliation & Packaging (`Bultos`):** Dock Count vs. Strict Customs modes, interactive Carton & Packaging String calculator (`39*60 1*78`), and `.xlsx` export.
  2. **360° Material Inspector:** 4-card breakdown (Identification, PO 43, Invoice 41 & Dock, Balances & NCR), child SKU sizes, and integrity alerts.
  3. **Status Semáforo Matrix:** Color-coded 3-way reconciliation grid with dynamic filters.
  4. **Prepack Curve Breakdown:** BOM size explosion per parent style.
  5. **SAP Clipboard Validator (Pre-Macro):** Sanitizes scientific notation, leading zero padding, and European/US numeric formats from raw SAP GUI exports.
  6. **Official Excel Export & Paste Strips:** 1-click TSV copy strips for `ZDESCARGA` (`AS:AT`) and `PACKING LIST` (`AW:AY`), plus multi-sheet `.xlsx` download.
  7. **Physical Binder Checklist:** Interactive 7-step printing checklist with live progress tracking.
  8. **SAP T-Code Encyclopedia & Diagnostic Guides:** Offline reference for `ME80FN`, `SE16N`, `MB51`, `VL31N`, `RFUI`, `MIGO`, `EINA`, `MAST`, `STPO`, and step-by-step troubleshooting guides.

## Running Locally

```bash
npm install
npm run dev
```

The development server runs on `http://0.0.0.0:3000`.
