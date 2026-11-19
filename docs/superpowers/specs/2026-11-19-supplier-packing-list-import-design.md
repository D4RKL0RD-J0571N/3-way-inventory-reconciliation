# Supplier Packing List Import and QC Comparison

**Status:** Draft for user review
**Target:** Active C# WPF desktop edition (`src/dotnet`)

## Goal

Read an optional supplier packing-list document and use its declared item quantities and carton counts to populate the existing Packing List values in the reconciliation workflow. Make supplier values comparable with QC so an operator can investigate whether the QC transcription agrees with the provider document.

The supplier list is optional. Existing blind receiving must continue to work when both QC and supplier packing list are absent. The current closure workbook remains the normal receiving input/template.

## Current behavior

`ReceivingPipeline` reads declared items from the closure workbook and optionally reads QC. If QC is absent, it creates an empty QC report with a blind setup warning and continues using closure declared values. The workbook reader obtains declared carton values from AW and declared units from the relevant closure fields. The current domain pipeline does not read supplier PDF, image, or standalone packing-list workbooks.

## Proposed workflow

1. Add an optional supplier packing-list path to the desktop file loader. It accepts PDF, Excel (`.xlsx`, `.xlsm`, `.xls`), and common image formats (PNG, JPG/JPEG, TIFF). The path may also be left empty.
2. Parse the supplier document into normalized item rows: supplier material/reference, description, declared units, declared cartons, and source location. Keep original text and parsed values available for review.
3. Match supplier rows to closure items by normalized SAP material code or the closure's `MAT. PROVEEDOR` reference. A unique supplier-reference match resolves to its SAP material automatically; ambiguous references show the candidate SAP materials for operator selection. When neither code matches, show description-based candidates. The closure item set is authoritative: a supplier-only material can be inspected but cannot be added to reconciliation/export by this feature. Closure items absent from the supplier document retain their existing closure values.
4. Aggregate repeated supplier rows for a material by summing each present numeric field (units and cartons independently). Show the contributing source rows and aggregate in review. A field absent from the document is not inferred; accepting only units preserves the existing closure carton value, and accepting only cartons preserves existing units.
5. Present a review table with supplier values next to closure Packing List values and QC values when QC is loaded. Show match status and differences in units/cartons. The operator can select a closure material for a candidate/unmatched row, choose which available fields to accept, or skip the row. No unresolved, skipped, or rejected field changes output. Require explicit confirmation of the reviewed import before accepted values populate the current run's existing Packing List fields and downstream Bultos/closure projections.
6. Preserve per-field provenance in the in-memory reconciliation result: supplier filename plus page/sheet and row/cell location, accepted value, and source kind (supplier import or closure workbook). Do not overwrite the source closure workbook merely by loading or accepting a file.

## Format handling

- **Excel:** Read modern `.xlsx` and `.xlsm` workbooks and legacy binary `.xls` files. Detect likely header rows and map a documented initial set of English/Spanish header aliases for material, description, units, and cartons. Unrecognized layouts should produce a clear review/import error rather than guessed values. Password-protected workbooks are not supported; report that the user must provide an unprotected copy.
- **Text PDF:** Extract selectable text locally and parse common table rows. The sample `EN26E01002-75 PL.pdf` has selectable text, but its extracted row ordering is not fully tabular; parser confidence and review are therefore required.
- **Scanned PDF and images:** Run local OCR for PNG, JPG/JPEG, TIF, and TIFF inputs and scanned PDF pages, then send candidates through the same review workflow. OCR must be an optional capability with a clear message when unavailable or unable to read the document. It must not be necessary for blind setup or for reconciliation without a supplier document. OCR-derived values with confidence below 0.85 are flagged low-confidence and are not preselected for acceptance; all OCR-derived fields require the same explicit operator confirmation as other extracted values.
- Encrypted/password-protected PDFs are not supported; report a clear request for an unprotected copy. No password prompt or remote parsing service is in scope.

## Reconciliation semantics

- The supplier document is the intended source for populated Packing List values when the operator accepts the import.
- QC remains the physical/control source. If QC is loaded, show supplier-vs-QC quantities and cartons as an inspection comparison; do not replace QC values with supplier values.
- If no supplier list is loaded or accepted, retain the current closure workbook values and current blind-receiving behavior. If a supplier file omits a closure material or a field, retain that item's/field's closure value.
- If QC is absent, supplier import may still populate declared Packing List values, but the run remains blind and must retain the existing warning/paused closure semantics.
- Existing material discrepancy rules remain unchanged; supplier-vs-QC differences are surfaced separately and do not algebraically offset shortage/overage results.

## Architecture and scope

- Add a supplier document reader in `Acr.Infrastructure`, separate from `Acr.Core` domain logic and the existing QC/closure workbook readers.
- Add normalized supplier-row and match/review result types at the infrastructure boundary; pass only explicitly accepted, field-specific values into pipeline enrichment. Supplier-only rows remain outside the declared item set.
- Extend the `ReceivingPipelineOptions` / result surface to carry the optional document and accepted rows without making the supplier file required.
- Add an optional picker and review step to the WPF loader/view model. Keep QC optional and leave the existing closure-template requirement/workflow intact.
- Integrate accepted values into the existing Packing List projection (including `BultosItems` and AW:AY output) and show QC comparison in the review table when available. Values affect the active run and the existing explicit export action; loading/reviewing does not write any workbook, and exports continue to use the user-selected destination and macro-safe exporter.
- This change targets the active desktop app. Legacy Streamlit and browser-only web editions are out of scope for this iteration.

## Failure handling

- Invalid or unsupported files produce an actionable error and leave the reconciliation run unchanged.
- No reliable rows, ambiguous material matches, and OCR uncertainty are visible in the review step; do not silently fill values from uncertain rows.
- A missing OCR engine affects only image/scanned-document import. Normal Excel, text-PDF, QC, and blind workflows remain usable.
- Missing QC and missing supplier files do not block blind receiving.

## Acceptance criteria

1. With only a closure workbook, the existing workflow and outputs remain unchanged.
2. With a closure workbook and accepted supplier Excel/PDF data, matching supplier cartons/units populate the existing Packing List fields and appear in the current run and explicitly requested export.
3. Repeated supplier rows aggregate units and cartons independently; missing fields preserve closure values; supplier-only materials never enter reconciliation/export.
4. With QC loaded, the review table displays supplier and QC values and their differences without changing QC physical values.
5. Ambiguous rows can be matched manually to a closure material or skipped; unmatched/skipped fields cannot be applied without an explicit selected match and field acceptance.
6. With the closure workbook but no QC and no supplier list, blind setup still runs and displays the existing blind warning.
7. With an unreadable scanned PDF/image, OCR confidence below 0.85, unavailable OCR support, encrypted PDF, or password-protected workbook, the user receives a clear message; uncertain/unreadable values are not preselected or applied.
8. The source closure workbook is not modified during import or review; generated output is written only through the existing explicit export action to its selected destination.

## Validation plan

Use representative `.xls`, `.xlsx`, selectable-text PDF, scanned PDF/image, encrypted/password-protected files, malformed input, repeated supplier rows, units-only/cartons-only rows, matching/unmatched/ambiguous material cases, QC-present and QC-absent runs, and a regression check for the no-QC/no-packing-list blind path. Verify imported values through Bultos and an explicitly requested closure export. Preserve existing blind receiving and macro-safe export behavior.

## Open implementation detail

Select and package a local OCR engine and the PDF/legacy Excel readers after checking the desktop deployment constraints. OCR must not require a remote service or network connection.
