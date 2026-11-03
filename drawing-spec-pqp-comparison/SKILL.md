---
name: drawing-spec-pqp-comparison
description: >-
  Audits and compares a Product Quality Plan (PQP) Excel sheet against one or more mechanical drawing PDFs
  (Top-Level Assembly, Interior/Exterior sub-assemblies). Validates torques, metadata, dimensions,
  Gage R&R, and process capability status.
---

# Drawing Specification & PQP Comparison Skill

## Overview
This skill provides automated checks and audits to compare a mechanical assembly's Product Quality Plan (PQP) workbook against its engineering drawing PDFs (e.g. ASSA ABLOY/HHI top-level assemblies, interior/exterior subassemblies, setup code labels, and packaging).

Specifically, it checks:
1. **Metadata Alignment**: Verifies Part Number, Description, Drawing Number, and Revision match across all visible sheets and align with drawings.
2. **Multi-Part & Family Part Support**: Intelligently handles family part submittals (e.g., multiline part numbers `58740-001\n58741-001`) and matches individual child dimensional sheets (`100% Dimensional_58740-001`, `100% Dimensional_58741-001`) without false-positive mismatch warnings.
3. **Control Plan Torques & Reaction Plan Statistical Control**: Checks if the torques specified in assembly steps match drawing specifications (mapped via flag notes). Audits **Control Plan Reaction Plan Robustness for Sampling Plans**: on periodic sampling plans (e.g. 5 pcs / 2 hrs), discarding/scrapping only the failed sample units (`Discard` / `Scrap`) is rejected as statistically non-robust. Mandates that the Reaction Plan must require containment and sorting of all suspect units produced since the last acceptable inspection (*"At minimum, should be sorting units that were produced since the last inspection"*), quarantine of non-conforming parts for scrap/rework, and process adjustment with re-verification before restart. Flags blank reaction plans on sampling rows as critical failures.
4. **MSA & Process Capability**: Audits Gage R&R and capability sheets to make sure they are populated and free of errors like `#DIV/0!`.
5. **Dimensional Drift Warnings**: Flags border-line measurements (disposition is `"Investigate"`) in dimensional inspection sheets.
6. **ASME Y14.5 No-Leading-Zero Inch Dimensions & Dual Units**: Robustly parses inch dimensions omitting leading zeros (e.g. `.43`, `.75`, `R.032`, `1.850`) and space-separated tolerance notes (`.75 .02`, `1.850 .020`), pairing them with metric dual callouts (`[11]`, `[19,1 ±0,5]`, `[0,8]`, `[46,99 ±0,5]`).
7. **Auto OCR Fallback**: If the engineering drawing PDF contains no extractable text objects (e.g., text was converted to geometry during CAD export), the tool automatically renders page images at 300 DPI and runs OCR using `pymupdf` and `rapidocr-onnxruntime` to extract dimensions and notes.
8. **Enhanced Fraction & Unit Checking**: Supports parsing complex fractional values (e.g., `8-5/8` to `8.625`, `4-7/8` to `4.875`) from PDF drawing text and cleans trailing character artifacts. Corrects unit auditing logic to respect explicit `INCH` designations.
9. **Robust Angle Auditing & CAD Adjacent Text Isolation**: Parses general angle tolerances (e.g., `ANGLES: +/- 1`) and matches angular dimensions (e.g., `45.0°`) in Excel sheet to drawing callouts. Crucially isolates angle dimensions from adjacent CAD drawing text (e.g. `8$` followed by `2X .060`, or `4$` followed by `.026\n4X`) to prevent erroneously consuming adjacent counts/dimensions as angle tolerances. Strictly separates angle callouts from linear dimensions in comparison matching loops.
10. **Limit & Reference Dimensions**: Auto-extracts suffixes (e.g., `MIN`, `MAX`, `REF`) and bypasses tolerance audits for limits as they do not have standard nominal tolerances.
11. **Adaptive 3F / FFF Photo & Content Verification**: Audits structured table cells in `FFF`/`3F` sheets (Verification: Form/Fit/Function, Requirement / Test Performed, Qty Tested, Disposition). Performs **Project Name Copy-Paste Residue Detection** (e.g. catches foreign project names like `FOSTER` when drawing is `TITAN`), validates sample size alignment between `Qty Tested` (e.g. 5 pcs) and 100% Dimensional measured sample count (e.g. 3 pcs), and alerts on overly generic descriptions (`comply with the drawing`). Dynamically adapts photo OCR verification to label type (Customer Support QR labels, Setup Code labels, Brand logos like Kwikset/Weiser, model descriptions, and bilingual English/French text) preventing false-positive missing FCC/IC warnings on non-regulatory items.
12. **Embedded Raw FAI Sub-Workbook Discrepancy Auditing, Multi-Layer Drawing Architecture & Copy-Paste Shift Detection**: Automatically inspects embedded raw FAI child workbooks (`xl/embeddings/Microsoft_Excel_Worksheet*.xlsx`) and cross-verifies sample measurements against the active visible PQP `100% Dimensional_*` worksheets, immediately detecting altered data or Out-of-Spec (OOS) conditions.
    - **Non-Numeric / Visual FAIL & Contradictory Pass Detection**: Parses text measurements (`FAIL`, `NG`, `REJECT`, `不合格`) on visual/note callouts (e.g. Balloon 37 `Paint Masking Area` on matte black finish 514). Flags fraudulent or erroneous reporting where measurement column states `FAIL` but result column marks `O` (OK) or when suppliers omit the item from the outer sheet.
    - **Outer PQP Table Truncation & Balloon Omission Audit**: Cross-compares the full balloon set of the embedded raw FAI sub-workbook against the outer visible PQP dimensional sheet. Catches critical vendor omissions where inspection tables are truncated (e.g. outer table stops at Balloon 31, omitting Balloons 32~37 covering material, appearance, raised characters, markings, brush stroke, and paint masking).
    - **Sibling Embedded Workbook Duplication / Cavity Mismatch Audit**: Scans sibling embedded workbooks attached to different cavity sheets (e.g. Cavity 1 vs Cavity 2) and flags 100% duplicate embedded files where suppliers copy-pasted Cavity 1 without embedding the true Cavity 2 raw report.
    - **Accurate Relationship-Based Host Worksheet Pairing**: Resolves exact host worksheet relationships from `xl/worksheets/_rels/` and `xl/drawings/_rels/` to pair embedded files directly with their host sheets, eliminating spurious cross-sheet mismatch noise.
    - **Copy-Paste Shift Detection across Family Parts**: Automatically scans sibling part sheets in the raw FAI and identifies if supplier mistakenly copied measurements from an adjacent part number (e.g. `50013-001` pasted into `50012-001`).
    - **Multi-Layer / Sub-Assembly Drawing Architecture**: Distinguishes Top-Level Assembly drawings (finished assembly dimensions, total thickness `0.85 ± 0.05 mm`, finished aperture `4 ± 0.2` & `3 ± 0.2 mm`) from Layer Component Drawings (Overlay, Adhesive, Shading paper, PC, Metal dome with single layer thickness `0.2 ± 0.02 mm`, aperture `6 mm`, `5 mm`). Audits multi-section sub-workbooks (300+ rows covering Section 0 Assembly + Sections 1~10 raw material layers) without generating false-positive missing balloon alarms on the top-level assembly drawing. Includes deduplication across multiple embedded file copies.
13. **Embedded Test Certificates, Product Name Cross-Verification, SGS/CTI Lab Tracking Filter & Material Cert Recognition**: Parses embedded OLE binary attachments (`oleObject*.bin`) across `Material Certs`, `Packaging Test Reports`, and `Finish Test Reports`, identifying and summarizing attached test certificates (SGS RoHS 2.0, CTI REACH SVHC, ISTA 3A Drop Test, Material Specifications) along with their host worksheet names (Sheet Name). Accurately classifies engineering drawing PDFs by prioritizing title blocks (`DO NOT SCALE`, `THIRD ANGLE PROJECTION`, `DWG NO`) over packaging notes. Crucially performs **Product Name / Description Cross-Check** between embedded Material Specification / Certificate of Compliance (CoC) and Drawing Title / PQP Part Description, immediately catching template reuse blunders where suppliers submit carton/packaging CoCs for divider/bracket parts. Filters table header tokens (`SN ID`, `SAMPLE ID`, `ITEM NO`) and prioritizes `Sample Name:` / `样品名称：`. Recognizes raw material certificates embedded in `[Material Certs]` (e.g. `White PET adhesive tape`, `聚碳酸酯薄膜及片材`, `3M SP7533`) as compliant component certifications rather than flagging finished assembly description mismatches.
14. **Sample Count (< 5 pcs) & Sequence Completeness Audit**: Audits 100% Dimensional Inspection worksheets to verify `Sample#` entries are non-empty, start from 1, and follow a contiguous integer sequence (`1, 2, 3... N`) across all balloon items. Crucially alerts if suppliers only measured 3 samples instead of ASSA ABLOY's standard 5-piece sample requirement.
15. **Multi-Drawing Part Number Consistency Audit**: Scans across all provided engineering drawing PDF files to detect any conflicting or inconsistent Part Numbers/Drawing Numbers and verifies alignment with PQP header specifications.
16. **Control Plan CTQ / SC Coverage, Traceability, Reaction Plan Statistical Control & Non-Dimensional SC Guidance**: Automatically extracts all Significant Characteristics (SC) and Critical To Quality (CTQ) dimensions from drawings and validates coverage in Control Plan. Crucially enforces that all sampling inspection reaction plans specify containment and sorting of suspect units produced since the last inspection (*"At minimum, should be sorting units that were produced since the last inspection"*), rather than merely discarding failed sample pieces. from drawings (including vector diamond `◇` flag notes, e.g. material specifications and perforated die-cut styles mapped to Note 10 SC definition), 100% Dimensional sheets, Variable R&R, and Process Capability, and cross-audits their coverage in the Control Plan. Flags missing SC classifications in Control Plan operations, generic specification descriptions, and cross-checks evaluation tools. Intelligently provides **Non-Dimensional SC Guidance**: when suppliers declare SC on visual/material checks (e.g. print artwork, sticker, material) without drawing diamonds, advises removing SC or substituting Attribute MSA rather than triggering false-positive Variable R&R failures.
17. **OpenPyXL Drawing Shape & Font Descriptor Resilience**: Automatically patches OpenPyXL drawing font descriptors (e.g. `pitchFamily.max`) to ensure reliable loading of workbooks containing complex CAD shapes without XML parsing crashes (`ValueError: Max value is 52`).
18. **Unrelated Embedded Files Audit (Section 8)**: Scans all embedded objects in the PQP workbook (`xl/embeddings/`) — including OLE `.bin` PDF attachments, embedded XLSX sub-workbooks, and raw binaries — and cross-checks every detected part number against the PQP's own known part numbers, while displaying their exact host worksheet names (Sheet Name). Supports 5-digit dash 3-digit family part numbers (`\b\d{5}-\d{3}\b`, e.g. `69633-002`) and base part numbers. Flags any embedded file whose part number does not match the current product, helping to detect copy-paste artifacts from other PQP templates. Reports three categories: ✅ Relevant, ❌ Unrelated, and ❓ No PN detected (needs manual review).
19. **ASSA ABLOY Multi-Release Title Block & Grid Zone Revision Parser**: Robustly parses multi-release title blocks (e.g., `Production 2 5058313`, `2 PRODUCTION RELEASED`) while filtering out drawing frame grid zone letters (`A`, `B`, `C`, `D` after `REV`). Expands Excel header search boundaries up to column 50 with offset 15 to reliably extract Revision fields located in non-standard columns (e.g. columns 27/37 of PQS/Index sheets).
20. **Embedded FAI Sub-Workbook Deep Auditing (Up to 10MB)**: Expands embedded Excel FAI inspection capacity to 10MB to support high-resolution drawing sheets (`image*.jpeg`). Inspects embedded sub-workbooks (such as `TLM FAI REPORT`) for formula errors (`#VALUE!`, `#REF!`, `#DIV/0!`, `#NAME?`), missing or blank header Revisions, and verifies CTQ/SC designation consistency between the main PQP and embedded FAI sheets.
21. **CTQ vs. SC Discrimination & Downgrade Prevention**: Distinguishes perpendicular cross-hair diamonds (`⯐`, Note 7 CTQ symbol) from pure empty diamonds (`◇`, Note 8 SC symbol). Rejects unauthorized supplier downgrades of drawing CTQs to SC in `100% Dimensional Results`, `Control Plan`, or `Inspection Plan`.
22. **Mandatory Process Capability & Index Audit for CTQs**: If CTQ characteristics exist on drawing or in PQP, `Process Capability` study is strictly mandatory ($C_{pk} \ge 1.33$ per QP-04-HHI-001). Flags critical failure if `Index` sheet marks `Process Capability` as `N/A` or if the `Process Capability` sheet is empty.
23. **Dual Unit Pairing Proximity, Phone/Fax, CAS Chemical & Lab Tracking Code Filtering**: Requires metric indicators (`[...]` or `mm`/`in`) before dual-unit pairing to prevent unrelated distant inch dimensions (e.g., `0.094"` and `⌀ 2.4"`) from falsely consuming each other. Automatically strips vendor telephone/fax/mobile/tax/zip codes (e.g. `TEL: 05-2785147`), CAS registry numbers (`\b\d{2,7}-\d{2}-\d\b`), and laboratory tracking numbers (`TW/2500126/VoC`, `SHA24-0244670-0001`, `CAN25-0017434-0001`, `ETR25100709`, `SZXEC...`, `TAOEC...`, `NGBEC...`, `PT-...`, `REPORT NO...`) from text prior to part number mining so supplier contact details and chemical test reports do not trigger false-positive copy-paste alarms.
24. **Index & PQS Checklist Alignment & VML Form Control Auditing**: Cross-audits status markings in the `Index` worksheet (Active `'x'` vs `'N/A'`) and `PQS` cover page checklist (Rows 54-72). Crucially parses Excel VML Form Control Checkboxes (`vmlDrawing*.vml` and `ctrlProps/ctrlProp*.xml`) to eliminate false negatives when cell text is blank. Strictly audits whether any item simultaneously has both `Approved` and `N/A` checked, flags multiple conflicting statuses, detects accidental row 53 header checks, and verifies 3-way alignment across Drawing, BOM, Gage R&R, and capability sheets.
25. **Strict Decimal Precision, ASME Note 4 Dimension Decimal Binding & Visual Rounding Auditing**: Validates that nominal dimensions and tolerance specifications in PQP sheets match the engineering drawing's exact decimal precision (both in numeric value and Excel `cell.number_format`, e.g. `0`, `0.0`, `0.00`, `0.000`).
    - **ASME Note 4 尺寸規格小數位數決定公差 (Dimension Precision Determines Tolerance)**: 圖面 Note 4 通用公差（如 `.XX ± .01`, `.XXX ± .005`, `ANGLES ± 1°`）之 `.XX` / `.XXX` **嚴格規範的是「尺寸規格（公稱值）本身標註的小數位數」**，由尺寸規格位數綁定對應公差等級。當圖面尺寸標註為 3 位小數（如 `8X .040`），其尺寸規格本體為 `.XXX`，故適用 `±.005`（3 位小數公差）。
    - **尺寸規格位數不足與公差位數矛盾檢測 (Dimension Precision Mismatch & Visual Rounding Bugs)**: 嚴格稽核 PQP 尺寸規格（Col D）、公差（Col E, F）與上下限（Col G, H）的 `cell.number_format`。若圖面為 3 位小數（`.040`），但 PQP 尺寸規格儲存格格式為 `0.00` 顯示成 `0.04`（2 位小數），工具會明確判定並警示：**「尺寸規格 2 位但公差 3 位（0.005）之位數規則矛盾」**，且上限公式進位偏差（$0.045 \rightarrow 0.05$，$0.035 \rightarrow 0.04$），要求統一將尺寸規格與上下限格式修正為 `0.000`（完整顯示 `0.040 ±0.005`，上限 `0.045` / 下限 `0.035`）。
    - Supports European comma decimals (`76,98±0,1`, `58,9±0,1`).
26. **Attribute Gage R&R `#DIV/0!` Sub-Workbook Diagnostic**: Intelligently diagnoses `#DIV/0!` formula conditions in Attribute MSA worksheets where primary summary rows (Rows 14+) are unpopulated but embedded multi-station inspection workbooks (e.g. `Worksheet4.xlsx` with 50 station sheets) exist, providing actionable guidance to summarize station data into the main sheet.
27. **Raw FAI Non-Conformance & Zero-Tolerance Concealment Audit**: Automatically inspects text and visual measurement callouts in raw embedded FAI files (`FAIL`, `NG`, `REJECT`, `不合格`). Catches fraudulent reporting where suppliers enter `FAIL` in measurement cells but mark `O` (OK) in results, or conceal failures by truncating outer PQP sheets.
28. **Outer PQP Table Truncation & Balloon Omission Audit**: Cross-compares the complete balloon list in raw embedded FAI workbooks against outer visible PQP dimensional sheets, catching cases where suppliers truncate the table at balloon 31 and omit later visual/material/paint masking requirements (balloons 32~37).
29. **Sibling Cavity Embedding Duplication Audit**: Scans sibling embedded workbooks attached to different cavities or surface finishes (e.g. Cavity 1 vs Cavity 2, US15 vs 514) and flags 100% duplicate embedded files where suppliers copy-pasted Cavity 1 without updating to the true Cavity 2 raw report.
30. **Supplier Abbreviated FAI Diagnostic & Balloon Remapping**: Identifies when suppliers submit an abbreviated FAI covering only key/SC dimensions (e.g. 7 items) while the outer PQP covers 40 full items with reference notes (`依重點尺寸量測請參閱附檔`). Remaps shifted balloon numbers by nominal matching (e.g. FAI Balloon 1 [0.031] -> PQP Balloon 38 [0.031] SC wall thickness) to verify data integrity without triggering false "altered measurement" alarms.
31. **Vector Stroke `±` Tolerance Recognition**: Fallbacks to extract tolerances when CAD draws `±` as vector graphic lines between text tokens (e.g. Detail C `.031` followed by `.003 ◇`).
32. **一鍵儲存格格式精度無損修復 (Cell Number Format Precision Auto-Fix, `--fix-formats` / `fix-formats`)**: 支援 `--fix-formats` 與獨立 `fix-formats` 子指令。當檢測到外層 PQP 尺寸規格儲存格格式因設為 `0.00` 或 `0.0` 導致末位 0 截斷、位數矛盾或極限進位失真時，自動調用 Windows Excel Win32 COM（備援 openpyxl 安全模式）在保留所有 OLE 物件、向量球標圖形與儲存格公式的前提下，自動將 Col D (Nominal)、Col E/F (Tol)、Col G (Upper Spec)、Col H (Lower Spec) 統一修正為 `0.000`，Col I (Measurement) 修正為 `0.0000`，並將 MAX R 尺寸的公差修正為 `0.00`。修復前自動建立 `.bak` 備份。

## Quick Start
Run the compliance check in a single command using `uv run`:

```bash
uv run C:/Users/jeffho/.gemini/config/skills/drawing-spec-pqp-comparison/scripts/compare_dwg_pqp.py all \
  --pqp "/path/to/pqp_workbook.xlsm" \
  --dwgs "/path/to/dwg1.pdf,/path/to/dwg2.pdf" \
  --output "/path/to/audit_report.md"
```

## Utility Scripts

The CLI script `compare_dwg_pqp.py` has four subcommands:

### 1. `extract`
Extracts specification metadata (drawing number, revision, weight, torque notes, and balloon dimensions) from drawing PDFs.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/drawing-spec-pqp-comparison/scripts/compare_dwg_pqp.py extract \
  --dwgs "/path/to/drawing.pdf" \
  --output "/path/to/spec.json"
```

### 2. `audit`
Audits PQP workbook compliance against the extracted specifications. Add `--fix-formats` to auto-repair cell number format precision.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/drawing-spec-pqp-comparison/scripts/compare_dwg_pqp.py audit \
  --pqp "/path/to/pqp_workbook.xlsm" \
  --output "/path/to/audit_report.md" \
  --fix-formats
```

### 3. `all`
Performs end-to-end extraction and PQP audit checks, writing the output report to a Markdown file. Add `--fix-formats` to auto-repair cell number format precision.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/drawing-spec-pqp-comparison/scripts/compare_dwg_pqp.py all \
  --pqp "/path/to/pqp_workbook.xlsm" \
  --dwgs "/path/to/dwg1.pdf,/path/to/dwg2.pdf" \
  --output "/path/to/audit_report.md" \
  --fix-formats
```

### 4. `fix-formats`
Auto-fixes outer PQP dimensional cell number format precision discrepancies (restoring Col D/E/F/G/H to `0.000`, Col I to `0.0000`, MAX R to `0.00`), eliminating visual rounding bugs and Note 4 tolerance contradictions while preserving all embedded OLE objects and shapes.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/drawing-spec-pqp-comparison/scripts/compare_dwg_pqp.py fix-formats \
  --pqp "/path/to/pqp_workbook.xlsm" \
  --dwgs "/path/to/drawing.pdf"
```

## Common Mistakes
1. **Wrong File Paths**: Always use absolute paths with forward slashes on Windows.
2. **Missing Sheets**: The tool scans visible worksheets. Make sure the workbook has visible sheets for the selected template.
3. **COM Server Rejected Error**: If converting `.xls` to `.xlsm` via Excel COM fails with "Call was rejected by callee", kill any hanging background `excel.exe` processes and try again.
