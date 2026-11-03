---
name: pqp-drawing-msa-auditor
description: >-
  Audits and cross-checks Product Quality Plan (PQP) Excel workbooks against mechanical drawing PDFs and spec files.
  Performs 10-dimension comprehensive quality audits: metadata consistency across sheets, Index & PQS checklist alignment,
  Control Plan torques and CTQ/SC traceability, Inspection Plan dimensional and SC balloon coverage, Gage R&R and Attribute MSA
  numerical threshold auditing (% Contribution, % Study Var, % Tolerance, Distinct Categories, Effectiveness Scores), contiguous
  100% dimensional inspection with drift alerts, embedded raw FAI sub-workbook discrepancy auditing, embedded OLE test certificates verification,
  unrelated copy-paste residue detection, and adaptive 3F/FFF photo OCR compliance.
---

# PQP Drawing Specification & MSA Quality Auditor (`pqp-drawing-msa-auditor`)

## Overview
This skill provides a unified, enterprise-grade automated auditing suite for manufacturing Product Quality Plans (PQP), Mechanical Drawing PDFs, and Measurement System Analysis (MSA) workbooks. It merges drawing specification extraction, tolerance cross-referencing, multi-part handling, checklist alignment, and statistical gage capability audits into a single robust tool.

Specifically, it audits **10 core quality dimensions**:
1. **Metadata & Multi-Part / Family Part Integrity**: Verifies Part Number, Description, Drawing Number, Revision, and Process/Product headers match across all visible sheets and align with drawing title blocks, handling multiline family part numbers (e.g. `58740-001\n58741-001`) and child sheets. Crucially audits formula references across all worksheets to detect and diagnose external workbook links (e.g. `=IF(ISBLANK([7]PQS!J5)," ",[7]PQS!J5)` where `[N]` references foreign cached files like `5057819`).
2. **Index & PQS Checklist Alignment & VML Form Control Checkbox Auditing**: Cross-audits status markings in the `Index` worksheet (Active `'x'` vs `'N/A'`) and `PQS` cover page checklist (Rows 54-72) against actual worksheet content, embedded images, and data presence. Crucially parses Excel VML Form Control Checkboxes (`vmlDrawing*.vml` and `ctrlProps/ctrlProp*.xml`) to eliminate openpyxl false negatives where cell text is empty. Strictly verifies that **no single item simultaneously has both `Approved` and `N/A` checked** (or multiple conflicting statuses), flags erroneous header row 53 checkbox marks, and validates 3-way consistency across Drawing, BOM/Matrices, Gage R&R, and Process Capability.
3. **Control Plan Assembly Torques & Reaction Plan Statistical Control**: Cross-verifies assembly step torque descriptions and tolerances against drawing flag notes. Crucially audits **Control Plan Reaction Plan Robustness for Sampling Plans**: on any in-process periodic sampling plan (e.g. 5 pcs / 2 hrs, per shift, per lot), strictly rejects merely discarding/scrapping the failed sample units (`Discard` / `Scrap`) as statistically non-robust. Mandates that the Reaction Plan must require 100% containment and sorting of all suspect units produced since the last acceptable inspection (*"At minimum, should be sorting units that were produced since the last inspection"*), quarantine of non-conforming parts for scrap/rework, and tooling/process adjustment with re-verification before resuming production. Flags empty reaction plans on sampling rows as critical omissions.
4. **Control Plan & Inspection Plan CTQ / SC Coverage & Traceability**: Distinguishes between Critical To Quality (CTQ, cross-hair diamond `⯐` per drawing Note 7 requiring Cpk >= 1.33) and Significant Characteristics (SC, empty diamond `◇` per Note 8). Detects circled index letters (`Ⓐ`~`Ⓛ`), flag notes, and vector perpendicular cross-diamonds directly adjacent to dimensions. Cross-verifies against `100% Dimensional Results` and flags critical misclassifications where suppliers downgrade drawing CTQ callouts to SC. Cross-checks `Inspection Plan` to verify all declared SC/CTQ features are explicitly listed with inspection standards, gages, and frequencies. Intelligently provides **Non-Dimensional SC Guidance**: when suppliers declare SC on visual/material checks (e.g. print artwork, sticker, material) without drawing diamonds, advises removing SC or substituting Attribute MSA rather than triggering false-positive Variable R&R failures.
5. **MSA Gage R&R & Process Capability Quality Thresholds & Attribute MSA Diagnostic**: Verifies 100% data presence and compliance with AIAG/ASSA ABLOY quality thresholds on Variable R&R, Attribute MSA, and Process Capability sheets. Strictly enforces that when CTQ characteristics exist on drawing or in dimensional sheets, `Process Capability` study is mandatory (Cpk >= 1.33); flags any empty sample data or invalid 'N/A' status in `Index`. Intelligently diagnoses **Attribute MSA `#DIV/0!` Sub-Workbook Patterns**: when primary worksheet summary cells (Rows 14+) are unpopulated while a massive embedded sub-workbook (e.g. 44.5MB `Worksheet4.xlsx` with 50 station sheets) exists, advises transferring station summary data rather than reporting calculation formula failure.
6. **100% Dimensional Inspection, Sample Count & Drift Alerts**: Audits dimensional inspection sheets for contiguous integer `Sample#` sequences (`1, 2, ... N`), detects Out-of-Spec (OOS) conditions, and flags borderline measurements (`"Investigate"` disposition). Crucially audits **Sample Count Completeness (< 5 pcs)** to alert if suppliers only measured 3 samples instead of ASSA ABLOY's standard 5-piece sample requirement.
7. **Drawing vs PQP Dimensional Spec Cross-Reference, ASME Note 4 & Strict Decimal Precision**: Matches balloon dimensions, ASME Y14.5 inch dimensions (e.g. `.75`, `1.850`, `⌀ 2.4`), metric dual units (`[19,1 ±0,5]`), angles (`45.0°`), fractions (`8-5/8`), and MIN/MAX/REF limits between drawing callouts and PQP entries. Dual dimension smart pairing strictly requires metric brackets `[...]` or explicit `mm` units to prevent distant unrelated inch dimensions from falsely consuming each other. Robustly parses European comma decimals (`76,98±0,1`, `58,9±0,1`).
   - **ASME Y14.5M Note 4 尺寸規格小數位數決定公差 (Dimension Precision Determines Tolerance)**: 圖面 Note 4 通用公差（如 `.XX ± .01`, `.XXX ± .005`, `ANGLES ± 1°`）之 `.XX` / `.XXX` **嚴格規範的是「尺寸規格（公稱值）本身標註的小數位數」**，由尺寸規格位數綁定對應公差等級。當圖面尺寸標註為 3 位小數（如 `8X .040`），其尺寸規格本體為 `.XXX`，故適用 `±.005`（3 位小數公差）。
   - **尺寸規格位數不足與公差位數矛盾檢測 (Dimension Precision Mismatch & Visual Rounding Bugs)**: 嚴格稽核 PQP 尺寸規格（Col D）、公差（Col E, F）與上下限（Col G, H）的 `cell.number_format`。若圖面為 3 位小數（`.040`），但 PQP 尺寸規格儲存格格式為 `0.00` 顯示成 `0.04`（2 位小數），工具會明確判定並警示：**「尺寸規格 2 位但公差 3 位（0.005）之位數規則矛盾」**，且上限公式進位偏差（$0.045 \rightarrow 0.05$，$0.035 \rightarrow 0.04$），要求統一將尺寸規格與上下限格式修正為 `0.000`（完整顯示 `0.040 ±0.005`，上限 `0.045` / 下限 `0.035`）。
   - **一鍵儲存格格式精度無損修復 (Cell Number Format Precision Auto-Fix, `--fix-formats` / `fix-formats`)**: 支援 `--fix-formats` 與獨立 `fix-formats` 子指令。當檢測到外層 PQP 尺寸規格儲存格格式因設為 `0.00` 或 `0.0` 導致末位 0 截斷、位數矛盾或極限進位失真時，自動調用 Windows Excel Win32 COM（備援 openpyxl 安全模式）在保留所有 OLE 物件、向量球標圖形與儲存格公式的前提下，自動將 Col D (Nominal)、Col E/F (Tol)、Col G (Upper Spec)、Col H (Lower Spec) 統一修正為 `0.000`，Col I (Measurement) 修正為 `0.0000`，並將 MAX R 尺寸的公差修正為 `0.00`。修復前自動建立 `.bak` 備份。
   - **CAD 鄰近文字公差識別邊界條件防誤判 (Adjacent Text Tolerance Isolation & Angle Matching Guard)**: 修正 CAD 圖面中角度尺寸標記（如 `8$`、`4$`）後方緊接其他文字（如 `2X .060`、`.026\n4X`）時，因公差前綴符號可選而誤將相鄰數字判為角度公差（如誤判為 `±2.0°` 或 `±0.026°`）之缺陷；嚴格排除無單位之首點小數（如 `\.\d{2,4}`）被誤配為角度公差，並在匹配迴圈中強制隔離角度 callout 與線性尺寸之交叉比對。
   - **Vector Stroke `±` Tolerance Recognition**: Fallbacks to extract tolerances when CAD draws `±` as vector graphic lines between text tokens (e.g. Detail C `.031` followed by `.003 ◇`).
8. **Form, Fit, Function (3F / FFF) Photo & Content Compliance**: Audits structured table cells in `FFF`/`3F` sheets (Verification, Requirement / Test Performed, Qty Tested, Disposition). Performs **Project Name Copy-Paste Residue Detection** (e.g. catches foreign project names like `FOSTER` when drawing is `TITAN`), validates sample size alignment between `Qty Tested` (e.g. 5 pcs) and 100% Dimensional measured sample count (e.g. 3 pcs), and alerts on overly generic descriptions (`comply with the drawing`). Adaptively runs RapidOCR on physical photos in 3F / Appearance worksheets to verify customer support QR labels, setup codes, brand logos (Kwikset, Weiser), and bilingual texts.
9. **Embedded Raw FAI Sub-Workbook Discrepancy Auditing & Multi-Layer Drawing Architecture**: Inspects raw embedded child FAI workbooks (`xl/embeddings/Microsoft_Excel_Worksheet*.xlsx`) and cross-checks sample measurements against active visible PQP sheets to flag altered data and OOS conditions.
   - **Supplier Abbreviated FAI Diagnostic & Balloon Remapping**: Identifies when suppliers submit an abbreviated FAI covering only key/SC dimensions (e.g. 7 items) while the outer PQP covers 40 full items with reference notes (`依重點尺寸量測請參閱附檔`). Remaps shifted balloon numbers by nominal matching (e.g. FAI Balloon 1 [0.031] -> PQP Balloon 38 [0.031] SC wall thickness) to verify data integrity without triggering false "altered measurement" alarms.
   - **Non-Numeric / Visual FAIL & Contradictory Pass Detection**: Parses text measurements (`FAIL`, `NG`, `REJECT`, `不合格`) on visual/note callouts (e.g. Balloon 37 `Paint Masking Area` on matte black finish 514). Flags fraudulent or erroneous reporting where measurement column states `FAIL` but result column marks `O` (OK) or when suppliers omit the item from the outer sheet.
   - **Outer PQP Table Truncation & Balloon Omission Audit**: Cross-compares the full balloon set of the embedded raw FAI sub-workbook against the outer visible PQP dimensional sheet. Catches critical vendor omissions where inspection tables are truncated (e.g. outer table stops at Balloon 31, omitting Balloons 32~37 covering material, appearance, raised characters, markings, brush stroke, and paint masking).
   - **Sibling Embedded Workbook Duplication / Cavity Mismatch Audit**: Scans sibling embedded workbooks attached to different cavity sheets (e.g. Cavity 1 vs Cavity 2) and flags 100% duplicate embedded files where suppliers copy-pasted Cavity 1 without embedding the true Cavity 2 raw report.
   - **Accurate Relationship-Based Host Worksheet Pairing**: Resolves exact host worksheet relationships from `xl/worksheets/_rels/` and `xl/drawings/_rels/` to pair embedded files directly with their host sheets, eliminating spurious cross-sheet mismatch noise.
   - **Copy-Paste Shift Detection across Family Parts**: Automatically scans sibling part sheets in the raw FAI and identifies if supplier mistakenly copied measurements from an adjacent part number (e.g. `50013-001` pasted into `50012-001`).
   - **Multi-Layer / Sub-Assembly Drawing Architecture**: Distinguishes Top-Level Assembly drawings (finished assembly dimensions, total thickness `0.85 ± 0.05 mm`, finished aperture `4 ± 0.2` & `3 ± 0.2 mm`) from Layer Component Drawings (Overlay, Adhesive, Shading paper, PC, Metal dome with single layer thickness `0.2 ± 0.02 mm`, aperture `6 mm`, `5 mm`). Audits multi-section sub-workbooks (300+ rows covering Section 0 Assembly + Sections 1~10 raw material layers) without generating false-positive missing balloon alarms on the top-level assembly drawing. Includes deduplication across multiple embedded file copies.
10. **Embedded Test Certificates, Product Name Cross-Check, SGS/CTI Lab Tracking Filter & Material Cert Recognition**: Summarizes embedded OLE test reports (`oleObject*.bin`, e.g. SGS RoHS 2.0, CTI REACH SVHC, ISTA 3A Drop Test, Material specs) with their hosting worksheets. Accurately classifies engineering drawing PDFs by prioritizing title blocks (`DO NOT SCALE`, `THIRD ANGLE PROJECTION`, `DWG NO`) over packaging notes. Crucially performs **Product Name / Description Cross-Check** between embedded Material Specification / Certificate of Compliance (CoC) and Drawing Title / PQP Part Description, immediately catching template reuse blunders where suppliers submit carton/packaging CoCs for divider/bracket parts. Supports 5-digit dash 3-digit family part numbers (`\b\d{5}-\d{3}\b`, e.g. `69633-002`) and strips vendor telephone/fax/mobile/tax/zip numbers (e.g. `TEL: 05-2785147`), CAS Registry Numbers (`\b\d{2,7}-\d{2}-\d\b`), and lab tracking codes (`TW/2500126/VoC`, `SHA24-0244670-0001`, `CAN25-0017434-0001`, `ETR25100709`, `SZXEC...`, `TAOEC...`, `NGBEC...`, `PT-...`, `REPORT NO...`) to eliminate false-positive copy-paste part number alarms. Filters table header tokens (`SN ID`, `SAMPLE ID`, `ITEM NO`) and prioritizes `Sample Name:` / `样品名称：`. Recognizes raw material certificates embedded in `[Material Certs]` (e.g. `White PET adhesive tape`, `聚碳酸酯薄膜及片材`, `3M SP7533`) as compliant component certifications rather than flagging finished assembly description mismatches.

---

## Quality Standards & Thresholds
- **SCREEN % EFFECTIVE SCORE** = 100%
- **SCREEN % EFFECTIVE SCORE vs. ATTRIBUTE (4)** = 100%
- **% Contribution** < 9.0%
- **% Study Var** < 20.0%
- **% Tolerance** < 20.0%
- **Number of Distinct Categories (ndc)** > 4
- **Process Capability (Cpk / Ppk)**: Fully populated without `#DIV/0!` or empty sample cells; Cpk >= 1.33 strictly required whenever CTQ features exist (Index cannot be 'N/A')
- **CTQ vs SC Classification**: CTQ features marked with cross-hair diamond `⯐` must not be downgraded to SC in Dimensional Results or Control Plan
- **Control Plan Reaction Plan Robustness for Sampling Plans**: On any in-process sampling plan (e.g. periodic 5 pcs every 2 hours), Reaction Plan must strictly require sorting/quarantining all units produced since the last acceptable inspection ('sorting units that were produced since the last inspection') and corrective verification before restart. Merely stating 'Discard' or 'Scrap' or leaving Reaction Plan blank is flagged as a Critical Non-Conformance (FAIL)
- **Sample Count Requirement**: 5 pcs standard (Sample #1~#5); flags warning if < 5 pcs
- **Sample # Sequence**: Contiguous integers `1, 2, ... N` across all balloon items
- **Inspection Plan SC Coverage**: 100% of declared SC/CTQ features must be explicitly listed with Inspection Standard, gage, and sample frequency
- **Inspection Plan Specificity**: Inspection Standard cells must not be blank when dimensional checks are defined
- **Strict Decimal Precision**: Nominal and tolerance decimal places in PQP must strictly match drawing callouts (e.g., drawing `±0.5` must not be formatted as `0.50`, and drawing `±0.05` must not be truncated to `±0.1`)
- **Raw FAI Non-Conformance & Zero-Tolerance Concealment**: 100% PASS required on all inspection items (visual, notes, dimensions). Zero tolerance for `FAIL`/`NG` items concealed by outer sheet truncation or contradictory `O` markings
- **Balloon Coverage Completeness**: 100% of balloons present in embedded raw FAI sub-workbooks must be visible in outer PQP dimensional results without truncation
- **Unique Cavity Reporting**: Sibling cavity sheets must have independent, non-duplicate embedded raw FAI reports
- **Multi-Layer Drawing Architecture Compliance**: Sub-assembly/layer-level dimensions must be audited against their respective breakdown drawing sheets (Sheet 3~5) rather than conflicting with top-level assembly drawings (Sheet 1~2)

---

## Quick Start

Run the full end-to-end 10-dimension audit in a single command using `uv run`:

```bash
uv run C:/Users/jeffho/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py all \
  --pqp "/path/to/pqp_workbook.xlsm" \
  --dwgs "/path/to/dwg1.pdf,/path/to/dwg2.pdf" \
  --output "/path/to/audit_report.md" \
  --json-output "/path/to/audit_report.json"
```

> **Note**: If drawing PDFs are not specified, the script automatically searches for `spec.json` in the target workbook directory.

---

## CLI Subcommands

### 1. `all` (Full End-to-End Extraction & 10-Dimension Audit)
Extracts specifications from drawing PDFs and executes the complete 10-dimension audit suite against the PQP workbook.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py all \
  --pqp "/path/to/pqp_workbook.xlsm" \
  --dwgs "/path/to/drawing.pdf" \
  --output "/path/to/audit_report.md"
```

### 2. `extract` (Extract Drawing Specifications to JSON)
Extracts title block metadata, flag notes/torques, SC/CTQ diamond notes, general/individual tolerances, and balloon dimensions from drawing PDFs.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py extract \
  --dwgs "/path/to/drawing.pdf" \
  --output "/path/to/spec.json"
```

### 3. `audit` (PQP Compliance Audit against Spec JSON / Drawings)
Audits the PQP workbook using pre-extracted `spec.json` or drawing files. Add `--fix-formats` to auto-repair cell number format precision.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py audit \
  --pqp "/path/to/pqp_workbook.xlsm" \
  --spec "/path/to/spec.json" \
  --output "/path/to/audit_report.md" \
  --fix-formats
```

### 3.1 `fix-formats` (Auto-Fix Cell Number Format Precision)
Auto-fixes outer PQP dimensional cell number format precision discrepancies (restoring Col D/E/F/G/H to `0.000`, Col I to `0.0000`, MAX R to `0.00`), eliminating visual rounding bugs and Note 4 tolerance contradictions while preserving all embedded OLE objects and shapes.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py fix-formats \
  --pqp "/path/to/pqp_workbook.xlsm" \
  --dwgs "/path/to/drawing.pdf"
```

### 4. `msa` (Standalone MSA Gage R&R & Capability Audit)
Audits Variable R&R, Attribute MSA, and Process Capability sheets for % Contribution, % Study Var, % Tolerance, Distinct Categories, and 100% effective scores.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py msa \
  --pqp "/path/to/pqp_workbook.xlsm"
```

### 5. `checklist` (Standalone Index & PQS Checklist Alignment Audit)
Audits whether status markings in the `Index` worksheet and `PQS` cover page match actual worksheet contents, images, and data activity.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py checklist \
  --pqp "/path/to/pqp_workbook.xlsm"
```

### 6. `metadata` (Standalone Metadata Consistency Audit)
Validates key metadata fields (Part Number, Description, Drawing Number, Revision, Process/Product) across all visible worksheets.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py metadata \
  --pqp "/path/to/pqp_workbook.xlsm" \
  --spec "/path/to/spec.json"
```

### 7. `inspection` (Standalone Inspection Plan Dimensional & SC Balloon Audit)
Audits whether all declared SC/CTQ features and 100% dimensional balloons are explicitly listed in the `Inspection Plan` worksheet, flagging generic descriptions and blank inspection standards.
```bash
uv run C:/Users/jeffho/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py inspection \
  --pqp "/path/to/pqp_workbook.xlsm"
```

---

## Common Mistakes & Troubleshooting
1. **File Paths on Windows**: Always use absolute paths with forward slashes (e.g., `C:/path/to/file.xlsm`).
2. **Temporary Lock Files**: Ensure the target workbook is not locked by Excel (`~$...`).
3. **Empty R&R Sheets**: If trial data rows (C24:L32 in Variable R&R) are blank, the script flags `#DIV/0!` formula errors.
4. **Drawing OCR Fallback**: If drawing PDF vector text is converted to curves/polygons, the tool automatically uses `pymupdf` + `rapidocr-onnxruntime` to render and OCR drawing pages at 300 DPI.
5. **OpenPyXL Font Limit Patch**: The tool automatically patches OpenPyXL drawing font descriptors (`Font.pitchFamily.max = 255`) to prevent XML parser crashes on complex CAD shape drawings.
