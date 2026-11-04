---
name: pcb-pqp-spec-comparison
description: >-
  Audits and compares a PCB Specification (PDF), FAB Drawing (PDF), a Product Quality Plan (PQP) Excel sheet,
  and a QA Inspection / Outgoing Report (PDF or XLS/XLSX 承認書) to determine manufacturing compliance, flagging differences
  in dimensions, stackup layers, drill tables, solder mask thickness, part numbers, revisions, tolerances, and missing documents.
---

# PCB PQP, Spec & QA Comparison Skill (PCB 品質與工程規格比對稽核技能)

## 概述 (Overview)
本技能提供 PCB 製造品質文件的自動化審查與比對。全面稽核客戶官方 PCB 規格書 (Specification)、FAB 工程圖面 (Drawing)、供應商提報的產品質量計畫 (PQP Excel)、以及工廠出貨檢驗報告 (QA Outgoing Report / 產品承認書)。

支援出貨報告多種格式：
- **PDF 數位文字檔** / **PDF 掃描圖檔**（自動呼叫 RapidOCR 進行光學字元辨識）
- **Excel 承認書檔案 (`.xls` / `.xlsx`)**（自動解析 `承认书1`、`承认书2`、`切片报告`、`沉金报告`、`承认书封面` 等多頁籤）

---

## 稽核比對重點 (Audit Scope)

1. **基本資訊一致性 (Metadata Alignment)**:
   - 料號 (Part Number)、版本 (Revision)、圖面號 (Drawing Number)、供應商名稱。
   - 容錯識別工廠內部衍生後綴（如 `-000`、`_Rev2.0`）。

2. **疊層結構比對 (Stack-up Table Comparison)**:
   - 核對 Top/Bottom 防焊層 (Solder Mask)、各銅箔層 (L1~L4 Cu)、PP 膠片 (Dielectric)、Core 芯板厚度與成品總板厚。
   - 三方比對：FAB 圖面設計規格 vs PQP FAI 實測 vs QA 切片微切片報告實測。

3. **鑽孔規格與孔徑檢驗 (Drill Table & Hole Size)**:
   - 比對圖面 DRILL TABLE 鑽孔符號、孔數、標稱孔徑、公差與 PTH/NPTH 鍍通屬性。
   - 稽核 PQP FAI 首件孔徑量測與 QA 承認書孔徑檢驗，自動標記公差消耗率 >80% 之需調查項目 (Investigate)。
   - 檢查圖面是否遺漏公差標註（如 3.5mm 孔公差空白）。

4. **表面處理與鍍層分析 (ENIG Plating & Solder Mask)**:
   - 化金厚度 (Au THK ≥ 2 u")、化鎳厚度 (Ni THK ≥ 120 u" 或依 IPC-4552)。
   - 防焊層顏色（Green/Black 等）、防焊層厚度（≥8µm / 50µm）、文字字符顏色 (Silkscreen)、UL94V-0 阻燃等級。

5. **外形尺寸與板邊公差 (Outline Dimensions & Tolerances)**:
   - 檢查長寬外形標稱值與公差設定（如圖面要求 ±0.127mm vs 供應商誤設 ±0.15mm）。
   - 板彎板翹度 (Bow & Twist ≤ 0.75%)。

6. **統計品質分析 (Cp/Cpk & Gage R&R)**:
   - 製程能力：Cpk ≥ 1.33（CTQ 關鍵特性）。
   - Variable Gage R&R：標準 % Contribution < 9%, % Study Var < 20%, % Tolerance < 20%, ndc > 4。
   - Attribute Gage R&R 填寫完整度審查。

7. **佐證報告完整性 (Report Completeness)**:
   - 稽核 PQP 工作簿中之 DFM、Electrical Test Report（電測報告）、Finish Test Reports（表面處理報告）、Material Certs 等是否空白或缺漏。

---

## 依賴套件 (Dependencies)
- `python-calamine` (高效讀取複雜 Excel 工作表)
- `pandas` (資料表處理)
- `pypdf` (PDF 規格書解析)
- `pymupdf` (PDF 渲染與圖片擷取)
- `rapidocr-onnxruntime` (快速輕量 ONNX OCR 辨識)

---

## 執行方式 (Quick Start)

使用 `uv run` 呼叫比對指令：

```bash
uv run C:/Users/jeffho/.gemini/config/skills/pcb-pqp-spec-comparison/scripts/compare_pcb_pqp.py compare \
  --spec "/path/to/spec.pdf" \
  --pqp "/path/to/pqp.xlsx" \
  --qa "/path/to/qa_report.xls" \
  --fab "/path/to/fab_drawing.pdf" \
  --output "/path/to/comparison_report.md"
```

### 參數說明 (Arguments):
- `--spec` (必填): PCB 規格書 PDF 檔案路徑。
- `--pqp` (必填): 產品質量計畫 PQP Excel 活頁簿路徑 (`.xlsx` 或 `.xls`)。
- `--qa` (必填): 工廠出貨品質報告路徑（支援 `.pdf` 數位檔/掃描檔 或 `.xls`/`.xlsx` 承認書格式）。
- `--fab` (選填): FAB 工程圖面 PDF 路徑（用於提取 Stack-up Table 與 Drill Table）。
- `--output` (必填): 輸出之 Markdown 稽核報告路徑。
- `--json-output` (選填): 輸出之完整結構化 JSON 檔案路徑。

---

## 輸出報告章節 (Generated Report Sections)
1. **一、 稽核摘要 (Audit Summary)**：整體判定狀態（PASS / HOLD / REJECT）、料號、版本、出貨批號。
2. **二、 關鍵不合規項目與缺失 (Critical Gaps)**：板邊公差不符、缺失測試報告、電鍍超差等。
3. **三、 警告與製程預警 (Warnings & Notes)**：孔徑公差使用率 >80%、圖面漏標提醒、料號後綴備註。
4. **四、 統計量測與製程品質 (Statistical Quality)**：Gage R&R 與 Cpk 完整數據表格。
5. **五、 關鍵規格參數核對 (Parameter Details)**：板厚、鍍金、鍍鎳、防焊厚度、板彎翹等。
6. **六、 Stack-up Table 疊層結構三方比對**：各層材料與厚度三方核對表。
7. **七、 Drill Table 鑽孔規格與實測比對**：孔徑符號、公差、實測範圍、PTH/NPTH 分類表。
8. **八、 行動建議 (Action Items)**：按高/中/低優先級列出供應商 SDY 與內部工程端需處理的事項清單。