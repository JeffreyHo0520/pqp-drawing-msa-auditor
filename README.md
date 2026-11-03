# PQP Drawing & MSA Quality Inspection Suite (雙軌技能套件)

> **適用於 Antigravity / Claude Code / OpenAI Codex / Cursor / Gemini CLI**  
> 企業級製造業產品品質計畫書（PQP）、工程圖面（PDF）與量測系統分析（MSA）自動化交叉稽核與精度修復套件。

---

## 📖 目錄 (Table of Contents)
- [套件簡介 (Overview)](#套件簡介-overview)
- [核心雙軌技能 (Included Skills)](#核心雙軌技能-included-skills)
- [10 維度全面品質稽核體系 (10-Dimension Audit)](#10-維度全面品質稽核體系-10-dimension-audit)
- [跨電腦一鍵安裝指南 (Installation Guide)](#跨電腦一鍵安裝指南-installation-guide)
- [命令列操作手冊 (CLI Usage & Subcommands)](#命令列操作手冊-cli-usage--subcommands)
- [AI Agent 呼叫方式 (How to Use in Agents)](#ai-agent-呼叫方式-how-to-use-in-agents)
- [依賴套件與環境 (Prerequisites)](#依賴套件與環境-prerequisites)
- [授權條款 (License)](#授權條款-license)

---

## 🎯 套件簡介 (Overview)
在製造業與供應商零件認證（Part Qualification Process）中，審查 PQP Excel 活頁簿與工程圖面 PDF 通常耗費大量工程師時間，且極易遺漏細微但致命的品質缺失（如：圖面 Note 4 公差位數矛盾、PQS 檢核表勾選倒置、抽樣計畫僅報廢不良品未回溯挑檢、內嵌原始 FAI 球標脫節、複製貼上誤植其他專案料號等）。

本套件整合兩大專業技能，提供端到端 100% 自動化解析、交叉比對與報告生成能力：
1. **`pqp-drawing-msa-auditor`**：10 維度全面品質稽核與儲存格精度修復引擎。
2. **`drawing-spec-pqp-comparison`**：工程圖面向量/OCR 規格萃取與 PQP 雙軌比對引擎。

---

## 🛠️ 核心雙軌技能 (Included Skills)

| 技能名稱 (Skill) | 指令入口 (Slash Command) | 核心功能 |
| :--- | :--- | :--- |
| **`pqp-drawing-msa-auditor`** | `/pqp-drawing-msa-auditor` | 執行 10 維度全方位深度稽核（元數據、PQS VML 核取方塊、Reaction Plan 圍堵管制、MSA 閾值、Cpk、100% 尺寸檢驗、內嵌原始 FAI 脫節與 26+ OLE 認證報告審查），並支援 `--fix-formats` 一鍵無損修正儲存格小數精度。 |
| **`drawing-spec-pqp-comparison`** | `/drawing-spec-pqp-comparison` | 解析工程圖面 PDF 標題欄、ASME Y14.5 通用公差（Note 4）、flag notes 裝配扭力、CTQ/SC 菱形標記與英吋/公制雙單位球標，並與 PQP 尺寸及管制計畫進行雙向比對。 |

---

## 🔍 10 維度全面品質稽核體系 (10-Dimension Audit)

1. **元數據與家族零件一致性 (Metadata & Multi-Part Integrity)**：
   - 稽核全活頁簿 20+ 分頁之 Part Number、Description、Drawing Number、Revision 是否一致並對齊圖面。
   - 深度掃描公式，偵測並標記跨檔案外部快取連結殘留（如 `=IF(ISBLANK([4]PQS!L13)...)`）。
2. **Index 與 PQS 檢核表對齊 (Index & PQS Checklist Alignment)**：
   - 解析底層 Excel VML Form Control 核取方塊（`vmlDrawing*.vml` 與 `ctrlProps`），解決 openpyxl 讀取文字為空的假陰性。
   - 嚴格稽核是否發生 **「Approved 與 N/A 同時勾選」** 或 **「無數據卻誤勾 Approved、有數據卻誤勾 N/A」** 之嚴重倒置。
3. **Control Plan 裝配扭力與異常處置統計管制 (Reaction Plan Robustness)**：
   - 比對工序扭力與圖面 flag notes。
   - **抽樣異常管制嚴格審查**：凡週期性抽樣檢驗工序，處置若僅填寫「報廢/丟棄當下樣品 (Discard/Scrap)」一律判定不合格，強制要求載明 **「回溯挑檢自前次檢驗合格以來生產的所有成品 (Sort units produced since last inspection)」**。
4. **CTQ / SC 涵蓋與降級防呆 (CTQ/SC Traceability & Downgrade Prevention)**：
   - 區分圖面實心十字菱形（`⯐` CTQ，要求 $C_{pk} \ge 1.33$）與空心菱形（`◇` SC）。
   - 嚴禁供應商私自將圖面 CTQ 降級標記為 SC。
5. **MSA Gage R&R 與製程能力評估 (MSA Quality Thresholds & Cpk)**：
   - 稽核 Variable Gage R&R 指標：% Contribution (< 9.0%)、% Study Var (< 20.0%)、% Tolerance (< 20.0%)、ndc (> 4)。
   - 稽核 Attribute MSA 100% Effective Scores 與 Process Capability 數據完整性。
6. **100% 尺寸檢驗與樣品數完整性 (Sample Count Completeness)**：
   - 嚴格檢查樣品量測數是否達到標準 **5 顆 (Sample #1~#5)**，若僅測量 3 顆即時發出特許警示。
   - 檢查連續整數樣品編號序列與公差邊界漂移警告（Investigate）。
7. **圖面 vs PQP 規格交叉比對與 ASME Note 4 位數綁定 (Strict Decimal Precision)**：
   - 智慧配對英吋（`.040`）、公制雙單位（`[1.02 ±0.13]`）、角度（`45.0°`）與分數（`8-5/8`）。
   - **Note 4 精度綁定**：嚴格稽核尺寸規格本體之小數位數，檢測「尺寸規格 2 位但公差 3 位（0.005）之位數規則矛盾」與極限進位偏差。
   - 支援 `--fix-formats` 自動調用 Win32 COM / openpyxl 將 Col D/E/F/G/H 格式無損修復為 `0.000`。
8. **Form, Fit, Function (3F / FFF) 照片與外觀合規性**：
   - 自動調用 RapidOCR 解析三防漆塗佈、零件佈局、QR Code、客戶服務序號與銘牌文字。
9. **內嵌原始 FAI 活頁簿脫節稽核 (Embedded Raw FAI Discrepancy)**：
   - 深入解開 `xl/embeddings/Microsoft_Excel_Worksheet*.xlsx`，比對原始量測值與外層 PQP。
   - 檢測外層已擴充至 37 項球標但內嵌附件仍為舊版 22 項球標之「體系脫節」與零容忍不合格隱匿。
10. **內嵌認證報告與料號關聯性審查 (Test Certificates & Copy-Paste Residue)**：
    - 分類並總結 26+ 件 OLE 二進位測試報告（SGS RoHS 2.0、CTI REACH SVHC、ISTA 3A、DFM）。
    - 比對內嵌檔案料號，過濾實驗室編號與電話傳真，精準揪出誤植之非本產品料號（殘留複製範本）。

---

## 🚀 跨電腦一鍵安裝指南 (Installation Guide)

在另一台電腦上下載並啟用這兩項技能，只需以下任一方式：

### 方式 A：PowerShell 一鍵線上安裝 (Windows 推薦)
直接在 PowerShell (或 Windows Terminal) 中執行以下指令，自動完成下載與安裝至全域技能目錄：

```powershell
powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/JeffreyHo0520/pqp-drawing-msa-auditor/main/install.ps1 | iex"
```

> **說明**：預設會將技能安裝至 `$HOME\.gemini\config\skills\`（所有工作區全域生效）。若需安裝至 Claude Code，可在下載後加上 `-Claude` 參數；若需安裝至當前專案，可加上 `-Workspace` 參數。

---

### 方式 B：Git Clone 本機安裝

```bash
# 1. 複製儲存庫
git clone https://github.com/JeffreyHo0520/pqp-drawing-msa-auditor.git
cd pqp-drawing-msa-auditor

# 2. Windows 執行安裝 (預設安裝至 ~/.gemini/config/skills)
.\install.ps1

# 或安裝至 Claude Code 目錄 (~/.claude/skills)
.\install.ps1 -Claude

# 或安裝至當前專案工作區 (.agents/skills)
.\install.ps1 -Workspace

# Linux / macOS 執行安裝
chmod +x install.sh
./install.sh
```

---

## 💻 命令列操作手冊 (CLI Usage & Subcommands)

本套件推薦使用高效 Python 套件管理器 `uv` 直接執行（無需繁瑣虛擬環境配置）：

### 1. 執行 10 維度全面端到端稽核 (Full Audit)
```bash
uv run ~/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py all \
  --pqp "PQP_Workbook.xlsx" \
  --dwgs "Drawing.pdf" \
  --output "audit_report.md" \
  --json-output "audit_report.json"
```

### 2. 圖面規格萃取 (Extract Spec JSON)
```bash
uv run ~/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py extract \
  --dwgs "Drawing.pdf" \
  --output "spec.json"
```

### 3. 一鍵自動修復儲存格小數格式 (Auto-Fix Decimal Formats)
自動修正外層 100% 尺寸表格式（Col D/E/F/G/H 統一為 `0.000`，Col I 為 `0.0000`），消除 Note 4 精度矛盾並保留所有內嵌 OLE 物件：
```bash
uv run ~/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py fix-formats \
  --pqp "PQP_Workbook.xlsx" \
  --dwgs "Drawing.pdf"
```

### 4. 獨立檢驗模組
```bash
# 單獨稽核 Gage R&R 與 Cpk 閾值
uv run ~/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py msa --pqp "PQP_Workbook.xlsx"

# 單獨稽核 Index 與 PQS 核取方塊
uv run ~/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py checklist --pqp "PQP_Workbook.xlsx"

# 單獨稽核元數據跨表一致性
uv run ~/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py metadata --pqp "PQP_Workbook.xlsx"

# 單獨稽核 Inspection Plan 球標涵蓋
uv run ~/.gemini/config/skills/pqp-drawing-msa-auditor/scripts/pqp_dwg_msa_auditor.py inspection --pqp "PQP_Workbook.xlsx"
```

---

## 🤖 AI Agent 呼叫方式 (How to Use in Agents)

安裝完成後，在任何支援的 AI Agent 對話框中即可直接觸發：

### 1. Antigravity / Gemini CLI
直接輸入斜線指令或自然語言：
```text
/pqp-drawing-msa-auditor
/drawing-spec-pqp-comparison
```
或對話輸入：
> *"幫我稽核當前目錄下的 PQP Excel 與最新工程圖面 PDF，產出完整 10 維度審查報告。"*

### 2. Claude Code / Cursor / Codex
在提示詞中指定：
> *"請載入 pqp-drawing-msa-auditor 技能，比對 PQP 文件與工程圖面，並檢查 PQS 檢核表與 Control Plan 的 Reaction Plan 是否符合回溯挑檢規範。"*

---

## 📦 依賴套件與環境 (Prerequisites)

* **Python 版本**：Python 3.10+
* **核心依賴庫**：
  * `openpyxl`（Excel 讀寫與底層 XML/VML 解析）
  * `pymupdf` (fitz)（PDF 向量文字萃取與 300 DPI 高畫質渲染）
  * `rapidocr-onnxruntime`（離線輕量級 OCR）
  * `pywin32`（Windows Excel Win32 COM 自動化無損修復）

使用 `uv` 執行時，指令會自動解析並配置相應環境。

---

## 📄 授權條款 (License)

本專案採用 [MIT License](LICENSE) 授權開源。
