---
name: gantt-chart-pro
version: "16.0"
description: >
  Excel 橫道圖（甘特圖）生成器 v16.0。觸發詞：甘特圖、橫道圖、施工進度、GanttChart、進度計劃。
  基於 v3.0 統一渲染引擎，數據驅動、配置分離、非交互式 API。
  從示例學校實際 進度表 提取色彩方案與排版規範，專業級工程橫道圖輸出。
triggers:
  - 甘特圖
  - 橫道圖
  - 施工進度
  - GanttChart
  - 進度計劃
  - 橫道
  - gantt
agent_created: true
priority: highest
---

# GanttChart Pro v16.0 — 技能文件

> 優先級：**最高**。凡涉及甘特圖、橫道圖、施工進度計劃，一律優先使用本技能。

**v16.0 升級要點**：從 v15.0 交互式腳本架構遷移至 **v3.0 統一渲染引擎**
（gen_gantt.py + config_v2.json + gantt_styles.json + time_utils.py），
配色方案與排版格式參考 **示例學校 示例舞蹈室翻新工程 進度表** 實際交付件校準。

---

## 一、系統架構（v3.0 Unified Rendering Engine）

### 1.1 文件清單

| 文件 | 角色 | 說明 |
|:---|:---|:---|
| `gen_gantt.py` | **主渲染引擎** | `GanttChart` 類，純函數式調用，無 input() |
| `time_utils.py` | 時間模組 | `TimeModule`，日期序列/星期/月份分組/工期計算 |
| `config_v2.json` | **系統配置** | 配色方案(schemes)、日曆、i18n、列印、驗證規則 |
| `gantt_styles.json` | 樣式定義 | 列寬、行高、字體、邊框樣式、凍結窗格 |
| `gantt_config.json` | **任務數據** | JSON 格式任務輸入（project + sections + tasks） |
| `a3_print_layout.py` | A3 橫向排版 | 後處理：重配行高列寬令內容鋪滿 A3 橫向單頁（見 3.9） |

### 1.2 運行方式（非交互式）

```bash
python gen_gantt.py [gantt_config.json] [output.xlsx]
```

- 默認讀取同目錄 `gantt_config.json`
- 默認輸出：`{工程名稱}-施工進度橫道圖.xlsx`
- **零交互**：不需要任何用戶輸入，全部由配置文件驅動

---

## 二、配色方案（從實際 進度表 校準）

### 2.1 blue_pro 方案（默認，推薦）

此方案直接對標 **示例學校施工進度橫道圖** 的實際色彩：

```
┌─────────────────────────────────────────────────────┐
│ 區域              │ 色值       │ 視覺效果           │
├─────────────────────────────────────────────────────┤
│ 標題行背景        │ #1E3B4A    │ 深海軍藍            │
│ 標題行文字        │ #FFFFFF    │ 純白粗體           │
│ 信息區背景        │ #F0F4FA    │ 淺藍灰              │
│ 信息區文字        │ #2C3E50    │ 深炭灰              │
│ 月份標題(奇)      │ #1E3B4A    │ 深海軍藍            │
│ 月份標題(偶)      │ lighten    │ ~#3A5A70 中海軍藍   │
│ 工作日表頭背景    │ #1E3B4A    │ 深海軍藍            │
│ 工作日表頭文字    │ #FFFFFF    │ 白色                │
│ 週末表頭背景      │ #E8E8E8    │ 淺灰                │
│ 週末表頭文字      │ #C0503D    │ 磚紅                │
│ 單元格格線        │ #C5D1DF    │ 淺銀灰              │
│ 粗格線/分隔線     │ #1E3B4A    │ 深海軍藍(medium)    │
│ 今日標記線        │ #D9534F    │ 珊瑚紅              │
│ 今日列背景        │ #FCEFE3    │ 淺珊瑚              │
│ 奇數段行背景      │ #F4F7FB    │ 極淺藍              │
│ 偶數段行背景      │ #FFFFFF    │ 白色                │
│ 關鍵路徑行背景    │ #FCEFE3    │ 淺珊瑚              │
└─────────────────────────────────────────────────────┘
```

### 2.2 Stage 階段配色（橫道條顏色）

按任務 category 字母分配，形成 6 色階段色譜：

| 字母 | 階段類型         | 主色      | 漸變左(亮)    | 漸變右(暗)    | 用途範例               |
|:---|:---|:---|:---|:---|:---|
| A | 開工準備         | `#2E86AB` | `#9BCDE0`    | `#256B8C`    | 交底、材料確認          |
| B | 拆卸及基層處理   | `#81C784` | `#B2DFB2`    | `#5DAE68`    | 拆卸、清運、基層        |
| C | 裝修主體施工     | `#FFD54F` | `#FFE082`    | `#E5B83A`    | 天花、牆身、地台、門窗  |
| D | 電氣及 AV 設備   | `#E57373` | `#EF9A9A`    | `#D45B5B`    | 配線、燈光、音響、AV    |
| E | 安裝及收尾工程   | `#9575CD` | `#C5B9E6`    | `#7E57C2`    | 把杆、家具、除醛、驗收  |
| F | （預留）         | `#64B5F6` | `#A6D4FA`    | `#4A96D4`    | 預留                   |
| G | （預留）         | `#BFCF91` | `#DDE8BD`    | `#A5B97D`    | 預留                   |

> **漲跌色原則**（工程慣例）：橫道色不涉股價，但整體色溫遵循「一文一色系」——
> 以藍色系為骨架（標題/格線/表頭），各 stage 色為點綴，低飽和度、不刺眼。
> 黃色(C)偏琥珀金而非亮黃，紅色(D)偏珊瑚而非警示紅。

### 2.3 其他配色方案

| 方案名 | 風格 | HEADER_BG | 適用場景 |
|:---|:---|:---|:---|
| `green_calm` | 自然舒緩綠 | `#2D5E4F` | 環保/綠建工程 |
| `gray_minimal` | 灰白極簡 | `#3A434C` | 打印友好、黑白輸出 |

---

## 三、排版格式規範（從 進度表 提取）

### 3.1 Excel 結構佈局

```
行1  : ══════════ 標題行（全欄合并）═══════════
       背景=#1E3B4A  文字=白色14pt粗體  行高=38px
       內容：「{工程名}  {副標題}」

行2  : ──── 信息行（全欄合并）─────────────
       背景=#F0F4FA  文字=#2C3E50 9pt  行高=24px
       內容：「工程編號 {id} | 工期：{n}日曆天 | 開工日期... | 承建商：...」

行3  :   ← 間距行（6px，無邊框）

行4  : ┌──────┬────┬───┬───┬───┬[月]─────[月]──┐
       │       │    │序│項 │天 │開│結│ 7月 │ 8月  │  ← 月份行
行5  : │       │    │號│目 │數│始│束│27 29 31│1  3 ...25│  ← 日期數字
行6  : │ info │欄  │  │名 │  │  │  │一二三四六│一二三四六日│  ← 星期
       └──────┴────┴───┴───┴───┴─┴─┴────────┴────────┘
       ↑ 凍結窗格在此分隔(G7 — v16 含 6 個資訊列 A~F：序號/項目名稱/天數/開始/結束/前置任務)

行7+ : [段落標題行] → 任務行 × N → [下一段落] ...
       段落標題：左對齊、10pt粗體、淺底色、底部粗線分隔
       任務行：序號(粗體) / 名稱(左對齊) / 天數 / 開始(MM/dd) / 結束(MM/dd) / [====橫道====]
```

### 3.2 列寬設定

| 列 | 內容 | 寬度(char) |
|:---|:---|:---|
| A | 序號 | 4.2 ~ 5.0 |
| B | 項目名稱 / 施工內容 | 36 ~ 40 |
| C | 天數（工期） | 6.0 |
| D | 開始日期 | 10.5 |
| E | 完成日期 | 10.5 |
| F 起 | 橫道區（每日一格） | **4.5** |

### 3.3 行高設定

| 行類型 | 高度(px) |
|:---|:---|
| 標題行 | 36 ~ 38 |
| 信息行 | 20 ~ 24 |
| 間距行 | 6 |
| 月份行 | 22 |
| 日期行 | 22 |
| 星期行 | 18 |
| 段落標題行 | 26 |
| 任務行 | 26 |

### 3.4 橫道條渲染技術

```
┌────────────────────────────────────────────┐
│  橫道條漸變效果（三段式）：                  │
│                                            │
│  ◇ 左端(第1格)：lighten(color, 0.35)      │
│  │ 中间(第2~N-1格)：color（主色）          │
│  ◇ 右端(最末格)：darken(color, 0.18)      │
│                                            │
│  長條(≥5天)：正中位置疊加白色 7pt 粗體     │
│             任務名稱文字                    │
│                                            │
│  有進度(progress)：已完成部分               │
│             darken(color, 0.12) 覆蓋        │
│                                            │
│  里程碑(milestone)：◆ 菱形符號             │
│             12pt 粗體，使用 task color      │
└────────────────────────────────────────────┘
```

### 3.5 週末處理

| 模式 | 週末列處理 | 節假日處理 |
|:---|:---|:---|
| 日曆天（默認） | `#E8E8E8` 淺灰底 | `#D0D0D0` 中灰底 |
| 工作日模式 | 不顯示週末列 | 同上 |

### 3.6 今日標記（Today Marker）

- **表頭**：今日所在列的日期格 → `#D9534F` 紅底白字；星期格 → `#D9534F` 紅字
- **主體**：今日列所有單元格左邊框改為 `medium` 紅線 (`#D9534F`)
- **關鍵路徑**行：額外以 `#FCEFE3` 淺珊瑚底色標識

### 3.7 圖例與備註

- **圖例區**：位於任務區下方 2 行，每行最多 3 個色塊+標籤
- **備註/關鍵路徑分析**：圖例區再下方，黃底(`#FFF8E1`)深棕字(`#5D4037`)
- **頁眉頁腳**：左=文件名(&F)，中=日期(&D)，右=頁碼(&P)；腳右="Page X of Y"

### 3.8 列印設置

| 參數 | 值 |
|:---|:---|
| 方向 | 橫向 (landscape) |
| 紙張 | A3 (paperSize=8)，自動縮放適配 A4 |
| 縮放 | fitToWidth=1, fitToHeight=0 (不限頁高) |
| 邊距 | 左右 0.4", 上下 0.5", 頁眉腳 0.2" |
| 重複行 | 1:6 (標題+信息+表頭) |
| 格線 | 交付文件關閉（print_options.gridLines=False） |

> ⚠️ **`fit_to_pages_wide` 唔可以係 0**（2026-08-01 修復）。Excel 將 `fitToWidth=0`
> 當「不限頁寬」，`fitToPage` 縮放會完全失效，結果橫向爆到幾頁。必須設 1。
> config_v2.json 內 `print_gridlines` / `print_headings` 舊版引擎從未套用（死設定），
> 已於 `_set_print()` 補上，並新增 `horizontally_centered`。

### 3.9 A3 橫向「單頁鋪滿」排版算法（2026-08-01 內化）

**問題**：引擎預設列寬（甘特 3.0 字符）+ 行高（任務行 47pt）會出現
「寬度只用一半、高度爆兩頁」——因為版面寬高比同紙張唔匹配。

**核心思路**：令**內容寬高比 = 紙張可打印區寬高比**，`fitToPage` 縮放時就會剛好鋪滿單頁。

換算基準（Calibri 11 為標準字體時）：
- 列寬 1 字符 = 7px @96dpi = **1.852 mm**
- 行高 1 pt = **0.3528 mm**

A3 橫向可打印區（邊距 L.59" R.40" T.91" B.71"）：
- 寬 `420 - (0.59+0.40)×25.4` = **394.9 mm**
- 高 `297 - (0.91+0.71)×25.4` = **255.9 mm**，比值 **1.543**

**步驟**：
1. 先壓行高定總高 `H`（任務行 24pt、節段行 26pt、表頭區 130pt、圖例說明區 ~174pt）
2. 算縮放 `s = 255.9 / (H × 0.3528)`
3. 反推目標總列寬 `W = 394.9 / 1.852 / s`（字符）
4. 扣除資訊列（A5 / B62 / C7 / D13 / E13 / F13 = 113），餘數平分俾甘特列
5. 校驗 `(W×1.852) / (H×0.3528) ≈ 1.543` 即成功

**實測案例**（21 條任務 / 18 甘特列）：H=860pt、W=253 字符 →
內容 468.3×303.4mm、比值 1.543、縮放 84.3%、甘特列由 5.6mm 加闊到 12.1mm。

**現成工具**：`a3_print_layout.py`（活體引擎目錄），改 `SRC`/`OUT` 即可套用。
B 列 62 字符 ≈ 32 個中文字單行，任務名唔好超過 32 字，否則要回調行高至 36pt。

---

## 四、gantt_config.json 任務數據格式

### 4.1 完整 JSON 結構

```json
{
  "project": {
    "name": "工程項目名稱",
    "subtitle": "施工進度橫道圖",
    "project_id": "SC26-xxxxxx",
    "contractor": "承建商名稱",
    "start_date": "2026-07-27",
    "end_date": "2026-08-25",
    "num_days": 30,
    "today": "2026-07-17"
  },
  "sections": [
    {
      "title": "段落標題",
      "tasks": [
        {
          "id": 1,
          "name": "任務名稱",
          "duration": 5,
          "start_day": 0,
          "deps": [],
          "category": "A",
          "milestone": false,
          "progress": 0,
          "material": "",
          "remark": ""
        }
      ]
    }
  ],
  "critical_path": [1, 3, 5],
  "legend": [
    {"color": "2E86AB", "label": "開工準備"},
    {"color": "81C784", "label": "拆卸及基層"}
  ],
  ["notes": ["關鍵路徑說明..."]
}
```

### 4.2 任務字段詳解

| 字段 | 類型 | 必填 | 說明 |
|:---|:---|:---|:---|
| id | int | ✓ | 唯一編號 |
| name | string | ✓ | 任務名稱 |
| duration | int | ✓ | 持續天數（含首尾） |
| start_day | int | ✓ | 相對起始日的偏移量（0=第一天） |
| deps | int[] | 前置依賴的 id 列表 |
| category | string | 階段字母(A~G)，決定橫道色 |
| milestone | bool | 是否里程碑（顯示◆符號） |
| progress | int | 0~100 完成百分比 |
| material | string | 材料訂貨期備註 |
| remark | string | 備註 |

### 4.3 自動計算字段（可省略）

- `end_day` = `start_day + duration - 1`（自動推算）
- `start_date`, `end_date` = 根據 project.start_date + day offset 自動轉換
- `critical_path` 未提供時，自動依賴關係計算關鍵路徑

---

## 五、config_v2.json 系統配置要點

```json
{
  "active_scheme": "blue_pro",        // 啟用的配色方案
  "calendar": {
    "type": "calendar_days",          // calendar_days 或 workdays
    "holidays": ["2026-05-01"],       // 公眾假期
    "makeup_days": ["2026-05-09"]     // 補班日
  },
  "gantt_layout": {
    "gantt_column_width": 4.5,        // 橫道區列寬
    "merge_months": true,             // 月份合并顯示
    "column_widths": { ... }          // 各信息列自定義寬度
  },
  "print": {
    "paper_size": 8,                  // 8=A3, 9=A4
    "fit_to_pages_wide": 1
  },
  "dependency": {
    "auto_calculate_critical_path": true
  }
}
```

---

## 六、AI 輔助生成流程（內化生成法）

### 6.1 推薦做法：複製 gen_gantt.py + 配置到工作目錄

```bash
# 1. 複製核心文件到當前工作目錄
cp "源目錄/gen_gantt.py" .
cp "源目錄/time_utils.py" .
cp "源目錄/config_v2.json" .
cp "源目錄/gantt_styles.json" .

# 2. 生成 gantt_config.json（任務數據）
# 3. 直接運行
python gen_gantt.py gantt_config.json output.xlsx
```

### 6.2 內化調用（嵌入 AI 生成腳本）

gen_gantt.py 的 `GanttChart` 類可直接 import 使用：

```python
import sys, os, json
sys.path.insert(0, r"包含 gen_gantt.py 和 time_utils.py 的目錄")
os.chdir(同上)

from gen_gantt import GanttChart

# 加載配置
with open("config_v2.json", encoding="utf-8") as f:
    sys_cfg = json.load(f)
with open("gantt_styles.json", encoding="utf-8") as f:
    sty_cfg = json.load(f)
with open("gantt_config.json", encoding="utf-8") as f:
    task_cfg = json.load(f)

# 生成（非交互，零彈窗）
chart = GanttChart(task_cfg, "output.xlsx", sys_cfg, sty_cfg)
print(f"Done -> {chart.output}")
if chart.warnings:
    for w in chart.warnings:
        print(f"  [!] {w}")
```

### 6.3 關鍵坑位

| 坑位 | 解決方案 |
|:---|:---|
| openpyxl 未安裝 | `pip install openpyxl`（僅需此一依賴） |
| 中文亂碼 | 確保所有 json 讀取指定 `encoding='utf-8'` |
| 輸出被沙箱攔截 | 先寫 temp 目錄，再 PowerShell Copy-Item 搬去桌面 |
| 日期格式 | `_parse_date()` 支持 `YYYY-MM-DD`、Excel 序列號 |
| 月份跨年 | 月份行自動帶年份前綴（如「2026年8月」）|

---

## 七、標準操作決策樹

```
用戶說「生成甘特圖/橫道圖」
    ↓
1. 收集任務數據 → 構建 gantt_config.json
   ├─ 有 Excel/EDF.xlsx → 解析轉為 JSON sections
   ├─ 有文字列表 → 結構化為 JSON
   └─ 有 PDF 進度表 → 參考其配色/排版生成新圖
    ↓
2. 確認工程元數據
   ├── name（工程名）
   ├── start_date / end_date 或 num_days
   ├── contractor（承建商）
   └── today（可選，默認今天）
    ↓
3. 分配 category（A~E）給每個任務
   ├── A = 開工準備
   ├── B = 拆卸/基層
   ├── C = 裝修主體
   ├── D = 電氣/機電/AV
   └── E = 安裝/收尾/驗收
    ↓
4. 選配色方案（未指定 → blue_pro）
    ↓
5. 選日曆模式（未指定 → calendar_days）
    ↓
6. 執行 gen_gantt.py → 輸出 xlsx
    ↓
7. 返回文件路徑
```

---

## 八、從 EDF.xlsx 轉換（兼容舊格式）

若用戶提供的是 EDF.xlsx（舊 v15.0 格式），轉換邏輯：

### 8.1 元數據區提取（前 14 行右側列）

搜索關鍵字：`工程項目名稱`、`工程項目編號`、`承建商`、`編制時間`、`總工期`

### 8.2 表頭識別

| 目標字段 | 譯別關鍵字 |
|:---|:---|
| 序號 | 序號、項次、No、# |
| 施工內容 | 施工內容、任務名稱、Description |
| 開始日期 | 開始日期、Start、Sd、開工 |
| 完成日期 | 完成日期、End、Ed、完工 |

### 8.3 任務類型

- **分類行**：序號含字母如 `A.`、`B.` → section header
- **子任務**：序號如 `A1`、`A.1` → 普通 task，category 取首字母
- **里程碑**：工期=1天且名稱含「里程碑」→ milestone=true

### 8.4 日期格式支援

`YYYY-MM-DD` / `YYYY/MM/DD` / `DD/MM/YYYY` / `YYYY年M月D日` / Excel 序列號

---

## 九、常見問題排除

| 問題 | 原因 | 解決 |
|:---|:---|:---|
| 找不到模組 | 未複製 time_utils.py | 確保 gen_gantt.py 同目錄有 time_utils.py |
| 日期解析失敗 | 非標準日期 | 改用 YYYY-MM-DD |
| 輸出空白 | gantt_config.json 缺少 sections | 檢查 JSON 格式 |
| 顏色不對 | active_scheme 拼寫錯誤 | 檢查 config_v2.json schemes 下是否存在 |
| 橫道太窄/太寬 | gantt_column_width 不合適 | 短工期(<30天)用 4.5~5.0；長工期用 3.5~4.0 |
| 週末沒有灰底 | calendar_type 設成了 workdays | 改為 calendar_days |
| openpyxl 缺失 | 新環境未安裝 | `pip install openpyxl` |

---

## 十、用戶 使用偏好（延續）

| 參數 | 默认值 | 說明 |
|:---|:---|:---|
| 配色 | blue_pro | 6色階段配色 |
| 日曆 | 日曆天 | 除非明確說工作日 |
| 短工期(<60天) | 周模式/逐日 | 每天一列 |
| 長工期(≥60天) | 自動推薦粒度 | 通常 2~7 天/列 |
| 輸出目錄 | 任務資料夾或桌面 | 撞名自動改名 |
| 臨時檔 | 用戶臨時目錄 | 絕不落地其他位置 |
| 工程貨幣 | MOP | 工程 |
| 字體 | 微軟正黑體 | 全文統一 |

---

## 十一、v3.2-v3.4 擴展（自適應 Excel 識別 + 顆粒度模式）

> 本技能係基礎引擎；以下擴展已**內部化**喺本技能目錄 **`scripts/`**（原桌面「橫道圖生成開發專案」活體項目，已於 2026-08-05 複製入此，避免外部路徑失效）。
> （同一套 `gen_gantt.py` + `config_v2.json` + `gantt_styles.json`，加咗 `edf_importer.py` v3.0 + `recognition_config.json` + `gui_gantt.py`；`time_utils.py` 已包含在 scripts/ 目錄）。
> 詳細變更記錄見倉庫 CHANGELOG。

### 11.1 配置驅動自適應 Excel 識別（edf_importer.py v3.0）
- 唔使逐個格式硬編：所有識別字典（field_patterns / metadata / category_keywords / date / duration / predecessor / template_profiles）抽到 **`recognition_config.json`**。
- **新 Excel 格式 → 只係喺 JSON 加關鍵詞就識得**（繁簡中性化 `_FOLD_MAP` + 多格式日期：文本/ datetime / Excel 序列號 + 雙模式前置依賴：行號 `2`/`2SS` 或代碼 `1.2`/`2.1;2.2`）。
- 運行：`python edf_importer.py 進度表.xlsx --report` 睇自動抽取咗咩；`gen_gantt.py` 會自動偵測 `.xlsx` 走 EDF 通道。

### 11.2 五種顆粒度模式（含 auto 自動選擇）
| 模式 | 行為 |
|------|------|
| `auto` | 自動：≤120日→day，121-450日→week，>450日→month |
| `day` | 每日一列，月份分組 |
| `week_grouped` | 每日一列 + 頂部 ISO 週分組欄「第N週 M/D–M/D」 |
| `week` | 每列一週 |
| `compressed` | 多天一列（**列寬驅動**：`col_width = panel_width / cols`，可讀 [2.0,15.0]pt 內揾最佳 nice_step，強制最小 2 日/列） |
| `month` | 每列一月 |

- A4 橫向面板 ~120pt / A3 橫向 ~180pt（`config_v2.json` 嘅 `granularity.recommend.panel_widths`）。
- CLI：`--granularity {auto,day,week_grouped,week,compressed,month}` / `--days-per-col N` / `--suggest`（只列印 A4/A3 最優日/列建議，唔生成）。
- ⚠️ **千祈唔好用舊「紙型容量減法」** `ceil(天數/容量列數)` —— 55 天項目會算出 1 日/列 = 零壓縮。必須用列寬驅動。

### 11.3 GUI headless 端到端測試
- sandbox 入面 `tkinter.Tk()` 可實例化 → 直接構建 `GanttGUI` + 叫 `_validate()` + 同步 `_run_generate()`（跳過 mainloop/thread/messagebox/os.startfile）做全鏈路驗證。
- `import gui_gantt` 前必須 `sys.path.insert(0, SCRIPT_DIR)`。

---

## 十二、版本記錄

| 版本 | 說明 |
|:---|:---|
| **v16.0** | **遷移至 v3.0 統一渲染引擎；配色/排版從示例學校 進度表 校準；新增 JSON 數據驅動流程；完整重寫 SKILL.md** |
| v15.0 | 舊版交互式腳本（GanttChart Pro v15.0.py + input()問答） |

---

## 十三、v16.1 修復記錄（2026-07-23 實測）

示例機房搬遷工程實測中發現 2 處與「GanttChart Pro 標準」（用戶要求 A3 橫向 + 每頁重複標題 1:6）不符，已烤入本技能 `scripts/gen_gantt.py` 的 `_set_print`：

| # | 問題 | 修復 | 狀態 |
|:---|:---|:---|:---|
| 1 | `_set_print` 在 `gantt_cols ≤ 60` 時**無視** config 的 `print.paper_size`，強制落 A4 | 優先尊重 `config_v2.json` 的 `print.paper_size`（8=A3）；僅當未設定時才自動推薦 | ✅ |
| 2 | 引擎**從未設定** `print_title_rows`，每頁不重複標題 | 永遠 `print_title_rows = "1:6"` | ✅ |

修復後實測（活體引擎直接重跑，無副本 patch）：`paperSize=8(A3)`、`print_title_rows='$1:$6'`、`landscape`、`fitToWidth=1`、`freeze G7` 全部原生生效；23 條橫道位置與日期對齊無倒退。

> ⚠️ **文檔與代碼一致性提醒**：本 SKILL.md 3.1 原寫凍結 `F7`（v15 五列慣例），v16 含 6 個資訊列實為 `G7`，已修正。類別配色實際渲染的漸變亮/暗端與 2.2 表格略有偏差（如 C 類別實際用 `FFE38C`/`D1AE40` 而非文檔 `FFE082`/`E5B83A`），屬文檔微飄，不影響輸出。

> 📌 **備份策略**（見 `\備忘.md`）：
> - **現行主力**：本技能 v16（引擎 = 本技能目錄 `scripts/gen_gantt.py`）
> - **備用主力**：若 v16 損壞，改用桌面 `舊版運行系統`
> - **已歸檔不用**：桌面 `舊版生成器`、v16 的 `` 子文件夾（已複製至備忘目錄）
