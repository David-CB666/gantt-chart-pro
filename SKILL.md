---
name: gantt-chart-pro
description: >
  Excel 橫道圖生成器。無需 MS Project，純 openpyxl 生成專業甘特圖。
  支持日曆天/工作天、多種配色方案、WBS 分組、里程碑標記、周末高亮、今日線。
  觸發詞：生成甘特圖、橫道圖、施工進度表、進度表、Gantt Chart。
version: 1.0.0
icon: 📊
author: David-CB666
agent_created: true
metadata:
  clawdbot:
    requires:
      bins:
        - python
      packages:
        - openpyxl
    commands:
      - name: generate
        description: 從 EDF.xlsx 生成甘特圖 Excel
        usage: python scripts/gantt_chart_pro.py --input EDF.xlsx --output schedule.xlsx
---

# GanttChart Pro v15.0 — Excel 橫道圖生成器

> **優勢**：無需 MS Project，純 Python + openpyxl 生成專業甘特圖 Excel

## 核心功能

| 功能 | 說明 |
|------|------|
| **雙日曆模式** | 日曆天（含周末）/ 工作天（排除周末假期） |
| **多配色方案** | 專業藍調 / 自然舒緩綠 / 灰白極簡 |
| **WBS 分組** | 根據任務編號自動分組（A=1級, A1=2級） |
| **里程碑標記** | 根據關鍵詞或工期=0自動標記 |
| **周末高亮** | 周末列淺灰背景，工作天模式深灰 |
| **今日線** | 當前日期紅色垂直線 |
| **前置關係** | 支持 FS/SS/FF/SF 四種關係 |
| **時間刻度** | 年/月/日 三級刻度自動適應 |
| **打印優化** | 橫向打印、自適應頁寬、標題行重複 |

## 前置條件

```bash
pip install openpyxl
```

## 快速開始

### 1. 準備 EDF.xlsx

**必需列**：
- **施工內容** — 任務名稱
- **開始日期** — 格式：YYYY-MM-DD 或 Excel 日期

**可選列**：
- **序號** — 任務編號（自動推斷 WBS 層級）
- **工期** — 天數（默認 1 天）
- **完成日期** — 自動計算或手動指定
- **前置任務** — 如 `A1FS` 或 `A1`（默認 FS）

**示例**（`examples/EDF.xlsx`）：

| 序號 | 施工內容 | 開始日期 | 工期 | 完成日期 | 前置任務 |
|------|----------|----------|------|----------|----------|
| A | 準備階段 | 2026-06-01 | 5 | 2026-06-05 | |
| A1 | 材料訂購 | 2026-06-01 | 3 | 2026-06-03 | |
| A2 | 人員進場 | 2026-06-02 | 2 | 2026-06-03 | A1 |
| B | 施工階段 | 2026-06-06 | 10 | 2026-06-15 | A |
| B1 | 結構施工 | 2026-06-06 | 5 | 2026-06-10 | |
| B2 | 機電安裝 | 2026-06-11 | 5 | 2026-06-15 | B1 |

### 2. 運行生成器

```bash
# 交互式運行（會詢問日曆模式）
python scripts/gantt_chart_pro.py

# 指定輸入輸出
python scripts/gantt_chart_pro.py --input EDF.xlsx --output schedule.xlsx

# 指定日曆模式
python scripts/gantt_chart_pro.py --input EDF.xlsx --calendar workdays
```

### 3. 查看結果

生成的 Excel 包含：
- **標題區** — 工程名稱、承建商、施工期間
- **任務表** — 序號、施工內容、開始、天數、完成
- **甘特圖** — 彩色橫道 + 時間刻度 + 周末高亮 + 今日線

## 配置文件

### config_v2.json

控制配色、日曆、輸出格式等。

```json
{
  "active_scheme": "blue_pro",
  "schemes": {
    "blue_pro": {
      "name": "專業藍調",
      "colors": { "HEADER_BG": "1E3B4A", ... },
      "stage_colors": { "A": "2E86AB", "B": "81C784", ... }
    }
  },
  "calendar": {
    "type": "ask_user",
    "options": ["calendar_days", "workdays"],
    "holidays": ["2026-05-01", "2026-12-20"]
  }
}
```

### gantt_styles.json

控制列寬、行高、字體、邊框等。

```json
{
  "column_widths": { "A": 5.0, "B": 30.0, "C": 10.0 },
  "row_heights": { "title_row": 36, "header_row": 30 },
  "fonts": { "default_name": "微軟正黑體", "sizes": { "title": 16 } }
}
```

## 三種配色方案

| 方案 | 名稱 | 適用場景 |
|------|------|----------|
| `blue_pro` | 專業藍調 | 工程報告、正式文件 |
| `green_calm` | 自然舒緩綠 | 環保項目、園林工程 |
| `gray_minimal` | 灰白極簡 | 打印輸出、黑白複印 |

## 高級功能

### WBS 自動分組

任務編號格式決定層級：
- `A`, `B`, `C` → 一級（階段）
- `A1`, `A2`, `B1` → 二級（工序）
- `A1.1`, `A1.2` → 三級（細項）

### 前置關係語法

| 格式 | 說明 |
|------|------|
| `A1` | FS（完成-開始），默認 |
| `A1FS` | 明確指定 FS |
| `A1SS` | 開始-開始 |
| `A1FF` | 完成-完成 |
| `A1SF` | 開始-完成 |
| `A1+2` | FS + 2 天延隔 |
| `A1-3` | FS - 3 天提前 |

### 里程碑自動標記

關鍵詞：`開工`、`竣工`、`驗收`、`移交`、`完成`

或工期 = 0 天自動標記為里程碑（菱形符號）

## 目錄結構

```
gantt-chart-pro/
├── SKILL.md                    # 技能說明（本文件）
├── scripts/
│   ├── gantt_chart_pro.py      # 核心腳本
│   └── time_utils.py           # 時間處理工具
├── examples/
│   └── EDF.xlsx                # 示例任務數據
├── references/
│   ├── config-guide.md         # 配置文件詳解
│   └── color-schemes.md        # 配色方案說明
├── config_v2.json              # 主配置文件
├── gantt_styles.json           # 樣式配置文件
└── README.md                   # GitHub README
```

## 與 msp-automation 的區別

| 特性 | gantt-chart-pro | msp-automation |
|------|-----------------|----------------|
| **依賴** | openpyxl | MS Project 2022/2024 |
| **輸出** | Excel (.xlsx) | MS Project (.mpp) |
| **費用** | 免費 | 需購買 MS Project |
| **功能** | 甘特圖生成 | 完整項目管理 |
| **適用** | 快速生成、無需 MS Project | 專業項目管理 |

## 授權

MIT License

## 作者

David-CB666 (Mike) — 澳門機電工程項目經理

## 版本歷史

- **v1.0.0** (2026-06-02) — 初始版本，從 v15.0 遷移
- **v15.0** (2026-05-12) — 原始版本（VBA 配色優化）
