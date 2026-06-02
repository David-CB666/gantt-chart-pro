# GanttChart Pro — Excel 橫道圖生成器

[![OpenClaw Skill](https://img.shields.io/badge/OpenClaw-Skill-blue)](https://clawd.org.cn/skills/david-cb666/gantt-chart-pro)
[![Python](https://img.shields.io/badge/Python-3.8%2B-green)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

> **無需 MS Project，純 Python + openpyxl 生成專業甘特圖 Excel**

---

## 為什麼選擇 GanttChart Pro？

✅ **零成本** — 無需購買 MS Project 許可證
✅ **輕量級** — 僅依賴 openpyxl，安裝簡單
✅ **專業輸出** — 多配色方案、WBS 分組、里程碑標記
✅ **靈活配置** — JSON 配置文件，支持自定義樣式
✅ **跨平台** — Windows/macOS/Linux 均可運行

---

## 快速開始

### 安裝依賴

```bash
pip install openpyxl
```

### 準備數據

創建 `EDF.xlsx`，必需列：
- **施工內容** — 任務名稱
- **開始日期** — YYYY-MM-DD 格式

可選列：序號、工期、完成日期、前置任務

### 生成甘特圖

```bash
python scripts/gantt_chart_pro.py --input EDF.xlsx --output schedule.xlsx
```

### 查看結果

打開 `schedule.xlsx`，包含：
- 標題區（工程名稱、承建商、施工期間）
- 任務表（序號、施工內容、開始、天數、完成）
- 甘特圖（彩色橫道 + 時間刻度 + 周末高亮）

---

## 功能特性

### 🎨 三種專業配色

| 方案 | 適用場景 |
|------|----------|
| 專業藍調 | 工程報告、正式文件 |
| 自然舒緩綠 | 環保項目、園林工程 |
| 灰白極簡 | 打印輸出、黑白複印 |

### 📅 雙日曆模式

- **日曆天** — 包含周末與節假日
- **工作天** — 僅計算工作日，周末灰顯

### 📊 WBS 自動分組

根據任務編號自動推斷層級：
- `A`, `B`, `C` → 一級（階段）
- `A1`, `A2` → 二級（工序）
- `A1.1`, `A1.2` → 三級（細項）

### ⚙️ 前置關係支持

FS/SS/FF/SF 四種關係，支持延隔天數

---

## 配置文件

### config_v2.json

```json
{
  "active_scheme": "blue_pro",
  "calendar": {
    "type": "workdays",
    "holidays": ["2026-05-01", "2026-12-20"]
  }
}
```

### gantt_styles.json

```json
{
  "column_widths": { "A": 5.0, "B": 30.0 },
  "fonts": { "default_name": "微軟正黑體" }
}
```

---

## 與 MS Project 的區別

| 特性 | GanttChart Pro | MS Project |
|------|----------------|------------|
| 成本 | 免費 | 需購買許可 |
| 依賴 | openpyxl | MS Office |
| 功能 | 甘特圖生成 | 完整項目管理 |
| 學習曲線 | 簡單 | 複雜 |

---

## 適用場景

✅ 快速生成施工進度表
✅ 無 MS Project 許可證
✅ 批量生成多個甘特圖
✅ 自動化腳本集成

---

## 授權

MIT License

---

## 作者

**David-CB666**
澳門機電工程項目經理

---

## 版本歷史

- **v1.0.0** (2026-06-02) — OpenClaw 技能包初始版本
- **v15.0** (2026-05-12) — 原始 Python 版本
