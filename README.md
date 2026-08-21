<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8%2B-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/License-MIT-yellow" alt="License">
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey" alt="Platform">
</p>

# GanttChart Pro

> **Professional Gantt charts in Excel — no MS Project required.**

Generate beautiful, print-ready construction schedules with pure Python + openpyxl.
Dual calendar modes, WBS auto-grouping, milestone markers, dependency support, and a
three-level timescale — everything you'd expect from MS Project, without the license or
complexity.

> 💡 See `examples/EDF.example.xlsx` and `scripts/gantt_config.example.json` for ready-to-run
> sample data — run `gen_gantt.py` to render a real `.xlsx` chart in seconds.

---

## 🎯 Why GanttChart Pro?

| | MS Project | GanttChart Pro |
|---|---|---|
| **Cost** | $$$ license | Free (MIT) |
| **Setup** | Complex install | `pip install openpyxl` |
| **Sharing** | .mpp files only | .xlsx (anyone can open) |
| **Automation** | COM hack | Native Python API |
| **Cross-platform** | Windows only | Windows / macOS / Linux |

---

## 🚀 Quick Start

```bash
git clone https://github.com/David-CB666/gantt-chart-pro.git
cd gantt-chart-pro
pip install openpyxl
```

### Generate a Chart (one command)

```bash
cd scripts
python gen_gantt.py ../scripts/gantt_config.example.json ../output.xlsx
```

Or simply drop your own `gantt_config.json` next to `gen_gantt.py` and run:

```bash
python gen_gantt.py
```

---

## 📦 What's Inside (v16.0)

| File | Role |
|---|---|
| `scripts/gen_gantt.py` | **Main rendering engine** — `GanttChart` class, data-driven, zero interaction |
| `scripts/time_utils.py` | Date series / weekday / month grouping / duration math |
| `scripts/edf_importer.py` | Config-driven EDF.xlsx → task JSON importer (v3.0) |
| `scripts/gui_gantt.py` | Optional headless GUI for end-to-end validation |
| `scripts/a3_print_layout.py` | A3 landscape single-page post-processing |
| `scripts/config_v2.json` | Color schemes, calendar, i18n, print & validation rules |
| `scripts/gantt_styles.json` | Column widths, row heights, fonts, borders |
| `scripts/recognition_config.json` | EDF field-pattern dictionary (extend without code changes) |
| `scripts/gantt_config.example.json` | Generic sample task data (works out of the box) |
| `SKILL.md` | Full skill documentation (engine architecture, color schemes, layout spec) |

---

## 🎨 Color Schemes

- **`blue_pro`** (default) — navy header/lines, 6-stage gradient bars, coral "today" line
- **`green_calm`** — natural green, eco/buildings friendly
- **`gray_minimal`** — grayscale, print-friendly B/W output

Stage colors (A–E): prep → demolition → fit-out → electrical/AV → install/handover.

---

## 🗓️ Calendar & Granularity

- **`calendar_days`** (default) — every day a column, weekends shaded
- **`workdays`** — hide weekends
- Granularity modes: `auto` / `day` / `week_grouped` / `week` / `compressed` / `month`
  - `auto`: ≤120d → day, 121–450d → week, >450d → month

---

## 📥 Import from EDF.xlsx

If you already have an Excel schedule in EDF format:

```bash
python edf_importer.py 你的進度表.xlsx --report   # preview what gets extracted
python gen_gantt.py 你的進度表.xlsx output.xlsx    # gen_gantt auto-detects .xlsx
```

All recognition patterns live in `recognition_config.json` — add a new column keyword
there and the importer understands your format, no code changes needed.

---

## 📄 License

MIT — see `LICENSE`. Free for commercial and personal use.
