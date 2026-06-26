<p align="center">
  <img src="https://img.shields.io/badge/Python-3.8%2B-blue?logo=python" alt="Python">
  <img src="https://img.shields.io/badge/License-MIT-yellow" alt="License">
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey" alt="Platform">
</p>

# GanttChart Pro

> **Professional Gantt charts in Excel — no MS Project required.**

<p align="center">
  <img src="demo/demo_preview.png" alt="GanttChart Pro Demo" width="800">
</p>

Generate beautiful, print-ready construction schedules with pure Python + openpyxl.

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
python scripts/gantt_chart_pro.py --input examples/EDF.xlsx --output schedule.xlsx --scheme blue_pro
```

### From Python

```python
from scripts.gantt_chart_pro import GanttGenerator

gen = GanttGenerator(
    input_file="examples/EDF.xlsx",
    output_file="schedule.xlsx",
    color_scheme="blue_pro",
    calendar_mode="calendar_days"
)
gen.generate()
```

---

## 📊 Features

- 🗓️ **Dual calendar modes** — calendar days or working days
- 🎨 **Multiple color schemes** — blue_pro, green_calm, gray_minimal
- 📊 **WBS auto-grouping** — task codes define hierarchy (A=Level 1, A1=Level 2)
- 🚩 **Milestone markers** — auto-detected or manual
- 📏 **Weekend highlight** — gray background for non-working days
- 🔴 **Today line** — red vertical line at current date
- 🔗 **Dependency arrows** — FS/SS/FF/SF relationships
- 📐 **Three-level timescale** — year / month / day
- 🖨️ **Print-ready** — landscape, auto-fit, repeating headers

---

## 📋 EDF Data Format

Prepare an Excel file with these columns:

| Column | Required | Description |
|--------|----------|-------------|
| **Task ID** | Recommended | A, A1, A2... (auto-infers hierarchy) |
| **Task Name** | Yes | Description |
| **Start Date** | Yes | YYYY-MM-DD |
| **Duration** | No | Days (default: 1) |
| **End Date** | No | Auto-calculated |
| **Predecessor** | No | e.g. "A1FS" or "A1" |

See [`examples/EDF.xlsx`](examples/EDF.xlsx) for a working example.

---

## 🎨 Color Schemes

| Scheme | Vibe | Best For |
|--------|------|----------|
| `blue_pro` | Professional navy | Client presentations |
| `green_calm` | Soft green | Internal planning |
| `gray_minimal` | Clean grayscale | Print/black & white |

Full customization: [`gantt_styles.json`](gantt_styles.json)

---

## 📖 Documentation

- **Configuration Guide**: [`references/config-guide.md`](references/config-guide.md)
- **Color Scheme Reference**: [`references/color-schemes.md`](references/color-schemes.md)

---

## 📄 License

MIT © [David-CB666](https://github.com/David-CB666)
