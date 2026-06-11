# neware-batch-export

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)](https://github.com/AtomRearch/neware-batch-export)
[![Release](https://img.shields.io/github/v/release/AtomRearch/neware-batch-export)](https://github.com/AtomRearch/neware-batch-export/releases)
[![NEWARE Developer Program](https://img.shields.io/badge/NEWARE-Developer%20Program-orange)](https://github.com/NEWARE-Tech/neware-official/discussions/2)

**Batch-export Neware `.ndax` files to complete, full-detail `.xlsx` — with GUI, cycle-mode display, parallel processing, and email notifications.**

> 📖 [中文说明 README_zh.md](README_zh.md)

> Developed as part of the [NEWARE Developer Program](https://github.com/NEWARE-Tech/neware-official/discussions/2).

---

## The Problem

Neware BTS 8.0 has a fundamental data quality gap between its two export modes:

| Mode | Sheets exported | Suitability |
|------|----------------|-------------|
| **Single-file export** (manual) | 8 full sheets — unit / test / cycle / **step** / **record** / log / idle / curve | ✅ Complete |
| **Built-in batch export** | Truncated — **step and record layers missing** | ❌ Incomplete |

The `step` and `record` sheets are the most important for machine learning and data-driven battery research — they contain the full time-series data at every recorded point. Without them, you cannot reconstruct charge/discharge curves, compute features, or build reliable ML models.

Yet the manual single-file export doesn't scale: clicking through hundreds of `.ndax` files one by one is not an option when you need to build a large dataset.

**This creates a bottleneck that affects every lab running AI/ML battery research with Neware equipment.** You either get volume or completeness — not both.

---

## The Solution

`BTSDAExReport.exe` (bundled with BTS 8.0) exposes an undocumented CLI:

```
BTSDAExReport.exe export custom "<input.ndax>" "<output.xlsx>"
```

The `custom` mode reads your saved `BTSDAConfig.iconf` and produces **the identical 8-sheet output as a manual single-file export** — including full step and record layers. We verified this by comparing row counts per sheet between manual and automated exports on real experimental data: they are identical.

This project wraps that CLI into a GUI tool with batch processing, parallel execution, and additional workflow features.

---

## Features

- ✅ **Full 8-sheet export** — identical to manual single-file export (step/record layers complete)
- ✅ **Cycle mode display** — shows the cycle-statistics mode currently configured in BTS (Step Default / Chg→Dchg / Dchg→Chg / Custom Step); read-only — this tool never modifies BTS configuration
- ✅ **Parallel processing** — configurable worker count (default: 3)
- ✅ **EN / 中文 UI toggle** — one-click language switch
- ✅ **Flexible path input** — paste single or multiple `r"..."` Python-style paths; `#`-prefixed lines auto-skipped
- ✅ **Folder scan** — browse directory, select `.ndax` files from a checklist
- ✅ **Custom output directory** — same folder as source, or a single destination
- ✅ **Skip existing** — skip files where `.xlsx` already exists (default on)
- ✅ **Summary CSV** — auto-generated after each batch (filename / cycles / size / elapsed time)
- ✅ **Email report** — optional post-export notification with summary table + CSV attachment
- ✅ **History** — last-used paths and settings restored on next launch
- ✅ **BTSDA conflict check** — warns if BTSDA.exe is running, which may interfere with the export

---

## Requirements

- Windows (BTSDA is Windows-only)
- Neware BTS 8.0 software installed
- Python ≥ 3.10
- [`customtkinter`](https://github.com/TomSchimansky/CustomTkinter)

```bash
pip install customtkinter
```

---

## Installation

1. Clone or download this repository
2. Install dependencies: `pip install customtkinter`
3. Configure paths (see [Configuration](#configuration))
4. Double-click **`neware-export.bat`** to launch

---

## Usage

### GUI

```bash
python neware_export_gui.py
```

or double-click `neware-export.bat`.

**Workflow:**
1. The header shows the **cycle mode currently configured in BTS** (read-only). If you need a different mode, set it in the BTS software first, then return here.
2. Paste `.ndax` paths — supports raw Python list format with inline comments:
   ```python
   r"E:\data\experiment_01.ndax",
   r"E:\data\experiment_02.ndax",
   # r"E:\data\skip_this.ndax",
   r"E:\data\experiment_03.ndax",
   ```
3. Or click **Scan Folder** to browse a directory and select files
4. Set output directory
5. Click **Export**

### PowerShell (headless)

```powershell
.\neware_batch_export.ps1 -InputDir "E:\path\to\ndax"
.\neware_batch_export.ps1 -InputDir "E:\data" -OutputDir "E:\xlsx" -Recurse -Force
```

---

## Configuration

Open `neware_export_gui.py` and edit the constants near the top:

```python
# ── Paths — adjust to match your BTS installation ──────────────────
BTSDA_EXE  = r"E:\software\BTSClient80\BTSDAExReport.exe"
ICONF_PATH = r"C:\Users\<you>\Documents\NEWARE\BTSClient\BTSDAConfig.iconf"
```

`ICONF_PATH` is only ever **read** — to display the cycle mode currently configured in BTS. This tool never writes to BTS configuration files.

### Email notifications

Email is **disabled by default**. To enable it, check *Email report on finish* in the GUI, then configure the SMTP section in `neware_export_gui.py`:

```python
SMTP_SERVER     = "smtp.qq.com"         # Change to your provider
SMTP_PORT       = 465
SENDER_EMAIL    = "your@email.com"
SENDER_PASSWORD = "your_smtp_auth_code" # App password / SMTP auth code
RECEIVER_EMAIL  = "recipient@email.com"
```

Common SMTP providers:

| Provider | Server | Port |
|----------|--------|------|
| QQ Mail  | `smtp.qq.com` | 465 |
| 163 Mail | `smtp.163.com` | 465 |
| Gmail    | `smtp.gmail.com` | 465 |
| Outlook  | `smtp.office365.com` | 587 |

> ⚠️ **Never commit real credentials.** Store them in environment variables or a local config file not tracked by git.

---

## Cycle Mode

Controls how BTS counts charge/discharge capacity per cycle. This is configured **in the BTS software itself** (`BTSDAConfig.iconf`) — the GUI displays the current value read-only and never changes it.

| Mode | Value | Description |
|------|-------|-------------|
| Step Default | 0 | Default step assignment (recommended) |
| Chg→Dchg | 1 | Charge-first cycle counting |
| Dchg→Chg | 2 | Discharge-first cycle counting |
| Custom Step | 3 | User-defined starting step |

To use a different mode, change it in BTS before exporting.

---

## Output sheets

| Sheet | Contents |
|-------|----------|
| `unit` | Unit settings |
| `test` | Channel / test metadata |
| `cycle` | Per-cycle statistics (capacity, CE, energy, DCIR…) |
| `step` | Per-step statistics |
| `record` | **Full time-series** — every recorded data point |
| `log` | Event log |
| `idle` | Rest segment data |
| `curve` | Computed curve data |

The `record` sheet is the primary source for ML feature extraction.

---

## Relationship to aurora-neware

[`aurora-neware`](https://github.com/EmpaEconversion/aurora-neware) controls Neware cyclers **during** experiments (start/stop channels, monitor status). This project handles **post-experiment data export**. They are complementary:

```
Run experiment → [aurora-neware: monitor & control]
                          ↓
                 Experiment complete
                          ↓
         [neware-batch-export: full .xlsx dataset]
                          ↓
                  ML / data analysis
```

---

## Verified on

- BTS 8.0 (`BTSDAExReport.exe` version 2024.06.24)
- Windows 10 / 11
- CT-series and BTS-series cyclers

---

## License

MIT

---
