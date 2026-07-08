<div align="center">

<img src="exp_logo.png" width="120" alt="DeltaV FHX Nameset Editor">

# DeltaV FHX Nameset Editor

[![GitHub stars](https://img.shields.io/github/stars/mobosa/DeltaV-FHX-Nameset-Editor?style=social)](https://github.com/mobosa/DeltaV-FHX-Nameset-Editor/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/mobosa/DeltaV-FHX-Nameset-Editor?style=social)](https://github.com/mobosa/DeltaV-FHX-Nameset-Editor/network/members)
[![GitHub issues](https://img.shields.io/github/issues/mobosa/DeltaV-FHX-Nameset-Editor)](https://github.com/mobosa/DeltaV-FHX-Nameset-Editor/issues)
[![GitHub license](https://img.shields.io/github/license/mobosa/DeltaV-FHX-Nameset-Editor)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)](https://www.python.org/)

**Batch translate/edit nameset values in DeltaV FHX configuration files**

[English](README.md) | [简体中文](README_zh-CN.md)

</div>

---

## 🔍 Overview

DeltaV FHX Nameset Editor is a tool designed for **Emerson DeltaV system engineers** to batch translate and edit nameset values in FHX configuration files. It resolves character encoding errors during DeltaV import by syncing nameset definitions between FHX files.

> 💡 Supports any FHX type (Library, Control Strategies, Setup, Recipes, etc.) with automatic detection.

## ✨ Key Features

| Feature | Description |
|---------|-------------|
| **ENUMERATION_SET Translation** | Batch replace value names like `VALUE=1 NAME="Stopped"` |
| **STRING_VALUE Reference Translation** | Replace field references in `SET="..." STRING_VALUE="value"` |
| **Expression Reference Translation** | Replace values in expressions like `'$phase_state:Held'` |
| **Alarm Block Translation** | SYSTEM_ALARM / USER_ALARM: DESCRIPTION, ALARM_WORD, MESSAGE, CATEGORY |
| **Priority Name Translation** | 危急→CRITICAL, 警告→WARNING, 建议→ADVISORY, 记录→LOG |
| **LOCALE Replacement** | Auto-replace locale strings based on New Database |
| **New Nameset Support** | Add complete ENUMERATION_SET definition blocks |
| **Built-in Mappings** | Auto-translate 20+ standard DeltaV namesets (`$phase_state`, `$recipe_state`, etc.) |

## 🚀 Quick Start

### Download Pre-built Executable (Recommended)

1. Download `DeltaV_FHX_Nameset_Editor.exe` from [Releases](https://github.com/mobosa/DeltaV-FHX-Nameset-Editor/releases)
2. Double-click to run — no Python installation required

### Run from Source

```bash
pip install openpyxl customtkinter
python fhx_migrator.py
```

## 📖 Workflow

```
┌─────────────────────────────────────────────────────────────┐
│  Step 1: Compare Original Database with New Database        │
│          → Export Excel (auto-fill suggested values)        │
└─────────────────────────────────────────────────────────────┘
                            ↓
         User reviews/edits "New Value" column in Excel
                            ↓
┌─────────────────────────────────────────────────────────────┐
│  Step 2: Read Excel → Generate new FHX file (suffix _NEW)   │
└─────────────────────────────────────────────────────────────┘
```

## 💻 CLI Mode

```bash
# Compare and export Excel
python fhx_migrator.py compare <Original FHX> --setup <New Database> [-o output.xlsx]

# Generate new FHX from edited Excel
python fhx_migrator.py generate <Original FHX> --setup <New Database> --excel edited.xlsx [-o output.fhx]
```

| Argument | Description |
|----------|-------------|
| `compare` | Compare Original Database with New Database, export Excel |
| `generate` | Read Excel and generate new FHX |
| `--setup` | **Required.** New Database reference file path |
| `--excel` | **Required for generate.** Edited Excel file path |
| `-o, --output` | Optional. Output file path |

## 📊 Excel Output Format

The exported Excel contains 5 worksheets:

| Sheet | Content | Status Values |
|-------|---------|---------------|
| **Namesets** | ENUMERATION_SET definitions comparison | `Both` / `Original only` |
| **String Values** | STRING_VALUE reference comparison | — |
| **Expression Refs** | Expression reference comparison | — |
| **Alarm Types** | SYSTEM_ALARM / USER_ALARM fields | `Both` / `Old only` / `New only` |
| **Alarm Priorities** | PRIORITY_NAME mapping | — |

**New Value column rules:**
- **Simple format**: Write name directly, e.g. `STOP`, `Running` → auto-converts to `VALUE=1 NAME="STOP"`
- **Full format**: `VALUE=1 NAME="STOP"`
- **Leave empty**: Skip this row

> 🔵 Blue background = auto-suggestion available | 🟡 Yellow background = manual input required

## ⚠️ Important Notes

- **Alarm Modification Limits**: Not all Alarm entries support direct modification. **Recommendation: Do not modify Alarm fields in project-exported FHX files.** Export Alarms separately if needed.
- **Nameset Not Found**: Different FHX export sources may have structural differences. Refer to `Normal Namesets` attachment if needed.
- **Re-import Issues**: Alarm properties won't be overwritten on re-import. Test on an empty database first.
- **Large Projects**: Import in batches to monitor errors and avoid overwriting.

## 🛠️ Build Executable

```bash
pip install pyinstaller
pyinstaller FHX_Migration_Tool.spec --noconfirm
```

Generated exe is in the `dist/` directory.

## 📁 Project Structure

```
DeltaV_FHX_Nameset_Editor/
├── fhx_core.py              # Backend logic (parsing, comparison, generation, Excel I/O)
├── fhx_migrator.py          # GUI + CLI entry point
├── test_fhxCoverage.py      # Unit tests
├── FHX_Migration_Tool.spec  # PyInstaller build config
├── exp_logo.ico             # App icon (ICO)
├── exp_logo.png             # App icon (PNG)
├── requirements.txt         # Python dependencies
├── README.md                # This file (English)
└── README_zh-CN.md          # Chinese documentation
```

## 📋 Dependencies

| Package | Purpose |
|---------|---------|
| `openpyxl` | Excel read/write |
| `customtkinter` | Modern GUI framework |
| Python 3.8+ | Runtime (source code only) |

## 🧪 Testing

```bash
pip install pytest
python -m pytest test_fhxCoverage.py -v
```

## 📄 License

This project is licensed under the MIT License.

## 👤 Author

**Jared.Ji** — Jared.Ji@emerson.com

---

<div align="center">

**Made with ❤️ for Emerson DeltaV Engineers**

</div>
