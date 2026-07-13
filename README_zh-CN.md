<div align="center">

# DeltaV FHX Nameset Editor

[![GitHub stars](https://img.shields.io/github/stars/mobosa/DeltaV-FHX-Nameset-Editor?style=social)](https://github.com/mobosa/DeltaV-FHX-Nameset-Editor/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/mobosa/DeltaV-FHX-Nameset-Editor?style=social)](https://github.com/mobosa/DeltaV-FHX-Nameset-Editor/network/members)
[![GitHub issues](https://img.shields.io/github/issues/mobosa/DeltaV-FHX-Nameset-Editor)](https://github.com/mobosa/DeltaV-FHX-Nameset-Editor/issues)
[![GitHub license](https://img.shields.io/github/license/mobosa/DeltaV-FHX-Nameset-Editor)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.8+-blue?logo=python&logoColor=white)](https://www.python.org/)

**批量翻译/编辑 DeltaV FHX 配置文件中的 nameset 值**

[English](README.md) | [简体中文](README_zh-CN.md)

</div>

---

## 🔍 概述

DeltaV FHX Nameset Editor 是一款专为 **Emerson DeltaV 系统工程师** 设计的工具，用于批量翻译和编辑 FHX 配置文件中的 nameset 值。通过同步 FHX 文件间的 nameset 定义，解决 DeltaV 导入时的字符编码错误。

> 💡 支持任何 FHX 类型（Library、Control Strategies、Setup、Recipes 等），自动识别处理。

## ✨ 功能特性

| 功能 | 说明 |
|------|------|
| **ENUMERATION_SET 翻译** | 批量替换 `VALUE=1 NAME="停止中"` 类型的值名称 |
| **STRING_VALUE 引用翻译** | 替换 `SET="..." STRING_VALUE="值"` 中的引用值 |
| **表达式引用翻译** | 替换 `'$phase_state:正在保持'` 等表达式中的值 |
| **报警块翻译** | SYSTEM_ALARM / USER_ALARM：DESCRIPTION、ALARM_WORD、MESSAGE、CATEGORY |
| **优先级名称翻译** | 危急→CRITICAL、警告→WARNING、建议→ADVISORY、记录→LOG |
| **LOCALE 替换** | 根据 New Database 自动替换 locale 字符串 |
| **新 nameset 添加** | 支持在 FHX 中新增完整的 ENUMERATION_SET 定义块 |
| **内建映射** | 自动翻译 20+ 套标准 DeltaV nameset（`$phase_state`、`$recipe_state` 等） |

## 🚀 快速开始

### 下载打包好的 exe（推荐）

1. 从 [Releases](https://github.com/mobosa/DeltaV-FHX-Nameset-Editor/releases) 下载 `DeltaV_FHX_Nameset_Editor.exe`
2. 双击运行，无需安装 Python

### 从源码运行

```bash
pip install openpyxl customtkinter
python fhx_migrator.py
```

## 📖 工作流程

```
┌─────────────────────────────────────────────────────────────┐
│  Step 1: 对比 Original Database 与 New Database             │
│          → 导出 Excel（自动填充建议值）                      │
└─────────────────────────────────────────────────────────────┘
                            ↓
              用户在 Excel 中审核/修改 "New Value" 列
                            ↓
┌─────────────────────────────────────────────────────────────┐
│  Step 2: 读取 Excel → 生成新的 FHX 文件（后缀 _NEW）        │
└─────────────────────────────────────────────────────────────┘
```

## 💻 CLI 命令行模式

```bash
# 对比并导出 Excel
python fhx_migrator.py compare <Original FHX> --setup <New Database> [-o output.xlsx]

# 从编辑后的 Excel 生成新 FHX
python fhx_migrator.py generate <Original FHX> --setup <New Database> --excel edited.xlsx [-o output.fhx]
```

| 参数 | 说明 |
|------|------|
| `compare` | 对比 Original Database 与 New Database，导出 Excel |
| `generate` | 读取 Excel 生成新 FHX |
| `--setup` | **必填。** New Database 参考文件路径 |
| `--excel` | **generate 必填。** 编辑后的 Excel 文件路径 |
| `-o, --output` | 可选。输出文件路径 |

## 📊 Excel 输出格式

导出的 Excel 包含 5 个工作表：

| Sheet | 内容 | 状态值 |
|-------|------|--------|
| **Namesets** | ENUMERATION_SET 定义对比 | `Both` / `Original only` |
| **String Values** | STRING_VALUE 引用对比 | — |
| **Expression Refs** | 表达式引用对比 | — |
| **Alarm Types** | SYSTEM_ALARM / USER_ALARM 字段对比 | `Both` / `Old only` / `New only` |
| **Alarm Priorities** | PRIORITY_NAME 映射 | — |

**New Value 列填写规则：**
- **简单格式**：直接写名称，如 `STOP`、`Running` → 自动转换为 `VALUE=1 NAME="STOP"`
- **完整格式**：`VALUE=1 NAME="STOP"`
- **留空**：跳过该行

> 🔵 蓝色底色 = 已有自动建议 | 🟡 黄色底色 = 需手动填写

## ⚠️ 注意事项

- **Alarm 修改限制**：并非所有 Alarm 条目都支持直接修改。**建议整个项目导出的 FHX 不要修改 Alarm 字段。** 如需修改，请单独导出 Alarm。
- **Nameset 找不到**：不同来源导出的 FHX 文件结构可能存在差异。可参考 `Normal Namesets` 附件补充。
- **重复导入问题**：已导入过的数据库再次导入时，Alarm 属性不会被覆盖。建议在空数据库上测试。
- **大型项目**：建议分批导入，及时关注 DeltaV 导入的报错信息。

## 🛠️ 打包 exe

```bash
pip install pyinstaller
pyinstaller FHX_Migration_Tool.spec --noconfirm
```

生成的 exe 在 `dist/` 目录下。

## 📁 项目结构

```
DeltaV_FHX_Nameset_Editor/
├── fhx_core.py              # 后端逻辑（解析、对比、生成、Excel I/O）
├── fhx_migrator.py          # GUI + CLI 入口
├── test_fhxCoverage.py      # 单元测试
├── FHX_Migration_Tool.spec  # PyInstaller 打包配置
├── exp_logo.ico             # 程序图标（ICO）
├── exp_logo.png             # 程序图标（PNG）
├── requirements.txt         # Python 依赖
├── README.md                # 英文文档
└── README_zh-CN.md          # 本文档（中文）
```

## 📋 依赖

| 包 | 用途 |
|---|------|
| `openpyxl` | Excel 读写 |
| `customtkinter` | 现代化 GUI 框架 |
| Python 3.8+ | 运行环境（仅源码运行时需要） |

## 🧪 测试

```bash
pip install pytest
python -m pytest test_fhxCoverage.py -v
```

## 📄 许可证

本项目采用 MIT 许可证。

## 👤 作者

**Jared.Ji** — Jared.Ji@emerson.com

---

<div align="center">

**为 Emerson DeltaV 工程师用心打造 ❤️**

</div>
