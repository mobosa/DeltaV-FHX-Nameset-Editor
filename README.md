# DeltaV FHX Nameset Editor

批量翻译/编辑 DeltaV FHX 配置文件中的 nameset 值。主要面向 Emerson DeltaV 系统工程师，用于将中文 nameset 值迁移为英文，解决 DeltaV 导入时的字符集错误。

## 功能特性

- **ENUMERATION_SET 定义翻译** — 批量替换 `VALUE=1 NAME="停止中"` 类型的值名称
- **STRING_VALUE 引用翻译** — 替换 `SET="..." STRING_VALUE="中文值"` 中的引用值
- **表达式引用翻译** — 替换 `'$phase_state:正在保持'` 等表达式中的中文值
- **报警块翻译** — SYSTEM_ALARM / USER_ALARM 的 DESCRIPTION、ALARM_WORD、MESSAGE、CATEGORY
- **优先级名称翻译** — 危急→CRITICAL、警告→WARNING、提示→ADVISORY、记录→LOG
- **LOCALE 替换** — 根据目标 Setup.fhx 替换 locale 字符串
- **新 nameset 添加** — 支持在 FHX 中新增完整的 ENUMERATION_SET 定义块
- **内建标准 DeltaV nameset 映射** — `$phase_state`、`$recipe_state`、`$sfc_action_states` 等 20+ 套标准 nameset 自动翻译，无需手动填写

## 工作流程

```
Step 1: 对比 FHX 与 Setup.fhx → 导出 Excel（自动填充英文建议值）
    ↓
用户在 Excel 中审核/修改 "New Value" 列
    ↓
Step 2: 读取 Excel → 生成新的英文 FHX 文件（后缀 _NEW）
```

## 快速开始

### 使用打包好的 exe（推荐）

1. 下载 `DeltaV_FHX_Nameset_Editor.exe`（见 [Releases](https://github.com/mobosa/delta-v-fhx-nameset-editor/releases)）
2. 双击运行，无需安装 Python
3. **Step 1 区域**：选择 FHX 文件 + Setup.fhx → 点击 **Compare and Export Excel**
4. 打开生成的 Excel，审核/修改 `New Value` 列
5. **Step 2 区域**：选择 FHX + Setup.fhx + 编辑后的 Excel → 点击 **Generate New FHX**

### 从源码运行

```bash
pip install openpyxl
python fhx_migrator.py          # 启动 GUI
```

## CLI 模式

不带参数启动 GUI，带参数进入命令行模式：

```bash
# 对比并导出 Excel
python fhx_migrator.py compare Library.fhx --setup Setup.fhx [-o output.xlsx]

# 从编辑后的 Excel 生成新 FHX
python fhx_migrator.py generate Library.fhx --setup Setup.fhx --excel edited.xlsx [-o output.fhx]
```

| 参数 | 说明 |
|------|------|
| `compare` | 对比 FHX 与 Setup，导出 Excel |
| `generate` | 读取 Excel 生成新 FHX |
| `--setup` | 必填。Setup.fhx 参考文件路径 |
| `--excel` | `generate` 必填。编辑后的 Excel 文件路径 |
| `-o, --output` | 可选。输出文件路径，不指定则自动生成 |

## Excel 输出格式

导出的 Excel 包含 5 个工作表：

| Sheet | 内容 |
|-------|------|
| **Namesets** | ENUMERATION_SET 定义对比（FHX 值 vs Setup 值） |
| **String Values** | STRING_VALUE 引用对比 |
| **Expression Refs** | 表达式引用对比 |
| **Alarm Types** | SYSTEM_ALARM / USER_ALARM 字段对比 |
| **Alarm Priorities** | PRIORITY_NAME 中英文映射 |

**New Value 列填写规则：**
- 简单格式：直接写名称，如 `STOP`、`Running`（工具自动转换为 `VALUE=1 NAME="STOP"`）
- 完整格式：`VALUE=1 NAME="STOP"`
- 留空：该行不修改

蓝色底色 = 已有自动建议，黄色底色 = 需手动填写。

## 打包 exe

```bash
pip install pyinstaller
pyinstaller FHX_Migration_Tool.spec --noconfirm
```

生成的 exe 在 `dist/` 目录下。

## 项目结构

```
fhx_migrator.py              # 主程序（GUI + CLI + 全部逻辑）
FHX_Migration_Tool.spec      # PyInstaller 打包配置
exp_logo.ico                  # 程序图标（ICO）
exp_logo.png                  # 程序图标（PNG）
使用说明.md                   # 详细使用说明（中文）
```

## 依赖

- Python 3.8+（仅源码运行时需要）
- `openpyxl` — Excel 读写
- `tkinter` — GUI（Python 标准库）

## 作者

Jared.Ji (Jared.Ji@emerson.com)
