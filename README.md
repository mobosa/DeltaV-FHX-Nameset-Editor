# DeltaV FHX Nameset Editor

批量翻译/编辑 DeltaV FHX 配置文件中的 nameset 值。面向 Emerson DeltaV 系统工程师，用于将中文 nameset 值迁移为英文，解决 DeltaV 导入时的字符集错误。支持任何 FHX 类型（Library、Control Strategies、Setup、Recipes 等），自动识别并处理。

## 功能特性

- **ENUMERATION_SET 定义翻译** — 批量替换 `VALUE=1 NAME="停止中"` 类型的值名称
- **STRING_VALUE 引用翻译** — 替换 `SET="..." STRING_VALUE="中文值"` 中的引用值
- **表达式引用翻译** — 替换 `'$phase_state:正在保持'` 等表达式中的中文值
- **报警块翻译** — SYSTEM_ALARM / USER_ALARM 的 DESCRIPTION、ALARM_WORD、MESSAGE、CATEGORY
- **优先级名称翻译** — 危急→CRITICAL、警告→WARNING、提示→ADVISORY、记录→LOG
- **LOCALE 替换** — 根据 New Database.fhx 替换 locale 字符串
- **新 nameset 添加** — 支持在 FHX 中新增完整的 ENUMERATION_SET 定义块
- **内建标准 DeltaV nameset 映射** — `$phase_state`、`$recipe_state`、`$sfc_action_states` 等 20+ 套标准 nameset 自动翻译，无需手动填写
- **自动处理任意 FHX 类型** — Library、Control Strategies、Setup、Recipes 等，无需手动选择类型

## 工作流程

```
Step 1: 对比任意 FHX 与 New Database.fhx → 导出 Excel（自动填充英文建议值）
    ↓
用户在 Excel 中审核/修改 "New Value" 列
    ↓
Step 2: 读取 Excel → 生成新的 FHX 文件（后缀 _NEW）
```

## 快速开始

### 使用打包好的 exe（推荐）

1. 下载 `DeltaV_FHX_Nameset_Editor.exe`（见 [Releases](https://github.com/mobosa/delta-v-fhx-nameset-editor/releases)）
2. 双击运行，无需安装 Python
3. **Step 1 区域**：选择 FHX 文件 + New Database.fhx → 点击 **Compare and Export Excel**
4. 打开生成的 Excel，审核/修改 `New Value` 列
5. **Step 2 区域**：选择 FHX + New Database.fhx + 编辑后的 Excel → 点击 **Generate New FHX**

### 从源码运行

```bash
pip install openpyxl
python fhx_migrator.py          # 启动 GUI
```

## CLI 模式

不带参数启动 GUI，带参数进入命令行模式：

```bash
# 对比并导出 Excel
python fhx_migrator.py compare <任意FHX文件> --setup <New Database.fhx> [-o output.xlsx]

# 从编辑后的 Excel 生成新 FHX
python fhx_migrator.py generate <任意FHX文件> --setup <New Database.fhx> --excel edited.xlsx [-o output.fhx]
```

| 参数 | 说明 |
|------|------|
| `compare` | 对比任意 FHX 与 New Database.fhx，导出 Excel |
| `generate` | 读取 Excel 生成新 FHX |
| `--setup` | 必填。New Database.fhx 参考文件路径 |
| `--excel` | `generate` 必填。编辑后的 Excel 文件路径 |
| `-o, --output` | 可选。输出文件路径，不指定则自动生成 |

> 注意：旧版的 `cs-compare`、`cs-generate`、`lib-compare`、`lib-generate` 子命令已移除，统一使用 `compare` / `generate`，工具会自动识别 FHX 类型。

## Excel 输出格式

导出的 Excel 包含 5 个工作表：

| Sheet | 内容 |
|-------|------|
| **Namesets** | ENUMERATION_SET 定义对比（FHX 值 vs New Database 值） |
| **String Values** | STRING_VALUE 引用对比 |
| **Expression Refs** | 表达式引用对比 |
| **Alarm Types** | SYSTEM_ALARM / USER_ALARM 字段对比 |
| **Alarm Priorities** | PRIORITY_NAME 中英文映射 |

**New Value 列填写规则：**
- 简单格式：直接写名称，如 `STOP`、`Running`（工具自动转换为 `VALUE=1 NAME="STOP"`）
- 完整格式：`VALUE=1 NAME="STOP"`
- 留空：该行不修改

蓝色底色 = 已有自动建议，黄色底色 = 需手动填写。

生成 FHX 前，工具会自动验证 Excel 数据格式，如发现问题会提示并阻止生成。

## 注意事项

- **Alarm 修改限制**：整个项目导出的 FHX 文件中，并非所有 Alarm 条目都支持直接修改，部分条目可能存在读写权限或结构性限制。**建议整个项目导出的 FHX 不要修改 Alarm 字段。**
- **如何修改 Alarm**：如需修改 Alarm，可将 Alarm 单独导出后操作 — 在 Step 1 和 Step 2 中两个文件选择窗格都选同一个 FHX 文件即可，编辑生成的 Excel 后重新生成 FHX 并导入。
- **Nameset 找不到**：不同来源导出的 FHX 文件结构可能存在差异，部分 Nameset 可能无法匹配到建议值。
- **重复导入问题**：已导入过的数据库再次导入修改过的 FHX 文件时，Alarm 相关属性不会被覆盖。只有空数据库才能完整导入所有属性。建议在空数据库上测试。
- **分开导出的项目**：如果项目按 Setup.fhx / Library.fhx / Control Strategies 等分开导出，部分 Nameset 可能提取不到建议值。可参考 `Normal Namesets` 附件补充。
- **大型项目建议拆分导入**：分批导入可以及时关注 DeltaV 导入的报错信息，避免相互覆盖，也便于定位问题所在。

## 打包 exe

```bash
pip install pyinstaller
pyinstaller FHX_Migration_Tool.spec --noconfirm
```

生成的 exe 在 `dist/` 目录下。

## 项目结构

```
fhx_core.py                # 后端逻辑（解析、对比、生成、Excel I/O，无 GUI 依赖）
fhx_migrator.py            # 入口文件（GUI + CLI，从 fhx_core 导入后端）
FHX_Migration_Tool.spec    # PyInstaller 打包配置
exp_logo.ico               # 程序图标（ICO）
exp_logo.png               # 程序图标（PNG）
test_fhxCoverage.py        # 单元测试
使用说明.md                 # 详细使用说明（中文）
```

## 依赖

- Python 3.8+（仅源码运行时需要）
- `openpyxl` — Excel 读写
- `customtkinter` — 现代化 GUI（可选，仅 GUI 模式需要）

## 测试

```bash
pip install pytest
python -m pytest test_fhxCoverage.py -v
```

## 作者

Jared.Ji (Jared.Ji@emerson.com)
