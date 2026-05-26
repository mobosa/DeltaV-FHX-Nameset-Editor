# DeltaV FHX Nameset Editor

批量编辑 DeltaV FHX 配置文件中的 nameset 值。典型用途是将中文 nameset 值翻译为英文，解决 DeltaV 不支持中文字符导致的导入错误。也可用于英文环境下的 nameset 值批量修改。

## 功能

- **Setup** — 编辑 `ENUMERATION_SET` 定义中的值名称
- **Library** — 编辑 `ENUMERATION_SET` 定义 + `STRING_VALUE` 引用 + 表达式引用，内置标准 DeltaV nameset 自动翻译
- **Control Strategies** — 编辑 `STRING_VALUE` 引用 + 表达式引用

额外支持：
- 报警块翻译（SYSTEM_ALARM / USER_ALARM 的 DESCRIPTION、ALARM_WORD、MESSAGE、CATEGORY）
- 优先级名称翻译（"危急"→"CRITICAL"、"警告"→"WARNING" 等）
- LOCALE 替换

## 工作流程

```
Step 1: 对比中文 FHX 与英文 Setup.fhx → 导出 Excel（自动填充英文建议值）
    ↓
用户在 Excel 中审核/修改 "New Value" 列
    ↓
Step 2: 读取 Excel → 生成新的英文 FHX 文件（后缀 _NEW）
```

## 快速开始

### 使用打包好的 exe

1. 下载 `DeltaV_FHX_Nameset_Editor.exe`
2. 双击运行
3. 选择标签页 → 加载 FHX + Setup.fhx → 点击 "Compare and Export Excel"
4. 编辑 Excel 中的 "New Value" 列
5. 回到工具 → 点击 "Generate New FHX"

### 从源码运行

```bash
pip install openpyxl
python fhx_migrator.py
```

### CLI 模式

不带参数运行启动 GUI，带参数则进入命令行模式：

```bash
# Setup 模式
python fhx_migrator.py compare input.fhx --setup Setup.fhx
python fhx_migrator.py generate input.fhx --setup Setup.fhx --excel edited.xlsx

# Library 模式
python fhx_migrator.py lib-compare Library.fhx --setup Setup.fhx
python fhx_migrator.py lib-generate Library.fhx --setup Setup.fhx --excel edited.xlsx

# Control Strategies 模式
python fhx_migrator.py cs-compare CS.fhx --setup Setup.fhx
python fhx_migrator.py cs-generate CS.fhx --excel edited.xlsx [--setup Setup.fhx]
```

所有子命令支持 `-o` / `--output` 指定输出路径，不指定则自动生成默认文件名。

### 打包 exe

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
使用说明.md                   # 详细使用说明
```

## 详细文档

参见 [使用说明.md](使用说明.md)

## 作者

Jared.Ji (Jared.Ji@emerson.com)
