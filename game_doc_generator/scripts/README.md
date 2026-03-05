# 文档解析工具脚本说明

## 脚本列表

### 1. read_docx.py - Word文档读取工具

**用途**：读取`.docx`格式的Word文档，提取标题、段落和表格内容

**使用方法**：
```bash
python read_docx.py <word文件路径> [输出格式]
```

**参数**：
- 文件路径：必填，Word文档的完整路径
- 输出格式：可选，`json` 或 `markdown`（默认为markdown）

**示例**：
```bash
python read_docx.py "g:/project/docs/功能设计.docx" markdown
```

**输出说明**：
- 自动识别标题（Heading样式）
- 提取所有段落文本
- 提取表格数据（最多前20行）
- 默认最多读取200个段落和50个表格

---

### 2. read_xlsx.py - Excel文件读取工具

**用途**：读取`.xlsx`格式的Excel文件，提取工作表和数据

**使用方法**：
```bash
python read_xlsx.py <excel文件路径> [输出格式]
```

**参数**：
- 文件路径：必填，Excel文件的完整路径
- 输出格式：可选，`json` 或 `markdown`（默认为markdown）

**示例**：
```bash
python read_xlsx.py "g:/project/data/配置表.xlsx" markdown
```

**输出说明**：
- 读取所有工作表
- 提取表格数据（最多100行×20列）
- 第一行自动识别为表头
- 显示前20行数据

---

### 3. generate_doc.py - 文档框架生成工具

**用途**：快速生成标准的游戏功能设计文档框架

**使用方法**：
```bash
python generate_doc.py --name "功能名称" --type "system"
```

**参数**：
- `--name`：功能名称（必填）
- `--type`：功能类型，可选值：system/building/activity/other（必填）
- `--output`：输出路径（可选，默认为当前目录）

---

## AI使用指南

当需要读取项目中的Word或Excel文档时，可以使用`run_command`工具调用这些脚本：

**读取Word文档示例**：
```python
run_command(
    CommandLine='python "g:/zmd works/skills/game_design_doc/scripts/read_docx.py" "文档路径.docx" markdown',
    Cwd='g:/zmd works/skills/game_design_doc/scripts',
    SafeToAutoRun=True
)
```

**读取Excel文件示例**：
```python
run_command(
    CommandLine='python "g:/zmd works/skills/game_design_doc/scripts/read_xlsx.py" "文件路径.xlsx" markdown',
    Cwd='g:/zmd works/skills/game_design_doc/scripts',
    SafeToAutoRun=True
)
```

---

## 依赖安装

这些脚本依赖以下Python库：

```bash
pip install python-docx openpyxl
```

确保在使用前已安装这些依赖。

---

### 4. convert_md_v2.py - Markdown转Word工具

**用途**：将 Markdown 文档转换为 Word (.docx) 格式，支持中文优化和表格/图片/代码块渲染。

**使用方法**：
```bash
python convert_md_v2.py <input.md> <output.docx>
```

**参数**：
- `input`：输入的 Markdown 文件路径
- `output`：输出的 Word 文件路径

**特性**：
- 自动设置中文字体（正文宋体，标题黑体）
- 支持表格及其边框央视
- 支持代码块和列表
- 支持加粗等行内样式

**示例**：
```bash
python convert_md_v2.py "../docs/design.md" "../docs/design.docx"
```

