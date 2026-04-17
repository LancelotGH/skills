---
name: md_to_word
description: 将项目开发文档目录下的 Markdown 文件转换为 Word（.docx）格式。当用户要求"生成 word"、"导出 word"、"转换为 word"或类似操作时使用此 skill。
---

# Markdown 转 Word 工具

## 工具说明

项目内置了专用的 Markdown → Word 转换脚本，支持中文排版、标题样式、表格和列表的正确渲染。

- **脚本路径**：`g:/zmd works/AIPartner/.agent/skills/md_to_word/generate_word_docs.py`
- **Python 路径**：`g:/zmd works/AIPartner/.venv/Scripts/python.exe`

## 使用方法

**转换单个文件**：

```bash
cd "g:/zmd works/AIPartner/开发文档"
"g:/zmd works/AIPartner/.venv/Scripts/python.exe" "g:/zmd works/AIPartner/.agent/skills/md_to_word/generate_word_docs.py" "文件名.md"
```

**转换当前目录下全部 .md 文件**（双击运行）：

```
g:/zmd works/AIPartner/开发文档/将当前目录md转word.bat
```

## 注意事项

- 脚本执行时必须 `cd` 到 .md 文件所在目录，再传入文件名（不含路径）。
- 输出的 .docx 文件与 .md 文件同名、同目录。
- 脚本会自动修复常见的 Markdown 格式问题（缺失空行、前置空格等）再执行转换。
