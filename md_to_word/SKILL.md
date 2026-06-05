---
name: md_to_word
description: 将 Markdown 文件导出为 Word（.docx）文档。全程离线，使用项目内置脚本，支持中文排版、标题样式、表格和列表正确渲染。当用户说"导出 Word""转成 word""生成 word 文件"时触发。
---

# Markdown 转 Word

## 工具说明

| 项目 | 路径 |
|------|------|
| 脚本 | `.agent/skills/md_to_word/generate_word_docs.py` |
| Python | `g:/zmd works/AIPartner/.venv/Scripts/python.exe` |
| 外部依赖 | 无（python-docx 已在 .venv 中安装） |

输出文件与源文件同目录同名，扩展名改为 `.docx`。

## 使用方法

脚本运行时**必须先 cd 到 .md 文件所在目录**，再传入文件名（不含路径）：

```bash
cd "<md文件所在目录>"
"g:/zmd works/AIPartner/.venv/Scripts/python.exe" "g:/zmd works/AIPartner/.agent/skills/md_to_word/generate_word_docs.py" "<文件名.md>"
```

**示例**：

```bash
cd "g:/zmd works/AIPartner/框架"
"g:/zmd works/AIPartner/.venv/Scripts/python.exe" "g:/zmd works/AIPartner/.agent/skills/md_to_word/generate_word_docs.py" "心情系统.md"
```

## 注意事项

- 必须 cd 到文件所在目录再执行，不能传绝对路径作为参数。
- 脚本会自动修复常见的 Markdown 格式问题（缺失空行、前置空格等）再执行转换。
- 完全离线，不依赖任何外部服务或网络。
