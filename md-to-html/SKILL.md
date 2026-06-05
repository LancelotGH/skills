---
name: md-to-html
description: 将 Markdown 文件导出为美观的 HTML 文档。全程离线，使用 pandoc 转换，注入内嵌 CSS（中文字体、侧边目录、表格斑马纹、blockquote 样式），输出与源文件同目录。当用户说"导出 HTML""转成 html""生成 html 文件"时触发。
---

# Markdown 转 HTML

## 工具说明

| 项目 | 路径 |
|------|------|
| 脚本 | `.agent/skills/md-to-html/scripts/convert.py` |
| Python | `g:/zmd works/AIPartner/.venv/Scripts/python.exe` |
| 外部依赖 | pandoc（已安装） |

输出文件与源文件同目录同名，扩展名改为 `.html`。

## 使用方法

```bash
"g:/zmd works/AIPartner/.venv/Scripts/python.exe" "g:/zmd works/AIPartner/.agent/skills/md-to-html/scripts/convert.py" "<文档绝对路径.md>"
```

**示例**：

```bash
"g:/zmd works/AIPartner/.venv/Scripts/python.exe" "g:/zmd works/AIPartner/.agent/skills/md-to-html/scripts/convert.py" "g:/zmd works/AIPartner/框架/心情系统.md"
```

## 注意事项

- 完全离线，CSS 和字体均内嵌，不依赖外网。
- 外部依赖 pandoc 已在系统 PATH 中安装，无需额外配置。
- TOC 锚点：pandoc 生成标题 ID 时会去掉编号前缀（如 `## 1. 概述` 的 ID 是 `概述`），目录链接应写 `[概述](#概述)` 而非 `[概述](#1-概述)`。
