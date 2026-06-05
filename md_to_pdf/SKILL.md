---
name: md_to_pdf
description: 将 Markdown 文件导出为 PDF 文档。全程离线，使用 pandoc 解析 Markdown、Chrome/Edge 渲染输出 A4 PDF，中文友好排版。当用户说"导出 PDF""转成 pdf""生成 pdf 文件"时触发。
---

# Markdown 转 PDF

## 工具说明

| 项目 | 路径 |
|------|------|
| 脚本 | `.agent/skills/md_to_pdf/scripts/convert.py` |
| Python | `g:/zmd works/AIPartner/.venv/Scripts/python.exe` |
| 外部依赖 | pandoc（已安装）、Google Chrome 或 Microsoft Edge（已安装） |

输出文件与源文件同目录同名，扩展名改为 `.pdf`。

## 使用方法

```bash
"g:/zmd works/AIPartner/.venv/Scripts/python.exe" "g:/zmd works/AIPartner/.agent/skills/md_to_pdf/scripts/convert.py" "<文档绝对路径.md>"
```

**示例**：

```bash
"g:/zmd works/AIPartner/.venv/Scripts/python.exe" "g:/zmd works/AIPartner/.agent/skills/md_to_pdf/scripts/convert.py" "g:/zmd works/AIPartner/框架/心情系统.md"
```

## 注意事项

- 脚本全程使用本地字体（Microsoft YaHei / PingFang SC），不依赖外网，完全离线运行。
- 如果 Chrome 未安装，脚本会自动查找 Microsoft Edge 作为备用浏览器。
- 外部依赖 pandoc 已在系统 PATH 中安装，无需额外配置。
