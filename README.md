# 游戏功能文档设计工具集 - 使用说明

## 概述

本工具集是一套协同工作的 AI skill，用于辅助游戏策划生成和维护功能设计文档。支持**框架设计文档**和**开发文档**两种模式。

## Skill 清单

| Skill | 目录 | 职责 | 独立调用 |
| --- | --- | --- | --- |
| 游戏功能开发文档生成器 | `game_doc_generator/` | 主流程入口 + 文档创建/修改工作流 | ✅ |
| 游戏功能文档结构模板 | `game_doc_template/` | 标准文档章节结构模板 | ✅ |
| 游戏功能规则编写规范 | `game_rule_writing/` | 规则描述规范 + 逻辑拆分 + 系统解耦 | ✅ |
| 游戏配置表设计指南 | `game_config_table/` | 配置表设计原则 + 命名规范 | ✅ |
| 游戏功能文档质量检查 | `game_doc_quality/` | 质量检查 + 关联影响校验 | ✅ |
| Markdown 文档排版规范 | `markdown_format_standard/` | Markdown 排版格式规范 | ✅ |

## 快速开始

### 创建新文档

向 AI 助手说："帮我生成一个 XXX 功能的设计文档"

AI 会自动：
1. 遍历 `框架/` 和 `开发文档/` 目录，查找是否已有相关文档
2. 若无相关文档，优先创建框架设计文档
3. 按照完整工作流生成文档并执行质量检查

### 独立使用某个 Skill

每个 skill 都可以单独调用：

- **质量检查**："帮我检查这个文档的质量"
- **配置表设计**："帮我设计 XXX 的配置表"
- **规则评审**："帮我优化这段规则描述"
- **模板查阅**："文档应该包含哪些章节"

### Word 文档生成（可选）

如需将 Markdown 文档转换为 Word 格式：

```bash
cd game_doc_generator/scripts
python convert_md_v2.py "文档路径.md" "输出路径.docx"
```

## 两种文档模式

| 维度 | 框架设计文档 | 开发文档 |
| --- | --- | --- |
| 存放目录 | `框架/` | `开发文档/` |
| 写作风格 | 通俗易懂，侧重设计意图 | 严谨精确，侧重实现细节 |
| C/S逻辑 | ❌ | ✅ |
| 配置表 | 仅标注"可配置" | 完整表结构设计 |

## 示例参考

在 `game_doc_generator/examples/` 文件夹中提供了完整的示例文档：

- **example_system.md** - 系统玩法示例
- **example_building.md** - 建筑功能示例
- **example_activity.md** - 活动功能示例

## 技术要求

**Python 脚本（可选）**：
- Python 3.6+
- 安装依赖：`pip install -r game_doc_generator/requirements.txt`

**注意**：Python 仅用于 Word 生成脚本，AI 使用 skill 本身不需要 Python。
