# 游戏功能开发文档生成器 - 使用说明

## 快速开始

本 skill 是游戏功能文档设计工具集的**主流程入口**，用于创建和修改游戏功能设计文档。

### 1. 创建文档

向 AI 助手说："帮我生成一个 XXX 功能的设计文档"

AI 会自动执行以下流程：
1. 遍历 `框架/` 和 `开发文档/` 目录，判断是否已有相关文档
2. 根据情况建议创建框架设计文档或开发文档
3. 分析项目上下文，生成文档并执行质量检查
4. 检查关联文档影响

### 2. 使用生成脚本（可选）

**默认输出**：AI 助手直接生成 markdown 格式（.md）文档

**如需 Word 格式**：使用 `md_to_word` skill 转换。

### 3. 查看示例

在 `examples` 文件夹中有完整的示例文档：

- **example_system.md** - 系统玩法示例
- **example_building.md** - 建筑功能示例
- **example_activity.md** - 活动功能示例

### 4. 导出为 Word

使用 `md_to_word` skill 转换。

## 关联 Skill

本 skill 在工作流中引用以下 skill（均可独立调用）：

| Skill | 目录 | 独立调用场景 |
| --- | --- | --- |
| 文档结构模板 | `game_doc_template/` | 查阅模板、检查章节完整性 |
| 规则编写规范 | `game_rule_writing/` | 评审/优化规则描述 |
| 配置表设计指南 | `game_config_table/` | 设计/修改/审查配置表 |
| 质量检查 | `game_doc_quality/` | 文档审查、关联影响检查 |
| Markdown 排版 | `markdown_format_standard/` | 检查 Markdown 格式 |

## 技术要求

**运行环境**（仅脚本需要）：
- Python 3.6+
- 安装依赖：`pip install -r requirements.txt`

## 反馈与改进

如果模板不符合你的需求，可以：
1. 修改对应 skill 的 `SKILL.md` 调整规范和模板
2. 在 `examples` 中添加新的示例类型
