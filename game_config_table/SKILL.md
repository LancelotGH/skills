---
name: 游戏配置表设计指南
description: 提供游戏功能文档中配置表结构的设计规范，包含复用原则、字段设计、命名规范和常见类型参考，可独立调用以设计、修改或审查配置表。
---

# 游戏配置表设计指南

## 概述

本 skill 提供游戏功能文档中**配置表结构的设计规范和最佳实践**，确保配置表设计规范、可维护、易扩展。

### 使用场景

- 被 `game_doc_generator` 主流程引用，指导配置表需求章节的编写
- **独立调用**：为某个功能独立设计配置表、修改已有配置表、审查配置表设计

### 独立调用触发词

- "设计配置表"
- "修改配置表"
- "审查配置表"
- "配置表字段不合理，帮我优化"

---

## 双模式行为差异

- **框架设计文档模式**：仅需在规则中标注参数是否"可配置"，用【配置】标记，并提供建议的表名
- **开发文档模式**：需要完整的配置表结构设计（字段名、类型、说明、对应规则）

---

## 配置表复用原则

⚠️ **最重要的原则**：

- **优先复用已有配置表** — 检查项目中已有的配置表清单
- **禁止重复创建表** — 如果已有表可以满足需求，只需在已有表中增加数据或字段
- **只有在没有合适的已有表时，才能创建新表**

**操作类型和标注方式**：

必须在每个表前明确标注操作类型，并说明为什么选择该操作：

**【增加数据】** 在已有表 xxx_config 中增加以下数据行：

| 字段名 | 数据值示例 |
|--------|------------|
| item_id | 10001 |
| item_name | 新道具 |
| item_count | 5 |

**【增加字段】** 在已有表 xxx_config 中增加以下字段：

| 字段名 | 类型 | 说明 | 对应规则参数 |
|--------|------|------|-------------|
| new_field | int | 新增字段说明 | 规则3.x中的"xxx参数" |

**【新建】** 新建表：xxx_config

**说明为什么需要新建：** 现有的 aaa_config 和 bbb_config 表都无法满足该功能的配置需求，因为...

| 字段名 | 类型 | 说明 | 对应规则参数 |
|--------|------|------|-------------|
| config_id | int | 配置ID，主键 | - |
| config_name | string | 配置名称 | - |

---

## 配置表字段设计原则 — 以 Excel 为主

⚠️ **核心要求**：

- **禁止使用JSON或复杂嵌套结构**：配置表必须以Excel编辑为主，所有字段必须是扁平化的基础类型
- **基础类型**：int、string、float等Excel直接支持的类型
- **处理多值字段**：使用逗号分隔的字符串（如："1,2,3,4"）
- **处理多属性字段**：展开为多组字段（如：attribute_1_id, attribute_1_value, attribute_2_id, attribute_2_value）
- **错误示例**：`unlock_rewards | json | {"dialogue": [1,2], "items": [10,20]}`
- **正确示例**：`unlock_dialogue_ids | string | "1,2"` 和 `unlock_item_ids | string | "10,20"`

**基础要求**：
- **配置表字段必须覆盖规则中提到的所有可配置参数** ⚠️
- 每个字段必须注明对应的规则章节，确保一致性
- 字段命名遵循项目规范
- 考虑扩展性和维护性

---

## 配置表命名规范

默认标准，若项目已有规范则以项目为准：

| 规范项 | 要求 | 示例 |
|-------|------|------|
| 表名 | PascalCase，以 Config 结尾 | `ItemConfig`、`QuestConfig`、`RewardConfig` |
| 字段名 | snake_case，清晰描述用途 | `item_id`、`unlock_level`、`max_count` |
| 主键 | `xxx_id` 格式 | `item_id`、`quest_id` |
| 外键 | 与关联表主键同名 | `reward_id` 关联 RewardConfig 表 |
| 布尔值 | 以 `is_` 或 `can_` 开头 | `is_active`、`can_trade` |
| 避免缩写 | 除行业通用缩写外禁止使用 | ❌ `atk` → ✅ `attack` |

---

## ID 分段设计（建议）

使用分段ID区分不同类别，便于管理和扩展：
- 示例：1000-1999 为物品类A，2000-2999 为物品类B，10000-19999 为任务类

---

## 表间关系设计

- **一对多关系**：从表中用外键字段关联主表ID（如 `StageConfig.chapter_id → ChapterConfig.chapter_id`）
- **多对多关系**：通过中间映射表关联（如 `QuestRewardMap` 连接任务表和奖励表）
- 使用 0 或 -1 表示"无关联"

---

## 常见配置表类型参考

根据功能类型选用：

| 配置表类型 | 典型表名 | 核心字段 |
|-----------|---------|----------|
| 物品配置 | ItemConfig | item_id, item_name, item_type, quality, max_stack, sell_price |
| 任务配置 | QuestConfig | quest_id, quest_type, target_type, target_count, reward_id |
| 奖励配置 | RewardConfig | reward_id, reward_type, item_id, item_count, currency_amount |
| 关卡配置 | StageConfig | stage_id, difficulty, cost_energy, time_limit, reward_id |
| 商店配置 | ShopConfig | shop_id, item_id, price_type, price_amount, stock_limit |
| 活动配置 | ActivityConfig | activity_id, start_time, end_time, unlock_level, params |
| 等级经验 | LevelExpConfig | level, exp_required, hp_bonus, attack_bonus |
| VIP特权 | VipConfig | vip_level, exp_required, discount_rate, energy_bonus |
| 消耗配置 | CostConfig | cost_id, action_type, cost_type, cost_amount |
| 技能配置 | SkillConfig | skill_id, skill_type, cost_type, cooldown, damage_base |

---

## 常见错误和注意事项

| 错误类型 | 错误做法 | 正确做法 |
|---------|---------|----------|
| 魔法数字 | `status: 1, 2, 3`（含义不明） | 在说明中列出枚举含义（1=正常, 2=维护, 3=下线） |
| 字段过度复杂 | `data: "1001:5|1002:10"`（自定义分隔） | 使用规范的分隔格式并在说明中注明解析规则 |
| 缺少默认值说明 | 没有说明 0 或空字符串的含义 | 明确说明 0 表示"无限制"，空字符串表示"无" |
| 字段名不一致 | 同一含义在不同表中使用不同字段名 | 统一命名，如都使用 `item_id` 而非混用 `prop_id` |

---

## 配置表设计示例

**新建表：activity_config**

| 字段名 | 类型 | 说明 | 对应规则参数 |
|--------|------|------|-------------|
| activity_id | int | 活动ID，主键 | - |
| activity_name | string | 活动名称 | - |
| start_time | datetime | 开始时间 | 规则3.2中的"活动开始时间" |
| end_time | datetime | 结束时间 | 规则3.5中的"活动结束时间" |
| open_level | int | 开启等级要求 | 规则3.2中的"开启条件" |
| reward_item_ids | string | 奖励道具ID列表，逗号分隔 | 规则3.5中的"奖励发放" |
| reward_item_counts | string | 奖励道具数量列表，逗号分隔 | 规则3.5中的"奖励发放" |
