# 游戏功能开发文档生成器 - 使用说明

## 快速开始

本skill用于生成标准化的游戏功能策划设计文档。

### 1. 查看SKILL.md

首先阅读 `SKILL.md` 了解完整的文档模板结构和填写指南。

### 2. 使用生成脚本（可选）

**默认输出**：AI助手直接生成markdown格式（.md）文档

**如需Word格式**，使用Python脚本快速生成文档框架：

```bash
cd scripts
python generate_doc.py --name "你的功能名称" --type "system"
```

**参数说明**：
- `--name`：功能名称（必填）
- `--type`：功能类型（必填），可选：
  - `system`：系统玩法
  - `building`：建筑功能
  - `activity`：活动功能
  - `other`：其他类型

**示例**：
```bash
# 生成系统玩法文档
python generate_doc.py --name "神石系统" --type "system"

# 生成建筑功能文档
python generate_doc.py --name "城堡建造" --type "building"

# 生成活动功能文档
python generate_doc.py --name "七日签到" --type "activity"
```

### 3. 查看示例

在 `examples` 文件夹中有3个完整的示例文档：

- **example_system.md** - 神石系统（系统玩法示例）
- **example_building.md** - 城堡建筑（建筑功能示例）
- **example_activity.md** - 活动中心（活动功能示例）

建议根据你要编写的功能类型查看对应示例，了解各章节的详细填写方式。

### 4. 填充文档内容

生成的文档包含：
- ✅ 完整的章节标题结构
- ✅ 填写提示和说明
- ✅ 预置的表格模板

你只需要根据实际功能需求填充各章节内容即可。

## 文档结构

标准文档包含4个核心章节：

```
一、设计目的
  1.1 功能定位
  1.2 期望体验

二、功能概述
  2.1 背景概述
  2.2 功能简介
  2.3 结构划分

三、规则说明
  3.1-3.8 根据功能类型调整

四、策划需求
  4.1 数值需求
  4.2 系统需求
  4.3 配置表需求
```

## 注意事项

⚠️ **文档聚焦策划设计**：
- 不包含UI原型图
- 不包含美术资源需求
- 不包含竞品分析
- 不包含界面交互细节

✅ **必须详细说明**：
- 功能的所有规则和逻辑
- 特殊情况处理方式
- 完整的数值配置
- 系统依赖和接口需求

## 技术要求

**运行环境**：
- Python 3.6+
- python-docx库

**安装依赖**：
```bash
pip install python-docx
```

## 反馈与改进

如果模板不符合你的需求，可以：
1. 修改 `SKILL.md` 调整模板结构
2. 修改 `scripts/generate_doc.py` 定制生成逻辑
3. 在 `examples` 中添加新的示例类型
