# 安装说明

本文档说明如何在不同的AI助手环境中安装和使用本skill。

---

## 📦 通用安装方法

### 前置条件

- AI助手支持skills功能（如Google Gemini Antigravity、Claude等）
- 有访问skills目录的权限

### 安装步骤

#### 步骤1：定位skills目录

不同AI助手的skills目录位置可能不同：

**Google Gemini Antigravity**:
- Windows: `%APPDATA%\.gemini\antigravity\skills\`
- macOS/Linux: `~/.gemini/antigravity/skills/`

**Claude Code/Codex（如适用）**:
- 查看设置/配置中的skills路径
- 通常类似于 `~/.config/claude/skills/` 或项目的 `.skills/` 目录

**通用查找方法**：
1. 询问AI助手："我的skills目录在哪里？"
2. 查看AI助手的设置/配置文档
3. 搜索现有的skill文件夹位置

#### 📍 安装级别选择

**选项1：Workspace级别（项目级别）**

将skill放在当前项目目录下，只对当前项目生效：

```
your-project/
└── .agent/
    └── skills/
        └── game_design_doc/
```

**优点**：
- 只影响当前项目，不污染全局环境
- 可以随项目一起版本控制（提交到Git）
- 团队成员自动获得相同的skill配置
- 不同项目可以使用不同版本

**适用场景**：团队协作、项目专用配置

**选项2：全局级别（所有项目可用）**

将skill放在用户配置目录，对所有项目生效：

```
~/.gemini/antigravity/skills/game_design_doc/
```

**优点**：
- 一次安装，所有项目都能使用
- 个人偏好和习惯设置

**适用场景**：个人常用工具、跨项目通用skill

#### 步骤2：复制skill文件夹

**Workspace级别安装**：
```bash
# 在项目根目录创建.agent/skills目录
mkdir -p .agent/skills

# 复制skill文件夹
cp -r game_design_doc .agent/skills/

# 或Windows:
mkdir .agent\skills
xcopy game_design_doc .agent\skills\game_design_doc\ /E /I
```

**全局级别安装**：

```bash
# 方法1：直接复制
将 game_design_doc 文件夹复制到 skills 目录

# 方法2：使用命令行（Linux/macOS）
cp -r game_design_doc /path/to/skills/

# 方法3：使用命令行（Windows）
xcopy game_design_doc "%APPDATA%\.gemini\antigravity\skills\game_design_doc\" /E /I
```

#### 步骤3：验证安装

**确认文件结构**：
```
skills/
└── game_design_doc/
    ├── SKILL.md              # 必需：主文件
    ├── README.md
    ├── prompts/
    │   └── expand_doc.md
    ├── scripts/
    │   └── generate_doc.py
    └── examples/
        ├── input_brief.md
        ├── workflow_demo.md
        ├── example_activity.md
        ├── example_building.md
        └── example_system.md
```

**检查SKILL.md**：
确保文件开头有正确的frontmatter：
```yaml
---
name: 游戏功能开发文档生成器
description: AI辅助从简要玩法描述生成详细的游戏功能开发文档，指导客户端和服务器程序员实现功能
---
```

#### 步骤4：重启AI助手（如需要）

某些AI助手可能需要重启才能加载新的skill。

#### 步骤5：测试skill

向AI助手发送测试请求：
```
请帮我生成一个每日签到功能的开发文档
```

如果AI助手提到使用了"游戏功能开发文档生成器"skill，说明安装成功。

---

## 🔧 特定环境安装

### Google Gemini Antigravity

1. **定位skills目录**：
   ```
   Windows: C:\Users\<用户名>\AppData\Roaming\.gemini\antigravity\skills\
   ```

2. **复制skill**：
   - 将`game_design_doc`文件夹整个复制到上述目录
   - 完整路径示例：`C:\Users\YourName\AppData\Roaming\.gemini\antigravity\skills\game_design_doc\`

3. **无需重启**：Antigravity会自动加载新的skill

### Claude Code（VS Code扩展）

1. **查找配置目录**：
   - 打开VS Code设置
   - 搜索"Claude"或"Skills"
   - 查看skills路径配置

2. **常见路径**：
   ```
   ~/.config/Code/User/globalStorage/anthropic.claude/skills/
   ```
   或
   ```
   项目根目录/.skills/
   ```

3. **复制并重启**：
   - 复制`game_design_doc`文件夹到skills目录
   - 重启VS Code或重新加载窗口

### Codex（如适用）

1. **检查文档**：
   - 查看Codex官方文档中关于skills的说明
   - 确认是否支持自定义skills

2. **按文档操作**：
   - 通常需要将skill放在项目目录或用户配置目录
   - 可能需要在配置文件中注册skill

---

## 🐍 Python依赖（可选）

如果需要使用Word文档生成脚本，安装Python依赖：

```bash
# 进入skill目录
cd /path/to/skills/game_design_doc/

# 创建虚拟环境（推荐）
python -m venv venv

# 激活虚拟环境
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 安装依赖
pip install python-docx
```

**注意**：Python依赖仅用于Word生成脚本。AI助手使用skill本身**不需要**安装Python。

---

## ✅ 验证清单

安装完成后，检查以下项：

- [ ] `game_design_doc`文件夹在正确的skills目录下
- [ ] `SKILL.md`文件存在且包含完整的frontmatter
- [ ] 所有子目录和文件都已复制（prompts、scripts、examples）
- [ ] AI助手能够识别到skill（询问"你有哪些skills？"）
- [ ] 测试生成一个简单的功能文档

---

## 🔍 故障排除

### 问题1：AI助手无法识别skill

**可能原因**：
- 文件夹位置不正确
- SKILL.md缺失或格式错误
- 需要重启AI助手

**解决方法**：
1. 确认skills目录位置
2. 检查SKILL.md的frontmatter格式
3. 尝试重启AI助手
4. 询问AI助手："为什么无法加载game_design_doc这个skill？"

### 问题2：skill加载了但无法使用

**可能原因**：
- 文件编码问题（非UTF-8）
- 文件内容损坏

**解决方法**：
1. 检查所有.md文件是UTF-8编码
2. 重新从源头复制skill文件夹
3. 查看AI助手的错误日志

### 问题3：Python脚本无法运行

**可能原因**：
- 未安装Python
- 未安装python-docx库

**解决方法**：
```bash
# 检查Python是否安装
python --version

# 安装依赖
pip install python-docx
```

**注意**：Python脚本是可选的，不影响skill的核心功能。

---

## 🔄 更新skill

当skill有新版本时：

1. **备份当前版本**：
   ```bash
   cp -r game_design_doc game_design_doc_backup
   ```

2. **覆盖安装**：
   - 删除旧的`game_design_doc`文件夹
   - 复制新版本到skills目录

3. **保留自定义内容**（如有）：
   - 如果您修改了示例文件，先备份
   - 安装新版本后，将自定义内容合并回去

---

## 📞 获取帮助

如果遇到安装问题：

1. **查看AI助手文档**：每个AI助手对skills的支持可能略有不同
2. **检查skill文件**：确保所有文件完整且未损坏
3. **询问AI助手**：直接问"我该如何安装skills？"

---

## 📋 快速参考

**最简安装（3步）**：
```bash
# 1. 复制文件夹到skills目录
cp -r game_design_doc ~/.gemini/antigravity/skills/

# 2. 验证文件存在
ls ~/.gemini/antigravity/skills/game_design_doc/SKILL.md

# 3. 测试使用
# 向AI助手发送："帮我生成一个签到功能的文档"
```

就这么简单！
