# 安装说明

本文档说明如何在不同 AI 助手环境中安装本工具集（包含 6 个 skill）。

---

## 📦 工具集文件结构

```
game_design_skills/              # 整体打包后的目录名
├── README.md                    # 工具集总说明
├── INSTALL.md                   # 本文件
├── game_doc_generator/          # 主流程入口
│   ├── SKILL.md
│   ├── README.md
│   ├── CLAUDE_TROUBLESHOOTING.md
│   ├── requirements.txt
│   ├── scripts/                 # Python 脚本工具
│   ├── examples/                # 示例文档
│   └── prompts/
├── game_doc_template/           # 文档结构模板
│   └── SKILL.md
├── game_rule_writing/           # 规则编写规范
│   └── SKILL.md
├── game_config_table/           # 配置表设计指南
│   └── SKILL.md
├── game_doc_quality/            # 质量检查与关联校验
│   └── SKILL.md
└── markdown_format_standard/    # Markdown 排版规范
    └── SKILL.md
```

---

## ⚠️ 安装前须知

- 工具集包含 **6 个独立 skill**，必须**全部安装**才能正常使用完整工作流
- 每个 skill 目录必须是 skills 根目录的**直接子目录**（不可嵌套额外父目录）
- 仅 `game_doc_generator` 包含 scripts/examples 等资源文件，其他 skill 只有 SKILL.md

---

## 📍 安装级别选择

### 选项1：Workspace 级别（项目级别，推荐）

将 skill 放在当前项目目录下，只对当前项目生效：

```
your-project/
└── .agent/             # 部分环境使用 .agent
    └── skills/
        ├── game_doc_generator/
        ├── game_doc_template/
        └── ...
```

或

```
your-project/
└── .cursor/            # Cursor 编辑器
    └── skills/
```

**优点**：
- 只影响当前项目，不污染全局环境
- 可以随项目一起版本控制（Git）
- 团队成员自动获得相同的 skill 配置

### 选项2：全局级别（所有项目可用）

将 skill 放在用户配置目录，对所有项目生效。不同环境的全局路径见下方各环境说明。

---

## 🔧 各环境安装指南

### Google Gemini（Antigravity / Jules / Gemini CLI）

**Skills 目录位置**：

| 安装级别 | 路径 |
| --- | --- |
| Workspace | `项目根目录/.agent/skills/` |
| 全局（Windows） | `%APPDATA%\.gemini\antigravity\skills\` |
| 全局（macOS/Linux） | `~/.gemini/antigravity/skills/` |

**安装命令**（Workspace 级别）：

```bash
# 在项目根目录执行
mkdir -p .agent/skills

# 复制所有 6 个 skill 文件夹
cp -r game_doc_generator .agent/skills/
cp -r game_doc_template .agent/skills/
cp -r game_rule_writing .agent/skills/
cp -r game_config_table .agent/skills/
cp -r game_doc_quality .agent/skills/
cp -r markdown_format_standard .agent/skills/
```

```powershell
# Windows PowerShell
$skills = @("game_doc_generator", "game_doc_template", "game_rule_writing", "game_config_table", "game_doc_quality", "markdown_format_standard")
foreach ($s in $skills) {
    Copy-Item -Path ".\$s" -Destination ".\.agent\skills\$s" -Recurse -Force
}
```

**注意**：Antigravity 会自动扫描并加载 skills 目录下的内容，无需重启。

---

### Claude Code（Anthropic 官方 CLI / VS Code 扩展）

**Skills 目录位置**：

| 安装级别 | 路径 |
| --- | --- |
| Workspace | `项目根目录/.claude/skills/` 或 `项目根目录/.agent/skills/` |
| 全局 | `~/.claude/skills/` |

**安装方式1：直接复制文件夹**

```bash
mkdir -p .claude/skills
cp -r game_doc_generator .claude/skills/
cp -r game_doc_template .claude/skills/
# ... 其他 skill 同理
```

**安装方式2：ZIP 上传（如界面支持）**

Claude 对文件夹命名有**严格限制**：
- ✅ 允许：字母（a-z, A-Z）、数字（0-9）、下划线（_）、连字符（-）
- ❌ 不允许：空格、中文字符、其他特殊字符

每个 skill 需单独打包为 zip 上传：

```powershell
# Windows PowerShell - 批量打包
$skills = @("game_doc_generator", "game_doc_template", "game_rule_writing", "game_config_table", "game_doc_quality", "markdown_format_standard")
foreach ($s in $skills) {
    Compress-Archive -Path ".\$s" -DestinationPath ".\$s.zip" -Force
}
```

```bash
# macOS/Linux - 批量打包
for s in game_doc_generator game_doc_template game_rule_writing game_config_table game_doc_quality markdown_format_standard; do
    zip -r "$s.zip" "$s/"
done
```

⚠️ 如遇命名报错，参考 `game_doc_generator/CLAUDE_TROUBLESHOOTING.md`

---

### Cursor（VS Code 分支编辑器）

**Skills 目录位置**：

| 安装级别 | 路径 |
| --- | --- |
| Workspace | `项目根目录/.cursor/skills/` 或 `项目根目录/.agent/skills/` |
| 全局 | 打开 Cursor 设置 → 搜索 "skills" → 查看配置路径 |

**安装步骤**：
1. 在项目根目录创建 `.cursor/skills/` 目录（如不存在）
2. 将 6 个 skill 文件夹复制到该目录
3. Cursor 会自动加载，无需重启

```bash
mkdir -p .cursor/skills
cp -r game_doc_generator game_doc_template game_rule_writing game_config_table game_doc_quality markdown_format_standard .cursor/skills/
```

---

### Windsurf（Codeium 编辑器）

**Skills 目录位置**：

| 安装级别 | 路径 |
| --- | --- |
| Workspace | `项目根目录/.windsurf/skills/` 或 `项目根目录/.agent/skills/` |
| 全局 | 查看 Windsurf 设置中的 skills 路径配置 |

**安装步骤**：
1. 确认 Windsurf 版本是否支持自定义 skills
2. 在对应目录创建 skills 文件夹
3. 复制 6 个 skill 文件夹到目录中

---

### VS Code + Copilot / 其他 AI 插件

**通用安装方式**：

不同 AI 插件的 skills 目录可能各异，通用查找方法：
1. 查看插件文档中关于 "skills"、"custom instructions"、"prompts" 的说明
2. 在 VS Code 设置中搜索 "skills" 关键词
3. 询问 AI 助手："我的 skills 目录在哪里？"

常见路径模式：
```
项目根目录/.vscode/skills/
项目根目录/.agent/skills/
项目根目录/.ai/skills/
~/.config/[插件名]/skills/
```

**替代方案（如不支持 skills 功能）**：

如果 AI 插件不原生支持 skills，可将 SKILL.md 内容添加到以下位置：
- **System Prompt / Custom Instructions**：将核心规则粘贴到全局指令中
- **项目级 .cursorrules / .clinerules / AGENTS.md**：将规则添加到项目级别的指令文件
- **手动引用**：在对话中发送 "请阅读 xxx/SKILL.md 并按照其规则执行"

---

### Codex（OpenAI）

**Skills 目录位置**：

| 安装级别 | 路径 |
| --- | --- |
| Workspace | `项目根目录/.codex/skills/` 或 `项目根目录/AGENTS.md` 引用 |

**安装步骤**：
1. 查看 Codex 官方文档确认 skills 支持方式
2. 如支持文件夹式 skills，按通用方法复制
3. 如使用 AGENTS.md 机制，在 AGENTS.md 中引用各 SKILL.md 文件路径

---

## 🐍 Python 依赖（可选）

如果需要使用 Word 文档生成/转换脚本：

```bash
cd game_doc_generator
pip install -r requirements.txt
```

核心依赖：
- `python-docx` — Word 文档生成
- Python 3.6+

**注意**：Python 依赖仅用于 Word 文档脚本。AI 使用 skill 本身**不需要**安装 Python。

---

## ✅ 验证清单

安装完成后，进行以下检查：

- [ ] 6 个 skill 文件夹都在 skills 目录下
- [ ] 每个文件夹中都有 `SKILL.md` 且包含正确的 YAML frontmatter
- [ ] `game_doc_generator` 中包含 scripts/、examples/、prompts/ 子目录
- [ ] AI 助手能识别到所有 skill（向助手询问："当前加载了哪些 skill？"）
- [ ] 测试："帮我生成一个签到功能的设计文档"

---

## 🔄 更新 Skill

更新时只需替换对应的 skill 文件夹即可，其他 skill 不受影响。这是拆分为多个独立 skill 的优势之一。

---

## 🔍 故障排除

### 问题1：AI 助手无法识别 skill

**排查步骤**：
1. 确认 skill 文件夹在正确的 skills 目录下
2. 检查 SKILL.md 是否存在且 frontmatter 格式正确
3. 尝试重启 AI 助手 / 编辑器
4. 确认文件编码为 UTF-8

### 问题2：skill 加载了但功能不完整

**原因**：可能只安装了部分 skill

**解决**：确保 6 个 skill 全部安装。`game_doc_generator` 会引用其他 skill，缺少任何一个都可能导致功能不完整。

### 问题3：scripts 脚本无法运行

**解决**：
```bash
python --version          # 确认 Python 已安装
pip install python-docx   # 安装依赖
```

**注意**：scripts 是可选工具，不影响 skill 核心功能。

---

## 📋 各环境 Skills 目录速查表

| 环境 | Workspace 路径 | 全局路径 | 安装方式 |
| --- | --- | --- | --- |
| Gemini Antigravity | `.agent/skills/` | `~/.gemini/antigravity/skills/` | 复制文件夹 |
| Claude Code | `.claude/skills/` | `~/.claude/skills/` | 复制文件夹 / ZIP上传 |
| Cursor | `.cursor/skills/` | 设置中查看 | 复制文件夹 |
| Windsurf | `.windsurf/skills/` | 设置中查看 | 复制文件夹 |
| VS Code + AI插件 | `.vscode/skills/` | 插件配置 | 视插件而定 |
| Codex | `.codex/skills/` | - | 视版本而定 |
