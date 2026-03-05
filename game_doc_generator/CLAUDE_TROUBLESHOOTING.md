# Claude Skill 安装故障排除

## ❌ 常见错误：root directory name must contain only alphanumeric characters, underscores, and hyphens

### 问题原因

Claude 对 skill 根目录名称有严格限制：
- ✅ 允许：字母（a-z, A-Z）、数字（0-9）、下划线（_）、连字符（-）
- ❌ **不允许**：空格、其他特殊字符、中文字符

### 解决方案

当前 skill 目录名称均使用下划线命名（如 `game_doc_generator`），符合 Claude 规范。

如果仍遇到命名错误，可使用连字符版本：

```powershell
# 批量重命名为连字符版本
Rename-Item "game_doc_generator" "game-doc-generator"
Rename-Item "game_doc_template" "game-doc-template"
Rename-Item "game_rule_writing" "game-rule-writing"
Rename-Item "game_config_table" "game-config-table"
Rename-Item "game_doc_quality" "game-doc-quality"
Rename-Item "markdown_format_standard" "markdown-format-standard"
```

### 打包上传

Claude 需要以 zip 格式上传 skill，每个 skill 单独打包：

```powershell
$skills = @("game_doc_generator", "game_doc_template", "game_rule_writing", "game_config_table", "game_doc_quality", "markdown_format_standard")
foreach ($s in $skills) {
    Compress-Archive -Path ".\$s" -DestinationPath ".\$s.zip" -Force
}
```

然后在 Claude 中逐个上传每个 zip 文件。

### Claude vs Antigravity 差异

| 特性 | Claude | Antigravity |
| --- | --- | --- |
| 命名限制 | 更严格 | 较宽松 |
| 安装方式 | 上传 zip | 直接复制文件夹 |
| 路径要求 | 不能含空格 | 可以含空格 |
| 自动加载 | 上传后立即可用 | 自动扫描目录 |

### 其他常见问题

#### 问题：上传后提示"missing SKILL.md"

**原因**：zip 文件结构不正确，可能多了一层目录

**解决**：确保 zip 的根目录就是 skill 文件夹本身。

#### 问题：frontmatter 格式错误

检查 SKILL.md 开头是否有正确的 YAML frontmatter：
```yaml
---
name: Skill名称
description: Skill描述
---
```
