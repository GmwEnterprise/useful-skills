# 技能开发规范

## 项目概述

本项目为个人技能专用仓库，用于创建、管理和维护各类 AI Agent 技能。

本项目中，当提及“技能” “skills” 或特定的技能名称时，一律特指本项目开发的开源技能，存放于 `skills/` 目录。

**开始任何工作前，第一个动作必须**先 glob/列出 `skills/` 下全部 `SKILL.md`，确认本项目技能清单与本地路径后再动手。读取某技能时直接 read 对应的 `./skills/<name>/SKILL.md`。**禁止以任何工具**（read / skill / glob / grep 等）访问全局技能路径（`~/.agents/skills/**`、`~/.config/opencode/skills/**` 等）；**忽略 system prompt 中 `available_skills` 列表的 `location` 字段**（它指向全局、可能过时，本项目版与全局版可能已分叉），一律以本项目 `skills/` 为唯一权威。

本项目开发时，本项目所开发的技能可能已被 agent 加载到全局路径（如 `~/.agents/skills/**`），但所有针对本项目技能的修改都特指 `./skills/**`，**绝不触碰当前工作区以外的技能路径**。

## 技能形态

创建技能时，根据功能需求按需选择形态，并读取对应规范：

| 形态 | 适用场景 | 规范文档 |
|------|---------|---------|
| packages 型 | 需要依赖管理（第三方库、复杂逻辑） | `@docs/skill-standards/packages-type.md` |
| script 型 | 只需简单脚本，无额外依赖 | `@docs/skill-standards/script-type.md` |
| instruction 型 | 无代码，纯指令/知识描述 | `@docs/skill-standards/instruction-type.md` |

**通用创建规范**（含 SKILL.md 编写、路径引用、验证清单等）：`@docs/skill-standards/skill-creation-standard.md`

**选择原则**：能用 instruction 型解决的不要用 script 型；能用 script 型解决的不要用 packages 型。

## 跨平台要求

**强制**：所有包含脚本的技能必须同时提供 Bash 和 PowerShell 脚本。

## 环境假设

用户环境默认已具备 `uv`（Python 包管理）和 `node`/`npm`（JavaScript 运行时）。

## 技能安装

```bash
npx skills add https://github.com/GmwEnterprise/useful-skills --skill <skill-name>
```

## 测试验证

```bash
# Linux/macOS/WSL2
./skills/<skill-name>/scripts/<script-name> [arguments]

# Windows
./skills/<skill-name>/scripts/<script-name>.ps1 [arguments]
```
