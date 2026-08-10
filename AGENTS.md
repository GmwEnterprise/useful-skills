# 技能开发规范

## 项目概述

本项目为个人技能专用仓库，用于创建、管理和维护各类 AI Agent 技能。

本项目中，当提及“技能” “skills”时，分两层含义:
- 本项目开发的开源技能，存放于 `skills/` 目录；
- 技能创建规范文档，存放于 `docs/skill-standards/` 目录，通过下方的 `@` 相对路径按需加载（创建/修改技能时再用 Read 工具读取，无需预加载）。

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
