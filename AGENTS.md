# 技能开发规范

## 项目概述

本项目为个人技能专用仓库，用于创建、管理和维护各类 AI Agent 技能。

本项目中，当提及“技能” “skills”时，分两层含义:
- 本项目 opencode 可以使用的技能，存放于 `.agents/skills/` 目录；
- 本项目开发的开源技能，存放于 `skills/` 目录。

## 技能形态

创建技能时，根据功能需求按需选择形态：

| 形态 | 适用场景 | 规范文档 |
|------|---------|---------|
| packages 型 | 需要依赖管理（第三方库、复杂逻辑） | 加载 `packages-type` 技能 |
| script 型 | 只需简单脚本，无额外依赖 | 加载 `script-type` 技能 |
| instruction 型 | 无代码，纯指令/知识描述 | 加载 `instruction-type` 技能 |

**通用创建规范**（含 SKILL.md 编写、路径引用、验证清单等）：加载 `skill-creation-standard` 技能。

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
