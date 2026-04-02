---
name: script-type
description: script 型技能的创建规范。当你需要创建只需简单脚本、无额外依赖的技能时加载此技能。适用于轻量级工具、文件操作、系统命令等场景。
---

# Script 型技能规范

script 型技能仅包含独立脚本，无需依赖管理。脚本直接实现功能逻辑。

## 目录结构

```
skills/<skill-name>/
├── scripts/
│   ├── <script-name>              # Bash 脚本
│   └── <script-name>.ps1          # PowerShell 脚本
└── SKILL.md
```

## 适用场景

- 无第三方依赖的文件处理
- 调用系统命令或内置工具
- 简单的数据转换或文本处理
- Shell/PowerShell 原生能力即可完成的任务

## 脚本编写要点

1. **保存原始工作目录** — 若需要处理用户传入的相对路径，在脚本开头保存 `pwd`
2. **参数校验** — 无参数时输出 usage 并退出
3. **相对路径转绝对路径** — 基于原始工作目录转换用户传入的路径参数
4. **直接执行功能逻辑** — 无需 `cd` 到其他目录，无依赖安装步骤

## Bash 模板

```bash
#!/bin/bash
set -e

ORIGINAL_PWD="$(pwd)"

if [ $# -eq 0 ]; then
    echo "Usage: $(basename "$0") <required_arg> [optional_arg]"
    exit 1
fi

REQUIRED_ARG="$1"
OPTIONAL_ARG="${2:-}"

if [[ "$REQUIRED_ARG" != /* ]]; then
    REQUIRED_ARG="$ORIGINAL_PWD/$REQUIRED_ARG"
fi

if [ -n "$OPTIONAL_ARG" ] && [[ "$OPTIONAL_ARG" != /* ]]; then
    OPTIONAL_ARG="$ORIGINAL_PWD/$OPTIONAL_ARG"
fi

# 在此实现功能逻辑，使用 $REQUIRED_ARG 和 $OPTIONAL_ARG
```

## PowerShell 模板

```powershell
#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$OriginalPwd = Get-Location

if ($args.Count -eq 0) {
    Write-Host "Usage: $(Split-Path -Leaf $PSCommandPath) <required_arg> [optional_arg]"
    exit 1
}

$RequiredArg = $args[0]
$OptionalArg = if ($args.Count -gt 1) { $args[1] } else { "" }

if (-not [System.IO.Path]::IsPathRooted($RequiredArg)) {
    $RequiredArg = Join-Path $OriginalPwd $RequiredArg
}

if ($OptionalArg -and -not [System.IO.Path]::IsPathRooted($OptionalArg)) {
    $OptionalArg = Join-Path $OriginalPwd $OptionalArg
}

# 在此实现功能逻辑，使用 $RequiredArg 和 $OptionalArg
```

## 与 packages 型的区别

| 对比项 | script 型 | packages 型 |
|-------|----------|------------|
| 依赖管理 | 无 | uv/npm |
| 目录切换 | 不需要 | 需要 cd 到 packages/ |
| 依赖安装 | 无 | 静默安装 |
| 复杂度 | 轻量 | 完整项目结构 |

当脚本功能变复杂、需要引入第三方依赖时，应迁移为 packages 型。
