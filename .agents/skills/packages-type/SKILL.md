---
name: packages-type
description: packages 型技能的创建规范。当你需要创建包含内嵌项目（packages/）的技能时加载此技能。适用于需要第三方依赖管理、复杂逻辑实现的场景。
---

# Packages 型技能规范

packages 型技能通过内嵌项目管理依赖和实现逻辑，脚本作为入口间接调用内嵌项目。

## 目录结构

```
skills/<skill-name>/
├── packages/<embedded_project>/
│   ├── main.py                    # 或 main.js，项目入口
│   ├── pyproject.toml             # 或 package.json，依赖声明
│   └── uv.lock                   # 或 package-lock.json，锁文件
├── scripts/
│   ├── <script-name>              # Bash 入口脚本
│   └── <script-name>.ps1          # PowerShell 入口脚本
└── SKILL.md
```

## 技术栈选择

| 场景 | 技术栈 |
|------|--------|
| Python 依赖 | Python (uv)，入口 `uv run python main.py` |
| JavaScript 依赖 | JavaScript (npm)，入口 `node main.js` |

## 脚本执行逻辑

所有 `scripts/` 下的脚本必须遵循以下执行流程：

1. **保存原始工作目录** — 用于后续相对路径转换
2. **参数校验** — 无参数时输出 usage 并退出
3. **相对路径转绝对路径** — 基于原始工作目录转换用户传入的路径参数
4. **切换到内嵌项目目录** — `cd` 到 `packages/<embedded_project>/`
5. **静默安装依赖** — `uv sync --quiet` 或 `npm install --silent`
6. **执行内嵌项目入口** — 传入已转换的绝对路径参数

### Bash 模板

```bash
#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PACKAGE_DIR="$(dirname "$SCRIPT_DIR")/packages/<embedded_project>"

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

cd "$PACKAGE_DIR"

# Python 项目
uv sync --quiet
uv run python main.py "$REQUIRED_ARG" ${OPTIONAL_ARG:+"$OPTIONAL_ARG"}

# JavaScript 项目（替换上面两行）
# npm install --silent
# node main.js "$REQUIRED_ARG" ${OPTIONAL_ARG:+"$OPTIONAL_ARG"}
```

### PowerShell 模板

```powershell
#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageDir = Join-Path (Split-Path $ScriptDir -Parent) "packages\<embedded_project>"

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

Push-Location $PackageDir

# Python 项目
uv sync --quiet
if ($OptionalArg) {
    uv run python main.py $RequiredArg $OptionalArg
} else {
    uv run python main.py $RequiredArg
}

# JavaScript 项目（替换上面几行）
# npm install --silent
# if ($OptionalArg) {
#     node main.js $RequiredArg $OptionalArg
# } else {
#     node main.js $RequiredArg
# }

Pop-Location
```

## 关键要点

- 脚本中的路径参数**必须**在步骤 3 转换为绝对路径，因为步骤 4 会 `cd` 到内嵌项目目录
- 依赖安装使用静默模式（`--quiet` / `--silent`），不输出安装日志
- 内嵌项目的依赖独立管理，不污染用户本机环境
- 每次执行都执行静默安装，确保环境一致性（`uv sync` 和 `npm install` 在依赖不变时几乎无开销）
