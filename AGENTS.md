# 技能开发规范

## 项目概述

本项目为个人技能专用仓库，用于创建、管理和维护各类 AI Agent 技能。所有技能统一存放在 `skills/` 目录下。

## 目录结构

```
skills/<skill-name>/
├── packages/<embedded_project>/   # 内嵌项目（可选）
├── scripts/                       # 调用脚本
│   ├── <script-name>              # Bash (Linux/macOS)
│   └── <script-name>.ps1          # PowerShell (Windows)
└── SKILL.md                       # 技能说明文档
```

## 开发规范

### 前置条件

- **必须**：创建技能前先加载 `skill-creator` 技能，确认所有创建细节
- **环境假设**：用户环境已具备 `uv` 和 `nodejs`

### 技术栈选择

| 场景 | 推荐技术栈 |
|------|-----------|
| 需要依赖管理 | Python (uv) 或 JavaScript (npm) |
| 无额外依赖 | 直接使用 `.py` 或 `.js` 文件 |

### 开发原则

1. **内嵌项目优先**：技能功能优先基于 `packages/<embedded_project>/` 实现
2. **脚本间接调用**：内嵌项目通过 `scripts/` 下的脚本间接调用
3. **跨平台支持**：同时提供 Bash 和 PowerShell 脚本
4. **换行符规范**：Bash/Python 脚本必须使用 LF 换行符，PowerShell 脚本使用 CRLF

### 换行符问题

在 WSL2/Windows 混合环境下，Git 默认会将 LF 转换为 CRLF，导致 Bash 脚本无法执行：

```
/bin/bash: line 1: script.sh: cannot execute: required file not found
```

**原因**：CRLF 换行符使 shebang 变成 `#!/bin/bash\r`，系统找不到该解释器。

**解决方案**：项目根目录必须包含 `.gitattributes` 文件：

```gitattributes
* text=auto
*.sh text eol=lf
*.ps1 text eol=crlf
*.py text eol=lf
```

**检查方法**：
```bash
file scripts/<script-name>
# 正确: ASCII text executable
# 错误: ASCII text executable, with CRLF line terminators
```

### 依赖管理

脚本必须包含依赖自动检查与初始化逻辑：

- **Python 项目**：检查 `uv.lock` 或虚拟环境是否存在，若不存在则执行 `uv sync`
- **JavaScript 项目**：检查 `node_modules` 是否存在，若不存在则执行 `npm install`

**目的**：
- 用户无需手动初始化，技能开箱即用
- 内嵌项目独立管理依赖，不污染用户本机环境

**示例（Bash）**：
```bash
# Python 项目
cd "$(dirname "$0")/../packages/<embedded_project>"
if [ ! -f "uv.lock" ] || [ ! -d ".venv" ]; then
    uv sync
fi
uv run python main.py "$@"

# JavaScript 项目
cd "$(dirname "$0")/../packages/<embedded_project>"
if [ ! -d "node_modules" ]; then
    npm install
fi
node main.js "$@"
```

## 技能安装

```bash
npx skills add https://github.com/GmwEnterprise/useful-skills --skill <skill-name>
```

## 测试验证

### 测试命令

```bash
# Linux/macOS
./skills/<skill-name>/scripts/<script-name> [arguments]

# Windows
./skills/<skill-name>/scripts/<script-name>.ps1 [arguments]
```

### 验证清单

- [ ] 脚本正常执行，输出符合预期
- [ ] 必选参数正常工作
- [ ] 可选参数正常工作
- [ ] 错误处理符合预期（参数缺失、无效输入等）
