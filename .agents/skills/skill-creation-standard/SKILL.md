---
name: skill-creation-standard
description: 创建技能时的通用规范。当你需要创建新技能或修改技能结构时加载此技能。包含技能形态选择、目录结构、SKILL.md 编写规范、路径引用规则等。
---

# 技能创建通用规范

## 技能形态选择

创建技能前，根据功能需求选择合适的形态：

| 形态 | 适用场景 | 详见 |
|------|---------|------|
| **packages 型** | 需要依赖管理（第三方库、复杂逻辑） | 加载 `packages-type` 技能 |
| **script 型** | 只需简单脚本，无额外依赖 | 加载 `script-type` 技能 |
| **instruction 型** | 无代码，纯指令/知识描述 | 加载 `instruction-type` 技能 |

**选择原则**：能用 instruction 型解决的不要用 script 型；能用 script 型解决的不要用 packages 型。

## 通用目录结构

```
skills/<skill-name>/
├── SKILL.md                       # 必需，技能说明文档
├── packages/<embedded_project>/   # 可选，packages 型专用
├── scripts/                       # 可选，packages 型和 script 型使用
│   ├── <script-name>              # Bash (Linux/macOS/WSL2)
│   └── <script-name>.ps1          # PowerShell (Windows)
└── refs/                          # 可选，参考文档
    └── *.md
```

## SKILL.md 编写规范

### 元数据头

每个 SKILL.md 必须以 YAML front matter 开头：

```yaml
---
name: <skill-name>
description: <一句话描述技能的触发条件、功能和使用场景>
---
```

`description` 应包含：
- 技能功能摘要
- 触发关键词/场景
- 核心输出或行为

### 内容结构

按需包含以下章节（不适用的章节可省略）：

```markdown
# 技能名称

## 前置要求（如有）
## 何时使用
## 快速开始
## 参数说明（如有）
## 输出格式（如有）
## 功能特性
## 常见问题（如有）
## 依赖要求（如有）
## 示例
```

### 编写原则

1. **面向 AI 消费**：SKILL.md 的主要读者是 AI agent，内容应精确、无歧义
2. **渐进式披露友好**：重要信息放在前面，细节向后排列
3. **包含错误场景**：列出常见错误及处理方式，帮助 agent 自行排查
4. **路径引用使用无前缀格式**：详见下方路径规范

## 路径引用规范

SKILL.md 中引用脚本或文档时，**不要使用 `./` 前缀**。

opencode 采用渐进式披露机制，技能安装后的实际目录为 `~/.agents/skills/<skill-name>/`，`./` 前缀会指向用户当前工作目录而非技能目录。

| 引用类型 | 正确写法 | 错误写法 |
|---------|---------|---------|
| 脚本调用 | `scripts/xxx <args>` | `./scripts/xxx <args>` |
| 参考文档 | `refs/xxx.md` | `./refs/xxx.md` |

## 跨平台要求

**强制**：所有包含脚本的技能必须同时提供：
- Bash 脚本（`scripts/<name>`）— 用于 Linux/macOS/WSL2
- PowerShell 脚本（`scripts/<name>.ps1`）— 用于 Windows

## 环境假设

用户环境默认已具备：
- `uv` — Python 包管理
- `node` / `npm` — JavaScript 运行时与包管理

简单技能无需创建脚本，直接通过 SKILL.md 描述即可。

## 技能安装

```bash
npx skills add https://github.com/GmwEnterprise/useful-skills --skill <skill-name>
```

## 验证清单

创建技能后，逐项验证：

- [ ] SKILL.md 元数据头完整（name、description）
- [ ] SKILL.md 路径引用无 `./` 前缀
- [ ] 脚本正常执行，输出符合预期（如有脚本）
- [ ] 必选参数正常工作（如有参数）
- [ ] 可选参数正常工作（如有可选参数）
- [ ] 错误处理符合预期（参数缺失、无效输入等）
- [ ] Bash 和 PowerShell 脚本均已提供（如有脚本）
