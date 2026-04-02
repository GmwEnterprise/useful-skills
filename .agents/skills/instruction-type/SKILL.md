---
name: instruction-type
description: instruction 型技能的创建规范。当你需要创建纯指令/知识描述型技能时加载此技能。适用于无代码、通过文本指令指导 AI 行为的场景。
---

# Instruction 型技能规范

instruction 型技能仅包含 SKILL.md，通过结构化的文本描述指导 AI agent 的行为。不包含任何可执行代码。

## 目录结构

```
skills/<skill-name>/
├── SKILL.md
└── refs/                          # 可选，参考文档
    └── *.md
```

## 适用场景

- 工作流程指导（如 brainstorming、debugging 流程）
- 编码规范与最佳实践
- 检查清单与审查标准
- 知识库与参考信息

## 编写要点

### 指令精确性

- 使用**明确的动词**：必须、应该、不要、禁止
- 避免**模糊表述**：可以用、或许、建议考虑
- 提供**决策条件**：当 X 时做 Y，否则做 Z

### 结构化呈现

- 用**流程图或步骤列表**描述执行顺序
- 用**表格**呈现对比或条件判断
- 用**代码块**展示输入/输出示例

### 错误预防

- 列出**常见误区**和需要避免的行为
- 提供**反面示例**说明什么不该做
- 设置**检查点**让 agent 在关键步骤自检

### 参考文档

复杂内容可拆分到 `refs/` 目录，SKILL.md 中引用：

```markdown
详见 refs/detailed-guide.md
```

## 示例结构

```markdown
---
name: my-workflow
description: ...
---

# My Workflow

## 何时使用
...

## 执行步骤
1. ...
2. ...

## 检查清单
- [ ] ...

## 常见误区
- ...

## 参考
详见 refs/advanced-patterns.md
```
