---
title: docx 技能重命名与创建能力设计
date: 2026-04-09
status: draft
---

# docx 技能重命名与创建能力设计

## 背景

现有 `skills/docx-reader/` 是一个 packages 型技能，已支持读取 `.docx` 文档并提取为 Markdown，同时导出图片和嵌入附件。

本次改造目标：

1. 将技能名从 `docx-reader` 重命名为 `docx`
2. 保留现有 `.docx -> Markdown` 的读取能力
3. 新增 `Markdown -> .docx` 的创建能力
4. 创建能力默认使用一套通用中文报告样式
5. 支持通过命令行参数和 JSON 配置文件覆盖标题、正文、表格、图片等样式
6. 当用户未提供 Markdown，但要求生成 docx 报告时，AI 先生成 Markdown，再调用脚本转为 docx

## 目标与非目标

### 目标

- 技能目录、SKILL 元数据、脚本入口统一改名为 `docx`
- 读取能力行为延续现有实现，尽量减少回归风险
- 创建能力基于本地 Markdown 文件生成可编辑的 Word 文档
- Markdown 创建侧支持常见结构与文本样式，包括：
  - 标题 1-6
  - 普通段落
  - 加粗、斜体、粗斜体
  - 删除线
  - 行内代码
  - 引用块
  - 无序列表、有序列表
  - 分隔线
  - 代码块
  - 表格
  - 图片
  - 链接
  - 软换行、硬换行
- 默认输出适合中文报告场景的文档版式
- 复杂样式通过 JSON 配置文件描述，简单调整通过 CLI 参数覆盖

### 非目标

- 不追求浏览器级 HTML 渲染完全保真
- 不保证兼容所有 Markdown 方言及其扩展语法
- 不在首版支持复杂内联 HTML 渲染
- 不在首版支持复杂嵌套表格、Mermaid、公式渲染等高复杂度扩展

## 目录与命名调整

技能目录从：

```text
skills/docx-reader/
```

调整为：

```text
skills/docx/
```

内部结构保持 packages 型：

```text
skills/docx/
├── SKILL.md
├── packages/docx/
│   ├── main.py
│   ├── pyproject.toml
│   ├── README.md
│   └── uv.lock
└── scripts/
    ├── docx-read
    ├── docx-read.ps1
    ├── docx-write
    └── docx-write.ps1
```

命名规则：

- 技能名：`docx`
- 读取脚本：`docx-read`
- 创建脚本：`docx-write`
- 内嵌 Python 包目录：`packages/docx`

这样可以避免技能名与单一能力名不一致的问题，同时保留清晰的双入口。

## 用户工作流

### 读取 `.docx`

```bash
scripts/docx-read <input.docx> [output_dir]
```

输出：

- `<output_dir>/<filename>.md`
- `<output_dir>/<filename>/` 附件目录

### 创建 `.docx`

```bash
scripts/docx-write <input.md> [output.docx] [--style-config style.json] [--body-size 12] [--h1-size 22]
```

输出：

- 指定输出路径的 `.docx`
- 未指定时默认在 Markdown 同目录生成同名 `.docx`

### AI 使用约定

当用户目标是“生成 docx 报告”而非“转换已有 Markdown 文件”时：

1. AI 先将内容整理为本地 Markdown 文件
2. 默认写入 `docs/` 或 `.tmp/`
3. 再调用 `scripts/docx-write` 生成 `.docx`
4. 若用户未指定样式，直接使用默认样式
5. 若用户指定字体、字号、表格或图片规则，再生成 JSON 配置文件或追加 CLI 覆盖参数

## 读取能力设计

读取能力沿用当前实现，主要做以下改动：

- 路径重命名：`docx-reader` -> `docx`
- 脚本重命名：`docx-reader` -> `docx-read`
- 文档更新：SKILL.md、README、示例、安装方式同步更新
- 尽量不改动既有 `.docx -> Markdown` 解析逻辑，避免引入无关回归

读取侧仍支持：

- 段落文本提取
- 标题层级识别
- 列表识别
- 粗体、斜体、删除线、下划线等基础文本样式
- Markdown 表格输出
- 图片导出与引用
- 嵌入附件导出
- 文档元数据输出

## 创建能力设计

### 总体策略

创建能力采用“先解析 Markdown 结构，再映射到 `python-docx`”的策略，而不是依赖正则直接拼装段落。

核心原因：

- 需要支持图片、表格和常见行内样式组合
- 需要稳定处理块级结构与嵌套关系
- 需要便于后续扩展样式配置

实现上，建议引入一个 Python Markdown 解析库，将 Markdown 先转为结构化 token/AST，再逐项映射为 Word 文档对象。

### Markdown 支持范围

创建侧首版必须支持以下 Markdown 元素：

- 标题 `#` - `######`
- 普通段落
- 加粗
- 斜体
- 粗斜体
- 删除线
- 行内代码
- 链接
- 引用块
- 无序列表
- 有序列表
- 分隔线
- 代码块
- 表格
- 图片
- 软换行
- 硬换行

### 元素映射

#### 标题

- `#` 到 `######` 分别映射到 Word Heading 1-6
- 每一级标题都可配置字体、字号、加粗、颜色、段前段后间距

#### 正文段落

- 普通 Markdown 段落映射到正文样式
- 支持 run 级别文本样式混排
- 换行按 Markdown 语义写入 Word 段落或手动换行

#### 行内样式

- `**text**` -> 加粗
- `*text*` -> 斜体
- `***text***` -> 粗斜体
- `` `code` `` -> 行内代码字体与底色
- `~~text~~` -> 删除线
- `[text](url)` -> 保留为链接文本与 URL 表现；若实现成本可控，可进一步生成 Word 真正超链接

#### 引用块

- 作为特殊段落样式写入
- 使用左缩进、可选斜体、可选浅灰色边框/底色增强辨识度

#### 列表

- 无序列表映射为 Word 项目符号列表
- 有序列表映射为 Word 编号列表
- 保留常见层级缩进

#### 分隔线

- 通过段落边框或视觉分隔段落实现

#### 代码块

- 作为独立段落块输出
- 使用等宽字体
- 支持背景色、边框、较小字号、单倍行距
- 保留原有换行与缩进

#### 表格

- Markdown 表格映射为 Word 表格
- 首行支持表头样式
- 支持边框、对齐和基础列宽控制

#### 图片

- 支持本地相对路径和绝对路径
- 路径解析相对 Markdown 文件所在目录进行
- 图片插入前可按页面宽度或最大宽度自动缩放
- 图片不存在时明确报错

## 默认样式设计

默认提供一套“通用中文报告”样式，适用于大多数无特殊要求的业务文档。

### 页面

- 纸张：A4
- 页边距：Word 常规页边距

### 正文

- 中文字体：`Microsoft YaHei`
- 西文字体：`Calibri`
- 字号：11pt
- 行距：1.5
- 段后：6pt

### 标题

- H1：20pt，加粗
- H2：18pt，加粗
- H3：16pt，加粗
- H4：14pt，加粗
- H5：13pt，加粗
- H6：12pt，加粗

### 引用块

- 左缩进
- 浅灰强调样式

### 代码块

- 等宽字体
- 浅灰底色
- 单倍行距

### 表格

- 表头加粗
- 浅色表头背景
- 细边框

### 图片

- 默认按页面内容宽度比例缩放
- 避免超出版心

默认样式目标是“开箱即用即可交付”，而不是追求高度装饰化。

## 样式配置设计

### 配置来源

样式由三层合并：

1. 内置默认配置
2. `--style-config <json>` 指定的 JSON 文件
3. CLI 单项覆盖参数

优先级：CLI > JSON > 默认值。

### CLI 常用覆盖参数

首版建议支持以下常用参数：

- `--style-config <path>`
- `--body-font <font>`
- `--body-size <pt>`
- `--h1-font <font>`
- `--h1-size <pt>`
- `--h2-font <font>`
- `--h2-size <pt>`
- `--table-style <style>`
- `--image-max-width <inches>`

必要时可继续补充 H3-H6 等级参数，但应避免一次性暴露过多低频选项。

### JSON 配置结构

复杂样式建议通过 JSON 文件描述。示例：

```json
{
  "page": {
    "size": "A4",
    "margin_top": 2.54,
    "margin_bottom": 2.54,
    "margin_left": 3.18,
    "margin_right": 3.18
  },
  "body": {
    "font": "Microsoft YaHei",
    "ascii_font": "Calibri",
    "size": 11,
    "line_spacing": 1.5,
    "space_after": 6
  },
  "headings": {
    "1": { "font": "Microsoft YaHei", "size": 20, "bold": true },
    "2": { "font": "Microsoft YaHei", "size": 18, "bold": true }
  },
  "blockquote": {
    "left_indent": 0.74,
    "italic": false
  },
  "code_block": {
    "font": "Consolas",
    "size": 10,
    "background_color": "F3F3F3"
  },
  "table": {
    "header_bold": true,
    "header_bg_color": "D9EAF7",
    "border": true
  },
  "image": {
    "max_width_inches": 5.8
  }
}
```

配置原则：

- 用户无要求时不必生成 JSON
- 只有复杂样式需求才生成 JSON 配置文件
- 临时配置文件放在 `.tmp/`

## 脚本接口设计

### 读取脚本

```bash
scripts/docx-read <input.docx> [output_dir]
```

规则：

- `input.docx` 必填
- `output_dir` 可选
- 未指定时默认输出到 `docs/extracted`

### 创建脚本

```bash
scripts/docx-write <input.md> [output.docx] [--style-config style.json] [--body-size 12] [--h1-size 22]
```

规则：

- `input.md` 必填
- `output.docx` 可选，默认与 Markdown 同目录同名
- `--style-config` 可选
- 其他样式参数可选

### 路径规则

- 脚本入口负责将用户输入的相对路径转换为绝对路径
- 创建脚本中的图片相对路径，以输入 Markdown 文件所在目录为基准解析
- 所有临时中间文件写入 `.tmp/`

## 错误处理

### 读取侧

- 文件不存在
- 文件扩展名不是 `.docx`
- 文档损坏或无法解析

### 创建侧

- Markdown 文件不存在
- 输入不是 `.md`
- JSON 配置文件不存在或格式错误
- 图片路径不存在
- 不支持的 Markdown 结构回退为纯文本或明确报错（按结构类型决定）

错误输出应保持简单明确，优先包含：

- 出错类型
- 出错路径或字段
- 可操作的修复建议

## 测试与验证策略

本次实现遵循 TDD。

### 读取侧回归验证

- 现有 `.docx -> Markdown` 基本用例继续通过
- 重命名后脚本与包路径仍能正常工作

### 创建侧新增验证

至少覆盖以下场景：

- Markdown 标题映射到 Heading 1-6
- 正文与加粗/斜体/删除线/行内代码的 run 样式写入
- 引用块样式写入
- 列表创建
- 分隔线创建
- 代码块创建
- 表格创建
- 图片插入
- 默认样式生效
- JSON 配置覆盖默认值
- CLI 参数覆盖 JSON 值

### 跨平台脚本验证

- Bash 脚本路径转换正确
- PowerShell 脚本路径转换正确
- 无 `uv` 时给出明确错误提示

### 编译与静态检查

- `uv sync --quiet`
- `uv run ruff check`
- 关键命令的真实执行验证

## 实施原则

- 尽量复用现有读取逻辑
- 为创建能力新增最小必要依赖
- Python 代码组织优先保持在单入口文件内，只有在明显需要时再拆分辅助函数
- 技能文档中明确“先生成 Markdown，再转换为 docx”的推荐工作流
- 只在用户明确要求时生成样式 JSON 文件

## 风险与缓解

### Markdown 解析复杂度上升

风险：如果直接用正则处理 Markdown，样式嵌套、表格和列表很容易出错。

缓解：使用结构化 Markdown 解析库，将块级与行内元素分开处理。

### Word 样式表达能力与 Markdown 语义不完全对齐

风险：部分扩展语法很难一比一映射为 Word。

缓解：首版只承诺常见 Markdown 结构，超出范围时明确降级策略。

### 图片与表格的版面控制

风险：大图或宽表格容易破坏版面。

缓解：图片默认限制最大宽度，表格采用基础样式并优先保证可读性。

## 结论

本次改造将 `docx-reader` 升级为统一的 `docx` 技能：

- 保留既有读取能力
- 新增基于本地 Markdown 的 docx 创建能力
- 提供默认中文报告样式
- 通过 CLI + JSON 配置实现样式覆盖
- 将“先生成 Markdown，再转 docx”作为标准 AI 工作流

该方案与仓库的 packages 型技能规范、跨平台脚本要求和工作目录规范保持一致，能够在最小化回归风险的前提下扩展出完整的双向 docx 工作流。
