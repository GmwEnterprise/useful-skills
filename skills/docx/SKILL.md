---
name: docx
description: 用于读取和创建 Word 文档（.docx）。读取时提取为 Markdown 并导出附件；创建时从本地 Markdown 生成带样式的 docx，支持标题、正文、表格、图片与常见文本样式配置。
---

# Docx

读取 `.docx` 为 Markdown，并从本地 Markdown 生成 `.docx` 文档。

## 前置要求

本技能依赖 `uv` 管理 Python 依赖。若系统未安装 `uv`，应先安装：

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

## 何时使用

- 需要将 `.docx` 读取为 Markdown 供 AI 继续处理
- 需要把 Markdown 报告输出为 `.docx`
- 需要保留图片、表格和常见 Markdown 文本样式
- 需要通过默认样式、CLI 参数或 JSON 配置控制 Word 文档格式

## 功能一：读取 Docx

```bash
# Bash/Linux/macOS
scripts/docx-read <input.docx> [output_dir]

# PowerShell/Windows
scripts/docx-read.ps1 <input.docx> [output_dir]
```

### 读取输出

- `<output_dir>/<filename>.md`
- `<output_dir>/<filename>/` 附件目录

默认输出目录：`docs/extracted`

## 功能二：创建 Docx

```bash
# Bash/Linux/macOS
scripts/docx-write <input.md> [output.docx] [--style-config style.json]

# PowerShell/Windows
scripts/docx-write.ps1 <input.md> [output.docx] [--style-config style.json]
```

### 创建支持范围

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

### 默认样式

未提供样式要求时，默认使用一套通用中文报告格式：

- 页面：A4
- 正文：`Microsoft YaHei` 11pt，1.5 倍行距
- 标题：H1-H6 逐级递减字号并加粗
- 代码块：等宽字体、浅灰底色
- 表格：首行加粗、浅色表头、细边框
- 图片：按页面宽度限制缩放

### 样式覆盖

简单覆盖优先使用 CLI 参数：

```bash
scripts/docx-write report.md report.docx --body-size 12 --h1-size 22
```

复杂样式使用 JSON 配置文件：

```bash
scripts/docx-write report.md report.docx --style-config .tmp/docx-style.json
```

合并优先级：CLI > JSON > 默认值。

## AI 使用约定

若用户没有直接提供 Markdown，而是要求生成一份 docx 报告：

1. 先在 `docs/` 或 `.tmp/` 生成 Markdown 文件
2. 默认直接使用内置样式
3. 用户明确要求字体、字号、图片或表格格式时，再生成 JSON 配置文件或追加 CLI 参数
4. 最后调用 `scripts/docx-write` 生成 `.docx`

## 依赖要求

- Python 3.12+
- `python-docx`
- `markdown-it-py`
- `uv`

## 示例

```bash
# 读取 docx
scripts/docx-read report.docx

# 创建 docx
scripts/docx-write report.md report.docx

# 使用样式配置
scripts/docx-write report.md report.docx --style-config .tmp/docx-style.json
```
