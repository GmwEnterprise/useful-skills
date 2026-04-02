---
name: docx-reader
description: 用于读取 Word 文档（.docx）提取内容为 Markdown 格式。当用户提到 Word、docx、文档读取时触发。支持提取文本（段落、标题、列表）、表格、图片、嵌入附件，输出 Markdown 文件和附件资源。
---

# Docx Reader

读取 .docx 文件并提取内容为 Markdown 格式，同时导出嵌入的图片和附件。

## 前置要求

**重要**：本技能依赖 `uv` 进行 Python 依赖管理。如果用户系统没有安装 `uv`，请停止任务执行并告知用户：

> 当前系统未安装 `uv`，无法执行此技能。请先安装 `uv`：
> ```bash
> curl -LsSf https://astral.sh/uv/install.sh | sh
> ```

## 何时使用

- 需要读取 .docx 文件内容
- 需要提取 Word 文档中的文本、表格、图片
- 需要将 Word 文档转换为 Markdown 格式
- 需要提取 Word 文档中嵌入的附件（xlsx、pdf 等）

## 快速开始

```bash
# Bash/Linux/macOS
scripts/docx-reader <docx_file> [output_dir]

# PowerShell/Windows
scripts/docx-reader.ps1 <docx_file> [output_dir]
```

## 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `docx_file` | 是 | Word 文件路径（.docx） |
| `output_dir` | 否 | 输出目录，默认为 `docs/extracted` |

## 输出结构

执行成功后，在输出目录下生成以下内容：

```
docs/extracted/
├── xxx.md                    # 提取的 Markdown 内容
└── xxx/                      # 附件资源目录
    ├── image1.png
    ├── image2.jpg
    ├── embedded.xlsx
    └── ...
```

**Markdown 文件**：包含文档的全部文本内容，保留标题层级、格式（粗体/斜体/删除线）、表格、图片引用和列表结构。文件头部包含 YAML 元数据（标题、作者、创建时间等）。

**附件目录**：以文档文件名命名，包含所有提取的图片和嵌入附件。

## 功能特性

### 文本提取

- 段落文本，保留粗体、斜体、删除线、下划线格式
- 标题层级（H1-H6）自动识别
- 列表项（有序/无序）识别
- 文本对齐（居中/右对齐）以 HTML 标签保留

### 表格支持

- 表格自动转为 Markdown 表格格式
- 第一行作为表头

### 图片提取

- 从文档中提取所有嵌入图片（png、jpg、gif、bmp、svg 等）
- 在 Markdown 中以相对路径引用：`![filename](assets_dir/filename)`
- 图片保存到附件目录

### 附件提取

- 提取嵌入的 OLE 对象附件（xlsx、pdf、pptx 等）
- 附件保存到附件目录

### 元数据

Markdown 文件头部包含 YAML front matter：

```yaml
---
title: "文档标题"
author: "作者"
created: "2024-01-01T00:00:00"
modified: "2024-06-01T00:00:00"
source: "document.docx"
---
```

## 常见问题

### 文件不存在

```
Error: File not found: /path/to/file.docx
```

检查文件路径是否正确。

### 不支持的格式

```
Error: Unsupported file format: .doc
```

只支持 `.docx` 格式。旧版 `.doc` 需先用 LibreOffice 转换：`soffice --headless --convert-to docx document.doc`

## 依赖要求

- Python 3.12+
- python-docx
- uv（用于依赖管理）

首次运行会自动初始化虚拟环境并安装依赖。

## 示例

```bash
# 使用默认输出目录 (docs/extracted)
scripts/docx-reader report.docx
# 输出:
# Success: report.docx
# Markdown: /current/path/docs/extracted/report.md
# Assets: /current/path/docs/extracted/report (3 files)

# 指定输出目录
scripts/docx-reader report.docx output

# 查看提取的 Markdown
cat docs/extracted/report.md
```
