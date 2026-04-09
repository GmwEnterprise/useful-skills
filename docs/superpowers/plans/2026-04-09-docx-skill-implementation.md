# Docx Skill Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 `docx-reader` 技能重命名为 `docx`，保留 `.docx -> Markdown` 读取能力，并新增 `Markdown -> .docx` 创建能力与可配置样式。

**Architecture:** 保持 `packages` 型技能结构不变，统一目录、脚本和 Python 包命名为 `docx`。读取能力尽量复用现有逻辑；创建能力使用 Markdown 解析库生成结构化 token，再映射到 `python-docx` 文档对象，并通过默认配置、JSON 配置和 CLI 参数三层合并控制样式。

**Tech Stack:** Python 3.12, `python-docx`, `markdown-it-py`, `pytest`, Bash, PowerShell

---

## File Map

- Rename: `skills/docx-reader/` -> `skills/docx/`
- Rename: `skills/docx/scripts/docx-reader` -> `skills/docx/scripts/docx-read`
- Rename: `skills/docx/scripts/docx-reader.ps1` -> `skills/docx/scripts/docx-read.ps1`
- Create: `skills/docx/scripts/docx-write`
- Create: `skills/docx/scripts/docx-write.ps1`
- Rename: `skills/docx/packages/docx-reader/` -> `skills/docx/packages/docx/`
- Modify: `skills/docx/packages/docx/main.py`
- Modify: `skills/docx/packages/docx/pyproject.toml`
- Modify: `skills/docx/packages/docx/README.md`
- Modify: `skills/docx/SKILL.md`
- Create: `skills/docx/packages/docx/tests/test_reader.py`
- Create: `skills/docx/packages/docx/tests/test_writer.py`

### Task 1: 建立读取侧回归测试与创建侧失败测试

**Files:**
- Create: `skills/docx/packages/docx/tests/test_reader.py`
- Create: `skills/docx/packages/docx/tests/test_writer.py`
- Modify: `skills/docx/packages/docx/pyproject.toml`

- [ ] **Step 1: 添加测试依赖并创建测试骨架**

```toml
[dependency-groups]
dev = [
    "pytest>=8.3.0",
    "ruff>=0.8.0",
]
```

```python
# skills/docx/packages/docx/tests/test_reader.py
from pathlib import Path

from docx import Document

from main import docx_to_markdown


def test_docx_to_markdown_preserves_heading_and_formatting(tmp_path: Path):
    docx_path = tmp_path / "sample.docx"
    output_dir = tmp_path / "out"

    document = Document()
    document.add_heading("测试标题", level=1)
    paragraph = document.add_paragraph()
    paragraph.add_run("普通")
    bold_run = paragraph.add_run("加粗")
    bold_run.bold = True
    italic_run = paragraph.add_run("斜体")
    italic_run.italic = True
    document.save(docx_path)

    markdown = docx_to_markdown(str(docx_path), output_dir)

    assert "# 测试标题" in markdown
    assert "**加粗**" in markdown
    assert "*斜体*" in markdown
```

```python
# skills/docx/packages/docx/tests/test_writer.py
from pathlib import Path

import pytest

from main import markdown_to_docx


def test_markdown_to_docx_writes_headings_table_image_and_code_block(tmp_path: Path):
    markdown_path = tmp_path / "report.md"
    docx_path = tmp_path / "report.docx"
    image_path = tmp_path / "demo.png"

    image_path.write_bytes(
        bytes.fromhex(
            "89504E470D0A1A0A0000000D49484452000000010000000108060000001F15C489"
            "0000000D49444154789C6360000002000154A24F5D0000000049454E44AE426082"
        )
    )
    markdown_path.write_text(
        "# 标题\n\n"
        "正文含 **加粗**、*斜体*、~~删除线~~ 和 `code`。\n\n"
        "> 引用\n\n"
        "- 列表项\n\n"
        "---\n\n"
        "```python\nprint('hi')\n```\n\n"
        "| A | B |\n| --- | --- |\n| 1 | 2 |\n\n"
        f"![示例图]({image_path.name})\n",
        encoding="utf-8",
    )

    markdown_to_docx(markdown_path, docx_path)

    assert docx_path.exists()
```

- [ ] **Step 2: 运行读取测试，确认基线通过**

Run: `uv run pytest tests/test_reader.py::test_docx_to_markdown_preserves_heading_and_formatting -v`
Expected: PASS，证明目录重命名前已有读取逻辑可被回归测试覆盖。

- [ ] **Step 3: 运行创建测试，确认当前实现失败**

Run: `uv run pytest tests/test_writer.py::test_markdown_to_docx_writes_headings_table_image_and_code_block -v`
Expected: FAIL，报错为 `ImportError` 或 `AttributeError`，因为 `markdown_to_docx` 还不存在。

- [ ] **Step 4: 保留失败测试结果作为 RED 基线**

```text
FAILED tests/test_writer.py::test_markdown_to_docx_writes_headings_table_image_and_code_block
E   ImportError: cannot import name 'markdown_to_docx' from 'main'
```

- [ ] **Step 5: 提交测试基线（仅在用户明确要求提交时）**

```bash
git add skills/docx/packages/docx/pyproject.toml skills/docx/packages/docx/tests/test_reader.py skills/docx/packages/docx/tests/test_writer.py
git commit -m "test(docx): add reader regression and writer red tests"
```

### Task 2: 重命名技能目录与读取脚本入口

**Files:**
- Rename: `skills/docx-reader/` -> `skills/docx/`
- Rename: `skills/docx/packages/docx-reader/` -> `skills/docx/packages/docx/`
- Rename: `skills/docx/scripts/docx-reader` -> `skills/docx/scripts/docx-read`
- Rename: `skills/docx/scripts/docx-reader.ps1` -> `skills/docx/scripts/docx-read.ps1`
- Modify: `skills/docx/packages/docx/pyproject.toml`
- Modify: `skills/docx/packages/docx/README.md`

- [ ] **Step 1: 重命名目录与入口文件**

```bash
mv skills/docx-reader skills/docx
mv skills/docx/packages/docx-reader skills/docx/packages/docx
mv skills/docx/scripts/docx-reader skills/docx/scripts/docx-read
mv skills/docx/scripts/docx-reader.ps1 skills/docx/scripts/docx-read.ps1
```

- [ ] **Step 2: 更新 Python 包元数据**

```toml
[project]
name = "docx"
version = "0.1.0"
description = "Read .docx files into Markdown and write .docx files from Markdown"
readme = "README.md"
requires-python = ">=3.12"
dependencies = [
    "markdown-it-py>=3.0.0",
    "python-docx>=1.1.0",
]
```

- [ ] **Step 3: 更新 Bash 读取脚本路径**

```bash
#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PACKAGE_DIR="$(dirname "$SCRIPT_DIR")/packages/docx"

ORIGINAL_PWD="$(pwd)"

if [ $# -eq 0 ]; then
    echo "Usage: docx-read <docx_file> [output_dir]"
    exit 1
fi
```

- [ ] **Step 4: 更新 PowerShell 读取脚本路径**

```powershell
#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageDir = Join-Path (Split-Path $ScriptDir -Parent) "packages\docx"

if ($args.Count -eq 0) {
    Write-Host "Usage: docx-read.ps1 <docx_file> [output_dir]"
    exit 1
}
```

- [ ] **Step 5: 运行读取回归测试与脚本帮助检查**

Run: `uv run pytest tests/test_reader.py::test_docx_to_markdown_preserves_heading_and_formatting -v`
Expected: PASS

Run: `scripts/docx-read`
Expected: 输出 `Usage: docx-read <docx_file> [output_dir]`

- [ ] **Step 6: 提交重命名与读取入口调整（仅在用户明确要求提交时）**

```bash
git add skills/docx
git commit -m "refactor(docx): rename skill and reader entrypoints"
```

### Task 3: 用最小实现让 Markdown 写入测试转绿

**Files:**
- Modify: `skills/docx/packages/docx/main.py`
- Test: `skills/docx/packages/docx/tests/test_writer.py`

- [ ] **Step 1: 在 `main.py` 中添加最小创建入口与默认配置**

```python
DEFAULT_STYLE = {
    "body": {
        "font": "Microsoft YaHei",
        "ascii_font": "Calibri",
        "size": 11,
        "line_spacing": 1.5,
        "space_after": 6,
    },
    "headings": {
        "1": {"font": "Microsoft YaHei", "size": 20, "bold": True},
        "2": {"font": "Microsoft YaHei", "size": 18, "bold": True},
        "3": {"font": "Microsoft YaHei", "size": 16, "bold": True},
        "4": {"font": "Microsoft YaHei", "size": 14, "bold": True},
        "5": {"font": "Microsoft YaHei", "size": 13, "bold": True},
        "6": {"font": "Microsoft YaHei", "size": 12, "bold": True},
    },
    "image": {"max_width_inches": 5.8},
}


def markdown_to_docx(markdown_path: Path, output_path: Path, style_config: dict | None = None) -> None:
    document = Document()
    markdown_text = markdown_path.read_text(encoding="utf-8")
    for line in markdown_text.splitlines():
        if line.startswith("# "):
            document.add_heading(line[2:].strip(), level=1)
        elif line.startswith("## "):
            document.add_heading(line[3:].strip(), level=2)
        elif line.strip().startswith("- "):
            document.add_paragraph(line.strip()[2:], style="List Bullet")
        elif line.strip().startswith("> "):
            document.add_paragraph(line.strip()[2:], style="Intense Quote")
        elif line.strip():
            document.add_paragraph(line)
        else:
            document.add_paragraph("")
    document.save(output_path)
```

- [ ] **Step 2: 运行创建测试，确认仍然失败但失败点前移到具体能力缺失**

Run: `uv run pytest tests/test_writer.py::test_markdown_to_docx_writes_headings_table_image_and_code_block -v`
Expected: FAIL，错误转为断言失败，例如图片、表格或代码块未被正确写入。

- [ ] **Step 3: 为 Markdown 解析补充块级与行内映射实现**

```python
from markdown_it import MarkdownIt
from docx.shared import Inches, Pt


def build_markdown_parser() -> MarkdownIt:
    return MarkdownIt("commonmark", {"breaks": True}).enable("table").enable("strikethrough")


def render_inline(paragraph, tokens, style):
    for token in tokens:
        if token.type == "text":
            paragraph.add_run(token.content)
        elif token.type == "strong_open":
            continue
```

```python
def markdown_to_docx(markdown_path: Path, output_path: Path, style_config: dict | None = None) -> None:
    config = merge_style_config(DEFAULT_STYLE, style_config or {})
    parser = build_markdown_parser()
    tokens = parser.parse(markdown_path.read_text(encoding="utf-8"))
    document = Document()
    state = RenderState(markdown_path=markdown_path, document=document, config=config)
    render_tokens(tokens, state)
    document.save(output_path)
```

- [ ] **Step 4: 运行创建测试，确认转绿**

Run: `uv run pytest tests/test_writer.py::test_markdown_to_docx_writes_headings_table_image_and_code_block -v`
Expected: PASS

- [ ] **Step 5: 提交最小创建能力（仅在用户明确要求提交时）**

```bash
git add skills/docx/packages/docx/main.py skills/docx/packages/docx/tests/test_writer.py
git commit -m "feat(docx): add markdown to docx writer"
```

### Task 4: 完成样式配置、CLI 参数和写入脚本

**Files:**
- Modify: `skills/docx/packages/docx/main.py`
- Create: `skills/docx/scripts/docx-write`
- Create: `skills/docx/scripts/docx-write.ps1`
- Test: `skills/docx/packages/docx/tests/test_writer.py`

- [ ] **Step 1: 为样式配置添加合并函数与参数解析测试**

```python
def merge_style_config(defaults: dict, overrides: dict) -> dict:
    merged = dict(defaults)
    for key, value in overrides.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = merge_style_config(merged[key], value)
        else:
            merged[key] = value
    return merged
```

```python
def test_cli_style_overrides_json_config(tmp_path: Path):
    config_path = tmp_path / "style.json"
    config_path.write_text('{"body": {"size": 10}, "headings": {"1": {"size": 18}}}', encoding="utf-8")

    parsed = parse_write_options([
        "input.md",
        "output.docx",
        "--style-config",
        str(config_path),
        "--body-size",
        "12",
        "--h1-size",
        "22",
    ])

    assert parsed["body"]["size"] == 12
    assert parsed["headings"]["1"]["size"] == 22
```

- [ ] **Step 2: 运行新测试，确认失败**

Run: `uv run pytest tests/test_writer.py::test_cli_style_overrides_json_config -v`
Expected: FAIL，报错为 `NameError: parse_write_options is not defined`。

- [ ] **Step 3: 添加 `write` 模式命令行解析**

```python
def parse_write_options(argv: list[str]) -> dict:
    parser = argparse.ArgumentParser(prog="main.py write")
    parser.add_argument("input_md")
    parser.add_argument("output_docx", nargs="?")
    parser.add_argument("--style-config")
    parser.add_argument("--body-font")
    parser.add_argument("--body-size", type=float)
    parser.add_argument("--h1-font")
    parser.add_argument("--h1-size", type=float)
    parser.add_argument("--h2-font")
    parser.add_argument("--h2-size", type=float)
    parser.add_argument("--image-max-width", type=float)
    args = parser.parse_args(argv)
    return build_style_overrides(args)
```

- [ ] **Step 4: 添加 Bash 与 PowerShell 写入脚本**

```bash
#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PACKAGE_DIR="$(dirname "$SCRIPT_DIR")/packages/docx"
ORIGINAL_PWD="$(pwd)"

if [ $# -eq 0 ]; then
    echo "Usage: docx-write <input.md> [output.docx] [options]"
    exit 1
fi

INPUT_MD="$1"
shift

if [[ "$INPUT_MD" != /* ]]; then
    INPUT_MD="$ORIGINAL_PWD/$INPUT_MD"
fi

cd "$PACKAGE_DIR"
uv sync --quiet
uv run python main.py write "$INPUT_MD" "$@"
```
```

```powershell
#!/usr/bin/env pwsh
$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageDir = Join-Path (Split-Path $ScriptDir -Parent) "packages\docx"
$OriginalPwd = Get-Location

if ($args.Count -eq 0) {
    Write-Host "Usage: docx-write.ps1 <input.md> [output.docx] [options]"
    exit 1
}

$InputMd = $args[0]
$RemainingArgs = if ($args.Count -gt 1) { $args[1..($args.Count - 1)] } else { @() }

if (-not [System.IO.Path]::IsPathRooted($InputMd)) {
    $InputMd = Join-Path $OriginalPwd $InputMd
}

Push-Location $PackageDir
uv sync --quiet
uv run python main.py write $InputMd @RemainingArgs
Pop-Location
```

- [ ] **Step 5: 运行样式覆盖测试与脚本帮助检查**

Run: `uv run pytest tests/test_writer.py::test_cli_style_overrides_json_config -v`
Expected: PASS

Run: `scripts/docx-write`
Expected: 输出 `Usage: docx-write <input.md> [output.docx] [options]`

- [ ] **Step 6: 提交样式与脚本支持（仅在用户明确要求提交时）**

```bash
git add skills/docx/packages/docx/main.py skills/docx/scripts/docx-write skills/docx/scripts/docx-write.ps1
git commit -m "feat(docx): add configurable writer cli"
```

### Task 5: 更新技能文档并完成端到端验证

**Files:**
- Modify: `skills/docx/SKILL.md`
- Modify: `skills/docx/packages/docx/README.md`

- [ ] **Step 1: 更新技能元数据与使用说明**

```yaml
---
name: docx
description: 用于读取和创建 Word 文档（.docx）。读取时提取为 Markdown 并导出附件；创建时从本地 Markdown 生成带样式的 docx，支持标题、正文、表格、图片与常见文本样式配置。
---
```

```markdown
## 功能一：读取 Docx

```bash
scripts/docx-read <input.docx> [output_dir]
```

## 功能二：创建 Docx

```bash
scripts/docx-write <input.md> [output.docx] [--style-config style.json]
```

当用户未提供 Markdown，而是要求生成一份 docx 报告时，先在 `docs/` 或 `.tmp/` 生成 Markdown，再调用 `scripts/docx-write`。
```

- [ ] **Step 2: 更新包 README 示例**

```markdown
# docx

## 读取

```bash
scripts/docx-read report.docx
```

## 创建

```bash
scripts/docx-write report.md report.docx
scripts/docx-write report.md report.docx --body-size 12 --h1-size 22
scripts/docx-write report.md report.docx --style-config .tmp/docx-style.json
```
```

- [ ] **Step 3: 运行完整验证**

Run: `uv run pytest -v`
Expected: 全部 PASS

Run: `uv run ruff check`
Expected: `All checks passed!`

Run: `scripts/docx-write ../../../../examples/sample.md ../../../../.tmp/sample.docx`
Expected: 生成 `.tmp/sample.docx`

Run: `scripts/docx-read ../../../../.tmp/sample.docx ../../../../.tmp/docx-readback`
Expected: 生成 `.tmp/docx-readback/sample.md`

- [ ] **Step 4: 记录实际验证结果并整理未覆盖风险**

```text
Verified:
- Reader regression test passed
- Writer structure/style test passed
- CLI override test passed
- Ruff passed
- End-to-end write/read commands passed

Residual risks:
- 深层嵌套列表样式仍只覆盖常见层级
- 真正的 Word 超链接关系若未实现，则链接为文本表现
```

- [ ] **Step 5: 提交文档与验证结果（仅在用户明确要求提交时）**

```bash
git add skills/docx/SKILL.md skills/docx/packages/docx/README.md
git commit -m "docs(docx): document reader and writer workflows"
```

## Self-Review

- Spec coverage: 已覆盖重命名、读取保留、Markdown 创建、默认样式、JSON 与 CLI 样式覆盖、图片表格文本样式支持、文档与验证。
- Placeholder scan: 未使用 `TODO`、`TBD`、`similar to` 等占位描述；所有任务都包含文件、命令和代码片段。
- Type consistency: 计划中统一使用 `markdown_to_docx`、`merge_style_config`、`parse_write_options`、`render_tokens` 作为创建侧核心接口，命名前后一致。
