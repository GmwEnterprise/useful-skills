# Feature Route Map

## Project Overview

- Application type: opencode 技能仓库；开源技能位于 `skills/<name>/`，每个技能含 `packages/<project>/`（uv 内嵌项目）、`scripts/`（Bash + PowerShell 跨平台入口）与 `SKILL.md`
- Main entry: `skills/<skill>/SKILL.md`（技能文档）；代码入口 `skills/<skill>/packages/<project>/main.py`
- Core directories: `skills/`（开源技能）、`.agents/skills/`（本项目 opencode 用技能与开发规范）
- Test entry: `skills/<skill>/packages/<project>/tests/`（`uv run pytest`）

## Module Index

- pptx: 读取 PPTX、更新 PPTX

## pptx

### 读取 PPTX

- Description: 将 `.pptx` 解析为 Markdown + JSON（文本 / 表格 / 备注 / 形状结构）
- Entry: `skills/pptx/scripts/pptx-reader`, `skills/pptx/scripts/pptx-reader.ps1`
- Core: `skills/pptx/packages/pptx/main.py`
- Tests: `skills/pptx/packages/pptx/tests/test_reader.py`
- Notes: 输出 `<name>.pptx_reader.md` 与 `.json`；不导出图片二进制，JSON 不含 run 级 `color`/`font`

### 更新 PPTX

- Description: 基于 JSON 指令更新已有 `.pptx`，9 种操作——形状级 `replace_text`/`add_picture`/`add_textbox`/`delete_shape`/`format_shape`/`move_shape`；幻灯片级 `insert_slide`/`move_slide`/`delete_slide`
- Entry: `skills/pptx/scripts/pptx-updater`, `skills/pptx/scripts/pptx-updater.ps1`，输入 `changes.json`
- Core: `skills/pptx/packages/pptx/updater.py`（`OP_DISPATCH` 分发表 + 各 `apply_*` 函数）
- Tests: `skills/pptx/packages/pptx/tests/test_replace_text.py`, `tests/test_shapes.py`, `tests/test_slides.py`
- Notes: 幻灯片级操作通过 `prs.slides._sldIdLst` + `prs.part.drop_rel` 实现（python-pptx 无高层 API）；新增操作只需添加 `apply_*` 函数并注册到 `OP_DISPATCH`
