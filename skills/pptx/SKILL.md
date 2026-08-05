---
name: pptx
description: 专门用于读取已有 PowerPoint 文件（.pptx），并基于 JSON 指令更新文本（可带样式）、插入图片、删除形状
---

# PPTX Reader & Updater

读取 `.pptx` 文件并输出 Markdown 与 JSON 两种格式；通过 JSON 描述的修改指令更新已有 `.pptx`：替换文本（可带样式）、插入本地图片、删除形状。

## 前置要求

**重要**：本技能依赖 `uv` 进行 Python 依赖管理。如果用户系统没有安装 `uv`，请停止任务执行并告知用户：

> 当前系统未安装 `uv`，无法执行此技能。请先安装 `uv`：
> ```bash
> curl -LsSf https://astral.sh/uv/install.sh | sh
> ```

## 何时使用

- 需要读取 `.pptx` 提取文本、表格供 AI 分析
- 需要查看每页幻灯片的形状结构、备注
- 需要批量替换已有 `.pptx` 中的文本（如改标题、改公司名、改日期）
- 需要在保留原幻灯片样式的前提下修改文字内容
- 需要替换文本并同步改变字号/加粗/斜体/颜色/字体
- 需要向已有幻灯片插入本地图片
- 需要删除幻灯片中的图片、表格或其它形状

---

## 功能一：读取 PPTX

### 快速开始

```bash
# Bash/Linux/macOS
scripts/pptx-reader <pptx_file> [output_dir]

# PowerShell/Windows
scripts/pptx-reader.ps1 <pptx_file> [output_dir]
```

### 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `pptx_file` | 是 | PowerPoint 文件路径（.pptx） |
| `output_dir` | 否 | 输出目录，默认源文件所在目录 |

### 输出格式

执行成功后，会在输出目录生成两个文件：

```
Success: <文件名>
Slides: <页数>
Markdown: <dir>/<filename>.pptx_reader.md
JSON:     <dir>/<filename>.pptx_reader.json
```

**Markdown 文件**：每页一个 `## Slide N` 章节，列出文本框文本（按段落缩进）、表格（markdown 表格）、备注。

**JSON 文件**：完整结构化数据，包含每页的形状（名称/类型/位置/文本/段落/run 级别的样式）、表格、备注，适合程序处理。

### 读取示例

```bash
# 读取到源文件同目录
scripts/pptx-reader deck.pptx

# 读取到指定目录
scripts/pptx-reader deck.pptx .tmp/extracted
```

---

## 功能二：更新已有 PPTX

通过 JSON 文件描述修改指令更新已有 `.pptx`：替换文本（可带样式）、插入本地图片、删除形状。未提供 `style` 时文本替换会保留原格式。

### 工作流程

1. （可选）先用 `pptx-reader` 读取，确认要替换的原文
2. AI 模型生成 JSON 修改指令文件
3. 调用脚本应用修改，输出新的 `.pptx`

### 快速开始

```bash
# Bash/Linux/macOS
scripts/pptx-updater <pptx_file> <changes.json> [output.pptx]

# PowerShell/Windows
scripts/pptx-updater.ps1 <pptx_file> <changes.json> [output.pptx]
```

### 参数说明

| 参数 | 必填 | 说明 |
|------|------|------|
| `pptx_file` | 是 | 源 PowerPoint 文件路径（.pptx） |
| `changes.json` | 是 | JSON 修改指令文件 |
| `output.pptx` | 否 | 输出路径，默认 `<源文件名>_updated.pptx` |

### 支持的操作

| 操作类型 | 说明 |
|---------|------|
| `replace_text` | 文本替换，可附加 `style` 设置新样式 |
| `add_picture` | 在指定页插入本地图片 |
| `add_textbox` | 在指定页新增文本框（可带样式） |
| `delete_shape` | 删除指定页的形状/图片 |
| `format_shape` | 批量设置形状样式（不改文本） |
| `move_shape` | 调整形状位置与尺寸 |
| `delete_slide` | 删除指定页 |
| `insert_slide` | 在指定位置插入新页 |
| `move_slide` | 调整页序 |

### JSON 指令格式

```json
{
  "operations": [
    { "type": "replace_text", "find": "旧公司名", "replace": "新公司名" },
    { "type": "replace_text", "slide": 1, "find": "Q3", "replace": "Q4",
      "style": { "bold": true, "size": 32, "color": "#FF0000" } },
    { "type": "add_picture", "slide": 2, "path": "logo.png",
      "left_cm": 5, "top_cm": 3, "width_cm": 6 },
    { "type": "delete_shape", "slide": 2, "name": "Picture 1" },
    { "type": "delete_shape", "slide": 2, "shape_type": "PICTURE" },
    { "type": "add_textbox", "slide": 3, "text": "新增说明",
      "left_cm": 2, "top_cm": 10, "width_cm": 15, "height_cm": 2,
      "style": { "bold": true, "size": 18 } },
    { "type": "format_shape", "slide": 1, "shape_type": "TEXT_BOX",
      "style": { "color": "#0066CC", "font": "微软雅黑" } },
    { "type": "move_shape", "slide": 2, "name": "Picture 3",
      "left_cm": 10, "top_cm": 8, "width_cm": 8 },
    { "type": "insert_slide", "position": 2, "layout": "Blank" },
    { "type": "move_slide", "source": 3, "target": 1 },
    { "type": "delete_slide", "slide": 5 }
  ]
}
```

### 通用字段

| 字段 | 层级 | 必填 | 说明 |
|------|------|------|------|
| `operations` | 顶层 | 是 | 修改操作数组，至少包含一个操作 |
| `type` | operation | 是 | 操作类型（见上方「支持的操作」表） |
| `slide` | operation | 视操作 | 页码（1-based）；`replace_text` 可选（默认全文档），其余操作必填 |

### replace_text 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `find` | 是 | 要查找的文本 |
| `replace` | 是 | 替换为的文本 |
| `slide` | 否 | 限定页码；不指定则对所有页（含表格、备注）生效 |
| `style` | 否 | 样式对象，应用到被替换文本所在的 run |

`style` 对象字段（均可选）：

| 字段 | 类型 | 说明 |
|------|------|------|
| `bold` | bool | 加粗 |
| `italic` | bool | 斜体 |
| `size` | number | 字号(pt) |
| `font` | string | 字体名（如 `微软雅黑`） |
| `color` | string | 颜色（`#RRGGBB` / `RRGGBB` / `#RGB`） |

> 样式作用于整个 run（PowerPoint 的最小样式单元），而非仅替换的文本片段。run 是 PowerPoint 中携带样式的最小单位，无法在 run 内部对片段单独设置样式。

### add_picture 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slide` | 是 | 目标页码（1-based） |
| `path` | 是 | 图片路径；相对路径基于 `changes.json` 所在目录 |
| `left_cm` | 是 | 左上角水平位置(cm) |
| `top_cm` | 是 | 左上角垂直位置(cm) |
| `width_cm` | 否 | 宽度(cm)；只给一个维度则等比缩放 |
| `height_cm` | 否 | 高度(cm)；都不给则用原图尺寸 |

### delete_shape 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slide` | 是 | 目标页码（1-based） |
| `name` | 二选一 | 精确匹配形状名（如 `Picture 1`） |
| `shape_type` | 二选一 | 删除该页所有指定类型形状（如 `PICTURE`） |

常用 `shape_type` 值：`PICTURE`（图片）、`TABLE`（表格）、`TEXT_BOX`（文本框）、`AUTO_SHAPE`（自选图形）、`PLACEHOLDER`（占位符）。形状名可在 `pptx-reader` 输出的 JSON 中查看（`name` 字段）。

### add_textbox 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slide` | 是 | 目标页码（1-based） |
| `text` | 是 | 文本框内容（`\n` 表示分段） |
| `left_cm` | 是 | 左上角水平位置(cm) |
| `top_cm` | 是 | 左上角垂直位置(cm) |
| `width_cm` | 是 | 文本框宽度(cm) |
| `height_cm` | 是 | 文本框高度(cm) |
| `style` | 否 | 样式对象，应用到文本框所有 run（字段同 replace_text 的 `style`） |

### format_shape 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slide` | 是 | 目标页码（1-based） |
| `style` | 是 | 样式对象，应用到匹配形状的所有 run（字段同 replace_text 的 `style`） |
| `name` | 否 | 精确匹配形状名（与 `shape_type` 二选一，都不给则作用于整页所有形状） |
| `shape_type` | 否 | 匹配该页所有指定类型形状（与 `name` 二选一，都不给则作用于整页所有形状） |

### move_shape 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slide` | 是 | 目标页码（1-based） |
| `name` | 是 | 精确匹配形状名（操作单个形状） |
| `left_cm` | 至少一个 | 新的水平位置(cm) |
| `top_cm` | 至少一个 | 新的垂直位置(cm) |
| `width_cm` | 至少一个 | 新的宽度(cm) |
| `height_cm` | 至少一个 | 新的高度(cm) |

### delete_slide 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slide` | 是 | 要删除的页码（1-based） |

### insert_slide 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `position` | 否 | 插入位置（1-based，插到该位置之前）；不填则追加到末尾 |
| `layout` | 否 | 布局：整数索引或名称（默认自动选择 `Blank` 布局） |

### move_slide 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `source` | 是 | 源页码（1-based） |
| `target` | 是 | 目标页码（1-based） |

### 行为说明

- **保留格式**：未提供 `style` 时，`replace_text` 在 run 级别替换并保留原样式。
- **覆盖范围**：`replace_text` 作用于所有文本框、表格单元格、幻灯片备注。
- **跨 run 限制**：目标文本被拆分到多个 run 时（混排样式常见），单 run 内查找可能无法命中。建议先用 `pptx-reader` 确认文本结构。
- **删除不可撤销**：`delete_shape`/`delete_slide` 直接修改幻灯片 XML，建议通过 `output` 另存为新文件以保留原件。
- **批量样式**：`format_shape` 只改样式不动文本；`replace_text` 带 `style` 只影响命中 run。两者搭配可覆盖「整形状美化」与「局部替换加粗」。
- **move_shape 单形状**：按 `name` 精确定位，调整位置与尺寸，四个尺寸字段任意子集、至少一个。
- **move_slide 语义**：将 `source` 页移到 `target` 位置，其余页自动顺延。
- **底层依赖**：`insert_slide`/`move_slide`/`delete_slide` 操作 python-pptx 的内部 `_sldIdLst` 结构（官方尚未提供高层 API），兼容 python-pptx 1.0.x。

### 更新示例

**修改指令文件** `.tmp/changes.json`：
```json
{
  "operations": [
    { "type": "replace_text", "find": "ACME 公司", "replace": "示例科技有限公司" },
    { "type": "replace_text", "slide": 1, "find": "Q3", "replace": "Q4",
      "style": { "bold": true, "size": 28 } },
    { "type": "add_picture", "slide": 2, "path": "logo.png",
      "left_cm": 1, "top_cm": 5, "width_cm": 4 }
  ]
}
```

**执行**：
```bash
scripts/pptx-updater deck.pptx .tmp/changes.json deck_updated.pptx
# 输出:
# Success: deck.pptx
# Operations: 3
#   [1] replace_text: 'ACME 公司' -> '示例科技有限公司' (all slides): 5 hit(s)
#   [2] replace_text: 'Q3' -> 'Q4' +style (slide 1): 1 hit(s)
#   [3] add_picture: added picture at (1,5) w=4cm
# Output: /path/to/deck_updated.pptx
```

### AI 使用模式

当用户需要修改已有 `.pptx` 文本时：

1. **读取确认**：先 `scripts/pptx-reader deck.pptx` 读取，定位需要替换的确切原文
2. **生成指令**：将替换项组织为上述 JSON，保存为临时文件（如 `.tmp/changes.json`）
3. **应用修改**：执行 `scripts/pptx-updater deck.pptx .tmp/changes.json [output.pptx]`
4. **告知用户**：输出新文件路径与命中次数

---

## 依赖要求

- Python 3.12+
- `python-pptx`
- `uv`（用于依赖管理）

首次运行会自动初始化虚拟环境并安装依赖。

## 常见问题

### 文件不存在

```
Error: PowerPoint file not found: /path/to/file.pptx
```
检查文件路径是否正确。

### 不支持的格式

```
Error: Unsupported file format: .ppt. Only .pptx is supported.
```
仅支持 `.pptx`，不支持旧版 `.ppt`。请先用 PowerPoint 另存为 `.pptx`。

### JSON 指令缺失 operations

```
Error: JSON must contain an 'operations' array
```
修改指令文件必须包含 `operations` 数组且至少一个操作。

### 替换命中数为 0

若报告 `0 hit(s)`，通常是目标文本被拆分到多个 run，或文本与 `find` 不完全一致（含空格/大小写）。用 `pptx-reader` 输出的 JSON 查看 run 级别文本，调整 `find` 为 run 内确实存在的连续片段。
