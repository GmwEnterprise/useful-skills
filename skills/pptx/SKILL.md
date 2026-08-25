---
name: pptx
description: 基于已有 PowerPoint（.pptx）做增量编辑（非从零创建）：读取提取文本/表格/形状/设计（母版/布局/主题），通过 JSON 指令替换文本（可带样式）、插入图片、删除形状、克隆幻灯片；新增/复制/克隆页面时保留现有模板与设计风格。适合在企业 PPT 上改文字、换图、加页、克隆页等场景。
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
- 需要整框重写标题/文本框内容（原文被拆分到多个 run、`replace_text` 无法命中时）
- 需要在保留原幻灯片样式的前提下修改文字内容
- 需要替换文本并同步改变字号/加粗/斜体/颜色/字体
- 需要向已有幻灯片插入本地图片
- 需要删除幻灯片中的图片、表格或其它形状
- 需要批量裁剪幻灯片，仅保留指定页面（如模板抽取、去附录）
- 需要在保留母版/背景/标题占位符结构的前提下新增一页（克隆现有页）
- 需要在插入图标/图片前确定框内文字的真实渲染位置（渲染导出 + 像素扫描定坐标）
- 需要给企业模板配图标（下载 iconify 图标并着色为透明 PNG）
- 生成或大改后需要视觉核验（改前后 diff + 渲染图复核）

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

**Markdown 文件**：顶部含 `## Design` 设计摘要（幻灯片尺寸、布局列表、主题字体、主题配色、`### Layout shapes` 各布局的形状明细、`### Slide -> layout` 页→布局映射）；其后每页一个 `## Slide N` 章节，列出文本框文本（按段落缩进）、表格（markdown 表格）、备注。

**JSON 文件**：完整结构化数据，适合程序处理。顶层含 `design`（`slide_size` / `layouts` / `slide_layout_map` / `theme` 字体与配色）。`design.layouts` 每项含 `shapes`（该布局全部形状的 id、名称、类型、位置尺寸、填充色、占位符类型——页头横幅、logo、封面装饰等设计元素通常在这一层）；`design.slide_layout_map` 是页码→布局名的映射。每页含 `layout`（布局名）；每个形状含 `id`（shape_id，**页内唯一**）、`name`、位置尺寸字段（`left_cm`/`top_cm`/`width_cm`/`height_cm`）、`fill`（形状填充色：`#RRGGBB` / `theme:ACCENT_1` / `null`，判断深浅底、配图标配色最需要它），为占位符时含 `placeholder`（`type` 如 `TITLE`/`CENTER_TITLE`/`BODY`、`idx`）；表格形状除 `table`（纯字符串矩阵）外含 `table_runs`（行×列×run 文本列表，用于预判表格内 `replace_text` 能否命中）；每个 run 含 `color`（`#RRGGBB` 或 `theme:ACCENT_1` 或 `null`）与 `font`（字体名或 `null`），另含文本/表格/备注/run 级 bold/italic/size。

> **null 样式语义**：run 的 `size`/`color`/`font` 为 `null` 时表示该 run 未直接设置样式，实际显示值继承自占位符→布局→母版的继承链，分析实际显示效果（如"页头标题统一 20pt"）需结合 `design.layouts` 的布局层信息判断，`null` 不等于"无字号/无颜色"。

### 读取示例

```bash
# 读取到源文件同目录
scripts/pptx-reader deck.pptx

# 读取到指定目录
scripts/pptx-reader deck.pptx .tmp/extracted
```

---

## 功能二：更新已有 PPTX

通过 JSON 文件描述修改指令更新已有 `.pptx`：替换文本（可带样式）、整框重写文本、插入本地图片、删除形状。未提供 `style` 时文本替换会保留原格式。

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
| `replace_text` | 段落内局部文本替换（单 run 内查找），可附加 `style` 设置新样式 |
| `set_text` | 形状级整框替换（跨 run 可用），保留首个 run 样式 |
| `add_picture` | 在指定页插入本地图片 |
| `add_textbox` | 在指定页新增文本框（可带样式） |
| `delete_shape` | 删除指定页的形状/图片 |
| `format_shape` | 批量设置形状样式（不改文本） |
| `move_shape` | 调整形状位置与尺寸 |
| `delete_slide` | 删除指定页 |
| `keep_slides` | 仅保留指定页，删除其余所有页 |
| `insert_slide` | 在指定位置插入新页 |
| `move_slide` | 调整页序 |
| `clone_slide` | 克隆（复制）已有页，保留母版/背景/占位符结构 |

### JSON 指令格式

```json
{
  "operations": [
    { "type": "replace_text", "find": "旧公司名", "replace": "新公司名" },
    { "type": "replace_text", "slide": 1, "find": "Q3", "replace": "Q4",
      "style": { "bold": true, "size": 32, "color": "#FF0000" } },
    { "type": "set_text", "slide": 1, "id": 4, "text": "新主标题" },
    { "type": "add_picture", "slide": 2, "path": "logo.png",
      "left_cm": 5, "top_cm": 3, "width_cm": 6 },
    { "type": "delete_shape", "slide": 2, "id": 7 },
    { "type": "delete_shape", "slide": 2, "shape_type": "PICTURE" },
    { "type": "add_textbox", "slide": 3, "text": "新增说明",
      "left_cm": 2, "top_cm": 10, "width_cm": 15, "height_cm": 2,
      "style": { "bold": true, "size": 18 } },
    { "type": "format_shape", "slide": 1, "shape_type": "TEXT_BOX",
      "style": { "color": "#0066CC", "font": "微软雅黑" } },
    { "type": "move_shape", "slide": 2, "id": 9,
      "left_cm": 10, "top_cm": 8, "width_cm": 8 },
    { "type": "insert_slide", "position": 2, "layout": "Blank" },
    { "type": "move_slide", "source": 3, "target": 1 },
    { "type": "clone_slide", "source": 1, "position": 3 },
    { "type": "keep_slides", "slides": [1, 10, 21, 22] },
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
| `id` | operation | 视操作 | 形状 id，取自 `pptx-reader` 输出 JSON 中每个形状的 `id` 字段（页内唯一）；支持的操作中优先级高于 `name`，同名形状歧义时必须用 `id` |

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

### set_text 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slide` | 是 | 目标页码（1-based） |
| `id` | 二选一 | 形状 id（优先于 `name`） |
| `name` | 二选一 | 形状名；同名形状多于一个时报错，改用 `id` |
| `text` | 是 | 新文本；`\n` 表示分段 |
| `style` | 否 | 样式对象，应用到新文本所有 run（字段同 replace_text 的 `style`） |

`set_text` 按形状定位后**整框替换**文本框内容：保留第 1 个 run 的字符样式与第 1 段的段落属性（对齐、缩进等）写入新文本，删除其余 run 与段落。适用口径：**换标题、换副标题、重写整个文本框**用 `set_text`；段落内局部改词用 `replace_text`。目标文本被拆分到多个 run 时 `replace_text` 无法命中，正是本操作要解决的场景。仅作用于文本框（含占位符），不作用于表格。

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
| `id` | 三选一 | 按 id 精确删除**单个**形状（优先于其余两者） |
| `name` | 三选一 | 删除所有同名形状（同名多于一个且只想删其一时用 `id`） |
| `shape_type` | 三选一 | 删除该页所有指定类型形状（如 `PICTURE`） |

常用 `shape_type` 值：`PICTURE`（图片）、`TABLE`（表格）、`TEXT_BOX`（文本框）、`AUTO_SHAPE`（自选图形）、`PLACEHOLDER`（占位符）。形状的 `id` 与 `name` 可在 `pptx-reader` 输出的 JSON 中查看。

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
| `id` | 否 | 按 id 精确定位单个形状（优先于 `name`/`shape_type`） |
| `name` | 否 | 精确匹配形状名（与 `shape_type` 二选一，都不给则作用于整页所有形状） |
| `shape_type` | 否 | 匹配该页所有指定类型形状（与 `name` 二选一，都不给则作用于整页所有形状） |

### move_shape 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slide` | 是 | 目标页码（1-based） |
| `id` | 二选一 | 形状 id（优先于 `name`） |
| `name` | 二选一 | 形状名；同名形状多于一个时报错，改用 `id` |
| `left_cm` | 至少一个 | 新的水平位置(cm) |
| `top_cm` | 至少一个 | 新的垂直位置(cm) |
| `width_cm` | 至少一个 | 新的宽度(cm) |
| `height_cm` | 至少一个 | 新的高度(cm) |

### delete_slide 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slide` | 是 | 要删除的页码（1-based） |

### keep_slides 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `slides` | 是 | 要保留的页码数组（1-based）；未列出的页全部删除 |

模板抽取、去附录等批量裁剪场景用一条 `keep_slides` 完成，避免写大量 `delete_slide` 指令；同时规避多条 `delete_slide` 顺序执行时页码逐条移位、必须倒序书写的问题。保留页的相对顺序不变。

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

### clone_slide 字段

| 字段 | 必填 | 说明 |
|------|------|------|
| `source` | 是 | 要克隆的源页码（1-based） |
| `position` | 否 | 新页插入位置（1-based，插到该位置之前）；不填则追加到末尾 |

`clone_slide` 复制源页的全部形状（文本/图片/表格等）及其所在布局，新页继承同一母版的背景与占位符结构，是企业模板下「新增一页且贴合原设计」最稳妥的方式。默认不复制源页备注（notesSlide）；克隆后可再用 `replace_text` 改写新页内容。嵌入图片/图表等关系会被正确重链接，结果可被 PowerPoint 正常打开。

### 行为说明

- **保留格式**：未提供 `style` 时，`replace_text` 在 run 级别替换并保留原样式。
- **set_text 样式陷阱**：`set_text` 将"首个 run 样式 + 首段段落属性"应用到整框所有新段落。对"标题行加粗大字 + 正文小字"的**多段多样式框**（如痛点/目标类对比框），整框替换会让正文也变成标题样式——此类框**禁用 set_text**，改用逐段 `replace_text`（每段恰好单 run 时保样式完美）。
- **clone_slide 连带资源**：克隆会把源页上的图片一起带过来。当克隆目的是"借版式换内容"时，流程应为：克隆 → `delete_shape shape_type=PICTURE`（先清图）→ `set_text`/`replace_text` 改文案 → 重新渲染扫描 → 加新图。
- **操作顺序法则**：operations 数组按顺序执行，`clone_slide`/`insert_slide`/`move_slide`/`delete_slide`/`keep_slides` 会使后续页码漂移。**结构操作先执行**（可单独跑一轮），内容操作基于结构变更后的最终页码；复杂任务拆成多趟跑 updater（如 clone → 文字 → 图标）最稳。
- **覆盖范围**：`replace_text` 作用于所有文本框、表格单元格、幻灯片备注。
- **跨 run 限制**：目标文本被拆分到多个 run 时（混排样式常见），单 run 内查找可能无法命中；此类场景改用 `set_text` 整框替换。建议先用 `pptx-reader` 确认文本结构。
- **形状定位**：`set_text`/`delete_shape`/`move_shape`/`format_shape` 支持按 `id` 定位（reader JSON 每个形状的 `id` 字段，页内唯一），优先级高于 `name`；同名形状歧义时必须用 `id`。
- **set_text 整框替换**：保留首个 run 样式与首段段落属性重写整个文本框，`\n` 分段；不作用于表格。
- **keep_slides 保留式删页**：按保留集一次裁剪，索引稳定；做"34 页裁到 5 页"类任务时优先于逐条 `delete_slide`。
- **删除不可撤销**：`delete_shape`/`delete_slide`/`keep_slides` 直接修改幻灯片 XML，建议通过 `output` 另存为新文件以保留原件。
- **批量样式**：`format_shape` 只改样式不动文本；`replace_text` 带 `style` 只影响命中 run。两者搭配可覆盖「整形状美化」与「局部替换加粗」。
- **move_shape 单形状**：按 `id` 或唯一 `name` 定位，调整位置与尺寸，四个尺寸字段任意子集、至少一个。
- **move_slide 语义**：将 `source` 页移到 `target` 位置，其余页自动顺延。
- **底层依赖**：`insert_slide`/`move_slide`/`delete_slide`/`keep_slides` 操作 python-pptx 的内部 `_sldIdLst` 结构（官方尚未提供高层 API），兼容 python-pptx 1.0.x。
- **克隆保真**：`clone_slide` 复用源页布局（含母版/背景/占位符），深拷贝形状并重链接嵌入关系（rId 重映射）；适合在保留设计语言的前提下批量生成结构相似的新页。

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

当用户需要修改已有 `.pptx` 时，按五步流程执行：

1. **读取与预检**：先 `scripts/pptx-reader deck.pptx` 读取。`replace_text` 前必须先看 `runs[].text` 做 **run 级预检**——确认目标文本是否被拆分到多个 run、`find` 是否是某个 run 内确实存在的连续片段（如标题拆两个 run、要改的前缀恰在首 run 内才可命中）。表格看 `table_runs`。run 拆散无法命中时换标题/整框用 `set_text`，局部改词调整 `find` 片段。
2. **定坐标（插入元素时）**：渲染导出 PNG 并做像素扫描，确定框内文字真实边界与留白（见「视觉验证闭环」），再计算插入位置。
3. **生成指令**：组织为 JSON 保存到 `.tmp/changes.json`；**结构操作在前**，内容操作使用结构变更后的最终页码。
4. **应用修改**：执行 `scripts/pptx-updater deck.pptx .tmp/changes.json [output.pptx]`。
5. **核验**：对插入的元素做改前后 diff 双条件核验，并目视复核渲染 PNG（见「视觉验证闭环」）；告知用户新文件路径与命中数。

---

## 视觉验证闭环

PPT 是视觉产物：**文字替换命中 ≠ 版面正确**（新文字可能溢出框、换行错位、样式被破坏）；JSON 只有形状框位置，没有文字真实渲染位置——凡"往已有页面插入元素（图标/图片/新形状）"或"生成/大改后核验"，都要走渲染验证。

### 导出页面 PNG

```bash
# Windows + 桌面版 PowerPoint（COM 导出，保真度最高）
scripts/pptx-render.ps1 deck.pptx .tmp/render 1600

# Linux/macOS/WSL2（LibreOffice + poppler：soffice 转 PDF、pdftoppm 逐页出图）
scripts/pptx-render deck.pptx .tmp/render 1600
```

输出 `slide1.png...slideN.png`，并打印 **px/cm 换算比例**（如 1600px / 33.87cm）。所有像素↔cm 换算一律用该比例，不要用屏幕 DPI。

已踩坑（脚本已内置规避，自写 COM 代码时须注意）：

- PowerPoint COM 对**中文路径直接 E_FAIL**：先复制到纯 ASCII 临时路径再 Open（ps1 脚本已自动处理）；
- `WithWindow=0`（不弹窗）在部分环境 E_FAIL，须用 `-1`；
- WPS 不提供等价 COM 导出接口，需桌面版 PowerPoint；无 Office 环境走 Bash 版（注意 `soffice --convert-to png` 对多页 PPT 只导第一页，必须走 PDF 中转）。

### 像素扫描定位框内文字真实边界（插元素前定坐标）

原理：框内出现次数最多的颜色 = 背景填充色；与背景差异超阈值的像素集合 = 文字/图形 bbox。用逃生舱口环境跑 PIL 脚本：

```python
from collections import Counter
# x0,y0,x1,y1 为框内缩进 0.1~0.2cm 后的像素范围（cm * 换算比例）
c = Counter(im.getpixel((x, y)) for y in range(y0, y1) for x in range(x0, x1))
bg = c.most_common(1)[0][0][:3]   # 背景色（顺带拿到框的填充色）
# 遍历求非背景像素的 min/max x/y → 文字 bbox（除以换算比例得 cm）
```

由此可判断文字对齐方式（bbox 居中/贴左）、左侧留白是否够放图标（实测比视觉模型可靠）。实测规律：**痛点/目标类对比框多为左对齐、框内无留白**，图标需放框外左/右侧；流程框多为居中、留白充足。reader JSON 的 `fill` 字段可直接给出框底色，可与扫描结果互验。

### 改前后 diff 自动核验（插元素后验证）

对每个新增元素的矩形区域，**双条件判定**（两者都过才算通过，可批量出结论，比视觉模型逐个看图可靠）：

1. **改前** PNG 该区域 ≥98% 纯色（证明没压到任何文字）；
2. **改后**与改前 diff 有 ≥5% 像素变化（证明元素真的渲染出来了）。

---

## 图标素材（iconify → 着色透明 PNG）

`add_picture` 只收本地图片；"给深蓝框配白色线性图标"这类需求用本管线一步产出：

```bash
# Bash / PowerShell 等价
scripts/pptx-icon .tmp/icons alert target shield-check --color FFFFFF --size 256
scripts/pptx-icon.ps1 .tmp/icons alert target shield-check --color FFFFFF
```

- 素材源为 iconify API（`https://api.iconify.design/<prefix>/<name>.svg`，默认 `mdi` 集，`--prefix` 可换）。**只支持 SVG**（PNG 端点 404）；不存在的图标名会逐一报 HTTP 404，可先 `curl -o /dev/null -s -w '%{http_code}' <url>` 批量校验再执行。
- 管线：SVG → 白底渲染 → 亮度反转作 alpha 通道 → 整体着目标色 → 裁剪 glyph bbox 居中到统一方形画布——抗锯齿边缘完美保留，插入时尺寸可控。
- 产出为透明 PNG，配合 `add_picture` 插入，`width_cm` 控制大小。

**语义配色默认策略**（企业模板实测规律，底色看形状的 `fill` 字段）：

| 底色类型 | 图标配色 | 常见用途 |
|---|---|---|
| 深蓝/深灰/深色底 | 白色 | 主流程框、栏目标题条 |
| 红/深红底 | 白色（框外则红本色） | 痛点/代价框 |
| 橙底 | 白色（框外则橙本色） | 目标/收益框 |
| 浅蓝/浅灰/白底 | 主题蓝（取自 theme accent） | 来源框、说明条 |

多页重复结构（如每页的"痛点/目标"框）用同一对图标（警示/靶心）可形成跨页视觉语言。

---

## 自写脚本（逃生舱口）

任务超出 JSON 指令能力时（如复杂形状遍历、占位符继承链解析、母版层操作），优先**复用技能已初始化的环境**运行自定义脚本，不要用 `uv run --with python-pptx` 等方式另建环境、重复下载依赖：

```bash
# <skill_dir> 为本技能目录；依赖已锁定在 packages/pptx/uv.lock，不触发下载
uv run --project <skill_dir>/packages/pptx python your_script.py
```

自定义脚本仍应遵守：

- 修改类操作输出到新文件，不覆写源文件
- 临时脚本与产物放入工作目录 `.tmp/` 下
- 自定义脚本能力成熟后，建议沉淀回本技能的 JSON 指令

---

## 依赖要求

- Python 3.12+
- `python-pptx`
- `uv`（用于依赖管理）

首次运行会自动初始化虚拟环境并安装依赖。

## 常见问题

### 输出文件被占用

```
Error: [Errno 13] Permission denied: ...
Hint: the file may be open in WPS/PowerPoint. Close it and retry, ...
```
目标文件正被 WPS/PowerPoint 打开。关闭后重试，或另存为新文件名。排查：`tasklist | findstr /i "wpp powerpnt"`，并检查源文件目录下是否有 `~$*.pptx` 锁文件。

### Windows + Git Bash 下 Bash 脚本参数异常

Git Bash 的 MSYS 路径转换会破坏 `D:\...` 形式的参数（实测被拆成 `D;C:\Program Files\Git\...`）。**Windows 下一律使用 `.ps1` 版脚本**。

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

### 同名形状歧义

```
Error: ambiguous shape name '文本框 486': 4 shapes match, locate by 'id' instead
```
该页存在多个同名形状。从 `pptx-reader` 输出的 JSON 中取目标形状的 `id` 字段，指令改用 `"id": <数值>` 定位。

### 替换命中数为 0

若报告 `0 hit(s)`，通常是目标文本被拆分到多个 run，或文本与 `find` 不完全一致（含空格/大小写）。用 `pptx-reader` 输出的 JSON 查看 run 级别文本：段落内局部改词时调整 `find` 为 run 内确实存在的连续片段；换标题/整框重写时改用 `set_text`。
