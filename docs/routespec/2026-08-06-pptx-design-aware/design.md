# Design: pptx-design-aware

## Goal
- 让 `pptx` 技能在读取时暴露 PPT 的设计上下文（母版布局、占位符角色、主题字体/配色、run 级颜色/字体），在更新时能新增与原设计完全一致的页面（`clone_slide`）。

## Scope
- In（读取端）：
  - JSON 顶层新增 `design`：`slide_size`、`layouts`（idx+name）、`theme`（major_font/minor_font/colors）。
  - 每个 slide 增加 `layout`（布局名）。
  - 每个形状增加 `placeholder`（`{type, idx}`，非占位符为 null）。
  - 每个 run 增加 `color`、`font`（取消路由图「JSON 不含 run 级 color/font」限制）。
  - Markdown 顶部增加 Design 摘要块。
- In（更新端）：
  - 新增 `clone_slide` 操作：克隆指定页（复制 spTree 形状 + 同布局 + 重链接嵌入关系），可选 `position`。
  - 默认不克隆备注（notesSlide），保持新页干净。
- Out：
  - 不改 `insert_slide` 默认布局（仍 Blank，保持向后兼容）。
  - 不新增 `set_placeholder` 或改 `add_textbox` 行为。
  - 不做「现有页新增占位符内容」的结构化继承（A 选项不含此项）。

## Assumptions
- 主题字体/配色从 `slide_masters[0]` 的 theme 关系 part 解析；不同模板 XML 结构可能有差异，统一用 try/except 降级为 null。
- `clone_slide` 复用源页 `slide_layout`，确保母版/背景绑定一致。

## Current Behavior
- `main.py`：仅输出 text/位置/run 的 bold/italic/size；无 layout、无占位符角色、无主题信息、无 run color/font。
- `updater.py`：`insert_slide`（`updater.py:310` `resolve_layout`）默认 Blank 布局，丢母版背景与标题占位符；`add_textbox`（`updater.py:270`）为自由文本框，不绑定母版；无克隆能力。

## Proposed Behavior
- 读取输出含完整设计上下文，AI 可据此分辨标题/正文、企业字体/配色。
- `clone_slide`：以源页为模板新增一页，1:1 复制形状文本与布局绑定（含母版背景与占位符结构），嵌入图片/图表等关系正确重链接；再配 `replace_text` 即可生成贴合原设计的新页。

## Implementation Direction
- 读取端（`main.py`）：
  - 新增 `extract_design(prs)`：slide 尺寸、`prs.slide_layouts` 列表、theme part 解析（`a:themeElements/a:fontScheme` major/minor latin typeface；`a:clrScheme` dk1/lt1/accent1..6 的 srgbClr/sysClr lastClr）。
  - `shape_to_dict`：占位符 → `placeholder={type, idx}`（`shape.placeholder_format`）。
  - `extract_text_frame`：run 增加 `color`（RGB→`#RRGGBB`；theme→`theme:ACCENT1`；无→null）与 `font`（name）。
  - `slide_to_dict`：增加 `layout`（`slide.slide_layout.name`）。
  - 输出顶层 `design`；Markdown 顶部追加 Design 摘要。
- 更新端（`updater.py`）：
  - 新增 `apply_clone_slide(prs, op)`：`source`（1-based）必填、`position` 可选；`prs.slides.add_slide(src.slide_layout)` 建新页 → 清空新页默认形状 → `deepcopy` 源页各 `shape._element` 插入 spTree → 复制源页 `part.rels`（跳过 notesSlide，外部/内部关系分别用 `get_or_add`）→ 经 `_sldIdLst` 调整 `position`。
  - 注册到 `OP_DISPATCH`。

## Affected Files
- `skills/pptx/packages/pptx/main.py` (route-map)：读取端增强
- `skills/pptx/packages/pptx/updater.py` (route-map)：新增 clone_slide
- `skills/pptx/SKILL.md` (route-map)：文档化新字段与 clone_slide
- `skills/pptx/packages/pptx/tests/test_reader.py` (route-map)：新增读取端断言
- `skills/pptx/packages/pptx/tests/test_slides.py` (route-map)：新增 clone_slide 测试
- `skills/pptx/packages/pptx/tests/conftest.py` (route-map)：可能补一个带标题占位符/带图片的测试 deck fixture
- `docs/routespec/feature-routes.md` (route-map)：route-sync（更新 pptx 读取/更新条目 Notes）

## Risks
- 主题 XML 路径差异：用 try/except 降级为 null，保证健壮。
- clone_slide 关系重链接：嵌入图片/图表若未正确复制 rId 将导致文件损坏；对策——复制全部非 notesSlide 关系，并以「克隆带图片页」用例验证。
- run color：theme 颜色返回 `MSO_THEME_COLOR` 枚举，需统一序列化为 `theme:<NAME>`。

## Test Strategy
- 读取端：构造带标题占位符（layout 0）+ 带显式 RGB 色 run + 主题色 run 的 deck，断言 `design.layouts`、`design.theme`、slide `layout`、形状 `placeholder`、run `color`/`font`。
- 更新端：克隆含文本页 → 断言页数 +1、新页文本与源一致、`position` 正确；克隆含图片页 → 重新打开文件成功且图片存在（验证关系复制）。
- 回归：现有 `uv run pytest` 全绿（字段为向后兼容增量）。

## RouteSpec Impact
- Need route-sync: yes
- Affected routes: pptx 模块「读取 PPTX」「更新 PPTX」条目 Notes（读取新增 design/layout/placeholder/color/font；更新新增 clone_slide；移除「JSON 不含 run 级 color/font」表述）

## Acceptance Criteria
- Reader JSON 含 `design`（slide_size/layouts/theme）、每页 `layout`、占位符形状含 `placeholder`、run 含 `color`/`font`；Markdown 含 Design 摘要。
- `clone_slide` 生成的新页与源页形状文本一致、布局绑定相同、嵌入关系完整；`uv run pytest`（含新用例）全绿。
- 现有测试无回归。
- SKILL.md 文档化读取新字段与 `clone_slide`（操作表/字段表/行为说明/示例）。
