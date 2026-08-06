# Execution Plan: pptx-design-aware

## Summary
- 读取端（`main.py`）输出设计上下文（design/layout/placeholder/run color+font）；更新端（`updater.py`）新增 `clone_slide` 操作；补测试、更新 `SKILL.md` 与路由图。

## Tasks
- [x] T1: 读取端增强 — 设计上下文
  - Files: `skills/pptx/packages/pptx/main.py`
  - Change:
    - 新增 `extract_design(prs)`：`slide_size={width_cm,height_cm}`；`layouts=[{idx,name}]`（遍历 `prs.slide_layouts`）；`theme={major_font,minor_font,colors{dk1,lt1,accent1..accent6}}`（经 `prs.slide_masters[0].part.rels` 取 theme part，解析 `a:themeElements/a:fontScheme` major/minor latin `typeface` 与 `a:clrScheme` 各色 `srgbClr@val` / `sysClr@lastClr`，try/except 降级 null）。
    - `extract_text_frame`：run 增加 `color`（`run.font.color.type` 为 RGB→`#RRGGBB`；THEME→`theme:ACCENT1`；否则 null）与 `font`（`run.font.name`）。
    - `shape_to_dict`：占位符增加 `placeholder={type, idx}`（`shape.placeholder_format`，非占位符不带此键或 null）。
    - `slide_to_dict`：增加 `layout=slide.slide_layout.name`。
    - `main()`：JSON 顶层加 `design`；Markdown 顶部追加 Design 摘要块。
  - Verify: `test_reader_design` 等用例（见 T2）。
  - Depends on: none

- [x] T2: 读取端测试
  - Files: `skills/pptx/packages/pptx/tests/conftest.py`, `skills/pptx/packages/pptx/tests/test_reader.py`
  - Change:
    - conftest 新增 `styled_deck` fixture：用 `prs.slide_layouts[0]`（Title Slide）建页，写 `slide.shapes.title.text`（产生 TITLE 占位符），并对某 run 设显式 RGB 色 + 字体；保留默认模板的 theme 字体/配色。
    - test_reader 新增用例：断言 `data["design"].slide_size`、`design.layouts` 非空且含名字、`design.theme.major_font/minor_font` 非 null、存在 `placeholder.type` 含 TITLE 的形状、slide `layout` 非空、存在 run 的 `color` 形如 `#RRGGBB` 且 `font` 非空；Markdown 含 Design 摘要。
  - Verify: `uv run pytest tests/test_reader.py`
  - Depends on: T1

- [x] T3: 更新端新增 clone_slide
  - Files: `skills/pptx/packages/pptx/updater.py`
  - Change:
    - 新增 `apply_clone_slide(prs, op)`：`source`（1-based int 必填）、`position`（int 可选）；取 `src=get_slide(prs, source)`；`new_slide=prs.slides.add_slide(src.slide_layout)`；清空 `new_slide.shapes` 默认形状（逐个 `shape._element.getparent().remove`）；`deepcopy(src 每个形状的 _element)` 用 `_spTree.insert_element_before(el, 'p:extLst')` 插入；复制 `src.part.rels`（跳过含 `notesSlide` 的关系；`is_external` 用 `get_or_add_ext_rel`，否则 `get_or_add(reltype, target_part)`）；按 `_sldIdLst` 将末尾新页移动到 `position`（同 `apply_insert_slide` 逻辑）。
    - 注册 `"clone_slide": apply_clone_slide` 到 `OP_DISPATCH`。
    - 返回串：`cloned slide {source} -> {position|end} (layout=..., now N slides)`。
  - Verify: T4 用例。
  - Depends on: none

- [x] T4: clone_slide 测试
  - Files: `skills/pptx/packages/pptx/tests/conftest.py`, `skills/pptx/packages/pptx/tests/test_slides.py`
  - Change:
    - conftest（复用现有 `deck_path`，已有文本页/表格页；克隆含图片场景用 `logo_path` 先 `apply_add_picture` 进某页再克隆）。
    - test_slides 新增：①克隆 slide 1 → 页数 +1、新页含「标题 Q3」、`position` 生效；②克隆含图片页 → 重新 `Presentation` 打开保存文件成功且新页存在 PICTURE 形状（验证关系复制）；③`source` 非法/越界抛错；④缺 `source` 抛错。
  - Verify: `uv run pytest tests/test_slides.py`
  - Depends on: T3

- [x] T5: SKILL.md 文档化
  - Files: `skills/pptx/SKILL.md`
  - Change:
    - 读取端「输出格式」：说明 JSON 含顶层 `design`、slide `layout`、形状 `placeholder`、run `color`/`font`；Markdown 含 Design 摘要。
    - 更新端「支持的操作」表加 `clone_slide`；新增 `clone_slide 字段` 小节（`source` 必填、`position` 可选、行为说明：克隆形状+同布局+嵌入关系，不克隆备注）；示例补一条 clone_slide；行为说明补一条。
  - Verify: 文档自洽，字段与代码一致。
  - Depends on: T1, T3

- [x] T6: route-sync
  - Files: `docs/routespec/feature-routes.md`
  - Change:
    - 「读取 PPTX」Notes：新增 design/layout/placeholder/color/font；移除「JSON 不含 run 级 color/font」表述。
    - 「更新 PPTX」Description/Notes：操作清单与 Notes 加入 `clone_slide`。
  - Verify: 条目与实现一致。
  - Depends on: T1, T3

## Verification
- Commands: `uv run pytest`（workdir=`skills/pptx/packages/pptx`）
- Manual checks: 用真实企业 .pptx 跑 `scripts/pptx-reader.ps1` 查看 design 块/placeholder/color；用 `scripts/pptx-updater.ps1` 跑一条 `clone_slide` 指令，PowerPoint 打开验证母版/背景/标题占位符一致。

## Task Relationships
- Strongly related: T1 + T2（同读取端，测试依赖实现）；T3 + T4（同更新端，测试依赖实现）
- Weakly related: T5 依赖 T1/T3 的最终接口形态
- Independent: 读取端（T1/T2）与更新端（T3/T4）文件边界清晰，可并行
- Conflict risks: `tests/conftest.py` 被 T2、T4 共同修改（各自新增不同 fixture）——建议 T2、T4 顺序执行或合并 conftest 改动；T5/T6 文档需待代码接口稳定后落笔

## RouteSync
- Need route-sync: yes（已在 T6 内）
- Expected updates: pptx 模块「读取 PPTX」「更新 PPTX」两条目的 Notes/Description

## Risks
- theme XML 解析跨模板差异 → try/except 降级 null（T1 内置）。
- clone_slide 关系重链接错误致文件损坏 → T4 含「重新打开 + 图片存在」验证。
