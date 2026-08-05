#!/usr/bin/env python3
import json
import sys
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Cm, Pt


def iter_text_frames(slide: Any):
    for shape in slide.shapes:
        if shape.has_text_frame:
            yield shape.text_frame
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    yield cell.text_frame
    if slide.has_notes_slide:
        yield slide.notes_slide.notes_text_frame


def replace_in_text_frame(
    tf: Any, find: str, replace: str, style: dict | None = None
) -> int:
    count = 0
    for para in tf.paragraphs:
        for run in para.runs:
            if find and find in run.text:
                count += run.text.count(find)
                run.text = run.text.replace(find, replace)
                if style:
                    apply_style_to_run(run, style)
    return count


def normalize_hex(color: str) -> str:
    c = color.lstrip("#")
    if len(c) == 3:
        c = "".join(ch * 2 for ch in c)
    if len(c) != 6:
        raise ValueError(f"invalid color: {color!r} (expect #RRGGBB)")
    return c.upper()


def apply_style_to_run(run: Any, style: dict) -> None:
    if "bold" in style:
        run.font.bold = bool(style["bold"])
    if "italic" in style:
        run.font.italic = bool(style["italic"])
    if "size" in style:
        run.font.size = Pt(style["size"])
    if "font" in style:
        run.font.name = style["font"]
    if "color" in style:
        run.font.color.rgb = RGBColor.from_string(normalize_hex(style["color"]))


def get_slide(prs: Any, slide_index: int) -> Any:
    slides = list(prs.slides)
    if slide_index < 1 or slide_index > len(slides):
        raise ValueError(f"slide {slide_index} out of range (1..{len(slides)})")
    return slides[slide_index - 1]


def shape_type_name(shape: Any) -> str | None:
    try:
        return shape.shape_type.name
    except Exception:
        return None


def apply_replace_text(prs: Any, op: dict) -> str:
    find = op.get("find")
    replace = op.get("replace")
    if not isinstance(find, str) or not isinstance(replace, str):
        raise ValueError("replace_text requires string 'find' and 'replace'")
    slide = op.get("slide")
    if slide is not None and not isinstance(slide, int):
        raise ValueError("'slide' must be an integer (1-based)")
    style = op.get("style")
    if style is not None and not isinstance(style, dict):
        raise ValueError("'style' must be an object")

    total = 0
    for idx, slide_obj in enumerate(prs.slides, start=1):
        if slide is not None and idx != slide:
            continue
        for tf in iter_text_frames(slide_obj):
            total += replace_in_text_frame(tf, find, replace, style)
    scope = f"slide {slide}" if slide is not None else "all slides"
    style_note = " +style" if style else ""
    return f"{find!r} -> {replace!r}{style_note} ({scope}): {total} hit(s)"


def apply_add_picture(prs: Any, op: dict) -> str:
    slide_index = op.get("slide")
    if not isinstance(slide_index, int):
        raise ValueError("add_picture requires integer 'slide' (1-based)")
    img_path = op.get("path")
    if not isinstance(img_path, str):
        raise ValueError("add_picture requires string 'path'")
    left_cm = op.get("left_cm")
    top_cm = op.get("top_cm")
    if not isinstance(left_cm, (int, float)) or not isinstance(
        top_cm, (int, float)
    ):
        raise ValueError("add_picture requires numeric 'left_cm' and 'top_cm'")

    img = Path(img_path)
    if not img.exists():
        raise ValueError(f"image not found: {img_path}")

    slide = get_slide(prs, slide_index)

    kwargs: dict = {}
    width_cm = op.get("width_cm")
    height_cm = op.get("height_cm")
    if width_cm is not None:
        kwargs["width"] = Cm(width_cm)
    if height_cm is not None:
        kwargs["height"] = Cm(height_cm)

    slide.shapes.add_picture(str(img), Cm(left_cm), Cm(top_cm), **kwargs)

    if width_cm is not None and height_cm is not None:
        size_note = f" {width_cm}x{height_cm}cm"
    elif width_cm is not None:
        size_note = f" w={width_cm}cm"
    elif height_cm is not None:
        size_note = f" h={height_cm}cm"
    else:
        size_note = " (original size)"
    return f"added picture at ({left_cm},{top_cm}){size_note}"


def apply_delete_shape(prs: Any, op: dict) -> str:
    slide_index = op.get("slide")
    if not isinstance(slide_index, int):
        raise ValueError("delete_shape requires integer 'slide' (1-based)")
    name = op.get("name")
    shape_type = op.get("shape_type")
    if name is None and shape_type is None:
        raise ValueError("delete_shape requires 'name' or 'shape_type'")

    slide = get_slide(prs, slide_index)

    to_remove = []
    for shape in slide.shapes:
        if name is not None and shape.name == name:
            to_remove.append(shape)
        elif shape_type is not None and shape_type_name(shape) == shape_type:
            to_remove.append(shape)

    for shape in to_remove:
        el = shape._element
        el.getparent().remove(el)

    if name is not None:
        return f"removed shape named {name!r}: {len(to_remove)}"
    return f"removed {len(to_remove)} {shape_type} shape(s)"


def find_shapes(
    slide: Any, name: str | None = None, shape_type: str | None = None
) -> list[Any]:
    result = []
    for shape in slide.shapes:
        if name is not None:
            if shape.name == name:
                result.append(shape)
        elif shape_type is not None:
            if shape_type_name(shape) == shape_type:
                result.append(shape)
        else:
            result.append(shape)
    return result


def apply_style_to_shape(shape: Any, style: dict) -> int:
    count = 0
    frames = []
    if shape.has_text_frame:
        frames.append(shape.text_frame)
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                frames.append(cell.text_frame)
    for tf in frames:
        for para in tf.paragraphs:
            for run in para.runs:
                apply_style_to_run(run, style)
                count += 1
    return count


def apply_format_shape(prs: Any, op: dict) -> str:
    slide_index = op.get("slide")
    if not isinstance(slide_index, int):
        raise ValueError("format_shape requires integer 'slide' (1-based)")
    style = op.get("style")
    if not isinstance(style, dict):
        raise ValueError("format_shape requires a 'style' object")
    name = op.get("name")
    shape_type = op.get("shape_type")

    slide = get_slide(prs, slide_index)
    shapes = find_shapes(slide, name, shape_type)
    runs = 0
    for shape in shapes:
        runs += apply_style_to_shape(shape, style)

    if name is not None:
        target = f"name={name!r}"
    elif shape_type is not None:
        target = f"type={shape_type!r}"
    else:
        target = "all shapes"
    return (
        f"styled {target} on slide {slide_index}: "
        f"{len(shapes)} shape(s), {runs} run(s)"
    )


def apply_move_shape(prs: Any, op: dict) -> str:
    slide_index = op.get("slide")
    if not isinstance(slide_index, int):
        raise ValueError("move_shape requires integer 'slide' (1-based)")
    name = op.get("name")
    if not isinstance(name, str):
        raise ValueError("move_shape requires string 'name'")
    fields = {
        "left_cm": op.get("left_cm"),
        "top_cm": op.get("top_cm"),
        "width_cm": op.get("width_cm"),
        "height_cm": op.get("height_cm"),
    }
    if all(v is None for v in fields.values()):
        raise ValueError(
            "move_shape requires at least one of "
            "left_cm/top_cm/width_cm/height_cm"
        )
    for k, v in fields.items():
        if v is not None and not isinstance(v, (int, float)):
            raise ValueError(f"'{k}' must be numeric")

    slide = get_slide(prs, slide_index)
    shapes = find_shapes(slide, name=name)
    if not shapes:
        raise ValueError(f"shape named {name!r} not found on slide {slide_index}")
    shape = shapes[0]

    parts = []
    if fields["left_cm"] is not None:
        shape.left = Cm(fields["left_cm"])
        parts.append(f"left={fields['left_cm']}cm")
    if fields["top_cm"] is not None:
        shape.top = Cm(fields["top_cm"])
        parts.append(f"top={fields['top_cm']}cm")
    if fields["width_cm"] is not None:
        shape.width = Cm(fields["width_cm"])
        parts.append(f"width={fields['width_cm']}cm")
    if fields["height_cm"] is not None:
        shape.height = Cm(fields["height_cm"])
        parts.append(f"height={fields['height_cm']}cm")
    return f"moved {name!r} on slide {slide_index}: {', '.join(parts)}"


def apply_add_textbox(prs: Any, op: dict) -> str:
    slide_index = op.get("slide")
    if not isinstance(slide_index, int):
        raise ValueError("add_textbox requires integer 'slide' (1-based)")
    text = op.get("text")
    if not isinstance(text, str):
        raise ValueError("add_textbox requires string 'text'")
    left_cm = op.get("left_cm")
    top_cm = op.get("top_cm")
    width_cm = op.get("width_cm")
    height_cm = op.get("height_cm")
    if not all(
        isinstance(v, (int, float))
        for v in (left_cm, top_cm, width_cm, height_cm)
    ):
        raise ValueError(
            "add_textbox requires numeric left_cm/top_cm/width_cm/height_cm"
        )

    slide = get_slide(prs, slide_index)
    shape = slide.shapes.add_textbox(
        Cm(left_cm), Cm(top_cm), Cm(width_cm), Cm(height_cm)
    )
    shape.text_frame.text = text
    style = op.get("style")
    runs = 0
    if isinstance(style, dict):
        runs = apply_style_to_shape(shape, style)

    style_note = f" +style({runs} runs)" if isinstance(style, dict) else ""
    return (
        f"added textbox on slide {slide_index}: {text!r}{style_note} "
        f"at ({left_cm},{top_cm}) {width_cm}x{height_cm}cm"
    )


def _sld_id_list(prs: Any):
    return prs.slides._sldIdLst


def resolve_layout(prs: Any, layout: Any) -> Any:
    layouts = prs.slide_layouts
    if layout is None:
        for lay in layouts:
            if "blank" in (lay.name or "").lower():
                return lay
        return layouts[min(6, len(layouts) - 1)]
    if isinstance(layout, bool):
        raise ValueError("layout must be int index or str name")
    if isinstance(layout, int):
        if layout < 0 or layout >= len(layouts):
            raise ValueError(
                f"layout index {layout} out of range (0..{len(layouts) - 1})"
            )
        return layouts[layout]
    if isinstance(layout, str):
        for lay in layouts:
            if lay.name == layout:
                return lay
        for lay in layouts:
            if layout.lower() in (lay.name or "").lower():
                return lay
        raise ValueError(f"layout {layout!r} not found")
    raise ValueError("layout must be int index or str name")


def apply_delete_slide(prs: Any, op: dict) -> str:
    slide_index = op.get("slide")
    if not isinstance(slide_index, int):
        raise ValueError("delete_slide requires integer 'slide' (1-based)")
    ids = _sld_id_list(prs)
    n = len(ids)
    if slide_index < 1 or slide_index > n:
        raise ValueError(f"slide {slide_index} out of range (1..{n})")
    el = ids[slide_index - 1]
    rId = el.rId
    prs.part.drop_rel(rId)
    ids.remove(el)
    return f"removed slide {slide_index} (now {n - 1} slides)"


def apply_insert_slide(prs: Any, op: dict) -> str:
    position = op.get("position")
    layout = resolve_layout(prs, op.get("layout"))
    ids = _sld_id_list(prs)
    n = len(ids)
    prs.slides.add_slide(layout)
    new_el = ids[n]
    ids.remove(new_el)
    if position is None:
        target = n
    else:
        if not isinstance(position, int):
            raise ValueError("'position' must be int (1-based)")
        if position < 1 or position > n + 1:
            raise ValueError(f"position {position} out of range (1..{n + 1})")
        target = position - 1
    ids.insert(target, new_el)
    pos_note = "end" if position is None else f"position {position}"
    lay_name = getattr(layout, "name", None)
    return (
        f"inserted new slide at {pos_note} "
        f"(layout={lay_name!r}, now {n + 1} slides)"
    )


def apply_move_slide(prs: Any, op: dict) -> str:
    source = op.get("source")
    target = op.get("target")
    if not isinstance(source, int) or not isinstance(target, int):
        raise ValueError(
            "move_slide requires integer 'source' and 'target' (1-based)"
        )
    ids = _sld_id_list(prs)
    n = len(ids)
    if source < 1 or source > n:
        raise ValueError(f"source {source} out of range (1..{n})")
    if target < 1 or target > n:
        raise ValueError(f"target {target} out of range (1..{n})")
    el = ids[source - 1]
    ids.remove(el)
    ids.insert(target - 1, el)
    return f"moved slide {source} -> position {target}"


OP_DISPATCH = {
    "replace_text": apply_replace_text,
    "add_picture": apply_add_picture,
    "add_textbox": apply_add_textbox,
    "delete_shape": apply_delete_shape,
    "format_shape": apply_format_shape,
    "move_shape": apply_move_shape,
    "delete_slide": apply_delete_slide,
    "insert_slide": apply_insert_slide,
    "move_slide": apply_move_slide,
}


def load_changes(path: Path, base_dir: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or "operations" not in data:
        raise ValueError("JSON must contain an 'operations' array")
    ops = data["operations"]
    if not isinstance(ops, list) or not ops:
        raise ValueError("'operations' must be a non-empty array")
    for op in ops:
        if (
            isinstance(op, dict)
            and op.get("type") == "add_picture"
            and isinstance(op.get("path"), str)
        ):
            pp = Path(op["path"])
            if not pp.is_absolute():
                op["path"] = str((base_dir / pp).resolve())
    return ops


def main() -> None:
    if len(sys.argv) < 3:
        print(
            "Usage: python updater.py <pptx_file> <changes.json> [output.pptx]",
            file=sys.stderr,
        )
        print("\nArguments:", file=sys.stderr)
        print(
            "  pptx_file    - Path to the source PowerPoint file (.pptx)",
            file=sys.stderr,
        )
        print(
            "  changes.json - JSON file describing update operations",
            file=sys.stderr,
        )
        print(
            "  output.pptx  - Optional: output path "
            "(default: <source>_updated.pptx)",
            file=sys.stderr,
        )
        sys.exit(1)

    file_path = sys.argv[1]
    changes_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else None

    try:
        path = Path(file_path)
        if not path.exists():
            print(f"Error: PowerPoint file not found: {file_path}", file=sys.stderr)
            sys.exit(1)
        if path.suffix.lower() != ".pptx":
            print(
                f"Error: Unsupported file format: {path.suffix}. "
                "Only .pptx supported.",
                file=sys.stderr,
            )
            sys.exit(1)

        changes_file = Path(changes_path)
        if not changes_file.exists():
            print(
                f"Error: Changes file not found: {changes_path}", file=sys.stderr
            )
            sys.exit(1)

        ops = load_changes(changes_file, changes_file.resolve().parent)

        prs = Presentation(str(path))

        report = []
        for i, op in enumerate(ops, start=1):
            if not isinstance(op, dict):
                raise ValueError(f"Each operation must be an object: {op!r}")
            op_type = op.get("type")
            handler = OP_DISPATCH.get(op_type)
            if handler is None:
                raise ValueError(f"Unsupported operation type: {op_type!r}")
            detail = handler(prs, op)
            report.append(f"  [{i}] {op_type}: {detail}")

        if output_path:
            out_file = Path(output_path)
        else:
            out_file = path.parent / f"{path.stem}_updated.pptx"
        out_file.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(out_file))

        print(f"Success: {path.name}")
        print(f"Operations: {len(ops)}")
        for line in report:
            print(line)
        print(f"Output: {out_file}")

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
