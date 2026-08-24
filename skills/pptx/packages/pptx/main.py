#!/usr/bin/env python3
import json
import sys
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.oxml.ns import qn


def emu_to_cm(emu: int | None) -> float | None:
    if emu is None:
        return None
    return round(emu / 360000, 2)


def shape_kind_name(shape: Any) -> str:
    try:
        return str(shape.shape_type)
    except Exception:
        return "UNKNOWN"


def run_color(run: Any) -> str | None:
    color = None
    try:
        ct = run.font.color.type
        if ct == MSO_COLOR_TYPE.RGB:
            color = "#" + str(run.font.color.rgb)
        elif ct == MSO_COLOR_TYPE.THEME:
            color = "theme:" + run.font.color.theme_color.name
    except Exception:
        color = None
    return color


def extract_text_frame(tf: Any) -> list[dict]:
    paragraphs = []
    for para in tf.paragraphs:
        runs = [
            {
                "text": r.text,
                "bold": r.font.bold,
                "italic": r.font.italic,
                "size": r.font.size.pt if r.font.size else None,
                "font": r.font.name,
                "color": run_color(r),
            }
            for r in para.runs
        ]
        paragraphs.append({"text": para.text, "level": para.level, "runs": runs})
    return paragraphs


def table_to_list(table: Any) -> list[list[str]]:
    rows = []
    for row in table.rows:
        rows.append([cell.text for cell in row.cells])
    return rows


def shape_to_dict(shape: Any) -> dict:
    d: dict = {
        "id": shape.shape_id,
        "name": shape.name,
        "type": shape_kind_name(shape),
        "left_cm": emu_to_cm(shape.left),
        "top_cm": emu_to_cm(shape.top),
        "width_cm": emu_to_cm(shape.width),
        "height_cm": emu_to_cm(shape.height),
    }
    if shape.is_placeholder:
        phf = shape.placeholder_format
        ph_type = phf.type.name if phf.type is not None else None
        d["placeholder"] = {"type": ph_type, "idx": phf.idx}
    if shape.has_text_frame:
        d["text"] = shape.text_frame.text
        d["paragraphs"] = extract_text_frame(shape.text_frame)
    if shape.has_table:
        d["table"] = table_to_list(shape.table)
    return d


def slide_to_dict(idx: int, slide: Any) -> dict:
    shapes = [shape_to_dict(s) for s in slide.shapes]
    notes = ""
    if slide.has_notes_slide:
        notes = slide.notes_slide.notes_text_frame.text
    return {
        "index": idx,
        "shapes": shapes,
        "notes": notes,
        "layout": slide.slide_layout.name,
    }


def cell_inline(text: str) -> str:
    return text.replace("\n", " ").replace("|", "\\|")


def layout_shape_summary(shape: Any) -> dict:
    d: dict = {
        "id": shape.shape_id,
        "name": shape.name,
        "type": shape_kind_name(shape),
        "left_cm": emu_to_cm(shape.left),
        "top_cm": emu_to_cm(shape.top),
        "width_cm": emu_to_cm(shape.width),
        "height_cm": emu_to_cm(shape.height),
    }
    if shape.is_placeholder:
        phf = shape.placeholder_format
        d["placeholder"] = phf.type.name if phf.type is not None else None
    return d


def extract_design(prs: Any) -> dict:
    design: dict = {
        "slide_size": {
            "width_cm": emu_to_cm(prs.slide_width),
            "height_cm": emu_to_cm(prs.slide_height),
        },
        "layouts": [
            {
                "idx": i,
                "name": lay.name,
                "shapes": [layout_shape_summary(s) for s in lay.shapes],
            }
            for i, lay in enumerate(prs.slide_layouts)
        ],
        "slide_layout_map": {
            str(idx): slide.slide_layout.name
            for idx, slide in enumerate(prs.slides, start=1)
        },
        "theme": {"major_font": None, "minor_font": None, "colors": {}},
    }
    try:
        master = prs.slide_masters[0]
        theme_part = None
        for rel in master.part.rels.values():
            if "theme" in rel.reltype:
                theme_part = rel.target_part
                break
        theme_el = etree.fromstring(theme_part.blob)
        te = theme_el.find(qn("a:themeElements"))
        font_scheme = te.find(qn("a:fontScheme"))
        major_font = (
            font_scheme.find(qn("a:majorFont")).find(qn("a:latin")).get("typeface")
        )
        minor_font = (
            font_scheme.find(qn("a:minorFont")).find(qn("a:latin")).get("typeface")
        )
        design["theme"]["major_font"] = major_font
        design["theme"]["minor_font"] = minor_font
        clr_scheme = te.find(qn("a:clrScheme"))
        colors = {}
        for child in clr_scheme:
            name = child.tag.split("}")[-1]
            srgb = child.find(qn("a:srgbClr"))
            sysc = child.find(qn("a:sysClr"))
            if srgb is not None:
                colors[name] = srgb.get("val")
            elif sysc is not None:
                colors[name] = sysc.get("lastClr")
        design["theme"]["colors"] = colors
    except Exception:
        pass
    return design


def format_slide_layout_map(mapping: dict) -> list[str]:
    items = sorted((int(k), v) for k, v in mapping.items())
    if not items:
        return []
    parts = []
    start = end = items[0][0]
    name = items[0][1]
    for idx, lay in items[1:]:
        if lay == name and idx == end + 1:
            end = idx
            continue
        span = f"{start}" if start == end else f"{start}-{end}"
        parts.append(f"{span}:{name}")
        start = end = idx
        name = lay
    span = f"{start}" if start == end else f"{start}-{end}"
    parts.append(f"{span}:{name}")
    return parts


def design_to_markdown(design: dict) -> str:
    size = design["slide_size"]
    w = size["width_cm"]
    h = size["height_cm"]
    layouts = ", ".join(f"[{lay['idx']}] {lay['name']}" for lay in design["layouts"])
    major = design["theme"]["major_font"] or "unknown"
    minor = design["theme"]["minor_font"] or "unknown"
    colors = design["theme"]["colors"]
    colors_str = " ".join(f"{k}={v}" for k, v in colors.items())
    lines = ["## Design", ""]
    lines.append(f"- Slide size: {w} x {h} cm")
    lines.append(f"- Layouts: {layouts}")
    lines.append(f"- Theme fonts: major={major}, minor={minor}")
    lines.append(f"- Theme colors: {colors_str}")
    lines.append("")
    lines.append("### Layout shapes")
    for lay in design["layouts"]:
        shapes = lay.get("shapes", [])
        if shapes:
            parts = " ".join(
                f"[{s['name']} {s['type']} ({s['left_cm']},{s['top_cm']}) "
                f"{s['width_cm']}x{s['height_cm']}]"
                for s in shapes
            )
        else:
            parts = "(no shapes)"
        lines.append(f"- [{lay['idx']}] {lay['name']}: {parts}")
    lines.append("")
    lines.append("### Slide -> layout")
    lines.append(
        "- " + ", ".join(format_slide_layout_map(design["slide_layout_map"]))
    )
    lines.append("")
    return "\n".join(lines)


def slide_to_markdown(idx: int, slide: Any) -> str:
    lines = [f"## Slide {idx}", ""]
    for shape in slide.shapes:
        kind = shape_kind_name(shape)
        if shape.has_text_frame and shape.text_frame.text.strip():
            lines.append(f"**[{shape.name}]** ({kind})")
            for para in shape.text_frame.paragraphs:
                if para.text:
                    prefix = "  " * (para.level or 0)
                    lines.append(f"{prefix}- {para.text}")
            lines.append("")
        if shape.has_table:
            rows = table_to_list(shape.table)
            if rows:
                header = rows[0]
                lines.append(
                    f"**[{shape.name}]** (TABLE {len(rows)}x{len(header)})"
                )
                lines.append(
                    "| " + " | ".join(cell_inline(c) for c in header) + " |"
                )
                lines.append("| " + " | ".join("---" for _ in header) + " |")
                for row in rows[1:]:
                    lines.append(
                        "| " + " | ".join(cell_inline(c) for c in row) + " |"
                    )
                lines.append("")
    if (
        slide.has_notes_slide
        and slide.notes_slide.notes_text_frame.text.strip()
    ):
        lines.append("> 备注: " + slide.notes_slide.notes_text_frame.text)
        lines.append("")
    return "\n".join(lines)


def main() -> None:
    if len(sys.argv) < 2:
        print(
            "Usage: python main.py <pptx_file> [output_dir]", file=sys.stderr
        )
        print("\nArguments:", file=sys.stderr)
        print("  pptx_file  - Path to the PowerPoint file (.pptx)", file=sys.stderr)
        print(
            "  output_dir - Optional: directory for output files "
            "(default: same dir as input)",
            file=sys.stderr,
        )
        sys.exit(1)

    file_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    try:
        path = Path(file_path)
        if not path.exists():
            print(f"Error: PowerPoint file not found: {file_path}", file=sys.stderr)
            sys.exit(1)

        if path.suffix.lower() != ".pptx":
            suffix = path.suffix
            msg = f"Error: Unsupported file format: {suffix}. "
            msg += "Only .pptx is supported."
            print(msg, file=sys.stderr)
            sys.exit(1)

        prs = Presentation(str(path))

        design = extract_design(prs)

        slides_data = []
        md_lines = [f"# {path.name}", ""]
        md_lines.append(design_to_markdown(design))
        for idx, slide in enumerate(prs.slides, start=1):
            slides_data.append(slide_to_dict(idx, slide))
            md_lines.append(slide_to_markdown(idx, slide))
            md_lines.append("")

        out_dir = Path(output_dir) if output_dir else path.parent
        out_dir.mkdir(parents=True, exist_ok=True)
        base = path.stem
        md_file = out_dir / f"{base}.pptx_reader.md"
        json_file = out_dir / f"{base}.pptx_reader.json"

        md_file.write_text("\n".join(md_lines), encoding="utf-8")
        json_data = {
            "file": str(path.absolute()),
            "slide_count": len(slides_data),
            "slides": slides_data,
            "design": design,
        }
        json_file.write_text(
            json.dumps(json_data, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        print(f"Success: {path.name}")
        print(f"Slides: {len(slides_data)}")
        print(f"Markdown: {md_file}")
        print(f"JSON: {json_file}")

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
