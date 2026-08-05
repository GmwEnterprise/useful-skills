#!/usr/bin/env python3
import json
import sys
from pathlib import Path
from typing import Any

from pptx import Presentation


def emu_to_cm(emu: int | None) -> float | None:
    if emu is None:
        return None
    return round(emu / 360000, 2)


def shape_kind_name(shape: Any) -> str:
    try:
        return str(shape.shape_type)
    except Exception:
        return "UNKNOWN"


def extract_text_frame(tf: Any) -> list[dict]:
    paragraphs = []
    for para in tf.paragraphs:
        runs = [
            {
                "text": r.text,
                "bold": r.font.bold,
                "italic": r.font.italic,
                "size": r.font.size.pt if r.font.size else None,
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
        "name": shape.name,
        "type": shape_kind_name(shape),
        "left_cm": emu_to_cm(shape.left),
        "top_cm": emu_to_cm(shape.top),
        "width_cm": emu_to_cm(shape.width),
        "height_cm": emu_to_cm(shape.height),
    }
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
    return {"index": idx, "shapes": shapes, "notes": notes}


def cell_inline(text: str) -> str:
    return text.replace("\n", " ").replace("|", "\\|")


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

        slides_data = []
        md_lines = [f"# {path.name}", ""]
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
