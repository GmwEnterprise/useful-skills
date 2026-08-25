#!/usr/bin/env python3
"""Fetch iconify icons and render them as tinted transparent PNGs.

Pipeline: iconify API (SVG only) -> svglib render on white -> luminance
inversion as alpha channel -> tint with target color -> crop glyph bbox ->
center on a square canvas.
"""
import argparse
import io
import sys
import urllib.error
import urllib.request
from pathlib import Path

from PIL import Image, ImageOps
from reportlab.graphics import renderPM
from svglib.svglib import svg2rlg

API_URL = "https://api.iconify.design/{prefix}/{name}.svg"


def normalize_hex(color: str) -> str:
    c = color.lstrip("#")
    if len(c) == 3:
        c = "".join(ch * 2 for ch in c)
    if len(c) != 6:
        raise ValueError(f"invalid color: {color!r} (expect RRGGBB)")
    return c.upper()


def fetch_svg(prefix: str, name: str) -> bytes:
    url = API_URL.format(prefix=prefix, name=name)
    req = urllib.request.Request(url, headers={"User-Agent": "pptx-skill"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read()


def render_tinted(svg_bytes: bytes, color: str, size: int) -> Image.Image:
    rgb = tuple(int(color[i : i + 2], 16) for i in (0, 2, 4))

    drawing = svg2rlg(io.StringIO(svg_bytes.decode("utf-8")))
    buf = io.BytesIO()
    renderPM.drawToFile(drawing, buf, fmt="PNG", bg=0xFFFFFFFF, dpi=96)
    im = Image.open(buf).convert("L")

    # black glyph on white -> invert luminance as alpha (anti-aliased edges
    # are preserved); tint the whole layer with the target color
    alpha = Image.eval(im, lambda v: 255 - v)
    out = Image.new("RGBA", im.size, rgb + (0,))
    out.putalpha(alpha)

    bbox = alpha.getbbox()
    glyph = out.crop(bbox) if bbox else out
    glyph = ImageOps.contain(glyph, (size, size))
    canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    canvas.alpha_composite(
        glyph, ((size - glyph.width) // 2, (size - glyph.height) // 2)
    )
    return canvas


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Download iconify icons as tinted transparent PNGs"
    )
    parser.add_argument("out_dir", help="output directory for PNG files")
    parser.add_argument("names", nargs="+", help="icon names (e.g. alert route)")
    parser.add_argument("--color", default="FFFFFF", help="RRGGBB tint color")
    parser.add_argument("--size", type=int, default=256, help="square canvas px")
    parser.add_argument("--prefix", default="mdi", help="iconify prefix (set)")
    args = parser.parse_args()

    color = normalize_hex(args.color)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    failed = []
    for name in args.names:
        try:
            svg_bytes = fetch_svg(args.prefix, name)
        except urllib.error.HTTPError as e:
            print(f"Error: icon {args.prefix}:{name} -> HTTP {e.code}")
            failed.append(name)
            continue
        except OSError as e:
            print(f"Error: icon {args.prefix}:{name} -> {e}")
            failed.append(name)
            continue
        png = render_tinted(svg_bytes, color, args.size)
        out_file = out_dir / f"{name}.png"
        png.save(str(out_file))
        print(f"Saved: {out_file} ({args.prefix}:{name}, #{color}, {args.size}px)")

    if failed:
        print(f"Failed: {len(failed)} icon(s): {', '.join(failed)}", file=sys.stderr)
        sample = API_URL.format(prefix=args.prefix, name="<name>")
        print(
            f"Hint: verify the icon exists first, e.g. "
            f"curl -o /dev/null -s -w '%{{http_code}}' {sample}",
            file=sys.stderr,
        )
        sys.exit(1)
    print(f"Success: {len(args.names) - len(failed)} icon(s) -> {out_dir}")


if __name__ == "__main__":
    main()
