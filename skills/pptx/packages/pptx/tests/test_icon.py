import pytest

import icon

RECT_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24">'
    '<rect x="4" y="8" width="16" height="8" fill="black"/></svg>'
)


def test_normalize_hex():
    assert icon.normalize_hex("FFFFFF") == "FFFFFF"
    assert icon.normalize_hex("#fff") == "FFFFFF"
    with pytest.raises(ValueError):
        icon.normalize_hex("ABCDE")


def test_render_tinted_tint_and_transparency():
    im = icon.render_tinted(RECT_SVG.encode(), "3C4650", 256)
    assert im.size == (256, 256)
    assert im.mode == "RGBA"
    # 2:1 glyph centered on square canvas: center opaque+tinted, top clear
    px = im.getpixel((128, 128))
    assert px[:3] == (0x3C, 0x46, 0x50)
    assert px[3] == 255
    assert im.getpixel((128, 16))[3] == 0
