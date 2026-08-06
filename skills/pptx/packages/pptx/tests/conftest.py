from pathlib import Path

import pytest
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Cm


@pytest.fixture
def deck_path(tmp_path: Path) -> Path:
    prs = Presentation()
    prs.slide_width = Cm(25)
    prs.slide_height = Cm(19)
    blank = prs.slide_layouts[6]

    s1 = prs.slides.add_slide(blank)
    tb1 = s1.shapes.add_textbox(Cm(2), Cm(2), Cm(20), Cm(5))
    tf1 = tb1.text_frame
    tf1.text = "标题 Q3"
    para = tf1.add_paragraph()
    para.text = "ACME 公司"
    s1.notes_slide.notes_text_frame.text = "ACME 备注"

    s2 = prs.slides.add_slide(blank)
    tb2 = s2.shapes.add_textbox(Cm(2), Cm(2), Cm(20), Cm(2))
    tb2.text_frame.text = "ACME 公司 业绩"
    table = s2.shapes.add_table(
        2, 2, Cm(2), Cm(8), Cm(15), Cm(2)
    ).table
    table.cell(0, 0).text = "公司"
    table.cell(0, 1).text = "年度"
    table.cell(1, 0).text = "ACME 公司"
    table.cell(1, 1).text = "2024"

    path = tmp_path / "deck.pptx"
    prs.save(str(path))
    return path


@pytest.fixture
def prs(deck_path: Path):
    return Presentation(str(deck_path))


@pytest.fixture
def logo_path(tmp_path: Path) -> Path:
    img = Image.new("RGB", (200, 100), (200, 30, 30))
    ImageDraw.Draw(img).rectangle(
        [5, 5, 195, 95], outline=(255, 255, 255)
    )
    p = tmp_path / "logo.png"
    img.save(str(p))
    return p


def _walk_text(prs):
    for idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if shape.has_text_frame:
                yield (idx, "text", shape.text_frame.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        yield (idx, "table", cell.text)
        if slide.has_notes_slide:
            yield (
                idx,
                "notes",
                slide.notes_slide.notes_text_frame.text,
            )


@pytest.fixture
def walk_text():
    return _walk_text


def _find_shape(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


@pytest.fixture
def find_shape():
    return _find_shape


def _shapes_by_type(slide, type_name):
    return [s for s in slide.shapes if s.shape_type.name == type_name]


@pytest.fixture
def shapes_by_type():
    return _shapes_by_type


def _runs(shape):
    runs = []
    if not shape.has_text_frame:
        return runs
    for para in shape.text_frame.paragraphs:
        runs.extend(para.runs)
    return runs


@pytest.fixture
def runs():
    return _runs


@pytest.fixture
def styled_deck_path(tmp_path: Path) -> Path:
    from pptx.dml.color import RGBColor
    from pptx.util import Cm, Pt

    prs = Presentation()
    prs.slide_width = Cm(25)
    prs.slide_height = Cm(19)
    title_layout = prs.slide_layouts[0]
    s = prs.slides.add_slide(title_layout)
    s.shapes.title.text = "主标题"
    run = s.shapes.title.text_frame.paragraphs[0].runs[0]
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    run.font.name = "微软雅黑"
    run.font.size = Pt(40)
    p = tmp_path / "styled.pptx"
    prs.save(str(p))
    return p
