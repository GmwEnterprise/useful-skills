import pytest

import updater


def test_add_picture(prs, logo_path, shapes_by_type):
    result = updater.apply_add_picture(
        prs,
        {
            "type": "add_picture",
            "slide": 1,
            "path": str(logo_path),
            "left_cm": 3,
            "top_cm": 4,
            "width_cm": 5,
            "height_cm": 2,
        },
    )
    pics = shapes_by_type(prs.slides[0], "PICTURE")
    assert len(pics) == 1
    pic = pics[0]
    assert abs(pic.left / 360000 - 3) < 0.01
    assert abs(pic.top / 360000 - 4) < 0.01
    assert abs(pic.width / 360000 - 5) < 0.01
    assert abs(pic.height / 360000 - 2) < 0.01
    assert "added picture" in result


def test_add_picture_only_width(prs, logo_path, shapes_by_type):
    updater.apply_add_picture(
        prs,
        {
            "type": "add_picture",
            "slide": 2,
            "path": str(logo_path),
            "left_cm": 1,
            "top_cm": 1,
            "width_cm": 6,
        },
    )
    pic = shapes_by_type(prs.slides[1], "PICTURE")[0]
    assert abs(pic.width / 360000 - 6) < 0.01
    assert pic.height > 0


def test_add_picture_missing_path_raises(prs):
    with pytest.raises(ValueError):
        updater.apply_add_picture(
            prs,
            {"type": "add_picture", "slide": 1, "left_cm": 1, "top_cm": 1},
        )


def test_add_textbox(prs):
    result = updater.apply_add_textbox(
        prs,
        {
            "type": "add_textbox",
            "slide": 1,
            "text": "新文本",
            "left_cm": 2,
            "top_cm": 10,
            "width_cm": 8,
            "height_cm": 2,
            "style": {"bold": True, "size": 18, "color": "0066CC"},
        },
    )
    tbs = [
        s
        for s in prs.slides[0].shapes
        if s.has_text_frame and s.text_frame.text == "新文本"
    ]
    assert len(tbs) == 1
    run = tbs[0].text_frame.paragraphs[0].runs[0]
    assert run.font.bold is True
    assert run.font.size.pt == 18
    assert str(run.font.color.rgb) == "0066CC"
    assert "added textbox" in result


def test_delete_shape_by_type(prs, shapes_by_type):
    assert len(shapes_by_type(prs.slides[1], "TABLE")) == 1
    updater.apply_delete_shape(
        prs, {"type": "delete_shape", "slide": 2, "shape_type": "TABLE"}
    )
    assert shapes_by_type(prs.slides[1], "TABLE") == []


def test_delete_shape_by_name(prs):
    table_shape = [s for s in prs.slides[1].shapes if s.has_table][0]
    name = table_shape.name
    updater.apply_delete_shape(
        prs, {"type": "delete_shape", "slide": 2, "name": name}
    )
    assert name not in [s.name for s in prs.slides[1].shapes]


def test_delete_shape_no_target_raises(prs):
    with pytest.raises(ValueError):
        updater.apply_delete_shape(prs, {"type": "delete_shape", "slide": 1})


def test_format_shape_by_type(prs):
    updater.apply_format_shape(
        prs,
        {
            "type": "format_shape",
            "slide": 1,
            "shape_type": "TEXT_BOX",
            "style": {"italic": True, "color": "FF0000"},
        },
    )
    for shape in prs.slides[0].shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                assert run.font.italic is True
                assert str(run.font.color.rgb) == "FF0000"


def test_format_shape_all_shapes_on_slide(prs):
    updater.apply_format_shape(
        prs, {"type": "format_shape", "slide": 1, "style": {"bold": True}}
    )
    for shape in prs.slides[0].shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                assert run.font.bold is True


def test_move_shape_full(prs):
    shape = prs.slides[0].shapes[0]
    name = shape.name
    result = updater.apply_move_shape(
        prs,
        {
            "type": "move_shape",
            "slide": 1,
            "name": name,
            "left_cm": 10,
            "top_cm": 3,
            "width_cm": 12,
            "height_cm": 4,
        },
    )
    assert abs(shape.left / 360000 - 10) < 0.01
    assert abs(shape.top / 360000 - 3) < 0.01
    assert abs(shape.width / 360000 - 12) < 0.01
    assert abs(shape.height / 360000 - 4) < 0.01
    assert "moved" in result


def test_move_shape_partial(prs):
    shape = prs.slides[0].shapes[0]
    name = shape.name
    orig_top = shape.top
    updater.apply_move_shape(
        prs, {"type": "move_shape", "slide": 1, "name": name, "left_cm": 9}
    )
    assert abs(shape.left / 360000 - 9) < 0.01
    assert shape.top == orig_top


def test_move_shape_not_found_raises(prs):
    with pytest.raises(ValueError):
        updater.apply_move_shape(
            prs,
            {"type": "move_shape", "slide": 1, "name": "不存在", "left_cm": 1},
        )


def test_move_shape_no_fields_raises(prs):
    shape = prs.slides[0].shapes[0]
    with pytest.raises(ValueError):
        updater.apply_move_shape(
            prs, {"type": "move_shape", "slide": 1, "name": shape.name}
        )
