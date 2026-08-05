import pytest

import updater


def test_insert_slide_at_position(prs):
    n = len(prs.slides._sldIdLst)
    result = updater.apply_insert_slide(
        prs, {"type": "insert_slide", "position": 1}
    )
    assert len(prs.slides._sldIdLst) == n + 1
    new_slide = prs.slides[0]
    assert not any(
        s.has_text_frame and s.text_frame.text.strip()
        for s in new_slide.shapes
    )
    assert "position 1" in result


def test_insert_slide_append(prs):
    n = len(prs.slides._sldIdLst)
    updater.apply_insert_slide(prs, {"type": "insert_slide"})
    assert len(prs.slides._sldIdLst) == n + 1
    slide1 = prs.slides[0]
    assert any(
        "标题 Q3" in s.text_frame.text
        for s in slide1.shapes
        if s.has_text_frame
    )


def test_insert_slide_layout_index(prs):
    result = updater.apply_insert_slide(
        prs, {"type": "insert_slide", "layout": 0}
    )
    assert "layout=" in result


def test_delete_slide(prs):
    n = len(prs.slides._sldIdLst)
    result = updater.apply_delete_slide(
        prs, {"type": "delete_slide", "slide": 1}
    )
    assert len(prs.slides._sldIdLst) == n - 1
    new_first = prs.slides[0]
    assert any(s.has_table for s in new_first.shapes)
    assert "removed slide 1" in result


def test_move_slide(prs):
    result = updater.apply_move_slide(
        prs, {"type": "move_slide", "source": 1, "target": 2}
    )
    first = prs.slides[0]
    assert any(s.has_table for s in first.shapes)
    assert "moved slide 1 -> position 2" in result


def test_delete_slide_out_of_range_raises(prs):
    with pytest.raises(ValueError):
        updater.apply_delete_slide(
            prs, {"type": "delete_slide", "slide": 5}
        )


def test_move_slide_target_out_of_range_raises(prs):
    with pytest.raises(ValueError):
        updater.apply_move_slide(
            prs, {"type": "move_slide", "source": 1, "target": 5}
        )


def test_insert_slide_position_out_of_range_raises(prs):
    with pytest.raises(ValueError):
        updater.apply_insert_slide(
            prs, {"type": "insert_slide", "position": 99}
        )


def test_move_slide_missing_target_raises(prs):
    with pytest.raises(ValueError):
        updater.apply_move_slide(prs, {"type": "move_slide", "source": 1})
