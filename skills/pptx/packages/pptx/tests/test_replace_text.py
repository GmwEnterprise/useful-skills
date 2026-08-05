import pytest

import updater


def test_replace_all_slides(prs, walk_text):
    result = updater.apply_replace_text(
        prs,
        {"type": "replace_text", "find": "ACME 公司", "replace": "新公司"},
    )
    joined = "|||".join(t for _, _, t in walk_text(prs))
    assert "ACME 公司" not in joined
    assert "新公司" in joined
    assert "3 hit" in result


def test_replace_scoped_to_slide(prs, walk_text):
    updater.apply_replace_text(
        prs,
        {
            "type": "replace_text",
            "slide": 1,
            "find": "ACME 公司",
            "replace": "新公司",
        },
    )
    by_slide: dict[int, list[str]] = {}
    for idx, _, t in walk_text(prs):
        by_slide.setdefault(idx, []).append(t)
    assert "新公司" in " ".join(by_slide[1])
    assert "ACME 公司" not in " ".join(by_slide[1])
    assert "ACME 公司" in " ".join(by_slide[2])
    assert "新公司" not in " ".join(by_slide[2])


def test_replace_in_table_cell(prs, walk_text):
    updater.apply_replace_text(
        prs, {"type": "replace_text", "find": "2024", "replace": "2025"}
    )
    table_texts = [t for _, o, t in walk_text(prs) if o == "table"]
    assert "2025" in table_texts
    assert "2024" not in table_texts


def test_replace_in_notes(prs, walk_text):
    updater.apply_replace_text(
        prs,
        {"type": "replace_text", "find": "ACME 备注", "replace": "新备注"},
    )
    notes_texts = [t for _, o, t in walk_text(prs) if o == "notes"]
    assert "新备注" in notes_texts
    assert all("ACME 备注" not in t for t in notes_texts)


def test_replace_with_style(prs):
    updater.apply_replace_text(
        prs,
        {
            "type": "replace_text",
            "slide": 1,
            "find": "Q3",
            "replace": "Q4",
            "style": {
                "bold": True,
                "size": 40,
                "color": "#FF0000",
                "font": "微软雅黑",
            },
        },
    )
    run = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    assert run.text == "标题 Q4"
    assert run.font.bold is True
    assert run.font.size.pt == 40
    assert str(run.font.color.rgb) == "FF0000"
    assert run.font.name == "微软雅黑"


def test_replace_style_only_on_hit_run(prs):
    updater.apply_replace_text(
        prs,
        {
            "type": "replace_text",
            "slide": 1,
            "find": "Q3",
            "replace": "Q4",
            "style": {"bold": True},
        },
    )
    paras = prs.slides[0].shapes[0].text_frame.paragraphs
    hit_run = paras[0].runs[0]
    other_run = paras[1].runs[0]
    assert hit_run.font.bold is True
    assert other_run.font.bold is None


def test_replace_no_match(prs):
    result = updater.apply_replace_text(
        prs, {"type": "replace_text", "find": "不存在", "replace": "X"}
    )
    assert "0 hit" in result


def test_replace_missing_find_raises(prs):
    with pytest.raises(ValueError):
        updater.apply_replace_text(
            prs, {"type": "replace_text", "replace": "X"}
        )
