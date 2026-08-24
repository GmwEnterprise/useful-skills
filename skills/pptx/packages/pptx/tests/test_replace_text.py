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


def test_set_text_across_runs_keeps_first_run_style(prs):
    from pptx.util import Cm, Pt

    slide = prs.slides[0]
    tb = slide.shapes.add_textbox(Cm(1), Cm(12), Cm(20), Cm(3))
    para = tb.text_frame.paragraphs[0]
    r1 = para.add_run()
    r1.text = "小安助手"
    r1.font.size = Pt(28)
    r1.font.bold = True
    r2 = para.add_run()
    r2.text = "产品下半年建设规划"
    r2.font.size = Pt(12)

    result = updater.apply_set_text(
        prs, {"type": "set_text", "slide": 1, "name": tb.name, "text": "新主标题"}
    )
    assert tb.text_frame.text == "新主标题"
    runs = tb.text_frame.paragraphs[0].runs
    assert len(runs) == 1
    assert runs[0].font.size.pt == 28
    assert runs[0].font.bold is True
    assert "set text" in result


def test_set_text_by_id(prs):
    from pptx.util import Cm

    slide = prs.slides[0]
    tb = slide.shapes.add_textbox(Cm(1), Cm(12), Cm(20), Cm(3))
    tb.text_frame.text = "旧标题"
    updater.apply_set_text(
        prs,
        {"type": "set_text", "slide": 1, "id": tb.shape_id, "text": "新标题"},
    )
    assert tb.text_frame.text == "新标题"


def test_set_text_multi_paragraph(prs):
    from pptx.util import Cm, Pt

    slide = prs.slides[0]
    tb = slide.shapes.add_textbox(Cm(1), Cm(12), Cm(20), Cm(3))
    tf = tb.text_frame
    tf.text = "第一段"
    tf.add_paragraph().text = "第二段"
    tf.paragraphs[0].runs[0].font.size = Pt(24)

    updater.apply_set_text(
        prs,
        {"type": "set_text", "slide": 1, "id": tb.shape_id, "text": "甲\n乙\n丙"},
    )
    paras = tf.paragraphs
    assert [p.text for p in paras] == ["甲", "乙", "丙"]
    for p in paras:
        assert p.runs[0].font.size.pt == 24


def test_set_text_ambiguous_name_raises(prs):
    from pptx.util import Cm

    slide = prs.slides[0]
    for i in range(2):
        tb = slide.shapes.add_textbox(Cm(1), Cm(13 + i), Cm(5), Cm(1))
        tb.name = "DUP"
        tb.text_frame.text = f"内容{i}"
    with pytest.raises(ValueError, match="ambiguous"):
        updater.apply_set_text(
            prs, {"type": "set_text", "slide": 1, "name": "DUP", "text": "X"}
        )


def test_set_text_missing_target_raises(prs):
    with pytest.raises(ValueError):
        updater.apply_set_text(prs, {"type": "set_text", "slide": 1, "text": "X"})


def test_set_text_with_style(prs):
    from pptx.util import Cm

    slide = prs.slides[0]
    tb = slide.shapes.add_textbox(Cm(1), Cm(12), Cm(20), Cm(3))
    tb.text_frame.text = "旧文本"
    updater.apply_set_text(
        prs,
        {
            "type": "set_text",
            "slide": 1,
            "name": tb.name,
            "text": "新文本",
            "style": {"size": 20, "color": "FF0000"},
        },
    )
    run = tb.text_frame.paragraphs[0].runs[0]
    assert run.text == "新文本"
    assert run.font.size.pt == 20
    assert str(run.font.color.rgb) == "FF0000"
