import json
import sys

import pytest

import main


def test_reader_generates_md_and_json(deck_path, tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    monkeypatch.setattr(sys, "argv", ["main.py", str(deck_path), str(out_dir)])
    main.main()

    base = deck_path.stem
    md = out_dir / f"{base}.pptx_reader.md"
    js = out_dir / f"{base}.pptx_reader.json"
    assert md.exists() and js.exists()

    md_text = md.read_text(encoding="utf-8")
    assert "## Slide 1" in md_text
    assert "## Slide 2" in md_text
    assert "标题 Q3" in md_text


def test_reader_json_has_table_and_notes(deck_path, tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    monkeypatch.setattr(sys, "argv", ["main.py", str(deck_path), str(out_dir)])
    main.main()

    data = json.loads(
        (out_dir / f"{deck_path.stem}.pptx_reader.json").read_text(
            encoding="utf-8"
        )
    )
    assert data["slide_count"] == 2
    assert any(sh.get("table") for sh in data["slides"][1]["shapes"])
    assert data["slides"][0]["notes"] == "ACME 备注"


def test_reader_invalid_suffix_exits(tmp_path, monkeypatch):
    p = tmp_path / "f.pdf"
    p.write_bytes(b"%PDF-1.4")
    monkeypatch.setattr(sys, "argv", ["main.py", str(p)])
    with pytest.raises(SystemExit):
        main.main()


def test_reader_missing_file_exits(tmp_path, monkeypatch):
    monkeypatch.setattr(
        sys, "argv", ["main.py", str(tmp_path / "no.pptx")]
    )
    with pytest.raises(SystemExit):
        main.main()


def _all_runs(data):
    runs = []
    for sl in data["slides"]:
        for sh in sl["shapes"]:
            for para in sh.get("paragraphs", []):
                runs.extend(para.get("runs", []))
    return runs


def test_reader_design_styled_deck(styled_deck_path, tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    monkeypatch.setattr(
        sys, "argv", ["main.py", str(styled_deck_path), str(out_dir)]
    )
    main.main()

    base = styled_deck_path.stem
    data = json.loads(
        (out_dir / f"{base}.pptx_reader.json").read_text(encoding="utf-8")
    )

    size = data["design"]["slide_size"]
    assert isinstance(size["width_cm"], (int, float))
    assert isinstance(size["height_cm"], (int, float))

    layouts = data["design"]["layouts"]
    assert layouts
    for lay in layouts:
        assert "idx" in lay
        assert "name" in lay

    major = data["design"]["theme"]["major_font"]
    assert isinstance(major, str) and major
    assert "accent1" in data["design"]["theme"]["colors"]

    title_shapes = [
        sh
        for sl in data["slides"]
        for sh in sl["shapes"]
        if sh.get("placeholder", {}).get("type")
        in ("TITLE", "CENTER_TITLE")
    ]
    assert title_shapes

    assert any(
        isinstance(sl["layout"], str) and sl["layout"] for sl in data["slides"]
    )

    runs = _all_runs(data)
    assert any(
        r.get("color") == "#FF0000" and r.get("font") == "微软雅黑"
        for r in runs
    )

    md_text = (out_dir / f"{base}.pptx_reader.md").read_text(encoding="utf-8")
    assert "## Design" in md_text


def test_reader_design_default_deck(deck_path, tmp_path, monkeypatch):
    out_dir = tmp_path / "out"
    monkeypatch.setattr(sys, "argv", ["main.py", str(deck_path), str(out_dir)])
    main.main()

    data = json.loads(
        (out_dir / f"{deck_path.stem}.pptx_reader.json").read_text(
            encoding="utf-8"
        )
    )
    assert "design" in data
    assert data["design"]["layouts"]
