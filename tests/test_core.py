import pytest
from PIL import Image

from amagi_deck import render
from amagi_deck.core import DeckError, convert, parse_slide_spec


def test_parse_slide_spec():
    assert parse_slide_spec(None, 5) is None
    assert parse_slide_spec("", 5) is None
    assert parse_slide_spec("1-3,5", 5) == {1, 2, 3, 5}
    assert parse_slide_spec(" 2 , 4- ", 5) == {2, 4, 5}
    with pytest.raises(DeckError, match="bad --slides"):
        parse_slide_spec("a-b", 5)
    with pytest.raises(DeckError, match="outside 1-5"):
        parse_slide_spec("4-6", 5)
    with pytest.raises(DeckError, match="outside"):
        parse_slide_spec("0", 5)


def test_convert_without_images(deck_pptx, tmp_path):
    deck, md = convert(deck_pptx, out_md=tmp_path / "out.md", renderer="none")
    assert all(s.image == "" for s in deck.slides)
    assert not (tmp_path / "out_slides").exists()
    assert "Slide images: none" in md
    assert "| Q2 | 20.5 | 9 |" in md


def test_convert_sketch_writes_pngs_with_relative_links(deck_pptx, tmp_path):
    out = tmp_path / "review" / "deck.md"
    deck, md = convert(deck_pptx, out_md=out, renderer="sketch", width=640, slides="1,7")
    assert [s.number for s in deck.slides] == [1, 7]
    files = sorted(p.name for p in (tmp_path / "review" / "deck_slides").iterdir())
    assert files == ["slide-01.png", "slide-07.png"]
    assert deck.slides[0].image == "deck_slides/slide-01.png"
    assert "![Slide 7](deck_slides/slide-07.png)" in md
    assert deck.renderer == "sketch"
    assert any("sketch renders" in w for w in deck.warnings)
    with Image.open(tmp_path / "review" / "deck_slides" / "slide-01.png") as im:
        assert im.size == (640, 360)
        assert im.getpixel((320, 300)) == (0x12, 0x2B, 0x45)  # the navy slide background


def test_rerun_clears_old_images(deck_pptx, tmp_path):
    images = tmp_path / "imgs"
    convert(deck_pptx, out_md=tmp_path / "a.md", images_dir=images, renderer="sketch", width=320)
    assert len(list(images.glob("slide-*.png"))) == 7
    convert(deck_pptx, out_md=tmp_path / "a.md", images_dir=images, renderer="sketch", width=320, slides="2")
    assert [p.name for p in images.glob("slide-*.png")] == ["slide-02.png"]


def test_auto_falls_back_to_sketch(deck_pptx, tmp_path, no_office):
    deck, _ = convert(deck_pptx, out_md=tmp_path / "x.md", width=320, slides="2")
    assert deck.renderer == "sketch"
    assert any("no working PowerPoint or LibreOffice" in w for w in deck.warnings)


def test_auto_records_failed_renderer(deck_pptx, tmp_path, monkeypatch):
    def broken(*args, **kwargs):
        raise render.RenderError("COM said no")

    monkeypatch.setattr(render, "has_powerpoint", lambda: True)
    monkeypatch.setattr(render, "find_soffice", lambda: None)
    monkeypatch.setattr(render, "render_powerpoint", broken)
    deck, _ = convert(deck_pptx, out_md=tmp_path / "x.md", width=320, slides="1")
    assert deck.renderer == "sketch"
    assert "powerpoint renderer failed (COM said no)" in deck.warnings[0]


def test_explicit_renderer_failure_is_an_error(deck_pptx, tmp_path, no_office):
    with pytest.raises(DeckError, match="--render soffice failed"):
        convert(deck_pptx, out_md=tmp_path / "x.md", renderer="soffice")


def test_pdf_always_renders_with_pdfium(deck_pdf, tmp_path):
    deck, md = convert(deck_pdf, out_md=tmp_path / "p.md", renderer="auto", width=480)
    assert deck.renderer == "pdfium"
    with Image.open(tmp_path / "p_slides" / "slide-02.png") as im:
        assert im.size == (480, 270)


def test_input_errors(tmp_path, deck_pptx):
    with pytest.raises(DeckError, match="file not found"):
        convert(tmp_path / "missing.pptx")
    txt = tmp_path / "notes.txt"
    txt.write_text("hi")
    with pytest.raises(DeckError, match="unsupported file type"):
        convert(txt)
    junk = tmp_path / "junk.pptx"
    junk.write_bytes(b"not a zip")
    with pytest.raises(DeckError, match="junk.pptx"):
        convert(junk, renderer="none")
    with pytest.raises(DeckError, match="--width"):
        convert(deck_pptx, width=10)
    with pytest.raises(DeckError, match="unknown renderer"):
        convert(deck_pptx, renderer="magic")


def test_keynote_without_converters_explains_export(tmp_path, no_office):
    key = tmp_path / "talk.key"
    key.write_bytes(b"PK\x03\x04")
    with pytest.raises(DeckError, match="File > Export To > PowerPoint"):
        convert(key)
