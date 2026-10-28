import shutil
import subprocess
from pathlib import Path

import pytest
from PIL import Image

from amagi_deck import render
from amagi_deck.core import convert


def test_find_soffice_env_override(monkeypatch, tmp_path):
    fake = tmp_path / "soffice"
    fake.write_text("")
    monkeypatch.setenv("AMAGI_DECK_SOFFICE", str(fake))
    assert render.find_soffice() == str(fake)
    monkeypatch.setenv("AMAGI_DECK_SOFFICE", str(tmp_path / "nope"))
    assert render.find_soffice() is None


def test_powerpoint_can_be_disabled(monkeypatch):
    monkeypatch.setenv("AMAGI_DECK_NO_POWERPOINT", "1")
    assert render.has_powerpoint() is False
    assert "not available" in render.available()["powerpoint"]


def test_soffice_convert_command(monkeypatch, tmp_path):
    calls = []

    def fake_run(cmd, env=None):
        calls.append(cmd)
        (tmp_path / "talk.pdf").write_bytes(b"%PDF")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(render, "find_soffice", lambda: "soffice")
    monkeypatch.setattr(render, "_run", fake_run)
    out = render.soffice_convert(Path("talk.key"), tmp_path, render.PDF_HIDDEN_FILTER)
    assert out == tmp_path / "talk.pdf"
    cmd = calls[0]
    assert cmd[0] == "soffice" and "--headless" in cmd
    assert cmd[1].startswith("-env:UserInstallation=file:")
    assert cmd[cmd.index("--convert-to") + 1] == render.PDF_HIDDEN_FILTER
    assert cmd[-1] == "talk.key"


def test_soffice_convert_reports_failure(monkeypatch, tmp_path):
    monkeypatch.setattr(render, "find_soffice", lambda: "soffice")
    monkeypatch.setattr(render, "_run", lambda cmd, env=None: subprocess.CompletedProcess(
        cmd, 1, "", "Error: source file could not be loaded"))
    with pytest.raises(render.RenderError, match="could not be loaded"):
        render.soffice_convert(Path("bad.pptx"), tmp_path, "pdf")


def _fake_pdf_export(pdf_source):
    def fake(src, out_dir, target):
        dest = Path(out_dir) / (Path(src).stem + ".pdf")
        shutil.copy(pdf_source, dest)
        return dest
    return fake


def test_render_soffice_maps_pages_when_hidden_slides_were_dropped(monkeypatch, tmp_path, deck_pdf):
    # deck_pdf has 2 pages; a 3-slide deck whose middle slide is hidden.
    monkeypatch.setattr(render, "soffice_convert", _fake_pdf_export(deck_pdf))
    targets = {n: tmp_path / ("s%d.png" % n) for n in (1, 2, 3)}
    done = render.render_soffice(Path("deck.pptx"), targets, 320, [False, True, False])
    assert sorted(done) == [1, 3]
    with Image.open(done[3]) as im:
        assert im.width == 320


def test_render_soffice_rejects_page_mismatch(monkeypatch, tmp_path, deck_pdf):
    monkeypatch.setattr(render, "soffice_convert", _fake_pdf_export(deck_pdf))
    with pytest.raises(render.RenderError, match="2 pages for 5 slides"):
        render.render_soffice(Path("deck.pptx"), {1: tmp_path / "a.png"}, 320, [False] * 5)


def test_powerpoint_runs_powershell_in_mta(monkeypatch, tmp_path):
    seen = {}

    def fake_run(cmd, env=None):
        seen["cmd"], seen["env"] = cmd, env
        Path(env["AMAGI_DECK_OUT"], "2.png").write_bytes(b"png")
        return subprocess.CompletedProcess(cmd, 0, "", "")

    monkeypatch.setattr(render, "has_powerpoint", lambda: True)
    monkeypatch.setattr(render, "_run", fake_run)
    out = tmp_path / "slide-02.png"
    done = render.render_powerpoint(Path("deck.pptx"), {2: out}, 1280, 720)
    assert done == {2: out} and out.read_bytes() == b"png"
    assert seen["cmd"][:2] == ["powershell", "-Mta"]
    assert seen["env"]["AMAGI_DECK_SLIDES"] == "2"
    assert (seen["env"]["AMAGI_DECK_W"], seen["env"]["AMAGI_DECK_H"]) == ("1280", "720")


def test_powerpoint_error_is_reported(monkeypatch):
    monkeypatch.setattr(render, "has_powerpoint", lambda: True)
    monkeypatch.setattr(render, "_run", lambda cmd, env=None: subprocess.CompletedProcess(
        cmd, 1, "", "\nException calling Open: file is corrupt\n"))
    with pytest.raises(render.RenderError, match="file is corrupt"):
        render.render_powerpoint(Path("deck.pptx"), {1: Path("x.png")}, 100, 56)


# --- real renderers, only where installed -------------------------------------------------

@pytest.mark.skipif(not render.has_powerpoint(), reason="Microsoft PowerPoint not installed")
def test_real_powerpoint_render(deck_pptx, tmp_path):
    deck, _ = convert(deck_pptx, out_md=tmp_path / "d.md", renderer="powerpoint", width=800, slides="2,7")
    assert deck.renderer == "powerpoint"
    for n in ("02", "07"):  # hidden slide 7 renders too
        with Image.open(tmp_path / "d_slides" / ("slide-%s.png" % n)) as im:
            assert im.size == (800, 450)


@pytest.mark.skipif(render.find_soffice() is None, reason="LibreOffice not installed")
def test_real_soffice_render(deck_pptx, tmp_path):
    deck, _ = convert(deck_pptx, out_md=tmp_path / "d.md", renderer="soffice", width=800, slides="2")
    assert deck.renderer == "soffice"
    with Image.open(tmp_path / "d_slides" / "slide-02.png") as im:
        assert im.width == 800


@pytest.mark.skipif(render.find_soffice() is None, reason="LibreOffice not installed")
def test_real_soffice_converts_odp(deck_pptx, tmp_path):
    odp = render.soffice_convert(deck_pptx, tmp_path, "odp")
    deck, md = convert(odp, out_md=tmp_path / "o.md", renderer="none")
    assert deck.kind == "odp"
    assert "| Q2 | 20.5 | 9 |" in md
