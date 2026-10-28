import json
import subprocess
import sys

import pytest

from amagi_deck import __version__
from amagi_deck.cli import main


def test_version(capsys):
    with pytest.raises(SystemExit) as exc:
        main(["--version"])
    assert exc.value.code == 0
    assert capsys.readouterr().out.strip() == "amagi-deck " + __version__


def test_renderers_list(capsys):
    assert main(["--renderers"]) == 0
    out = capsys.readouterr().out
    assert "powerpoint" in out and "soffice" in out and "sketch" in out


def test_no_input_is_usage_error(capsys):
    assert main([]) == 2
    assert "give a deck file" in capsys.readouterr().err


def test_missing_file_exit_1(capsys, tmp_path):
    assert main([str(tmp_path / "nope.pptx")]) == 1
    assert "file not found" in capsys.readouterr().err


def test_bad_slides_exit_2(deck_pptx, capsys):
    assert main([str(deck_pptx), "--slides", "9", "--render", "none", "-o", "-"]) == 2
    assert "outside 1-7" in capsys.readouterr().err


def test_default_output_in_cwd(deck_pptx, tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)
    assert main([str(deck_pptx), "--render", "sketch", "--width", "320", "--slides", "1-2"]) == 0
    assert (tmp_path / "fixture.md").is_file()
    assert sorted(p.name for p in (tmp_path / "fixture_slides").iterdir()) == ["slide-01.png", "slide-02.png"]
    err = capsys.readouterr().err
    assert "wrote fixture.md (2 slides, 2 images via sketch, 1 chart, 0 tables, 1 with notes)" in err
    assert "note: slide images are sketch renders" in err


def test_stdout_markdown_without_notes(deck_pptx, capsys):
    assert main([str(deck_pptx), "-o", "-", "--render", "none", "--no-notes"]) == 0
    out = capsys.readouterr().out
    assert out.startswith("# Fixture deck")
    assert "Speaker notes" not in out


def test_json_output(deck_pptx, tmp_path, capsys):
    target = tmp_path / "deck.json"
    assert main([str(deck_pptx), "--json", "--render", "none", "-o", str(target)]) == 0
    data = json.loads(target.read_text(encoding="utf-8"))
    assert data["title"] == "Fixture deck"
    assert data["slides"][1]["charts"][0]["series"][0]["values"] == [10.0, 20.5, 31.0]
    assert data["slides"][6]["hidden"] is True


def test_module_entry_point(deck_pptx):
    proc = subprocess.run([sys.executable, "-m", "amagi_deck", str(deck_pptx), "-o", "-", "--render", "none"],
                          capture_output=True, text=True, encoding="utf-8")
    assert proc.returncode == 0, proc.stderr
    assert "## Slide 2: Revenue chart" in proc.stdout
