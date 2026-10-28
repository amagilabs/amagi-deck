"""convert(): deck file in, Deck + markdown out, slide PNGs on disk."""

from __future__ import annotations

import os
import tempfile
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

from . import render
from .markdown import deck_md
from .model import Deck
from .pdf_extract import extract_pdf, render_pdf_pages
from .pptx_extract import extract_pptx

PPTX_EXTS = {".pptx"}
CONVERT_EXTS = {".key", ".ppt", ".pps", ".odp", ".pptm", ".ppsx", ".potx"}
PDF_EXTS = {".pdf"}
RENDERERS = ("auto", "powerpoint", "soffice", "sketch", "none")


class DeckError(Exception):
    pass


def parse_slide_spec(spec: Optional[str], total: int) -> Optional[Set[int]]:
    """'1-3,7' -> {1, 2, 3, 7}. None/'' means every slide."""
    if not spec:
        return None
    chosen: Set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        try:
            if "-" in part:
                a, b = part.split("-", 1)
                lo, hi = int(a), int(b) if b.strip() else total
                chosen.update(range(lo, hi + 1))
            else:
                chosen.add(int(part))
        except ValueError:
            raise DeckError("bad --slides value %r (use e.g. 1-3,7)" % spec)
    bad = sorted(n for n in chosen if n < 1 or n > total)
    if bad:
        raise DeckError("--slides %s is outside 1-%d" % (",".join(map(str, bad)), total))
    return chosen


def _to_pptx(src: Path, tmp: Path) -> Path:
    """Convert .key/.ppt/.odp/etc. to .pptx with LibreOffice, or PowerPoint where it can open it."""
    errors: List[str] = []
    if render.find_soffice():
        try:
            return render.soffice_convert(src, tmp, "pptx:Impress MS PowerPoint 2007 XML")
        except render.RenderError as exc:
            errors.append(str(exc))
    if src.suffix.lower() != ".key" and render.has_powerpoint():
        try:
            return render.powerpoint_to_pptx(src, tmp / (src.stem + ".pptx"))
        except render.RenderError as exc:
            errors.append(str(exc))
    if src.suffix.lower() == ".key":
        hint = ("Keynote files need LibreOffice (soffice) to convert. Or open the deck in Keynote "
                "and use File > Export To > PowerPoint (or PDF), then run amagi-deck on that file.")
    else:
        hint = "%s files need LibreOffice or Microsoft PowerPoint to convert." % src.suffix
    raise DeckError(hint + (" (" + "; ".join(errors) + ")" if errors else ""))


def _clear_old_images(images_dir: Path) -> None:
    if images_dir.is_dir():
        for old in images_dir.glob("slide-*.png"):
            old.unlink()


def _render(deck: Deck, src: Path, pptx: Optional[Path], targets: Dict[int, Path],
            renderer: str, width: int, hidden: List[bool]) -> Dict[int, Path]:
    if deck.kind == "pdf":
        deck.renderer = "pdfium"
        return render_pdf_pages(src, targets, width)

    height = round(width * deck.height_in / deck.width_in) if deck.width_in else round(width * 9 / 16)

    def attempt(name: str) -> Dict[int, Path]:
        if name == "powerpoint":
            return render.render_powerpoint(pptx, targets, width, height)
        if name == "soffice":
            return render.render_soffice(src, targets, width, hidden)
        if name == "sketch":
            from .sketch import render_sketch

            return render_sketch(pptx, targets, width)
        raise DeckError("unknown renderer %r" % name)

    if renderer != "auto":
        try:
            done = attempt(renderer)
        except render.RenderError as exc:
            raise DeckError("--render %s failed: %s" % (renderer, exc))
        deck.renderer = renderer
        return done

    order = []
    if render.has_powerpoint():
        order.append("powerpoint")
    if render.find_soffice():
        order.append("soffice")
    order.append("sketch")
    for name in order:
        try:
            done = attempt(name)
        except render.RenderError as exc:
            deck.warnings.append("%s renderer failed (%s); fell back to the next one" % (name, exc))
            continue
        deck.renderer = name
        if name == "sketch":
            deck.warnings.append(
                "slide images are sketch renders: no PowerPoint or LibreOffice was found, so layout "
                "is approximate and theme colors, gradients and effects are missing. Install "
                "LibreOffice for faithful renders.")
        return done
    return {}


def convert(src: Path, out_md: Optional[Path] = None, images_dir: Optional[Path] = None,
            renderer: str = "auto", width: int = 1280, slides: Optional[str] = None,
            notes: bool = True) -> Tuple[Deck, str]:
    """Parse ``src``, render the chosen slides, and return (deck, markdown text).

    ``out_md`` decides where image links point from; it is not written here.
    """
    src = Path(src)
    if not src.is_file():
        raise DeckError("file not found: %s" % src)
    if renderer not in RENDERERS:
        raise DeckError("unknown renderer %r (choose from %s)" % (renderer, ", ".join(RENDERERS)))
    if width < 64 or width > 8000:
        raise DeckError("--width must be between 64 and 8000 pixels")
    ext = src.suffix.lower()
    if ext not in PPTX_EXTS | CONVERT_EXTS | PDF_EXTS:
        raise DeckError("unsupported file type %r (use .pptx, .key, .ppt, .odp or .pdf)" % src.suffix)
    base = Path(out_md).parent if out_md else Path.cwd()
    stem = Path(out_md).stem if out_md else src.stem

    with tempfile.TemporaryDirectory(prefix="amagi-deck-") as tmp:
        pptx = None
        try:
            if ext in PDF_EXTS:
                deck = extract_pdf(src)
            else:
                pptx = src if ext in PPTX_EXTS else _to_pptx(src, Path(tmp))
                deck = extract_pptx(pptx, with_notes=notes)
                if pptx is not src:
                    deck.source = src.name
                    deck.kind = ext.lstrip(".")
        except DeckError:
            raise
        except (zipfile.BadZipFile, KeyError) as exc:
            raise DeckError("%s is not a valid %s file (%s)" % (src.name, ext, exc))
        except Exception as exc:
            if type(exc).__name__ in ("PackageNotFoundError", "PdfiumError"):
                raise DeckError("could not open %s: %s" % (src.name, exc))
            raise

        total = len(deck.slides)
        hidden = [s.hidden for s in deck.slides]
        chosen = parse_slide_spec(slides, total)
        if chosen is not None:
            deck.slides = [s for s in deck.slides if s.number in chosen]

        if renderer != "none" and deck.slides:
            images_dir = Path(images_dir) if images_dir else base / (stem + "_slides")
            images_dir.mkdir(parents=True, exist_ok=True)
            _clear_old_images(images_dir)
            pad = max(2, len(str(total)))
            targets = {s.number: images_dir / ("slide-%0*d.png" % (pad, s.number)) for s in deck.slides}
            done = _render(deck, src, pptx, targets, renderer, width, hidden)
            missing = [n for n in targets if n not in done]
            if missing:
                deck.warnings.append("no image for slide(s) %s" % ", ".join(map(str, sorted(missing))))
            for s in deck.slides:
                if s.number in done:
                    s.image = Path(os.path.relpath(done[s.number], base)).as_posix()

    return deck, deck_md(deck, include_notes=notes)
