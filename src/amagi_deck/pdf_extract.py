"""Slide PDFs: one page = one slide. Text comes from the PDF text layer, renders from PDFium."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict

import pypdfium2 as pdfium

from .model import Deck, Slide, TextBlock
from .pptx_extract import BULLET_RE

URL_RE = re.compile(r"https?://[^\s<>()\"']+")


def join_lines(lines):
    """Undo the PDF's visual line wrapping: a line that starts in lower case continues the
    previous one. A typed bullet glyph always starts a new item and is dropped."""
    out = []
    for line in lines:
        bullet = BULLET_RE.match(line)
        if bullet:
            out.append(line[bullet.end():])
        elif out and line[:1].islower():
            out[-1] += " " + line
        else:
            out.append(line)
    return [line for line in out if line]


def extract_pdf(path: Path) -> Deck:
    doc = pdfium.PdfDocument(str(path))
    try:
        slides = []
        width_pt = height_pt = 0.0
        for i in range(len(doc)):
            page = doc[i]
            if i == 0:
                width_pt, height_pt = page.get_size()
            textpage = page.get_textpage()
            raw = textpage.get_text_bounded() or ""
            textpage.close()
            page.close()
            lines = [" ".join(line.split()) for line in raw.replace("\r", "\n").split("\n")]
            lines = [line for line in lines if line]
            slide = Slide(number=i + 1)
            if lines:
                slide.title = lines[0][:80]
                body = join_lines(lines[1:])
                if body:
                    slide.text.append(TextBlock(shape="page text", lines=[(0, t) for t in body]))
            for url in URL_RE.findall(raw):
                url = url.rstrip(".,;:")
                if url not in slide.links:
                    slide.links.append(url)
            slides.append(slide)
        title = ""
        try:
            title = (doc.get_metadata_dict().get("Title") or "").strip()
        except Exception:
            pass
    finally:
        doc.close()
    deck = Deck(
        source=path.name,
        kind="pdf",
        slides=slides,
        width_in=round(width_pt / 72.0, 2),
        height_in=round(height_pt / 72.0, 2),
        title=title or (slides[0].title if slides else ""),
    )
    return deck


def render_pdf_pages(pdf_path: Path, targets: Dict[int, Path], width: int) -> Dict[int, Path]:
    """Render 1-based page numbers to PNG files at ``width`` pixels wide."""
    done: Dict[int, Path] = {}
    doc = pdfium.PdfDocument(str(pdf_path))
    try:
        for number, out in sorted(targets.items()):
            if number < 1 or number > len(doc):
                continue
            page = doc[number - 1]
            page_w, _ = page.get_size()
            scale = width / page_w if page_w else 1.0
            bitmap = page.render(scale=scale, draw_annots=True)
            image = bitmap.to_pil().convert("RGB")
            image.save(out, "PNG", optimize=True)
            bitmap.close()
            page.close()
            done[number] = out
    finally:
        doc.close()
    return done


def page_count(pdf_path: Path) -> int:
    doc = pdfium.PdfDocument(str(pdf_path))
    try:
        return len(doc)
    finally:
        doc.close()
