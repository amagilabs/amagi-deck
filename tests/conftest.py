"""Fixtures: a .pptx with every feature amagi-deck reads, and a small hand-written slide PDF."""

from __future__ import annotations

import io
from pathlib import Path

import pytest
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

NAVY = RGBColor(0x12, 0x2B, 0x45)


def _png(color=(200, 30, 30)) -> bytes:
    buf = io.BytesIO()
    Image.new("RGB", (40, 40), color).save(buf, "PNG")
    return buf.getvalue()


def build_pptx(path: Path) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    prs.core_properties.title = "Fixture deck"
    title_only = prs.slide_layouts[5]
    blank = prs.slide_layouts[6]

    # 1: title, typed bullets with a sub-level, notes, solid background
    s = prs.slides.add_slide(title_only)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = NAVY
    s.shapes.title.text = "Opening slide"
    box = s.shapes.add_textbox(Inches(1), Inches(2), Inches(10), Inches(3))
    tf = box.text_frame
    tf.text = "• First point"
    p = tf.add_paragraph()
    p.text = "– Sub point"
    p.level = 1
    p = tf.add_paragraph()
    p.text = "Plain point | with a pipe"
    s.notes_slide.notes_text_frame.text = "Say hello.\nThen the second line."

    # 2: clustered column chart with two series
    s = prs.slides.add_slide(title_only)
    s.shapes.title.text = "Revenue chart"
    data = CategoryChartData()
    data.categories = ["Q1", "Q2", "Q3"]
    data.add_series("Revenue", (10, 20.5, 31))
    data.add_series("Costs", (8, 9, 12))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1.5), Inches(8), Inches(5), data)
    gf.chart.has_title = True
    gf.chart.chart_title.text_frame.text = "Money by quarter"

    # 3: table with a merged header cell
    s = prs.slides.add_slide(title_only)
    s.shapes.title.text = "Price table"
    tbl = s.shapes.add_table(3, 3, Inches(1), Inches(1.5), Inches(9), Inches(2)).table
    values = [["Plan", "Price", "Seats"], ["Basic", "$5", "1"], ["Team", "$20", "5"]]
    for r, row in enumerate(values):
        for c, v in enumerate(row):
            tbl.cell(r, c).text = v

    # 4: picture with alt text, a group with text, a hyperlink
    s = prs.slides.add_slide(title_only)
    s.shapes.title.text = "Product"
    pic = s.shapes.add_picture(io.BytesIO(_png()), Inches(9), Inches(2), Inches(2), Inches(2))
    pic.name = "Hero image"
    pic._element.nvPicPr.cNvPr.set("descr", "A red square")
    group = s.shapes.add_group_shape()
    card = group.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2), Inches(3), Inches(2))
    card.text_frame.text = "Inside a group"
    link_box = s.shapes.add_textbox(Inches(1), Inches(5), Inches(6), Inches(1))
    run = link_box.text_frame.paragraphs[0].add_run()
    run.text = "Docs"
    run.font.size = Pt(18)
    run.hyperlink.address = "https://example.com/docs"

    # 5: scatter chart
    s = prs.slides.add_slide(title_only)
    s.shapes.title.text = "Scatter"
    xy = XyChartData()
    ser = xy.add_series("Points")
    ser.add_data_point(1, 2)
    ser.add_data_point(3, 4.5)
    s.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER, Inches(1), Inches(1.5), Inches(8), Inches(5), xy)

    # 6: no title placeholder: the first text line becomes the label
    s = prs.slides.add_slide(blank)
    tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    tb.text_frame.text = "Untitled layout text"

    # 7: hidden slide with notes
    s = prs.slides.add_slide(title_only)
    s.shapes.title.text = "Backup numbers"
    s._element.set("show", "0")
    s.notes_slide.notes_text_frame.text = "Hidden on purpose."

    prs.save(str(path))
    return path


def build_pdf(path: Path, pages) -> Path:
    """Write a minimal valid PDF, one 960x540 pt page per list of text lines."""
    objects = []

    def add(body: bytes) -> int:
        objects.append(body)
        return len(objects)

    font = add(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")
    page_ids = []
    pages_id = len(objects) + 1 + 2 * len(pages)  # reserved: after every page + content pair
    for lines in pages:
        ops = ["BT", "/F1 28 Tf", "60 470 Td", "34 TL"]
        for line in lines:
            safe = line.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
            ops.append("(%s) Tj T*" % safe)
        ops.append("ET")
        stream = "\n".join(ops).encode("latin-1")
        content = add(b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream")
        page_ids.append(add(
            b"<< /Type /Page /Parent %d 0 R /MediaBox [0 0 960 540] /Contents %d 0 R "
            b"/Resources << /Font << /F1 %d 0 R >> >> >>" % (pages_id, content, font)))
    kids = b" ".join(b"%d 0 R" % i for i in page_ids)
    assert add(b"<< /Type /Pages /Kids [%s] /Count %d >>" % (kids, len(page_ids))) == pages_id
    info = add(b"<< /Title (PDF fixture) >>")
    catalog = add(b"<< /Type /Catalog /Pages %d 0 R >>" % pages_id)

    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for i, body in enumerate(objects, start=1):
        offsets.append(len(out))
        out += b"%d 0 obj\n" % i + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    for off in offsets:
        out += b"%010d 00000 n \n" % off
    out += b"trailer\n<< /Size %d /Root %d 0 R /Info %d 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (
        len(objects) + 1, catalog, info, xref)
    path.write_bytes(bytes(out))
    return path


PDF_PAGES = [
    ["Quarterly review", "- Revenue grew", "and costs held flat", "See https://example.com/q3."],
    ["Next steps", "Hire two people"],
]


@pytest.fixture(scope="session")
def deck_pptx(tmp_path_factory) -> Path:
    return build_pptx(tmp_path_factory.mktemp("pptx") / "fixture.pptx")


@pytest.fixture(scope="session")
def deck_pdf(tmp_path_factory) -> Path:
    return build_pdf(tmp_path_factory.mktemp("pdf") / "slides.pdf", PDF_PAGES)


@pytest.fixture
def no_office(monkeypatch):
    """Pretend neither PowerPoint nor LibreOffice is installed."""
    from amagi_deck import render

    monkeypatch.setattr(render, "has_powerpoint", lambda: False)
    monkeypatch.setattr(render, "find_soffice", lambda: None)
