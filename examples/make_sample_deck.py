"""Build examples/pitch-deck.pptx, a made-up seed pitch for a made-up company (Tidewell).

Every name and number in it is fictional. Run: python examples/make_sample_deck.py
"""

from __future__ import annotations

import io
import sys
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

NAVY = RGBColor(0x12, 0x2B, 0x45)
TEAL = RGBColor(0x1F, 0x9E, 0x9A)
SAND = RGBColor(0xF4, 0xEF, 0xE6)
INK = RGBColor(0x22, 0x22, 0x22)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


def sensor_png() -> bytes:
    img = Image.new("RGB", (600, 600), (244, 239, 230))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([120, 140, 480, 460], radius=60, fill=(18, 43, 69))
    d.ellipse([240, 220, 360, 340], fill=(31, 158, 154))
    d.rounded_rectangle([200, 380, 400, 410], radius=12, fill=(244, 239, 230))
    for i in range(3):
        d.arc([300 - 90 - 40 * i, 60 - 40 * i, 300 + 90 + 40 * i, 220 + 40 * i], 230, 310,
              fill=(31, 158, 154), width=10)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def band(slide, prs, color=NAVY, height=Inches(0.18)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def title(slide, text, size=36, color=NAVY):
    ph = slide.shapes.title
    ph.left, ph.top, ph.width, ph.height = Inches(0.6), Inches(0.45), Inches(12.1), Inches(1.0)
    ph.text = text
    para = ph.text_frame.paragraphs[0]
    para.alignment = PP_ALIGN.LEFT
    run = para.runs[0]
    run.font.size = Pt(size)
    run.font.bold = True
    run.font.color.rgb = color
    return ph


def bullets(slide, items, left, top, width, height, size=22):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        level, text = item if isinstance(item, tuple) else (0, item)
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.level = level
        run = para.add_run()
        run.text = ("• " if level == 0 else "– ") + text
        run.font.size = Pt(size if level == 0 else size - 4)
        run.font.color.rgb = INK
        para.space_after = Pt(10)
    return box


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def build(path: Path) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    prs.core_properties.title = "Tidewell seed deck"
    prs.core_properties.author = "Tidewell (fictional)"
    title_only = prs.slide_layouts[5]

    # 1. Cover
    s = prs.slides.add_slide(title_only)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = NAVY
    t = title(s, "Tidewell", size=60, color=WHITE)
    t.top = Inches(2.4)
    sub = s.shapes.add_textbox(Inches(0.6), Inches(3.6), Inches(12), Inches(1.2))
    p = sub.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "Leak sensors and a shut-off valve for landlords with 2 to 40 units"
    r.font.size, r.font.color.rgb = Pt(26), RGBColor(0xCF, 0xE8, 0xE6)
    p2 = sub.text_frame.add_paragraph()
    r2 = p2.add_run()
    r2.text = "Seed round, 2026"
    r2.font.size, r2.font.color.rgb = Pt(18), RGBColor(0x9F, 0xB6, 0xC8)
    notes(s, "Open with the burst pipe story: one unit, one weekend, $12,400 in repairs.\n"
             "Keep this slide up for under 20 seconds.")

    # 2. Problem
    s = prs.slides.add_slide(title_only)
    band(s, prs)
    title(s, "Small landlords find leaks days late")
    bullets(s, [
        "A slow leak under a sink runs for 4 days on average before a tenant reports it",
        (1, "Water damage was the largest repair cost in 9 of our 12 pilot portfolios"),
        "Smart-home leak kits are built for one house, not 12 units across 3 buildings",
        "Property managers want one alert, with the unit number in it",
    ], Inches(0.6), Inches(1.7), Inches(12), Inches(4.8))
    notes(s, "The 4-day figure is from our 31 pilot interviews, not a market study. Say so if asked.")

    # 3. Product, with a picture
    s = prs.slides.add_slide(title_only)
    band(s, prs)
    title(s, "One sensor per wet area, one valve per building")
    bullets(s, [
        "Battery sensor: 3-year life, sticks under sinks and water heaters",
        "Valve closes the main within 30 seconds of a confirmed leak",
        "Alerts by text and email, with the unit and the room",
    ], Inches(0.6), Inches(1.7), Inches(7.2), Inches(4.5))
    pic = s.shapes.add_picture(io.BytesIO(sensor_png()), Inches(8.3), Inches(1.6), Inches(4.4), Inches(4.4))
    pic.name = "Sensor photo"
    pic._element.nvPicPr.cNvPr.set("descr", "Tidewell sensor: a navy puck with a teal status light")

    # 4. Traction chart
    s = prs.slides.add_slide(title_only)
    band(s, prs)
    title(s, "Traction: 9 months of pilot installs")
    data = CategoryChartData()
    data.categories = ["Q1 2026", "Q2 2026", "Q3 2026"]
    data.add_series("Sensors installed", (140, 410, 960))
    data.add_series("Paying buildings", (6, 19, 41))
    frame = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.6), Inches(1.6),
                               Inches(8.2), Inches(5.3), data)
    chart = frame.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "Pilot installs by quarter"
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    for i, color in enumerate((NAVY, TEAL)):
        fill = chart.plots[0].series[i].format.fill
        fill.solid()
        fill.fore_color.rgb = color
    bullets(s, ["41 paying buildings", "3 leaks caught before tenant reports", "0 churned buildings"],
            Inches(9.2), Inches(2.0), Inches(3.8), Inches(4), size=20)
    notes(s, "Q3 is through September 30 only.\nThe 3 caught leaks are the strongest proof point: "
             "tell the Elm Street story.")

    # 5. Revenue line chart
    s = prs.slides.add_slide(title_only)
    band(s, prs)
    title(s, "Monthly recurring revenue")
    data = CategoryChartData()
    data.categories = ["Apr", "May", "Jun", "Jul", "Aug", "Sep"]
    data.add_series("MRR ($)", (1200, 1850, 2600, 3900, 5100, 6650))
    frame = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.6), Inches(1.6),
                               Inches(12), Inches(5.3), data)
    frame.chart.has_legend = False
    frame.chart.plots[0].series[0].format.line.color.rgb = TEAL
    frame.chart.plots[0].series[0].format.line.width = Pt(3.5)

    # 6. Pricing table
    s = prs.slides.add_slide(title_only)
    band(s, prs)
    title(s, "Pricing")
    rows = [
        ("Plan", "Per unit / month", "Includes", "Buildings"),
        ("Starter", "$4", "Sensors, alerts", "1 to 3"),
        ("Portfolio", "$3", "Sensors, alerts, valve", "4 to 15"),
        ("Manager", "$2.50", "Everything, plus API", "16+"),
    ]
    tbl = s.shapes.add_table(len(rows), 4, Inches(0.6), Inches(1.8), Inches(12), Inches(3.2)).table
    for r, row in enumerate(rows):
        for c, value in enumerate(row):
            cell = tbl.cell(r, c)
            cell.text = value
            para = cell.text_frame.paragraphs[0]
            para.runs[0].font.size = Pt(20)
    notes(s, "Hardware is sold at cost. The margin is in the monthly fee.")

    # 7. Team, as a group of cards
    s = prs.slides.add_slide(title_only)
    band(s, prs)
    title(s, "Team")
    group = s.shapes.add_group_shape()
    people = [("Dana Okafor", "CEO. 8 years running a 60-unit portfolio"),
              ("Luis Ferreira", "CTO. Built firmware for a smart-meter maker"),
              ("Priya Raman", "Ops. Led installs for a regional plumbing firm")]
    for i, (name, role) in enumerate(people):
        left = Inches(0.6 + i * 4.15)
        card = group.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.9), Inches(3.9), Inches(3.6))
        card.fill.solid()
        card.fill.fore_color.rgb = SAND
        card.line.fill.background()
        tf = card.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = name
        r.font.size, r.font.bold, r.font.color.rgb = Pt(24), True, NAVY
        p2 = tf.add_paragraph()
        r2 = p2.add_run()
        r2.text = role
        r2.font.size, r2.font.color.rgb = Pt(16), INK

    # 8. The ask, with a link
    s = prs.slides.add_slide(title_only)
    band(s, prs, TEAL)
    title(s, "The ask: $1.2M seed")
    box = bullets(s, [
        "18 months of runway to 400 paying buildings",
        "Hire 2 installers and 1 firmware engineer",
        "UL certification for the valve",
    ], Inches(0.6), Inches(1.7), Inches(12), Inches(3.6))
    p = box.text_frame.add_paragraph()
    r = p.add_run()
    r.text = "Data room: example.com/tidewell"
    r.font.size, r.font.color.rgb = Pt(18), TEAL
    r.hyperlink.address = "https://example.com/tidewell"
    notes(s, "Stop talking after the number. Let them ask the first question.")

    # 9. Hidden appendix
    s = prs.slides.add_slide(title_only)
    band(s, prs)
    title(s, "Appendix: unit economics")
    rows = [("Metric", "Value"), ("Hardware cost per building", "$310"),
            ("Install time", "2.5 hours"), ("Gross margin (subscription)", "78%")]
    tbl = s.shapes.add_table(len(rows), 2, Inches(0.6), Inches(1.8), Inches(8), Inches(2.6)).table
    for r_, row in enumerate(rows):
        for c, value in enumerate(row):
            tbl.cell(r_, c).text = value
    s._element.set("show", "0")
    notes(s, "Hidden. Show only if someone asks about margins.")

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return path


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("pitch-deck.pptx")
    print(build(out))
