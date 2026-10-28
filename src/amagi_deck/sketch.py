"""Approximate slide drawing with Pillow, for machines with no PowerPoint or LibreOffice.

It places solid-filled shapes, pictures, text, tables and simple bar/line charts where the
deck puts them. Theme colors, gradients, effects and exact fonts are not resolved, so every
image gets a small "sketch render" label in the corner.
"""

from __future__ import annotations

import io
from functools import lru_cache
from pathlib import Path
from typing import Optional, Tuple

from PIL import Image, ImageDraw, ImageFont
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_FILL
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN

from .pptx_extract import parse_chart_xml

EMU_PER_PT = 12700
FONT_FILES = ("arial.ttf", "Arial.ttf", "DejaVuSans.ttf", "LiberationSans-Regular.ttf",
              "Helvetica.ttc", "/System/Library/Fonts/Helvetica.ttc",
              "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf")
BOLD_FILES = ("arialbd.ttf", "Arial Bold.ttf", "DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")
PALETTE = [(38, 99, 186), (232, 126, 4), (46, 158, 92), (196, 58, 58), (128, 90, 180), (120, 120, 120)]
LABEL = "amagi-deck sketch render (approximate)"


@lru_cache(maxsize=64)
def font(size: int, bold: bool = False):
    size = max(6, int(size))
    for name in (BOLD_FILES if bold else ()) + FONT_FILES:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    try:
        return ImageFont.load_default(size=size)
    except TypeError:  # Pillow < 10.1
        return ImageFont.load_default()


def _rgb(color_format) -> Optional[Tuple[int, int, int]]:
    try:
        if color_format.type == MSO_COLOR_TYPE.RGB:
            rgb = color_format.rgb
            return (rgb[0], rgb[1], rgb[2])
    except Exception:
        pass
    return None


def _solid_fill(fill) -> Optional[Tuple[int, int, int]]:
    try:
        if fill.type == MSO_FILL.SOLID:
            return _rgb(fill.fore_color)
    except Exception:
        pass
    return None


def _dark(rgb) -> bool:
    return rgb is not None and (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) < 128


def _wrap(draw, text: str, fnt, max_w: int):
    words = text.split()
    lines, line = [], ""
    for word in words:
        trial = (line + " " + word).strip()
        if not line or draw.textlength(trial, font=fnt) <= max_w:
            line = trial
        else:
            lines.append(line)
            line = word
    if line:
        lines.append(line)
    return lines


class Sketcher:
    def __init__(self, prs, width: int):
        self.prs = prs
        self.width = width
        self.scale = width / float(prs.slide_width)
        self.height = max(1, round(prs.slide_height * self.scale))

    def px(self, emu) -> int:
        return int(round((emu or 0) * self.scale))

    def render(self, slide, out: Path) -> Path:
        bg = None
        try:
            bg = _solid_fill(slide.background.fill)
        except Exception:
            bg = None
        img = Image.new("RGB", (self.width, self.height), bg or (255, 255, 255))
        draw = ImageDraw.Draw(img)
        self._shapes(img, draw, slide.shapes, (0, 0, 1.0, 1.0), bg)
        lab = font(max(10, self.height // 60))
        tw = draw.textlength(LABEL, font=lab)
        pad = max(4, self.height // 150)
        y = self.height - pad - lab.size - pad
        draw.rectangle([self.width - tw - 3 * pad, y - pad, self.width, self.height], fill=(240, 240, 240))
        draw.text((self.width - tw - 2 * pad, y), LABEL, font=lab, fill=(90, 90, 90))
        img.save(out, "PNG", optimize=True)
        return out

    def _box(self, shape, xf):
        ox, oy, sx, sy = xf
        left = ox + (shape.left or 0) * sx
        top = oy + (shape.top or 0) * sy
        w = (shape.width or 0) * sx
        h = (shape.height or 0) * sy
        return [self.px(left), self.px(top), self.px(left + w), self.px(top + h)]

    def _shapes(self, img, draw, shapes, xf, bg):
        for shape in shapes:
            try:
                self._shape(img, draw, shape, xf, bg)
            except Exception:
                # One odd shape should not cost the whole slide.
                continue

    def _shape(self, img, draw, shape, xf, bg):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            xfrm = shape._element.grpSpPr.find(
                "{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
            ox, oy, sx, sy = xf
            if xfrm is not None and xfrm.find("{*}chExt") is not None:
                off, ext = xfrm.find("{*}off"), xfrm.find("{*}ext")
                ch_off, ch_ext = xfrm.find("{*}chOff"), xfrm.find("{*}chExt")
                cx = float(ch_ext.get("cx")) or 1.0
                cy = float(ch_ext.get("cy")) or 1.0
                gsx = float(ext.get("cx")) / cx * sx
                gsy = float(ext.get("cy")) / cy * sy
                gox = ox + float(off.get("x")) * sx - float(ch_off.get("x")) * gsx
                goy = oy + float(off.get("y")) * sy - float(ch_off.get("y")) * gsy
                xf = (gox, goy, gsx, gsy)
            self._shapes(img, draw, shape.shapes, xf, bg)
            return
        box = self._box(shape, xf)
        if box[2] - box[0] < 1 or box[3] - box[1] < 1:
            return
        if getattr(shape, "has_chart", False) and shape.has_chart:
            self._chart(draw, shape, box)
            return
        if getattr(shape, "has_table", False) and shape.has_table:
            self._table(draw, shape, box)
            return
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or hasattr(shape, "image"):
            try:
                pic = Image.open(io.BytesIO(shape.image.blob)).convert("RGBA")
                pic = pic.resize((box[2] - box[0], box[3] - box[1]))
                img.paste(pic, (box[0], box[1]), pic)
                return
            except Exception:
                pass
        fill = None
        if hasattr(shape, "fill"):
            fill = _solid_fill(shape.fill)
        if fill is not None:
            oval = False
            try:
                oval = shape.auto_shape_type == MSO_AUTO_SHAPE_TYPE.OVAL
            except Exception:
                pass
            (draw.ellipse if oval else draw.rectangle)(box, fill=fill)
        if shape.has_text_frame:
            self._text(draw, shape, box, fill if fill is not None else bg)

    def _text(self, draw, shape, box, under):
        ptype = None
        if shape.is_placeholder:
            try:
                ptype = shape.placeholder_format.type
            except ValueError:
                ptype = None
        is_title = ptype in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE)
        default_pt = 40 if is_title else 18
        default_color = (255, 255, 255) if _dark(under) else (34, 34, 34)
        inset = self.px(91440)
        x0, y, x1 = box[0] + inset, box[1] + inset, box[2] - inset
        max_w = max(10, x1 - x0)
        for para in shape.text_frame.paragraphs:
            text = " ".join(para.text.replace("\v", " ").split())
            if not text:
                y += self.px(default_pt * EMU_PER_PT) // 2
                continue
            size_pt, bold, color = default_pt, is_title, default_color
            for run in para.runs:
                if run.font.size is not None:
                    size_pt = run.font.size.pt
                if run.font.bold is not None:
                    bold = run.font.bold
                rgb = _rgb(run.font.color) if run.font.color is not None else None
                if rgb is not None:
                    color = rgb
                break
            fnt = font(self.px(size_pt * EMU_PER_PT), bold)
            indent = self.px(para.level * 457200)
            lines = _wrap(draw, text, fnt, max_w - indent)
            align = para.alignment
            if align is None and ptype == PP_PLACEHOLDER.CENTER_TITLE:
                align = PP_ALIGN.CENTER
            for line in lines:
                w = draw.textlength(line, font=fnt)
                if align == PP_ALIGN.CENTER:
                    x = x0 + (max_w - w) / 2
                elif align == PP_ALIGN.RIGHT:
                    x = x1 - w
                else:
                    x = x0 + indent
                draw.text((x, y), line, font=fnt, fill=color)
                y += int(fnt.size * 1.2)

    def _table(self, draw, shape, box):
        table = shape.table
        rows = list(table.rows)
        cols = list(table.columns)
        sx = (box[2] - box[0]) / float(sum(c.width for c in cols) or 1)
        sy = (box[3] - box[1]) / float(sum(r.height for r in rows) or 1)
        fnt = font(self.px(12 * EMU_PER_PT))
        y = box[1]
        for r, row in enumerate(rows):
            x = box[0]
            rh = row.height * sy
            for c, col in enumerate(cols):
                cw = col.width * sx
                cell = table.cell(r, c)
                fill = _solid_fill(cell.fill) or ((68, 84, 106) if r == 0 else (248, 248, 248))
                draw.rectangle([x, y, x + cw, y + rh], fill=fill, outline=(200, 200, 200))
                if not cell.is_spanned:
                    text = " ".join(cell.text.split())
                    lines = _wrap(draw, text, fnt, int(cw) - 8)
                    ty = y + 4
                    for line in lines:
                        draw.text((x + 4, ty), line, font=fnt,
                                  fill=(255, 255, 255) if _dark(fill) else (34, 34, 34))
                        ty += int(fnt.size * 1.2)
                x += cw
            y += rh

    def _chart(self, draw, shape, box):
        chart = parse_chart_xml(shape.chart._chartSpace)
        draw.rectangle(box, fill=(252, 252, 252), outline=(210, 210, 210))
        tf = font(self.px(14 * EMU_PER_PT), True)
        lf = font(self.px(10 * EMU_PER_PT))
        title = chart.title or chart.kind
        draw.text((box[0] + 8, box[1] + 6), title, font=tf, fill=(34, 34, 34))
        series = [s for s in chart.series if any(v is not None for v in s.values)]
        if not series:
            return
        top = box[1] + 12 + tf.size
        bottom = box[3] - 10 - 2 * lf.size
        left, right = box[0] + 12, box[2] - 12
        values = [v for s in series for v in s.values if v is not None]
        vmax = max(max(values), 0.0)
        vmin = min(min(values), 0.0)
        span = (vmax - vmin) or 1.0
        n = max(len(s.values) for s in series)

        def ypos(v):
            return bottom - (v - vmin) / span * (bottom - top)

        draw.line([left, ypos(0), right, ypos(0)], fill=(160, 160, 160))
        slot = (right - left) / float(n or 1)
        kind = chart.kind.split(" ")[0]
        if kind in ("line", "scatter", "area", "radar"):
            for si, s in enumerate(series):
                pts = [(left + slot * (i + 0.5), ypos(v)) for i, v in enumerate(s.values) if v is not None]
                if len(pts) > 1:
                    draw.line(pts, fill=PALETTE[si % len(PALETTE)], width=max(2, self.height // 300))
        else:
            bw = slot * 0.8 / len(series)
            for si, s in enumerate(series):
                for i, v in enumerate(s.values):
                    if v is None:
                        continue
                    x = left + slot * i + slot * 0.1 + bw * si
                    y0, y1 = sorted((ypos(0), ypos(v)))
                    draw.rectangle([x, y0, x + bw - 1, y1], fill=PALETTE[si % len(PALETTE)])
        for i, cat in enumerate(chart.categories[:n]):
            w = draw.textlength(cat, font=lf)
            draw.text((left + slot * (i + 0.5) - w / 2, bottom + 4), cat, font=lf, fill=(80, 80, 80))
        lx = left
        for si, s in enumerate(series):
            draw.rectangle([lx, box[3] - lf.size - 6, lx + lf.size, box[3] - 6], fill=PALETTE[si % len(PALETTE)])
            draw.text((lx + lf.size + 4, box[3] - lf.size - 7), s.name, font=lf, fill=(60, 60, 60))
            lx += lf.size + 12 + draw.textlength(s.name, font=lf)


def render_sketch(path: Path, targets, width: int):
    from pptx import Presentation

    prs = Presentation(str(path))
    sketch = Sketcher(prs, width)
    done = {}
    for number, slide in enumerate(prs.slides, start=1):
        if number in targets:
            done[number] = sketch.render(slide, targets[number])
    return done
