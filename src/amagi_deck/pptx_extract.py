"""Pull text, notes, tables, chart data, pictures and links out of a .pptx."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, List, Optional

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.shapes.picture import Picture as PptxPicture
from pptx.util import Emu

from .model import Chart, Deck, Picture, Series, Slide, Table, TextBlock

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
}
DIAGRAM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
TITLE_TYPES = {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}
SKIP_PLACEHOLDERS = {PP_PLACEHOLDER.SLIDE_NUMBER, PP_PLACEHOLDER.DATE, PP_PLACEHOLDER.FOOTER}


def clean(text: str) -> str:
    """Collapse PowerPoint soft line breaks and stray whitespace."""
    return " ".join(text.replace("\v", " ").replace("\x0b", " ").split())


def iter_shapes(shapes) -> Iterable:
    """Yield shapes in reading order (top to bottom, then left to right), entering groups."""

    def key(shape):
        return (shape.top or 0, shape.left or 0)

    for shape in sorted(shapes, key=key):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(shape.shapes)
        else:
            yield shape


def placeholder_type(shape):
    if not shape.is_placeholder:
        return None
    try:
        return shape.placeholder_format.type
    except ValueError:
        return None


def _num(text: Optional[str]) -> Optional[float]:
    if text is None:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _points(ref) -> List[Optional[str]]:
    """Read a c:strCache / c:numCache / c:multiLvlStrCache into a list indexed by c:pt/@idx."""
    if ref is None:
        return []
    cache = None
    for name in ("strCache", "numCache", "multiLvlStrCache", "strLit", "numLit"):
        cache = ref.find(".//c:%s" % name, NS)
        if cache is not None:
            break
    if cache is None:
        return []
    if cache.tag.endswith("multiLvlStrCache"):
        # Use the innermost level (the one closest to the axis).
        lvl = cache.find("c:lvl", NS)
        if lvl is None:
            return []
        count_el = cache.find("c:ptCount", NS)
        pts = lvl.findall("c:pt", NS)
    else:
        count_el = cache.find("c:ptCount", NS)
        pts = cache.findall("c:pt", NS)
    count = int(count_el.get("val")) if count_el is not None else len(pts)
    values: List[Optional[str]] = [None] * count
    for pt in pts:
        idx = int(pt.get("idx", "0"))
        v = pt.find("c:v", NS)
        if idx >= len(values):
            values.extend([None] * (idx + 1 - len(values)))
        values[idx] = v.text if v is not None else None
    return values


def _chart_kind(plot) -> str:
    kind = etree.QName(plot).localname.replace("Chart", "")
    details = []
    bar_dir = plot.find("c:barDir", NS)
    if bar_dir is not None:
        details.append("column" if bar_dir.get("val") == "col" else "horizontal")
    grouping = plot.find("c:grouping", NS)
    if grouping is not None and grouping.get("val") not in (None, "standard"):
        details.append(grouping.get("val"))
    return kind + (" (%s)" % ", ".join(details) if details else "")


def parse_chart_xml(chart_space) -> Chart:
    """Turn a c:chartSpace element into a Chart with categories and series values."""
    title_el = chart_space.find("c:chart/c:title", NS)
    title = ""
    if title_el is not None:
        title = clean(" ".join(t.text or "" for t in title_el.iterfind(".//a:t", NS)))
        if not title:
            title = clean(" ".join(p or "" for p in _points(title_el.find(".//c:strRef", NS))))
    plot_area = chart_space.find("c:chart/c:plotArea", NS)
    kinds: List[str] = []
    categories: List[str] = []
    series: List[Series] = []
    if plot_area is not None:
        for plot in plot_area:
            if not etree.QName(plot).localname.endswith("Chart"):
                continue
            kinds.append(_chart_kind(plot))
            for ser in plot.findall("c:ser", NS):
                name_el = ser.find("c:tx", NS)
                name = ""
                if name_el is not None:
                    pts = _points(name_el)
                    if pts:
                        name = clean(pts[0] or "")
                    else:
                        v = name_el.find("c:v", NS)
                        name = clean(v.text or "") if v is not None else ""
                cats = _points(ser.find("c:cat", NS))
                if cats and not categories:
                    categories = [clean(c or "") for c in cats]
                x_raw = _points(ser.find("c:xVal", NS))
                y_el = ser.find("c:val", NS)
                if y_el is None:
                    y_el = ser.find("c:yVal", NS)
                values = [_num(v) for v in _points(y_el)]
                x_values = [_num(v) for v in x_raw]
                if x_raw and all(v is None for v in x_values):
                    # Text x values in a scatter chart act as category labels.
                    if not categories:
                        categories = [clean(v or "") for v in x_raw]
                    x_values = []
                series.append(Series(name=name or "Series %d" % (len(series) + 1),
                                     values=values, x_values=x_values))
    return Chart(kind=" + ".join(kinds) or "chart", title=title, categories=categories, series=series)


def _smartart_text(shape, slide_part) -> List[str]:
    rel_ids = shape._element.find(".//dgm:relIds", NS)
    if rel_ids is None:
        return []
    rid = rel_ids.get("{%s}dm" % NS["r"])
    try:
        part = slide_part.related_part(rid)
    except KeyError:
        return []
    root = etree.fromstring(part.blob)
    lines = []
    for pt in root.iterfind(".//dgm:pt", NS):
        if pt.get("type") not in (None, "node"):
            continue
        text = clean("".join(t.text or "" for t in pt.iterfind(".//a:t", NS)))
        if text:
            lines.append(text)
    return lines


def _alt_text(shape) -> str:
    for el in shape._element.iter():
        if etree.QName(el).localname == "cNvPr":
            return clean(el.get("descr") or el.get("title") or "")
    return ""


def _links(shape) -> List[str]:
    out = []
    try:
        address = shape.click_action.hyperlink.address
        if address:
            out.append(address)
    except Exception:
        pass
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                address = run.hyperlink.address
                if address:
                    out.append(address)
    return out


def _text_lines(text_frame) -> List[tuple]:
    lines = []
    for para in text_frame.paragraphs:
        text = clean(para.text)
        if text:
            lines.append((para.level, text))
    return lines


def _table(shape) -> Table:
    rows = []
    for row in shape.table.rows:
        cells = []
        for cell in row.cells:
            cells.append("" if cell.is_spanned else clean(cell.text_frame.text if cell.text_frame else ""))
        rows.append(cells)
    return Table(rows=rows)


def parse_slide(slide, number: int, with_notes: bool = True) -> Slide:
    out = Slide(number=number)
    out.hidden = slide._element.get("show") == "0"
    try:
        out.layout = slide.slide_layout.name or ""
    except Exception:
        out.layout = ""

    for shape in iter_shapes(slide.shapes):
        ptype = placeholder_type(shape)
        if ptype in SKIP_PLACEHOLDERS:
            continue
        for link in _links(shape):
            if link not in out.links:
                out.links.append(link)

        if getattr(shape, "has_chart", False) and shape.has_chart:
            out.charts.append(parse_chart_xml(shape.chart._chartSpace))
            continue
        if getattr(shape, "has_table", False) and shape.has_table:
            out.tables.append(_table(shape))
            continue
        graphic = shape._element.find(".//a:graphicData", NS)
        if graphic is not None and graphic.get("uri") == DIAGRAM_URI:
            lines = _smartart_text(shape, slide.part)
            if lines:
                out.text.append(TextBlock(shape="SmartArt", lines=[(0, t) for t in lines]))
            continue
        if isinstance(shape, PptxPicture):
            content_type = ""
            try:
                content_type = shape.image.content_type
            except Exception:
                pass
            out.pictures.append(Picture(name=shape.name, alt_text=_alt_text(shape),
                                        content_type=content_type))
            continue
        if not shape.has_text_frame:
            continue
        lines = _text_lines(shape.text_frame)
        if not lines:
            continue
        if ptype in TITLE_TYPES and not out.title:
            out.title = " ".join(t for _, t in lines)
            continue
        out.text.append(TextBlock(shape=shape.name, lines=lines))

    if not out.title:
        # No title placeholder: use the first line of the top-most text as a label only.
        for block in out.text:
            if block.lines:
                out.title = block.lines[0][1][:80]
                break

    if with_notes and slide.has_notes_slide:
        frame = slide.notes_slide.notes_text_frame
        if frame is not None:
            paras = [p.text.replace("\v", "\n").strip() for p in frame.paragraphs]
            out.notes = "\n".join(p for p in paras if p)
    return out


def extract_pptx(path: Path, with_notes: bool = True) -> Deck:
    prs = Presentation(str(path))
    slides = [parse_slide(s, i, with_notes) for i, s in enumerate(prs.slides, start=1)]
    deck = Deck(
        source=path.name,
        kind="pptx",
        slides=slides,
        width_in=round(Emu(prs.slide_width or 0).inches, 2),
        height_in=round(Emu(prs.slide_height or 0).inches, 2),
    )
    deck.title = clean(prs.core_properties.title or "") or (slides[0].title if slides else "")
    return deck
