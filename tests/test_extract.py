from lxml import etree

from amagi_deck.pdf_extract import extract_pdf, join_lines
from amagi_deck.pptx_extract import extract_pptx, parse_chart_xml, strip_bullet


def test_deck_level_fields(deck_pptx):
    deck = extract_pptx(deck_pptx)
    assert deck.title == "Fixture deck"
    assert deck.kind == "pptx"
    assert (deck.width_in, deck.height_in) == (13.33, 7.5)
    assert [s.number for s in deck.slides] == [1, 2, 3, 4, 5, 6, 7]


def test_titles_bullets_levels_and_notes(deck_pptx):
    s1 = extract_pptx(deck_pptx).slides[0]
    assert s1.title == "Opening slide"
    lines = [line for block in s1.text for line in block.lines]
    assert lines == [(0, "First point"), (1, "Sub point"), (0, "Plain point | with a pipe")]
    assert s1.notes == "Say hello.\nThen the second line."


def test_notes_can_be_skipped(deck_pptx):
    deck = extract_pptx(deck_pptx, with_notes=False)
    assert all(s.notes == "" for s in deck.slides)


def test_column_chart_data(deck_pptx):
    chart = extract_pptx(deck_pptx).slides[1].charts[0]
    assert chart.title == "Money by quarter"
    assert chart.kind == "column (clustered)"
    assert chart.categories == ["Q1", "Q2", "Q3"]
    assert [(s.name, s.values) for s in chart.series] == [
        ("Revenue", [10.0, 20.5, 31.0]), ("Costs", [8.0, 9.0, 12.0])]


def test_scatter_chart_keeps_x_values(deck_pptx):
    chart = extract_pptx(deck_pptx).slides[4].charts[0]
    assert chart.kind == "scatter"
    assert chart.series[0].x_values == [1.0, 3.0]
    assert chart.series[0].values == [2.0, 4.5]


def test_table_rows(deck_pptx):
    table = extract_pptx(deck_pptx).slides[2].tables[0]
    assert table.rows[0] == ["Plan", "Price", "Seats"]
    assert table.rows[2] == ["Team", "$20", "5"]


def test_picture_group_text_and_links(deck_pptx):
    s4 = extract_pptx(deck_pptx).slides[3]
    assert [(p.name, p.alt_text, p.content_type) for p in s4.pictures] == [
        ("Hero image", "A red square", "image/png")]
    texts = [t for block in s4.text for _, t in block.lines]
    assert "Inside a group" in texts
    assert s4.links == ["https://example.com/docs"]


def test_untitled_slide_uses_first_line(deck_pptx):
    assert extract_pptx(deck_pptx).slides[5].title == "Untitled layout text"


def test_hidden_slide(deck_pptx):
    slides = extract_pptx(deck_pptx).slides
    assert [s.hidden for s in slides] == [False] * 6 + [True]
    assert slides[6].notes == "Hidden on purpose."


def test_strip_bullet():
    assert strip_bullet("• Point") == "Point"
    assert strip_bullet("- Point") == "Point"
    assert strip_bullet("-5% churn") == "-5% churn"
    assert strip_bullet("–\tSub") == "Sub"
    assert strip_bullet("Plain") == "Plain"


def test_chart_xml_with_missing_points_and_horizontal_bars():
    xml = b"""<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea>
      <c:barChart><c:barDir val="bar"/><c:grouping val="stacked"/>
        <c:ser><c:tx><c:v>Direct</c:v></c:tx>
          <c:cat><c:strRef><c:strCache><c:ptCount val="3"/>
            <c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt>
          </c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="3"/>
            <c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>3.25</c:v></c:pt>
          </c:numCache></c:numRef></c:val></c:ser>
      </c:barChart></c:plotArea></c:chart></c:chartSpace>"""
    chart = parse_chart_xml(etree.fromstring(xml))
    assert chart.kind == "bar (stacked)"
    assert chart.categories == ["A", "", "C"]
    assert chart.series[0].name == "Direct"
    assert chart.series[0].values == [1.0, None, 3.25]


def test_pdf_pages(deck_pdf):
    deck = extract_pdf(deck_pdf)
    assert deck.kind == "pdf"
    assert deck.title == "PDF fixture"
    assert (deck.width_in, deck.height_in) == (13.33, 7.5)
    s1, s2 = deck.slides
    assert s1.title == "Quarterly review"
    assert [t for _, t in s1.text[0].lines] == ["Revenue grew and costs held flat", "See https://example.com/q3."]
    assert s1.links == ["https://example.com/q3"]
    assert s2.title == "Next steps"


def test_join_lines():
    assert join_lines(["• One", "wrapped", "Two", "- Three"]) == ["One wrapped", "Two", "Three"]
    assert join_lines(["lower first"]) == ["lower first"]
