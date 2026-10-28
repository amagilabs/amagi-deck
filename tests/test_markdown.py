from amagi_deck.markdown import chart_rows, deck_md, fmt_num, md_table, plural, slide_md
from amagi_deck.model import Chart, Deck, Picture, Series, Slide, Table, TextBlock


def test_fmt_num():
    assert fmt_num(None) == ""
    assert fmt_num(3.0) == "3"
    assert fmt_num(2.5) == "2.5"
    assert fmt_num(1 / 3) == "0.3333333333"


def test_plural():
    assert plural(1, "slide") == "1 slide"
    assert plural(0, "chart") == "0 charts"


def test_md_table_escapes_and_pads():
    out = md_table([["a|b", "c"], ["only one"]])
    assert out == ["| a\\|b | c |", "|---|---|", "| only one |  |"]


def test_chart_rows_category_and_scatter():
    cat = Chart(kind="line", categories=["Jan", ""], series=[Series("S", [1.0, None])])
    assert chart_rows(cat) == [["Category", "S"], ["Jan", "1"], ["2", ""]]
    xy = Chart(kind="scatter", series=[Series("P", [2.0], x_values=[1.5])])
    assert chart_rows(xy) == [["Series", "X", "Y"], ["P", "1.5", "2"]]
    assert chart_rows(Chart(kind="pie")) == []


def test_slide_md_sections_in_order():
    slide = Slide(number=3, title="Plan", hidden=True,
                  text=[TextBlock("body", [(0, "Top"), (1, "Nested")])],
                  tables=[Table([["k", "v"], ["x", "1"]])],
                  charts=[Chart(kind="pie", title="Mix", categories=["A"], series=[Series("S", [1.0])]),
                          Chart(kind="line")],
                  pictures=[Picture("Logo", ""), Picture("Photo", "A dog")],
                  links=["https://example.com"], notes="Line one\n\nLine two",
                  image="my deck_slides/slide-03.png")
    text = "\n".join(slide_md(slide))
    assert text.startswith("## Slide 3: Plan (hidden)\n\n![Slide 3](my%20deck_slides/slide-03.png)")
    assert "- Top\n  - Nested" in text
    assert "**Chart: Mix** (pie)" in text
    assert "**Chart** (line)\n\n(no data cached in the file)" in text
    assert '- Logo (no alt text)\n- Photo: "A dog"' in text
    assert "> Line one\n>\n> Line two" in text
    order = [text.index(k) for k in ("**Table**", "**Chart: Mix**", "**Pictures**", "**Links**", "**Speaker notes**")]
    assert order == sorted(order)
    assert "Speaker notes" not in "\n".join(slide_md(slide, include_notes=False))


def test_deck_md_header_outline_and_warnings():
    deck = Deck(source="d.pptx", kind="pptx", width_in=10, height_in=7.5, title="Board deck",
                renderer="sketch", warnings=["approximate"],
                slides=[Slide(number=1, title="One", notes="n"), Slide(number=2, hidden=True)])
    md = deck_md(deck)
    assert md.startswith("# Board deck\n\nSource: `d.pptx` (pptx, 10 x 7.5 in). Slide images: sketch. "
                         "2 slides, 1 hidden, 1 with notes, 0 charts, 0 tables.")
    assert "> Note: approximate" in md
    assert "## Outline\n\n1. One\n2. (untitled) (hidden)" in md
    assert "## Slide 2: (untitled) (hidden)" in md
    assert md.endswith("\n") and not md.endswith("\n\n")
