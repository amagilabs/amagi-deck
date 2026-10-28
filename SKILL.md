---
name: amagi-deck
description: Use when the user hands over a slide deck (.pptx, .key, .ppt, .odp or a slide PDF) and wants it reviewed, summarized, critiqued, rewritten, or turned into a blog post, memo or skill. amagi-deck turns the deck into one markdown file with a PNG of every slide plus its text, speaker notes, tables and the numbers behind each chart, so you can see the slides and read the data instead of guessing from extracted text.
---

# amagi-deck

## When to use

Use this skill when the user gives you a presentation file and asks you to do something with it:
"review my pitch deck", "what's wrong with slide 6", "turn this webinar deck into a blog post",
"pull the numbers out of these charts", "summarize the board deck". Text extraction alone
leaves out the layout, and often the chart values. amagi-deck gives you a rendered image of each
slide to look at and each chart's data as a table.

## Workflow

1. **Convert the deck.** Run it in a scratch folder so the output does not land in the user's
   project by accident:

   ```
   amagi-deck path/to/deck.pptx -o /tmp/deck-review/deck.md
   ```

   This writes `deck.md` and `deck_slides/slide-01.png`, `slide-02.png`, ... next to it. The
   last line on stderr reports how many slides, images, charts and tables it found.

2. **Read `deck.md` first.** It has an outline, then one section per slide:
   - `## Slide N: title`, with `(hidden)` on slides hidden in the deck
   - the slide image link
   - bullet text, with sub-bullets indented
   - `**Table**` and `**Chart: title** (kind)` sections, written as markdown tables
   - `**Pictures**`, listing each picture's alt text, or `(no alt text)` when it has none
   - `**Links**` and `**Speaker notes**`

3. **Look at the slide images** for anything about design, layout or readability. Open the PNGs
   for the slides you are commenting on. Do not judge visual design from the text alone.

4. **Check the header line** before trusting the images. `Slide images: powerpoint` or
   `soffice` means a real render. `Slide images: sketch` means neither PowerPoint nor
   LibreOffice was available. The images are then rough drawings: boxes, text, pictures and
   simple bar/line charts, without theme colors, gradients or effects. Say so if you
   comment on how slides look. `> Note:` lines under the header explain any fallback.

## Useful options

- `--slides 1-3,7`: convert only some slides.
- `--render none`: text, notes and chart data only, with no images. This is the fastest option.
- `--no-notes`: leave out speaker notes (for example, before sharing the markdown).
- `-o -`: print the markdown to stdout. Images still go to `<name>_slides/` in the current folder.
- `--json`: the same content as structured JSON.
- `amagi-deck --renderers`: shows which renderers this machine has.

## Limits to tell the user about

- Chart numbers come from the data cached inside the .pptx. A chart pasted in as a picture has
  no data, so only the image shows it.
- A slide PDF has no notes, charts or tables as data. Its text comes from the PDF text layer,
  so chart axis labels can show up as bullet lines.
- `.key`, `.ppt` and `.odp` need LibreOffice to convert (PowerPoint also works for .ppt and
  .odp). LibreOffice's Keynote import does not cover every Keynote feature. If a Keynote
  deck comes out empty or wrong, ask the user to export it from Keynote (File > Export To >
  PowerPoint or PDF) and run that file.
- Speaker notes can hold private remarks. Ask before you quote them in anything that leaves the
  conversation.
