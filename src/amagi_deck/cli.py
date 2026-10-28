"""Command line entry point."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import List, Optional

from . import __version__, render
from .core import RENDERERS, DeckError, convert

DESCRIPTION = (
    "Turn a .pptx, .key or slide PDF into one markdown file: a PNG render of every slide plus "
    "its text, speaker notes, tables and chart data, so Claude can review a deck it cannot open."
)
EPILOG = """examples:
  amagi-deck pitch.pptx                    writes pitch.md and pitch_slides/slide-01.png ...
  amagi-deck pitch.pptx -o review/deck.md  writes review/deck.md and review/deck_slides/
  amagi-deck pitch.pdf --slides 1-3        first three slides only
  amagi-deck pitch.pptx --render none      text, notes and chart data only (fast)
  amagi-deck pitch.pptx -o - --json        JSON on stdout
  amagi-deck --renderers                   show which slide renderers this machine has

exit codes: 0 ok, 1 could not read or convert the file, 2 bad arguments"""


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="amagi-deck", description=DESCRIPTION, epilog=EPILOG,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("input", nargs="?", help="deck file: .pptx, .key, .ppt, .odp or .pdf")
    p.add_argument("-o", "--output", help="markdown file to write (default: <input name>.md in the "
                   "current folder; '-' prints to stdout)")
    p.add_argument("--images-dir", help="folder for slide PNGs (default: <output name>_slides next "
                   "to the markdown)")
    p.add_argument("--render", choices=RENDERERS, default="auto",
                   help="slide renderer (default auto: PowerPoint, then LibreOffice, then sketch)")
    p.add_argument("--width", type=int, default=1280, help="slide image width in pixels (default 1280)")
    p.add_argument("--slides", help="only these slides, e.g. 1-3,7")
    p.add_argument("--no-notes", action="store_true", help="leave out speaker notes")
    p.add_argument("--json", action="store_true", help="write JSON instead of markdown")
    p.add_argument("--renderers", action="store_true", help="list available renderers and exit")
    p.add_argument("--version", action="version", version="amagi-deck " + __version__)
    return p


def summary(deck, target: str) -> str:
    n_img = sum(1 for s in deck.slides if s.image)
    parts = ["%d slides" % len(deck.slides)]
    if n_img:
        parts.append("%d images via %s" % (n_img, deck.renderer))
    parts.append("%d charts" % sum(len(s.charts) for s in deck.slides))
    parts.append("%d tables" % sum(len(s.tables) for s in deck.slides))
    parts.append("%d with notes" % sum(1 for s in deck.slides if s.notes))
    return "amagi-deck: wrote %s (%s)" % (target, ", ".join(parts))


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.renderers:
        for name, status in render.available().items():
            print("%-11s %s" % (name, status))
        return 0
    if not args.input:
        parser.print_usage(sys.stderr)
        print("amagi-deck: error: give a deck file (or --renderers)", file=sys.stderr)
        return 2

    src = Path(args.input)
    to_stdout = args.output == "-"
    if to_stdout:
        out = None
    elif args.output:
        out = Path(args.output)
    else:
        out = Path.cwd() / (src.stem + (".json" if args.json else ".md"))
    images_dir = Path(args.images_dir) if args.images_dir else None
    if to_stdout and images_dir is None and args.render != "none":
        images_dir = Path.cwd() / (src.stem + "_slides")

    try:
        deck, text = convert(src, out_md=out, images_dir=images_dir, renderer=args.render,
                             width=args.width, slides=args.slides, notes=not args.no_notes)
    except DeckError as exc:
        msg = str(exc)
        print("amagi-deck: error: %s" % msg, file=sys.stderr)
        return 2 if msg.startswith(("bad --slides", "--slides", "--width")) else 1

    if args.json:
        text = json.dumps(deck.to_dict(), indent=2, ensure_ascii=False) + "\n"
    for w in deck.warnings:
        print("amagi-deck: note: %s" % w, file=sys.stderr)
    if to_stdout:
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
        sys.stdout.write(text)
        return 0
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(text, encoding="utf-8", newline="\n")
    try:
        shown = str(out.relative_to(Path.cwd()))
    except ValueError:
        shown = str(out)
    print(summary(deck, shown), file=sys.stderr)
    return 0
