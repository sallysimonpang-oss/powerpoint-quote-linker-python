"""Command-line interface for the PowerPoint quote linker."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

from .linker import DEFAULT_TARGET_TEMPLATE, link_presentation


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ppt-quote-linker",
        description="Add Google Books hyperlinks to bold phrases in quoted SmartArt text.",
    )
    parser.add_argument("input", type=Path, help="input .pptx file")
    parser.add_argument("output", type=Path, help="output .pptx file")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.input.suffix.lower() != ".pptx":
        raise SystemExit("input must be a .pptx file")
    if args.output.suffix.lower() != ".pptx":
        raise SystemExit("output must be a .pptx file")
    if not args.input.is_file():
        raise SystemExit(f"input file does not exist: {args.input}")
    if args.input.resolve() == args.output.resolve():
        raise SystemExit("input and output paths must be different")

    count = link_presentation(args.input, args.output, DEFAULT_TARGET_TEMPLATE)
    noun = "hyperlink" if count == 1 else "hyperlinks"
    print(f"Created {args.output} with {count} {noun}.")
    return 0
