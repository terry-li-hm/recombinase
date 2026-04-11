"""CLI entry point for recombinase.

Subcommands:
- inspect  : print structural metadata of a pptx template
- init     : write a scaffold config YAML from a template's shape names
- generate : populate a template from YAML records and write to an output pptx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from recombinase import __version__
from recombinase.config import load_config, write_scaffold_config
from recombinase.generate import generate_deck, load_records
from recombinase.inspect import (
    format_template_info,
    inspect_template,
    shape_names_from_slide,
)


def cmd_inspect(args: argparse.Namespace) -> int:
    """Print structural metadata of a pptx template."""
    info = inspect_template(args.template)
    print(format_template_info(info))
    return 0


def cmd_init(args: argparse.Namespace) -> int:
    """Write a scaffold config YAML from a template's source slide shapes."""
    template_path = Path(args.template).expanduser().resolve()
    info = inspect_template(template_path)
    shape_names = shape_names_from_slide(info, args.source_slide_index)

    if not shape_names:
        print(
            f"No shapes found on slide {args.source_slide_index} of {template_path}",
            file=sys.stderr,
        )
        return 1

    output_path = Path(args.output).expanduser().resolve()
    if output_path.exists() and not args.force:
        print(
            f"Config file already exists: {output_path} (use --force to overwrite)",
            file=sys.stderr,
        )
        return 1

    write_scaffold_config(template_path, shape_names, output_path)
    print(f"Wrote scaffold config: {output_path}")
    print(f"Found {len(shape_names)} shape(s) on slide {args.source_slide_index}.")
    print("Edit the placeholders section to map your data fields to shape names.")
    return 0


def cmd_generate(args: argparse.Namespace) -> int:
    """Generate a populated pptx deck from a template + YAML data directory."""
    config = load_config(args.config)
    records = load_records(args.data_dir)

    if not records:
        print(f"No YAML records found in {args.data_dir}", file=sys.stderr)
        return 1

    result = generate_deck(config, records, args.output)

    print(f"Generated: {result['output']}")
    print(f"Records: {result['records_generated']}")
    if result["warnings"]:
        print(f"Warnings ({len(result['warnings'])}):", file=sys.stderr)
        for warning in result["warnings"]:
            print(f"  - {warning}", file=sys.stderr)
        return 2 if args.strict else 0
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="recombinase",
        description=(
            "Template-guided pptx synthesis: inspect templates, scaffold "
            "configs, and generate populated decks from structured YAML data."
        ),
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")

    sub = parser.add_subparsers(dest="command", required=True, metavar="COMMAND")

    # inspect
    p_inspect = sub.add_parser("inspect", help="Print structural metadata of a pptx template.")
    p_inspect.add_argument("template", type=str, help="Path to a .pptx/.pptm file")
    p_inspect.set_defaults(func=cmd_inspect)

    # init
    p_init = sub.add_parser(
        "init",
        help="Write a scaffold config YAML from a template's shape names.",
    )
    p_init.add_argument("template", type=str, help="Path to a .pptx/.pptm file")
    p_init.add_argument(
        "--source-slide-index",
        type=int,
        default=1,
        help="1-based index of the source slide (default: 1)",
    )
    p_init.add_argument(
        "--output",
        "-o",
        type=str,
        default="template-config.yaml",
        help="Path to write the scaffold config (default: ./template-config.yaml)",
    )
    p_init.add_argument(
        "--force",
        action="store_true",
        help="Overwrite existing output file",
    )
    p_init.set_defaults(func=cmd_init)

    # generate
    p_gen = sub.add_parser(
        "generate",
        help="Generate a populated pptx deck from a template + YAML data.",
    )
    p_gen.add_argument(
        "--config",
        "-c",
        type=str,
        required=True,
        help="Path to the template config YAML",
    )
    p_gen.add_argument(
        "--data-dir",
        "-d",
        type=str,
        required=True,
        help="Directory containing per-record YAML files",
    )
    p_gen.add_argument(
        "--output",
        "-o",
        type=str,
        required=True,
        help="Path to write the generated pptx",
    )
    p_gen.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero if any record produced warnings",
    )
    p_gen.set_defaults(func=cmd_generate)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
