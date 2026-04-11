"""CLI entry point for recombinase, built on Typer for rich human UX.

Subcommands:
- new      : scaffold a project directory (template/, cv-data/, output/)
- inspect  : print structural metadata of a pptx template
- init     : write a scaffold config YAML from a template's shape names
- generate : populate a template from YAML records and write to an output pptx
"""

from __future__ import annotations

import os
import traceback
from collections.abc import Callable
from pathlib import Path

import typer
import yaml
from pptx.exc import PackageNotFoundError

from recombinase import __version__
from recombinase.config import load_config, write_scaffold_config
from recombinase.generate import generate_deck, load_records
from recombinase.inspect import (
    format_template_info,
    inspect_template,
    shape_names_from_slide,
)


def _default_project_dir() -> Path:
    """Pick a sensible default directory for `recombinase new`.

    On Windows with OneDrive configured, use `$env:OneDrive\\cv`. On other
    platforms (or if OneDrive isn't set), fall back to `~/cv`. Users can
    always override by passing an explicit path. Windows environment variables
    are case-insensitive, so using uppercase keys here works on all platforms.
    """
    onedrive = os.environ.get("ONEDRIVE") or os.environ.get("ONEDRIVECOMMERCIAL")
    if onedrive:
        return Path(onedrive) / "cv"
    return Path.home() / "cv"


_WORKFLOW_EPILOG = """\
[bold]Typical workflow[/bold] (run in order):

  1. [cyan]recombinase new[/cyan]                           scaffold a project folder
  2. [cyan]recombinase inspect[/cyan] TEMPLATE              see template shape names
  3. [cyan]recombinase init[/cyan] TEMPLATE                 write a config to edit
     [dim](then edit the config to map field names to shape names)[/dim]
  4. [cyan]recombinase generate[/cyan] -c CONFIG -d DATA -o OUT.pptx

Run [cyan]recombinase <command> --help[/cyan] for details on a single command.
"""


app = typer.Typer(
    name="recombinase",
    help=(
        "Template-guided pptx synthesis: inspect templates, scaffold configs, "
        "and generate populated decks from structured YAML data."
    ),
    epilog=_WORKFLOW_EPILOG,
    no_args_is_help=True,
    add_completion=False,
    rich_markup_mode="rich",
)


def _version_callback(value: bool) -> None:
    if value:
        typer.echo(f"recombinase {__version__}")
        raise typer.Exit()


@app.callback()
def _root(
    version: bool = typer.Option(
        False,
        "--version",
        "-V",
        callback=_version_callback,
        is_eager=True,
        help="Show version and exit.",
    ),
) -> None:
    """Recombinase — template-guided pptx synthesis."""


@app.command("new")
def cmd_new(
    project_dir: Path = typer.Argument(
        None,
        help=(
            "Path to the project directory to create. If omitted, defaults to "
            "$env:OneDrive/cv on Windows (when OneDrive is configured), "
            "otherwise ~/cv."
        ),
        resolve_path=True,
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Proceed even if the target directory already exists and is non-empty.",
    ),
) -> None:
    """Scaffold a new project directory with template/, cv-data/, and output/ subfolders.

    Creates a conventional folder layout under the given path:

        <project-dir>/
          template/   — put your .pptx/.pptm template here
          cv-data/    — put your per-record YAML files here
          output/     — generated decks will land here

    Safe to run in OneDrive — it's a plain mkdir + README write, no sync surprises.
    """
    if project_dir is None:
        project_dir = _default_project_dir()
        typer.secho(
            f"No path given — defaulting to '{project_dir}'",
            fg=typer.colors.CYAN,
        )

    if project_dir.exists() and any(project_dir.iterdir()) and not force:
        typer.secho(
            f"Directory already exists and is not empty: '{project_dir}' (use --force to proceed)",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    subfolders = ("template", "cv-data", "output")
    for sub in subfolders:
        (project_dir / sub).mkdir(parents=True, exist_ok=True)

    readme_path = project_dir / "README.md"
    if not readme_path.exists():
        readme_path.write_text(
            "# Recombinase project\n\n"
            "This folder was scaffolded by `recombinase new`.\n\n"
            "## Layout\n\n"
            "- `template/` — place the source .pptx/.pptm template file here\n"
            "- `cv-data/` — one YAML file per record (consultant, use case, etc.)\n"
            "- `output/` — generated decks land here\n\n"
            "## Typical workflow\n\n"
            "```\n"
            "recombinase inspect template/YOUR_TEMPLATE.pptm\n"
            "recombinase init template/YOUR_TEMPLATE.pptm -o template/config.yaml\n"
            "# edit template/config.yaml to map field names to shape names\n"
            "# write one .yaml file per record in cv-data/\n"
            "recombinase generate -c template/config.yaml -d cv-data/ "
            "-o output/deck.pptx\n"
            "```\n",
            encoding="utf-8",
        )

    typer.secho(f"Created project: '{project_dir}'", fg=typer.colors.GREEN)
    for sub in subfolders:
        typer.echo(f"  {project_dir / sub}")
    typer.echo("\nNext: place your template file in the template/ folder, then run:")
    typer.echo(f'  recombinase inspect "{project_dir / "template" / "YOUR_TEMPLATE.pptm"}"')


@app.command("inspect")
def cmd_inspect(
    template: Path = typer.Argument(
        ...,
        exists=True,
        file_okay=True,
        dir_okay=False,
        readable=True,
        resolve_path=True,
        help="Path to a .pptx/.pptm template file.",
    ),
) -> None:
    """Print structural metadata of a pptx template.

    Shape names, types, placeholder info, and text character counts only —
    no actual slide text is read or printed, so output is safe to share.
    Run this first on any new template to discover the shape names you'll
    reference in the config's `placeholders:` section.
    """
    info = inspect_template(template)
    typer.echo(format_template_info(info))


@app.command("init")
def cmd_init(
    template: Path = typer.Argument(
        ...,
        exists=True,
        file_okay=True,
        dir_okay=False,
        readable=True,
        resolve_path=True,
        help="Path to a .pptx/.pptm template file.",
    ),
    source_slide_index: int = typer.Option(
        1,
        "--source-slide-index",
        "-s",
        min=1,
        help="1-based index of the slide to read shape names from.",
    ),
    output: Path = typer.Option(
        Path("template-config.yaml"),
        "--output",
        "-o",
        help="Path to write the scaffold config YAML.",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Overwrite an existing output file.",
    ),
) -> None:
    """Write a scaffold config YAML from a template's shape names."""
    info = inspect_template(template)
    shape_names = shape_names_from_slide(info, source_slide_index)

    if not shape_names:
        typer.secho(
            f"No shapes found on slide {source_slide_index} of '{template}'",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    if output.exists() and not force:
        typer.secho(
            f"Config file already exists: '{output}' (use --force to overwrite)",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    output.parent.mkdir(parents=True, exist_ok=True)
    write_scaffold_config(template, shape_names, output)
    typer.secho(f"Wrote scaffold config: '{output}'", fg=typer.colors.GREEN)
    typer.echo(f"Found {len(shape_names)} shape(s) on slide {source_slide_index}.")
    typer.echo(
        f'Next: edit "{output}" to map your data field names (left) to the shape names (right).'
    )


@app.command("generate")
def cmd_generate(
    config: Path = typer.Option(
        ...,
        "--config",
        "-c",
        exists=True,
        file_okay=True,
        dir_okay=False,
        readable=True,
        resolve_path=True,
        help="Path to the template config YAML.",
    ),
    data_dir: Path = typer.Option(
        ...,
        "--data-dir",
        "-d",
        exists=True,
        file_okay=False,
        dir_okay=True,
        readable=True,
        resolve_path=True,
        help="Directory containing per-record YAML files.",
    ),
    output: Path = typer.Option(
        ...,
        "--output",
        "-o",
        resolve_path=True,
        help="Path to write the generated pptx.",
    ),
    force: bool = typer.Option(
        False,
        "--force",
        "-f",
        help="Overwrite an existing output file without prompting.",
    ),
    strict: bool = typer.Option(
        False,
        "--strict",
        help=(
            "Exit non-zero if any record had missing-field or missing-shape "
            "warnings. Default: warnings print but exit 0."
        ),
    ),
) -> None:
    """Generate a populated pptx deck from a template + YAML data directory."""
    if output.suffix.lower() != ".pptx":
        typer.secho(
            f"Warning: output path '{output}' does not end in .pptx — "
            "PowerPoint may not open it correctly.",
            fg=typer.colors.YELLOW,
            err=True,
        )

    if output.exists() and not force:
        typer.secho(
            f"Output file already exists: '{output}' (use --force to overwrite)",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    output.parent.mkdir(parents=True, exist_ok=True)

    typer.secho(f"Loading config: '{config}'", fg=typer.colors.CYAN)
    cfg = load_config(config)

    typer.secho(f"Loading records from: '{data_dir}'", fg=typer.colors.CYAN)
    records = load_records(data_dir)

    if not records:
        typer.secho(
            f"No YAML records found in '{data_dir}'",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    typer.secho(
        f"Generating {len(records)} slide(s) from '{cfg.template.name}'...",
        fg=typer.colors.CYAN,
    )
    result = generate_deck(cfg, records, output)

    typer.secho(f"Generated: '{result['output']}'", fg=typer.colors.GREEN)
    typer.echo(f"Records: {result['records_generated']}")

    if result["warnings"]:
        typer.secho(
            f"Warnings ({len(result['warnings'])}):",
            fg=typer.colors.YELLOW,
            err=True,
        )
        for warning in result["warnings"]:
            typer.secho(f"  - {warning}", fg=typer.colors.YELLOW, err=True)
        if strict:
            raise typer.Exit(code=2)


def _print_error(message: str) -> None:
    typer.secho(f"Error: {message}", fg=typer.colors.RED, err=True)


def _format_permission_error(exc: BaseException) -> str:
    target = getattr(exc, "filename", None)
    if target:
        return f"Cannot write '{target}': the file may be open in PowerPoint. Close it and re-run."
    return f"Permission denied: {exc}"


def _format_package_not_found(exc: BaseException) -> str:
    return (
        f"Not a valid pptx file: {exc}. The file may be corrupt, empty, "
        "or in a format python-pptx can't read."
    )


# Dispatch table mapping exception classes to (formatter, exit_code) tuples.
# Order matters — more specific classes must come before their superclasses.
# Adding a new trap is a one-line addition to this table.
_Formatter = Callable[[BaseException], str]

_ERROR_HANDLERS: tuple[tuple[type[BaseException], _Formatter, int], ...] = (
    (PermissionError, _format_permission_error, 1),
    (PackageNotFoundError, _format_package_not_found, 1),
    (yaml.YAMLError, lambda exc: f"Invalid YAML: {exc}", 1),
    (FileNotFoundError, lambda exc: str(exc), 1),
    (NotADirectoryError, lambda exc: str(exc), 1),
    (ValueError, lambda exc: str(exc), 1),
)


def main(argv: list[str] | None = None) -> int:
    """Entry point. Accepts optional argv list for programmatic / test use.

    Traps known failure modes via `_ERROR_HANDLERS` and prints a clean error
    message instead of a traceback. Set `RECOMBINASE_DEBUG=1` to restore
    the traceback for debugging.
    """
    debug = os.environ.get("RECOMBINASE_DEBUG") == "1"
    try:
        app(args=argv, standalone_mode=False)
    except typer.Exit as exc:
        return exc.exit_code
    except BaseException as exc:
        for exc_type, formatter, exit_code in _ERROR_HANDLERS:
            if isinstance(exc, exc_type):
                _print_error(formatter(exc))
                if debug:
                    traceback.print_exc()
                return exit_code
        # Unclassified: last-resort guard so a human-facing CLI never shows
        # a raw Python traceback on an unexpected exception class.
        _print_error(f"Unexpected {type(exc).__name__}: {exc}")
        if debug:
            traceback.print_exc()
        else:
            typer.secho(
                "  (set RECOMBINASE_DEBUG=1 to see the full traceback)",
                fg=typer.colors.RED,
                err=True,
            )
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
