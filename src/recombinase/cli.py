"""CLI entry point for recombinase, built on Typer for rich human UX.

Subcommands:
- inspect  : print structural metadata of a pptx template
- init     : write a scaffold config YAML from a template's shape names
- generate : populate a template from YAML records and write to an output pptx
"""

from __future__ import annotations

import os
from pathlib import Path

import typer

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

app = typer.Typer(
    name="recombinase",
    help=(
        "Template-guided pptx synthesis: inspect templates, scaffold configs, "
        "and generate populated decks from structured YAML data."
    ),
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
            "$env:OneDrive\\cv on Windows (when OneDrive is configured), "
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
            f"No path given — defaulting to {project_dir}",
            fg=typer.colors.CYAN,
        )

    if project_dir.exists() and any(project_dir.iterdir()) and not force:
        typer.secho(
            f"Directory already exists and is not empty: {project_dir} (use --force to proceed)",
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
            "- `output/` — generated decks land here (ignored by default conventions)\n\n"
            "## Typical workflow\n\n"
            "```\n"
            "recombinase inspect template/<your-template>.pptm\n"
            "recombinase init template/<your-template>.pptm -o template/config.yaml\n"
            "# edit template/config.yaml to map field names → shape names\n"
            "# write one .yaml file per record in cv-data/\n"
            "recombinase generate -c template/config.yaml -d cv-data/ "
            "-o output/deck.pptx\n"
            "```\n",
            encoding="utf-8",
        )

    typer.secho(f"Created project: {project_dir}", fg=typer.colors.GREEN)
    for sub in subfolders:
        typer.echo(f"  {project_dir / sub}")
    typer.echo("\nNext: place your template file in the template/ folder, then run:")
    typer.echo(f'  recombinase inspect "{project_dir / "template"}/<your-template>.pptm"')


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
    """Print structural metadata of a pptx template (no text content)."""
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
            f"No shapes found on slide {source_slide_index} of {template}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    if output.exists() and not force:
        typer.secho(
            f"Config file already exists: {output} (use --force to overwrite)",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    write_scaffold_config(template, shape_names, output)
    typer.secho(f"Wrote scaffold config: {output}", fg=typer.colors.GREEN)
    typer.echo(f"Found {len(shape_names)} shape(s) on slide {source_slide_index}.")
    typer.echo("Edit the placeholders section to map your data fields to shape names.")


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
        help="Path to write the generated pptx.",
    ),
    strict: bool = typer.Option(
        False,
        "--strict",
        help="Exit non-zero if any record produced warnings.",
    ),
) -> None:
    """Generate a populated pptx deck from a template + YAML data directory."""
    cfg = load_config(config)
    records = load_records(data_dir)

    if not records:
        typer.secho(
            f"No YAML records found in {data_dir}",
            fg=typer.colors.RED,
            err=True,
        )
        raise typer.Exit(code=1)

    result = generate_deck(cfg, records, output)

    typer.secho(f"Generated: {result['output']}", fg=typer.colors.GREEN)
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


def main(argv: list[str] | None = None) -> int:
    """Entry point. Accepts optional argv list for programmatic / test use."""
    try:
        app(args=argv, standalone_mode=False)
    except typer.Exit as exc:
        return exc.exit_code
    except (FileNotFoundError, ValueError) as exc:
        typer.secho(f"Error: {exc}", fg=typer.colors.RED, err=True)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
