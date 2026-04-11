"""Tests for v0.1.8: `validate` subcommand and `generate --dry-run` flag."""

from __future__ import annotations

import shutil
from pathlib import Path

import yaml
from typer.testing import CliRunner

from recombinase.cli import app

runner = CliRunner()


# -- validate: happy path --------------------------------------------------


def test_validate_succeeds_when_all_shapes_match(
    simple_template: Path, tmp_path: Path, write_config
) -> None:
    config_path = tmp_path / "config.yaml"
    write_config(config_path, simple_template, {"name": "Field"})

    result = runner.invoke(app, ["validate", "-c", str(config_path)])
    assert result.exit_code == 0, result.output
    assert "Config is valid" in result.output
    assert "Matched 1/1" in result.output


def test_validate_auto_detects_config_in_scaffolded_project(
    simple_template: Path, tmp_path: Path, monkeypatch, write_config
) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    (project / "template").mkdir()
    shutil.copy(simple_template, project / "template" / "t.pptx")
    write_config(
        project / "template" / "config.yaml",
        project / "template" / "t.pptx",
        {"name": "Field"},
    )

    monkeypatch.chdir(project)
    result = runner.invoke(app, ["validate"])
    assert result.exit_code == 0, result.output
    assert "Config is valid" in result.output


# -- validate: missing-shape error ----------------------------------------


def test_validate_fails_on_missing_shape(
    simple_template: Path, tmp_path: Path, write_config
) -> None:
    config_path = tmp_path / "config.yaml"
    write_config(config_path, simple_template, {"name": "Nonexistent_Shape"})

    result = runner.invoke(app, ["validate", "-c", str(config_path)])
    assert result.exit_code == 1
    combined = result.output + (result.stderr if hasattr(result, "stderr") else "")
    assert "Missing shapes" in combined
    assert "Nonexistent_Shape" in combined
    assert "name" in combined  # field name surfaced in error


def test_validate_missing_shape_suggests_inspect(
    simple_template: Path, tmp_path: Path, write_config
) -> None:
    config_path = tmp_path / "config.yaml"
    write_config(config_path, simple_template, {"name": "Typo_Name"})

    result = runner.invoke(app, ["validate", "-c", str(config_path)])
    assert result.exit_code == 1
    combined = result.output + (result.stderr if hasattr(result, "stderr") else "")
    assert "recombinase inspect" in combined


# -- validate: unused shapes warning --------------------------------------


def test_validate_warns_on_unused_shapes(rich_template: Path, tmp_path: Path, write_config) -> None:
    """rich_template has 5 shapes. Config only maps 2 → 3 unused, warning-level."""
    config_path = tmp_path / "config.yaml"
    write_config(
        config_path,
        rich_template,
        {"name": "Consultant_Name", "role": "Role_Title"},
    )

    result = runner.invoke(app, ["validate", "-c", str(config_path)])
    # Unused shapes are a warning, not an error — exit 0 without --strict
    assert result.exit_code == 0
    combined = result.output + (result.stderr if hasattr(result, "stderr") else "")
    assert "Unused shapes" in combined
    assert "Summary_Body" in combined  # one of the unused ones


def test_validate_strict_fails_on_unused_shapes(
    rich_template: Path, tmp_path: Path, write_config
) -> None:
    config_path = tmp_path / "config.yaml"
    write_config(
        config_path,
        rich_template,
        {"name": "Consultant_Name"},
    )

    result = runner.invoke(app, ["validate", "-c", str(config_path), "--strict"])
    # Strict escalates unused to exit 2
    assert result.exit_code == 2


# -- validate: config lookup failures -------------------------------------


def test_validate_errors_when_no_config_and_nothing_to_autodetect(
    tmp_path: Path, monkeypatch
) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    monkeypatch.chdir(project)

    result = runner.invoke(app, ["validate"])
    assert result.exit_code == 1
    combined = result.output + (result.stderr if hasattr(result, "stderr") else "")
    assert "No --config" in combined or "not found" in combined


# -- generate --dry-run ---------------------------------------------------


def test_generate_dry_run_does_not_write_output(
    simple_template: Path, tmp_path: Path, write_config
) -> None:
    config_path = tmp_path / "config.yaml"
    write_config(config_path, simple_template, {"name": "Field"}, overflow_ratio=0)
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    (data_dir / "alpha.yaml").write_text(
        yaml.safe_dump({"id": "alpha", "name": "Alpha"}), encoding="utf-8"
    )
    output = tmp_path / "out.pptx"

    result = runner.invoke(
        app,
        [
            "generate",
            "-c",
            str(config_path),
            "-d",
            str(data_dir),
            "-o",
            str(output),
            "--dry-run",
        ],
    )
    assert result.exit_code == 0, result.output
    assert not output.exists()
    assert "DRY RUN" in result.output


def test_generate_dry_run_reports_warnings_without_writing(
    simple_template: Path, tmp_path: Path, write_config
) -> None:
    """Dry run should still produce the warning list users rely on for QA."""
    config_path = tmp_path / "config.yaml"
    write_config(
        config_path,
        simple_template,
        {"name": "Field", "extra": "Nonexistent_Shape"},
        overflow_ratio=0,
    )
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    (data_dir / "r.yaml").write_text(yaml.safe_dump({"name": "A", "extra": "B"}), encoding="utf-8")
    output = tmp_path / "out.pptx"

    result = runner.invoke(
        app,
        [
            "generate",
            "-c",
            str(config_path),
            "-d",
            str(data_dir),
            "-o",
            str(output),
            "--dry-run",
        ],
    )
    assert result.exit_code == 0, result.output
    assert not output.exists()
    combined = result.output + (result.stderr if hasattr(result, "stderr") else "")
    assert "Nonexistent_Shape" in combined  # missing-shape warning surfaced


def test_generate_dry_run_ignores_existing_output(
    simple_template: Path, tmp_path: Path, write_config
) -> None:
    """Dry-run should not trip the overwrite guard — it's not writing anything."""
    config_path = tmp_path / "config.yaml"
    write_config(config_path, simple_template, {"name": "Field"}, overflow_ratio=0)
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    (data_dir / "r.yaml").write_text(yaml.safe_dump({"name": "A"}), encoding="utf-8")

    # Pre-existing output file that would normally trigger the refuse-overwrite guard
    output = tmp_path / "existing.pptx"
    output.write_bytes(b"placeholder")

    result = runner.invoke(
        app,
        [
            "generate",
            "-c",
            str(config_path),
            "-d",
            str(data_dir),
            "-o",
            str(output),
            "--dry-run",
        ],
    )
    assert result.exit_code == 0, result.output
    assert output.read_bytes() == b"placeholder"  # untouched


def test_generate_dry_run_exits_2_strict_with_warnings(
    simple_template: Path, tmp_path: Path, write_config
) -> None:
    config_path = tmp_path / "config.yaml"
    write_config(config_path, simple_template, {"name": "Nonexistent"}, overflow_ratio=0)
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    (data_dir / "r.yaml").write_text(yaml.safe_dump({"name": "A"}), encoding="utf-8")
    output = tmp_path / "out.pptx"

    result = runner.invoke(
        app,
        [
            "generate",
            "-c",
            str(config_path),
            "-d",
            str(data_dir),
            "-o",
            str(output),
            "--dry-run",
            "--strict",
        ],
    )
    assert result.exit_code == 2
    assert not output.exists()
