"""Tests for v0.1.7 auto-detection: zero-arg CLI use inside a scaffolded folder."""

from __future__ import annotations

import shutil
from pathlib import Path

import yaml
from typer.testing import CliRunner

from recombinase.cli import _find_template_in_cwd, app

runner = CliRunner()


def test_find_template_in_cwd_picks_template_subfolder(
    simple_template: Path, tmp_path: Path, monkeypatch
) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    (project / "template").mkdir()
    shutil.copy(simple_template, project / "template" / "mytemplate.pptx")

    monkeypatch.chdir(project)
    result = _find_template_in_cwd()
    assert result is not None
    assert result.name == "mytemplate.pptx"
    assert "template" in result.parts


def test_find_template_in_cwd_falls_back_to_cwd(
    simple_template: Path, tmp_path: Path, monkeypatch
) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    shutil.copy(simple_template, project / "cv.pptx")

    monkeypatch.chdir(project)
    result = _find_template_in_cwd()
    assert result is not None
    assert result.name == "cv.pptx"


def test_find_template_in_cwd_returns_none_when_ambiguous(
    simple_template: Path, tmp_path: Path, monkeypatch
) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    (project / "template").mkdir()
    shutil.copy(simple_template, project / "template" / "a.pptx")
    shutil.copy(simple_template, project / "template" / "b.pptx")

    monkeypatch.chdir(project)
    assert _find_template_in_cwd() is None


def test_find_template_in_cwd_returns_none_when_missing(tmp_path: Path, monkeypatch) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    monkeypatch.chdir(project)
    assert _find_template_in_cwd() is None


def test_cli_inspect_no_args_auto_detects(
    simple_template: Path, tmp_path: Path, monkeypatch
) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    (project / "template").mkdir()
    shutil.copy(simple_template, project / "template" / "t.pptx")

    monkeypatch.chdir(project)
    result = runner.invoke(app, ["inspect"])
    assert result.exit_code == 0, result.output
    assert "Auto-detected template" in result.output
    assert "Field" in result.output  # the named shape from simple_template


def test_cli_inspect_no_args_errors_cleanly_when_missing(tmp_path: Path, monkeypatch) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    monkeypatch.chdir(project)
    result = runner.invoke(app, ["inspect"])
    assert result.exit_code == 1
    assert "No template specified" in result.output or "No template specified" in (
        result.stderr if hasattr(result, "stderr") else ""
    )


def test_cli_init_no_args_auto_detects_and_writes_scaffold(
    simple_template: Path, tmp_path: Path, monkeypatch
) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    (project / "template").mkdir()
    shutil.copy(simple_template, project / "template" / "t.pptx")

    monkeypatch.chdir(project)
    result = runner.invoke(app, ["init"])
    assert result.exit_code == 0, result.output

    # Default output should land inside template/ because that folder exists
    expected_config = project / "template" / "config.yaml"
    assert expected_config.exists()
    content = expected_config.read_text(encoding="utf-8")
    assert "placeholders:" in content
    assert "Field" in content


def test_cli_generate_zero_arg_inside_scaffolded_project(
    simple_template: Path, tmp_path: Path, monkeypatch
) -> None:
    """End-to-end zero-arg `recombinase generate` inside a full scaffolded project."""
    project = tmp_path / "proj"
    project.mkdir()

    # Set up scaffolded layout manually (what `recombinase new` would create)
    (project / "template").mkdir()
    (project / "data").mkdir()
    (project / "output").mkdir()
    shutil.copy(simple_template, project / "template" / "t.pptx")

    # Write config
    (project / "template" / "config.yaml").write_text(
        yaml.safe_dump(
            {
                "template": str(project / "template" / "t.pptx"),
                "source_slide_index": 1,
                "clear_source_slide": True,
                "overflow_ratio": 0,
                "placeholders": {"name": "Field"},
            }
        ),
        encoding="utf-8",
    )

    # Write one record
    (project / "data" / "alpha.yaml").write_text(
        yaml.safe_dump({"id": "alpha", "name": "Alpha Jones"}), encoding="utf-8"
    )

    monkeypatch.chdir(project)
    result = runner.invoke(app, ["generate"])
    assert result.exit_code == 0, result.output

    # Default output is ./output/deck.pptx
    expected_output = project / "output" / "deck.pptx"
    assert expected_output.exists()


def test_cli_generate_zero_arg_errors_cleanly_when_no_config(tmp_path: Path, monkeypatch) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    monkeypatch.chdir(project)
    result = runner.invoke(app, ["generate"])
    assert result.exit_code == 1
    combined = result.output + (result.stderr if hasattr(result, "stderr") else "")
    assert "No --config" in combined or "not found" in combined


def test_cli_generate_zero_arg_errors_cleanly_when_no_data_dir(tmp_path: Path, monkeypatch) -> None:
    project = tmp_path / "proj"
    project.mkdir()
    (project / "template").mkdir()
    (project / "template" / "config.yaml").write_text(
        yaml.safe_dump(
            {
                "template": "nonexistent.pptx",
                "placeholders": {"name": "Field"},
            }
        ),
        encoding="utf-8",
    )

    monkeypatch.chdir(project)
    result = runner.invoke(app, ["generate"])
    assert result.exit_code == 1
    combined = result.output + (result.stderr if hasattr(result, "stderr") else "")
    assert "No --data-dir" in combined or "data" in combined
