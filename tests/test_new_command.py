"""Tests for the `recombinase new` project-scaffold subcommand."""

from __future__ import annotations

from pathlib import Path

import pytest
from typer.testing import CliRunner

from recombinase.cli import _default_project_dir, app

runner = CliRunner()


def test_new_creates_layout(tmp_path: Path) -> None:
    project = tmp_path / "my-pack"
    result = runner.invoke(app, ["new", str(project)])
    assert result.exit_code == 0, result.output

    assert project.exists()
    assert (project / "template").is_dir()
    assert (project / "cv-data").is_dir()
    assert (project / "output").is_dir()
    assert (project / "README.md").is_file()


def test_new_into_empty_existing_dir(tmp_path: Path) -> None:
    project = tmp_path / "my-pack"
    project.mkdir()
    # Empty existing directory is fine
    result = runner.invoke(app, ["new", str(project)])
    assert result.exit_code == 0, result.output
    assert (project / "template").is_dir()


def test_new_refuses_non_empty_dir_without_force(tmp_path: Path) -> None:
    project = tmp_path / "my-pack"
    project.mkdir()
    (project / "existing.txt").write_text("hello")

    result = runner.invoke(app, ["new", str(project)])
    assert result.exit_code == 1
    assert "already exists" in result.output or "already exists" in (
        result.stderr if hasattr(result, "stderr") else ""
    )


def test_new_force_overrides_non_empty_check(tmp_path: Path) -> None:
    project = tmp_path / "my-pack"
    project.mkdir()
    (project / "existing.txt").write_text("hello")

    result = runner.invoke(app, ["new", str(project), "--force"])
    assert result.exit_code == 0, result.output
    assert (project / "template").is_dir()
    assert (project / "existing.txt").exists()  # pre-existing file untouched


def test_new_readme_is_written(tmp_path: Path) -> None:
    project = tmp_path / "my-pack"
    result = runner.invoke(app, ["new", str(project)])
    assert result.exit_code == 0

    readme = project / "README.md"
    assert readme.exists()
    content = readme.read_text(encoding="utf-8")
    assert "Recombinase project" in content
    assert "template/" in content
    assert "cv-data/" in content
    assert "output/" in content


def test_version_flag() -> None:
    result = runner.invoke(app, ["--version"])
    assert result.exit_code == 0
    assert "recombinase" in result.output


def test_default_project_dir_prefers_onedrive(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    # Use uppercase env var names because os.environ is case-sensitive on
    # Linux/macOS (the code path we're actually testing here). Windows
    # resolves case-insensitively so real Windows users still get a hit.
    monkeypatch.setenv("ONEDRIVE", "C:\\Users\\test\\OneDrive - Test")
    monkeypatch.delenv("ONEDRIVECOMMERCIAL", raising=False)
    result = _default_project_dir()
    assert result.name == "cv"
    assert "OneDrive" in str(result)


def test_default_project_dir_falls_back_to_home(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("ONEDRIVE", raising=False)
    monkeypatch.delenv("ONEDRIVECOMMERCIAL", raising=False)
    result = _default_project_dir()
    assert result == Path.home() / "cv"


def test_new_with_no_arg_uses_default(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    # Redirect the default to a tmp path by setting ONEDRIVE env var
    fake_onedrive = tmp_path / "OneDrive"
    fake_onedrive.mkdir()
    monkeypatch.setenv("ONEDRIVE", str(fake_onedrive))

    result = runner.invoke(app, ["new"])
    assert result.exit_code == 0, result.output

    expected_project = fake_onedrive / "cv"
    assert expected_project.exists()
    assert (expected_project / "template").is_dir()
    assert (expected_project / "cv-data").is_dir()
    assert (expected_project / "output").is_dir()
