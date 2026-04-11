# Changelog

All notable changes to this project will be documented here.

The format is loosely based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.4] - 2026-04-11

### Added
- `recombinase new` now works with no arguments — defaults to
  `$env:OneDrive\cv` on Windows when OneDrive is configured, otherwise `~/cv`.
- `CLAUDE.md` documenting project-specific hard constraints, the
  duplicate-and-populate invariant, release process, and non-goals.
- `.pre-commit-config.yaml` with ruff, ruff-format, and stock hygiene hooks.
- GitHub Actions CI workflow running ruff / mypy / pytest on Ubuntu, Windows,
  and macOS across Python 3.10-3.13.

## [0.1.3] - 2026-04-11

### Added
- `recombinase new <project-dir>` subcommand scaffolds a conventional project
  folder layout (`template/`, `cv-data/`, `output/`) with a README, suitable
  for running directly inside a OneDrive folder.

### Changed
- Removed inline `# noqa: UP045` suppression in `cli.py` by switching the
  Typer `--version` callback to use `bool` instead of `Optional[bool]`.
- Added mypy to the dev tooling and verified the package passes a strict run.

## [0.1.2] - 2026-04-11

### Changed
- Swapped `argparse` for `typer` as the CLI framework. Same subcommand surface,
  but with rich help output, colours, and better validation on path arguments.
- Fixed `set_shape_value` to cleanly handle lists, newline-separated strings,
  and scalars without leaking run-level formatting concerns into callers.

## [0.1.1] - 2026-04-11

### Added
- `ruff` + `mypy` + `types-PyYAML` dev dependencies.
- Windows 11 install notes and a recommended OneDrive project layout section
  in the README.

### Changed
- Project pyproject.toml now includes lint/format/type-check configuration.

## [0.1.0] - 2026-04-11

### Added
- Initial release.
- `recombinase inspect` — print structural metadata of a pptx template.
- `recombinase init` — scaffold a template config YAML from shape names.
- `recombinase generate` — populate a template from YAML records.
- 13 end-to-end tests covering config loading, data loading, slide duplication,
  shape text setting (scalar / list / newline), and full pipeline.
