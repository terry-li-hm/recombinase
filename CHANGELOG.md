# Changelog

All notable changes to this project will be documented here.

The format is loosely based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.5] - 2026-04-11

### Fixed

- **`duplicate_slide` now preserves shape relationships** — previously the
  lxml deep-copy of shape XML left `r:id` / `r:embed` references pointing at
  the source slide's rels file, so any pptx template containing pictures,
  hyperlinks, or embedded charts would silently produce broken content on
  the duplicated slide. The new implementation copies the source slide's
  relationships onto the new slide and rewrites the r: references in the
  copied XML against the new rId map. Surfaced by the v0.1.5 code review;
  would have fired on the first real CV template with a headshot.
- **Group-shape recursion** — `find_shape_by_name`, `inspect_template`, and
  `shape_names_from_slide` now descend into `p:grpSp` group shapes so every
  named shape at any nesting depth is reachable. CV templates commonly group
  `Name + Role + Headshot` as a single unit, which was previously invisible
  to both discovery and population.
- **Config validation is now defensive** — `load_config` raises a clean
  `ValueError` with the file path prefix on non-dict YAML, empty files,
  invalid YAML syntax, wrong field types (e.g. `source_slide_index: two`),
  and malformed placeholder mappings. Previously these raised raw
  `TypeError`/`AttributeError` tracebacks.
- **`cmd_generate` hardening** — creates parent directories automatically,
  refuses to overwrite an existing output file without `--force`, warns
  when the output suffix is not `.pptx`, and shows progress messages for
  each pipeline phase so slow runs no longer look hung.
- **Top-level exception handler** in `main()` now traps
  `PermissionError` (with a PowerPoint-lock hint), `PackageNotFoundError`
  (corrupt pptx), `yaml.YAMLError`, `NotADirectoryError`, and a
  last-resort bare `Exception` to guarantee a CLI user never sees a raw
  Python traceback. Set `RECOMBINASE_DEBUG=1` to restore tracebacks.

### Changed

- `recombinase --help` epilog now shows the full `new → inspect → init →
  generate` workflow order so a first-time user does not need to read four
  subcommand helps.
- README §Usage gains a `0. Scaffold a project` section explaining the new
  subcommand before the existing inspect/init/generate flow.
- `cmd_new` scaffold README template now uses `YOUR_TEMPLATE.pptm` instead
  of `<your-template>.pptm`, which was mis-parsed as shell redirection in
  `cmd.exe`.
- `set_shape_value` simplified: the int/float branch collapsed into the
  scalar fallthrough, the recursive newline-split call replaced by a
  direct call to a `_write_paragraphs` helper. The docstring now correctly
  lists the supported types.
- Scaffold config comment fixed: the generated YAML no longer claims it
  came from a non-existent `recombinase inspect --write-config` command.
- Slug generation in `write_scaffold_config` now handles any non-alphanumeric
  character in shape names via a single regex, instead of the old ad-hoc
  chain of `.replace(" ", "_").replace("-", "_").replace(".", "_")`.

### Removed

- Dead `bullet_join` field from `TemplateConfig` — declared, parsed, and
  never read anywhere in the package. Its docstring also described behavior
  `set_shape_value` never implemented.
- Dead `_source_file` metadata injection in `load_records` — silently
  polluted every loaded record with a key that was never consumed
  downstream and could collide with user-defined fields.

### Added

- 29 new regression tests covering every P0/P1 fix above:
  image-rel preservation across slide duplication (with a round-trip
  save/reopen assertion), group-shape walking and lookup, parametrized
  `set_shape_value` coverage of int/float/None/empty-list/empty-string,
  seven config-validation error paths, `generate_deck` out-of-range
  index, CLI `generate` parent-dir creation, overwrite refusal, `--force`
  override, non-pptx suffix warning, clean `main()` handling of missing
  files and corrupt pptx, and workflow-epilog visibility in `--help`.
  Test count: 22 → 51.
- `.pre-commit-config.yaml` gains no new hooks; the existing ruff +
  hygiene pass now runs against the larger codebase unchanged.

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
