# Changelog

All notable changes to this project will be documented here.

The format is loosely based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.12] - 2026-04-11

### Added

- **Table cell population.** New `tables:` section in the template config
  maps table field names to a `TableConfig` (shape name, ordered column
  list, header_row flag, list_joiner string). `generate_deck` walks each
  configured table after the placeholders loop, reads the record's list
  of dicts, and populates rows cell-by-cell. Header row 0 is skipped by
  default. List values within a cell are joined with `list_joiner`
  (default newline) so bullet-style cells inherit the template's
  paragraph-level formatting via `_write_paragraphs`. Warnings fire on
  over-capacity (more rows than template), missing columns, and
  non-dict row data.
- **Picture placeholder insertion.** `generate_deck` now routes shapes to
  `set_picture` when they are `PicturePlaceholder` instances instead of
  setting their text frame. `set_picture` accepts a file path (relative
  paths resolve against an optional `base_dir`), calls
  `shape.insert_picture(path)`, and silently no-ops on empty/None values
  so the template's example headshot stays in place when a record omits
  the field. Missing files raise `FileNotFoundError` which flows into
  `generate_deck`'s per-record warning collector.
- 15 new regression tests covering: table population with header skip,
  list-joiner cell rendering, over-capacity truncation warning,
  header-row=false write into row 0, non-table-shape warning, picture
  placeholder detection, picture insertion round-trip, missing-file
  error, relative-path resolution against base_dir, empty-value no-op,
  end-to-end table population via generate_deck, table record type
  validation, config parsing for tables section, malformed tables
  rejection. 116 → 131 tests.

### Changed

- `TemplateConfig.validate()` now accepts a config with populated
  `tables:` and empty `placeholders:` (previously required at least one
  placeholder). At least one of the two must be non-empty — a config
  with both empty still fails validation with a clearer message.
- Added `TableConfig` dataclass to the `recombinase.config` module.

### Closes

- The two remaining client template feature gaps (`recent_projects` table
  and `photo` picture placeholder) now work without manual
  post-processing. A client CV run can populate every template field in
  one `recombinase generate` command.

## [0.1.11] - 2026-04-11

### Added

- **PowerPoint lock detection in `PackageNotFoundError` handler.** When
  python-pptx can't open a template file, `_format_package_not_found` now
  parses the path out of the exception message and checks for a matching
  `~$<filename>.pptx` lock-file marker in the same directory — the standard
  Windows Office lock-file convention. If found, emits a specific,
  actionable error: *"Template '...' is currently open in PowerPoint
  (lock file '~$...' detected). Close the file in PowerPoint and re-run."*
  Otherwise falls back to a generic message that still mentions PowerPoint
  as the most likely cause.
- 4 regression tests covering: lock file present → PowerPoint hint; lock
  file absent → generic message still mentions PowerPoint; exception
  without a parseable path → safe fallback; lock file in a different
  directory → not false-positive.

### Why now

Caught on the first real client CV run. The template was open in PowerPoint
(user had just renamed a shape via Selection Pane and hadn't closed the
file) and recombinase's `inspect`/`generate` commands emitted a generic
"not a valid pptx file" message — user had to guess that the Windows file
lock was the cause. Now the error message says it directly.

### Deferred

- Table cell population and picture placeholder insertion remain on
  v0.1.12 (was v0.1.11 plan). Shipping a 4-test hot-fix to unblock the
  active Pack run took precedence over the bigger features.

## [0.1.10] - 2026-04-11

### Fixed

- **Bullet/paragraph formatting now preserved across list expansion.**
  v0.1.9's `_write_paragraphs` called `text_frame.clear()` which wiped all
  `<a:pPr>` elements, so only the first output paragraph inherited the
  template's bullet styling (via master-level defaults). Items 2..N came
  out as bare paragraphs — a 3-item background list rendered as one bullet
  followed by two un-bulleted lines. v0.1.10 captures the first existing
  paragraph's `<a:pPr>` AND the first run's `<a:rPr>` BEFORE clearing the
  text frame, then re-injects both into every new paragraph via a new
  `_apply_preserved_format` helper. Caught on the first real-template run
  where `background`, `education`, `languages`, and `key_competencies`
  all rendered with only one bullet each.

### Added

- 4 regression tests (`test_v0_1_10_bullet_preservation.py`) that verify
  pPr and rPr preservation across list expansion, including a save/reopen
  round-trip assertion. Tests build a synthetic template with an
  injected `marL` pPr marker and `sz` rPr marker, then assert every output
  paragraph carries the preserved attributes. 108 → 112 tests.

### Deferred

- Table cell population and picture placeholder insertion — now planned
  for v0.1.11. The bullet bug was a higher-priority hot-fix because it
  affected every bullet-list field in every CV output on v0.1.9, and the
  table/picture features only affect two shapes on the template.

## [0.1.9] - 2026-04-11

### Added

- **14 new robustness tests** (94 → 108) closing the remaining coverage
  gaps flagged by the v0.1.5 testing reviewer re-review:
  - `load_records` error branches: missing dir (`FileNotFoundError`),
    path-is-file (`NotADirectoryError`), non-dict YAML record
    (`ValueError`), silent-skip for empty/comment-only YAML files,
    and `.yml` extension globbing.
  - `remove_slide` direct test with slide-count and relationship-drop
    verification. Previously only covered indirectly.
  - `duplicate_slide` external-hyperlink rel preservation — rounds a real
    hyperlink through save/reopen and asserts the external rel survives.
  - `duplicate_slide` `notesSlide` skip — verifies source-slide presenter
    notes do NOT carry over to the duplicated slide.
  - `_walk_shapes` nested group recursion — group-inside-a-group template
    fixture, asserts every level is reachable.
  - `set_shape_value` with `bool` values (`True` → `"True"`, `False` →
    `"False"` — pin the bool-is-int-subclass gotcha).
  - `load_config` relative template path resolution against config dir.
  - `main()` `PermissionError` branch with the PowerPoint-lock hint —
    monkeypatches `Presentation.save` to force the error path.
  - `main()` `yaml.YAMLError` branch with clean error output.

### Why these and not the others

The testing reviewer's remaining-gap list also flagged Unicode field
values, `RECOMBINASE_DEBUG` env var, non-string placeholder type guards,
and `test_regressions.py` file splitting. Those are P2/P3 — they catch
edge cases that may never fire. The 14 above catch code paths that WILL
fire on a real CV template with a hyperlink, or a typo'd data directory,
or presenter notes, or a PowerPoint lock. Shipping the high-signal tests
and deferring the rest.

## [0.1.8] - 2026-04-11

### Added

- **`recombinase validate`** — pre-flight check that cross-references a
  config against its template. Verifies the config loads, the template
  exists, every configured shape name is actually present on the source
  slide, and reports any template shapes that aren't mapped. Surfaces the
  field name alongside missing shapes so you can see exactly which
  placeholder key maps to the broken shape name. Exit 0 on clean, 1 on
  missing shapes, 2 on unused shapes if `--strict`. Zero-arg auto-detect
  matches the other commands: run it inside a scaffolded project and it
  picks up `./template/config.yaml`.
- **`recombinase generate --dry-run`** — simulate the full pipeline
  (load config, load records, duplicate slides, set values, collect
  warnings) but do NOT write the output pptx. Ignores the overwrite
  guard since nothing is being written. Use this to preview overflow and
  missing-field warnings before committing to a real run.
- 11 new tests (83 → 94) covering validate happy path, missing-shape
  errors, unused-shape warnings, --strict escalation, auto-detect,
  dry-run output suppression, dry-run warning surfacing, dry-run
  overwrite-guard bypass.

### Why these two

From the v0.1.7 brainstorm's "11 possible features" list, these are the
only two that meaningfully bite before the first real run. The others
(extract mode, watch mode, PDF export, multi-template, etc.) depend on
usage-driven signal that doesn't exist yet. Shipping them speculatively
would violate the project's "simplify, don't engineer" discipline.
Revisit after the first real pack use.

## [0.1.7] - 2026-04-11

### Added

- **Zero-argument workflow inside a scaffolded project.** Every command now
  has sensible defaults so a user inside a `recombinase new` project folder
  can run:
  ```
  recombinase inspect     # auto-detects ./template/*.pptx or *.pptm
  recombinase init        # same, writes config to ./template/config.yaml
  recombinase generate    # uses ./template/config.yaml, ./cv-data/, ./output/deck.pptx
  ```
  No paths required. `recombinase new` → `cd <dir>` → three commands in
  sequence and you have a populated deck.
- **`_find_template_in_cwd()`** helper used by `inspect` and `init`. Searches
  `./template/*.pptm`, `./template/*.pptx`, `./*.pptm`, `./*.pptx` in that
  order. Returns the unique match or `None` — refuses to guess when
  multiple candidates exist.
- **`cmd_generate` default path resolution**: `-c` defaults to
  `./template/config.yaml`, `-d` defaults to `./cv-data/`, `-o` defaults to
  `./output/deck.pptx`. Friendly errors when the scaffolded paths are
  missing, pointing at the next command to run.
- **`cmd_init` smart output default**: if a `./template/` folder exists
  (scaffolded layout), the scaffold config is written to
  `./template/config.yaml`. Otherwise it falls back to
  `./template-config.yaml` in the current directory.
- 10 new tests for the auto-detection paths including end-to-end zero-arg
  `recombinase generate` inside a synthetic scaffolded project.

### Changed

- `inspect` and `init` template argument is now optional (was required).
  Passing an explicit path still works unchanged.

## [0.1.6] - 2026-04-11

### Added

- **Overflow detection heuristic in `generate_deck`.** For each populated
  field, recombinase now compares the new text length against the source
  slide's baseline and emits a warning when the ratio exceeds
  `overflow_ratio` (default `1.5`). Set `overflow_ratio: 0` in the config
  to disable. Caught by `--strict` to escalate to exit code 2.
- **Preset geometry detection in `inspect`.** Each `ShapeInfo` now carries
  a `preset_geom` field recording `<a:prstGeom prst="...">` values
  (`ellipse`, `roundRect`, `rect`, etc.). Surfaced in `format_template_info`
  output as `geom=ellipse` so circle-cropped profile pictures are
  discoverable without opening PowerPoint. The underlying bitmap is still
  a square — the preset is a display mask only.
- **`tests/conftest.py` with shared pytest fixtures**:
  `simple_template`, `rich_template`, `template_with_picture`,
  `template_with_group`, `tiny_png`, `sample_data_dir`, `write_config`.
  Reduces duplication across test files and makes future tests cheaper
  to write.
- **Regression test for the `r:id` rewrite branch in `duplicate_slide`**:
  verifies the rewrite actually fires via a round-trip save/reopen with
  `Picture.image.blob` access.

### Changed

- **`main()` exception handler refactored to a dispatch table**
  (`_ERROR_HANDLERS`). Six nearly-identical `except` blocks collapsed
  into one loop plus a data structure. Adding a new trapped exception
  class is now a one-line addition instead of a four-line block.
- **`_format_permission_error` and `_format_package_not_found`** take
  `BaseException` instead of specific subclasses, which satisfies the
  contravariant `Callable[[BaseException], str]` type the dispatch
  table expects.

### Fixed

- **Test fixture duplication surfaced by the v0.1.5 re-review**: new tests
  in `test_v0_1_6_features.py` use conftest fixtures directly.
  `test_end_to_end.py` and `test_regressions.py` still use their legacy
  inline helpers and will migrate in a future release — this was a
  deliberate scope choice to keep the v0.1.6 ship small.

### Deferred to v0.2

The v0.1.5 re-review surfaced several remaining items that are not
addressed in this release:

- CLI `inspect` / `init` / `new` runner smoke tests
- `generate_deck` with `records=[]` at the library layer
- `RECOMBINASE_DEBUG` env var documentation in README
- `cmd_init` default output path matching the scaffolded layout
- `_default_project_dir` missing `OneDriveConsumer`
- Nested group shape recursion test
- External hyperlink rel duplication test
- `notesSlide` skip test
- `test_regressions.py` split into focused files

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
