# CLAUDE.md

Project-specific instructions for Claude Code (and other AI coding assistants)
working in this repository.

## What this is

`recombinase` is a small, public PyPI package that populates PowerPoint
templates from structured YAML data. It is **personal infrastructure**,
authored by Terry Li for use in consulting credentials-pack workflows and
published under MIT license for anyone who wants it.

## Hard constraints

### No client data, ever

- **No real CV content**, no colleague names, no engagement details, no
  client templates committed to this repo.
- **No REDACTED-internal template files** — the package is tested against
  synthetic pptx files built programmatically inside the tests.
- Data processed by the tool lives **on the user's device**, not in the repo.
  The README explicitly documents a OneDrive-side project layout for this.

### No inline suppressions

This repo follows the broader vivesca rule: **never** use `# noqa`,
`# type: ignore`, `# pyright: ignore`, `# pragma: no cover`. If a lint rule
fires, fix the code or fix the config. The one exception is the
per-file-ignore entry in `pyproject.toml` for `B008` in `cli.py`, because
Typer's API *requires* function calls in argument defaults.

### The duplicate-and-populate invariant

The generator **must not** create new slides from a blank layout. It **must**
duplicate a known-good filled example slide from the template and overwrite
the text values in-place. This is the single most important design decision
in the package — it's why output decks inherit 100% of the template's visual
styling (fonts, colours, positions, master-slide elements) for free. See
`generate.duplicate_slide` and the accompanying tests.

If you're tempted to change this to create-from-layout, stop. The lxml
deep-copy approach looks clumsy but is correct.

## Dev commands

Run all of these before pushing:

```bash
ruff check src/ tests/
ruff format --check src/ tests/
mypy src/
pytest
```

All must pass. The `pre-commit` config enforces this on commit.

## Release process

1. Bump `version` in `pyproject.toml` AND `src/recombinase/__init__.py`
   (keep them in sync — there's no single source of truth yet, because
   that adds complexity we don't need).
2. Add an entry to `CHANGELOG.md` under a new version heading.
3. Run the dev commands above. Must all pass.
4. Build and publish:
   ```bash
   rm -rf dist/
   python -m build
   twine upload dist/*
   ```
5. Commit and push with a message like `vX.Y.Z: <one-line summary>`.

## Design rules

- **YAML, not CSV or JSON.** The data model has variable-length lists
  (background bullets, skills) that don't fit flat tables naturally and
  are hostile to edit in raw JSON. YAML is the sweet spot for human-edited
  structured data.
- **One file per record.** Each consultant / each use-case / each record
  is its own YAML file in a data directory. Keeps review, git history, and
  revision per-record.
- **Template config is external.** The package knows nothing about specific
  shape names. A per-template YAML config declares the mapping. Same package,
  different configs, different templates.
- **Typer over argparse/click.** Chosen for rich help output and human UX.
  Its "function calls in argument defaults" pattern is idiomatic; suppress
  B008 only in `cli.py`.
- **Warnings, not errors, on missing data.** `generate` emits warnings and
  keeps going rather than failing on the first missing field. Use `--strict`
  to escalate to non-zero exit.

## What NOT to build here

- Not a replacement for PowerPoint's native features.
- Not a general-purpose document automation tool.
- Not an LLM-assisted extractor — all processing stays local and deterministic.
- Not a service or API — CLI and library only.

## Non-goals for scope discipline

- No GUI
- No web interface
- No cloud/server component
- No account system
- No telemetry

If a future feature feels like it's heading toward any of these, stop and
revisit whether the core package should own it.
