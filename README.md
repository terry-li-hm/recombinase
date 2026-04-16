# recombinase

> Biology: a recombinase is an enzyme that extracts DNA fragments and recombines them into new molecules using a homologous template as the structural guide. This package does the same for PowerPoint documents.

Template-guided pptx synthesis. Take a styled "filled example" slide in a `.pptx`/`.pptm` template, a folder of per-record YAML data files, and produce a populated output deck — one slide per record, visually identical to the template because the fill operation duplicates the source slide and replaces text in-place by shape name.

## Install

```
pip install --user recombinase
```

Dependencies: `python-pptx`, `pyyaml`. That's it.

### Windows 11 install notes

On a standard Windows 11 machine with Python installed:

```powershell
python -m pip install --user recombinase
```

The `--user` flag installs into your user profile (`%APPDATA%\Python\Python3XX\site-packages`) which doesn't require admin rights and usually slips past managed-device policy that blocks system-wide installs. After install, the `recombinase` command is on your PATH at `%APPDATA%\Python\Python3XX\Scripts\`. If `recombinase --version` isn't found after install, use `python -m recombinase.cli --version` as a fallback.

### From source (dev)

```
git clone https://github.com/terry-li-hm/recombinase.git
cd recombinase
pip install --user -e ".[dev]"
```

## Concepts

**Three steps**, loosely coupled by file:

1. **Template** — a `.pptx`/`.pptm` file with at least one slide where every field you want to populate is a named shape (e.g. a text box named `Consultant_Name`).
2. **Config** — a small YAML file declaring which shape name on the template corresponds to which data field. One config per template. Templates change; configs go with them.
3. **Data** — a directory of per-record YAML files, one file per record (e.g. one per consultant). Each file has a flat map of field names to values. List values become bullet paragraphs automatically.

The template config is intentionally per-template rather than hardcoded in the package. Same package, different config → different template. You can build a CV pack, a use-case slide deck, or a client case study collection with the same tool and three different configs.

## Usage

### 0. Scaffold a project (optional, Windows/OneDrive friendly)

Before touching any template, `recombinase new` creates a conventional folder layout you can drop your pptx into:

```
recombinase new              # defaults to %OneDrive%/cv, or ~/cv
recombinase new C:\Work\Pack # explicit path
```

Result:

```
<project>/
├── README.md
├── template/   # drop your .pptm/.pptx here
├── data/    # one YAML file per record (consultant, use case, etc.)
└── output/     # generated decks land here
```

Then `cd` into the scaffolded directory and use the remaining commands with relative paths.

### 1. Inspect a template

Discover the shape names on each slide — structural metadata only, never the actual text content. Safe to share the output.

```
recombinase inspect "path/to/template.pptm"
```

Example output:

```
File: /path/to/template.pptm
Slide count: 1

=== Slide 1 (layout: 'Blank') ===
  - 'Consultant_Name' | type=TEXT_BOX (17) | text_chars=12 | paras=1, runs=1
  - 'Role_Title' | type=TEXT_BOX (17) | text_chars=18 | paras=1, runs=1
  - 'Summary_Body' | type=TEXT_BOX (17) | text_chars=140 | paras=2, runs=3
  - 'Background_Bullets' | type=TEXT_BOX (17) | text_chars=220 | paras=5, runs=5
```

### 2. Scaffold a config

Generate a starter config file from the template's shape names:

```
recombinase init "path/to/template.pptm" --output template-config.yaml
```

This reads the shape names from slide 1 and writes a config like:

```yaml
template: /path/to/template.pptm
source_slide_index: 1
clear_source_slide: true

placeholders:
  consultant_name: Consultant_Name
  role_title: Role_Title
  summary_body: Summary_Body
  background_bullets: Background_Bullets
```

Edit the left side (data field names) to match how your records are keyed. For example, if your YAML data files use `name:` not `consultant_name:`, rename the left side:

```yaml
placeholders:
  name: Consultant_Name
  role: Role_Title
  summary: Summary_Body
  background: Background_Bullets
```

### 3. Write per-record data files

Create a directory with one YAML file per record. Filenames become the sort order:

```
data/
├── 01-jane-doe.yaml
├── 02-john-smith.yaml
└── 03-alice-wong.yaml
```

Each file is a flat map. How list values are handled depends on the template shape:

- **Single-run shape** (e.g. a bullet list): each item becomes a separate paragraph
- **Multi-run shape** (e.g. mixed bold + grey in one line): each item replaces the corresponding run, preserving its formatting

```yaml
id: jane-doe
name: Jane Doe
role: Senior Consultant
summary: >-
  Twelve years across global wealth management with a focus on
  regulatory data and risk modelling.
background:
  - Bank A — Risk modelling lead (2010-2015)
  - Bank B — Head of analytics (2015-2020)
  - Bank C — CDO, Asia Pacific (2020-present)
key_skills:
  - Risk modelling
  - Governance
  - Wealth data architecture

# Rich text: list items map to runs in multi-run shapes
header:
  - "Bold part of the title. "
  - "Grey subtitle part."
```

The field names on the left must match the keys in your template config's `placeholders:` section.

### 4. Generate the output deck

```
recombinase generate \
  --config template-config.yaml \
  --data-dir data/ \
  --output output/deck.pptx
```

Produces a populated pptx with one slide per YAML file. If `clear_source_slide: true` in the config, the original example slide is removed from the output.

### One-line end-to-end

After the config exists:

```
recombinase generate -c template-config.yaml -d data/ -o out.pptx
```

## Design notes

### Why duplicate a filled example slide?

The alternative is creating slides from a layout and writing text into empty placeholders. That approach loses any hand-tweaks the template designer made (custom colours, tweaked positions, decorative shapes, non-placeholder elements). Duplicating a known-good filled slide inherits 100% of its visual styling by design — `deepcopy` of the shape tree carries every property.

Trade-off: the template must contain one "canonical good example" slide to clone from. This is usually natural for CV templates and pack-prep work.

### Rich text and the flattening caveat

When a value is written into a shape with `shape.text_frame.text = "..."`, rich-text runs within that shape (bold name, italic subtitle in one text frame) collapse to the placeholder's default run style. For most modern consulting templates this isn't an issue — each styled fragment lives in its own shape. If your template has a multi-run placeholder, either split it into separate shapes or accept the flattening.

### Variable-length lists

List values in the YAML data become separate paragraphs in the target text frame, inheriting the placeholder's paragraph-level bullet formatting automatically. No bullet markers in the source data — the template supplies them. A consultant with three background bullets and another with seven both work without any config change.

### Warnings, not errors

If a config references a shape name that doesn't exist, or a record is missing a field, `generate` produces a **warning** but continues. This is deliberate: partial output is more useful than total failure during iteration. Pass `--strict` if you want non-zero exit on warnings.

## Recommended file layout (Windows + OneDrive)

For work that involves colleague personal data (CVs, HR records, etc.), keep the package install and the data separate:

```
%USERPROFILE%\AppData\Roaming\Python\...\   # package install (local, not synced)
C:\Users\<you>\OneDrive - <Org>\Pack\       # work product (synced, backed up)
├── template\
│   ├── CV_template.pptm                   # the canonical pptx
│   └── template-config.yaml               # mapping shape names → data fields
├── data\                                # per-consultant YAML records
│   ├── 01-jane-doe.yaml
│   ├── 02-john-smith.yaml
│   └── 03-alice-wong.yaml
└── output\
    └── deck.pptx                           # generated
```

The package itself is generic tooling and lives in your Python site-packages — it has no opinions about any particular template or data set. The work data and configs live in OneDrive where they're backed up, versioned by OneDrive's history, and remain inside your organization's sanctioned storage. Run `recombinase` commands with full paths pointing at OneDrive locations:

```powershell
cd "C:\Users\<you>\OneDrive - <Org>\Pack"
recombinase inspect "template\CV_template.pptm"
recombinase init "template\CV_template.pptm" -o "template\template-config.yaml"
recombinase generate -c "template\template-config.yaml" -d "data" -o "output\deck.pptx"
```

## Development

```
git clone https://github.com/terry-li-hm/recombinase.git
cd recombinase
pip install -e ".[dev]"
ruff check src/ tests/
ruff format src/ tests/
mypy src/
pytest
```

## Scope (v0.1)

- ✓ Inspect: print template structural metadata
- ✓ Init: scaffold a config from shape names
- ✓ Generate: populate template from YAML records
- ✗ Extract: reverse direction (pptx → YAML) — v0.2 — needs a sample source file for structure before it can be implemented reliably

## License

MIT
