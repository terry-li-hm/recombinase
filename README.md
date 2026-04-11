# cv-scripts

Personal toolkit for generating PowerPoint CV decks from structured data
(Excel, CSV, or YAML) into a pre-formatted pptx template.

## Scope

Generic tooling. No client content lives in this repo — no templates, no real
CV data, no colleague names. Scripts are designed to run against any
placeholder-based pptx template. Data stays on the device running the scripts.

## Scripts

- `inspect_pptx_shapes.py` — Print structural metadata of a pptx template
  (shape names, types, placeholder info, text character counts). Used to
  discover shape names before writing a generator mapping. Outputs no text
  content, only structure.

## Requirements

```
python -m pip install --user python-pptx
```

Optional for Excel input:

```
python -m pip install --user openpyxl
```

## Workflow

1. Obtain a pptx template with named placeholder shapes.
2. Run `inspect_pptx_shapes.py` against it to discover the shape names.
3. Prepare structured input data (Excel/CSV/YAML) with columns matching the
   placeholder names.
4. Run the generator script (to be added) to produce the populated output
   pptx in a new file.

## Non-goals

- Not a distributable package.
- Not a replacement for PowerPoint's native template features.
- Not for processing data that shouldn't leave the local machine via any
  channel — runs locally, reads locally, writes locally.
