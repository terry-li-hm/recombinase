"""Generate a populated pptx deck from a template + structured YAML data.

Pattern: duplicate a well-formatted source slide in the template once per
record, populate the duplicated slide's shapes by name, save the output.
Relies on the `source slide` in the template being already styled correctly
(the 'filled example' approach) so that visual fidelity is inherited for
free via slide duplication.
"""

from __future__ import annotations

from collections.abc import Iterator
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.placeholder import PicturePlaceholder
from pptx.slide import Slide

from recombinase.config import TableConfig, TemplateConfig

# OOXML namespace used by r:id / r:embed / r:link attributes inside shape XML.
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def load_records(data_dir: Path | str) -> list[dict[str, Any]]:
    """Load all YAML files from a directory as a list of records.

    Each .yaml or .yml file becomes one record. Files are loaded in
    sorted filename order so output is deterministic.
    """
    data_dir = Path(data_dir).expanduser().resolve()
    if not data_dir.exists():
        raise FileNotFoundError(f"Data directory not found: {data_dir}")
    if not data_dir.is_dir():
        raise NotADirectoryError(f"Not a directory: {data_dir}")

    records: list[dict[str, Any]] = []
    yaml_files = sorted([*data_dir.glob("*.yaml"), *data_dir.glob("*.yml")])
    for yaml_file in yaml_files:
        with yaml_file.open("r", encoding="utf-8") as fh:
            data = yaml.safe_load(fh)
        if data is None:
            continue
        if not isinstance(data, dict):
            raise ValueError(f"{yaml_file}: expected top-level mapping, got {type(data).__name__}")
        records.append(data)
    return records


def _walk_shapes(shapes: Any) -> Iterator[Any]:
    """Walk a shape collection, yielding every shape including group members.

    PowerPoint allows shapes to be nested inside GroupShape elements
    (`<p:grpSp>`), and iterating `slide.shapes` directly only yields the
    top-level children. This walker recurses into groups so every named
    shape is reachable by `find_shape_by_name` and visible to `inspect`.
    """
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk_shapes(shape.shapes)


def duplicate_slide(presentation: Any, source_slide: Slide) -> Slide:
    """Duplicate a slide, preserving shapes, formatting, and relationships.

    python-pptx has no native `slide.duplicate()` — this is the canonical
    workaround extended to also copy the source slide's relationships and
    rewrite `r:id` / `r:embed` / `r:link` references in the copied XML.
    Without the rel copy, deep-copied shapes referencing pictures, hyperlinks,
    or embedded charts produce dangling references on the new slide and
    those elements render as broken or empty.

    Ref: https://github.com/scanny/python-pptx/issues/132
    """
    blank_layout = source_slide.slide_layout
    new_slide = presentation.slides.add_slide(blank_layout)

    # Remove any shapes the layout added automatically — we want only the
    # source slide's shapes, not the layout's defaults.
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    # Copy relationships from the source slide onto the new slide, building an
    # old-rId -> new-rId map so that copied shape XML can be rewritten. Skip
    # notesSlide relationships — those belong to the notes subsystem and
    # copying them would attach the source's notes to the new slide.
    rel_id_map: dict[str, str] = {}
    for old_rid, rel in source_slide.part.rels.items():
        if "notesSlide" in rel.reltype:
            continue
        if rel.is_external:
            new_rid = new_slide.part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_rid = new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)
        rel_id_map[old_rid] = new_rid

    # Copy each top-level shape XML element from the source slide into the new
    # slide, rewriting any r: references so they resolve against the new rels.
    for shape in source_slide.shapes:
        new_el = deepcopy(shape._element)
        for descendant in new_el.iter():
            for attr in ("id", "embed", "link"):
                qname = f"{{{_R_NS}}}{attr}"
                val = descendant.get(qname)
                if val and val in rel_id_map and val != rel_id_map[val]:
                    descendant.set(qname, rel_id_map[val])
        new_slide.shapes._spTree.append(new_el)

    return new_slide


def find_shape_by_name(slide: Slide, name: str) -> Any | None:
    """Find a shape on a slide by its `.name` property, recursing into groups.

    Returns the first matching shape at any nesting depth, or None if no
    shape with that name exists on the slide. CV templates commonly group
    Name + Role + Headshot as a single unit, and without group recursion
    those named shapes would be unreachable here.
    """
    for shape in _walk_shapes(slide.shapes):
        if shape.name == name:
            return shape
    return None


def set_shape_value(shape: Any, value: Any) -> None:
    """Write a value into a shape's text frame.

    Handles:
    - None or empty string -> clear the text frame
    - list -> each item becomes a separate paragraph (bullet points inherit
      from the placeholder's paragraph-level formatting in the template)
    - scalar (str / int / float / any str-convertible) -> a single text
      value; strings containing `\\n` are split into paragraphs

    Caveat: setting text frame `.text` flattens rich-text runs within the
    shape to the placeholder's default run style. If a placeholder needs
    rich-text preservation (e.g. bold name + italic subtitle in one shape),
    consider splitting it into two separate shapes in the template.
    """
    if not getattr(shape, "has_text_frame", False):
        return

    tf = shape.text_frame

    if value is None or value == "":
        tf.clear()
        return

    if isinstance(value, list):
        items = [str(item) for item in value if item is not None and item != ""]
        _write_paragraphs(tf, items)
        return

    # Scalar: stringify first, then either split on newlines (paragraph form)
    # or write as a single text value.
    text = str(value)
    if "\n" in text:
        _write_paragraphs(tf, text.split("\n"))
        return

    tf.text = text


def _write_paragraphs(text_frame: Any, items: list[str]) -> None:
    """Write a list of strings as paragraphs into a text frame, one per item.

    Preserves paragraph-level formatting (bullet style, indent, alignment)
    by capturing the first existing paragraph's `<a:pPr>` before clearing
    the text frame, then re-injecting that pPr into every new paragraph.

    Without this preservation step, `text_frame.clear()` wipes all pPr and
    `add_paragraph()` produces bare paragraphs that lose the template's
    bullet styling — so a 3-item list ends up with one bullet on the first
    item and none on the rest. This is the v0.1.10 fix.

    Empty input clears the text frame. Non-empty input keeps the first
    existing paragraph's style profile for every output paragraph.
    """
    from pptx.oxml.ns import qn

    if not items:
        text_frame.clear()
        return

    # Capture the first existing paragraph's pPr + first run's rPr BEFORE
    # clear() wipes them. These encode the template's bullet/indent/font style.
    first_pPr_copy = None
    first_rPr_copy = None
    existing_paras = text_frame.paragraphs
    if existing_paras:
        first_p_xml = existing_paras[0]._p
        first_pPr = first_p_xml.find(qn("a:pPr"))
        if first_pPr is not None:
            first_pPr_copy = deepcopy(first_pPr)
        # Also capture the first run's rPr (font size, color, weight, etc.)
        first_r = first_p_xml.find(qn("a:r"))
        if first_r is not None:
            first_rPr = first_r.find(qn("a:rPr"))
            if first_rPr is not None:
                first_rPr_copy = deepcopy(first_rPr)

    text_frame.clear()

    # After clear(), text_frame has exactly one empty paragraph to reuse.
    first_p = text_frame.paragraphs[0]
    first_p.text = items[0]
    _apply_preserved_format(first_p, first_pPr_copy, first_rPr_copy)

    for item in items[1:]:
        new_p = text_frame.add_paragraph()
        new_p.text = item
        _apply_preserved_format(new_p, first_pPr_copy, first_rPr_copy)


def _apply_preserved_format(paragraph: Any, pPr_copy: Any, rPr_copy: Any) -> None:
    """Inject a preserved pPr and the first run's rPr into a paragraph.

    pPr is inserted at the start of the paragraph element (OOXML requires
    pPr to precede text runs). rPr is injected at the start of the
    paragraph's first run.
    """
    from pptx.oxml.ns import qn

    if pPr_copy is not None:
        # Remove any existing pPr (added by add_paragraph) then insert ours
        existing_pPr = paragraph._p.find(qn("a:pPr"))
        if existing_pPr is not None:
            paragraph._p.remove(existing_pPr)
        paragraph._p.insert(0, deepcopy(pPr_copy))

    if rPr_copy is not None:
        first_r = paragraph._p.find(qn("a:r"))
        if first_r is not None:
            existing_rPr = first_r.find(qn("a:rPr"))
            if existing_rPr is not None:
                first_r.remove(existing_rPr)
            first_r.insert(0, deepcopy(rPr_copy))


def is_picture_placeholder(shape: Any) -> bool:
    """Return True if the shape is a python-pptx PicturePlaceholder.

    PicturePlaceholders have an `insert_picture(path)` method to embed an
    image into the placeholder slot, keeping the shape's crop geometry and
    position. Regular Picture shapes (inserted via `add_picture`) are a
    different class and need different handling.
    """
    return isinstance(shape, PicturePlaceholder)


def set_picture(shape: Any, value: Any, base_dir: Path | None = None) -> None:
    """Insert an image file into a PicturePlaceholder shape.

    `value` is interpreted as a path to an image file. If it's a relative
    path and `base_dir` is provided, it resolves against `base_dir` (the
    directory of the YAML record file, typically). Otherwise the path is
    used as-is.

    Empty string or None is treated as "no picture" and silently no-ops —
    the placeholder retains whatever it already had (from the source slide
    inheritance), so leaving a `photo:` field empty keeps the example
    headshot as a fallback.
    """
    if value is None or value == "":
        return

    image_path = Path(str(value)).expanduser()
    if not image_path.is_absolute() and base_dir is not None:
        image_path = (base_dir / image_path).resolve()

    if not image_path.exists():
        raise FileNotFoundError(f"picture file not found: {image_path}")

    shape.insert_picture(str(image_path))


def populate_table(shape: Any, table_config: TableConfig, rows: list[Any]) -> list[str]:
    """Populate a table shape from a list of row dicts.

    Each element of `rows` is a dict whose keys should match the names in
    `table_config.columns`. Column values can be scalars or lists; lists
    are joined with `table_config.list_joiner` to produce multi-line cell
    text that inherits the template's paragraph-level bullet formatting.

    Returns a list of warning strings (empty if everything populated cleanly).
    Common warnings:
    - data has more rows than the template table can hold (truncation)
    - a row dict is missing a column key (that cell is left blank)
    - the shape isn't actually a table (recombinase is called on the wrong shape)
    """
    warnings: list[str] = []

    if not getattr(shape, "has_table", False):
        warnings.append(
            f"shape {shape.name!r} is configured as a table but is not a "
            "table shape on the template; skipping"
        )
        return warnings

    table = shape.table
    all_rows = list(table.rows)
    start_row = 1 if table_config.header_row else 0
    data_rows = all_rows[start_row:]

    if len(rows) > len(data_rows):
        warnings.append(
            f"table {shape.name!r} has {len(data_rows)} data rows but "
            f"{len(rows)} records were provided; truncating to {len(data_rows)}"
        )

    for row_index, row_dict in enumerate(rows[: len(data_rows)]):
        target_row = data_rows[row_index]
        for col_index, column_name in enumerate(table_config.columns):
            if col_index >= len(target_row.cells):
                warnings.append(
                    f"table {shape.name!r} row {row_index} has only "
                    f"{len(target_row.cells)} cells; column {column_name!r} "
                    "exceeds the template width, skipping"
                )
                continue
            cell = target_row.cells[col_index]
            if not isinstance(row_dict, dict):
                warnings.append(
                    f"table {shape.name!r} row {row_index}: expected a "
                    f"dict, got {type(row_dict).__name__}; skipping row"
                )
                break
            value = row_dict.get(column_name)
            if value is None:
                continue
            if isinstance(value, list):
                items = [str(item) for item in value if item is not None and item != ""]
                text = table_config.list_joiner.join(items)
            else:
                text = str(value)
            # Preserve cell's existing paragraph formatting just like
            # _write_paragraphs does for placeholders.
            if "\n" in text:
                _write_paragraphs(cell.text_frame, text.split("\n"))
            else:
                cell.text_frame.text = text

    return warnings


def remove_slide(presentation: Any, slide: Slide) -> None:
    """Remove a slide from the presentation.

    python-pptx has no native `slides.remove()` — this walks the slide id
    list on the presentation part, finds the relationship whose target
    matches the target slide, and drops both the rel and the id list entry.

    Ref: https://github.com/scanny/python-pptx/issues/67
    """
    slide_id = slide.slide_id
    xml_slides = presentation.slides._sldIdLst

    for sl in xml_slides:
        rid = sl.get(f"{{{_R_NS}}}id")
        if rid is None:
            continue
        slide_from_rel = presentation.part.related_part(rid).slide
        if slide_from_rel.slide_id == slide_id:
            presentation.part.drop_rel(rid)
            xml_slides.remove(sl)
            return


def _value_char_length(value: Any) -> int:
    """Estimate the text-frame character length a value would occupy."""
    if value is None or value == "":
        return 0
    if isinstance(value, list):
        items = [str(item) for item in value if item is not None and item != ""]
        if not items:
            return 0
        return sum(len(item) for item in items) + (len(items) - 1)  # inter-para newlines
    return len(str(value))


def _capture_baseline_lengths(
    source_slide: Slide, placeholder_map: dict[str, str]
) -> dict[str, int]:
    """Record the text length of each configured placeholder on the source slide.

    Returned dict maps data field name to the source's current character count.
    Shapes without a text frame are skipped. These baselines are compared
    against per-record values to flag probable overflow.
    """
    baselines: dict[str, int] = {}
    for field_name, shape_name in placeholder_map.items():
        shape = find_shape_by_name(source_slide, shape_name)
        if shape is None or not getattr(shape, "has_text_frame", False):
            continue
        baselines[field_name] = len(shape.text_frame.text or "")
    return baselines


def generate_deck(
    config: TemplateConfig,
    records: list[dict[str, Any]],
    output_path: Path | str,
) -> dict[str, Any]:
    """Generate an output pptx from a template + records.

    Returns a dict summary with counts and any warnings. Warnings include
    missing shapes, missing fields, set-value exceptions, and — when
    `config.overflow_ratio > 0` — probable overflow cases where a record's
    field is substantially larger than the source-slide baseline.
    """
    output_path = Path(output_path).expanduser().resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    presentation = Presentation(str(config.template))

    total_slides = len(presentation.slides)
    if config.source_slide_index < 1 or config.source_slide_index > total_slides:
        raise ValueError(
            f"source_slide_index={config.source_slide_index} out of range "
            f"(template has {total_slides} slide(s))"
        )

    # 1-based -> 0-based
    source_slide = presentation.slides[config.source_slide_index - 1]

    # Capture baseline text lengths BEFORE generation so overflow comparison
    # is against the original example, not the most recently populated slide.
    baselines: dict[str, int] = (
        _capture_baseline_lengths(source_slide, config.placeholders)
        if config.overflow_ratio > 0
        else {}
    )

    warnings: list[str] = []
    generated_count = 0

    for record_index, record in enumerate(records, start=1):
        record_id = record.get("id") or record.get("name") or f"record_{record_index}"
        new_slide = duplicate_slide(presentation, source_slide)

        for field_name, shape_name in config.placeholders.items():
            shape = find_shape_by_name(new_slide, shape_name)
            if shape is None:
                warnings.append(
                    f"record {record_id!r}: shape {shape_name!r} not found on new slide"
                )
                continue
            if field_name not in record:
                warnings.append(f"record {record_id!r}: no value for field {field_name!r}")
                continue

            value = record[field_name]

            # Picture placeholder: route to insert_picture instead of setting text
            if is_picture_placeholder(shape):
                try:
                    record_source = record.get("_recombinase_record_dir")
                    base_dir = Path(record_source) if isinstance(record_source, str) else None
                    set_picture(shape, value, base_dir=base_dir)
                except FileNotFoundError as exc:
                    warnings.append(
                        f"record {record_id!r}: picture for {field_name!r} not found: {exc}"
                    )
                except Exception as exc:
                    warnings.append(
                        f"record {record_id!r}: failed to insert picture "
                        f"{shape_name!r} ({field_name!r}): {exc}"
                    )
                continue

            # Text placeholder: existing path
            try:
                set_shape_value(shape, value)
            except Exception as exc:
                warnings.append(
                    f"record {record_id!r}: failed to set {shape_name!r} ({field_name!r}): {exc}"
                )
                continue

            # Overflow heuristic: compare new text length against source baseline.
            baseline = baselines.get(field_name, 0)
            if baseline > 0:
                new_length = _value_char_length(value)
                if new_length > baseline * config.overflow_ratio:
                    ratio = new_length / baseline
                    warnings.append(
                        f"record {record_id!r}: field {field_name!r} is "
                        f"{ratio:.1f}x the source baseline ({new_length} vs "
                        f"{baseline} chars) — may overflow shape {shape_name!r}"
                    )

        # Tables: populate from config.tables using the record's list-of-dicts data
        for table_field_name, table_config in config.tables.items():
            table_shape = find_shape_by_name(new_slide, table_config.shape)
            if table_shape is None:
                warnings.append(
                    f"record {record_id!r}: table shape {table_config.shape!r} "
                    "not found on new slide"
                )
                continue
            if table_field_name not in record:
                warnings.append(f"record {record_id!r}: no value for table {table_field_name!r}")
                continue
            table_value = record[table_field_name]
            if not isinstance(table_value, list):
                warnings.append(
                    f"record {record_id!r}: table {table_field_name!r} expects "
                    f"a list of row dicts, got {type(table_value).__name__}"
                )
                continue
            try:
                table_warnings = populate_table(table_shape, table_config, table_value)
            except Exception as exc:
                warnings.append(
                    f"record {record_id!r}: failed to populate table "
                    f"{table_config.shape!r} ({table_field_name!r}): {exc}"
                )
                continue
            warnings.extend(
                f"record {record_id!r}: {tw}" for tw in table_warnings
            )

        generated_count += 1

    if config.clear_source_slide:
        remove_slide(presentation, source_slide)

    presentation.save(str(output_path))

    return {
        "output": str(output_path),
        "records_generated": generated_count,
        "warnings": warnings,
        "generated_at": datetime.now().isoformat(timespec="seconds"),
    }
