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
from pptx.slide import Slide

from recombinase.config import TemplateConfig

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

    Empty input clears the text frame. Non-empty input reuses the existing
    single paragraph for the first item and appends new paragraphs for the
    rest, so bullet formatting from the template is preserved.
    """
    if not items:
        text_frame.clear()
        return
    text_frame.clear()
    # After clear(), text_frame has exactly one empty paragraph to reuse.
    text_frame.paragraphs[0].text = items[0]
    for item in items[1:]:
        text_frame.add_paragraph().text = item


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


def generate_deck(
    config: TemplateConfig,
    records: list[dict[str, Any]],
    output_path: Path | str,
) -> dict[str, Any]:
    """Generate an output pptx from a template + records.

    Returns a dict summary with counts and any warnings.
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
            try:
                set_shape_value(shape, record[field_name])
            except Exception as exc:
                warnings.append(
                    f"record {record_id!r}: failed to set {shape_name!r} ({field_name!r}): {exc}"
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
