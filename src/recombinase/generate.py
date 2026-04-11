"""Generate a populated pptx deck from a template + structured YAML data.

Pattern: duplicate a well-formatted source slide in the template once per
record, populate the duplicated slide's shapes by name, save the output.
Relies on the `source slide` in the template being already styled correctly
(the 'filled example' approach) so that visual fidelity is inherited for
free via slide duplication.
"""

from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml
from pptx import Presentation
from pptx.slide import Slide

from recombinase.config import TemplateConfig


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
        data.setdefault("_source_file", str(yaml_file))
        records.append(data)
    return records


def duplicate_slide(presentation: Any, source_slide: Slide) -> Slide:
    """Duplicate a slide within a presentation, preserving all shapes and formatting.

    python-pptx does not provide a native slide.duplicate() — this is the
    canonical workaround using lxml deep-copy of the shape tree.

    Ref: https://github.com/scanny/python-pptx/issues/132
    """
    blank_layout = source_slide.slide_layout
    new_slide = presentation.slides.add_slide(blank_layout)

    # Remove any shapes the layout added automatically — we want only the
    # source slide's shapes, not the layout's defaults.
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)

    # Copy each shape XML element from the source slide into the new slide.
    for shape in source_slide.shapes:
        el = shape._element
        new_el = deepcopy(el)
        new_slide.shapes._spTree.append(new_el)

    return new_slide


def find_shape_by_name(slide: Slide, name: str) -> Any | None:
    """Find a shape on a slide by its .name property. Returns None if missing."""
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def set_shape_value(shape: Any, value: Any) -> None:
    """Write a value into a shape's text frame.

    Handles three cases:
    - scalar string → single text value, replaces whatever was there
    - list of strings → each item becomes a separate paragraph (bullet point)
    - None or empty → clear the text frame

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
        # Variable-length list → one paragraph per item (bullets inherit style
        # from the placeholder's paragraph-level formatting).
        items = [str(item) for item in value if item is not None]
        if not items:
            tf.clear()
            return
        tf.clear()
        # After clear(), text_frame has exactly one empty paragraph.
        first_p = tf.paragraphs[0]
        first_p.text = items[0]
        for item in items[1:]:
            p = tf.add_paragraph()
            p.text = item
        return

    if isinstance(value, (int, float)):
        value = str(value)

    # Scalar string: may contain newlines, which we interpret as paragraphs.
    text = str(value)
    if "\n" in text:
        set_shape_value(shape, text.split("\n"))
        return

    tf.text = text


def remove_slide(presentation: Any, slide: Slide) -> None:
    """Remove a slide from the presentation (python-pptx has no native API).

    Ref: https://github.com/scanny/python-pptx/issues/67
    """
    slide_id = slide.slide_id
    xml_slides = presentation.slides._sldIdLst
    slides_list = list(xml_slides)

    for sl in slides_list:
        rid = sl.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
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

    # 1-based → 0-based
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
