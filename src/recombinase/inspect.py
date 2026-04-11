"""Inspect a pptx template and print structural metadata.

Prints only structural information (shape names, types, placeholder info,
text character counts). Does NOT print actual text content, so it's safe
to share the output publicly even if the source template contains draft
sample data.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn


@dataclass
class ShapeInfo:
    """Structural summary of a single shape (no text content)."""

    name: str
    shape_type: str
    is_placeholder: bool
    placeholder_type: str | None
    placeholder_idx: int | None
    text_chars: int | None
    paragraph_count: int | None
    run_count: int | None
    depth: int = 0
    """Nesting depth inside group shapes. 0 = top-level, 1 = inside a group, etc."""

    preset_geom: str | None = None
    """The `<a:prstGeom prst="...">` value if the shape uses a preset geometry
    (`ellipse`, `roundRect`, `rect`, etc.). `None` if the shape has no preset
    geometry or uses a custom geometry. Useful for detecting circle-cropped
    profile pictures — if a Picture shape has `preset_geom="ellipse"` the
    underlying image bitmap is still a square, it's just being rendered with
    a circle mask."""

    has_table: bool = False
    """True if the shape is a table shape (has a `shape.table` attribute).
    Used by `recombinase validate` to catch config errors where a text
    placeholder is pointed at a table shape or vice versa."""


@dataclass
class SlideInfo:
    """Structural summary of a single slide."""

    index: int
    layout_name: str
    shapes: list[ShapeInfo]


@dataclass
class TemplateInfo:
    """Structural summary of an entire pptx template."""

    path: Path
    slide_count: int
    slides: list[SlideInfo]
    layout_names: list[str]


def _detect_preset_geom(shape: Any) -> str | None:
    """Return the value of `<a:prstGeom prst="...">` for this shape, or None.

    Walks the shape's XML element looking for the first `prstGeom` descendant.
    Returns the preset name string (e.g. `"ellipse"`, `"roundRect"`, `"rect"`).
    Used by the inspector to flag circle-cropped profile pictures so users
    can see at a glance which shapes apply a mask to the underlying bitmap.
    """
    element = getattr(shape, "_element", None)
    if element is None:
        return None
    prst_geom = element.find(f".//{qn('a:prstGeom')}")
    if prst_geom is None:
        return None
    return prst_geom.get("prst")


def _shape_info(shape: Any, depth: int = 0) -> ShapeInfo:
    placeholder_type: str | None = None
    placeholder_idx: int | None = None
    if getattr(shape, "is_placeholder", False):
        ph = shape.placeholder_format
        placeholder_type = str(ph.type) if ph.type is not None else None
        placeholder_idx = ph.idx

    text_chars: int | None = None
    paragraph_count: int | None = None
    run_count: int | None = None
    if getattr(shape, "has_text_frame", False):
        tf = shape.text_frame
        text_chars = len(tf.text or "")
        paragraph_count = len(tf.paragraphs)
        run_count = sum(len(p.runs) for p in tf.paragraphs)

    return ShapeInfo(
        name=shape.name,
        shape_type=str(shape.shape_type),
        is_placeholder=bool(getattr(shape, "is_placeholder", False)),
        placeholder_type=placeholder_type,
        placeholder_idx=placeholder_idx,
        text_chars=text_chars,
        paragraph_count=paragraph_count,
        run_count=run_count,
        depth=depth,
        preset_geom=_detect_preset_geom(shape),
        has_table=bool(getattr(shape, "has_table", False)),
    )


def _collect_shapes(shapes: Any, depth: int = 0) -> list[ShapeInfo]:
    """Walk a shape collection recursively, descending into groups.

    Each shape is emitted once; group shapes are emitted before their
    children, with `depth` incremented for nested shapes so formatters
    can indent them.
    """
    result: list[ShapeInfo] = []
    for shape in shapes:
        result.append(_shape_info(shape, depth))
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            result.extend(_collect_shapes(shape.shapes, depth + 1))
    return result


def inspect_template(path: Path | str) -> TemplateInfo:
    """Inspect a pptx/pptm template and return its structural metadata."""
    path = Path(path).expanduser().resolve()
    prs = Presentation(str(path))

    slides: list[SlideInfo] = []
    for slide_index, slide in enumerate(prs.slides, start=1):
        shape_infos = _collect_shapes(slide.shapes)
        slides.append(
            SlideInfo(
                index=slide_index,
                layout_name=slide.slide_layout.name,
                shapes=shape_infos,
            )
        )

    return TemplateInfo(
        path=path,
        slide_count=len(prs.slides),
        slides=slides,
        layout_names=[layout.name for layout in prs.slide_layouts],
    )


def format_template_info(info: TemplateInfo) -> str:
    """Format a TemplateInfo for human-readable printing."""
    lines: list[str] = [
        f"File: {info.path}",
        f"Slide count: {info.slide_count}",
        "",
    ]

    for slide in info.slides:
        lines.append(f"=== Slide {slide.index} (layout: {slide.layout_name!r}) ===")
        if not slide.shapes:
            lines.append("  (no shapes)")
        for shape in slide.shapes:
            indent = "  " + ("  " * shape.depth)
            bits = [repr(shape.name), f"type={shape.shape_type}"]
            if shape.is_placeholder:
                bits.append(
                    f"placeholder(type={shape.placeholder_type}, idx={shape.placeholder_idx})"
                )
            if shape.text_chars is not None:
                bits.append(f"text_chars={shape.text_chars}")
                bits.append(f"paras={shape.paragraph_count}, runs={shape.run_count}")
            if shape.preset_geom is not None:
                bits.append(f"geom={shape.preset_geom}")
            if shape.depth > 0:
                bits.append(f"depth={shape.depth}")
            lines.append(indent + "- " + " | ".join(bits))
        lines.append("")

    lines.append("=== Slide Layouts available on master ===")
    for layout_index, layout_name in enumerate(info.layout_names):
        lines.append(f"  [{layout_index}] {layout_name!r}")

    return "\n".join(lines)


def shape_names_from_slide(info: TemplateInfo, slide_index: int) -> list[str]:
    """Return the list of shape names on a specific slide (1-based index).

    Includes shapes nested inside groups so scaffold configs written from
    this list cover every addressable shape on the slide.
    """
    for slide in info.slides:
        if slide.index == slide_index:
            return [shape.name for shape in slide.shapes]
    return []


def shape_types_from_slide(info: TemplateInfo, slide_index: int) -> dict[str, bool]:
    """Return a dict mapping shape name -> is_table for a specific slide.

    Used by ``recombinase validate`` to check that placeholders aren't
    pointed at table shapes and vice versa. Later shapes with the same
    name win — normally shapes have unique names on a slide anyway.
    """
    for slide in info.slides:
        if slide.index == slide_index:
            return {shape.name: shape.has_table for shape in slide.shapes}
    return {}
