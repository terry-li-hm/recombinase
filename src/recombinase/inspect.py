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
