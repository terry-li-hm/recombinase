"""Template configuration: load and validate a per-template YAML config."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass
class TableConfig:
    """Configuration for populating a table shape from a list of record rows.

    One `TableConfig` per table shape in the template. The record's data
    for this table is expected to be a list of dicts, one dict per row,
    with keys matching the `columns` list. Each column value can be a
    scalar (written as single-line cell text) or a list (joined with
    `list_joiner` — defaults to newline — for bullet-style cells).
    """

    shape: str
    """The `.Name` of the table shape on the source slide."""

    columns: list[str] = field(default_factory=list)
    """Ordered list of data field names, one per table column."""

    header_row: bool = True
    """If True, row 0 of the table is treated as a static header and
    skipped during population. Data rows start at row 1."""

    footer_rows: int = 0
    """Number of trailing rows to treat as static footer (e.g. a totals
    row or a notes row at the bottom of the table). These rows are never
    populated and never cleared — their example content is preserved
    through the duplicate-and-populate cycle. Default 0 means no footer.

    When set, populate_table treats only ``rows[start_row : -footer_rows]``
    as data rows. Excess-row clearing still runs but stops before the
    footer region, so static totals/notes survive the clear pass.
    """

    list_joiner: str = "\n"
    """String used to join list values into cell text. Default is newline,
    which renders as multiple bullets per cell when the template cell has
    paragraph-level bullet formatting."""


@dataclass
class SectionConfig:
    """Configuration for populating a "sectioned list" shape.

    A sectioned list is a single text frame that renders N groups, each
    with one header paragraph followed by a variable-length list of
    bullet paragraphs. Real-world example: a CV "Key Competencies" cell
    with four sections — FS Industry, Functional, Technical, Methodical
    — where each section header uses a distinct bold/unbulleted style
    and the items below it inherit a bulleted list style.

    One `SectionConfig` per sectioned-list shape. The record's data for
    this field is expected to be a list of dicts, each with keys
    ``header`` (string) and ``items`` (list of strings).

    Rendering strategy: the template cell must already contain at least
    two styled example paragraphs — one at ``header_index`` showing the
    header profile (pPr + run rPr), and one at ``bullet_index`` showing
    the bullet profile. `populate_sections` captures both profiles
    before clearing, then emits paragraphs alternating between them in
    the order prescribed by the record data.
    """

    shape: str
    """The `.Name` of the text frame shape on the source slide."""

    header_index: int = 0
    """Zero-based index of the template paragraph whose pPr + first-run
    rPr define the *header* profile. Defaults to 0 (first paragraph)."""

    bullet_index: int = 1
    """Zero-based index of the template paragraph whose pPr + first-run
    rPr define the *bullet* profile. Defaults to 1 (second paragraph)."""


@dataclass
class TemplateConfig:
    """Per-template configuration declaring shape name <-> data field mapping.

    One config file per pptx template. The config decouples the template's
    internal shape names (which vary per template) from the data fields
    (which stay consistent across templates).
    """

    template: Path
    """Path to the source pptx/pptm template."""

    source_slide_index: int = 1
    """1-based index of the slide inside the template that should be
    duplicated per record. Usually the 'filled example' slide."""

    placeholders: dict[str, str] = field(default_factory=dict)
    """Mapping from data field name (key) to shape .Name in the template (value).

    Example:
        placeholders:
          name: Consultant_Name
          role: Role_Title
          background: Background_Bullets

    Picture placeholders work automatically — if the shape on the template
    is a PicturePlaceholder and the data value is a file path, recombinase
    calls `shape.insert_picture(path)` instead of setting text.
    """

    tables: dict[str, TableConfig] = field(default_factory=dict)
    """Mapping from data field name to a `TableConfig`. Used when a record's
    data for a field is a list of dicts and the target shape is a table."""

    sections: dict[str, SectionConfig] = field(default_factory=dict)
    """Mapping from data field name to a `SectionConfig`. Used when a record's
    data for a field is a list of {header, items} dicts and the target
    shape is a single text frame that should render as a sectioned list
    (header paragraph followed by bullet paragraphs, repeated per section).
    Real-world example: the CV "Key Competencies" cell with four named
    sections — FS Industry, Functional, Technical, Methodical."""

    clear_source_slide: bool = True
    """If True, the source (example) slide is removed from the final output
    so only the generated per-record slides remain."""

    overflow_ratio: float = 1.5
    """If a populated shape's text length is greater than this multiple of
    the source-example baseline, emit an overflow warning. Set to 0 to
    disable overflow detection entirely."""

    def validate(self) -> list[str]:
        """Return a list of validation error messages, or empty list if OK."""
        errors: list[str] = []
        if not self.template.exists():
            errors.append(f"Template file not found: {self.template}")
        if self.source_slide_index < 1:
            errors.append(f"source_slide_index must be >= 1 (got {self.source_slide_index})")
        if not self.placeholders and not self.tables and not self.sections:
            errors.append(
                "config has no placeholders, tables, or sections — at least "
                "one of them must be populated, otherwise recombinase has "
                "nothing to do"
            )
        for field_name, section in self.sections.items():
            if section.header_index < 0:
                errors.append(
                    f"sections.{field_name}.header_index must be >= 0 (got {section.header_index})"
                )
            if section.bullet_index < 0:
                errors.append(
                    f"sections.{field_name}.bullet_index must be >= 0 (got {section.bullet_index})"
                )
            if section.header_index == section.bullet_index:
                errors.append(
                    f"sections.{field_name}.header_index and bullet_index "
                    f"must differ (both are {section.header_index}); the two "
                    "profiles need to come from distinct template paragraphs"
                )
        return errors


def load_config(path: Path | str) -> TemplateConfig:
    """Load a TemplateConfig from a YAML file.

    Resolves the `template` path relative to the config file's directory
    unless it's absolute. Raises `ValueError` with a file-path prefix on
    any malformed or missing field instead of letting a raw TypeError or
    AttributeError escape.
    """
    path = Path(path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")

    with path.open("r", encoding="utf-8") as fh:
        try:
            raw = yaml.safe_load(fh)
        except yaml.YAMLError as exc:
            raise ValueError(f"{path}: invalid YAML: {exc}") from exc

    if raw is None:
        raise ValueError(f"{path}: config file is empty")
    if not isinstance(raw, dict):
        raise ValueError(f"{path}: expected top-level mapping, got {type(raw).__name__}")

    data: dict[str, Any] = raw

    template_raw = data.get("template")
    if template_raw is None:
        raise ValueError(f"{path}: missing required key 'template'")
    if not isinstance(template_raw, str):
        raise ValueError(f"{path}: 'template' must be a string, got {type(template_raw).__name__}")

    template_path = Path(template_raw).expanduser()
    if not template_path.is_absolute():
        template_path = (path.parent / template_path).resolve()

    source_slide_index_raw = data.get("source_slide_index", 1)
    if not isinstance(source_slide_index_raw, int) or isinstance(source_slide_index_raw, bool):
        raise ValueError(
            f"{path}: 'source_slide_index' must be an integer, got "
            f"{type(source_slide_index_raw).__name__}"
        )

    placeholders_raw = data.get("placeholders")
    if placeholders_raw is None:
        placeholders_raw = {}
    # Explicit None handling above — using `or {}` would silently coerce wrong
    # types like `[]` or `0` into an empty dict and hide real config errors.
    if not isinstance(placeholders_raw, dict):
        raise ValueError(
            f"{path}: 'placeholders' must be a mapping, got {type(placeholders_raw).__name__}"
        )
    for key, value in placeholders_raw.items():
        if not isinstance(key, str) or not isinstance(value, str):
            raise ValueError(
                f"{path}: 'placeholders' entries must be string -> string, got "
                f"{type(key).__name__} -> {type(value).__name__}"
            )

    clear_source_slide_raw = data.get("clear_source_slide", True)
    if not isinstance(clear_source_slide_raw, bool):
        raise ValueError(
            f"{path}: 'clear_source_slide' must be a boolean, got "
            f"{type(clear_source_slide_raw).__name__}"
        )

    overflow_ratio_raw = data.get("overflow_ratio", 1.5)
    if isinstance(overflow_ratio_raw, bool) or not isinstance(overflow_ratio_raw, (int, float)):
        raise ValueError(
            f"{path}: 'overflow_ratio' must be a number, got {type(overflow_ratio_raw).__name__}"
        )
    if overflow_ratio_raw < 0:
        raise ValueError(f"{path}: 'overflow_ratio' must be >= 0 (got {overflow_ratio_raw})")

    tables_raw = data.get("tables")
    if tables_raw is None:
        tables_raw = {}
    if not isinstance(tables_raw, dict):
        raise ValueError(f"{path}: 'tables' must be a mapping, got {type(tables_raw).__name__}")
    tables: dict[str, TableConfig] = {}
    for field_name, table_data in tables_raw.items():
        if not isinstance(field_name, str):
            raise ValueError(
                f"{path}: 'tables' keys must be strings, got {type(field_name).__name__}"
            )
        if not isinstance(table_data, dict):
            raise ValueError(
                f"{path}: 'tables.{field_name}' must be a mapping, got {type(table_data).__name__}"
            )
        shape_name_raw = table_data.get("shape")
        if not isinstance(shape_name_raw, str):
            raise ValueError(f"{path}: 'tables.{field_name}.shape' must be a string")
        columns_raw = table_data.get("columns")
        if columns_raw is None:
            columns_raw = []
        if not isinstance(columns_raw, list) or not all(
            isinstance(col, str) for col in columns_raw
        ):
            raise ValueError(f"{path}: 'tables.{field_name}.columns' must be a list of strings")
        header_row_raw = table_data.get("header_row", True)
        if not isinstance(header_row_raw, bool):
            raise ValueError(f"{path}: 'tables.{field_name}.header_row' must be a boolean")
        footer_rows_raw = table_data.get("footer_rows", 0)
        if not isinstance(footer_rows_raw, int) or isinstance(footer_rows_raw, bool):
            raise ValueError(
                f"{path}: 'tables.{field_name}.footer_rows' must be a non-negative integer"
            )
        if footer_rows_raw < 0:
            raise ValueError(
                f"{path}: 'tables.{field_name}.footer_rows' must be >= 0 (got {footer_rows_raw})"
            )
        list_joiner_raw = table_data.get("list_joiner", "\n")
        if not isinstance(list_joiner_raw, str):
            raise ValueError(f"{path}: 'tables.{field_name}.list_joiner' must be a string")
        tables[field_name] = TableConfig(
            shape=shape_name_raw,
            columns=list(columns_raw),
            header_row=header_row_raw,
            footer_rows=footer_rows_raw,
            list_joiner=list_joiner_raw,
        )

    sections_raw = data.get("sections")
    if sections_raw is None:
        sections_raw = {}
    if not isinstance(sections_raw, dict):
        raise ValueError(f"{path}: 'sections' must be a mapping, got {type(sections_raw).__name__}")
    sections: dict[str, SectionConfig] = {}
    for field_name, section_data in sections_raw.items():
        if not isinstance(field_name, str):
            raise ValueError(
                f"{path}: 'sections' keys must be strings, got {type(field_name).__name__}"
            )
        if not isinstance(section_data, dict):
            raise ValueError(
                f"{path}: 'sections.{field_name}' must be a mapping, got "
                f"{type(section_data).__name__}"
            )
        section_shape_raw = section_data.get("shape")
        if not isinstance(section_shape_raw, str):
            raise ValueError(f"{path}: 'sections.{field_name}.shape' must be a string")
        header_index_raw = section_data.get("header_index", 0)
        if not isinstance(header_index_raw, int) or isinstance(header_index_raw, bool):
            raise ValueError(
                f"{path}: 'sections.{field_name}.header_index' must be a non-negative integer"
            )
        bullet_index_raw = section_data.get("bullet_index", 1)
        if not isinstance(bullet_index_raw, int) or isinstance(bullet_index_raw, bool):
            raise ValueError(
                f"{path}: 'sections.{field_name}.bullet_index' must be a non-negative integer"
            )
        sections[field_name] = SectionConfig(
            shape=section_shape_raw,
            header_index=header_index_raw,
            bullet_index=bullet_index_raw,
        )

    config = TemplateConfig(
        template=template_path,
        source_slide_index=source_slide_index_raw,
        placeholders=dict(placeholders_raw),
        tables=tables,
        sections=sections,
        clear_source_slide=clear_source_slide_raw,
        overflow_ratio=float(overflow_ratio_raw),
    )

    errors = config.validate()
    if errors:
        joined = "\n  - ".join(errors)
        raise ValueError(f"{path}: config validation failed:\n  - {joined}")

    return config


_SLUG_RE = re.compile(r"[^a-z0-9]+")


def _slug_from_shape_name(name: str) -> str:
    """Convert a shape name to a valid YAML key (lowercase, underscored)."""
    slug = _SLUG_RE.sub("_", name.lower()).strip("_")
    return slug or "field"


def write_scaffold_config(
    template_path: Path,
    shape_names: list[str],
    output_path: Path,
) -> None:
    """Write a starter config file from a list of shape names.

    The user then edits the placeholders section to map semantic field names
    to the actual shape names they want to populate.

    Uses `yaml.safe_dump` rather than hand-formatted YAML so that shape
    names containing YAML-significant characters (`:`, `#`, `{}`, `[]`,
    leading `-`, etc.) are quoted correctly and produce a file that
    round-trips cleanly through `load_config`.
    """
    placeholders: dict[str, str] = {}
    if shape_names:
        for name in shape_names:
            slug = _slug_from_shape_name(name)
            # Disambiguate slug collisions by suffixing with an index so every
            # shape ends up in the scaffold — otherwise two shapes that slug
            # to the same key silently lose one.
            unique_slug = slug
            suffix = 2
            while unique_slug in placeholders:
                unique_slug = f"{slug}_{suffix}"
                suffix += 1
            placeholders[unique_slug] = name

    scaffold: dict[str, Any] = {
        "template": str(template_path),
        "source_slide_index": 1,
        "clear_source_slide": True,
        "placeholders": placeholders,
    }

    header_lines = [
        f"# Template config for {template_path.name}",
        "# Generated by `recombinase init`.",
        "# Edit the placeholders section to map your data fields to shape names.",
        "",
    ]
    body = yaml.safe_dump(scaffold, sort_keys=False, allow_unicode=True, default_flow_style=False)
    if not shape_names:
        body += "# (no named shapes found — template has only default names)\n"
    output_path.write_text("\n".join(header_lines) + body, encoding="utf-8")
