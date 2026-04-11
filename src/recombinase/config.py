"""Template configuration: load and validate a per-template YAML config."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml


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
    """

    clear_source_slide: bool = True
    """If True, the source (example) slide is removed from the final output
    so only the generated per-record slides remain."""

    def validate(self) -> list[str]:
        """Return a list of validation error messages, or empty list if OK."""
        errors: list[str] = []
        if not self.template.exists():
            errors.append(f"Template file not found: {self.template}")
        if self.source_slide_index < 1:
            errors.append(f"source_slide_index must be >= 1 (got {self.source_slide_index})")
        if not self.placeholders:
            errors.append("placeholders mapping is empty — at least one field is required")
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

    placeholders_raw = data.get("placeholders") or {}
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

    config = TemplateConfig(
        template=template_path,
        source_slide_index=source_slide_index_raw,
        placeholders=dict(placeholders_raw),
        clear_source_slide=clear_source_slide_raw,
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
    """
    lines = [
        f"# Template config for {template_path.name}",
        "# Generated by `recombinase init`.",
        "# Edit the placeholders section to map your data fields to shape names.",
        "",
        f"template: {template_path}",
        "source_slide_index: 1",
        "clear_source_slide: true",
        "",
        "# placeholders: data_field_name -> shape_name_in_template",
        "placeholders:",
    ]
    if shape_names:
        for name in shape_names:
            slug = _slug_from_shape_name(name)
            lines.append(f"  {slug}: {name}")
    else:
        lines.append("  # (no named shapes found — template has only default names)")

    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
