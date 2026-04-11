"""v0.1.12: table cell population + picture placeholder insertion.

Two new features covering the remaining real-template gaps:
- populate_table walks a table.rows x cells grid driven by a TableConfig,
  with header-row skipping and list-as-multi-line-cell support
- set_picture calls PicturePlaceholder.insert_picture on picture
  placeholder shapes, with optional relative-path resolution
- generate_deck routes placeholder shapes to set_shape_value OR
  set_picture based on runtime type, and processes config.tables after
  the main placeholders loop
"""

from __future__ import annotations

import struct
import zlib
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.util import Inches

from recombinase.config import TableConfig, load_config
from recombinase.generate import (
    find_shape_by_name,
    generate_deck,
    is_picture_placeholder,
    load_records,
    populate_table,
    set_picture,
)

# -- populate_table: basic population --------------------------------------


def _build_template_with_table(path: Path) -> None:
    """Build a template with a named 3x3 table (header row + 2 data rows)."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # Add a 3-row x 3-col table (1 header row + 2 body rows)
    table_shape = slide.shapes.add_table(
        rows=3, cols=3, left=Inches(0.5), top=Inches(0.5), width=Inches(8), height=Inches(2)
    )
    table_shape.name = "recent_projects"
    # Populate header row
    table = table_shape.table
    table.cell(0, 0).text = "Client & Project"
    table.cell(0, 1).text = "Role"
    table.cell(0, 2).text = "Achievements"
    # Pre-fill body rows with example content to simulate a real template
    table.cell(1, 0).text = "Example Client A — Example Project"
    table.cell(1, 1).text = "Example Role"
    table.cell(1, 2).text = "Example achievement"
    table.cell(2, 0).text = "Example Client B — Example Project"
    table.cell(2, 1).text = "Example Role"
    table.cell(2, 2).text = "Example achievement"
    prs.save(str(path))


def test_populate_table_skips_header_and_writes_data_rows(tmp_path: Path) -> None:
    template_path = tmp_path / "table.pptx"
    _build_template_with_table(template_path)

    prs = Presentation(str(template_path))
    shape = find_shape_by_name(prs.slides[0], "recent_projects")
    assert shape is not None

    config = TableConfig(
        shape="recent_projects",
        columns=["client_project", "role", "achievements"],
        header_row=True,
    )

    rows = [
        {
            "client_project": "Bank A — Data governance",
            "role": "Lead consultant",
            "achievements": "Delivered framework",
        },
        {
            "client_project": "Bank B — AI tiering",
            "role": "Advisor",
            "achievements": "Designed taxonomy",
        },
    ]

    warnings = populate_table(shape, config, rows)
    assert warnings == []

    table = shape.table
    # Header row unchanged
    assert table.cell(0, 0).text == "Client & Project"
    assert table.cell(0, 1).text == "Role"
    assert table.cell(0, 2).text == "Achievements"
    # Data row 1
    assert table.cell(1, 0).text == "Bank A — Data governance"
    assert table.cell(1, 1).text == "Lead consultant"
    assert table.cell(1, 2).text == "Delivered framework"
    # Data row 2
    assert table.cell(2, 0).text == "Bank B — AI tiering"
    assert table.cell(2, 1).text == "Advisor"
    assert table.cell(2, 2).text == "Designed taxonomy"


def test_populate_table_joins_list_values_with_newlines(tmp_path: Path) -> None:
    template_path = tmp_path / "table.pptx"
    _build_template_with_table(template_path)

    prs = Presentation(str(template_path))
    shape = find_shape_by_name(prs.slides[0], "recent_projects")
    assert shape is not None

    config = TableConfig(
        shape="recent_projects",
        columns=["client_project", "role", "achievements"],
        header_row=True,
    )

    rows = [
        {
            "client_project": "Bank A — Data governance",
            "role": "Lead",
            "achievements": [
                "Delivered framework",
                "Completed audit cycle",
                "Trained 20 reviewers",
            ],
        },
    ]

    warnings = populate_table(shape, config, rows)
    assert warnings == []

    table = shape.table
    cell_text = table.cell(1, 2).text
    assert "Delivered framework" in cell_text
    assert "Completed audit cycle" in cell_text
    assert "Trained 20 reviewers" in cell_text
    # Three lines separated by newlines
    assert cell_text.count("\n") == 2


def test_populate_table_truncates_when_rows_exceed_capacity(tmp_path: Path) -> None:
    template_path = tmp_path / "table.pptx"
    _build_template_with_table(template_path)

    prs = Presentation(str(template_path))
    shape = find_shape_by_name(prs.slides[0], "recent_projects")
    assert shape is not None

    config = TableConfig(
        shape="recent_projects",
        columns=["client_project", "role", "achievements"],
        header_row=True,
    )

    # Template has 2 data rows (plus header). Provide 4 records.
    rows = [
        {"client_project": f"Client {i}", "role": "Role", "achievements": "Ach"} for i in range(4)
    ]

    warnings = populate_table(shape, config, rows)
    assert len(warnings) == 1
    assert "truncating" in warnings[0].lower()
    assert "2 data rows" in warnings[0]

    table = shape.table
    # Only the first 2 records wrote
    assert table.cell(1, 0).text == "Client 0"
    assert table.cell(2, 0).text == "Client 1"


def test_populate_table_no_header_row_writes_into_row_0(tmp_path: Path) -> None:
    template_path = tmp_path / "table.pptx"
    _build_template_with_table(template_path)

    prs = Presentation(str(template_path))
    shape = find_shape_by_name(prs.slides[0], "recent_projects")
    assert shape is not None

    config = TableConfig(
        shape="recent_projects",
        columns=["client_project", "role", "achievements"],
        header_row=False,
    )

    rows = [
        {"client_project": "New Client", "role": "New Role", "achievements": "New Ach"},
    ]

    warnings = populate_table(shape, config, rows)
    assert warnings == []

    table = shape.table
    # Row 0 overwritten (header treated as data)
    assert table.cell(0, 0).text == "New Client"


def test_populate_table_warns_on_non_table_shape(tmp_path: Path) -> None:
    """If the configured shape isn't a table, emit a warning and skip."""
    from recombinase.generate import _write_paragraphs  # noqa: F401 (keep import side-effect)

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(4), Inches(1))
    tb.name = "not_a_table"
    tb.text_frame.text = "I am a text box"

    path = tmp_path / "textbox.pptx"
    prs.save(str(path))
    prs = Presentation(str(path))
    shape = find_shape_by_name(prs.slides[0], "not_a_table")
    assert shape is not None

    config = TableConfig(shape="not_a_table", columns=["a"])
    warnings = populate_table(shape, config, [{"a": "x"}])
    assert len(warnings) == 1
    assert "not a table shape" in warnings[0]


# -- set_picture: picture placeholder insertion ----------------------------


def _tiny_png(path: Path) -> None:
    """Write a minimal 1x1 red PNG to disk."""
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = b"\x00\x00\x00\rIHDR" + struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    ihdr_crc = struct.pack(">I", zlib.crc32(ihdr[4:]) & 0xFFFFFFFF)
    idat_data = zlib.compress(b"\x00\xff\x00\x00")
    idat = struct.pack(">I", len(idat_data)) + b"IDAT" + idat_data
    idat_crc = struct.pack(">I", zlib.crc32(b"IDAT" + idat_data) & 0xFFFFFFFF)
    iend = b"\x00\x00\x00\x00IEND" + struct.pack(">I", zlib.crc32(b"IEND") & 0xFFFFFFFF)
    path.write_bytes(sig + ihdr + ihdr_crc + idat + idat_crc + iend)


def _build_template_with_picture_placeholder(pptx_path: Path) -> None:
    """Build a template whose layout includes a picture placeholder.

    python-pptx doesn't expose a public API for adding a picture placeholder
    to a layout, so we use the layout index from the default pptx template
    that already includes one: layout 8 ("Picture with Caption") has a
    picture placeholder at idx 1.
    """
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[8])  # Picture with Caption
    # Rename the picture placeholder to our test name
    for shape in slide.shapes:
        if is_picture_placeholder(shape):
            shape.name = "photo"
            break
    prs.save(str(pptx_path))


def test_is_picture_placeholder_detects_picture_placeholders(tmp_path: Path) -> None:
    template_path = tmp_path / "pic.pptx"
    _build_template_with_picture_placeholder(template_path)

    prs = Presentation(str(template_path))
    shape = find_shape_by_name(prs.slides[0], "photo")
    assert shape is not None
    assert is_picture_placeholder(shape) is True


def test_is_picture_placeholder_false_for_regular_text_box(tmp_path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(4), Inches(1))
    tb.name = "not_a_picture"
    tb.text_frame.text = "text"
    assert is_picture_placeholder(tb) is False


def test_set_picture_inserts_image_into_placeholder(tmp_path: Path) -> None:
    template_path = tmp_path / "pic.pptx"
    _build_template_with_picture_placeholder(template_path)

    png_path = tmp_path / "headshot.png"
    _tiny_png(png_path)

    prs = Presentation(str(template_path))
    shape = find_shape_by_name(prs.slides[0], "photo")
    assert shape is not None

    set_picture(shape, str(png_path))

    # After insert_picture, the shape should still have a matching name and
    # the save-reopen cycle should still find it.
    output = tmp_path / "out.pptx"
    prs.save(str(output))
    reopened = Presentation(str(output))
    new_shape = find_shape_by_name(reopened.slides[0], "photo")
    assert new_shape is not None


def test_set_picture_raises_on_missing_file(tmp_path: Path) -> None:
    import pytest

    template_path = tmp_path / "pic.pptx"
    _build_template_with_picture_placeholder(template_path)

    prs = Presentation(str(template_path))
    shape = find_shape_by_name(prs.slides[0], "photo")
    assert shape is not None

    with pytest.raises(FileNotFoundError, match="picture file not found"):
        set_picture(shape, str(tmp_path / "nonexistent.jpg"))


def test_set_picture_resolves_relative_path_against_base_dir(tmp_path: Path) -> None:
    template_path = tmp_path / "pic.pptx"
    _build_template_with_picture_placeholder(template_path)

    # Create a png in a sub directory
    base_dir = tmp_path / "data"
    base_dir.mkdir()
    png_path = base_dir / "headshot.png"
    _tiny_png(png_path)

    prs = Presentation(str(template_path))
    shape = find_shape_by_name(prs.slides[0], "photo")
    assert shape is not None

    # Pass a relative path plus the base_dir
    set_picture(shape, "headshot.png", base_dir=base_dir)

    # If it worked, there's no exception and the shape remains
    output = tmp_path / "out.pptx"
    prs.save(str(output))


def test_set_picture_noop_on_empty_value(tmp_path: Path) -> None:
    template_path = tmp_path / "pic.pptx"
    _build_template_with_picture_placeholder(template_path)

    prs = Presentation(str(template_path))
    shape = find_shape_by_name(prs.slides[0], "photo")
    assert shape is not None

    # Should not raise
    set_picture(shape, "")
    set_picture(shape, None)


# -- end-to-end: generate_deck with table config + placeholders ------------


def test_generate_deck_populates_table_from_record(tmp_path: Path, write_config) -> None:
    """Full pipeline: template with a table shape, config has `tables:`,
    records have list-of-dicts, generated output has the rows populated."""
    template_path = tmp_path / "with_table.pptx"
    _build_template_with_table(template_path)

    config_path = tmp_path / "config.yaml"
    config_path.write_text(
        yaml.safe_dump(
            {
                "template": str(template_path),
                "source_slide_index": 1,
                "clear_source_slide": True,
                "overflow_ratio": 0,  # disable overflow for this test
                "placeholders": {},
                "tables": {
                    "recent_projects": {
                        "shape": "recent_projects",
                        "header_row": True,
                        "columns": ["client_project", "role", "achievements"],
                    }
                },
            }
        ),
        encoding="utf-8",
    )

    data_dir = tmp_path / "data"
    data_dir.mkdir()
    (data_dir / "r1.yaml").write_text(
        yaml.safe_dump(
            {
                "id": "alpha",
                "recent_projects": [
                    {
                        "client_project": "Bank A — Framework",
                        "role": "Lead",
                        "achievements": ["One", "Two"],
                    },
                    {
                        "client_project": "Bank B — AI",
                        "role": "Advisor",
                        "achievements": ["Three"],
                    },
                ],
            }
        ),
        encoding="utf-8",
    )

    config = load_config(config_path)
    records = load_records(data_dir)
    output = tmp_path / "out.pptx"
    result = generate_deck(config, records, output)

    assert result["records_generated"] == 1
    assert result["warnings"] == []

    # Re-open and verify the table was populated on the first generated slide
    prs = Presentation(str(output))
    shape = find_shape_by_name(prs.slides[0], "recent_projects")
    assert shape is not None
    assert shape.has_table

    table = shape.table
    # Header row preserved
    assert table.cell(0, 0).text == "Client & Project"
    # Data rows populated
    assert "Bank A — Framework" in table.cell(1, 0).text
    assert table.cell(1, 1).text == "Lead"
    assert "One" in table.cell(1, 2).text
    assert "Two" in table.cell(1, 2).text
    assert "Bank B — AI" in table.cell(2, 0).text


def test_generate_deck_warns_when_table_record_is_not_list(tmp_path: Path, write_config) -> None:
    template_path = tmp_path / "with_table.pptx"
    _build_template_with_table(template_path)

    config_path = tmp_path / "config.yaml"
    config_path.write_text(
        yaml.safe_dump(
            {
                "template": str(template_path),
                "source_slide_index": 1,
                "clear_source_slide": True,
                "overflow_ratio": 0,
                "placeholders": {},
                "tables": {
                    "recent_projects": {
                        "shape": "recent_projects",
                        "columns": ["client_project"],
                    }
                },
            }
        ),
        encoding="utf-8",
    )

    data_dir = tmp_path / "data"
    data_dir.mkdir()
    (data_dir / "r1.yaml").write_text(
        yaml.safe_dump({"id": "x", "recent_projects": "not a list"}),
        encoding="utf-8",
    )

    config = load_config(config_path)
    records = load_records(data_dir)
    result = generate_deck(config, records, tmp_path / "out.pptx")

    assert result["records_generated"] == 1
    assert any("expects a list" in w for w in result["warnings"])


def test_load_config_parses_tables_section(tmp_path: Path, simple_template: Path) -> None:
    config_path = tmp_path / "c.yaml"
    config_path.write_text(
        yaml.safe_dump(
            {
                "template": str(simple_template),
                "placeholders": {"name": "Field"},
                "tables": {
                    "projects": {
                        "shape": "ProjTable",
                        "columns": ["a", "b"],
                        "header_row": False,
                        "list_joiner": " | ",
                    }
                },
            }
        ),
        encoding="utf-8",
    )

    cfg = load_config(config_path)
    assert "projects" in cfg.tables
    tc = cfg.tables["projects"]
    assert tc.shape == "ProjTable"
    assert tc.columns == ["a", "b"]
    assert tc.header_row is False
    assert tc.list_joiner == " | "


def test_load_config_rejects_malformed_tables_section(
    tmp_path: Path, simple_template: Path
) -> None:
    import pytest

    config_path = tmp_path / "c.yaml"
    config_path.write_text(
        yaml.safe_dump(
            {
                "template": str(simple_template),
                "placeholders": {"name": "Field"},
                "tables": "not a dict",
            }
        ),
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="'tables' must be a mapping"):
        load_config(config_path)
