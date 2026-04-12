"""Tests for data loading robustness: duplicate keys, empty file warning, size guard."""

from __future__ import annotations

import pytest

from recombinase.config import load_config
from recombinase.generate import load_records


def test_yaml_duplicate_key_in_record_raises(tmp_path):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    (data_dir / "record.yaml").write_text("name: A\nname: B\n", encoding="utf-8")
    with pytest.raises(ValueError, match="duplicate key"):
        load_records(data_dir)


def test_yaml_duplicate_key_in_config_raises(tmp_path):
    config_file = tmp_path / "config.yaml"
    config_file.write_text(
        "template: deck.pptx\ntemplate: other.pptx\n",
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="duplicate key"):
        load_config(config_file)


def test_empty_yaml_file_warns(tmp_path):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    (data_dir / "empty.yaml").write_text("", encoding="utf-8")
    with pytest.warns(UserWarning, match="empty"):
        load_records(data_dir)


def test_record_file_size_guard(tmp_path):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    large_file = data_dir / "big.yaml"
    # Write ~11 MB of valid YAML lines to exceed the 10 MB limit
    chunk = "padding: x\n" * 1000
    with large_file.open("w", encoding="utf-8") as fh:
        for _ in range(1100):
            fh.write(chunk)
    with pytest.raises(ValueError, match="10 MB"):
        load_records(data_dir)


def test_config_file_size_guard(tmp_path):
    config_file = tmp_path / "config.yaml"
    # Write ~11 MB to exceed the 10 MB limit
    chunk = "# padding\n" * 1000
    with config_file.open("w", encoding="utf-8") as fh:
        fh.write("template: deck.pptx\n")
        for _ in range(1100):
            fh.write(chunk)
    with pytest.raises(ValueError, match="10 MB"):
        load_config(config_file)
