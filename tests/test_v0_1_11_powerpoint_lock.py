"""v0.1.11: PackageNotFoundError formatter detects PowerPoint lock files.

When a Windows user opens a pptx in PowerPoint, the OS creates a ~$filename.pptx
lock-file marker alongside the original file. When recombinase hits
python-pptx's PackageNotFoundError, it should check for that lock file and
emit a specific "open in PowerPoint" message instead of the generic
"corrupt / empty / wrong format" message.
"""

from __future__ import annotations

from pathlib import Path

from pptx.exc import PackageNotFoundError

from recombinase.cli import _format_package_not_found


def test_format_package_not_found_detects_lock_file(tmp_path: Path) -> None:
    """When a ~$ lock file is present, emit the PowerPoint-open hint."""
    fake_pptx = tmp_path / "cv_template.pptx"
    fake_pptx.write_bytes(b"not actually a valid pptx")
    lock_file = tmp_path / "~$cv_template.pptx"
    lock_file.write_bytes(b"")

    exc = PackageNotFoundError(f"Package not found at '{fake_pptx}'")
    msg = _format_package_not_found(exc)

    assert "open in PowerPoint" in msg
    assert "cv_template.pptx" in msg
    assert "Close" in msg or "close" in msg


def test_format_package_not_found_generic_when_no_lock_file(tmp_path: Path) -> None:
    """Without a lock file, fall back to the generic message but still mention PowerPoint."""
    fake_pptx = tmp_path / "some.pptx"
    fake_pptx.write_bytes(b"not actually a valid pptx")

    exc = PackageNotFoundError(f"Package not found at '{fake_pptx}'")
    msg = _format_package_not_found(exc)

    # Generic message hints at the most common cause
    assert "PowerPoint" in msg
    # Still falls back to explaining the other possibilities
    assert "corrupt" in msg or "empty" in msg


def test_format_package_not_found_handles_path_without_quotes(tmp_path: Path) -> None:
    """If the exception message doesn't contain a quoted path, fall back to generic."""

    class _WeirdError(Exception):
        pass

    exc = _WeirdError("something unusual happened without a path")
    msg = _format_package_not_found(exc)

    # Should still produce a valid message, not crash
    assert "PowerPoint" in msg or "pptx" in msg


def test_format_package_not_found_ignores_lock_file_in_different_dir(
    tmp_path: Path,
) -> None:
    """Lock file must be in the same directory as the pptx, not anywhere."""
    fake_pptx = tmp_path / "sub" / "cv_template.pptx"
    fake_pptx.parent.mkdir()
    fake_pptx.write_bytes(b"not a real pptx")
    # Lock file in a DIFFERENT directory — should not trigger the hint
    (tmp_path / "~$cv_template.pptx").write_bytes(b"")

    exc = PackageNotFoundError(f"Package not found at '{fake_pptx}'")
    msg = _format_package_not_found(exc)

    # The specific "open in PowerPoint (lock file detected)" message should NOT fire
    # because the lock file isn't in the same directory
    assert "lock file" not in msg.lower()
