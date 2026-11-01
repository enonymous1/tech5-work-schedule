"""File output for tech5-work-schedule: filenames and workbook saving.

Previously this lived inside the generator alongside rendering. The
split matters because filename rules (Unicode handling, collision
numbering) and rendering rules change for different reasons and are
tested separately; ``generate`` is the only function that touches the
filesystem.
"""

from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from pathlib import Path

from .config import load_config
from .render import build_workbook

_RESERVED_CHARS = set('/\\?%*:|"<>')
"""Characters Windows forbids in filenames, plus path separators."""


def slugify(project_name: str) -> str:
    """Make a project name filename-safe, keeping Unicode letters.

    Only path separators, control characters, and Windows-reserved
    characters are removed; accented letters survive, so distinct
    project names stay distinct. Remaining runs of non-word characters
    (including spaces) become a single underscore. The result is capped
    at 60 characters and falls back to "Output" for an empty name.
    """
    cleaned = "".join(
        ch
        for ch in unicodedata.normalize("NFC", project_name)
        if ch not in _RESERVED_CHARS and ord(ch) >= 32
    )
    slug = re.sub(r"[^\w]+", "_", cleaned, flags=re.UNICODE).strip("_")
    return slug[:60] or "Output"


def _default_output_path(project_name: str) -> Path:
    """Return a non-existent default output path, numbering collisions.

    The timestamp has one-second resolution, so two runs in the same
    second would otherwise return the same filename and the second run
    would silently overwrite the first. Numbered suffixes (``-2``,
    ``-3``, ...) keep every default-named run's output.
    """
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = f"TECH5_{slugify(project_name)}_{timestamp}"
    candidate = output_dir / f"{stem}.xlsx"
    counter = 2
    while candidate.exists():
        candidate = output_dir / f"{stem}-{counter}.xlsx"
        counter += 1
    return candidate


def generate(config_path: str | Path, output: str | Path | None = None) -> Path:
    """Generate a workbook from a config file and return its path.

    With no ``output``, the workbook goes to ``output/`` under the current
    working directory, named ``TECH5_<project>_<timestamp>.xlsx`` (with a
    ``-2``, ``-3``, ... suffix if that name already exists, so a
    default-named run never overwrites an earlier file). With an
    explicit ``output``, that exact path is used — and overwritten if it
    exists — and parent directories are created.
    """
    config = load_config(config_path)
    wb = build_workbook(config)
    if output is None:
        output_path = _default_output_path(config.meta.project_name)
    else:
        output_path = Path(output)
        # ``Path("file.xlsx").parent`` is ``Path(".")``, which always
        # exists; the old ``parent != Path("")`` check was dead code
        # (``Path("")`` is also ``Path(".")``) and is gone.
        output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
