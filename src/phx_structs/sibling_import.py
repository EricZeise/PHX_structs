"""Makes PHX_pyxl's pure-Python phpp_tool package importable from this sibling repo.

Same "no vendoring, no pip install, direct sibling-path import" choice
phpp-shape-sync's normalizer.py makes for the same modules -- map_parser.py,
locators.py, and reader.py have no third-party dependencies beyond openpyxl,
which this repo already requires.
"""

from __future__ import annotations

import sys
from pathlib import Path

PHX_PYXL_SRC = Path("/Users/smini/Documents/Coding/PHX_pyxl/src")


def ensure_phpp_tool_importable() -> None:
    if str(PHX_PYXL_SRC) not in sys.path:
        sys.path.insert(0, str(PHX_PYXL_SRC))
