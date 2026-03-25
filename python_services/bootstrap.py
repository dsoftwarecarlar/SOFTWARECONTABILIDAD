from __future__ import annotations

import os
import site
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VENDOR = ROOT / "vendor"


def bootstrap_vendor() -> None:
    if not VENDOR.is_dir():
        return

    site.addsitedir(str(VENDOR))

    pywin32_system32 = VENDOR / "pywin32_system32"
    if pywin32_system32.is_dir():
        if hasattr(os, "add_dll_directory"):
            os.add_dll_directory(str(pywin32_system32))
        if str(pywin32_system32) not in sys.path:
            sys.path.insert(0, str(pywin32_system32))
        current_path = os.environ.get("PATH", "")
        if str(pywin32_system32) not in current_path.split(os.pathsep):
            os.environ["PATH"] = str(pywin32_system32) + os.pathsep + current_path
