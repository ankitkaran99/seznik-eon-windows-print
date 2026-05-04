import os
import sys


def _bundle_path(*parts: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    return os.path.join(base, *parts)


os.environ["TCL_LIBRARY"] = _bundle_path("tcl", "tcl8.6")
os.environ["TK_LIBRARY"] = _bundle_path("tcl", "tk8.6")
