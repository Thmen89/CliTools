import sys
from pathlib import Path

LIB = str(Path(__file__).resolve().parents[2] / "lib")


def run(ctx):
    if LIB not in sys.path:
        sys.path.insert(0, LIB)
    from ilst_tools import app
    app.run(ctx, "settings")
