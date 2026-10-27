"""Rule 1 in app/README.md: render/ and domain/ import nothing from web, storage or publish."""

import ast
from pathlib import Path

SRC = Path(__file__).resolve().parents[2] / "src" / "trmnl_calendar"
FORBIDDEN = ("web", "storage", "publish")


def test_pure_packages_do_not_import_io_layers():
    bad = []
    for pkg in ("render", "domain"):
        for f in (SRC / pkg).rglob("*.py") if (SRC / pkg).exists() else []:
            for node in ast.walk(ast.parse(f.read_text())):
                names = []
                if isinstance(node, ast.Import):
                    names = [a.name for a in node.names]
                elif isinstance(node, ast.ImportFrom) and node.module:
                    names = [node.module]
                bad += [
                    f"{f.name}: {n}"
                    for n in names
                    if any(n.startswith(f"trmnl_calendar.{x}") for x in FORBIDDEN)
                ]
    assert not bad, bad
