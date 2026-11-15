"""Layering (docs/05 section 2): render and domain stay pure; only `publish` and `main` glue."""

import ast
from pathlib import Path

SRC = Path(__file__).resolve().parents[2] / "src" / "esp_screen"
PURE = ("render", "domain")
IO_PACKAGES = (
    "esp_screen.device",
    "esp_screen.storage",
    "esp_screen.web",
    "esp_screen.beszel",
    "esp_screen.publish",
    "esp_screen.main",
    "httpx",
    "fastapi",
    "sqlite3",
)


def imports(path: Path) -> set[str]:
    tree = ast.parse(path.read_text())
    out = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            out.update(a.name for a in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            out.add(node.module)
    return out


def test_pure_packages_do_no_io():
    for pkg in PURE:
        for f in (SRC / pkg).rglob("*.py") if (SRC / pkg).exists() else []:
            bad = [m for m in imports(f) if m.startswith(IO_PACKAGES)]
            assert not bad, f"{f.relative_to(SRC)} imports {bad}"


def test_device_client_knows_nothing_else():
    for f in (SRC / "device").rglob("*.py"):
        bad = [m for m in imports(f) if m.startswith("esp_screen.") and m != "esp_screen.device"]
        assert not bad, f"{f.relative_to(SRC)} imports {bad}"
