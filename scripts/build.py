"""Create a deterministic manifest for the checked-in static publication."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SITE = ROOT / "site"


def digest(path: Path) -> dict[str, object]:
    data = path.read_bytes()
    return {
        "path": path.relative_to(SITE).as_posix(),
        "bytes": len(data),
        "sha256": hashlib.sha256(data).hexdigest(),
    }


def main() -> None:
    files = [p for p in SITE.rglob("*") if p.is_file() and p.name != "manifest.json"]
    manifest = {"title": "УП ПМ.01 · 09.01.05", "files": [digest(p) for p in sorted(files)]}
    (SITE / "manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(f"BUILT: manifest for {len(files)} published files")


if __name__ == "__main__":
    main()
