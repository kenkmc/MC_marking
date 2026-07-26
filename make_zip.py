"""Create the portable CheckMate release archive."""

from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent


def app_version() -> str:
    source = (ROOT / "omr_software.py").read_text(encoding="utf-8")
    match = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"', source, re.MULTILINE)
    if not match:
        raise RuntimeError("APP_VERSION was not found in omr_software.py")
    return match.group(1)


def build_archive(source: Path, output: Path) -> tuple[int, int]:
    if not (source / "CheckMate.exe").is_file():
        raise FileNotFoundError(f"CheckMate.exe was not found in {source}")
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()

    count = 0
    with zipfile.ZipFile(
        output,
        "w",
        zipfile.ZIP_DEFLATED,
        compresslevel=6,
        allowZip64=True,
    ) as archive:
        for path in sorted(source.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(source))
                count += 1
    return count, output.stat().st_size


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--source",
        type=Path,
        default=ROOT / "dist" / "CheckMate",
    )
    parser.add_argument("--version", default=app_version())
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    output = args.output or ROOT / "dist" / f"CheckMate_v{args.version}.zip"
    count, size = build_archive(args.source.resolve(), output.resolve())
    print(f"Created {output}: {count} files, {size / 1024 / 1024:.1f} MB")


if __name__ == "__main__":
    main()
