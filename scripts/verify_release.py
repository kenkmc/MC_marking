"""Fail a release build when its tag and embedded versions disagree."""

from __future__ import annotations

import argparse
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SEMVER = re.compile(r"^v?(\d+)\.(\d+)\.(\d+)$")


def extract(pattern: str, path: Path) -> str:
    match = re.search(pattern, path.read_text(encoding="utf-8"), re.MULTILINE)
    if not match:
        raise RuntimeError(f"Version was not found in {path.name}")
    return match.group(1)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("tag")
    args = parser.parse_args()
    match = SEMVER.fullmatch(args.tag)
    if not match:
        raise SystemExit(f"Release tag must use vMAJOR.MINOR.PATCH: {args.tag}")
    expected = ".".join(match.groups())

    app = extract(
        r'^\s*APP_VERSION\s*=\s*"([^"]+)"',
        ROOT / "omr_software.py",
    )
    installer = extract(
        r'^\s*#define\s+AppVersion\s+"([^"]+)"',
        ROOT / "installer.iss",
    )
    file_version = extract(
        r"""StringStruct\(["']FileVersion["'],\s*["']([^"']+)["']\)""",
        ROOT / "version_info.txt",
    )
    if app != expected or installer != expected or file_version != expected:
        raise SystemExit(
            "Version mismatch: "
            f"tag={expected}, app={app}, installer={installer}, "
            f"file={file_version}"
        )
    print(f"Release metadata verified: v{expected}")


if __name__ == "__main__":
    main()
