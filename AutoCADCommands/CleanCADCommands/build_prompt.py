#!/usr/bin/env python3
"""
Concatenate project files into prompt.txt with clear section headers and code fences.
Place this file in the same directory as your project and run:  python build_prompt.py
"""

from pathlib import Path
from datetime import datetime

# --- Config ---------------------------------------------------------
# Files will be written to this target (overwritten, with a timestamped backup if it exists)
TARGET_NAME = "prompt.txt"

# Order priority; anything not listed falls back to alphabetical after these
ORDER_HINT = []

# Include only these extensions (set to None to include everything)
ALLOW_EXTS = {'.cs'}

# Map extensions to code-fence languages
LANG_MAP = {
    ".cs": "csharp",
}
# -------------------------------------------------------------------

def lang_for(path: Path) -> str:
    return LANG_MAP.get(path.suffix.lower(), "")

def main():
    root = Path(__file__).resolve().parent
    target = root / TARGET_NAME
    self_name = Path(__file__).name

    # Collect candidate files
    files = []
    for p in root.iterdir():
        if not p.is_file():
            continue
        if p.name in {self_name, TARGET_NAME}:  # skip myself and the target
            continue
        if p.name.startswith("."):              # skip dotfiles
            continue
        if ALLOW_EXTS is not None and p.suffix.lower() not in ALLOW_EXTS:
            continue
        files.append(p)

    # Sort: ORDER_HINT first (in given order), then remaining alphabetically
    def sort_key(p: Path):
        try:
            idx = ORDER_HINT.index(p.name)
        except ValueError:
            idx = 10_000  # push behind hinted entries
        return (idx, p.name.lower())

    files.sort(key=sort_key)

    # Make a timestamped backup if prompt.txt exists
    if target.exists():
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = target.with_name(f"{target.stem}.{stamp}.bak{target.suffix}")
        target.replace(backup)
        print(f"[info] Backed up existing {TARGET_NAME} -> {backup.name}")

    lines = []
    header = (
        f"# Bundled project files → {TARGET_NAME}\n"
        f"# Generated: {datetime.now().isoformat(timespec='seconds')}\n"
        f"# Directory: {root}\n\n"
        f"## File list (in order):\n"
        + "".join(f"- {p.name}\n" for p in files)
        + "\n"
    )
    lines.append(header)

    for p in files:
        lang = lang_for(p)
        lines.append(f"===== BEGIN {p.name} =====\n")
        lines.append(f"```{lang}\n")
        try:
            text = p.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            # Fall back to latin-1 if a file isn't UTF-8; avoids hard crashes
            text = p.read_text(encoding="latin-1")
        lines.append(text.rstrip() + "\n")
        lines.append("```\n")
        lines.append(f"===== END {p.name} =====\n\n")

    target.write_text("".join(lines), encoding="utf-8", newline="\n")
    print(f"[done] Wrote {TARGET_NAME} with {len(files)} file(s).")

if __name__ == "__main__":
    main()