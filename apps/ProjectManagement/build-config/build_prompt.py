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

# OPTION: Hardcode a list of files to include.
# If this list is NOT empty, ONLY these files will be included, in this order.
# If this list IS empty, the script will proceed with standard file discovery and sorting.
HARDCODED_FILES = [
    "index.html",
    "styles.css",
]

# Order priority; anything not listed falls back to alphabetical after these
ORDER_HINT = [
    "index.html",
    "styles.css",
    "script.js",
]

# Include only these extensions (set to None to include everything)
ALLOW_EXTS = {".html", ".css", ".js", ".py", ".json", ".md", ".ps1"}

# Map extensions to code-fence languages
LANG_MAP = {
    ".html": "html",
    ".css": "css",
    ".js": "javascript",
    ".py": "python",
    ".json": "json",
    ".md": "markdown",
    ".txt": "",  # plain
}
# -------------------------------------------------------------------


def lang_for(path: Path) -> str:
    return LANG_MAP.get(path.suffix.lower(), "")


def main():
    root = Path(__file__).resolve().parent
    target = root / TARGET_NAME
    self_name = Path(__file__).name

    files = []
    use_hardcoded = bool(HARDCODED_FILES)

    if use_hardcoded:
        print(f"[info] Using hardcoded file list: {HARDCODED_FILES}")
        # Use hardcoded files, maintaining their order
        for filename in HARDCODED_FILES:
            p = root / filename
            if p.is_file():
                files.append(p)
            else:
                print(
                    f"[warn] Hardcoded file not found or is not a file: {filename}")
        # Skip standard sorting, the order is the hardcoded list's order
    else:
        # Standard file discovery
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
