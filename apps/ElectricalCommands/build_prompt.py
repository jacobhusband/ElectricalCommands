#!/usr/bin/env python3
"""
Correctly finds all dependencies for a C# CommandMethod, concatenates the
relevant files, and copies the result to the clipboard.

Place this file in the same directory as your project and run:
python build_prompt.py <CommandMethodName>
"""

import sys
import re
from pathlib import Path
from datetime import datetime
import pyperclip

# --- Config ---------------------------------------------------------
# Order priority; anything not listed falls back to alphabetical after these
ORDER_HINT = []

# Include only these extensions
ALLOW_EXTS = {'.cs'}

# Map extensions to code-fence languages
LANG_MAP = {".cs": "csharp"}
# -------------------------------------------------------------------

def lang_for(path: Path) -> str:
    """Returns the code-fence language for a given file path."""
    return LANG_MAP.get(path.suffix.lower(), "")

def get_entry_point(target_command: str, file_contents: dict):
    """
    Finds the file and method name associated with the [CommandMethod] attribute.
    """
    command_pattern = re.compile(r'\[CommandMethod\s*\(\s*"' + re.escape(target_command) + r'"[^\]]*\]', re.IGNORECASE)
    method_sig_pattern = re.compile(r'\s*(?:public|private|internal|static|protected|async)?\s+[\w\.<>\[\]]+\s+([a-zA-Z_]\w*)\s*\(')

    for path, content in file_contents.items():
        match = command_pattern.search(content)
        if match:
            following_text = content[match.end():]
            method_match = method_sig_pattern.search(following_text)
            if method_match:
                method_name = method_match.group(1)
                return method_name, path
    return None, None

def create_method_definition_index(file_contents: dict):
    """
    Scans all files and creates a map of {method_name: file_path}.
    """
    index = {}
    definition_pattern = re.compile(r'^\s*(?:public|private|internal|static|protected|async)\s+.*?([a-zA-Z_]\w*)\s*\(.*?\)\s*\{', re.MULTILINE)
    
    for path, content in file_contents.items():
        for match in definition_pattern.finditer(content):
            method_name = match.group(1)
            if method_name not in index:
                index[method_name] = path
    return index

def find_method_body(content: str, method_name: str) -> str:
    """
    Returns the full definition of a specific method using brace counting.
    """
    pattern = re.compile(r'(?:public|private|internal|static|protected|async)?\s+.*?\s+' + re.escape(method_name) + r'\s*\(.*?\)\s*\{', re.DOTALL)
    match = pattern.search(content)

    if not match:
        return ""

    start_index = match.end()
    brace_count = 1
    
    for i in range(start_index, len(content)):
        if content[i] == '{':
            brace_count += 1
        elif content[i] == '}':
            brace_count -= 1
        
        if brace_count == 0:
            return content[match.start():i+1]
    return ""

def find_called_methods(body: str) -> set:
    """
    Finds potential method calls (PascalCase) within a given method body.
    """
    pattern = re.compile(r'\b([A-Z][a-zA-Z0-9_]*)(?=\s*\()')
    calls = set(pattern.findall(body))
    return calls

def main():
    if len(sys.argv) < 2:
        print(f"Usage: python {Path(__file__).name} <CommandMethodName>")
        print("Example: python build_prompt.py EMBEDIMAGES")
        return
    target_command = sys.argv[1]

    root = Path(__file__).resolve().parent
    self_name = Path(__file__).name

    # Recursively find all .cs files in the directory and subdirectories
    all_cs_files = [p for p in root.rglob("*.cs") if p.is_file() and p.name != self_name and not p.name.startswith(".")]
    file_contents = {p: p.read_text(encoding="utf-8", errors="ignore") for p in all_cs_files}

    if not file_contents:
        print("Error: No .cs files found in the directory or subdirectories.")
        return

    method_index = create_method_definition_index(file_contents)
    print(f"Indexed {len(method_index)} methods from {len(file_contents)} files.")

    entry_method, entry_file = get_entry_point(target_command, file_contents)

    if not entry_method:
        print(f"\nError: Could not find [CommandMethod(\"{target_command}\")] in any .cs file.")
        return

    files_to_include = {entry_file}
    methods_to_process = {entry_method}
    processed_methods = set()

    print("\nStarting dependency analysis...")
    print(f"Entry point: Method '{entry_method}' in file '{entry_file.name}'")

    while methods_to_process:
        method_name = methods_to_process.pop()
        if method_name in processed_methods:
            continue
        processed_methods.add(method_name)

        if method_name not in method_index:
            continue 
        
        file_path = method_index[method_name]
        files_to_include.add(file_path)
        content = file_contents[file_path]

        method_body = find_method_body(content, method_name)
        if not method_body:
            continue

        called_methods = find_called_methods(method_body)
        project_dependencies = {call for call in called_methods if call in method_index}

        if project_dependencies:
            print(f" - Analyzing '{method_name}': Found dependencies on: {', '.join(project_dependencies)}")
        
        for called_method in project_dependencies:
            if called_method not in processed_methods:
                methods_to_process.add(called_method)

    files = sorted(list(files_to_include), key=lambda p: (ORDER_HINT.index(p.name) if p.name in ORDER_HINT else 1000, p.name.lower()))

    # --- MODIFICATION IS HERE ---
    lines = [
        f"# Bundled project files for command: {target_command} and its dependencies\n"
        f"# Generated: {datetime.now().isoformat(timespec='seconds')}\n"
        f"# Directory: {root}\n\n"
        f"## File list ({len(files)} file(s) included):\n" + "".join(f"- {p.name}\n" for p in files) + "\n"
        f"# INSTRUCTION:\n"
        f"# Return copy & paste drop in .cs files for only the files that have been modified.\n\n"
    ]
    # --- END MODIFICATION ---

    for p in files:
        lines.append(f"===== BEGIN {p.name} =====\n")
        lines.append(f"```{lang_for(p)}\n")
        lines.append(file_contents[p].rstrip() + "\n")
        lines.append("```\n")
        lines.append(f"===== END {p.name} =====\n\n")

    output_text = "".join(lines)
    pyperclip.copy(output_text)
    
    print("\n[done] Copied content of the following file(s) to clipboard:")
    for p in files:
        print(f" - {p.name}")

if __name__ == "__main__":
    main()
