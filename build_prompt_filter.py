#!/usr/bin/env python3
"""
Finds a C# CommandMethod and any custom functions it calls directly,
compiles them into a report, and copies the result to the clipboard.

This script performs a one-level dependency check only.

Place this file in the same directory as your project and run:
python build_prompt_filter.py <CommandMethodName>
"""

import sys
import re
from pathlib import Path
from datetime import datetime
import pyperclip
from collections import defaultdict

# --- Config ---------------------------------------------------------
# Order priority for files in the report; anything not listed falls back to alphabetical
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
    This is used to identify which function calls are "custom" (defined in the project).
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
    Finds potential method calls (PascalCase followed by an opening parenthesis).
    """
    pattern = re.compile(r'\b([A-Z][a-zA-Z0-9_]*)(?=\s*\()')
    calls = set(pattern.findall(body))
    return calls

def main():
    if len(sys.argv) < 2:
        print(f"Usage: python {Path(__file__).name} <CommandMethodName>")
        print("Example: python build_prompt_filter.py EMBEDIMAGES")
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
    print(f"Indexed {len(method_index)} custom methods from {len(file_contents)} files.")

    entry_method, entry_file = get_entry_point(target_command, file_contents)

    if not entry_method:
        print(f"\nError: Could not find [CommandMethod(\"{target_command}\")] in any .cs file.")
        return

    # --- REVISED: ONE-LEVEL DEPENDENCY CHECK ---
    included_functions = defaultdict(set)

    print("\nStarting analysis for the CommandMethod and its direct dependencies...")
    print(f"Entry point: Method '{entry_method}' in file '{entry_file.name}'")

    # 1. Find and add the entry CommandMethod itself to the report.
    entry_content = file_contents[entry_file]
    entry_body = find_method_body(entry_content, entry_method)
    
    if not entry_body:
        print(f"Error: Could not find the body of the entry method '{entry_method}'.")
        return
    
    included_functions[entry_file].add(entry_body)
    print(f" - Added entry method '{entry_method}' to the report.")

    # 2. Find all methods called directly within the entry method's body.
    direct_calls = find_called_methods(entry_body)

    # 3. Filter these calls to include only custom functions defined in the project.
    # Exclude the entry method itself in case of recursion.
    custom_dependencies = {call for call in direct_calls if call in method_index and call != entry_method}

    if custom_dependencies:
        print(f" - Found direct dependencies on: {', '.join(custom_dependencies)}")
        # 4. Loop through ONLY the direct dependencies and add them.
        for method_name in custom_dependencies:
            file_path = method_index[method_name]
            content = file_contents[file_path]
            method_body = find_method_body(content, method_name)
            
            if method_body:
                included_functions[file_path].add(method_body)
                print(f"   - Added function '{method_name}' from '{file_path.name}'")
    else:
        print(" - No custom, direct dependencies found.")

    # --- REPORT GENERATION ---
    sorted_files = sorted(list(included_functions.keys()), key=lambda p: (ORDER_HINT.index(p.name) if p.name in ORDER_HINT else 1000, p.name.lower()))

    lines = [
        f"# Report for CommandMethod '{target_command}' and its direct dependencies\n"
        f"# Generated: {datetime.now().isoformat(timespec='seconds')}\n\n"
        f"## This report contains the entry CommandMethod and all custom functions it directly calls.\n"
    ]

    for path in sorted_files:
        functions = sorted(list(included_functions[path]))
        if not functions:
            continue

        lines.append(f"\n===== BEGIN FUNCTIONS FROM: {path.name} =====\n")
        lines.append(f"```{lang_for(path)}\n")
        # Join functions with two newlines for better spacing
        lines.append("\n\n".join(functions))
        lines.append(f"\n```\n===== END {path.name} =====")

    output_text = "".join(lines)
    pyperclip.copy(output_text)
    
    print("\n[done] Copied a report of the following functions to the clipboard:")
    for path in sorted_files:
        print(f" - From file {path.name} ({len(included_functions[path])} function(s))")

if __name__ == "__main__":
    main()
