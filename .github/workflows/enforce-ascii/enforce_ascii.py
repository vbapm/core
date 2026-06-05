"""
Enforce ASCII-only characters in VBA source files.

Inspired by DecimalTurn/Enforce-CRLF.
Scans files with given extensions and fails if any non-ASCII characters are found.
"""

import os
import sys


def is_ascii(char: str) -> bool:
    """Check if a single character is ASCII (code point 0-127)."""
    return ord(char) < 128


def find_non_ascii(filepath: str) -> list[tuple[int, int, str]]:
    """
    Scan a file for non-ASCII characters.
    Returns a list of (line_number, col_number, char) tuples.
    """
    issues: list[tuple[int, int, str]] = []
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            for line_no, line in enumerate(f, start=1):
                for col_no, char in enumerate(line, start=1):
                    if not is_ascii(char):
                        issues.append((line_no, col_no, char))
    except Exception as e:
        print(f"  ⚠ Error reading {filepath}: {e}")
    return issues


def main(extensions: str):
    repo_dir = os.environ.get("GITHUB_WORKSPACE", os.getcwd())
    extensions_tuple = tuple(ext.strip() for ext in extensions.split(","))

    files: list[str] = []
    files_with_issues: list[str] = []
    total_issues = 0

    for root, _, filenames in os.walk(repo_dir):
        # Skip .git directory
        if ".git" in root.split(os.sep):
            continue
        for filename in filenames:
            if filename.endswith(extensions_tuple):
                filepath = os.path.join(root, filename)
                files.append(filepath)

                issues = find_non_ascii(filepath)
                if issues:
                    files_with_issues.append(filepath)
                    total_issues += len(issues)
                    rel_path = os.path.relpath(filepath, repo_dir)
                    print(f"🔴 {rel_path} has {len(issues)} non-ASCII character(s):")
                    for line_no, col_no, char in issues:
                        codepoint = ord(char)
                        print(f"    Line {line_no}, col {col_no}: U+{codepoint:04X} ({char!r})")
                else:
                    rel_path = os.path.relpath(filepath, repo_dir)
                    print(f"🟢 {rel_path} is ASCII-clean")

    if not files:
        print("No files with the specified extensions found in the repository.")
        sys.exit(0)

    print(f"\nScanned {len(files)} file(s) with the specified extensions.")

    if files_with_issues:
        print(f"\n🔴 {len(files_with_issues)} file(s) contain non-ASCII characters ({total_issues} total):")
        for f in files_with_issues:
            print(f"  - {os.path.relpath(f, repo_dir)}")
        print("\nVBA files must contain only ASCII characters (U+0000 to U+007F).")
        print("Please replace the non-ASCII characters listed above and try again.")
        sys.exit(2)

    print("✅ All files are ASCII-clean.")


def parse_arguments():
    import argparse
    parser = argparse.ArgumentParser(
        description="Enforce ASCII-only characters in files with specified extensions."
    )
    parser.add_argument(
        "--extensions", type=str, required=True,
        help="Comma-separated list of file extensions to process (e.g. .bas,.frm,.cls)"
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_arguments()
    main(args.extensions)
