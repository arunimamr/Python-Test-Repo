"""
Generate folder structure as JSON for a given directory.

Usage:
    python generate_folder_json.py <directory_path> [--output output.json]

If no directory_path is provided, defaults to current directory.
"""
import os
import sys
import json
from pathlib import Path
from datetime import datetime


def build_tree(path: Path):
    """Recursively build a folder tree starting at path."""
    node = {
        'name': path.name,
        'path': str(path.resolve()),
    }
    try:
        if path.is_dir():
            node['type'] = 'folder'
            children = []
            try:
                for entry in sorted(path.iterdir(), key=lambda p: (not p.is_dir(), p.name.lower())):
                    children.append(build_tree(entry))
            except PermissionError:
                node['error'] = 'PermissionError'
            node['children'] = children
        else:
            node['type'] = 'file'
            stat = path.stat()
            node['size'] = stat.st_size
            node['modified'] = datetime.fromtimestamp(stat.st_mtime).isoformat()
    except Exception as e:
        node['type'] = 'unknown'
        node['error'] = str(e)
    return node


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Generate folder structure JSON')
    parser.add_argument('directory', nargs='?', default='.', help='Directory to scan')
    parser.add_argument('--output', '-o', default=None, help='Output JSON file path')
    args = parser.parse_args()

    dir_path = Path(args.directory).expanduser()
    if not dir_path.exists():
        print(f"ERROR: Directory not found: {dir_path}")
        sys.exit(2)

    print(f"Scanning: {dir_path}")
    tree = build_tree(dir_path)

    # By default save the JSON into the same folder where this script lives
    script_dir = Path(__file__).parent.resolve()
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = script_dir / f"folder_structure_{dir_path.name}.json"

    # Ensure parent directory exists for the chosen output
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(tree, f, indent=2, ensure_ascii=False)

    print(f"Saved JSON to: {output_path}")

if __name__ == '__main__':
    main()

