"""
Command-line interface for the Markdown → DOCX converter.

Usage:
    py cli.py input.md                      # outputs input.docx
    py cli.py input.md -o report.docx       # custom output path
    py cli.py input.md output.docx          # positional output path
"""

import sys
import os
import argparse

# Fix Unicode display on Windows terminals (emoji in cp1252)
if sys.platform == "win32":
    sys.stdout = open(sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1)

from converter import convert_markdown_to_docx


def main():
    parser = argparse.ArgumentParser(
        description="Convert a Markdown file to a formatted Word (.docx) document.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  py cli.py README.md
  py cli.py notes.md -o report.docx
  py cli.py docs/guide.md -o output/guide.docx
        """
    )
    parser.add_argument("input",  help="Path to the input .md file")
    parser.add_argument("output", nargs="?", help="Path for the output .docx file (optional)")
    parser.add_argument("-o", "--out", dest="out_flag", help="Output .docx path (alternative to positional)")

    args = parser.parse_args()

    # Resolve input
    input_path = args.input
    if not os.path.isfile(input_path):
        print(f"❌  Error: File not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    # Resolve output
    output_path = args.out_flag or args.output
    if not output_path:
        base = os.path.splitext(input_path)[0]
        output_path = base + ".docx"

    # Ensure output directory exists
    out_dir = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(out_dir, exist_ok=True)

    # Read markdown
    try:
        with open(input_path, "r", encoding="utf-8") as f:
            markdown_text = f.read()
    except Exception as e:
        print(f"❌  Error reading file: {e}", file=sys.stderr)
        sys.exit(1)

    if not markdown_text.strip():
        print("⚠️   Warning: Input file is empty.", file=sys.stderr)

    # Convert
    print(f"⚙️   Converting: {input_path}")
    try:
        docx_bytes = convert_markdown_to_docx(markdown_text)
    except Exception as e:
        print(f"❌  Conversion error: {e}", file=sys.stderr)
        sys.exit(1)

    # Write output
    try:
        with open(output_path, "wb") as f:
            f.write(docx_bytes)
    except Exception as e:
        print(f"❌  Error writing output: {e}", file=sys.stderr)
        sys.exit(1)

    size_kb = len(docx_bytes) / 1024
    print(f"✅  Done! → {output_path}  ({size_kb:.1f} KB)")


if __name__ == "__main__":
    main()
