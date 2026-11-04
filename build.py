#!/usr/bin/env python3
"""
Bundle the editor into ONE self-contained HTML file.

The single file has ExcelJS, SheetJS, JsBarcode, the styles and the app code
inlined, so it works with no server, no internet connection and no CDN.

Usage:
    python3 build.py
Output:
    dist/putwall-trouble-editor.html
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "dist", "putwall-trouble-editor.html")


def read(path):
    with open(os.path.join(ROOT, path), "r", encoding="utf-8") as fh:
        return fh.read()


def safe_js(code):
    # Prevent an embedded "</script>" from terminating the inline tag early.
    return re.sub(r"</script", r"<\\/script", code, flags=re.IGNORECASE)


def main():
    html = read("index.html")

    # Inline the stylesheet.
    css_tag = '<link rel="stylesheet" href="src/styles.css" />'
    if css_tag not in html:
        print("Could not find the stylesheet link in index.html", file=sys.stderr)
        return 1
    html = html.replace(css_tag, "<style>\n" + read("src/styles.css") + "\n</style>")

    # Inline every local <script src="..."></script>.
    def repl(match):
        src = match.group(1)
        code = read(src)
        return "<script>\n" + safe_js(code) + "\n</script>"

    html, count = re.subn(
        r'<script src="([^"]+)"></script>', repl, html
    )
    if count == 0:
        print("No local script tags found to inline", file=sys.stderr)
        return 1

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write(html)

    size = os.path.getsize(OUT)
    print("Wrote %s (%.2f MB, inlined %d scripts)" % (OUT, size / 1e6, count))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
