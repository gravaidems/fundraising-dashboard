#!/usr/bin/env python3
"""
build_dashboard.py — assemble the self-contained digital-report.html.

The page ships with NO data in it. It is a viewer: the reader loads an .xlsx and
the page parses it in the browser. That keeps client figures out of the shared
file entirely, and means the page can never display stale numbers.

What gets inlined:
  lib/xlsx-reader.js      ZIP + XML reader
  lib/extract-report.js   the extractor, the same one node uses at build time

Usage:
    python3 build_dashboard.py [-t dashboard_template.html] [-o digital-report.html]
"""

import argparse
import os

LIBS = [
    ("/*__LIB_XLSX_READER__*/", "lib/xlsx-reader.js"),
    ("/*__LIB_EXTRACT_REPORT__*/", "lib/extract-report.js"),
]

# The page starts with nothing loaded. These markers exist so the payload slots
# are explicit rather than implied.
DATA_MARKERS = ["/*__DATA__*/", "/*__PM_DATA__*/"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-t", "--template", default="dashboard_template.html")
    ap.add_argument("-o", "--out", default="digital-report.html")
    args = ap.parse_args()

    with open(args.template) as fh:
        html = fh.read()

    for marker, path in LIBS:
        if marker not in html:
            raise SystemExit("template is missing the %s marker" % marker)
        if not os.path.exists(path):
            raise SystemExit("missing library file: %s" % path)
        with open(path) as fh:
            src = fh.read()
        # inlining would break out of the script element
        if "</script" in src.lower():
            raise SystemExit("%s contains a closing script tag and cannot be inlined" % path)
        html = html.replace(marker, src)

    for marker in DATA_MARKERS:
        if marker not in html:
            raise SystemExit("template is missing the %s marker" % marker)
        html = html.replace(marker, "null")

    # A shipped page must not carry data. Check structurally: the payload slots
    # must be literal nulls, and no serialised payload key may appear anywhere.
    # (Words like "IMPORTRANGE" legitimately occur in the extractor's own source,
    # so matching on prose would give false positives.)
    for expect in ("let DATA = null;", "let PM = null;"):
        if expect not in html:
            raise SystemExit("refusing to write %s: expected %r in the output" % (args.out, expect))
    for key in ('"source_workbook"', '"month_keys"', '"metric_order"', '"pending_months"'):
        if key in html:
            raise SystemExit(
                "refusing to write %s: found the payload key %s, so workbook data has leaked in"
                % (args.out, key))

    with open(args.out, "w") as fh:
        fh.write(html)

    print("wrote %s (%.1f KB, no data baked in — loads a workbook at runtime)"
          % (args.out, len(html) / 1024.0))


if __name__ == "__main__":
    main()
