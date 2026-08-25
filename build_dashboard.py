#!/usr/bin/env python3
"""
Build a self-contained digital-report.html from digital-report-data.json.

The data is inlined into the page, so the resulting file opens offline with no
server and no external requests. Charts are hand-rolled inline SVG for the same
reason. Regenerate by re-running extract -> verify -> build.
"""

import argparse
import json
import os
import re

TEMPLATE_PATH = "dashboard_template.html"

# Which workbook tab ultimately feeds each (period-kind, field) pair. Derived
# from the Email row's IMPORTRANGE formulas, which are the only ones the xlsx
# preserved in full; every other row in the same column shares the query.
SOURCE_PATTERNS = [
    "Import - Quarterly Goals",
    "Quarterly Fundraising Toplines",
    "Import - Monthly Goals",
    "Monthly Fundraising Toplines",
]


def derive_sources(channels):
    """Read the Email row formulas to label each column's upstream source tab."""
    email = next((c for c in channels if c["row"] == 14), None)
    sources = {}
    if not email:
        return sources
    for pkey, p in email["periods"].items():
        for field in ("goal", "actual"):
            f = p.get(field + "_formula") or ""
            hit = next((s for s in SOURCE_PATTERNS if s in f), None)
            if hit:
                sources["%s_%s" % (pkey, field)] = hit
    return sources


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-d", "--data", default="digital-report-data.json")
    ap.add_argument("-p", "--paid-media", default="paid-media-data.json")
    ap.add_argument("-t", "--template", default=TEMPLATE_PATH)
    ap.add_argument("-o", "--out", default="digital-report.html")
    args = ap.parse_args()

    with open(args.data) as fh:
        data = json.load(fh)

    data["sources"] = derive_sources(data["channels"])

    # A month is 'pending' when no channel has reported an actual for it. The
    # Overall row shows a computed 0 for such months; we must not render that
    # as "raised nothing".
    pending = []
    for m in data["month_keys"]:
        if all(ch["periods"][m]["actual"] is None for ch in data["channels"]):
            pending.append(m)
    data["pending_months"] = pending

    # Persist the enriched payload so the JSON audit trail on disk is exactly
    # what the page consumes (verify_rendered.js reads it back).
    with open(args.data, "w") as fh:
        json.dump(data, fh, indent=2)

    with open(args.template) as fh:
        html = fh.read()

    def inject(marker, payload):
        nonlocal html
        blob = json.dumps(payload, indent=None, separators=(",", ":"))
        blob = blob.replace("</", "<\\/")  # keep a stray </script> from closing the tag
        if marker not in html:
            raise SystemExit("template is missing the %s marker" % marker)
        html = html.replace(marker, blob)

    inject("/*__DATA__*/", data)

    pm = None
    if os.path.exists(args.paid_media):
        with open(args.paid_media) as fh:
            pm = json.load(fh)
    inject("/*__PM_DATA__*/", pm)

    with open(args.out, "w") as fh:
        fh.write(html)

    print("wrote %s (%.1f KB, %d pending month(s): %s; paid media: %s)"
          % (args.out, len(html) / 1024.0, len(pending), ", ".join(pending) or "none",
             "%d quarters / %d months / %d detail rows"
             % (len(pm["quarters"]), len(pm["months"]), len(pm["detail"])) if pm else "absent"))


if __name__ == "__main__":
    main()
