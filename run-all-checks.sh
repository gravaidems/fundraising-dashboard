#!/bin/sh
# Full pipeline: rebuild the snapshot from the workbook, rebuild the page, then
# run every check. Any failure exits non-zero.
#
#   ./run-all-checks.sh ["Some Other Workbook.xlsx"]
set -e
WB="${1:-EXTERNAL_ CAB Digital Report.xlsx}"
: "${JSDOM_PATH:=/tmp/vfy/node_modules}"
export JSDOM_PATH

echo "== extract (shared JS extractor, same code the browser runs) =="
node build-data.js "$WB"

echo "\n== build page =="
python3 build_dashboard.py

echo "\n== provenance audit (openpyxl, independent of the extractor) =="
python3 verify_provenance.py "$WB"

echo "\n== arithmetic identities =="
python3 verify_digital_report.py digital-report-data.json
python3 verify_paid_media.py paid-media-data.json
python3 verify_p2p.py p2p-data.json
python3 verify_email_stats.py email-stats-data.json
python3 verify_high_dollar.py high-dollar-data.json

echo "\n== rendered page vs workbook =="
node verify_rendered.js digital-report.html digital-report-data.json paid-media-data.json

echo "\n== drop-in path (real + synthetic workbooks) =="
node verify_dropin.js digital-report.html "$WB"

echo "\nall checks passed"
