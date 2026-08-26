#!/usr/bin/env python3
"""
Structural and arithmetic verification of the extracted High-Dollar
Donations data.

This sheet is simpler than the other detail-table tabs: a single additive
metric (amount) and no ratio metrics at all, so there is no ratio identity
to check. It also has no "Year:"/"Quarter:" label cells anywhere in its
header rows, so meta.year/meta.quarter are derived on this page from the
`date` column instead of being read from a labelled cell — that derivation
is checked here directly against the detail rows.

Checks:
  1. Structural sanity: every detail row has receipt_id, amount, date and
     category (category may legitimately be blank on the workbook itself,
     so it is checked for presence of the key, not for a non-empty value).
  2. receipt_id values are unique.
  3. amount values are non-negative numbers.
  4. Detail rows sum to their computed monthly rollup.
  5. Monthly rollups sum to the computed quarter-to-date total.
  6. meta.year/meta.quarter, when set, match every dated detail row (this
     sheet has no labelled cell for either, so this is exactly what "the
     scope banner names the right quarter" is grounded in).
  7. Cumulative series' final point equals the quarter-to-date total, both
     overall and per category.

Exit code is non-zero on any failure.
"""

import json
import sys

TOL = 0.02

results = []


def check(name, ok, detail):
    results.append((ok, name, detail))


def close(a, b, tol=TOL):
    if a is None or b is None:
        return a is None and b is None
    return abs(a - b) <= tol


def main(path="high-dollar-data.json"):
    with open(path) as fh:
        d = json.load(fh)

    detail = d["detail"]
    check("workbook has at least one detail row", len(detail) > 0, "no detail rows")

    # --- 1. structural sanity ------------------------------------------
    check("every detail row has a receipt_id", all(x.get("receipt_id") for x in detail),
          "one or more detail rows missing receipt_id")
    check("every detail row has an amount", all(x["metrics"]["amount"]["value"] is not None for x in detail),
          "one or more detail rows missing amount")
    check("every detail row has a date", all(x.get("date") for x in detail),
          "one or more detail rows missing date")
    check("every detail row has a category key present", all("category" in x for x in detail),
          "one or more detail rows missing the category key entirely")

    # --- 2. receipt_id values are unique --------------------------------
    ids = [x["receipt_id"] for x in detail]
    check("receipt_id values are unique", len(set(ids)) == len(ids),
          "%d rows but only %d distinct receipt_id values" % (len(ids), len(set(ids))))

    # --- 3. amount values are non-negative numbers ----------------------
    amounts = [x["metrics"]["amount"]["value"] for x in detail]
    check("all amount values are non-negative", all(isinstance(a, (int, float)) and a >= 0 for a in amounts),
          "amounts: %r" % amounts)

    # --- 4. detail rows sum to their computed monthly rollup ------------
    for m in d["months"]:
        kids = [x for x in detail if x["month"] == m["label"]]
        check("month %s has detail rows" % m["label"], len(kids) > 0, "no detail rows matched")
        check("month %s rollup lists exactly its own rows" % m["label"],
              sorted(m["rows"]) == sorted(x["row"] for x in kids),
              "rollup rows %r vs detail rows %r" % (m["rows"], [x["row"] for x in kids]))
        s = sum(x["metrics"]["amount"]["value"] or 0.0 for x in kids)
        stored = m["metrics"]["amount"]["value"]
        check("detail rows sum to month %s amount" % m["label"], close(s, stored),
              "computed %r vs detail sum %r over %d rows" % (stored, s, len(kids)))
        check("month %s count matches its own row count" % m["label"], m.get("count") == len(kids),
              "rollup count %r vs %d matching detail rows" % (m.get("count"), len(kids)))

    # --- 5. monthly rollups sum to the quarter-to-date total ------------
    s = sum(m["metrics"]["amount"]["value"] or 0.0 for m in d["months"])
    stored = d["totals"]["metrics"]["amount"]["value"]
    check("months sum to quarter-to-date total for amount", close(s, stored),
          "total %r vs monthly sum %r" % (stored, s))
    check("quarter-to-date total equals the sum of every detail row's amount",
          close(sum(amounts), stored), "sum(detail amounts) %r vs totals %r" % (sum(amounts), stored))
    check("quarter-to-date count equals the number of detail rows", d["totals"].get("count") == len(detail),
          "totals count %r vs %d detail rows" % (d["totals"].get("count"), len(detail)))

    # --- 6. derived year/quarter match every dated detail row -----------
    dated = [x for x in detail if x.get("year") is not None]
    if d["meta"]["year"]["value"] is not None:
        check("meta.year matches every dated detail row",
              all(x["year"] == d["meta"]["year"]["value"] for x in dated),
              "meta.year=%r but detail years=%r" % (d["meta"]["year"]["value"], sorted(set(x["year"] for x in dated))))
        check("meta.quarter matches every dated detail row",
              all(x["quarter"] == d["meta"]["quarter"]["value"] for x in dated),
              "meta.quarter=%r but detail quarters=%r" %
              (d["meta"]["quarter"]["value"], sorted(set(x["quarter"] for x in dated))))
        check("meta.year/meta.quarter are marked as derived, not read from a labelled cell",
              d["meta"]["year"].get("derived") is True and d["meta"]["quarter"].get("derived") is True,
              "meta.year=%r meta.quarter=%r" % (d["meta"]["year"], d["meta"]["quarter"]))
    else:
        check("meta.year is blank only because detail rows span more than one quarter",
              len(set((x["year"], x["quarter"]) for x in dated)) != 1,
              "meta.year is blank but detail rows all share one year/quarter — should have been derived")

    # --- 7. cumulative series: final point equals the quarter-to-date total ---
    final = d["cumulative"]["overall"]["amount"][-1]["value"]
    check("cumulative overall amount final point equals quarter-to-date total",
          close(final, stored), "cumulative final %r vs total %r" % (final, stored))
    for cat in d["categories"]:
        series = d["cumulative"]["by_category"][cat]["amount"]
        rows = [x for x in detail if x["category"] == cat]
        s = sum(x["metrics"]["amount"]["value"] or 0.0 for x in rows)
        check("cumulative by_category[%s] final point equals that category's row sum" % cat,
              close(series[-1]["value"], s), "cumulative final %r vs row sum %r" % (series[-1]["value"], s))
    check("categories list matches the detail rows (ignoring blank category)",
          set(x["category"] for x in detail if x["category"]) == set(d["categories"]),
          "detail categories vs d['categories'] mismatch")

    failures = [r for r in results if not r[0]]
    print("%d checks run, %d passed, %d failed"
          % (len(results), len(results) - len(failures), len(failures)))
    for ok, name, detail_msg in failures:
        print("  FAIL  %s\n        %s" % (name, detail_msg))
    if not failures:
        print("all checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
