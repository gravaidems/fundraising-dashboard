#!/usr/bin/env python3
"""
Arithmetic verification of the extracted Paid Media Report data.

Checks:
  1. Gross CPA = Gross Spend / NTL on every row (quarterly, monthly, detail).
  2. Each ROAS = its Raised figure / Gross Spend on every row.
  3. Detail rows sum to their month's topline row (additive metrics).
  4. Monthly rows sum to the quarterly row for the quarter the tab is set to.
  5. Total Raised is NOT Immediate + Lifetime (documents the de-dup in G9,
     so a future change that silently "fixes" it gets caught).
  6. The year-to-date rollup sums the additive metrics and RECOMPUTES ratios.

Exit code is non-zero on any failure.
"""

import json
import sys

TOL = 0.02      # dollars
RTOL = 1e-6     # ratios

results = []


def check(name, ok, detail):
    results.append((ok, name, detail))


def close(a, b, tol=TOL):
    if a is None or b is None:
        return a is None and b is None
    return abs(a - b) <= tol


def rows_of(d):
    """Every metric-bearing row, labelled for error messages."""
    out = []
    for q in d["quarters"]:
        out.append(("quarter " + q["label"], q))
    for m in d["months"]:
        out.append(("month " + m["label"], m))
    for x in d["detail"]:
        out.append(("detail r%d %s / %s / %s" % (x["row"], x["month"], x["channel"], x["placement"]), x))
    return out


def main(path="paid-media-data.json"):
    with open(path) as fh:
        d = json.load(fh)

    # --- 1 & 2. per-row ratio identities --------------------------------
    for label, row in rows_of(d):
        M = row["metrics"]
        spend, ntl, cpa = M["gross_spend"]["value"], M["ntl"]["value"], M["gross_cpa"]["value"]
        if ntl:
            check("%s: CPA = spend/NTL" % label,
                  close(cpa, spend / ntl, max(TOL, abs(spend / ntl) * 1e-9)),
                  "%s = %r vs %r/%r = %r" % (M["gross_cpa"]["cell"], cpa, spend, ntl,
                                             spend / ntl if ntl else None))
        else:
            check("%s: CPA blank when NTL is zero" % label, cpa is None,
                  "NTL=%r so %s should be blank; got %r" % (ntl, M["gross_cpa"]["cell"], cpa))

        for rk, raised_key in [("immediate_roas", "immediate_raised"),
                               ("lifetime_roas", "lifetime_raised"),
                               ("total_roas", "total_raised")]:
            roas, raised = M[rk]["value"], M[raised_key]["value"]
            if spend:
                check("%s: %s = %s/spend" % (label, rk, raised_key),
                      close(roas, raised / spend, RTOL),
                      "%s = %r vs %r/%r = %r" % (M[rk]["cell"], roas, raised, spend,
                                                 raised / spend if spend else None))

    # --- 3. detail rows sum to their month ------------------------------
    for m in d["months"]:
        kids = [x for x in d["detail"] if x["month"].strip() == m["label"].strip()]
        check("month %s has detail rows" % m["label"], len(kids) > 0, "no detail rows matched")
        for key in d["additive"]:
            s = sum((x["metrics"][key]["value"] or 0.0) for x in kids)
            stored = m["metrics"][key]["value"]
            check("detail rows sum to month %s %s" % (m["label"], key),
                  close(s, stored, TOL if key != "ntl" else 0.5),
                  "%s = %r vs detail sum %r over %d rows"
                  % (m["metrics"][key]["cell"], stored, s, len(kids)))

    # --- 4. months sum to the selected quarter --------------------------
    qn = int(d["meta"]["quarter"]["value"])
    q = next((x for x in d["quarters"] if x["quarter"] == qn), None)
    check("selected quarter Q%d present in quarterly table" % qn, q is not None,
          "quarters present: %s" % [x["quarter"] for x in d["quarters"]])
    if q:
        for key in d["additive"]:
            s = sum((m["metrics"][key]["value"] or 0.0) for m in d["months"])
            stored = q["metrics"][key]["value"]
            check("months sum to Q%d %s" % (qn, key),
                  close(s, stored, TOL if key != "ntl" else 0.5),
                  "%s = %r vs monthly sum %r" % (q["metrics"][key]["cell"], stored, s))
        # ratios of the selected quarter must also hold against the monthly sums
        sp = sum((m["metrics"]["gross_spend"]["value"] or 0.0) for m in d["months"])
        tr = sum((m["metrics"]["total_raised"]["value"] or 0.0) for m in d["months"])
        check("Q%d Total ROAS matches monthly sums" % qn,
              close(q["metrics"]["total_roas"]["value"], tr / sp, 1e-6),
              "%s = %r vs %r/%r" % (q["metrics"]["total_roas"]["cell"],
                                    q["metrics"]["total_roas"]["value"], tr, sp))

    # --- 5. Total Raised is deliberately not Immediate + Lifetime -------
    for label, row in rows_of(d):
        M = row["metrics"]
        i, l, t = M["immediate_raised"]["value"], M["lifetime_raised"]["value"], M["total_raised"]["value"]
        if i is None or l is None or t is None:
            continue
        # t must never exceed i + l (immediate is removed from NTL, not added)
        check("%s: Total Raised <= Immediate + Lifetime" % label, t <= i + l + TOL,
              "%s = %r exceeds %r + %r = %r" % (M["total_raised"]["cell"], t, i, l, i + l))
    qsel = next((x for x in d["quarters"] if x["quarter"] == qn), None)
    if qsel:
        M = qsel["metrics"]
        i, l, t = (M["immediate_raised"]["value"], M["lifetime_raised"]["value"],
                   M["total_raised"]["value"])
        check("Q%d Total Raised differs from Immediate + Lifetime (de-dup per G9)" % qn,
              abs(t - (i + l)) > TOL,
              "expected a de-dup gap; t=%r i+l=%r" % (t, i + l))

    # --- 6. year-to-date rollup ----------------------------------------
    y = d["ytd"]["metrics"]
    for key in d["additive"]:
        s = sum((q["metrics"][key]["value"] or 0.0) for q in d["quarters"])
        check("YTD %s sums the quarterly rows" % key, close(y[key]["value"], s, 0.5),
              "ytd %r vs sum %r" % (y[key]["value"], s))
    sp, ntl = y["gross_spend"]["value"], y["ntl"]["value"]
    check("YTD Gross CPA recomputed from summed spend and NTL",
          close(y["gross_cpa"]["value"], sp / ntl, 1e-9),
          "ytd %r vs %r/%r" % (y["gross_cpa"]["value"], sp, ntl))
    for rk, raised_key in [("immediate_roas", "immediate_raised"),
                           ("lifetime_roas", "lifetime_raised"),
                           ("total_roas", "total_raised")]:
        check("YTD %s recomputed from summed components" % rk,
              close(y[rk]["value"], y[raised_key]["value"] / sp, 1e-9),
              "ytd %r vs %r/%r" % (y[rk]["value"], y[raised_key]["value"], sp))
    # a YTD ratio must not equal the naive average of the quarterly ratios
    naive = sum(q["metrics"]["total_roas"]["value"] for q in d["quarters"]) / len(d["quarters"])
    check("YTD Total ROAS is not the naive average of quarterly ROAS",
          abs(y["total_roas"]["value"] - naive) > 1e-6,
          "ytd %r vs naive average %r — ratios must be recomputed, not averaged"
          % (y["total_roas"]["value"], naive))

    # --- structural sanity ---------------------------------------------
    check("every detail row has month, channel, placement and objective",
          all(x["month"] and x["channel"] and x["placement"] and x["objective"] for x in d["detail"]),
          "one or more detail rows missing a dimension")
    check("detail months are a subset of the monthly toplines",
          set(x["month"].strip() for x in d["detail"]) <= set(m["label"].strip() for m in d["months"]),
          "detail months: %s vs toplines: %s"
          % (sorted(set(x["month"].strip() for x in d["detail"])),
             sorted(set(m["label"].strip() for m in d["months"]))))

    failures = [r for r in results if not r[0]]
    print("%d checks run, %d passed, %d failed"
          % (len(results), len(results) - len(failures), len(failures)))
    for ok, name, detail in failures:
        print("  FAIL  %s\n        %s" % (name, detail))
    if not failures:
        print("all checks passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
