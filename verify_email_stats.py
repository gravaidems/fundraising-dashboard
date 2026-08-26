#!/usr/bin/env python3
"""
Arithmetic verification of the extracted Email Statistics data.

Like P2P Statistics, the Email Statistics sheet is a single flat detail
table (one row per email send) with no stored monthly or quarterly rollup.
So the checks here split into two kinds: identities that must hold against
the *workbook's own stored cells* (the per-row ratios that are actually
resolvable — Average, Donate Rate, Action Rate, Unsub Rate all recombine
from stored Raised/Donors/Recipients/Actions/Unsubs columns), and identities
that must hold in the rollups *this page computes* (monthly and
quarter-to-date totals), which are regression checks on the rollup
arithmetic itself rather than an audit of the workbook.

Open Rate and Click Rate have no stored opens/clicks count anywhere on this
sheet, so there is no real numerator to check them against — they are
expected to be blank (None) in every rollup rather than averaged, and that
expectation is checked instead of a fabricated identity.

Checks:
  1. Average = Raised / Donors on every detail row (stored cells).
  2. Donate Rate = Donors / Recipients on every detail row.
  3. Action Rate = Actions / Recipients on every detail row.
  4. Unsub Rate = Unsubs / Recipients on every detail row.
  5. Detail rows sum to their computed monthly rollup (additive metrics).
  6. Monthly rollups sum to the computed quarter-to-date totals.
  7. Every rollup ratio is recomputed from summed components, not averaged.
  8. Unresolvable ratios (Open Rate, Click Rate) are blank in every rollup.
  9. Structural sanity: every detail row has a month, date, audience, label.
  10. Cumulative series' final point equals the quarter-to-date total.

Exit code is non-zero on any failure.
"""

import json
import sys

TOL = 0.02
RTOL = 1e-6

results = []


def check(name, ok, detail):
    results.append((ok, name, detail))


def close(a, b, tol=TOL):
    if a is None or b is None:
        return a is None and b is None
    return abs(a - b) <= tol


RATIO_IDENTITIES = [
    ("average", "raised", "donors", "raised/donors"),
    ("donate_rate", "donors", "recipients", "donors/recipients"),
    ("action_rate", "actions", "recipients", "actions/recipients"),
    ("unsub_rate", "unsubs", "recipients", "unsubs/recipients"),
]


def main(path="email-stats-data.json"):
    with open(path) as fh:
        d = json.load(fh)

    metric_order = set(d["metric_order"])

    # --- 1-4. per-row ratio identities against the workbook's own cells ----
    for rk, nk, dk, desc in RATIO_IDENTITIES:
        if rk not in metric_order or nk not in metric_order or dk not in metric_order:
            continue
        for x in d["detail"]:
            M = x["metrics"]
            num, den, ratio = M[nk]["value"], M[dk]["value"], M[rk]["value"]
            label = "detail r%d %s/%s/%s: %s" % (x["row"], x["month"], x["topic"], x["audience"], rk)
            if den:
                check("%s = %s" % (label, desc),
                      close(ratio, num / den, max(TOL, abs(num / den) * 1e-9)),
                      "%s = %r vs %r/%r = %r" % (M[rk]["cell"], ratio, num, den, num / den))
            else:
                check("%s blank when denominator is zero" % label, ratio is None,
                      "denominator=%r so %s should be blank; got %r" % (den, M[rk]["cell"], ratio))

    # --- 6. detail rows sum to their computed monthly rollup ---------------
    for m in d["months"]:
        kids = [x for x in d["detail"] if x["month"] == m["label"]]
        check("month %s has detail rows" % m["label"], len(kids) > 0, "no detail rows matched")
        check("month %s rollup lists exactly its own rows" % m["label"],
              sorted(m["rows"]) == sorted(x["row"] for x in kids),
              "rollup rows %r vs detail rows %r" % (m["rows"], [x["row"] for x in kids]))
        for key in d["additive"]:
            s = sum((x["metrics"][key]["value"] or 0.0) for x in kids)
            stored = m["metrics"][key]["value"]
            check("detail rows sum to month %s %s" % (m["label"], key),
                  close(s, stored, TOL if "recipients" not in key else 0.5),
                  "computed %r vs detail sum %r over %d rows" % (stored, s, len(kids)))

    # --- 7. monthly rollups sum to the quarter-to-date totals --------------
    for key in d["additive"]:
        s = sum((m["metrics"][key]["value"] or 0.0) for m in d["months"])
        stored = d["totals"]["metrics"][key]["value"]
        check("months sum to quarter-to-date total for %s" % key,
              close(s, stored, TOL if "recipients" not in key else 0.5),
              "total %r vs monthly sum %r" % (stored, s))

    # --- 8. rollup ratios are recomputed from summed components ------------
    comps = d["ratio_components"]
    for label, row in [("month " + m["label"], m) for m in d["months"]] + [("quarter to date", d["totals"])]:
        M = row["metrics"]
        for key in d["ratios"]:
            comp = comps.get(key)
            if not comp:
                check("%s: %s has no components, so is blank" % (label, key),
                      M[key]["value"] is None,
                      "expected None for unresolvable ratio %s, got %r" % (key, M[key]["value"]))
                continue
            nk, dk = comp
            nv, dv = M[nk]["value"], M[dk]["value"]
            expected = (nv / dv) if dv else None
            check("%s: %s recomputed from %s/%s" % (label, key, nk, dk),
                  close(M[key]["value"], expected, RTOL),
                  "%r vs %r/%r = %r" % (M[key]["value"], nv, dv, expected))

    # --- 9. unresolvable ratios are consistently blank ----------------------
    for key in d.get("unresolvable_ratios", []):
        for row in d["months"] + [d["totals"]]:
            check("unresolvable ratio %s is blank in %s" % (key, row["label"]),
                  row["metrics"][key]["value"] is None,
                  "expected None, got %r" % row["metrics"][key]["value"])
    check("Open Rate and Click Rate are the only unresolvable ratios",
          set(d.get("unresolvable_ratios", [])) == {"open_rate", "click_rate"},
          "unresolvable_ratios = %r" % d.get("unresolvable_ratios"))

    # --- structural sanity ---------------------------------------------
    check("every detail row has a month, date, audience, mailing topic and label",
          all(x["month"] and x["date"] and x["audience"] and x["topic"] and x["label"] for x in d["detail"]),
          "one or more detail rows missing a dimension")
    check("audiences list matches the detail rows",
          set(x["audience"] for x in d["detail"]) == set(d["audiences"]),
          "detail audiences vs d['audiences'] mismatch")
    check("topics (mailing topic) list matches the detail rows",
          set(x["topic"] for x in d["detail"]) == set(d["topics"]),
          "detail topics vs d['topics'] mismatch")
    check("month_keys count matches months count",
          len(d["month_keys"]) == len(d["months"]),
          "%d month_keys vs %d months" % (len(d["month_keys"]), len(d["months"])))
    check("labels are unique",
          len(set(x["label"] for x in d["detail"])) == len(d["detail"]),
          "duplicate Label value(s) found among detail rows")

    # --- cumulative series: final point equals the quarter-to-date total ---
    for key in d["accum_keys"]:
        final = d["cumulative"]["overall"][key][-1]["value"]
        stored = d["totals"]["metrics"][key]["value"]
        check("cumulative overall %s final point equals quarter-to-date total" % key,
              close(final, stored, TOL if "recipients" not in key else 0.5),
              "cumulative final %r vs total %r" % (final, stored))
        for aud in d["audiences"]:
            series = d["cumulative"]["by_audience"][aud][key]
            rows = [x for x in d["detail"] if x["audience"] == aud]
            s = sum((x["metrics"][key]["value"] or 0.0) for x in rows)
            check("cumulative by_audience[%s][%s] final point equals that audience's row sum" % (aud, key),
                  close(series[-1]["value"], s, TOL if "recipients" not in key else 0.5),
                  "cumulative final %r vs row sum %r" % (series[-1]["value"], s))
        for tp in d["topics"]:
            series = d["cumulative"]["by_topic"][tp][key]
            rows = [x for x in d["detail"] if x["topic"] == tp]
            s = sum((x["metrics"][key]["value"] or 0.0) for x in rows)
            check("cumulative by_topic[%s][%s] final point equals that topic's row sum" % (tp, key),
                  close(series[-1]["value"], s, TOL if "recipients" not in key else 0.5),
                  "cumulative final %r vs row sum %r" % (series[-1]["value"], s))

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
