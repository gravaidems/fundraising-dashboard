#!/usr/bin/env python3
"""
Arithmetic verification of the extracted Digital Report data.

Checks, in order:
  1. Every '% to Goal' cell equals Actual / Goal.
  2. Quarterly goal C = SUM(Q:S) for every channel row.
  3. The 'Display Dynamic Goals' toggles (O9/O10) are FALSE, so Q/R/S mirror
     the monthly goals G/J/M rather than the actuals.
  4. Channel rows sum to the Overall 'Total Raised' row, per period.
  5. Average Donation = Total Raised / Total Donations.

Exit code is non-zero if any check fails, so this can gate the build.
"""

import json
import sys

TOL = 0.02          # currency tolerance, in dollars
PCT_TOL = 1e-6      # ratio tolerance

results = []


def check(name, ok, detail):
    results.append((ok, name, detail))


def close(a, b, tol=TOL):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


def main(path="digital-report-data.json"):
    with open(path) as fh:
        d = json.load(fh)

    channels = d["channels"]
    months = d["month_keys"]

    # --- 1. percentage recomputation ------------------------------------
    for ch in channels:
        for pkey, p in ch["periods"].items():
            g, a, pct = p["goal"], p["actual"], p["pct"]
            if not g:
                # No goal: workbook stores the '-' sentinel, extractor gives None.
                check(
                    "%s / %s pct sentinel" % (ch["name"], pkey),
                    pct is None,
                    "goal=%r so %s should be '-'; got %r" % (g, p["pct_cell"], p["pct_raw"]),
                )
                continue
            if a is None:
                continue  # month not yet reported
            expected = a / g
            check(
                "%s / %s pct" % (ch["name"], pkey),
                close(expected, pct, PCT_TOL),
                "%s: %s/%s = %.10f vs stored %.10f"
                % (p["pct_cell"], p["actual_cell"], p["goal_cell"], expected,
                   pct if pct is not None else float("nan")),
            )

    # --- 2. quarterly goal = SUM of helper columns ----------------------
    for ch in channels:
        helper_sum = sum((ch["helpers"][m]["value"] or 0.0) for m in months)
        q_goal = ch["periods"]["quarter"]["goal"] or 0.0
        check(
            "%s quarterly goal = SUM(Q:S)" % ch["name"],
            close(helper_sum, q_goal),
            "%s = %.2f vs SUM(%s) = %.2f"
            % (ch["periods"]["quarter"]["goal_cell"], q_goal,
               ",".join(ch["helpers"][m]["cell"] for m in months), helper_sum),
        )

    # --- 3. dynamic goal toggles off, helpers mirror monthly goals ------
    dg = d["meta"]["dynamic_goals"]
    toggles_off = not dg["july"]["value"] and not dg["august"]["value"]
    check(
        "dynamic goal toggles are FALSE",
        toggles_off,
        "O9=%r O10=%r" % (dg["july"]["value"], dg["august"]["value"]),
    )
    if toggles_off:
        for ch in channels:
            for m in months:
                check(
                    "%s helper %s mirrors monthly goal" % (ch["name"], m),
                    close(ch["helpers"][m]["value"] or 0.0,
                          ch["periods"][m]["goal"] or 0.0),
                    "%s = %r vs %s = %r"
                    % (ch["helpers"][m]["cell"], ch["helpers"][m]["value"],
                       ch["periods"][m]["goal_cell"], ch["periods"][m]["goal"]),
                )

    # --- 4. channels sum to Overall Total Raised ------------------------
    tr = d["overall"]["total_raised"]
    for pkey in ["quarter"] + months:
        op = tr["periods"][pkey]
        for field in ("goal", "actual"):
            ch_sum = sum((ch["periods"][pkey][field] or 0.0) for ch in channels)
            stored = op[field]
            if stored is None:
                check(
                    "Overall %s %s blank while channels sum to %.2f" % (pkey, field, ch_sum),
                    ch_sum == 0.0,
                    "%s is blank but channel sum is %.2f" % (op[field + "_cell"], ch_sum),
                )
                continue
            check(
                "channels sum to Overall %s %s" % (pkey, field),
                close(ch_sum, stored),
                "%s = %.2f vs channel sum = %.2f (diff %.4f)"
                % (op[field + "_cell"], stored, ch_sum, ch_sum - stored),
            )

    # --- 5. average donation -------------------------------------------
    td = d["overall"]["total_donations"]
    ad = d["overall"]["average_donation"]
    for pkey in ["quarter"] + months:
        raised = tr["periods"][pkey]["actual"]
        count = td["periods"][pkey]["actual"]
        stored = ad["periods"][pkey]["actual"]
        if not raised or not count or stored is None:
            continue
        check(
            "average donation %s" % pkey,
            close(raised / count, stored, 1e-6),
            "%s: %.2f/%.0f = %.8f vs stored %.8f"
            % (ad["periods"][pkey]["actual_cell"], raised, count, raised / count, stored),
        )

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
