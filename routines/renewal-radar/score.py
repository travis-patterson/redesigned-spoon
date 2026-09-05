#!/usr/bin/env python3
"""Rank a 90-day renewal book by ARR at risk.

Deterministic scoring so the same book produces the same ranking week to week.
Reads three files written by the routine from Salesforce query results:

  accounts.json   [{id,name,end,arr,commit,owner}, ...]      the renewal book
  snapshots.json  raw SOQL result for the utilization snapshots
  coverage.json   raw SOQL result for the per-account snapshot census
                  (COUNT(Id), MAX(Date_Month_of_Snapshot__c) grouped by account)

Usage: python3 score.py accounts.json snapshots.json coverage.json
"""
import json
import sys
import datetime

RISING, DECLINING = 0.05, -0.05
OVERAGE = 1.10          # above this, overage outranks trend when setting posture
MIN_RATEABLE_COMMIT = 5  # below this a per-unit rate is noise, not a price
STALE_DAYS = 62
TREND_MONTHS = 3


def money(v):
    return "$%s" % format(v, ",.0f")


def load(path):
    with open(path) as fh:
        return json.load(fh)


def records(path):
    return load(path).get("records", [])


def main(accounts_path, snapshots_path, coverage_path):
    today = datetime.date.today()
    accounts = load(accounts_path)

    by_account = {}
    for r in records(snapshots_path):
        by_account.setdefault(r["Account__c"], []).append(r)
    for rows in by_account.values():
        rows.sort(key=lambda r: r["Date_Month_of_Snapshot__c"])

    # Distinguish "never instrumented" from "instrumentation stopped". Both are
    # blind spots but they have different owners and different fixes.
    coverage = {r["Account__c"]: r for r in records(coverage_path)}

    rows, frozen, instrumented = [], 0, 0

    for a in accounts:
        end = datetime.date(*map(int, a["end"].split("-")))
        arr, commit = float(a["arr"]), float(a["commit"])
        row = dict(name=a["name"], owner=a["owner"], end=a["end"],
                   days=(end - today).days, arr=arr, commit=commit,
                   latest=None, util=None, trend=None, posture=None,
                   risk=0.0, shelf=0.0, recover=0, flags=[])
        snaps = by_account.get(a["id"], [])

        if not snaps:
            cov = coverage.get(a["id"])
            if cov and cov.get("snaps"):
                row["flags"].append("snapshots stopped %s" % cov.get("latest"))
            else:
                row["flags"].append("never instrumented, zero snapshots on record")
            row.update(posture="BLIND", risk=arr)
            rows.append(row)
            continue

        instrumented += 1
        latest = snaps[-1]
        in_use = latest["Number_of_Linked_Accounts__c"] or 0
        paused = latest["Number_of_Paused_Linked_Accounts__c"] or 0
        relink = latest["Number_of_Relink_Needed_Linked_Accounts__c"] or 0

        # Direction comes from the trailing 3 months, not the full window. A
        # 7-month baseline often lands mid-onboarding, so the whole book reads
        # as growth that actually finished months ago.
        window = snaps[-(TREND_MONTHS + 1):]
        base = window[0]["Number_of_Linked_Accounts__c"] or 0
        trend = ((in_use - base) / base) if base else None
        direction = ("rising" if (trend or 0) > RISING else
                     "declining" if (trend or 0) < DECLINING else "flat")

        util = (in_use / commit) if commit > 0 else None
        if commit <= 0:
            row["flags"].append("contracted units = %.0f, utilization not computable" % commit)

        month = datetime.date(*map(int, latest["Date_Month_of_Snapshot__c"].split("-")))
        if (today - month).days > STALE_DAYS:
            row["flags"].append("latest snapshot %s, stale" % latest["Date_Month_of_Snapshot__c"])

        paused_series = [s["Number_of_Paused_Linked_Accounts__c"] for s in snaps]
        if len(paused_series) >= 4 and len(set(paused_series[-4:])) == 1 and paused_series[-1]:
            frozen += 1

        if util is None:
            posture = "UNKNOWN"
        elif util > OVERAGE:
            posture = "Expansion"
        elif util < 0.60:
            posture = "At risk" if direction == "declining" else "Retention"
        elif util <= 0.85:
            posture = "Growth" if direction == "rising" else "Hold"
        else:
            posture = "Expansion" if direction == "rising" else "Hold"

        rate = arr / commit if commit > 0 else 0
        shelf = max(0.0, (commit - in_use) * rate) if commit > 0 else 0.0

        if posture in ("At risk", "Retention", "UNKNOWN"):
            risk = arr                      # revenue that may not renew
        elif posture == "Expansion":
            if commit < MIN_RATEABLE_COMMIT:
                risk = 0.0
                row["flags"].append(
                    "commit of %.0f units too small to derive a rate, upside not sized" % commit)
            else:
                risk = max(0.0, (in_use - commit) * rate)   # upside forgone
        else:
            risk = 0.0

        row.update(latest=in_use, util=util, trend=trend, posture=posture,
                   risk=risk, shelf=shelf, recover=paused + relink)
        rows.append(row)

    rows.sort(key=lambda r: (-r["risk"], r["days"]))

    head = ("%-30s%-15s%-12s%4s%9s%6s%6s%6s%7s  %-10s%9s" %
            ("ACCOUNT", "OWNER", "RENEWS", "D", "ARR", "CMT", "USE", "UTIL", "3MO",
             "POSTURE", "AT RISK"))
    out = [head, "-" * len(head)]
    for r in rows:
        out.append("%-30s%-15s%-12s%4d%9s%6.0f%6.0f%6s%7s  %-10s%9s" % (
            r["name"][:29], r["owner"][:14], r["end"], r["days"],
            format(r["arr"], ",.0f"), r["commit"], r["latest"] or 0,
            ("%.0f%%" % (r["util"] * 100)) if r["util"] is not None else "-",
            ("%+.0f%%" % (r["trend"] * 100)) if r["trend"] is not None else "-",
            r["posture"], format(r["risk"], ",.0f")))

    flagged = [r for r in rows if r["risk"] > 0]
    out += ["",
            "ARR up for renewal in 90d  $%12s  (%d accounts)" % (
                format(sum(r["arr"] for r in rows), ",.0f"), len(rows)),
            "Flagged needing attention  $%12s  (%d accounts)" % (
                format(sum(r["risk"] for r in flagged), ",.0f"), len(flagged)),
            "Shelfware in the book      $%12s" % format(sum(r["shelf"] for r in rows), ",.0f")]

    if frozen:
        out += ["", "SYSTEMIC: paused-unit count frozen for 4+ months on %d of %d instrumented "
                    "accounts. Treat the field as unreliable org-wide; do not quote recoverable-unit "
                    "figures to a customer." % (frozen, instrumented)]

    defects = [(r["name"], f) for r in rows for f in r["flags"]]
    if defects:
        out += ["", "PER-ACCOUNT DEFECTS"]
        out += ["  %s: %s" % d for d in defects]

    print("\n".join(out))
    with open("radar.json", "w") as fh:
        json.dump(rows, fh, indent=2)

    # Slack delivery. Slack renders no HTML and silently truncates long text, so
    # the report is split: a lead message that stands on its own, then the full
    # table in chunks small enough to post as threaded replies.
    lead = ["*Renewal radar* \u2014 %s in the 90-day book across %d accounts, "
            "%s flagged." % (money(sum(r["arr"] for r in rows)), len(rows),
                             money(sum(r["risk"] for r in flagged)))]
    blind = [r for r in rows if r["posture"] == "BLIND"]
    if blind:
        lead.append("%s renews with no usage data at all: %s."
                    % (money(sum(r["arr"] for r in blind)),
                       ", ".join(r["name"] for r in blind)))
    lead.append("")
    lead.append("Run /contract on these:")
    for r in flagged[:5]:
        # Show the at-risk figure, not ARR: at-risk is the sort key, and on an
        # expansion row it is upside forgone rather than the account's ARR, so
        # printing ARR here makes the ordering look wrong.
        lead.append("  \u2022 *%s* \u2014 %s at risk, %s ARR, renews %s (%dd), %s"
                    % (r["name"], money(r["risk"]), money(r["arr"]), r["end"],
                       r["days"], r["posture"]))
    if frozen:
        lead.append("")
        lead.append("_Paused-unit field is frozen on %d of %d instrumented "
                    "accounts; recoverable-unit figures are not quotable._"
                    % (frozen, instrumented))
    with open("slack-lead.txt", "w") as fh:
        fh.write("\n".join(lead))

    # Chunk the monospace table. 3500 chars keeps each reply clear of Slack's
    # truncation point once the code fence is added.
    chunks, cur = [], []
    size = 0
    for line in out:
        if size + len(line) + 1 > 3500 and cur:
            chunks.append("\n".join(cur))
            cur, size = [head, "-" * len(head)], len(head) * 2
        cur.append(line)
        size += len(line) + 1
    if cur:
        chunks.append("\n".join(cur))
    for i, ch in enumerate(chunks):
        with open("slack-table-%02d.txt" % (i + 1), "w") as fh:
            fh.write("```\n%s\n```" % ch)
    print("slack: 1 lead + %d table chunk(s)" % len(chunks))


if __name__ == "__main__":
    main(*sys.argv[1:4])
