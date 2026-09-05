# ROUTINE: weekly-trend-dashboard
# Schedule: 0 16 * * 5  (Fridays 12:00 ET during EDT)
# Runtime: Claude + Agent Handler (Google Sheets, Drive, Gmail). Source sheet is read-only.

GOAL
Chart the weekly GTM trends from the WEEKLY tab of the GTM Weekly Metrics sheet,
publish the dashboard to Drive, and email the link with a short read of what moved.

HARD GUARDRAILS  (violate any one of these -> stop and report)
- NEVER call the Artifact tool. It requires an interactive approval that nobody is
  present to give on a scheduled run, and the routine will hang at the prompt
  instead of failing. This is what broke the previous version of this routine.
  The same applies to any other tool that opens a confirmation: if a tool asks,
  the run is already wrong. Deliver through Drive and Gmail only.
- The source spreadsheet is READ ONLY. No update_values, append_values,
  batch_update, add_sheet, or any other write to spreadsheet
  1FCFjss1No-3f2Su93ahKcIIbzzQREX6nv4AjgHktadk.
- Do not set sharing permissions on the Drive file and do not publish it to the
  web. It lands in the owner's own Drive and inherits default access.
- Gmail: the only permitted write is a single message to travis@merge.dev.

TOOLING
- Sheets, Drive and Gmail all go through Agent Handler. Load exact schemas first:
  tool_search("agent handler google sheets get values drive create file gmail send")
- Fall back to a native connector only if an Agent Handler call fails twice, and
  say so in the final report.

SHEET
  id:  1FCFjss1No-3f2Su93ahKcIIbzzQREX6nv4AjgHktadk   ("GTM Weekly Metrics")
  tab: WEEKLY  (titled "GTM Weekly Metrics - COMBINED")

STEP 1 - ESTABLISH THE WINDOW
Read `WEEKLY!D5`. That is the as-of date the sheet was last updated (e.g. 8/31/2026).

Week start dates live in row 13, running left to right from column F. Read
`WEEKLY!F13:FR13`.

Three traps in that row, all of which have to be handled or the chart lies:
- Weeks continue PAST the as-of date as empty placeholders. A naive read charts
  them as zeros and the series appears to collapse.
- After the weekly run there is a run of blank cells and then a QUARTERLY summary
  block ("Q4 2023", "Q1-2024", ...). Reading through it concatenates quarterly
  totals onto a weekly series.
- The week labelled with the as-of date is the week currently in progress.

So: keep only columns whose row-13 value parses as a date AND is <= the as-of
date. Stop at the first blank. Set `partial_last` to true, because the final
week is still running.

Do not hardcode column letters. The window grows by one column every week.

STEP 2 - LOCATE THE METRIC ROWS BY LABEL
Read `WEEKLY!A1:C450` once and find the row number for each of these labels in
column B or C. Match on exact text.

  "New Business Meetings Held"          -> meetings, a count
  "Total Pipeline Generation ($)"       -> total pipeline, dollars
  and the three segment rows immediately under Total Pipeline Generation:
  "Velocity", "Commercial", "Enterprise"   (the FIRST occurrence of each below
  the Total Pipeline Generation row; these labels repeat many times in the tab)

As of this writing those are rows 113, 135, 136, 140 and 144, but look them up
rather than trusting the numbers. Rows shift whenever anyone inserts a section.

Read each metric row across the window established in step 1, with
value_render_option UNFORMATTED_VALUE so you get numbers rather than "$1,242,800".

STEP 3 - RECONCILE BEFORE CHARTING
Velocity + Commercial + Enterprise must equal Total Pipeline Generation for every
week in the window. It does today, exactly, for all 52 weeks.

If any week is off by more than $1, you have picked up the wrong segment rows.
Stop, say which weeks disagree and what row numbers you used, and do not publish.
A dashboard built on the wrong rows is worse than no dashboard.

STEP 4 - BUILD
Write `series.json`:

    {"as_of": "<D5>",
     "source_url": "https://docs.google.com/spreadsheets/d/1FCFjss1No-3f2Su93ahKcIIbzzQREX6nv4AjgHktadk/edit",
     "weeks": [...],
     "partial_last": true,
     "charts": [
       {"title": "Total pipeline generation", "unit": "usd",
        "note": "New business pipeline created, all segments.",
        "series": [{"name": "Total", "values": [...]}]},
       {"title": "New business meetings held", "unit": "count",
        "note": "All sources: AE and exec outbound, SDR outbound, inbound.",
        "series": [{"name": "Meetings held", "values": [...]}]},
       {"title": "Pipeline generation by segment", "unit": "usd",
        "note": "Velocity, Commercial and Enterprise sum exactly to the total above.",
        "series": [{"name": "Velocity", ...}, {"name": "Commercial", ...},
                   {"name": "Enterprise", ...}]}]}

Use the trailing 52 weeks. Longer compresses the recent weeks into noise.

Write the script below to `build_dashboard.py` verbatim, then run:

    python3 build_dashboard.py series.json dashboard.html

Do not hand-write HTML or restyle the output. The layout, the palette and the
axis logic are fixed so that readers learn the dashboard once and only the
numbers change. The palette was validated for colour-vision deficiency in both
light and dark mode; substituting colours breaks that.

<<<BUILD_PY>>>

STEP 5 - PUBLISH TO DRIVE
Upload `dashboard.html` to Drive with `content_mime_type: text/html` and
`disable_conversion_to_google_type: true`. Without that flag Drive converts the
file to a Google Doc and the entire layout is destroyed.

Name it `GTM Weekly Trends [YYYY.MM.DD].html` using the as-of date. Take the
`viewUrl` from the response. Verify the returned `fileSize` matches the file on
disk; a mismatch means the upload truncated and should be redone.

STEP 6 - EMAIL
Send one message to travis@merge.dev.

    Subject: GTM weekly trends, week of {as-of}

    Four to six sentences, no preamble. What moved this week against last
    COMPLETE week, which segment drove it, and anything that looks like a data
    problem rather than a business event. Then the Drive link.

    Then a small HTML table: metric, last complete week, prior week, change.

Style: tight, no em dashes, no emojis, short declarative sentences, no sign-off.
Do not describe the charts. The reader can open them.

Note in the email that the final point on each chart is a week in progress, so
the visible drop at the right edge is not a real decline.

STEP 7 - COMPLETION GATE
Verify: the number of weeks charted equals the number of columns kept in step 1,
and the step 3 reconciliation passed for every week. Close with:
"{N} weeks charted through {as-of}. Sheet unchanged, dashboard in Drive, one
email sent."
