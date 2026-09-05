# Weekly trend dashboard

Fridays 12:00 ET. Charts the trailing 52 weeks of GTM metrics from the WEEKLY tab
of the GTM Weekly Metrics sheet, uploads the dashboard to Drive, and posts the
link as a Slack DM. Source sheet is read-only.

## Why the old version hung

It did not crash. It built the dashboard correctly, then called the `Artifact`
tool to publish it. That raises an interactive approval prompt, and a scheduled
run has nobody present to approve it, so the session sat in
`SESSION_STATUS_REQUIRES_ACTION` indefinitely and the routine never reported.

The fix is delivery through Drive and Gmail, which run unattended, plus an
explicit guardrail against `Artifact` and anything else that prompts.

Four other defects in the old prompt, all corrected:

- It asked for bar graphs in one line and line charts in the next.
- It named Tableau, Power BI and Google Data Studio as the tool order. None of
  the three is reachable from a routine session.
- Its Inputs section ended with an empty `If missing: .` fallback.
- It asked for "pipeline by product category", which does not exist on the
  WEEKLY tab. See the open question below.

## Files

| File | Purpose |
|---|---|
| `PROMPT.md` | The routine, with `<<<BUILD_PY>>>` as a placeholder for the builder |
| `build_dashboard.py` | Chart generator, kept separate so it can be tested |
| `routine-prompt.txt` | `PROMPT.md` with the builder inlined. Paste this into the routine |
| `sample-light.png` / `sample-dark.png` | Rendered output, real data through 8/31/2026 |

Regenerate the deployable prompt after editing either source file:

```bash
python3 -c "p=open('PROMPT.md').read(); s=open('build_dashboard.py').read(); \
open('routine-prompt.txt','w').write(p.replace('<<<BUILD_PY>>>','\`\`\`python\n'+s.rstrip()+'\n\`\`\`'))"
```

## Sheet traps the routine has to handle

Each of these silently corrupts the chart rather than raising an error:

- **Row 13 runs past the last real week.** Future weeks are pre-populated as
  blanks and the series appears to collapse to zero. Cut at `WEEKLY!D5`, the
  as-of date.
- **A quarterly summary block follows the weekly run.** After a gap, the same row
  continues with `Q4 2023`, `Q1-2024` and so on. Reading through it appends
  quarterly totals to a weekly series.
- **The as-of week is still in progress.** Its partial numbers read as a crash.
  The builder takes stat tiles from the last complete week for this reason and
  labels the final chart point.
- **Row numbers move.** The metric rows are 113, 135, 136, 140 and 144 today, but
  the tab has row groups that people insert into. The routine matches on label
  text instead.
- **Segment labels repeat.** "Velocity", "Commercial" and "Enterprise" appear
  under nearly every section. Only the first occurrence below Total Pipeline
  Generation is the right one, and the routine verifies the choice by
  reconciling the three against the total before it charts anything.

## Chart decisions

Line charts, one y-scale each, never a dual axis. Palette is categorical slots
1-3 from the dataviz skill, validated all-pairs in both light and dark mode with
`validate_palette.js`; orange is excluded per Merge's convention. Light mode
flags aqua and yellow below 3:1 contrast, which obligates direct labels and a
data table, and both ship.

Verified on real data: Velocity + Commercial + Enterprise equals Total Pipeline
Generation exactly for all 52 weeks, zero mismatches.

## Open question

The original prompt asked for "pipeline by product category". That does not exist
on the WEEKLY tab, and the three candidate readings disagree:

- **By segment** (Velocity / Commercial / Enterprise) reconciles exactly to the
  total. This is what ships today.
- **Merge Unified vs CORE PG** (rows 159 and 168). CORE PG exceeds Total Pipeline
  Generation in the latest week, $349,480 against $331,380, so the two are not a
  clean split and a chart of them would not add up.
- **Product line breakdown** (Accelerate / Cruise / Neutral / Stop / TBD) exists,
  but on the UNIFIED tab under "$ PG - by Energy", not on WEEKLY, and it covers
  Merge Unified only.

Pick one and the third chart can be swapped or added.

## Delivery

Slack DM to travis@merge.dev. Slack renders no HTML and a message cannot display
an image on its own, so the HTML dashboard goes to Drive and Slack gets a text
version: current values, week-over-week change, and a block-character sparkline
per series inside a code fence, plus the Drive link.

Sparklines cover the trailing 26 weeks rather than the full 52. Over 52 weeks the
single $6.01M spike in Sep 2025 pins every later bar to the floor and the shape
carries no information.

Run `build_dashboard.py series.json slack.txt --slack "<drive url>"` to generate
it. The HTML path is unchanged.
