# Renewal radar

Weekly triage of every customer account with a contract event in the next 90 days,
ranked by ARR at risk. Names the three to five accounts worth a full `/contract`
brief that week. Read-only against Salesforce; delivers one self-email.

## Files

| File | Purpose |
|---|---|
| `PROMPT.md` | The routine, with `<<<SCORE_PY>>>` as a placeholder for the scorer |
| `score.py` | Deterministic scorer, kept separate so it can be edited and tested |
| `routine-prompt.txt` | `PROMPT.md` with the scorer inlined. This is what goes in the trigger |

Regenerate the deployable prompt after editing either source file:

```bash
python3 -c "p=open('PROMPT.md').read(); s=open('score.py').read(); \
open('routine-prompt.txt','w').write(p.replace('<<<SCORE_PY>>>','\`\`\`python\n'+s.rstrip()+'\n\`\`\`'))"
```

The routine runs with `sources: []`, so it gets no repo checkout. It writes the
scorer to disk from its own prompt. That is why the scorer is inlined rather than
referenced by path.

## Why the scoring is a script, not a judgment call

Sixty accounts is past the point where reading a table produces a stable answer.
The same book has to produce the same ranking week to week, so that a change in
the report means the business moved rather than the model read the rows
differently. Thresholds live at the top of `score.py`.

## Posture rules

Taken from the `/contract` skill's posture table, with three additions that came
out of running it against the live book:

- **Overage outranks trend.** Above 110% utilization the posture is Expansion
  regardless of direction. Without this, an account at 438% of commit with a
  flat month classified as Hold.
- **Direction comes from the trailing 3 months, not the full window.** A 7-month
  baseline frequently lands mid-onboarding, which made most of the book read as
  growth that had actually finished months earlier.
- **Commits under 5 units are not rateable.** Deriving $/unit from a commit of 1
  produced a $20,000 upside figure from two extra units.

At-risk dollars are the full ARR on a retention or blind account, upside forgone
on an expansion account, and zero on a flat one.

## Salesforce field notes

Verified against the org, and each one is load-bearing:

- `Total_ARR__c` is zero on every customer record. Use `Active_ARR_New__c`.
- `Contracted_Linked_Accounts__c` carries zero and negative values in the live
  book. The scorer flags rather than divides.
- `Snapshot_Contracted_Linked_Accounts__c` and `Snapshot_utilization__c` are
  known-wrong per the `/contract` skill and are not read.
- `Date_Month_of_Snapshot__c >= LAST_N_MONTHS:7`, never `=`. The equality form
  excludes the current month and drops the freshest snapshot from every account.
- `Renewal_Date__c` is `Current_Subscription_End_Date__c` plus one day by
  convention. Not a discrepancy.

## First run, 2026-09-05

60 accounts, $3,359,399 up for renewal inside 90 days, $1,691,867 flagged,
$804,061 of shelfware.

Four accounts are blind. Uber ($288k, renews in 24 days), Syndio ($70.1k) and
Atlas ($20.1k) have never had a utilization snapshot written. Center ($55.8k)
has 20 snapshots that stop at 2025-12-02. That is $434,040 renewing with no
usage basis for the conversation.

Largest scored risk is Miro: $255,400 renewing in 24 days, 509 of 900 units in
use, down 43% over three months.

The paused-unit count is frozen for 4 or more months on 33 of 56 instrumented
accounts, so recoverable-unit figures are not quotable to a customer until Sales
Ops fixes the field.

## Deployment

Trigger `trig_01DbXmdnRpdWNjxEFzBmtyJs`, `0 12 * * 1`, fresh session per fire,
push notification on completion.

**The trigger was created without MCP connectors attached.** Routines created
through the MCP API can only inherit connectors the calling session itself holds
as passable grants, and this org does not permit setting them on the API call.
Until Agent Handler and Gmail are attached to this routine in the claude.ai
Routines UI, every fire will fail at step 1 with no Salesforce tools available.
The other five routines on this account have their connectors populated because
they were created through that UI.
