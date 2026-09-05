# ROUTINE: renewal-radar
# Schedule: 0 12 * * 1  (Mondays 08:00 ET during EDT, 07:00 ET during EST)
# Runtime: Claude + Agent Handler (Salesforce, Slack). Read-only against Salesforce.

GOAL
Rank every customer account with a contract event in the next 90 days by ARR at
risk, classify each as expansion, flat renewal, or retention risk, and name the
three to five accounts that need a CRO touch this week. Deliver as a Slack DM.

HARD GUARDRAILS  (violate any one of these -> stop and report)
- Salesforce is READ ONLY. Permitted: run_soql_query and the salesforce list/get/
  search tools. Forbidden: any create, update, upsert or delete against any
  Salesforce object, including Tasks, Notes and Chatter.
- Slack: the only permitted write is posting to the DIRECT MESSAGE with
  travis@merge.dev, resolved by email lookup at run time. Never post to a
  channel, public or private, however well named. Never @-mention anyone.
- This report carries per-account ARR, discount and at-risk figures. It is a DM
  to one person and nothing else. Do not write it to a shared Drive folder and do
  not include it in anything customer-facing.
- Do not open a full contract brief for any account in this run. This routine
  triages and names candidates; the /contract skill does the deep work on demand.

TOOLING
- Every Salesforce and Slack call goes through Agent Handler. Load exact schemas
  first: tool_search("agent handler salesforce soql query slack post message
  lookup user by email"). Use the field names the schema returns; do not guess
  parameter names.
- Fall back to a native connector only if an Agent Handler call fails twice. If
  you fall back, say so in the final report.

STEP 1 - PULL THE RENEWAL BOOK
  Agent Handler:salesforce__run_soql_query

    SELECT Id, Name, Current_Subscription_End_Date__c, Account_Current_Renewal_Date__c,
           Active_ARR_New__c, Contracted_Linked_Accounts__c, Owner.Name
    FROM Account
    WHERE Type = 'Customer'
      AND Current_Subscription_End_Date__c >= TODAY
      AND Current_Subscription_End_Date__c <= NEXT_N_DAYS:90
    ORDER BY Current_Subscription_End_Date__c ASC

Field notes, all verified against the org:
- Use `Active_ARR_New__c` for ARR. `Total_ARR__c` is zero on every customer record
  and is not a usable field.
- `Current_Subscription_End_Date__c` is canonical. `Renewal_Date__c` is that date
  plus one day by convention, so do not treat the two as disagreeing.
- `Account_Current_Renewal_Date__c` usually matches but sometimes does not. Where
  it differs from `Current_Subscription_End_Date__c` by more than 3 days, list the
  account under DATE DISAGREEMENTS with both values. Do not pick a winner.
- `Contracted_Linked_Accounts__c` is the commit, and it is dirty: zero and negative
  values both occur in the live book. The scorer flags these rather than dividing
  by them.

Write the result to `accounts.json` as a list of objects with exactly these keys:
`id`, `name`, `end` (YYYY-MM-DD), `arr`, `commit`, `owner`.

STEP 2 - PULL UTILIZATION
Two queries, both scoped by a semi-join so no account id list has to be pasted.

Snapshot history:

    SELECT Account__c, Date_Month_of_Snapshot__c, Number_of_Linked_Accounts__c,
           Number_of_Paused_Linked_Accounts__c, Number_of_Relink_Needed_Linked_Accounts__c
    FROM Monthly_Account_Utilization_Snapshot__c
    WHERE Account__c IN (SELECT Id FROM Account WHERE Type = 'Customer'
          AND Current_Subscription_End_Date__c >= TODAY
          AND Current_Subscription_End_Date__c <= NEXT_N_DAYS:90)
      AND Date_Month_of_Snapshot__c >= LAST_N_MONTHS:7
    ORDER BY Account__c, Date_Month_of_Snapshot__c ASC

Use `>= LAST_N_MONTHS:7`, not `= LAST_N_MONTHS:7`. The equality form excludes the
current month and silently drops the freshest snapshot from every account.

Coverage census, which separates "never instrumented" from "instrumentation stopped":

    SELECT Account__c, COUNT(Id) snaps, MAX(Date_Month_of_Snapshot__c) latest
    FROM Monthly_Account_Utilization_Snapshot__c
    WHERE Account__c IN (SELECT Id FROM Account WHERE Type = 'Customer'
          AND Current_Subscription_End_Date__c >= TODAY
          AND Current_Subscription_End_Date__c <= NEXT_N_DAYS:90)
    GROUP BY Account__c

The snapshot query returns 400+ rows and will exceed the inline tool-result limit.
When that happens the harness saves the full result to a file and gives you the
path. Use that path directly as `snapshots.json`; do not try to read it into
context or re-run the query in smaller pieces. Save the coverage result as
`coverage.json` in the same `{"records": [...]}` shape.

`Number_of_Linked_Accounts__c` is the billable in-use count. Paused and
relink-needed are separate degraded states, not subsets of it. Never use
`Snapshot_Contracted_Linked_Accounts__c` or `Snapshot_utilization__c`: both are
known-wrong in this org.

STEP 3 - SCORE
Write the script below to `score.py` verbatim and run it:

    python3 score.py accounts.json snapshots.json coverage.json

Do not rank, classify or compute at-risk figures yourself. The whole point of the
script is that the same book produces the same ranking every week, so a change in
the report means the business changed, not that the model read the table
differently. If the script errors, fix the input files, not the thresholds.

<<<SCORE_PY>>>

STEP 4 - COMPOSE
The script prints the ranked table, writes `radar.json`, and writes the Slack
payloads: `slack-lead.txt` plus one or more `slack-table-NN.txt`. Do not
recompute any number and do not rebuild the table by hand.

Write a short read of the week, three to five sentences, to go at the top of the
lead message. Lead with the largest at-risk line and what makes it at risk. Name
anything renewing inside 30 days that is not in a good posture. Do not restate
the table; it is directly below.

Style: tight, no preamble, no em dashes, no emojis, short declarative sentences,
no sign-off. A blind account outranks a merely under-consuming one of the same
size, because there is no basis for a renewal conversation at all.

STEP 5 - DELIVER TO SLACK
Resolve the DM target first:

    Agent Handler:slack__lookup_user_by_email   email: travis@merge.dev

Take the user id from the response and pass it as `channel` on post_message.
Posting to a user id opens the direct message with that person. Post with
mrkdwn true and unfurl_links false.

If lookup_user_by_email fails, STOP and report. Do not substitute a channel,
and do not guess an id. If Slack returns reauth_required, say so plainly: the
Agent Handler Slack connection needs to be reauthorized and no amount of
retrying will fix it.

Post in this order:

1. One message: your step 4 read, then a blank line, then the contents of
   `slack-lead.txt` verbatim. Keep the `ts` from the response.
2. Each `slack-table-NN.txt` in ascending order as a threaded reply, passing the
   `ts` from step 1 as `thread_ts`. These are already wrapped in code fences,
   which is what holds the column alignment. Post them verbatim.

The table goes in the thread rather than the channel on purpose: 60 accounts is
several screens, and the lead message has to stay readable on a phone.

Do not send email. Slack is the only delivery channel for this routine.

STEP 6 - COMPLETION GATE
Before reporting, verify: rows in the script's table == accounts in `accounts.json`.
If they do not reconcile, the semi-join and the account query disagreed. Say so
rather than reporting a partial book. Close with:
"{N} accounts scored, ${total} in the 90-day book, ${risk} flagged. Salesforce
unchanged, posted to Slack DM as 1 message plus {K} threaded replies."
