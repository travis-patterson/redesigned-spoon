# ROUTINE: renewal-radar
# Schedule: 0 12 * * 1  (Mondays 08:00 ET during EDT, 07:00 ET during EST)
# Runtime: Claude + Agent Handler (Salesforce, Gmail). Read-only against Salesforce.

GOAL
Rank every customer account with a contract event in the next 90 days by ARR at
risk, classify each as expansion, flat renewal, or retention risk, and name the
three to five accounts that need a CRO touch this week. Deliver as one self-email.

HARD GUARDRAILS  (violate any one of these -> stop and report)
- Salesforce is READ ONLY. Permitted: run_soql_query and the salesforce list/get/
  search tools. Forbidden: any create, update, upsert or delete against any
  Salesforce object, including Tasks, Notes and Chatter.
- Gmail: the only permitted write is a single message to travis@merge.dev.
  Forbidden: sending to any other recipient, any delete, trash, spam or label
  operation, and modifying drafts not created in this run.
- This report contains ARR, discount and at-risk figures per account. Do not post
  it to Slack, do not write it to a shared Drive folder, and do not include it in
  anything customer-facing. Self-email only.
- Do not open a full contract brief for any account in this run. This routine
  triages and names candidates; the /contract skill does the deep work on demand.

TOOLING
- Every Salesforce and Gmail call goes through Agent Handler. Load exact schemas
  first: tool_search("agent handler salesforce soql query"). Use the field names
  the schema returns; do not guess parameter names.
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
The script prints the ranked table and writes `radar.json`. Build the email from
its output. Do not recompute any number.

    Subject: Renewal radar, {N} accounts, ${total} in 90 days

    ## This week
    Three to five sentences. Lead with the largest at-risk line and what makes it
    at risk. Name the accounts renewing inside 30 days that are not yet in a good
    posture. Do not restate the table.

    ## Run /contract on these
    The top 3 to 5 rows by AT RISK, each one line: account, ARR, renewal date,
    posture, and the single reason it made the list.

    ## Full book
    The script's table, verbatim, in a <pre> block.

    ## Data defects
    The script's SYSTEMIC line and PER-ACCOUNT DEFECTS list, plus any DATE
    DISAGREEMENTS from step 1. These have owners: a blind account is a Sales Ops
    instrumentation gap, not a rep problem, and saying so keeps the list credible.

Style: tight, no preamble, no em dashes, no emojis, short declarative sentences,
no sign-off. A blind account outranks a merely under-consuming one of the same
size, because there is no basis for a renewal conversation at all.

STEP 5 - DELIVER
Agent Handler:gmail__create_draft, then gmail__send_message, to travis@merge.dev
only, HTML body. If send fails, leave the draft and say so.

STEP 6 - COMPLETION GATE
Before reporting, verify: rows in the script's table == accounts in `accounts.json`.
If they do not reconcile, the semi-join and the account query disagreed. Say so
rather than reporting a partial book. Close with:
"{N} accounts scored, ${total} in the 90-day book, ${risk} flagged. Salesforce
unchanged, one self-email sent."
