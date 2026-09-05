# Routines

State of the scheduled routines on this account, as of 2026-09-05.

| Routine | Schedule | State | Notes |
|---|---|---|---|
| auto draft email replies | Daily 12:00 ET | on | Working. Last run succeeded 9/4. |
| Weekly trend dashboard | Fri 12:00 ET | on, needs prompt swap | Rebuilt here in `weekly-trends/`, now delivering to Slack. Paste `weekly-trends/routine-prompt.txt` into it. |
| Monthly rep scorecard | Fri 12:00 ET | on | Cron is weekly (`0 16 * * 5`) despite the name. Use `0 16 1 * *` for monthly. |
| Pipeline Hygiene | Fri 12:00 ET | off, dropped | Depended on a zip that no longer exists. Decision on 9/5: leave it off. Delete in the UI if you want it gone. |
| Travis Morning Brief | Weekdays 07:00 ET | off | Turned off while it was Gmail and Calendar only. Agent Handler now covers Salesforce and Gong, so the constraint in its prompt is stale. |
| Renewal radar | Mon 08:00 ET | on, needs connectors | Built here in `renewal-radar/`, delivering to Slack. Attach Agent Handler in the UI or it fails at step 1. |

## What an agent can and cannot do to a routine from a session

This constrained every fix in this directory, so it is worth stating plainly:

- **Creating** a routine works, but `connectors` is blocked at the org level, and
  the fallback only passes through grants the calling session holds as passable.
  A routine created this way reaches no connectors and cannot be repaired from a
  session.
- **Updating** a routine created through the claude.ai Routines UI is refused
  outright, including just flipping it to disabled. `update_trigger` only accepts
  routines an agent created itself.

So the deployable artifact from a session is the prompt text, not the routine.
Each subdirectory ships a `routine-prompt.txt` to paste into the UI.

## Slack delivery

Both rebuilt routines DM travis@merge.dev rather than sending email. The pattern
is the same in each and is worth reusing for the others:

1. `slack__lookup_user_by_email` with `travis@merge.dev`, at run time. Never a
   hardcoded id, and never a channel name.
2. The returned user id is the `channel` argument on `post_message`. Posting to a
   user id opens the DM with that person.
3. If the lookup fails, stop. Do not fall back to a channel. These reports carry
   per-account ARR and at-risk figures.
4. Anything wider than a phone screen goes in a code fence, and anything longer
   than ~3,500 characters goes in a thread rather than the message body.

**Agent Handler's Slack connection is currently unauthorized** on this account:
`slack__lookup_user_by_email` returns `reauth_required`. Reauthorize it or every
Slack-delivering routine stops at the post step.
