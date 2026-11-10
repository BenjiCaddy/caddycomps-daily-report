# CaddyComps daily sales report (cloud)

Every morning:
1. About 08:30 to 09:00 UK: a Google Apps Script in Benji's Google account reads the comps sheet (read-only) and saves `data/sheet.json` to this repo.
2. 09:45 UK: a Claude cloud routine checks out this repo, builds the 8-slide deck for yesterday, checks the totals, publishes the deck to the usual link, posts to #performance-team and emails Benji. Failures go to Benji as a Slack DM, never to the channel.

Folders: `skill/` (report scripts and rules), `tools/` (export converter and data check), `apps-script/` (the Google script), `ROUTINE.md` (the cloud routine's instructions).

Going back to the Mac version: pause the routine at https://claude.ai/code/routines and turn the "Daily sales report (09:45)" scheduled task back on in the Claude app. Nothing here needs deleting.

Keep this repo private: `data/sheet.json` is the comp budget data.
