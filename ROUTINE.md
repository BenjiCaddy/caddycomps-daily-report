Run CaddyComps' daily sales report in the cloud. It runs at 09:45 UK time and reports on YESTERDAY (the last full day of sales). Benji has approved the wording and destinations, so send automatically once every check passes. Keep it light on credits: the scripts do the maths. Only output the short message, links and any problems. No screenshots and no re-checking of slides.

You are in a checkout of the private repo caddycomps-daily-report. It holds:
- skill/ : the report skill (read skill/SKILL.md once and follow its rules: flags, W.C week labels, 8 slides, no en or em dashes, planning window of report day + 8 to + 35 because comps must be live 7 days).
- data/sheet.json : this morning's export of the comps Google Sheet, written by a Google Apps Script at about 08:30 to 09:00.
- tools/json_to_xlsx.py and tools/check_total.py.

RULES
- Never read data/sheet.json into the conversation (it is about 800 KB). Only scripts read it.
- Google Drive is read-only: never edit, share or trash anything. Don't read any Gmail inbox.
- No en or em dashes in anything you write.
- If any step fails, do NOT post in the channel. Send Benji a Slack direct message instead (slack_send_message with channel_id U08JN5U1Q6T) saying in one or two lines which step failed and why, then stop. Never post partial or unchecked numbers.

STEPS
1. Freshness: python3 -c "import json;print(json.load(open('data/sheet.json'))['exported_at'])". If the file is missing, or exported_at is not today's date (UTC), the morning export did not run: that is a failure ("the Google export did not run this morning").
2. Set up: pip install -q openpyxl. Then
   python3 -I tools/json_to_xlsx.py data/sheet.json /tmp/master.xlsx
   YDAY=$(TZ=Europe/London date -d yesterday +%F)
3. Report (takes 2 to 4 minutes, use a 600000 ms Bash timeout):
   python3 -I skill/scripts/daily_report.py --workbook /tmp/master.xlsx --date "$YDAY" --out /tmp/out
   If it exits with code 3 (NO_SALES_FOR_DATE), yesterday's sales are not in the sheet yet: DM Benji "Daily sales report: yesterday's sales are not in the sheet yet" and stop.
4. Check: python3 -I tools/check_total.py skill/scripts /tmp/master.xlsx /tmp/out
   If "ok" is false, that is a failure: include "error" in the DM. If "ok" is true, keep "notes", "ignored" and "zero_sold" for the email.
5. Duplicate guard: read the last 10 messages of Slack channel C0C023ERG3D (slack_read_channel). If one already starts with the first line of /tmp/out/subject.txt, stop quietly.
6. Publish the deck to the SAME artifact every day so the link never changes:
   a. Artifact action "read" on https://claude.ai/artifact/YMERD6g4U8KQsGjrCPqHKp with path "project/deck.json".
   b. Artifact action "publish": url = https://claude.ai/artifact/YMERD6g4U8KQsGjrCPqHKp, root = /tmp/out, file_path = /tmp/out/project/deck.json, files = the 8 slides project/slides/{cover,today,trends,alerts,pricing,fatigue,plan,ideas}.html each mapped to the same relative path. If refused because of a newer version, follow the refusal and publish again. Refused 3 times = failure.
7. Send:
   a. Post the contents of /tmp/out/slack.md exactly as written to Slack channel C0C023ERG3D with slack_send_message.
   b. Email benji@caddycomps.com: subject = /tmp/out/subject.txt, plain text body = /tmp/out/email.txt, plus a final "DATA NOTES" section if check_total returned notes, ignored rows or zero_sold comps. Use the Gmail send_message tool (find it with ToolSearch "select:mcp__Gmail__send_message"). If there is no send tool, create a Gmail draft to benji@caddycomps.com with the same subject and body instead.
8. Finish with 3 lines: the overall %, the Slack message link and the deck link.
