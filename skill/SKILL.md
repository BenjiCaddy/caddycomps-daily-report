---
name: "daily-sell-through-report"
description: "Build CaddyComps' end-of-day sales report from the weekly comps Google Sheet called JUNE/JULY/AUGUST 2026 MASTER COMP BUDGET. Covers today vs the same day last week, overall trends, the flag and alert system, price point and ticket count analysis, fatigue suggestions with replacement ideas, and suggested changes to the upcoming plan (comps must be live 7 days, plans run 1 to 4 weeks ahead). Produces an 8-slide Slides artifact plus a Slack or email summary. Use whenever Benji asks for the daily report, today's sales review, an end-of-day update, or to run the daily, even if he does not name the skill. For the Monday weekly deck use weekly-sell-through-review instead."
---

# Daily sell-through report

This is a WEEK-TO-DATE report ending at the last day with sales. Same rules as the weekly skill: made 70% = sell-through of 70% or more; flags are Continuous (beige, under 70% in 3+ of the last 4 weeks), Serious (orange, under 64% in 2+), Alert (crimson, under 59% this week so far), Red Alert (dark red, under 59% for 2+ weeks running); only one flag per item, the most severe; week labels are W.C dates; never overlap text; no en or em dashes; feeders excluded from made-70 and group stats; day blocks end at the TOTAL row and stray rows are ignored (the script handles this).

Planning: the team plans 1 week to 1 month ahead and every comp must be on the website for at least 7 days before its draw. So anything drawing up to report day + 7 is already live and cannot be changed; suggestions apply to draws from report day + 8 up to report day + 35. The script reads the rest of the current week tab and the "Week NN Planning" tabs (one column to the right of week tabs) to know what is already planned.

Sheet: see `config.json`. The Google Sheet has one tab per week (names like "WEEK 23  (021126 - 08112026)"). Benji's team updates it during the day.

## Getting the data in (important)

The whole workbook is about 840 KB. Do NOT pull it through `read_file_content` or `download_file_content` into the conversation: it is far too large and wastes credits. Use one of these instead:
1. Benji uploads the workbook (File > Download > .xlsx) or a smaller export (the current week's tab plus the last 4 weeks).
2. If the environment saves downloaded Drive files to disk (for example Claude Cowork or Claude Code), download the sheet as .xlsx (`exportMimeType` application/vnd.openxmlformats-officedocument.spreadsheetml.sheet) straight to a file and use it.
If neither works, say so and ask for the upload.

## Run

1. `python scripts/daily_report.py --workbook <all_weeks.xlsx> [--extra <newer.xlsx>] [--day Sunday] [--date YYYY-MM-DD] --out <folder>`
   The scheduled run is at 09:45 and passes `--date` = yesterday: the script reports that day and ignores any later sales already in the sheet. If the sheet has no sales for that date it exits with code 3 (NO_SALES_FOR_DATE).
   The script reads every week tab, takes the latest week and its last day with sales as "today", and writes `<folder>/project/slides/*.html`, `<folder>/project/deck.json`, `<folder>/summary.txt` (Slack/email text) and `<folder>/daily.json`. Read `summary.txt` and check the Data check: today's overall % must equal the sheet's TOTAL row for that day (column K). If it does not, check the TOTAL row's SUM ranges (column L must cover the same rows as column H) and tell Benji; do not post.
2. Create the artifact with the Artifact tool: `publish` with `type_url` = the slides type URL in `config.json` and a title like "Daily Sales Review Sunday 8 Nov". It returns a URL and a folder; use that folder as `--out` (re-run the script into it), then publish `project/deck.json` plus the eight slides (`cover, today, trends, alerts, pricing, fatigue, plan, ideas`) with `url`. Do not verify or screenshot unless asked.
3. Deliver: the script also writes short sectioned messages, `slack.md` (Slack markdown), `email.txt` (plain text) and `subject.txt`. Send those, not `summary.txt`: the deck holds the detail (Benji's preference). They show today (overall, last same weekday, average over N weeks, comps made 70%), week to date, and the deck link; no £ gap. Set env `DECK_URL` if the deck link changes. Previously: post `summary.txt` plus the deck link to the Slack channel in `config.json` (ask Benji to confirm the channel and permission the first time; never post without confirmation). Email: only if a Gmail connector is connected, otherwise give Benji the text to send.

Comp names: the same product is sometimes worded differently in the sheet (for example R7 QUAD 15.5° MINI DRIVER and R7 QUAD MINI DRIVER). The script merges loft, flex, hand size, brackets, stray #, trailing PUTTER, GOLF BALLS and GPS WATCH automatically. Colour, model and edition words (X, MAX, PRO, BLACK, CHROME, SHADOWFALL) are kept apart because they can be different products; when Benji confirms two names are the same product, add them to `references/aliases.json`.


Sheet layouts: from Week 25 the week tabs use the new layout (same as the Planning tabs: A Planned, B slot or day header, D prize, E price, F tickets, I total, K sold, L %, M revenue, blocks end at "Daily Comp Total"). The scripts read both layouts.

Triple Whale (optional): pass `--tw tw.json` with {sales, orders, ad_spend, roas, prev_*} for the report day and the same day last week (fetched with the Triple Whale connector's run-sql on orders_table and ads_table). It adds one line to the email and the Today slide; Slack is unchanged.

## The eight slides

1. Title: "Daily Sales Review" with the day and date
2. Today vs the same day last week: overall %, comps made 70%, revenue, £ gap, plus product and brand tables (today, last week, change)
3. Trends: week to date vs last week to the same day, last 14 days chart, last 4 weeks to date
4. Alerts: flags for products and brands that ran today, with the key
5. Which prices and ticket counts are working: two cards, Working and Struggling. Each price band and ticket count band is judged on how often its comps made 70% over the last 8 weeks against the rate for all comps (5+ points either way; bands with under 10 runs left out), with the best and weakest price and ticket mix added and this week as a check. Plain words, no tables of percentages
6. Possible fatigue and what could replace it: comps that ran in the last 3 days whose last 3 runs are 8+ points below their earlier runs and under 70% (or 5+ runs in 14 days under 72%); replacements are comps of the same product type averaging 70%+ over 3+ runs in 8 weeks, or 78%+ in a different product type at the same price band, rested 3+ days and not one-off prizes (limited edition, custom, seasonal). They are suggestions from past results, never forecasts. Brands that average under 70% across history (for example Takomo) and brands in `references/ideas.json` `avoid_brands` are never suggested as replacements
   Each fatigue row also says where the comp sits in the plan: which planned draws are already live, the first draw it can be swapped from, or that it is not in the plan for the next 4 weeks. Rest for replacements is counted to that swap date, and a replacement is never suggested if it is already planned within 3 days of it.
7. Upcoming plan, suggested changes: every planned comp drawing between report day + 8 and report day + 35 is checked. Reasons, most serious first: showing fatigue; averaging under 64% over 3+ runs in 8 weeks; brand averages under 70% across our history; booked within 2 days of its previous run. Top 6 shown, one row per comp at its first draw in the window, each with a swap from the replacement rules. Suggestions from past results, never forecasts
8. New comp ideas: three columns, Idea, What the market says (short, sourced, from `references/ideas.json` `market`, sources in speaker notes) and Our history (has it run and how did it do, plus the closest product type). No 'Benji idea' or 'Claude idea' labels and no suggested price or ticket structure (it ignored product cost). Location-bound prizes (lessons, golf breaks, venue vouchers) are left out because we sell across the UK and Ireland. PuttOUT and Srixon are brands Benji rates

## Market research (do this before the ideas slide when you can)

The script only knows our own history. Real market research needs web search:
1. If a web search tool is available, search for what is selling and trending in golf equipment right now (brands, products, at-home practice and simulator products) and what competing prize sites are running. Add only claims you can source, with the source and date, to `references/ideas.json` (or to the slide notes), and put brands Benji says are popular (currently Srixon) at the top.
2. If no web search is available, say plainly that no live research was run and that the ideas come from Benji's list and our history. Never invent market trends or sales figures.
3. Benji's inputs override mine: if he says a brand is popular or has not performed (for example Takomo), update `ideas.json` (`ideas`, `avoid_brands`).

## Triple Whale (optional, currently off)

The report does not need Triple Whale. Moby cannot schedule emails on Benji's plan, so there is no daily Triple Whale email to read; do not search the inbox for one. If a Triple Whale connector has been connected (Benji has asked his account admin), use it to add a traffic and spend line (sessions, orders, revenue, ad spend against the same weekday last week) to the summary and the Today slide. Never put API keys in chat and never read other emails in the inbox.

Delivery by email: use the Gmail connector (benji@caddycomps.com). Confirm recipients and wording with Benji before the first send.

## Notes

- Partial weeks: flags on 1 to 3 days of data are less reliable; the footer says "week to date".
- Credits: the script does all the maths; keep chat output to the summary and the links.
- If a comp shows 0 sold, tell Benji and check it is really in that day's block before presenting.
