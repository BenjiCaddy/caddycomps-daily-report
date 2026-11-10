/**
 * CaddyComps daily report: export the comps sheet to the private GitHub repo every morning.
 *
 * Runs in YOUR Google account (benji@caddycomps.com). View access to the sheet is enough:
 * this script only READS the sheet. It writes one file, data/sheet.json, to the GitHub repo.
 *
 * Set-up (once):
 *  1. Go to https://script.google.com, New project, paste this whole file in, Save.
 *  2. Change REPO below to your repo, e.g. 'benji-caddycomps/caddycomps-daily-report'.
 *  3. Project Settings (gear icon) > Script properties > Add property:
 *       GITHUB_TOKEN = your fine-grained GitHub token (Contents: Read and write, this repo only).
 *     Type it there yourself. Never paste it into a chat.
 *  4. Back in the editor, pick exportSheetToGitHub and press Run. Approve the Google permissions.
 *     Check data/sheet.json appears in the repo.
 *  5. Pick createDailyTrigger and press Run once. It exports every morning between about 08:30 and 09:00 UK time.
 */
const SHEET_ID = '1Y3Eq5t9j38kXMDaRWFua7ZKrRuoVImB7GWdlCdU2fj8';
const REPO = 'OWNER/caddycomps-daily-report';   // <-- change OWNER to your GitHub user name
const PATH = 'data/sheet.json';
const BRANCH = 'main';
const COLS = 16;                                 // columns A to P are all the report needs

function exportSheetToGitHub() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const tabs = ss.getSheets()
    .filter(function (s) { return /week\s*\d+/i.test(s.getName()); })   // week tabs and Week NN Planning tabs only
    .map(function (s) {
      const last = s.getLastRow();
      const rows = last < 1 ? [] : s.getRange(1, 1, last, COLS).getValues().map(function (r) {
        return r.map(function (v) { return v instanceof Date ? v.toISOString() : v; });
      });
      return { title: s.getName(), rows: rows };
    });
  const body = JSON.stringify({ sheet: ss.getName(), exported_at: new Date().toISOString(), tabs: tabs });

  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) throw new Error('Add GITHUB_TOKEN under Project Settings > Script properties first.');
  const api = 'https://api.github.com/repos/' + REPO + '/contents/' + PATH;
  const headers = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };

  const current = UrlFetchApp.fetch(api + '?ref=' + BRANCH, { headers: headers, muteHttpExceptions: true });
  const sha = current.getResponseCode() === 200 ? JSON.parse(current.getContentText()).sha : undefined;
  const res = UrlFetchApp.fetch(api, {
    method: 'put', headers: headers, contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({
      message: 'Sheet export ' + new Date().toISOString(),
      content: Utilities.base64Encode(body, Utilities.Charset.UTF_8),
      sha: sha, branch: BRANCH
    })
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('GitHub upload failed: ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 300));
  }
  Logger.log('Exported ' + tabs.length + ' tabs, ' + Math.round(body.length / 1024) + ' KB');
}

/** Run once: export every morning between about 08:30 and 09:00 UK time. */
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'exportSheetToGitHub') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('exportSheetToGitHub').timeBased().everyDays(1).atHour(8).nearMinute(45)
    .inTimezone('Europe/London').create();
}
