// Budget app backend. Container-bound: open the Google Sheet > Extensions > Apps Script.
// The sheet does all the maths; this file only reads results and appends entries.

// Optional: public https link to icon.png (see setup notes). Leave '' to skip.
var FAVICON_URL = 'https://raw.githubusercontent.com/Nesh0573/budget/main/icon.png';

function doGet() {
  var out = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('My Budget')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  if (FAVICON_URL) out.setFaviconUrl(FAVICON_URL);
  return out;
}

function getState() {
  var ss = SpreadsheetApp.getActive();
  var b = ss.getSheetByName('Budget'), s = ss.getSheetByName('Savings'), v = ss.getSheetByName('Invest');
  var tz = ss.getSpreadsheetTimeZone();
  return {
    amount: b.getRange('B3').getValue(),
    month: Utilities.formatDate(new Date(b.getRange('B4').getValue()), tz, 'MMMM yyyy'),
    cats: b.getRange('A9:G20').getValues().map(function (r) {
      return { name: r[0], group: r[1], share: r[2], alloc: r[3], used: r[4], left: r[5], pct: r[6] };
    }),
    savings: s.getRange('H4:K7').getValues(),
    held: s.getRange('H10:I17').getValues(),
    market: { cbr: v.getRange('B4').getValue(), infl: v.getRange('B5').getValue() },
    invest: {
      total: v.getRange('B9').getValue(), extra: v.getRange('B8').getValue(),
      profile: v.getRange('B10').getValue(), status: v.getRange('B11').getValue(),
      blended: v.getRange('D22').getValue(), est: v.getRange('D23').getValue(),
      rows: v.getRange('A14:J20').getValues()
    }
  };
}

// items: [{id, type: 'expense'|'saving'|'budget'|'profile'|'extra', ...}]
// Entry IDs are generated on the phone, so a retry after a bad connection never double-posts.
function addEntries(items) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActive(), tz = ss.getSpreadsheetTimeZone();
    var ex = ss.getSheetByName('Expenses'), sv = ss.getSheetByName('Savings'), v = ss.getSheetByName('Invest');
    var seen = {};
    [[ex, 5], [sv, 6]].forEach(function (p) {
      var n = p[0].getLastRow();
      if (n > 1) p[0].getRange(2, p[1], n - 1, 1).getValues().forEach(function (r) { seen[r[0]] = 1; });
    });
    items.forEach(function (it) {
      if (it.type === 'budget') return ss.getSheetByName('Budget').getRange('B3').setValue(Number(it.amount));
      if (it.type === 'profile') return v.getRange('B10').setValue(it.value);
      if (it.type === 'extra') return v.getRange('B8').setValue(Number(it.amount) || 0);
      if (seen[it.id]) return;
      var d = Utilities.parseDate(it.date, tz, 'yyyy-MM-dd'), amt = Number(it.amount);
      if (!(amt > 0)) return;
      if (it.type === 'expense') ex.appendRow([d, it.category, amt, it.note || '', it.id]);
      if (it.type === 'saving') sv.appendRow([d, it.category, amt, it.where || '', it.note || '', it.id]);
      seen[it.id] = 1;
    });
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return getState();
}
