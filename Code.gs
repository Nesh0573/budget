// Budget app backend (container-bound: Google Sheet > Extensions > Apps Script).
// The sheet stores plain data (Categories, Months, Plan, Expenses, Savings, Invest inputs).
// All totals are calculated here from that data, so the app never depends on formula results.

var APP_VER = 'v10'; // must match VER in Index.html
var FAVICON_URL = 'https://raw.githubusercontent.com/Nesh0573/budget/main/icon.png';

function doGet() {
  var out = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('My Budget')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  if (FAVICON_URL) out.setFaviconUrl(FAVICON_URL);
  return out;
}

function sh_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function key_(d) { return Utilities.formatDate(new Date(d), tz_(), 'yyyy-MM'); }
function mdate_(k) { return Utilities.parseDate(k + '-01', tz_(), 'yyyy-MM-dd'); }
function cur_() { return key_(new Date()); }
function rows_(name, cols) {
  var s = sh_(name), n = s.getLastRow();
  return n < 2 ? [] : s.getRange(2, 1, n - 1, cols).getValues();
}
function label_(k) { return Utilities.formatDate(mdate_(k), tz_(), 'MMMM yyyy'); }

function getState(key) { return JSON.stringify(buildState_(key)); }

function buildState_(key) {
  var cur = cur_(); key = key || cur; ensureIds_();
  var cats = rows_('Categories', 3).filter(function (r) { return r[0]; })
    .map(function (r) { return { name: r[0], group: r[1], share: Number(r[2]) || 0 }; });
  var budgeted = false, amount = 0;
  rows_('Months', 2).forEach(function (r) {
    if (r[0] && key_(r[0]) === key) { budgeted = true; amount = Number(r[1]) || 0; }
  });
  var plan = {}, curInv = 0, invName = cats.filter(function (c) { return c.group === 'Savings'; })[1];
  invName = invName ? invName.name : 'Investments';
  rows_('Plan', 3).forEach(function (r) {
    if (!r[0]) return;
    var k = key_(r[0]), a = Number(r[2]) || 0;
    if (k === key) plan[r[1]] = (plan[r[1]] || 0) + a;
    if (k === cur && r[1] === invName) curInv += a;
  });
  var usedE = {}, usedS = {}, saved = {}, held = {}, entries = [];
  rows_('Expenses', 5).forEach(function (r) {
    if (budgeted && r[0] && key_(r[0]) === key) {
      usedE[r[1]] = (usedE[r[1]] || 0) + (Number(r[2]) || 0);
      entries.push({ id: r[4], kind: 'expense', date: Utilities.formatDate(new Date(r[0]), tz_(), 'yyyy-MM-dd'), category: r[1], amount: Number(r[2]) || 0, note: r[3], where: '' });
    }
  });
  rows_('Savings', 6).forEach(function (r) {
    if (!r[0]) return;
    var a = Number(r[2]) || 0;
    saved[r[1]] = (saved[r[1]] || 0) + a;
    held[r[3]] = (held[r[3]] || 0) + a;
    if (budgeted && key_(r[0]) === key) {
      usedS[r[1]] = (usedS[r[1]] || 0) + a;
      entries.push({ id: r[5], kind: 'saving', date: Utilities.formatDate(new Date(r[0]), tz_(), 'yyyy-MM-dd'), category: r[1], amount: a, note: r[4], where: r[3] });
    }
  });
  var allocTotal = 0;
  var out = cats.map(function (c) {
    var a = plan[c.name] || 0, u = budgeted ? ((c.group === 'Savings' ? usedS : usedE)[c.name] || 0) : 0;
    allocTotal += a;
    return { name: c.name, group: c.group, dshare: c.share, share: amount ? a / amount : 0, alloc: a, used: u, left: a - u, pct: a ? u / a : 0 };
  });

  var sv = sh_('Savings'), targets = [sv.getRange('J4').getValue(), 0, sv.getRange('J6').getValue()];
  var buckets = cats.filter(function (c) { return c.group === 'Savings'; }).map(function (c, i) {
    var t = Number(targets[i]) || 0, tot = saved[c.name] || 0;
    return { name: c.name, total: tot, target: t, prog: t ? Math.min(1, tot / t) : 0 };
  });
  var total = 0; Object.keys(saved).forEach(function (k) { total += saved[k]; });

  var inv = sh_('Invest'), infl = Number(inv.getRange('B5').getValue()) || 0;
  var names = inv.getRange('A14:A20').getValues(), meta = inv.getRange('A14:G20').getValues();
  var wts = inv.getRange('B27:D33').getValues(), profile = inv.getRange('B10').getValue();
  var pi = ['Conservative', 'Balanced', 'Growth'].indexOf(profile); if (pi < 0) pi = 1;
  var extra = Number(inv.getRange('B8').getValue()) || 0, tInv = curInv + extra, blended = 0;
  var rows = meta.map(function (r, i) {
    var rate = Number(r[1]) || 0, tax = Number(r[2]) || 0, min = Number(r[6]) || 0, w = Number(wts[i][pi]) || 0;
    var at = rate * (1 - tax), amt = tInv * w; blended += w * at;
    return [r[0], rate, tax, at, at - infl, r[5], min, w, amt, (min && amt < min) ? 'Below minimum - use the MMF instead' : 'OK'];
  });
  var em = buckets[0] ? buckets[0].prog : 0;
  return {
    ver: APP_VER, key: key, cur: cur, month: label_(key), status: key < cur ? 'past' : key === cur ? 'current' : 'future',
    budgeted: budgeted, amount: amount, entries: entries.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; }), allocTotal: allocTotal, unalloc: amount - allocTotal, cats: out,
    savings: { buckets: buckets, total: total },
    held: names.map(function (n) { return [n[0], held[n[0]] || 0]; }).concat([['Other / unassigned', total - names.reduce(function (s, n) { return s + (held[n[0]] || 0); }, 0)]]),
    market: { cbr: Number(inv.getRange('B4').getValue()) || 0, infl: infl },
    invest: { total: tInv, extra: extra, profile: ['Conservative', 'Balanced', 'Growth'][pi],
      status: em >= 1 ? 'Funded - fine to invest the rest' : 'Emergency fund at ' + Math.round(em * 100) + '% - fill it before investing more',
      blended: blended, est: tInv * blended, rows: rows }
  };
}

// Items come from the phone (possibly queued offline). Entry IDs stop double-posting on retry.
function addEntries(items, viewKey) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var notes = [];
  try {
    var cur = cur_(), seen = {};
    [['Expenses', 5], ['Savings', 6]].forEach(function (p) {
      rows_(p[0], p[1]).forEach(function (r) { seen[r[p[1] - 1]] = 1; });
    });
    items.forEach(function (it) {
      try { apply_(it, cur, seen, notes); } catch (e) { notes.push('Could not save: ' + e.message); }
    });
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  var st = buildState_(viewKey); st.notes = notes;
  return JSON.stringify(st);
}

function hasBudget_(k) {
  return rows_('Months', 2).some(function (r) { return r[0] && key_(r[0]) === k; });
}

function apply_(it, cur, seen, notes) {
  var inv = sh_('Invest');
  if (it.type === 'profile') return inv.getRange('B10').setValue(it.value);
  if (it.type === 'extra') return inv.getRange('B8').setValue(Number(it.amount) || 0);
  if (it.type === 'budget' || it.type === 'alloc' || it.type === 'reset') {
    if (it.month < cur) return notes.push(label_(it.month) + ' is closed and cannot be changed.');
    var amt = it.type === 'reset' ? 0 : Number(it.amount);
    if (!(amt >= 0)) return notes.push('Enter a valid amount.');
    if (it.type === 'reset') {
      var m = rows_('Months', 2).filter(function (r) { return r[0] && key_(r[0]) === it.month; })[0];
      if (!m) return notes.push('Create the ' + label_(it.month) + ' budget first.');
      return setBudget_(it.month, Number(m[1]) || 0, true);
    }
    if (it.type === 'budget') return setBudget_(it.month, amt);
    if (!hasBudget_(it.month)) return notes.push('Create the ' + label_(it.month) + ' budget first.');
    return setPlan_(it.month, it.category, amt);
  }
  if (it.type === 'edit' || it.type === 'delete') return change_(it, cur, notes);
  if (seen[it.id]) return;
  var k = String(it.date).slice(0, 7), amount = Number(it.amount);
  if (!(amount > 0)) return;
  if (k < cur) return notes.push(label_(k) + ' is closed. Entry not saved.');
  if (!hasBudget_(k)) return notes.push('Create the ' + label_(k) + ' budget first. Entry not saved.');
  var d = Utilities.parseDate(it.date, tz_(), 'yyyy-MM-dd');
  if (it.type === 'expense') sh_('Expenses').appendRow([d, it.category, amount, it.note || '', it.id]);
  if (it.type === 'saving') sh_('Savings').appendRow([d, it.category, amount, it.where || '', it.note || '', it.id]);
  seen[it.id] = 1;
}

function setPlan_(month, category, amount) {
  var p = sh_('Plan'), n = p.getLastRow(), rows = n > 1 ? p.getRange(2, 1, n - 1, 3).getValues() : [];
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] && key_(rows[i][0]) === month && rows[i][1] === category) { p.getRange(i + 2, 3).setValue(amount); return; }
  }
  p.appendRow([mdate_(month), category, amount]);
  p.getRange(p.getLastRow(), 1).setNumberFormat('mmm yyyy');
}

// New month: split by default shares. Existing month: scale the current split to the new amount.
// force = true re-splits by the default shares (reset). 0 is a valid amount everywhere.
function setBudget_(month, amount, force) {
  var ms = sh_('Months'), n = ms.getLastRow(), rows = n > 1 ? ms.getRange(2, 1, n - 1, 2).getValues() : [], row = -1, old = 0;
  rows.forEach(function (r, i) { if (r[0] && key_(r[0]) === month) { row = i + 2; old = Number(r[1]) || 0; } });
  var plan = {};
  rows_('Plan', 3).forEach(function (r) { if (r[0] && key_(r[0]) === month) plan[r[1]] = Number(r[2]) || 0; });
  var cats = rows_('Categories', 3).filter(function (r) { return r[0]; });
  var vals = cats.map(function (c) {
    return (row > 0 && old > 0 && !force && plan[c[0]] !== undefined) ? Math.round(plan[c[0]] * amount / old) : Math.round(amount * (Number(c[2]) || 0));
  });
  var oldSum = 0; Object.keys(plan).forEach(function (k) { oldSum += plan[k]; });
  var target = (row > 0 && old > 0 && !force) ? Math.round(oldSum * amount / old) : amount;
  var diff = target - vals.reduce(function (s, x) { return s + x; }, 0), big = 0;
  vals.forEach(function (x, i) { if (x > vals[big]) big = i; });
  if (vals.length) vals[big] += diff;
  if (row > 0) ms.getRange(row, 2).setValue(amount);
  else { ms.appendRow([mdate_(month), amount]); ms.getRange(ms.getLastRow(), 1).setNumberFormat('mmm yyyy'); }
  cats.forEach(function (c, i) { setPlan_(month, c[0], vals[i]); });
}

// Give every logged row a permanent ID (rows typed by hand in the sheet have none).
function ensureIds_() {
  [['Expenses', 5], ['Savings', 6]].forEach(function (p) {
    var s = sh_(p[0]), n = s.getLastRow(); if (n < 2) return;
    s.getRange(2, 1, n - 1, p[1]).getValues().forEach(function (r, i) {
      if (r[0] && !r[p[1] - 1]) s.getRange(i + 2, p[1]).setValue('m' + Utilities.getUuid().slice(0, 8));
    });
  });
}

// Edit or delete one logged entry. Closed months are locked.
function change_(it, cur, notes) {
  var sv = it.kind === 'saving', s = sh_(sv ? 'Savings' : 'Expenses'), w = sv ? 6 : 5, n = s.getLastRow();
  var rows = n > 1 ? s.getRange(2, 1, n - 1, w).getValues() : [], at = -1;
  rows.forEach(function (r, i) { if (r[w - 1] === it.id) at = i; });
  if (at < 0) { if (it.type === 'edit') notes.push('That entry was not found. It may already be deleted.'); return; }
  var ok = key_(rows[at][0]);
  if (ok < cur) return notes.push(label_(ok) + ' is closed. Entry not changed.');
  if (it.type === 'delete') return s.deleteRow(at + 2);
  var nk = String(it.date).slice(0, 7), amt = Number(it.amount);
  if (!(amt > 0)) return notes.push('Enter a valid amount.');
  if (nk < cur) return notes.push(label_(nk) + ' is closed. Pick a date this month or later.');
  if (!hasBudget_(nk)) return notes.push('Create the ' + label_(nk) + ' budget first.');
  var d = Utilities.parseDate(it.date, tz_(), 'yyyy-MM-dd');
  var vals = sv ? [d, it.category, amt, it.where || '', it.note || ''] : [d, it.category, amt, it.note || ''];
  s.getRange(at + 2, 1, 1, vals.length).setValues([vals]);
}
