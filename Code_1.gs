/**
 * Spend tracker
 * =============
 * Gmail (labelled bank alerts) -> raw -> transactions -> Telegram.
 *
 * SHEET TABS
 *   raw           msg_id | received_at | source | from | subject | body
 *   transactions  msg_id | txn_date | txn_time | amount | currency | merchant | source | account | type
 *   exclude       (optional) col A: merchant or account values to leave out of totals
 *
 * TRIGGERS
 *   runAll          -> time-driven, hourly
 *   pollCommands    -> time-driven, every minute
 *   sendDailyBrief  -> time-driven, daily 10-11pm
 *
 * SCRIPT PROPERTIES (Project Settings -> Script Properties)
 *   TG_TOKEN            required
 *   TG_CHAT             required
 *   ANTHROPIC_API_KEY   optional, enables free-form questions
 *   TG_OFFSET           managed automatically, do not set by hand
 *
 * NOTE: uses polling, not webhooks. Apps Script /exec returns a 302 that
 * Telegram's webhook client won't follow. No deployment step is needed —
 * code changes take effect on save.
 */

// ---------------------------------------------------------------- CONFIG

const SOURCES = [
  { label: 'DBS PAYLAH',                                        source: 'PayLah' },
  { label: 'UOB Transaction Alerts',                            source: 'UOB'    },
  { label: 'UOB Personal Internet Banking Notification Alerts',  source: 'UOB'    },
];

const SEARCH_CAP = 500;   // raise before a large backfill
const BODY_CAP   = 5000;
const MAX_LINES  = 12;    // transactions listed before truncating

const CLAUDE_MODEL = 'claude-haiku-4-5-20251001';

const HEADERS = {
  raw: ['msg_id', 'received_at', 'source', 'from', 'subject', 'body'],
  transactions: ['msg_id', 'txn_date', 'txn_time', 'amount', 'currency',
                 'merchant', 'source', 'account', 'type'],
};

const MONTHS = { Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6,
                 Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12 };

const DOW = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];

// Everything counts as spend except rows whose amount is unknown.
const NON_SPEND_TYPES = ['paynow_noamt'];

// Sub-breakdown only — both are included in the headline total.
const EVERYDAY_TYPES = ['card', 'wallet'];

// ---------------------------------------------------------------- ENTRY POINTS

function runAll() {
  assertHeaders();
  pullAlerts();
  parseRaw();
}

function assertHeaders() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEADERS).forEach(function (name) {
    const cols = HEADERS[name];
    const sh = ss.getSheetByName(name);
    if (!sh) throw new Error('Missing tab: "' + name + '"');
    const got = sh.getRange(1, 1, 1, cols.length).getValues()[0];
    cols.forEach(function (want, i) {
      if (String(got[i]).trim() !== want) {
        throw new Error(name + ' column ' + (i + 1) +
                        ': expected "' + want + '", got "' + got[i] + '"');
      }
    });
  });
}

// ---------------------------------------------------------------- FETCH

function pullAlerts() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('raw');
  const lastRow = sheet.getLastRow();
  const seen = new Set(
    lastRow > 1
      ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String)
      : []
  );

  const rows = [];

  SOURCES.forEach(function (src) {
    const threads = GmailApp.search('label:"' + src.label + '"', 0, SEARCH_CAP);
    console.log(src.label + ': ' + threads.length + ' threads');

    if (threads.length === 0) {
      console.log('WARNING: "' + src.label + '" returned 0 threads — check the label name in Gmail');
    }
    if (threads.length >= SEARCH_CAP) {
      console.log('WARNING: hit SEARCH_CAP for "' + src.label + '" — raise it');
    }

    threads.forEach(function (thread) {
      thread.getMessages().forEach(function (msg) {
        const id = msg.getId();
        if (seen.has(id)) return;
        rows.push([
          id,
          msg.getDate(),
          src.source,
          msg.getFrom(),
          msg.getSubject(),
          msg.getPlainBody().slice(0, BODY_CAP),
        ]);
        seen.add(id);
      });
    });
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
  console.log('Added ' + rows.length + ' rows to raw');
  return rows.length;
}

// ---------------------------------------------------------------- PARSERS

/**
 * DBS PayLah! wallet.
 * Anchored on the wallet line so the Outlook forward header's "To:" is skipped.
 */
function parsePayLah(body) {
  if (!body) return null;

  const amt = body.match(/Amount:\s*([A-Z]{3})\s*([\d,]+\.\d{2})/);
  if (!amt) return null;

  const dt = body.match(/Date & Time:\s*(\d{1,2})\s+(\w{3})\s+(\d{2}:\d{2})/);
  const block = body.match(
    /From:\s*PayLah! Wallet \(Mobile ending (\d+)\)\s*[\r\n]+\s*To:\s*(.+)/
  );

  return {
    date: dt ? isoFromDayMon(dt[1], dt[2]) : null,
    time: dt ? dt[3] : '',
    amount: parseFloat(amt[2].replace(/,/g, '')),
    currency: amt[1],
    merchant: block ? block[2].trim() : '',
    account: block ? block[1] : 'PayLah',
    type: 'wallet',
  };
}

/** UOB — tries each known format in turn. */
function parseUob(body) {
  if (!body) return null;
  return parseUobCard(body)
      || parseUobPayNow(body)
      || parseUobPayNowBare(body);
}

/** "A transaction of SGD 12.34 was made with your UOB Card ending 1234 on 01/01/26 at MERCHANT." */
function parseUobCard(body) {
  const m = body.match(
    /A transaction of ([A-Z]{3})\s*([\d,]+\.\d{2}) was made with your (.+?) ending (\d{4}) on (\d{2})\/(\d{2})\/(\d{2}) at (.+?)(?:\.\s*If unauthorised|\.\s*$)/s
  );
  if (!m) return null;

  return {
    date: '20' + m[7] + '-' + m[6] + '-' + m[5],
    time: '',
    amount: parseFloat(m[2].replace(/,/g, '')),
    currency: m[1],
    merchant: m[8].trim(),
    account: m[4],
    type: 'card',
  };
}

/** "You made a PayNow transfer of SGD 100.00 to NAME (UEN ending XXXX) on your a/c ending 0000 at 11:47AM SGT, 01 Jan 26." */
function parseUobPayNow(body) {
  const m = body.match(
    /You made a PayNow transfer of ([A-Z]{3})\s*([\d,]+\.\d{2}) to (.+?)(?:\s*\((?:UEN|Mobile|NRIC)[^)]*\))? on your a\/c ending (\d{4}) at (\d{1,2}):(\d{2})\s*(AM|PM) SGT,\s*(\d{1,2})\s+(\w{3})\s+(\d{2})/i
  );
  if (!m) return null;

  let hh = parseInt(m[5], 10);
  const ampm = m[7].toUpperCase();
  if (ampm === 'PM' && hh !== 12) hh += 12;
  if (ampm === 'AM' && hh === 12) hh = 0;

  const mm = MONTHS[titleCase(m[9])];
  if (!mm) return null;

  return {
    date: '20' + m[10] + '-' + pad(mm) + '-' + pad(m[8]),
    time: pad(hh) + ':' + m[6],
    amount: parseFloat(m[2].replace(/,/g, '')),
    currency: m[1],
    merchant: m[3].trim(),
    account: m[4],
    type: 'paynow',
  };
}

/**
 * "UOB - Your PayNow transfer to NAME on 01-Jan-2026 is successful."
 * No amount in the email. Captured with a blank amount so the gap stays
 * visible rather than being silently dropped.
 */
function parseUobPayNowBare(body) {
  const m = body.match(
    /Your PayNow transfer to (.+?) on (\d{1,2})-(\w{3})-(\d{4}) is successful/i
  );
  if (!m) return null;

  const mm = MONTHS[titleCase(m[3])];
  if (!mm) return null;

  return {
    date: m[4] + '-' + pad(mm) + '-' + pad(m[2]),
    time: '',
    amount: '',
    currency: 'SGD',
    merchant: m[1].trim(),
    account: '',
    type: 'paynow_noamt',
  };
}

const PARSERS = { PayLah: parsePayLah, UOB: parseUob };

// ---------------------------------------------------------------- PARSE

function parseRaw() {
  const ss  = SpreadsheetApp.getActive();
  const raw = ss.getSheetByName('raw');
  const txn = ss.getSheetByName('transactions');

  const lastRaw = raw.getLastRow();
  console.log('raw has ' + lastRaw + ' rows');
  if (lastRaw < 2) { console.log('raw is empty — run pullAlerts first'); return 0; }

  const done = new Set(
    txn.getLastRow() > 1
      ? txn.getRange(2, 1, txn.getLastRow() - 1, 1).getValues().flat().filter(String)
      : []
  );

  const data = raw.getRange(2, 1, lastRaw - 1, 6).getValues();
  const rows = [];
  const failed = [];

  data.forEach(function (r) {
    const msgId = r[0], source = r[2], body = r[5];
    if (!msgId || done.has(msgId)) return;

    const fn = PARSERS[source];
    if (!fn) { failed.push(msgId + ' (no parser for "' + source + '")'); return; }

    const p = fn(body);
    if (!p)      { failed.push(msgId + ' (' + source + ' regex miss)'); return; }
    if (!p.date) { failed.push(msgId + ' (no date parsed)'); return; }

    rows.push([msgId, p.date, "'" + p.time, p.amount, p.currency,
               p.merchant, source, p.account, p.type]);
  });

  if (rows.length) {
    txn.getRange(txn.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
  }

  console.log('Parsed ' + rows.length + ', failed ' + failed.length);
  if (failed.length) console.log(failed);
  flagDuplicates(rows);
  return rows.length;
}

/**
 * UOB alerts carry no transaction reference, so a re-forwarded email is
 * indistinguishable from a genuine repeat charge. Flag, never auto-merge:
 * two identical $4.50 charges in one day are legitimate.
 */
function flagDuplicates(rows) {
  const counts = {};
  rows.forEach(function (r) {
    const key = [r[1], r[3], r[5], r[7]].join('|');
    counts[key] = (counts[key] || 0) + 1;
  });
  const dupes = Object.keys(counts).filter(function (k) { return counts[k] > 1; });
  if (dupes.length) {
    console.log('POSSIBLE DUPLICATES (check manually): ' + dupes.join(' ;; '));
  }
}

// ---------------------------------------------------------------- REPORTING

/**
 * Optional `exclude` tab, column A: merchant names or account numbers to keep
 * out of totals — self-transfers, moving money to savings, paying your own card.
 * Matching is case-insensitive and exact on either merchant or account.
 */
function excludeSet() {
  const sh = SpreadsheetApp.getActive().getSheetByName('exclude');
  if (!sh || sh.getLastRow() < 2) return new Set();
  return new Set(
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
      .flat()
      .filter(String)
      .map(function (v) { return String(v).trim().toLowerCase(); })
  );
}

/** SGD rows with txn_date between fromDate and toDate inclusive (yyyy-MM-dd). */
function rowsBetween(fromDate, toDate) {
  const txn = SpreadsheetApp.getActive().getSheetByName('transactions');
  if (txn.getLastRow() < 2) return [];

  const skip = excludeSet();

  return txn.getRange(2, 1, txn.getLastRow() - 1, 9).getValues()
    .filter(function (r) {
      const d = asDateStr(r[1]);
      if (d < fromDate || d > toDate) return false;
      if (r[4] !== 'SGD') return false;
      if (skip.has(String(r[5]).trim().toLowerCase())) return false;
      if (skip.has(String(r[7]).trim().toLowerCase())) return false;
      return true;
    });
}

/** Rows that count toward a spend total (amount known). */
function spendRows(rows) {
  return rows.filter(function (r) { return NON_SPEND_TYPES.indexOf(r[8]) === -1; });
}

/** Generic summary used by every command and by the nightly brief. */
function summarise(fromDate, toDate, title) {
  const all = rowsBetween(fromDate, toDate);
  if (!all.length) return '*' + title + ': $0.00* — nothing recorded';

  const spend    = spendRows(all);
  const everyday = spend.filter(function (r) { return EVERYDAY_TYPES.indexOf(r[8]) !== -1; });
  const transfers= spend.filter(function (r) { return EVERYDAY_TYPES.indexOf(r[8]) === -1; });
  const unknown  = all.filter(function (r) { return NON_SPEND_TYPES.indexOf(r[8]) !== -1; });

  const total = sumAmounts(spend);

  let msg = '*' + title + ': $' + total.toFixed(2) + '*  (' + spend.length + ' txns)';

  if (everyday.length && transfers.length) {
    msg += '\n_$' + sumAmounts(everyday).toFixed(2) + ' everyday · $'
         + sumAmounts(transfers).toFixed(2) + ' transfers_';
  }

  const byAmount = function (a, b) { return Number(b[3]) - Number(a[3]); };
  const line = function (r) {
    return '  ' + Number(r[3]).toFixed(2) + '  ' + safe(r[5]);
  };

  const sorted = spend.slice().sort(byAmount);
  msg += '\n' + sorted.slice(0, MAX_LINES).map(line).join('\n');
  if (sorted.length > MAX_LINES) {
    msg += '\n  …and ' + (sorted.length - MAX_LINES) + ' more';
  }

  if (unknown.length) {
    msg += '\n\n⚠️ ' + unknown.length + ' transfer(s) with no amount in the alert — check the UOB app';
  }

  return msg;
}

/** Mon-anchored week containing today. */
function weekDays() {
  const tz  = Session.getScriptTimeZone();
  const now = new Date();
  const idx = DOW.indexOf(Utilities.formatDate(now, tz, 'EEE'));

  const out = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(now.getTime());
    d.setDate(d.getDate() - idx + i);
    out.push({
      label: DOW[i],
      date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
      isToday: i === idx,
      isFuture: i > idx,
    });
  }
  return out;
}

function weekSummary() {
  const days = weekDays();
  const rows = spendRows(rowsBetween(days[0].date, days[6].date));

  const totals = {};
  rows.forEach(function (r) {
    const d = asDateStr(r[1]);
    totals[d] = (totals[d] || 0) + Number(r[3] || 0);
  });

  let weekTotal = 0;
  const lines = days.map(function (d) {
    if (d.isFuture) return d.label + '   pending';
    const amt = totals[d.date] || 0;
    weekTotal += amt;
    return d.label + '   $' + amt.toFixed(2) + (d.isToday ? '  (so far)' : '');
  });

  return '*This week*\n```\n' + lines.join('\n')
       + '\n───────────────\nTotal  $' + weekTotal.toFixed(2) + '\n```';
}

/** Per-day breakdown for a calendar month, plus totals. */
function monthSummary(offset) {
  const r = monthRange(offset);
  const all = rowsBetween(r.from, r.to);
  const spend = spendRows(all);

  const byDay = {};
  spend.forEach(function (x) {
    const d = asDateStr(x[1]);
    byDay[d] = (byDay[d] || 0) + Number(x[3] || 0);
  });

  const days = Object.keys(byDay).sort();
  const total = sumAmounts(spend);

  // Average over elapsed days, so a mid-month check isn't misleading.
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const lastDay = r.to < todayStr ? r.to : todayStr;
  const elapsed = Number(lastDay.slice(8, 10));

  let msg = '*' + r.name + ': $' + total.toFixed(2) + '*  (' + spend.length + ' txns)';
  msg += '\n_avg $' + (elapsed ? total / elapsed : 0).toFixed(2) +
         '/day over ' + elapsed + ' days_';

  if (days.length) {
    msg += '\n```\n' + days.map(function (d) {
      return d.slice(5) + '   $' + byDay[d].toFixed(2);
    }).join('\n') + '\n```';
  }

  return msg;
}

// ---------------------------------------------------------------- TELEGRAM

function sendDailyBrief() {
  const today = shift(0);
  let msg = summarise(today, today, 'Today');
  const wk = weekSummary();
  if (wk) msg += '\n\n' + wk;
  tgSend(msg, true);
}

/**
 * Polls Telegram for new messages. Runs on a per-minute trigger.
 * TG_OFFSET is what stops old commands being reprocessed — Telegram only
 * clears an update once acknowledged with a higher offset.
 */
function pollCommands() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TG_TOKEN');
  if (!token) return;

  const offset = Number(props.getProperty('TG_OFFSET') || 0);

  const res = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + token + '/getUpdates?timeout=0&offset=' + offset,
    { muteHttpExceptions: true }
  );

  let data;
  try {
    data = JSON.parse(res.getContentText());
  } catch (err) {
    console.log('getUpdates parse error: ' + res.getContentText());
    return;
  }
  if (!data.ok) { console.log('getUpdates: ' + res.getContentText()); return; }
  if (!data.result.length) return;

  let maxId = offset;
  data.result.forEach(function (u) {
    maxId = Math.max(maxId, u.update_id + 1);
    const m = u.message || u.edited_message;
    if (!m || !m.text) return;
    if (String(m.chat.id) !== String(props.getProperty('TG_CHAT'))) return;

    const text = m.text.trim().split('@')[0];
    try {
      tgSend(handleCommand(text));
    } catch (err) {
      tgSend('Error: ' + err);
    }
  });

  props.setProperty('TG_OFFSET', String(maxId));
}

function handleCommand(raw) {
  const text = raw.toLowerCase();

  switch (text) {
    case '/today':
      return summarise(shift(0), shift(0), 'Today');
    case '/yesterday':
      return summarise(shift(-1), shift(-1), 'Yesterday');
    case '/week':
      return weekSummary();
    case '/lastweek': {
      const days = weekDays();
      return summarise(shiftFrom(days[0].date, -7),
                       shiftFrom(days[0].date, -1), 'Last week');
    }
    case '/month':
      return monthSummary(0);
    case '/lastmonth':
      return monthSummary(-1);
    case '/sync': {
      const added  = pullAlerts();
      const parsed = parseRaw();
      return 'Synced: ' + added + ' new email(s), ' + parsed + ' parsed.';
    }
    case '/start':
    case '/help':
      return 'Commands:\n/today\n/yesterday\n/week\n/lastweek\n/month\n/lastmonth\n/sync'
           + '\n\nOr just ask a question in plain English.';
    default:
      if (text.charAt(0) === '/') return 'Unknown command. Try /help';
      return askClaude(raw);
  }
}

function tgSend(msg, throwOnError) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TG_TOKEN');
  const chat  = props.getProperty('TG_CHAT');
  if (!token || !chat) throw new Error('Set TG_TOKEN and TG_CHAT in Script Properties');

  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    payload: { chat_id: chat, text: msg, parse_mode: 'Markdown' },
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  if (code !== 200) {
    const detail = 'Telegram ' + code + ': ' + res.getContentText();
    console.log(detail);
    // Retry once without Markdown — an odd merchant name can break the parser.
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      payload: { chat_id: chat, text: msg },
      muteHttpExceptions: true,
    });
    if (throwOnError) throw new Error(detail);
  }
}

// ---------------------------------------------------------------- CLAUDE

/** Precomputed facts plus raw rows for the current month. */
function buildContext() {
  const m = monthRange(0);
  const all = rowsBetween(m.from, m.to);
  const spend = spendRows(all);

  const byDay = {};
  const byMerchant = {};
  spend.forEach(function (r) {
    const d = asDateStr(r[1]);
    byDay[d] = (byDay[d] || 0) + Number(r[3] || 0);
    byMerchant[r[5]] = (byMerchant[r[5]] || 0) + Number(r[3] || 0);
  });

  const top = Object.keys(byMerchant)
    .sort(function (a, b) { return byMerchant[b] - byMerchant[a]; })
    .slice(0, 15)
    .map(function (k) { return k + ': $' + byMerchant[k].toFixed(2); });

  const total = sumAmounts(spend);
  const elapsed = Number(shift(0).slice(8, 10));

  return [
    'Today is ' + shift(0) + ' (Singapore). Currency is SGD.',
    'Transaction types: card and wallet are everyday spending; paynow is a bank',
    'transfer (also real spending). All are included in totals.',
    '',
    'MONTH TO DATE (' + m.name + '): $' + total.toFixed(2) +
      ' across ' + spend.length + ' transactions',
    'Average per elapsed day: $' + (elapsed ? total / elapsed : 0).toFixed(2) +
      ' over ' + elapsed + ' days',
    '',
    'DAILY TOTALS:',
    Object.keys(byDay).sort().map(function (d) {
      return d + '  $' + byDay[d].toFixed(2);
    }).join('\n'),
    '',
    'TOP MERCHANTS THIS MONTH:',
    top.join('\n'),
    '',
    'TRANSACTIONS (date, time, amount, merchant, account, type):',
    spend.map(function (r) {
      return [asDateStr(r[1]), r[2], r[3], r[5], r[7], r[8]].join('\t');
    }).join('\n'),
  ].join('\n');
}

function askClaude(question) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) {
    return 'Free-form questions need ANTHROPIC_API_KEY in Script Properties. Try /help.';
  }

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': key,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify({
      model: CLAUDE_MODEL,
      max_tokens: 500,
      system: 'You answer questions about the user\'s spending, in Singapore dollars. ' +
              'The totals given in the context are precomputed and authoritative — use ' +
              'them exactly, never recalculate them. For anything not covered by a ' +
              'precomputed total you may sum the transaction rows. If the data does not ' +
              'contain the answer, say so plainly rather than guessing. Be terse: two or ' +
              'three sentences unless asked for detail. Plain text, no markdown headers.',
      messages: [{
        role: 'user',
        content: buildContext() + '\n\nQuestion: ' + question,
      }],
    }),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    console.log('Claude error: ' + res.getContentText());
    return 'Could not reach Claude (' + res.getResponseCode() + ').';
  }

  return JSON.parse(res.getContentText()).content
    .filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; })
    .join('\n');
}

// ---------------------------------------------------------------- DIAGNOSTICS

/**
 * Sends a test message. Logs only whether the credentials are present,
 * never their values, so the execution log is safe to screenshot or share.
 */
function testTelegram() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TG_TOKEN');
  const chat  = props.getProperty('TG_CHAT');
  console.log('token present: ' + !!token + ', chat present: ' + !!chat);
  if (!token || !chat) return;

  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
    method: 'post',
    payload: { chat_id: chat, text: 'test' },
    muteHttpExceptions: true,
  });
  console.log('Telegram responded ' + res.getResponseCode());
}

// ---------------------------------------------------------------- HELPERS

/**
 * Sheets may store txn_date as a Date object or as text depending on how the
 * cell was written. Normalise on read so comparisons always work.
 */
function asDateStr(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).trim();
}

/** yyyy-MM-dd for today + days. */
function shift(days) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** yyyy-MM-dd for a given yyyy-MM-dd + days. */
function shiftFrom(dateStr, days) {
  const p = dateStr.split('-');
  const d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** offset 0 = this month, -1 = last month. */
function monthRange(offset) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const first = new Date(now.getFullYear(), now.getMonth() + offset, 1);
  const last  = new Date(now.getFullYear(), now.getMonth() + offset + 1, 0);
  return {
    from: Utilities.formatDate(first, tz, 'yyyy-MM-dd'),
    to:   Utilities.formatDate(last,  tz, 'yyyy-MM-dd'),
    name: Utilities.formatDate(first, tz, 'MMMM yyyy'),
  };
}

function sumAmounts(rows) {
  return rows.reduce(function (s, r) { return s + Number(r[3] || 0); }, 0);
}

/** Strip characters that break Telegram's legacy Markdown parser. */
function safe(s) {
  return String(s).replace(/[_*`\[\]]/g, '');
}

function isoFromDayMon(day, mon) {
  const mm = MONTHS[titleCase(mon)];
  if (!mm) return null;
  const now = new Date();
  let year = now.getFullYear();
  if (mm === 12 && now.getMonth() === 0) year -= 1;   // Dec txn, Jan alert
  return year + '-' + pad(mm) + '-' + pad(day);
}

function pad(n) { return String(n).padStart(2, '0'); }

function titleCase(s) {
  return String(s).charAt(0).toUpperCase() + String(s).slice(1).toLowerCase();
}
