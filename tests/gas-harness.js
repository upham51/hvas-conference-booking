// In-memory stand-ins for the Apps Script services, so Code.gs can run in
// plain Node. Used by backend-check.js.

const fs = require('fs');
const path = require('path');
const vm = require('vm');

function makeSheet(rows) {
  const width = () => Math.max(1, ...rows.map((r) => r.length));
  const clean = (v) => (typeof v === 'string' && v[0] === "'" ? v.slice(1) : v);
  const set = (r, c, v) => {
    while (rows.length < r) rows.push([]);
    while (rows[r - 1].length < c) rows[r - 1].push('');
    rows[r - 1][c - 1] = clean(v);
  };
  return {
    rows,
    getLastRow: () => rows.length,
    getLastColumn: width,
    getDataRange: () => ({ getValues: () => rows.map((r) => { const o = r.slice(); while (o.length < width()) o.push(''); return o; }) }),
    getRange: (r, c, nr, nc) => ({
      getValues: () => {
        const out = [];
        for (let i = 0; i < (nr || 1); i++) {
          const row = [];
          for (let j = 0; j < (nc || 1); j++) row.push((rows[r - 1 + i] || [])[c - 1 + j] === undefined ? '' : rows[r - 1 + i][c - 1 + j]);
          out.push(row);
        }
        return out;
      },
      setValue: (v) => set(r, c, v),
      setValues: (vs) => vs.forEach((vr, i) => vr.forEach((v, j) => set(r + i, c + j, v))),
    }),
    appendRow: (a) => rows.push(a.map(clean)),
    deleteRow: (n) => rows.splice(n - 1, 1),
  };
}

function load(classRows, bookingRows) {
  const sheets = {
    Classes: makeSheet(classRows),
    Bookings: makeSheet(bookingRows || [['BookingID', 'ClassID', 'Day', 'StartTime', 'EndTime', 'ParentName', 'StudentName', 'Email', 'Phone', 'Timestamp', 'ManageToken', 'ReminderSent']]),
  };
  const mail = [];
  const cache = {};
  let uuid = 0;
  const RealDate = Date;
  // The clock is fixed before the conference days so booking is never "too late".
  class FakeDate extends RealDate {
    constructor(...a) { if (a.length) super(...a); else super('2026-10-01T19:00:00Z'); }
    static now() { return new RealDate('2026-10-01T19:00:00Z').getTime(); }
  }
  const ctx = {
    Date: FakeDate, console, JSON, Math, Number, String, Object, Array, RegExp, Error,
    Logger: { log() {} },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: (n) => sheets[n] }), flush() {} },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
    CacheService: { getScriptCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'a-long-enough-pin' }) },
    Utilities: {
      getUuid: () => 'u' + String(++uuid).padStart(7, '0') + '-aaaa-bbbb-cccc-dddddddddddd',
      formatDate: (d, tz, f) => {
        const p = Object.fromEntries(new Intl.DateTimeFormat('en-US', { timeZone: tz, hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: 'numeric', minute: 'numeric' })
          .formatToParts(d).map((x) => [x.type, x.value]));
        return f === 'yyyy-MM-dd' ? `${p.year}-${p.month}-${p.day}` : f === 'H' ? String(Number(p.hour) % 24) : String(Number(p.minute));
      },
    },
    MailApp: { sendEmail: (to, subject, body) => mail.push({ to, subject, body }), getRemainingDailyQuota: () => 100 },
    ContentService: {},
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), ctx);
  const call = (fn, ...args) => { ctx.__a = args; const out = JSON.stringify(vm.runInContext(fn + '.apply(null, __a)', ctx)); return out === undefined ? undefined : JSON.parse(out); };
  return { sheets, mail, call, ctx };
}

module.exports = { load, makeSheet };
