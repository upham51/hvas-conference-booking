/**
 * Hood View Adventist School: Parent-Teacher Conference Booking
 * Backend: Google Sheets + Apps Script Web App
 */

const SHEET_NAME_CLASSES = 'Classes';
const SHEET_NAME_BOOKINGS = 'Bookings';

// Classes sheet columns. StartTime/EndTime are the default hours; each day can
// override them with Day1Start/Day1End (Wednesday) and Day2Start/Day2End
// (Thursday), and Day1Closed/Day2Closed (TRUE) turns a day off for that class.
// A blank Day cell falls back to StartTime/EndTime.
const CLASS_HEADERS = [
  'ClassID', 'Grade', 'Teacher', 'ParentCode', 'TeacherToken',
  'StartTime', 'EndTime', 'IncrementMin', 'Active', 'BreakMin',
  'Day1Start', 'Day1End', 'Day1Closed', 'Day2Start', 'Day2End', 'Day2Closed',
];
const CONF_DAYS = [
  { key: 'day1', date: '2026-11-04', label: 'Wednesday, Nov 4' },
  { key: 'day2', date: '2026-11-05', label: 'Thursday, Nov 5' },
];
const DEFAULT_START = '14:00';
const DEFAULT_END = '19:00';
const DEFAULT_INCREMENT = 15;

// Gap between back-to-back conferences, set per class in the Classes sheet's
// BreakMin column. A missing column or blank cell counts as the default.
const ALLOWED_BREAKS = [0, 5, 10];
const DEFAULT_BREAK_MIN = 5;

// Windows when no conference can start or run, for every class on both days.
// Slots that would overlap a window are skipped, and the next slot starts
// exactly when the window ends. Parents see the window as a marker in the grid.
const BLOCKED_TIMES = [
  { start: '17:00', end: '17:45', label: 'Dinner break' },
];

// Where teachers open their private schedule page. Their emails link here
// with ?t=<their private token>. Update it if the site moves to another host.
const TEACHER_PAGE_URL = 'https://hvjaparentteacherconference.upham51.workers.dev/teacher';

// The principal's practice class. It works exactly like a real class (parents
// can book it with the code, the teacher page has every feature) but it is
// left out of the dashboard totals. Created by applyLaunchSettings().
const TEST_CLASS_ID = 'principaltest';
const TEST_CLASS = { classId: TEST_CLASS_ID, grade: 'Practice class', teacher: 'Principal test class', parentCode: 'PTEST' };

// The address used to build the "manage my booking" link in confirmation
// emails. This is the school's own short link, which redirects to wherever the
// booking site is hosted, so the links in already-sent emails keep working if
// the site ever moves. The redirect lives in the Wix dashboard under
// SEO & GEO -> Tools and settings -> URL Redirect Manager; if it is ever
// removed, these cancellation links break, so point this straight at the host
// instead of deleting the redirect.
// Verified that the redirect preserves the ?manage= token.
const SITE_BASE_URL = 'https://hvja.org/parentteacher';

// School's own time zone, used when deciding "is tomorrow a conference day"
// for the automatic reminder emails.
const SCHOOL_TIMEZONE = 'America/Los_Angeles';

// Each class lists one or more teacher emails. Every address in the list
// gets an instant email the moment a parent books a conference in that
// class. Use an empty array to skip notifications for a class, nothing
// breaks either way.
// Note: fhaffner@hvja.org (Fae Haffner, 1st Grade co-teacher) was given
// to us as "likely" rather than confirmed, double check it with the office.
const TEACHER_EMAILS = {
  walker: ['walker@hvja.org'],
  leslie: ['mleslie@hvja.org'],
  segoria: ['lsegoria@hvja.org'],
  mcginnis: ['tmcginnis@hvja.org'],
  sexton: ['dsexton@hvja.org', 'fhaffner@hvja.org'],
  neely: ['jneely@hvja.org'],
  white: ['hwhite@hvja.org'],
  pazitka: ['kpazitka@hvja.org'],
  hays: ['bhays@hvja.org'],
  waymire: ['swaymire@hvja.org'],
  morita: ['tmorita@hvja.org'],
  saliba: ['vsaliba@hvja.org'],
  principaltest: ['principal@hvja.org'],
};

const CLASS_SEED = [
  { classId: 'walker',   grade: 'Preschool',     teacher: 'Selina Walker' },
  { classId: 'leslie',   grade: 'Pre-K',          teacher: 'Melissa Leslie' },
  { classId: 'segoria',  grade: 'Forest School',  teacher: 'Lizzie Segoria' },
  { classId: 'mcginnis', grade: 'Kindergarten',   teacher: 'Tamma-Lynn McGinnis' },
  { classId: 'sexton',   grade: '1st Grade',      teacher: 'Debbie Sexton & Fae Haffner' },
  { classId: 'neely',    grade: '2nd Grade',      teacher: 'Jada Neely' },
  { classId: 'white',    grade: '3rd Grade',      teacher: 'Haley White' },
  { classId: 'pazitka',  grade: '4th Grade',      teacher: 'Kristin Pazitka' },
  { classId: 'hays',     grade: '5th Grade',      teacher: 'Brian Hays' },
  { classId: 'waymire',  grade: '6th Grade',      teacher: 'Shelly Waymire' },
  { classId: 'morita',   grade: '7th Grade',      teacher: 'Timothy Morita' },
  { classId: 'saliba',   grade: '8th Grade',      teacher: 'Victoria Saliba' },
];

// Families cannot book a slot that starts less than this many minutes from
// now (school time). Gives the teacher notice and stops someone booking a
// time that already started.
const BOOKING_CUTOFF_MIN = 120;

const ALLOWED_INCREMENTS = [10, 15, 20, 30];
const MAX_NAME_LEN = 100;
const MAX_EMAIL_LEN = 100;
const MAX_PHONE_LEN = 30;
const MIN_ADMIN_PIN_LEN = 8;
const PIN_MAX_FAILS = 5;
const PIN_LOCKOUT_SEC = 15 * 60;
const REQUEST_CACHE_SEC = 15 * 60;
const LOCK_WAIT_MS = 30000;
const MAIL_FROM_NAME = 'Hood View Adventist School';

const BUSY_MSG = 'The site is busy right now. Please wait a moment and try again. Nothing was booked or changed.';
const BAD_TOKEN_MSG = 'This link is no longer valid, it may already have been used. Contact the school office for help.';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Running setup again would wipe every booking and give every teacher a
  // new private link. Refuse once real data exists.
  const existing = ss.getSheetByName(SHEET_NAME_BOOKINGS);
  if (existing && existing.getLastRow() > 1) {
    throw new Error('setup() was stopped: the Bookings tab already has ' +
      (existing.getLastRow() - 1) + ' booking(s), and setup would erase them. Nothing was changed.');
  }

  let classes = ss.getSheetByName(SHEET_NAME_CLASSES);
  if (!classes) classes = ss.insertSheet(SHEET_NAME_CLASSES);
  classes.clear();
  classes.appendRow(CLASS_HEADERS);
  classes.setFrozenRows(1);

  const rows = CLASS_SEED.concat([TEST_CLASS]).map(function (c) { return classRow_(CLASS_HEADERS, c); });
  classes.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  classes.autoResizeColumns(1, CLASS_HEADERS.length);

  let bookings = existing;
  if (!bookings) bookings = ss.insertSheet(SHEET_NAME_BOOKINGS);
  bookings.clear();
  bookings.appendRow([
    'BookingID', 'ClassID', 'Day', 'StartTime', 'EndTime',
    'ParentName', 'StudentName', 'Email', 'Phone', 'Timestamp',
    'ManageToken', 'ReminderSent',
  ]);
  bookings.setFrozenRows(1);

  ss.setActiveSheet(classes);
  notify_(
    'Setup complete. Open the Classes tab to see each teacher\'s parent code ' +
    'and private schedule link. Before anyone uses the principal dashboard, set ' +
    'an admin PIN of at least ' + MIN_ADMIN_PIN_LEN + ' characters: Project Settings > ' +
    'Script properties > ADMIN_PIN.'
  );
}

// Adds any of the named columns that the Classes sheet doesn't have yet.
// Callers hold the lock or run from the editor.
function ensureColumns_(sheet, names) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = names.filter(function (n) { return headers.indexOf(n) === -1; });
  if (missing.length) sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  return missing;
}

function newToken_() {
  const a = Utilities.getUuid().split('-');
  return a[0] + a[1];
}

// One Classes row, laid out by the sheet's own header order.
function classRow_(headers, c) {
  const v = {
    ClassID: c.classId, Grade: c.grade, Teacher: c.teacher,
    ParentCode: c.parentCode || c.classId.toUpperCase(), TeacherToken: newToken_(),
    StartTime: "'" + DEFAULT_START, EndTime: "'" + DEFAULT_END,
    IncrementMin: DEFAULT_INCREMENT, Active: true, BreakMin: DEFAULT_BREAK_MIN,
  };
  return headers.map(function (h) { return Object.prototype.hasOwnProperty.call(v, h) ? v[h] : ''; });
}

// The one function to run after redeploying. It is safe to run again, and it
// never overwrites a choice a teacher or the principal has made:
//   1. adds the BreakMin and per-day columns to the Classes sheet if missing,
//   2. writes the full teacher names,
//   3. gives any class with no break set a 5 minute break (a break someone
//      already chose is left alone),
//   4. adds the principal's practice class (code PTEST) if it isn't there.
function applyLaunchSettings() {
  const outcome = withLock_(function () {
    const sheet = getSheet_(SHEET_NAME_CLASSES);
    const added = ensureColumns_(sheet, CLASS_HEADERS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idCol = headers.indexOf('ClassID'), teacherCol = headers.indexOf('Teacher'), breakCol = headers.indexOf('BreakMin');
    if (idCol === -1 || teacherCol === -1) {
      throw new Error('The Classes tab is missing its ClassID or Teacher column. Nothing was changed.');
    }

    const names = {};
    CLASS_SEED.concat([TEST_CLASS]).forEach(function (c) { names[c.classId] = c.teacher; });
    let renamed = 0, breaksSet = 0, hasTest = false;
    const unmatched = [];
    for (let r = 1; r < values.length; r++) {
      const id = values[r][idCol];
      if (!id) continue;
      if (id === TEST_CLASS_ID) hasTest = true;
      if (Object.prototype.hasOwnProperty.call(names, id)) {
        sheet.getRange(r + 1, teacherCol + 1).setValue(names[id]);
        renamed++;
      } else {
        unmatched.push(id);
      }
      const cur = values[r][breakCol];
      if (cur === '' || cur === null || cur === undefined) {
        sheet.getRange(r + 1, breakCol + 1).setValue(DEFAULT_BREAK_MIN);
        breaksSet++;
      }
    }
    if (!hasTest) sheet.appendRow(classRow_(headers, TEST_CLASS));
    SpreadsheetApp.flush();
    return { added: added, renamed: renamed, breaksSet: breaksSet, addedTest: !hasTest, unmatched: unmatched };
  });
  if (outcome.ok === false) throw new Error(outcome.error);

  notify_('Launch settings applied.\n' +
    'Teacher names updated: ' + outcome.renamed + '.\n' +
    'Classes given a ' + DEFAULT_BREAK_MIN + ' minute break: ' + outcome.breaksSet + ' (breaks already set were kept).\n' +
    (outcome.added.length ? 'Columns added: ' + outcome.added.join(', ') + '.\n' : '') +
    (outcome.addedTest ? 'Practice class added: parent code ' + TEST_CLASS.parentCode + '. Its teacher link is on the principal dashboard.\n' : 'Practice class was already there.\n') +
    (outcome.unmatched.length ? 'Not in the name list, names left alone: ' + outcome.unmatched.join(', ') + '.' : ''));
}

// One-time upgrade for a Bookings sheet that was created before the
// self-serve cancel and reminder-email features existed. Safe to run more
// than once, it only adds columns that are missing.
function migrateAddColumns_v2() {
  const sheet = getSheet_(SHEET_NAME_BOOKINGS);
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const toAdd = ['ManageToken', 'ReminderSent'].filter(function (h) {
    return headers.indexOf(h) === -1;
  });
  if (toAdd.length) {
    sheet.getRange(1, headers.length + 1, 1, toAdd.length).setValues([toAdd]);
  }
  notify_(toAdd.length
    ? 'Added columns: ' + toAdd.join(', ') + '. Existing bookings made before this upgrade will not have a cancel link.'
    : 'Already up to date, nothing to add.');
}

// One-time setup for the automatic day-before reminder emails. Safe to
// re-run, it clears any previous version of this trigger first so you never
// end up with duplicates sending double reminders.
function createDailyReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyReminders_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendDailyReminders_')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .inTimezone(SCHOOL_TIMEZONE)
    .create();
  notify_(
    'Reminder emails are now scheduled. Every day at around 9am, anyone ' +
    'booked for the next day\'s conferences will get a reminder automatically.'
  );
}

// Run from the Apps Script editor before sending the link to families. It
// changes nothing, it only reports what is and isn't ready. Results show in
// the execution log (and as a popup if the spreadsheet is open).
function checkLaunchReadiness() {
  const lines = [];
  const reminder = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'sendDailyReminders_';
  });
  lines.push((reminder ? 'OK   ' : 'FIX  ') + 'Day-before reminder trigger ' +
    (reminder ? 'is installed.' : 'is NOT installed. Run createDailyReminderTrigger().'));

  const pin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || '';
  const pinOk = pin.length >= MIN_ADMIN_PIN_LEN && pin !== '1289';
  lines.push((pinOk ? 'OK   ' : 'FIX  ') + 'Admin PIN ' + (pinOk
    ? 'is set and long enough.'
    : 'is missing, too short, or still the old default. The dashboard stays locked until ADMIN_PIN in Script properties has ' + MIN_ADMIN_PIN_LEN + '+ characters.'));

  const quota = MailApp.getRemainingDailyQuota();
  lines.push((quota >= 500 ? 'OK   ' : 'WARN ') + 'Emails left today: ' + quota +
    '. Each booking sends 2 to 3 emails, so 150 families need about 450.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssTz = ss.getSpreadsheetTimeZone();
  const scriptTz = Session.getScriptTimeZone();
  lines.push((ssTz === scriptTz ? 'OK   ' : 'FIX  ') + 'Time zones: spreadsheet ' + ssTz +
    ', script ' + scriptTz + (ssTz === scriptTz ? '.' : '. These must match or class hours will be read wrong.'));

  const sheet = getSheet_(SHEET_NAME_BOOKINGS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const hasCols = ['BookingID', 'ManageToken', 'ReminderSent'].every(function (h) { return headers.indexOf(h) !== -1; });
  lines.push((hasCols ? 'OK   ' : 'FIX  ') + 'Bookings columns ' + (hasCols ? 'are complete.' : 'are missing. Run migrateAddColumns_v2().'));
  lines.push('INFO ' + (sheet.getLastRow() - 1) + ' booking(s) on the sheet right now.');

  const report = lines.join('\n');
  Logger.log(report);
  notify_(report);
  return report;
}

// Popup when run from the open spreadsheet, log line otherwise.
function notify_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (err) { /* run from the editor */ }
}

function setAdminPin(newPin) {
  if (!newPin || String(newPin).length < MIN_ADMIN_PIN_LEN) {
    throw new Error('The admin PIN must be at least ' + MIN_ADMIN_PIN_LEN + ' characters.');
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', String(newPin));
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    if (action === 'schedule') return jsonOut(getScheduleForCode_(e.parameter.code));
    if (action === 'teacher') return jsonOut(getTeacherView_(e.parameter.token));
    if (action === 'admin') return jsonOut(getAdminView_(e.parameter.pin));
    if (action === 'config') return jsonOut(getPublicConfig_());
    if (action === 'manage') return jsonOut(getManageView_(e.parameter.token));
    return jsonOut({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    console.error(err);
    return jsonOut({ ok: false, error: 'Something went wrong on our end. Please try again in a moment.' });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'book') return jsonOut(bookSlot_(body));
    if (action === 'updateSchedule') return jsonOut(updateSchedule_(body));
    if (action === 'teacherCancel') return jsonOut(teacherCancel_(body));
    if (action === 'cancel') return jsonOut(cancelBooking_(body));
    if (action === 'manageCancel') return jsonOut(cancelByToken_(body));
    return jsonOut({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    console.error(err);
    return jsonOut({ ok: false, error: 'Something went wrong on our end. Nothing was changed. Please try again in a moment.' });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getPublicConfig_() {
  return { ok: true, days: CONF_DAYS };
}

function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function readTable_(sheet) {
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

// ---- Locking and repeat-request protection ----

// Runs fn while holding the script-wide lock. If the lock can't be had in
// 30 seconds, returns a friendly "busy" answer instead of throwing.
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_WAIT_MS)) return { ok: false, error: BUSY_MSG, code: 'busy' };
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// The browser sends a random requestId with every booking and cancel, and
// sends the same one again if it has to retry. The first result is kept for
// 15 minutes, so a retry gets the original answer instead of booking twice.
function cleanRequestId_(id) {
  const s = String(id || '');
  return /^[A-Za-z0-9-]{8,64}$/.test(s) ? s : '';
}
function getCachedResult_(requestId) {
  if (!requestId) return null;
  const hit = CacheService.getScriptCache().get('req_' + requestId);
  return hit ? JSON.parse(hit) : null;
}
function putCachedResult_(requestId, result) {
  if (!requestId) return;
  CacheService.getScriptCache().put('req_' + requestId, JSON.stringify(result), REQUEST_CACHE_SEC);
}

// Writes one new row to Bookings by column name rather than a fixed
// position, so it never misaligns columns if the sheet has (or is missing)
// newer columns like ManageToken and ReminderSent. Every text value gets a
// leading apostrophe, which tells Sheets "store this exactly as typed": no
// turning "18:45" into a time, a phone into a number, or "=..." into a formula.
function appendBookingRow_(fields) {
  const sheet = getSheet_(SHEET_NAME_BOOKINGS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(function (h) {
    if (!Object.prototype.hasOwnProperty.call(fields, h)) return '';
    const v = fields[h];
    return (typeof v === 'string' && v !== '') ? "'" + v : v;
  });
  sheet.appendRow(row);
}

function bookingHeadersHave_(name) {
  const sheet = getSheet_(SHEET_NAME_BOOKINGS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(name) !== -1;
}

function findClassByCode_(code) {
  if (!code) return null;
  const norm = String(code).trim().toUpperCase();
  const classes = readTable_(getSheet_(SHEET_NAME_CLASSES));
  return classes.find(function (c) {
    return String(c.ParentCode).toUpperCase() === norm && c.Active;
  }) || null;
}

function findClassByToken_(token) {
  if (!token) return null;
  const classes = readTable_(getSheet_(SHEET_NAME_CLASSES));
  return classes.find(function (c) { return String(c.TeacherToken) === String(token); }) || null;
}

function findClassById_(classId) {
  const classes = readTable_(getSheet_(SHEET_NAME_CLASSES));
  return classes.find(function (c) { return c.ClassID === classId; }) || null;
}

// ---- Time helpers ----

// Google Sheets silently converts strings like "14:00" into its own time
// serial value when they're written to a cell. When we read that cell back,
// Apps Script hands it to us as a JS Date (epoch-anchored to Dec 30 1899,
// with the hours/minutes we care about). This helper accepts either that
// Date, a real number (fraction-of-a-day serial), or a plain "HH:MM" string,
// and always returns minutes-since-midnight as a number.
function minutesFromTimeValue_(v) {
  if (v instanceof Date) {
    return v.getHours() * 60 + v.getMinutes();
  }
  if (typeof v === 'number') {
    return Math.round(v * 24 * 60);
  }
  const parts = String(v).replace(/^'/, '').split(':');
  return Number(parts[0]) * 60 + Number(parts[1]);
}

function toHHMM_(m) {
  const h = Math.floor(m / 60), mm = m % 60;
  return (h < 10 ? '0' : '') + h + ':' + (mm < 10 ? '0' : '') + mm;
}

// Every booking time is compared, sorted and returned through this, so a
// cell Sheets turned into a time value and a cell stored as "18:45" text
// always look the same.
function slotKey_(v) {
  return toHHMM_(minutesFromTimeValue_(v));
}

// "18:45" -> "6:45 PM", for emails.
function formatTime12_(hhmm) {
  const m = minutesFromTimeValue_(hhmm);
  const h = Math.floor(m / 60), mm = m % 60;
  const h12 = h % 12 === 0 ? 12 : h % 12;
  return h12 + ':' + (mm < 10 ? '0' : '') + mm + ' ' + (h >= 12 ? 'PM' : 'AM');
}

function isBlank_(v) { return v === '' || v === null || v === undefined; }
function isTrue_(v) { return v === true || /^true$/i.test(String(v)); }

// 'day1' -> the Classes columns that hold that day's hours.
function dayColumns_(dayKey) {
  const name = dayKey.charAt(0).toUpperCase() + dayKey.slice(1);
  return { start: name + 'Start', end: name + 'End', closed: name + 'Closed' };
}

// Hours for one conference day: that day's own cells if filled in, otherwise
// the class's default StartTime/EndTime. `closed` means no conferences that day.
function hoursForDay_(cls, dayKey) {
  const c = dayColumns_(dayKey);
  const s = cls[c.start], e = cls[c.end];
  return {
    startTime: slotKey_(isBlank_(s) ? cls.StartTime : s),
    endTime: slotKey_(isBlank_(e) ? cls.EndTime : e),
    closed: isTrue_(cls[c.closed]),
  };
}

function slotsForDay_(cls, dayKey) {
  return generateSlots_(cls, BLOCKED_TIMES, dayKey);
}

// Minutes between conferences for a class. Missing column, blank cell or a
// value outside 0/5/10 all fall back to the default.
function breakMinFor_(cls) {
  const v = cls ? cls.BreakMin : undefined;
  if (v === undefined || v === null || v === '') return DEFAULT_BREAK_MIN;
  const n = Number(v);
  return ALLOWED_BREAKS.indexOf(n) === -1 ? DEFAULT_BREAK_MIN : n;
}

// Slots step by conference length plus the class's break. A slot that would
// overlap a blocked window is skipped and the next one starts exactly when
// the window ends, so the grid restarts after the gap instead of running
// through it. Every caller (booking, parent grid, schedule check, admin
// totals) goes through here, so they always agree.
function generateSlots_(cls, blockedTimes, dayKey) {
  const inc = Number(cls.IncrementMin) || DEFAULT_INCREMENT;
  const brk = breakMinFor_(cls);
  let startMin = minutesFromTimeValue_(cls.StartTime);
  let endMin = minutesFromTimeValue_(cls.EndTime);
  if (dayKey) {
    const h = hoursForDay_(cls, dayKey);
    if (h.closed) return [];
    startMin = minutesFromTimeValue_(h.startTime);
    endMin = minutesFromTimeValue_(h.endTime);
  }
  const blocks = (blockedTimes || BLOCKED_TIMES).map(function (b) {
    return { s: minutesFromTimeValue_(b.start), e: minutesFromTimeValue_(b.end) };
  });
  const slots = [];
  let t = startMin;
  while (t + inc <= endMin) {
    let resumeAt = -1;
    blocks.forEach(function (b) {
      if (t < b.e && b.s < t + inc && b.e > resumeAt) resumeAt = b.e;
    });
    if (resumeAt !== -1) { t = resumeAt; continue; }
    slots.push({ start: toHHMM_(t), end: toHHMM_(t + inc) });
    t += inc + brk;
  }
  return slots;
}

// True if a slot on this conference day starts too soon (or already started)
// to be booked, judged by the clock at the school.
function isPastCutoff_(dayDef, slotStart) {
  const now = new Date();
  const today = Utilities.formatDate(now, SCHOOL_TIMEZONE, 'yyyy-MM-dd');
  if (dayDef.date < today) return true;
  if (dayDef.date > today) return false;
  const nowMin = Number(Utilities.formatDate(now, SCHOOL_TIMEZONE, 'H')) * 60 +
    Number(Utilities.formatDate(now, SCHOOL_TIMEZONE, 'm'));
  return minutesFromTimeValue_(slotStart) < nowMin + BOOKING_CUTOFF_MIN;
}

// True if the slot overlaps any booking on that day. Uses overlap rather
// than "same start time" so a booking can never be hidden by a grid change.
function bookingForSlot_(dayKey, slot, bookings) {
  const s = minutesFromTimeValue_(slot.start), e = minutesFromTimeValue_(slot.end);
  return bookings.find(function (b) {
    if (b.Day !== dayKey) return false;
    const bs = minutesFromTimeValue_(b.StartTime), be = minutesFromTimeValue_(b.EndTime);
    return bs < e && s < be;
  }) || null;
}
function slotOverlapsBooking_(dayKey, slot, bookings) {
  return !!bookingForSlot_(dayKey, slot, bookings);
}

function sortBookings_(list) {
  return list.sort(function (a, b) {
    if (a.Day !== b.Day) return a.Day < b.Day ? -1 : 1;
    return minutesFromTimeValue_(a.StartTime) - minutesFromTimeValue_(b.StartTime);
  });
}

function getBookingsForClass_(classId) {
  const bookings = readTable_(getSheet_(SHEET_NAME_BOOKINGS));
  return bookings.filter(function (b) { return b.ClassID === classId; });
}

// ---- Parent: schedule and booking ----

function getScheduleForCode_(code) {
  const cls = findClassByCode_(code);
  if (!cls) return { ok: false, error: 'That classroom code was not found. Please check with your teacher.' };

  const bookings = getBookingsForClass_(cls.ClassID);

  const days = CONF_DAYS.map(function (d) {
    const closedDay = hoursForDay_(cls, d.key).closed;
    return {
      key: d.key,
      date: d.date,
      label: d.label,
      closed: closedDay,
      slots: slotsForDay_(cls, d.key).map(function (s) {
        const taken = slotOverlapsBooking_(d.key, s, bookings);
        const closed = !taken && isPastCutoff_(d, s.start);
        return { start: s.start, end: s.end, taken: taken || closed, closed: closed };
      }),
    };
  });

  return {
    ok: true,
    classId: cls.ClassID,
    grade: cls.Grade,
    teacher: cls.Teacher,
    incrementMin: Number(cls.IncrementMin) || DEFAULT_INCREMENT,
    breakMin: breakMinFor_(cls),
    blocked: BLOCKED_TIMES,
    days: days,
  };
}

function isValidEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s);
}

function bookSlot_(body) {
  const requestId = cleanRequestId_(body.requestId);
  const cached = getCachedResult_(requestId);
  if (cached) return cached;

  const required = ['code', 'day', 'start', 'parentName', 'studentName', 'email'];
  for (let i = 0; i < required.length; i++) {
    if (!body[required[i]] || !String(body[required[i]]).trim()) {
      return { ok: false, error: 'Please fill in every field.' };
    }
  }
  const parentName = String(body.parentName).trim();
  const studentName = String(body.studentName).trim();
  const email = String(body.email).trim();
  const phone = body.phone ? String(body.phone).trim() : '';
  if (parentName.length > MAX_NAME_LEN || studentName.length > MAX_NAME_LEN) {
    return { ok: false, error: 'Names can be up to ' + MAX_NAME_LEN + ' characters.' };
  }
  if (email.length > MAX_EMAIL_LEN || !isValidEmail_(email)) {
    return { ok: false, error: 'That email address doesn\'t look complete. Please check it, we send your confirmation and cancel link there.' };
  }
  if (phone.length > MAX_PHONE_LEN) {
    return { ok: false, error: 'Phone number can be up to ' + MAX_PHONE_LEN + ' characters.' };
  }

  const cls = findClassByCode_(body.code);
  if (!cls) return { ok: false, error: 'That classroom code was not found.' };

  const dayDef = CONF_DAYS.find(function (d) { return d.key === body.day; });
  if (!dayDef) return { ok: false, error: 'Invalid conference day.' };

  if (hoursForDay_(cls, dayDef.key).closed) {
    return { ok: false, code: 'day_closed', error: 'This teacher is not holding conferences that day. Please pick the other day.' };
  }

  const slot = slotsForDay_(cls, dayDef.key).find(function (s) { return s.start === slotKey_(body.start); });
  if (!slot) {
    return { ok: false, code: 'slot_gone', error: 'That time is no longer on the schedule. Please pick another.' };
  }
  if (isPastCutoff_(dayDef, slot.start)) {
    return { ok: false, code: 'slot_closed', error: 'That time is too close to book online now. Please pick a later time or contact the school office.' };
  }

  let mail = null;
  const result = withLock_(function () {
    // A retry may have waited on the lock while the first attempt finished.
    const again = getCachedResult_(requestId);
    if (again) return again;

    const bookings = getBookingsForClass_(cls.ClassID);
    if (slotOverlapsBooking_(body.day, slot, bookings)) {
      return { ok: false, code: 'slot_taken', error: 'Sorry, someone just booked that time. Please pick another.' };
    }

    // Only issue a manage token (and therefore a cancel link) once the sheet
    // has been migrated to support it.
    const manageToken = bookingHeadersHave_('ManageToken') ? Utilities.getUuid() : null;

    appendBookingRow_({
      BookingID: Utilities.getUuid(),
      ClassID: cls.ClassID,
      Day: body.day,
      StartTime: slot.start,
      EndTime: slot.end,
      ParentName: parentName,
      StudentName: studentName,
      Email: email,
      Phone: phone,
      Timestamp: new Date(),
      ManageToken: manageToken || '',
      ReminderSent: '',
    });
    SpreadsheetApp.flush();

    const res = {
      ok: true,
      confirmation: {
        teacher: cls.Teacher,
        grade: cls.Grade,
        day: dayDef.label,
        start: slot.start,
        end: slot.end,
        manageToken: manageToken,
      },
    };
    putCachedResult_(requestId, res);
    mail = { manageToken: manageToken };
    return res;
  });

  // Emails go out after the lock is released, so one family's emails never
  // hold up another family's booking.
  if (mail) {
    const info = { parentName: parentName, studentName: studentName, email: email, phone: phone };
    sendConfirmation_(cls, dayDef, slot, info, mail.manageToken);
    sendTeacherNotification_(cls, dayDef, slot, info);
  }
  return result;
}

// Original cancel path: requires knowing the exact bookingId and the email
// used to book. Not used by the site; kept for office tooling. Parent-facing
// cancels go through cancelByToken_ instead, see below.
function cancelBooking_(body) {
  if (!body.bookingId || !body.email) return { ok: false, error: 'Missing booking info.' };
  return withLock_(function () {
    const sheet = getSheet_(SHEET_NAME_BOOKINGS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idCol = headers.indexOf('BookingID');
    const emailCol = headers.indexOf('Email');
    for (let r = 1; r < values.length; r++) {
      if (values[r][idCol] === body.bookingId && values[r][emailCol] === body.email) {
        sheet.deleteRow(r + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'Booking not found.' };
  });
}

// Self-serve cancel, used by the "manage my booking" link in the
// confirmation email. The token is random and unguessable, so no email or
// other personal info needs to travel in the URL.
function findBookingByManageToken_(token) {
  if (!token) return null;
  const sheet = getSheet_(SHEET_NAME_BOOKINGS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const tokenCol = headers.indexOf('ManageToken');
  if (tokenCol === -1) return null;
  for (let r = 1; r < values.length; r++) {
    if (values[r][tokenCol] && String(values[r][tokenCol]) === String(token)) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = values[r][i]; });
      obj.__row = r + 1;
      return obj;
    }
  }
  return null;
}

function getManageView_(token) {
  const booking = findBookingByManageToken_(token);
  if (!booking) return { ok: false, error: BAD_TOKEN_MSG };
  const cls = findClassById_(booking.ClassID);
  const dayDef = CONF_DAYS.find(function (d) { return d.key === booking.Day; });
  return {
    ok: true,
    teacher: cls ? cls.Teacher : booking.ClassID,
    grade: cls ? cls.Grade : '',
    parentCode: cls ? cls.ParentCode : '',
    dayLabel: dayDef ? dayDef.label : booking.Day,
    start: slotKey_(booking.StartTime),
    end: slotKey_(booking.EndTime),
    studentName: booking.StudentName,
    parentName: booking.ParentName,
  };
}

function cancelByToken_(body) {
  const token = body && body.token;
  const requestId = cleanRequestId_(body && body.requestId);
  const cached = getCachedResult_(requestId);
  if (cached) return cached;

  let cancelled = null;
  const result = withLock_(function () {
    const again = getCachedResult_(requestId);
    if (again) return again;
    // Find the row again inside the lock: another cancel may have shifted
    // the rows since this request started.
    const booking = findBookingByManageToken_(token);
    if (!booking) return { ok: false, code: 'bad_token', error: BAD_TOKEN_MSG };
    getSheet_(SHEET_NAME_BOOKINGS).deleteRow(booking.__row);
    SpreadsheetApp.flush();
    const cls = findClassById_(booking.ClassID);
    const res = { ok: true, parentCode: cls ? cls.ParentCode : '', classId: booking.ClassID };
    putCachedResult_(requestId, res);
    cancelled = { booking: booking, cls: cls };
    return res;
  });

  if (cancelled) sendCancellationEmails_(cancelled.cls, cancelled.booking);
  return result;
}

// ---- Email ----

function sendMail_(to, subject, body) {
  MailApp.sendEmail(String(to).trim(), subject, body, { name: MAIL_FROM_NAME });
}

function sendConfirmation_(cls, dayDef, slot, info, manageToken) {
  try {
    const subject = 'Conference confirmed: ' + dayDef.label + ' with ' + cls.Teacher;
    let msg = 'Hi ' + info.parentName + ',\n\n' +
      'Your parent-teacher conference is confirmed:\n\n' +
      'Student: ' + info.studentName + '\n' +
      'Teacher: ' + cls.Teacher + ' (' + cls.Grade + ')\n' +
      'Day: ' + dayDef.label + '\n' +
      'Time: ' + formatTime12_(slot.start) + ' to ' + formatTime12_(slot.end) + '\n\n';
    if (manageToken) {
      msg += 'Need to cancel or pick a different time? Use this link:\n' +
        SITE_BASE_URL + '?manage=' + manageToken + '\n\n';
    } else {
      msg += 'If your plans change, please contact the school office.\n\n';
    }
    msg += 'Hood View Adventist School';
    sendMail_(info.email, subject, msg);
  } catch (err) {
    // Email failure should never block a successful booking.
    console.error(err);
  }
}

function sendToTeachers_(classId, subject, msg) {
  const teacherEmails = TEACHER_EMAILS[classId];
  if (!teacherEmails || !teacherEmails.length) return; // not configured for this class
  // Send one at a time, and catch per address, so a typo in one teacher's
  // email (like the unconfirmed Haffner address) never blocks the other.
  teacherEmails.forEach(function (addr) {
    try {
      sendMail_(addr, subject, msg);
    } catch (err) {
      console.error(err);
    }
  });
}

// The teacher's private schedule page, for the "view your schedule" line in
// the emails that go to teachers.
function teacherLink_(cls) {
  return cls && cls.TeacherToken ? TEACHER_PAGE_URL + '?t=' + encodeURIComponent(cls.TeacherToken) : '';
}

function sendTeacherNotification_(cls, dayDef, slot, info) {
  const subject = 'New sign-up: ' + info.studentName + ', ' + dayDef.label + ' ' + formatTime12_(slot.start);
  const msg = 'Hi,\n\n' +
    'A parent just booked a conference time with you:\n\n' +
    'Student: ' + info.studentName + '\n' +
    'Parent: ' + info.parentName + '\n' +
    'Day: ' + dayDef.label + '\n' +
    'Time: ' + formatTime12_(slot.start) + ' to ' + formatTime12_(slot.end) + '\n' +
    'Contact: ' + info.email + (info.phone ? ' / ' + info.phone : '') + '\n\n' +
    'View, print or cancel from your private schedule page:\n' + teacherLink_(cls) + '\n\n' +
    'Hood View Adventist School';
  sendToTeachers_(cls.ClassID, subject, msg);
}

function sendCancellationEmails_(cls, booking, byTeacher) {
  const dayDef = CONF_DAYS.find(function (d) { return d.key === booking.Day; });
  const dayLabel = dayDef ? dayDef.label : booking.Day;
  const teacherName = cls ? cls.Teacher : booking.ClassID;
  const time = formatTime12_(slotKey_(booking.StartTime)) + ' to ' + formatTime12_(slotKey_(booking.EndTime));

  try {
    sendMail_(booking.Email,
      'Conference cancelled: ' + dayLabel + ' with ' + teacherName,
      'Hi ' + booking.ParentName + ',\n\n' +
      (byTeacher
        ? 'Your teacher had to cancel your parent-teacher conference. We are sorry for the change:\n\n'
        : 'Your parent-teacher conference has been cancelled:\n\n') +
      'Student: ' + booking.StudentName + '\n' +
      'Teacher: ' + teacherName + '\n' +
      'Day: ' + dayLabel + '\n' +
      'Time: ' + time + '\n\n' +
      'To pick a new time, go to ' + SITE_BASE_URL +
      (cls ? ' and enter the classroom code ' + cls.ParentCode : '') + '.\n\n' +
      'Hood View Adventist School');
  } catch (err) {
    console.error(err);
  }

  const link = teacherLink_(cls);
  sendToTeachers_(booking.ClassID,
    'Cancelled: ' + booking.StudentName + ', ' + dayLabel + ' ' + formatTime12_(slotKey_(booking.StartTime)),
    'Hi,\n\n' +
    (byTeacher
      ? 'A conference was cancelled from the teacher schedule page, and the family has been emailed. The time is open again for other families.\n\n'
      : 'A family cancelled their conference. The time is open again for other families.\n\n') +
    'Student: ' + booking.StudentName + '\n' +
    'Parent: ' + booking.ParentName + '\n' +
    'Day: ' + dayLabel + '\n' +
    'Time: ' + time + '\n\n' +
    (link ? 'Your schedule: ' + link + '\n\n' : '') +
    'Hood View Adventist School');
}

// Runs once a day (see createDailyReminderTrigger). Sends a reminder to
// every parent whose conference is tomorrow and who hasn't been reminded
// yet, then marks them as reminded so they don't get a second one.
// Emails go out without holding the lock (150 emails would block bookings
// for minutes); the "reminded" marks are written under the lock, finding
// each row again by BookingID in case a cancel moved the rows.
function sendDailyReminders_() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, SCHOOL_TIMEZONE, 'yyyy-MM-dd');

  const dayDef = CONF_DAYS.find(function (d) { return d.date === tomorrowStr; });
  if (!dayDef) return; // no conference tomorrow, nothing to do

  const sheet = getSheet_(SHEET_NAME_BOOKINGS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const col = {};
  headers.forEach(function (h, i) { col[h] = i; });
  if (col.ReminderSent === undefined || col.BookingID === undefined) return; // sheet not migrated

  const classesById = {};
  readTable_(getSheet_(SHEET_NAME_CLASSES)).forEach(function (c) { classesById[c.ClassID] = c; });

  let sentIds = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row[col.Day] !== dayDef.key) continue;
    if (row[col.ReminderSent]) continue;

    const classId = row[col.ClassID];
    const cls = classesById[classId];
    const teacherName = cls ? cls.Teacher : classId;
    const grade = cls ? cls.Grade : '';

    try {
      const subject = 'Reminder: conference with ' + teacherName + ' tomorrow';
      const msg = 'Hi ' + row[col.ParentName] + ',\n\n' +
        'Just a reminder, your parent-teacher conference is tomorrow:\n\n' +
        'Student: ' + row[col.StudentName] + '\n' +
        'Teacher: ' + teacherName + (grade ? ' (' + grade + ')' : '') + '\n' +
        'Day: ' + dayDef.label + '\n' +
        'Time: ' + formatTime12_(slotKey_(row[col.StartTime])) + ' to ' + formatTime12_(slotKey_(row[col.EndTime])) + '\n\n' +
        'See you then.\n\n' +
        'Hood View Adventist School';
      sendMail_(row[col.Email], subject, msg);
      sentIds.push(row[col.BookingID]);
    } catch (err) {
      console.error(err); // keep going, one failed reminder shouldn't stop the rest
    }
    if (sentIds.length >= 25) { markReminded_(sentIds); sentIds = []; }
  }
  if (sentIds.length) markReminded_(sentIds);
}

function markReminded_(bookingIds) {
  const want = {};
  bookingIds.forEach(function (id) { want[id] = true; });
  const lock = LockService.getScriptLock();
  lock.waitLock(5 * 60 * 1000);
  try {
    const sheet = getSheet_(SHEET_NAME_BOOKINGS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idCol = headers.indexOf('BookingID');
    const sentCol = headers.indexOf('ReminderSent');
    const now = new Date();
    for (let r = 1; r < values.length; r++) {
      if (want[values[r][idCol]]) sheet.getRange(r + 1, sentCol + 1).setValue(now);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

// ---- Teacher ----

function bookingOut_(b) {
  const dayDef = CONF_DAYS.find(function (d) { return d.key === b.Day; });
  return {
    bookingId: b.BookingID,
    day: b.Day,
    dayLabel: dayDef ? dayDef.label : b.Day,
    start: slotKey_(b.StartTime),
    end: slotKey_(b.EndTime),
    parentName: b.ParentName,
    studentName: b.StudentName,
    email: b.Email,
    phone: b.Phone,
  };
}

function getTeacherView_(token) {
  const cls = findClassByToken_(token);
  if (!cls) return { ok: false, error: 'Invalid teacher link.' };

  const bookings = sortBookings_(getBookingsForClass_(cls.ClassID));

  // Every slot of each day, with the family in it (or null), so the page can
  // print a full timetable as well as the sign-up list.
  const days = CONF_DAYS.map(function (d) {
    const h = hoursForDay_(cls, d.key);
    return {
      key: d.key,
      date: d.date,
      label: d.label,
      closed: h.closed,
      startTime: h.startTime,
      endTime: h.endTime,
      slots: slotsForDay_(cls, d.key).map(function (s) {
        const b = bookingForSlot_(d.key, s, bookings);
        return { start: s.start, end: s.end, booking: b ? bookingOut_(b) : null };
      }),
    };
  });

  return {
    ok: true,
    classId: cls.ClassID,
    isTest: cls.ClassID === TEST_CLASS_ID,
    grade: cls.Grade,
    teacher: cls.Teacher,
    parentCode: cls.ParentCode,
    startTime: slotKey_(cls.StartTime),
    endTime: slotKey_(cls.EndTime),
    incrementMin: Number(cls.IncrementMin) || DEFAULT_INCREMENT,
    breakMin: breakMinFor_(cls),
    blocked: BLOCKED_TIMES,
    days: days,
    bookings: bookings.map(bookingOut_),
  };
}

// Saves a class's schedule. `body.days` holds each day's hours and whether the
// day is off: { day1: { closed, startTime, endTime }, day2: { ... } }. A page
// that only sends startTime/endTime applies them to both days and leaves each
// day's open/closed state alone. Nothing is saved unless every family already
// booked still lands exactly on a slot of the new schedule.
function updateSchedule_(body) {
  const cls = findClassByToken_(body.token);
  if (!cls) return { ok: false, error: 'Invalid teacher link.' };

  const hhmm = /^([01]\d|2[0-3]):[0-5]\d$/;
  const inc = Number(body.incrementMin);
  // An older page that doesn't send a break keeps the class's current one.
  const brk = (body.breakMin === undefined || body.breakMin === null || body.breakMin === '')
    ? breakMinFor_(cls) : Number(body.breakMin);
  if (ALLOWED_INCREMENTS.indexOf(inc) === -1) {
    return { ok: false, error: 'Conference length must be 10, 15, 20 or 30 minutes.' };
  }
  if (ALLOWED_BREAKS.indexOf(brk) === -1) {
    return { ok: false, error: 'Break between conferences must be none, 5 or 10 minutes.' };
  }

  const plan = {};
  for (let i = 0; i < CONF_DAYS.length; i++) {
    const d = CONF_DAYS[i];
    const cur = hoursForDay_(cls, d.key);
    let src = body.days && typeof body.days === 'object' ? body.days[d.key] : null;
    if (!body.days) src = { startTime: body.startTime, endTime: body.endTime, closed: cur.closed };
    if (!src) return { ok: false, error: 'Please set the hours for ' + d.label + ', or turn that day off.' };

    const closed = src.closed === true || src.closed === 'true';
    let start = String(src.startTime || ''), end = String(src.endTime || '');
    if (closed) {
      // A day that is off keeps whatever hours it had, so turning it back on
      // later restores them.
      if (!hhmm.test(start) || !hhmm.test(end)) { start = cur.startTime; end = cur.endTime; }
    } else {
      if (!hhmm.test(start) || !hhmm.test(end)) return { ok: false, error: 'Please pick a start and end time for ' + d.label + '.' };
      if (minutesFromTimeValue_(start) >= minutesFromTimeValue_(end)) {
        return { ok: false, error: 'End time needs to be after start time on ' + d.label + '.' };
      }
    }
    plan[d.key] = { closed: closed, startTime: start, endTime: end };
  }

  const result = withLock_(function () {
    // Every family already booked must land exactly on a slot of the new
    // schedule. Otherwise they'd drop off the grid, overlap a new slot, or be
    // left on a day that is now off.
    const fits = {};
    CONF_DAYS.forEach(function (d) {
      const p = plan[d.key];
      fits[d.key] = {};
      if (p.closed) return;
      generateSlots_({ StartTime: p.startTime, EndTime: p.endTime, IncrementMin: inc, BreakMin: brk }, BLOCKED_TIMES)
        .forEach(function (s) { fits[d.key][s.start + '-' + s.end] = true; });
    });
    const conflicts = sortBookings_(getBookingsForClass_(cls.ClassID)).filter(function (b) {
      if (!plan[b.Day]) return false;
      if (plan[b.Day].closed) return true;
      return !fits[b.Day][slotKey_(b.StartTime) + '-' + slotKey_(b.EndTime)];
    });
    if (conflicts.length) {
      const list = conflicts.slice(0, 5).map(function (b) {
        const d = CONF_DAYS.find(function (x) { return x.key === b.Day; });
        return (d ? d.label : b.Day) + ' at ' + formatTime12_(slotKey_(b.StartTime));
      }).join('; ');
      return {
        ok: false,
        error: 'Nothing was saved. ' + conflicts.length + (conflicts.length === 1 ? ' family is' : ' families are') +
          ' already booked at a time the new schedule would remove or overlap (' + list +
          (conflicts.length > 5 ? '; and more' : '') + '). Keep hours, a conference length, a break and open days that include those times, ' +
          'or cancel those conferences first from the sign-up list.',
      };
    }

    const sheet = getSheet_(SHEET_NAME_CLASSES);
    ensureColumns_(sheet, CLASS_HEADERS);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const col = function (name) { return headers.indexOf(name) + 1; };
    const idCol = headers.indexOf('ClassID');
    for (let r = 1; r < values.length; r++) {
      if (values[r][idCol] === cls.ClassID) {
        CONF_DAYS.forEach(function (d) {
          const c = dayColumns_(d.key), p = plan[d.key];
          sheet.getRange(r + 1, col(c.start)).setValue("'" + p.startTime);
          sheet.getRange(r + 1, col(c.end)).setValue("'" + p.endTime);
          sheet.getRange(r + 1, col(c.closed)).setValue(p.closed ? true : '');
        });
        // The default hours follow the first open day, so they stay meaningful.
        const firstOpen = CONF_DAYS.find(function (d) { return !plan[d.key].closed; });
        if (firstOpen) {
          sheet.getRange(r + 1, col('StartTime')).setValue("'" + plan[firstOpen.key].startTime);
          sheet.getRange(r + 1, col('EndTime')).setValue("'" + plan[firstOpen.key].endTime);
        }
        sheet.getRange(r + 1, col('IncrementMin')).setValue(inc);
        sheet.getRange(r + 1, col('BreakMin')).setValue(brk);
        break;
      }
    }
    SpreadsheetApp.flush();
    return null;
  });
  if (result) return result;

  return getTeacherView_(body.token);
}

// A teacher cancels one of their own conferences. Only a booking in the
// teacher's own class can be cancelled with their token. The family is told
// by email, and the time opens up for others.
function findBookingById_(classId, bookingId) {
  const sheet = getSheet_(SHEET_NAME_BOOKINGS);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('BookingID'), classCol = headers.indexOf('ClassID');
  if (idCol === -1 || classCol === -1) return null;
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(bookingId) && values[r][classCol] === classId) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = values[r][i]; });
      obj.__row = r + 1;
      return obj;
    }
  }
  return null;
}

function teacherCancel_(body) {
  const requestId = cleanRequestId_(body && body.requestId);
  const cached = getCachedResult_(requestId);
  if (cached) return cached;

  const cls = findClassByToken_(body && body.token);
  if (!cls) return { ok: false, error: 'Invalid teacher link.' };
  if (!body.bookingId) return { ok: false, error: 'Missing booking info.' };

  let cancelled = null;
  const result = withLock_(function () {
    const again = getCachedResult_(requestId);
    if (again) return again;
    // Find the row inside the lock: another cancel may have shifted the rows.
    const found = findBookingById_(cls.ClassID, body.bookingId);
    if (!found) {
      return { ok: false, code: 'not_found', error: 'That conference is no longer on your schedule, it may already have been cancelled. Nothing else was changed.' };
    }
    getSheet_(SHEET_NAME_BOOKINGS).deleteRow(found.__row);
    SpreadsheetApp.flush();
    const res = getTeacherView_(body.token);
    putCachedResult_(requestId, res);
    cancelled = found;
    return res;
  });

  if (cancelled) sendCancellationEmails_(cls, cancelled, true);
  return result;
}

// ---- Principal dashboard ----

function checkAdminPin_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pin_fails') || 0);
  if (fails >= PIN_MAX_FAILS) {
    return 'Too many wrong PINs. The dashboard is locked for 15 minutes. Nothing else is affected.';
  }
  const realPin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || '';
  if (realPin.length < MIN_ADMIN_PIN_LEN) {
    return 'The dashboard is locked until a longer admin PIN is set. In Apps Script, open ' +
      'Project Settings > Script properties and set ADMIN_PIN to at least ' + MIN_ADMIN_PIN_LEN + ' characters.';
  }
  if (!pin || String(pin) !== realPin) {
    cache.put('pin_fails', String(fails + 1), PIN_LOCKOUT_SEC);
    return 'Incorrect PIN.';
  }
  cache.remove('pin_fails');
  return null;
}

function getAdminView_(pin) {
  const pinError = checkAdminPin_(pin);
  if (pinError) return { ok: false, error: pinError };

  const classes = readTable_(getSheet_(SHEET_NAME_CLASSES));
  const allBookings = readTable_(getSheet_(SHEET_NAME_BOOKINGS));

  const classSummaries = classes.map(function (cls) {
    const days = CONF_DAYS.map(function (d) {
      const h = hoursForDay_(cls, d.key);
      return { key: d.key, label: d.label, closed: h.closed, startTime: h.startTime, endTime: h.endTime,
        slotCount: slotsForDay_(cls, d.key).length };
    });
    const totalSlots = days.reduce(function (n, d) { return n + d.slotCount; }, 0);
    const bookings = allBookings.filter(function (b) { return b.ClassID === cls.ClassID; });
    return {
      classId: cls.ClassID,
      isTest: cls.ClassID === TEST_CLASS_ID,
      grade: cls.Grade,
      teacher: cls.Teacher,
      parentCode: cls.ParentCode,
      teacherToken: cls.TeacherToken,
      startTime: slotKey_(cls.StartTime),
      endTime: slotKey_(cls.EndTime),
      days: days,
      incrementMin: Number(cls.IncrementMin) || DEFAULT_INCREMENT,
      breakMin: breakMinFor_(cls),
      totalSlots: totalSlots,
      booked: bookings.length,
      remaining: Math.max(0, totalSlots - bookings.length),
      bookings: sortBookings_(bookings).map(bookingOut_),
    };
  });

  // The practice class is left out of the school-wide numbers.
  const real = classSummaries.filter(function (c) { return !c.isTest; });
  const totalSlots = real.reduce(function (s, c) { return s + c.totalSlots; }, 0);
  const totalBooked = real.reduce(function (s, c) { return s + c.booked; }, 0);

  let emailQuota = null;
  try { emailQuota = MailApp.getRemainingDailyQuota(); } catch (err) { /* not critical */ }

  return {
    ok: true,
    generatedAt: new Date().toISOString(),
    totals: { totalSlots: totalSlots, totalBooked: totalBooked },
    blocked: BLOCKED_TIMES,
    emailQuota: emailQuota,
    classes: classSummaries,
  };
}
