/**
 * Hood View Adventist School: Parent-Teacher Conference Booking
 * Backend: Google Sheets + Apps Script Web App
 */

const SHEET_NAME_CLASSES = 'Classes';
const SHEET_NAME_BOOKINGS = 'Bookings';
const CONF_DAYS = [
  { key: 'day1', date: '2026-11-04', label: 'Wednesday, Nov 4' },
  { key: 'day2', date: '2026-11-05', label: 'Thursday, Nov 5' },
];
const DEFAULT_START = '14:00';
const DEFAULT_END = '19:00';
const DEFAULT_INCREMENT = 15;

// The live site's own URL, used to build the "manage my booking" link that
// goes out in the confirmation email.
const SITE_BASE_URL = 'https://hvjaparentteacherconference.upham51.workers.dev/';

// School's own time zone, used when deciding "is tomorrow a conference day"
// for the automatic reminder emails.
const SCHOOL_TIMEZONE = 'America/Los_Angeles';

// Each class lists one or more teacher emails. Every address in the list
// gets an instant email the moment a parent books a conference in that
// class. Use an empty array to skip notifications for a class, nothing
// breaks either way.
// Note: fhaffner@hvja.org (Mrs. Haffner, 1st Grade co-teacher) was given
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
};

const CLASS_SEED = [
  { classId: 'walker',   grade: 'Preschool',     teacher: 'Mrs. Walker' },
  { classId: 'leslie',   grade: 'Pre-K',          teacher: 'Mrs. Leslie' },
  { classId: 'segoria',  grade: 'Forest School',  teacher: 'Ms. Segoria' },
  { classId: 'mcginnis', grade: 'Kindergarten',   teacher: 'Mrs. McGinnis' },
  { classId: 'sexton',   grade: '1st Grade',      teacher: 'Mrs. Sexton / Mrs. Haffner' },
  { classId: 'neely',    grade: '2nd Grade',      teacher: 'Mrs. Neely' },
  { classId: 'white',    grade: '3rd Grade',      teacher: 'Mrs. White' },
  { classId: 'pazitka',  grade: '4th Grade',      teacher: 'Mrs. Pazitka' },
  { classId: 'hays',     grade: '5th Grade',      teacher: 'Mrs. Hays' },
  { classId: 'waymire',  grade: '6th Grade',      teacher: 'Mr. Waymire' },
  { classId: 'morita',   grade: '7th Grade',      teacher: 'Mrs. Morita' },
  { classId: 'saliba',   grade: '8th Grade',      teacher: 'Mr. Saliba' },
];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let classes = ss.getSheetByName(SHEET_NAME_CLASSES);
  if (!classes) classes = ss.insertSheet(SHEET_NAME_CLASSES);
  classes.clear();
  classes.appendRow([
    'ClassID', 'Grade', 'Teacher', 'ParentCode', 'TeacherToken',
    'StartTime', 'EndTime', 'IncrementMin', 'Active',
  ]);
  classes.setFrozenRows(1);

  const rows = CLASS_SEED.map(function (c) {
    return [
      c.classId,
      c.grade,
      c.teacher,
      c.classId.toUpperCase(),
      Utilities.getUuid().split('-')[0] + Utilities.getUuid().split('-')[1],
      DEFAULT_START,
      DEFAULT_END,
      DEFAULT_INCREMENT,
      true,
    ];
  });
  classes.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  classes.autoResizeColumns(1, 9);

  let bookings = ss.getSheetByName(SHEET_NAME_BOOKINGS);
  if (!bookings) bookings = ss.insertSheet(SHEET_NAME_BOOKINGS);
  bookings.clear();
  bookings.appendRow([
    'BookingID', 'ClassID', 'Day', 'StartTime', 'EndTime',
    'ParentName', 'StudentName', 'Email', 'Phone', 'Timestamp',
    'ManageToken', 'ReminderSent',
  ]);
  bookings.setFrozenRows(1);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PIN')) {
    props.setProperty('ADMIN_PIN', '1289');
  }

  ss.setActiveSheet(classes);
  SpreadsheetApp.getUi().alert(
    'Setup complete! Open the Classes tab to see each teacher\'s parent code ' +
    'and private schedule link. Default admin PIN is 1289, change it by ' +
    'running setAdminPin() with your own value.'
  );
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
  if (!toAdd.length) {
    SpreadsheetApp.getUi().alert('Already up to date, nothing to add.');
    return;
  }
  sheet.getRange(1, headers.length + 1, 1, toAdd.length).setValues([toAdd]);
  SpreadsheetApp.getUi().alert(
    'Added columns: ' + toAdd.join(', ') + '. New bookings from now on will ' +
    'support self-serve cancel links and reminder emails. Existing bookings ' +
    'made before this upgrade will not have a cancel link.'
  );
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
    .create();
  SpreadsheetApp.getUi().alert(
    'Reminder emails are now scheduled. Every day at around 9am, anyone ' +
    'booked for the next day\'s conferences will get a reminder automatically.'
  );
}

function setAdminPin(newPin) {
  if (!newPin) {
    throw new Error('Call setAdminPin("yourNewPin") with a value.');
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
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'book') return jsonOut(bookSlot_(body));
    if (action === 'updateSchedule') return jsonOut(updateSchedule_(body));
    if (action === 'cancel') return jsonOut(cancelBooking_(body));
    if (action === 'manageCancel') return jsonOut(cancelByToken_(body.token));
    return jsonOut({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
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

// Writes one new row to Bookings by column name rather than a fixed
// position, so it never misaligns columns if the sheet has (or is missing)
// newer columns like ManageToken and ReminderSent.
function appendBookingRow_(fields) {
  const sheet = getSheet_(SHEET_NAME_BOOKINGS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(function (h) {
    return Object.prototype.hasOwnProperty.call(fields, h) ? fields[h] : '';
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
  const parts = String(v).split(':');
  return Number(parts[0]) * 60 + Number(parts[1]);
}

function toHHMM_(m) {
  const h = Math.floor(m / 60), mm = m % 60;
  return (h < 10 ? '0' : '') + h + ':' + (mm < 10 ? '0' : '') + mm;
}

function generateSlots_(cls) {
  const inc = Number(cls.IncrementMin) || DEFAULT_INCREMENT;
  const startMin = minutesFromTimeValue_(cls.StartTime);
  const endMin = minutesFromTimeValue_(cls.EndTime);
  const slots = [];
  for (let t = startMin; t + inc <= endMin; t += inc) {
    slots.push({ start: toHHMM_(t), end: toHHMM_(t + inc) });
  }
  return slots;
}

function getBookingsForClass_(classId) {
  const bookings = readTable_(getSheet_(SHEET_NAME_BOOKINGS));
  return bookings.filter(function (b) { return b.ClassID === classId; });
}

function getScheduleForCode_(code) {
  const cls = findClassByCode_(code);
  if (!cls) return { ok: false, error: 'That classroom code was not found. Please check with your teacher.' };

  const slots = generateSlots_(cls);
  const bookings = getBookingsForClass_(cls.ClassID);

  const days = CONF_DAYS.map(function (d) {
    const taken = {};
    bookings.filter(function (b) { return b.Day === d.key; }).forEach(function (b) {
      taken[b.StartTime] = true;
    });
    return {
      key: d.key,
      date: d.date,
      label: d.label,
      slots: slots.map(function (s) {
        return { start: s.start, end: s.end, taken: !!taken[s.start] };
      }),
    };
  });

  return {
    ok: true,
    classId: cls.ClassID,
    grade: cls.Grade,
    teacher: cls.Teacher,
    days: days,
  };
}

function bookSlot_(body) {
  const required = ['code', 'day', 'start', 'parentName', 'studentName', 'email'];
  for (let i = 0; i < required.length; i++) {
    if (!body[required[i]] || !String(body[required[i]]).trim()) {
      return { ok: false, error: 'Please fill in every field.' };
    }
  }

  const cls = findClassByCode_(body.code);
  if (!cls) return { ok: false, error: 'That classroom code was not found.' };

  const dayDef = CONF_DAYS.find(function (d) { return d.key === body.day; });
  if (!dayDef) return { ok: false, error: 'Invalid conference day.' };

  const validSlots = generateSlots_(cls);
  const slot = validSlots.find(function (s) { return s.start === body.start; });
  if (!slot) return { ok: false, error: 'That time slot is no longer available. Please pick another.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const bookings = getBookingsForClass_(cls.ClassID);
    const alreadyTaken = bookings.some(function (b) {
      return b.Day === body.day && b.StartTime === body.start;
    });
    if (alreadyTaken) {
      return { ok: false, error: 'Sorry, someone just booked that slot. Please pick another.' };
    }

    const bookingId = Utilities.getUuid();
    // Only issue a manage token (and therefore a cancel link) once the sheet
    // has been migrated to support it. Older sheets keep working exactly as
    // before, they just won't get a cancel link in the email.
    const manageToken = bookingHeadersHave_('ManageToken') ? Utilities.getUuid() : null;

    appendBookingRow_({
      BookingID: bookingId,
      ClassID: cls.ClassID,
      Day: body.day,
      StartTime: slot.start,
      EndTime: slot.end,
      ParentName: String(body.parentName).trim(),
      StudentName: String(body.studentName).trim(),
      Email: String(body.email).trim(),
      Phone: body.phone ? String(body.phone).trim() : '',
      Timestamp: new Date(),
      ManageToken: manageToken || '',
      ReminderSent: '',
    });

    sendConfirmation_(cls, dayDef, slot, body, manageToken);
    sendTeacherNotification_(cls, dayDef, slot, body);

    return {
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
  } finally {
    lock.releaseLock();
  }
}

// Original cancel path: requires knowing the exact bookingId and the email
// used to book. Kept for backward compatibility (e.g. office/admin tooling).
// Parent-facing cancels go through cancelByToken_ instead, see below.
function cancelBooking_(body) {
  if (!body.bookingId || !body.email) return { ok: false, error: 'Missing booking info.' };
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
    if (values[r][tokenCol] && values[r][tokenCol] === token) {
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
  if (!booking) {
    return {
      ok: false,
      error: 'This link is no longer valid, it may already have been used. Contact the school office for help.',
    };
  }
  const cls = findClassById_(booking.ClassID);
  const dayDef = CONF_DAYS.find(function (d) { return d.key === booking.Day; });
  return {
    ok: true,
    teacher: cls ? cls.Teacher : booking.ClassID,
    grade: cls ? cls.Grade : '',
    parentCode: cls ? cls.ParentCode : '',
    dayLabel: dayDef ? dayDef.label : booking.Day,
    start: toHHMM_(minutesFromTimeValue_(booking.StartTime)),
    end: toHHMM_(minutesFromTimeValue_(booking.EndTime)),
    studentName: booking.StudentName,
    parentName: booking.ParentName,
  };
}

function cancelByToken_(token) {
  const booking = findBookingByManageToken_(token);
  if (!booking) {
    return {
      ok: false,
      error: 'This link is no longer valid, it may already have been used. Contact the school office for help.',
    };
  }
  const cls = findClassById_(booking.ClassID);
  const sheet = getSheet_(SHEET_NAME_BOOKINGS);
  sheet.deleteRow(booking.__row);
  return { ok: true, parentCode: cls ? cls.ParentCode : '', classId: booking.ClassID };
}

function sendConfirmation_(cls, dayDef, slot, body, manageToken) {
  try {
    const subject = 'Conference Confirmed: ' + dayDef.label + ' with ' + cls.Teacher;
    let msg = 'Hi ' + body.parentName + ',\n\n' +
      'Your parent-teacher conference is confirmed:\n\n' +
      'Student: ' + body.studentName + '\n' +
      'Teacher: ' + cls.Teacher + ' (' + cls.Grade + ')\n' +
      'Day: ' + dayDef.label + '\n' +
      'Time: ' + slot.start + ' - ' + slot.end + '\n\n';
    if (manageToken) {
      msg += 'Need to cancel or pick a different time? Use this link:\n' +
        SITE_BASE_URL + '?manage=' + manageToken + '\n\n';
    } else {
      msg += 'If your plans change, please contact the school office.\n\n';
    }
    msg += 'Hood View Adventist School';
    MailApp.sendEmail(String(body.email).trim(), subject, msg);
  } catch (err) {
    // Email failure should never block a successful booking.
  }
}

function sendTeacherNotification_(cls, dayDef, slot, body) {
  const teacherEmails = TEACHER_EMAILS[cls.ClassID];
  if (!teacherEmails || !teacherEmails.length) return; // not configured yet for this class
  const subject = 'New sign-up: ' + body.studentName + ', ' + dayDef.label + ' ' + slot.start;
  const msg = 'Hi,\n\n' +
    'A parent just booked a conference time with you:\n\n' +
    'Student: ' + body.studentName + '\n' +
    'Parent: ' + body.parentName + '\n' +
    'Day: ' + dayDef.label + '\n' +
    'Time: ' + slot.start + ' - ' + slot.end + '\n' +
    'Contact: ' + body.email + (body.phone ? ' / ' + body.phone : '') + '\n\n' +
    'View your full schedule any time at your private link.\n\n' +
    'Hood View Adventist School';
  // Send one at a time, and catch per address, so a typo in one teacher's
  // email (like the unconfirmed Haffner address) never blocks the booking
  // or the other teacher's notification.
  teacherEmails.forEach(function (addr) {
    try {
      MailApp.sendEmail(addr, subject, msg);
    } catch (err) {
      // Never block a booking over a notification failure.
    }
  });
}

// Runs once a day (see createDailyReminderTrigger). Sends a reminder to
// every parent whose conference is tomorrow and who hasn't been reminded
// yet, then marks them as reminded so they don't get a second one.
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
  if (col.ReminderSent === undefined) return; // sheet not migrated yet

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row[col.Day] !== dayDef.key) continue;
    if (row[col.ReminderSent]) continue;

    const classId = row[col.ClassID];
    const cls = findClassById_(classId);
    const teacherName = cls ? cls.Teacher : classId;
    const grade = cls ? cls.Grade : '';
    const email = row[col.Email];

    try {
      const subject = 'Reminder: Conference with ' + teacherName + ' tomorrow';
      const msg = 'Hi ' + row[col.ParentName] + ',\n\n' +
        'Just a reminder, your parent-teacher conference is tomorrow:\n\n' +
        'Student: ' + row[col.StudentName] + '\n' +
        'Teacher: ' + teacherName + (grade ? ' (' + grade + ')' : '') + '\n' +
        'Day: ' + dayDef.label + '\n' +
        'Time: ' + toHHMM_(minutesFromTimeValue_(row[col.StartTime])) + ' - ' + toHHMM_(minutesFromTimeValue_(row[col.EndTime])) + '\n\n' +
        'See you then!\n\n' +
        'Hood View Adventist School';
      MailApp.sendEmail(String(email).trim(), subject, msg);
      sheet.getRange(r + 1, col.ReminderSent + 1).setValue(new Date());
    } catch (err) {
      // Keep going, one failed reminder shouldn't stop the rest.
    }
  }
}

function getTeacherView_(token) {
  const cls = findClassByToken_(token);
  if (!cls) return { ok: false, error: 'Invalid teacher link.' };

  const bookings = getBookingsForClass_(cls.ClassID).sort(function (a, b) {
    if (a.Day !== b.Day) return a.Day < b.Day ? -1 : 1;
    return a.StartTime < b.StartTime ? -1 : 1;
  });

  return {
    ok: true,
    classId: cls.ClassID,
    grade: cls.Grade,
    teacher: cls.Teacher,
    parentCode: cls.ParentCode,
    startTime: toHHMM_(minutesFromTimeValue_(cls.StartTime)),
    endTime: toHHMM_(minutesFromTimeValue_(cls.EndTime)),
    incrementMin: cls.IncrementMin,
    days: CONF_DAYS,
    bookings: bookings.map(function (b) {
      const dayDef = CONF_DAYS.find(function (d) { return d.key === b.Day; });
      return {
        day: b.Day,
        dayLabel: dayDef ? dayDef.label : b.Day,
        start: b.StartTime,
        end: b.EndTime,
        parentName: b.ParentName,
        studentName: b.StudentName,
        email: b.Email,
        phone: b.Phone,
      };
    }),
  };
}

function updateSchedule_(body) {
  const cls = findClassByToken_(body.token);
  if (!cls) return { ok: false, error: 'Invalid teacher link.' };

  const start = body.startTime, end = body.endTime, inc = Number(body.incrementMin);
  if (!start || !end || !inc) return { ok: false, error: 'Missing schedule fields.' };

  const sheet = getSheet_(SHEET_NAME_CLASSES);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('ClassID');
  const startCol = headers.indexOf('StartTime');
  const endCol = headers.indexOf('EndTime');
  const incCol = headers.indexOf('IncrementMin');

  for (let r = 1; r < values.length; r++) {
    if (values[r][idCol] === cls.ClassID) {
      sheet.getRange(r + 1, startCol + 1).setValue(start);
      sheet.getRange(r + 1, endCol + 1).setValue(end);
      sheet.getRange(r + 1, incCol + 1).setValue(inc);
      break;
    }
  }

  return getTeacherView_(body.token);
}

function getAdminView_(pin) {
  const realPin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!pin || String(pin) !== String(realPin)) {
    return { ok: false, error: 'Incorrect PIN.' };
  }

  const classes = readTable_(getSheet_(SHEET_NAME_CLASSES));
  const allBookings = readTable_(getSheet_(SHEET_NAME_BOOKINGS));

  const classSummaries = classes.map(function (cls) {
    const slots = generateSlots_(cls);
    const totalSlots = slots.length * CONF_DAYS.length;
    const bookings = allBookings.filter(function (b) { return b.ClassID === cls.ClassID; });
    return {
      classId: cls.ClassID,
      grade: cls.Grade,
      teacher: cls.Teacher,
      parentCode: cls.ParentCode,
      teacherToken: cls.TeacherToken,
      startTime: toHHMM_(minutesFromTimeValue_(cls.StartTime)),
      endTime: toHHMM_(minutesFromTimeValue_(cls.EndTime)),
      incrementMin: cls.IncrementMin,
      totalSlots: totalSlots,
      booked: bookings.length,
      remaining: totalSlots - bookings.length,
      bookings: bookings
        .sort(function (a, b) {
          if (a.Day !== b.Day) return a.Day < b.Day ? -1 : 1;
          return a.StartTime < b.StartTime ? -1 : 1;
        })
        .map(function (b) {
          const dayDef = CONF_DAYS.find(function (d) { return d.key === b.Day; });
          return {
            day: b.Day,
            dayLabel: dayDef ? dayDef.label : b.Day,
            start: b.StartTime,
            end: b.EndTime,
            parentName: b.ParentName,
            studentName: b.StudentName,
            email: b.Email,
            phone: b.Phone,
          };
        }),
    };
  });

  const totalSlots = classSummaries.reduce(function (s, c) { return s + c.totalSlots; }, 0);
  const totalBooked = classSummaries.reduce(function (s, c) { return s + c.booked; }, 0);

  return {
    ok: true,
    generatedAt: new Date().toISOString(),
    totals: { totalSlots: totalSlots, totalBooked: totalBooked },
    classes: classSummaries,
  };
}
