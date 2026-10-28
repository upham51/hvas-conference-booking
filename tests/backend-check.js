// Backend check for Code.gs. Runs the real script against an in-memory
// spreadsheet, so schedule saving, closing a day, teacher cancels, the
// teacher emails and applyLaunchSettings() are exercised without Apps Script.
//
// Run: node tests/backend-check.js

const failures = [];
function expect(name, actual, wanted) {
  const a = JSON.stringify(actual), w = JSON.stringify(wanted);
  if (a === w) { console.log('ok    ' + name); return; }
  failures.push(name);
  console.log('FAIL  ' + name + '\n        got    ' + a + '\n        wanted ' + w);
}

const { load } = require('./gas-harness');

const OLD_HEADERS = ['ClassID', 'Grade', 'Teacher', 'ParentCode', 'TeacherToken', 'StartTime', 'EndTime', 'IncrementMin', 'Active'];
const oldRow = (id, teacher, token) => [id, '1st Grade', teacher, id.toUpperCase(), token, '14:00', '19:00', 15, true];
const book = (t, day, start, extra) => t.call('bookSlot_', Object.assign({ code: 'WALKER', day, start, parentName: 'Pat Parent', studentName: 'Sam', email: 'pat@example.com' }, extra));
const hours = (st, en, closed) => ({ startTime: st, endTime: en, closed: !!closed });

// A class sheet that already has every column, as setup() would leave it.
function freshClass() {
  const t = load([
    ['ClassID', 'Grade', 'Teacher', 'ParentCode', 'TeacherToken', 'StartTime', 'EndTime', 'IncrementMin', 'Active', 'BreakMin', 'Day1Start', 'Day1End', 'Day1Closed', 'Day2Start', 'Day2End', 'Day2Closed'],
    ['walker', 'Preschool', 'Selina Walker', 'WALKER', 'tok-walker', '14:00', '19:00', 15, true, 5, '', '', '', '', '', ''],
    ['leslie', 'Pre-K', 'Melissa Leslie', 'LESLIE', 'tok-leslie', '14:00', '19:00', 15, true, 5, '', '', '', '', '', ''],
  ]);
  return t;
}

// ---- dinner window and break
{
  const t = freshClass();
  const sched = t.call('getScheduleForCode_', 'WALKER');
  expect('dinner skips 5:00 to 5:45 and restarts at 5:45',
    sched.days[0].slots.map((s) => s.start).join(' '),
    '14:00 14:20 14:40 15:00 15:20 15:40 16:00 16:20 16:40 17:45 18:05 18:25 18:45');
  expect('parent schedule carries the blocked window', sched.blocked, [{ start: '17:00', end: '17:45', label: 'Dinner break' }]);
}

// ---- class codes
{
  const t = load([
    ['ClassID', 'Grade', 'Teacher', 'ParentCode', 'TeacherToken', 'StartTime', 'EndTime', 'IncrementMin', 'Active'],
    ['walker', 'Preschool', 'Selina Walker', 'WALKER', 'tok-walker', '14:00', '19:00', 15, true],
    ['segoria', 'Forest School', 'Lizzie Segoria', 'SEGORIA', 'tok-seg', '14:00', '19:00', 15, true],
    ['sexton', '1st Grade', 'Debbie Sexton & Fae Haffner', 'SEXTON', 'tok-sexton', '14:00', '19:00', 15, true],
    ['hays', '5th Grade', 'Brian Hays', 'HAYS', 'tok-hays', '14:00', '19:00', 15, true],
  ]);
  const id = (code) => { const r = t.call('findClassByCode_', code); return r && r.ClassID; };
  expect('the teacher code works in any case', [id('sexton'), id(' Sexton ')], ['sexton', 'sexton']);
  expect('a grade works: firstgrade, 1st grade, 1stgrade, 1st-Grade, grade1',
    ['firstgrade', 'First Grade', '1st grade', '1stgrade', '1st-Grade', 'grade1'].map(id), Array(6).fill('sexton'));
  expect('other grade names work', [id('fifthgrade'), id('forest school'), id('preschool')], ['hays', 'segoria', 'walker']);
  expect('an unknown code is not found', [id('ninthgrade'), id(''), id('!!')], [null, null, null]);

  t.sheets.Classes.rows.push(['sexton2', '1st Grade', 'Someone Else', 'OTHER', 'tok-x', '14:00', '19:00', 15, true]);
  expect('a grade shared by two classes matches neither, a teacher code still works', [id('firstgrade'), id('other')], [null, 'sexton2']);
}

// ---- per-day hours and closing a day
{
  const t = freshClass();
  const save = (days, extra) => t.call('updateSchedule_', Object.assign({ token: 'tok-walker', incrementMin: 15, breakMin: 5, days }, extra));

  let r = save({ day1: hours('14:00', '17:00'), day2: hours('15:00', '19:00') });
  expect('different hours per day save', r.ok, true);
  expect('Wednesday and Thursday get their own slot counts',
    r.days.map((d) => d.slots.length), [9, 10]);

  r = save({ day1: hours('14:00', '19:00'), day2: hours('14:00', '19:00', true) });
  expect('Thursday can be turned off', r.days.map((d) => d.closed), [false, true]);
  const parent = t.call('getScheduleForCode_', 'WALKER');
  expect('parents see Thursday as closed with no times', [parent.days[1].closed, parent.days[1].slots.length], [true, 0]);
  expect('a closed day does not count toward the admin total',
    t.call('getAdminView_', 'a-long-enough-pin').classes.find((c) => c.classId === 'walker').totalSlots, 13);
  expect('booking on a closed day is refused', book(t, 'day2', '15:00').code, 'day_closed');

  r = save({ day1: hours('14:00', '19:00', true), day2: hours('14:00', '19:00') });
  expect('Wednesday can be turned off instead', r.days.map((d) => d.closed), [true, false]);

  r = save({ day1: hours('14:00', '19:00'), day2: hours('14:00', '19:00') });
  expect('turning a day back on restores it', r.days.map((d) => d.closed), [false, false]);

  expect('a break outside none/5/10 is refused', save({ day1: hours('14:00', '19:00'), day2: hours('14:00', '19:00') }, { breakMin: 7 }).ok, false);
  expect('end before start on one day is refused', save({ day1: hours('14:00', '19:00'), day2: hours('16:00', '15:00') }).ok, false);

  // Old pages send one start/end for both days and no open/closed state.
  save({ day1: hours('14:00', '19:00'), day2: hours('14:00', '19:00', true) });
  r = t.call('updateSchedule_', { token: 'tok-walker', startTime: '15:00', endTime: '19:00', incrementMin: 15 });
  expect('an older page keeps a closed day closed', r.days.map((d) => [d.closed, d.startTime]), [[false, '15:00'], [true, '15:00']]);
}

// ---- bookings keep the schedule honest
{
  const t = freshClass();
  const ok = book(t, 'day2', '15:00');
  expect('a family can book', ok.ok, true);
  expect('the booking email goes to the parent and the teacher', t.mail.map((m) => m.to).sort(), ['pat@example.com', 'walker@hvja.org']);

  const closing = t.call('updateSchedule_', { token: 'tok-walker', incrementMin: 15, breakMin: 5,
    days: { day1: hours('14:00', '19:00'), day2: hours('14:00', '19:00', true) } });
  expect('closing a day with a booked family is refused and saves nothing', [closing.ok, /Nothing was saved/.test(closing.error)], [false, true]);
  expect('the day stays open after the refusal', t.call('getTeacherView_', 'tok-walker').days[1].closed, false);

  const moving = t.call('updateSchedule_', { token: 'tok-walker', incrementMin: 15, breakMin: 10,
    days: { day1: hours('14:00', '19:00'), day2: hours('14:00', '19:00') } });
  expect('a new break that moves a booked time is refused', moving.ok, false);
}

// ---- teacher email links
{
  const t = freshClass();
  book(t, 'day1', '14:00');
  const toTeacher = t.mail.find((m) => m.to === 'walker@hvja.org');
  expect('the teacher sign-up email links to their schedule page',
    toTeacher && toTeacher.body.indexOf('/teacher?t=tok-walker') !== -1, true);
}

// ---- teacher cancels
{
  const t = freshClass();
  book(t, 'day1', '14:00');
  const view = t.call('getTeacherView_', 'tok-walker');
  const bookingId = view.bookings[0].bookingId;
  expect('the teacher view shows the booked family in its slot', view.days[0].slots[0].booking.studentName, 'Sam');
  t.mail.length = 0;

  expect('another teacher cannot cancel it', t.call('teacherCancel_', { token: 'tok-leslie', bookingId, requestId: 'cancel-req-0001' }).code, 'not_found');
  expect('a wrong link is refused', t.call('teacherCancel_', { token: 'nope', bookingId }).ok, false);
  expect('the booking is still there', t.sheets.Bookings.rows.length, 2);

  const done = t.call('teacherCancel_', { token: 'tok-walker', bookingId, requestId: 'cancel-req-0002' });
  expect('the teacher can cancel', [done.ok, done.bookings.length, done.days[0].slots[0].booking], [true, 0, null]);
  expect('the booking row is gone', t.sheets.Bookings.rows.length, 1);
  const toParent = t.mail.find((m) => m.to === 'pat@example.com');
  expect('the family is told their teacher cancelled', toParent && /Your teacher had to cancel/.test(toParent.body), true);
  expect('the teacher is emailed with a link', t.mail.some((m) => m.to === 'walker@hvja.org' && /teacher\?t=tok-walker/.test(m.body)), true);

  const mails = t.mail.length;
  const again = t.call('teacherCancel_', { token: 'tok-walker', bookingId, requestId: 'cancel-req-0002' });
  expect('a retried cancel gets the first answer and sends nothing more', [again.ok, t.mail.length], [true, mails]);
  expect('the time is open for families again', t.call('getScheduleForCode_', 'WALKER').days[0].slots[0].taken, false);
}

// ---- launch settings
{
  const t = load([
    OLD_HEADERS,
    oldRow('walker', 'Mrs. Walker', 'tok-walker'),
    oldRow('sexton', 'Mrs. Sexton / Mrs. Haffner', 'tok-sexton'),
  ]);
  t.call('applyLaunchSettings');
  const rows = t.sheets.Classes.rows, h = rows[0];
  const cell = (r, name) => rows[r][h.indexOf(name)];
  expect('launch settings add the new columns', ['BreakMin', 'Day1Start', 'Day2Closed'].every((n) => h.indexOf(n) !== -1), true);
  expect('full names are written', [cell(1, 'Teacher'), cell(2, 'Teacher')], ['Selina Walker', 'Debbie Sexton & Fae Haffner']);
  expect('every class gets a 5 minute break', [cell(1, 'BreakMin'), cell(2, 'BreakMin')], [5, 5]);
  expect('the practice class is added', [cell(3, 'ClassID'), cell(3, 'ParentCode'), cell(3, 'Active')], ['principaltest', 'PTEST', true]);

  t.sheets.Classes.rows[1][h.indexOf('BreakMin')] = 10;
  t.call('applyLaunchSettings');
  expect('running it again keeps a break someone chose', cell(1, 'BreakMin'), 10);
  expect('running it again does not add a second practice class', rows.length, 4);

  const admin = t.call('getAdminView_', 'a-long-enough-pin');
  expect('the practice class is flagged and left out of the totals',
    [admin.classes.filter((c) => c.isTest).length, admin.totals.totalSlots], [1, 46]); // walker 10 a day on a 10 min break + sexton 13 a day, over two days
}

// ---- confirmation email wording, ping, and the admin lockout
{
  const t = freshClass();
  book(t, 'day2', '15:00');
  const toParent = t.mail.find((m) => m.to === 'pat@example.com');
  expect('the confirmation email uses the agreed wording, link on the next line',
    /Need to change your time\? Use this link to cancel, then pick a new time\nhttps?:\/\/\S+\?manage=\S+/.test(toParent.body), true);
  expect('the old wording is gone', /pick a different time/.test(toParent.body), false);

  t.ctx.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: (text) => ({ text, setMimeType() { return this; } }) };
  t.sheets.Classes = null; // a ping must not need the sheet
  expect('ping answers ok without touching the sheet', JSON.parse(t.ctx.doGet({ parameter: { action: 'ping' } }).text), { ok: true });
}

{
  const t = freshClass();
  for (let i = 0; i < 20; i++) t.call('getAdminView_', undefined);
  for (let i = 0; i < 20; i++) t.call('getAdminView_', '');
  expect('requests with no PIN never lock the dashboard', t.call('getAdminView_', 'a-long-enough-pin').ok, true);
  expect('a request with no PIN gets the same error', t.call('getAdminView_', '').error, 'Incorrect PIN.');

  for (let i = 0; i < 5; i++) t.call('getAdminView_', 'wrong-pin-123');
  expect('wrong PINs still lock the dashboard', /Too many wrong PINs/.test(t.call('getAdminView_', 'a-long-enough-pin').error), true);
}

if (failures.length) {
  console.error('\n' + failures.length + ' backend check(s) failed.');
  process.exit(1);
}
console.log('\nAll backend checks pass.');
