// Slot grid check for Code.gs: conference length, break between conferences
// and the dinner-break window. Loads Code.gs as plain JavaScript, no Apps
// Script services are called.
//
// Run: node tests/slots-check.js

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ctx = {};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), ctx);

const slots = (cls) => vm.runInContext('generateSlots_(' + JSON.stringify(cls) + ')', ctx);
const starts = (cls) => slots(cls).map((s) => s.start).join(' ');

const failures = [];
function expect(name, actual, wanted) {
  if (actual === wanted) { console.log('ok    ' + name); return; }
  failures.push(name);
  console.log('FAIL  ' + name + '\n        got    ' + actual + '\n        wanted ' + wanted);
}

const day = { StartTime: '14:00', EndTime: '19:00', IncrementMin: 15 };

expect('2:00 to 7:00, 15 min, 5 min break',
  starts({ ...day, BreakMin: 5 }),
  '14:00 14:20 14:40 15:00 15:20 15:40 16:00 16:20 16:40 17:45 18:05 18:25 18:45');
expect('2:00 start gives 13 slots', slots({ ...day, BreakMin: 5 }).length, 13);
expect('3:00 start gives 10 slots', slots({ ...day, StartTime: '15:00', BreakMin: 5 }).length, 10);
expect('blank break counts as 5 minutes', starts({ ...day, BreakMin: '' }), starts({ ...day, BreakMin: 5 }));
expect('missing break counts as 5 minutes', starts(day), starts({ ...day, BreakMin: 5 }));
expect('a break outside 0, 5, 10 counts as 5 minutes', starts({ ...day, BreakMin: 7 }), starts({ ...day, BreakMin: 5 }));
expect('no break, dinner still skipped',
  starts({ ...day, BreakMin: 0 }).split(' ').slice(11, 14).join(' '), '16:45 17:45 18:00');
expect('10 min break, 30 min conferences',
  slots({ ...day, IncrementMin: 30, BreakMin: 10 }).map((s) => s.start + '-' + s.end).join(' '),
  '14:00-14:30 14:40-15:10 15:20-15:50 16:00-16:30 17:45-18:15 18:25-18:55');

const overlapsDinner = slots({ ...day, BreakMin: 5 }).filter((s) => s.start < '17:45' && s.end > '17:00');
expect('no slot touches 5:00 to 5:45', overlapsDinner.length, 0);

if (failures.length) {
  console.error('\n' + failures.length + ' slot check(s) failed.');
  process.exit(1);
}
console.log('\nAll slot checks pass.');
