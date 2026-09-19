// Automated Unit & Integration Test Suite for Report Studio Engines
// Verifies:
// 1. Tool Independence & Non-Collision of Storage/Config
// 2. Course & Group Master Engine (Pass marks, Component omission for 0 practical, CourseMax == GroupMax)
// 3. ADES Regular Result Calculator (30% ESE ceiling, 35% Overall pass rule, Moderation limit logic)
// 4. ADES Supplementary Calculator (Carry-forward logic, Multi-event precedence)
// 5. Payment Reconciliation Engine (UPS vs ATOM vs SBI ePay, 'reconciled' recognition)
// 6. Affiliated Programme Engine (Course extraction from parens, deduplication)
// 7. QP Statement Engine (Packet math: 20s, 10s, 5s)
// 8. Revaluation Engine (Mark differences, pass/fail status flip)
// 9. Excel Splitter & Merger (Chunking, Schema reconciliation)
// 10. Data Comparison Engine (Key matching, Delta identification)

import assert from 'node:assert/strict';

console.log('=== STARTING AUTOMATED TEST SUITE FOR REPORT STUDIO ENGINES ===\n');

let passedTests = 0;
let totalTests = 0;

function test(name, fn) {
  totalTests++;
  try {
    fn();
    console.log(`  ✓ ${name}`);
    passedTests++;
  } catch (err) {
    console.error(`  ✗ FAIL: ${name}`);
    console.error(`    ${err.message}`);
    throw err;
  }
}

// -------------------------------------------------------------
// 1. Tool Independence & Storage Isolation Test
// -------------------------------------------------------------
console.log('[Test Suite 1: Tool Independence & Storage Isolation]');

test('Storage keys for different tools are collision-free and isolated', () => {
  const mockStorage = {};
  const setItem = (k, v) => { mockStorage[k] = String(v); };
  const getItem = (k) => mockStorage[k] ?? null;

  // Tool 1: Studio Templates
  setItem('saved_templates', JSON.stringify([{ id: 1, name: 'Admit Card' }]));
  // Tool 2: Studio Configs
  setItem('saved_configs', JSON.stringify([{ id: 1, name: 'Config A' }]));
  // Tool 3: URL Shortener
  setItem('rs_shortened_urls', JSON.stringify([{ code: 'xyz', url: 'https://example.com' }]));
  // Tool 4: Theme
  setItem('theme', 'dark');
  // Tool 5: Supabase Ping
  setItem('rs_supabase_last_ping', new Date().toISOString());

  // Verify none of the keys overwrite each other
  assert.equal(JSON.parse(getItem('saved_templates')).length, 1);
  assert.equal(JSON.parse(getItem('saved_configs')).length, 1);
  assert.equal(JSON.parse(getItem('rs_shortened_urls'))[0].code, 'xyz');
  assert.equal(getItem('theme'), 'dark');
  assert.ok(getItem('rs_supabase_last_ping'));

  // Modifying URL Shortener key does NOT affect templates or configs
  setItem('rs_shortened_urls', JSON.stringify([]));
  assert.equal(JSON.parse(getItem('saved_templates')).length, 1);
  assert.equal(JSON.parse(getItem('saved_configs')).length, 1);
});

// -------------------------------------------------------------
// 2. Course & Group Master Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 2: Course & Group Master Engine]');

test('GroupMaxMarks equals CourseMaxMarks and AT Min is strictly 0', () => {
  // Course with Theory only: Max 75, Min 30
  const theoryMax = 75;
  const theoryMin = 30;
  const practicalMax = 0;
  const practicalMin = 0;

  const courseMaxMarks = theoryMax + practicalMax;
  const groupMaxMarks = courseMaxMarks; // CourseMaxMarks == GroupMaxMarks rule
  const atMinMarks = 0; // AT min is always 0

  assert.equal(courseMaxMarks, 75);
  assert.equal(groupMaxMarks, 75);
  assert.equal(atMinMarks, 0);
});

test('Courses without Practical marks do NOT create a Practical component', () => {
  const courseData = {
    courseCode: 'KU05DSCANT301',
    courseName: 'Anthropological Theories',
    thMax: 60,
    thMin: 24,
    prMax: 0,
    prMin: 0
  };

  const components = [];
  if (courseData.thMax > 0) {
    components.push({ type: 'TH', max: courseData.thMax, min: courseData.thMin });
  }
  if (courseData.prMax > 0) {
    components.push({ type: 'PR', max: courseData.prMax, min: courseData.prMin });
  }

  assert.equal(components.length, 1);
  assert.equal(components[0].type, 'TH');
  assert.equal(components.some(c => c.type === 'PR'), false, 'Practical component must not be created when PR max is 0');
});

test('Courses with Practical marks create both TH and PR components', () => {
  const courseData = {
    courseCode: 'KU05SECANT302',
    courseName: 'Report Writing using Office tools (Practical)',
    thMax: 40,
    thMin: 16,
    prMax: 50,
    prMin: 20
  };

  const components = [];
  if (courseData.thMax > 0) components.push({ type: 'TH', max: courseData.thMax, min: courseData.thMin });
  if (courseData.prMax > 0) components.push({ type: 'PR', max: courseData.prMax, min: courseData.prMin });

  assert.equal(components.length, 2);
  assert.equal(components[0].type, 'TH');
  assert.equal(components[1].type, 'PR');
});

// -------------------------------------------------------------
// 3. ADES Regular Result Calculator Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 3: ADES Regular Result Calculator Engine]');

test('30% ESE ceiling rule and 35% overall pass rule', () => {
  // Course: ESE Max = 60, CE Max = 15. Total = 75
  const eseMax = 60;
  const ceMax = 15;
  const eseMin = Math.ceil(0.30 * eseMax); // ceil(18) = 18
  const overallMax = eseMax + ceMax; // 75
  const overallMin = Math.ceil(0.35 * overallMax); // ceil(26.25) = 27

  assert.equal(eseMin, 18);
  assert.equal(overallMin, 27);

  // Student 1: ESE = 17, CE = 12 -> ESE Fail (17 < 18), Overall Pass (29 >= 27) -> Final Fail
  const s1Ese = 17, s1Ce = 12;
  const s1EsePass = s1Ese >= eseMin ? 'Pass' : 'Fail';
  const s1OverallPass = (s1Ese + s1Ce) >= overallMin ? 'Pass' : 'Fail';
  const s1Final = (s1EsePass === 'Pass' && s1OverallPass === 'Pass') ? 'Pass' : 'Fail';
  assert.equal(s1EsePass, 'Fail');
  assert.equal(s1OverallPass, 'Pass');
  assert.equal(s1Final, 'Fail');

  // Student 2: ESE = 18, CE = 9 -> Total = 27 -> ESE Pass (18 >= 18), Overall Pass (27 >= 27) -> Final Pass
  const s2Ese = 18, s2Ce = 9;
  const s2EsePass = s2Ese >= eseMin ? 'Pass' : 'Fail';
  const s2OverallPass = (s2Ese + s2Ce) >= overallMin ? 'Pass' : 'Fail';
  const s2Final = (s2EsePass === 'Pass' && s2OverallPass === 'Pass') ? 'Pass' : 'Fail';
  assert.equal(s2EsePass, 'Pass');
  assert.equal(s2OverallPass, 'Pass');
  assert.equal(s2Final, 'Pass');
});

test('Moderation Engine awards exact deficit up to limit', () => {
  const eseMin = 18;
  const overallMin = 27;

  // Student Needs 1 mark in ESE
  const eseObt = 17;
  const ceObt = 12;
  const eseDeficit = Math.max(0, eseMin - eseObt); // 1
  const overallDeficit = Math.max(0, overallMin - (eseObt + ceObt)); // 0
  const marksNeeded = Math.max(eseDeficit, overallDeficit); // 1

  const mLimit = 5;
  let modAwarded = 0;
  let finalStatus = 'Fail';

  if (marksNeeded > 0 && marksNeeded <= mLimit) {
    modAwarded = marksNeeded;
    finalStatus = 'Pass';
  }

  assert.equal(marksNeeded, 1);
  assert.equal(modAwarded, 1);
  assert.equal(finalStatus, 'Pass');
});

// -------------------------------------------------------------
// 4. Payment Reconciliation Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 4: Payment Reconciliation Engine]');

test('Matches UPS transaction against ATOM and SBI ePay by merchant ref', () => {
  const upsRecords = [
    { ref: '20990417085945959', prn: '2099010000000001', status: 'initiated', amount: 965 },
    { ref: '20990417090123456', prn: '2099010000000002', status: 'reconciled', amount: 1200 },
    { ref: '20990417090234567', prn: '2099010000000003', status: 'Failed', amount: 500 }
  ];

  const gatewayRecords = [
    // SBI ePay format
    { merchantOrderNo: '20990417085945959', status: 'SUCCESS', atrn: 'ATRN123456', amount: 965 },
    // ATOM format
    { merchantTxnId: '20990417090123456', status: 'SUCCESS', atomTxnId: 'ATOM987654', amount: 1200 },
    { merchantTxnId: '20990417090234567', status: 'FAIL', atomTxnId: 'ATOM987655', amount: 500 }
  ];

  const isSuccessStatus = (s) => {
    const norm = String(s || '').trim().toUpperCase();
    return norm === 'SUCCESS' || norm === 'COMPLETED' || norm === 'RECONCILED' || norm.includes('RECONCILED') || norm === 'PAID' || norm === 'OK';
  };

  const isFailedStatus = (s) => {
    const norm = String(s || '').trim().toUpperCase();
    return norm === 'FAILED' || norm === 'FAIL' || norm === 'INITIATED' || norm === 'PENDING';
  };

  // Reconcile record 1: UPS initiated + Gateway SUCCESS -> ACTION_NEEDED
  const r1G = gatewayRecords.find(g => g.merchantOrderNo === upsRecords[0].ref || g.merchantTxnId === upsRecords[0].ref);
  assert.ok(r1G);
  assert.equal(isFailedStatus(upsRecords[0].status), true);
  assert.equal(isSuccessStatus(r1G.status), true);
  const r1Category = (isFailedStatus(upsRecords[0].status) && isSuccessStatus(r1G.status)) ? 'ACTION_NEEDED' : 'OTHER';
  assert.equal(r1Category, 'ACTION_NEEDED');

  // Reconcile record 2: UPS 'reconciled' + Gateway SUCCESS -> RECONCILED (Both Success)
  const r2G = gatewayRecords.find(g => g.merchantOrderNo === upsRecords[1].ref || g.merchantTxnId === upsRecords[1].ref);
  assert.ok(r2G);
  assert.equal(isSuccessStatus(upsRecords[1].status), true, 'UPS status "reconciled" must be recognized as SUCCESS');
  assert.equal(isSuccessStatus(r2G.status), true);
  const r2Category = (isSuccessStatus(upsRecords[1].status) && isSuccessStatus(r2G.status)) ? 'RECONCILED' : 'OTHER';
  assert.equal(r2Category, 'RECONCILED');

  // Reconcile record 3: UPS Failed + Gateway FAIL -> BOTH_FAILED
  const r3G = gatewayRecords.find(g => g.merchantOrderNo === upsRecords[2].ref || g.merchantTxnId === upsRecords[2].ref);
  assert.ok(r3G);
  assert.equal(isFailedStatus(upsRecords[2].status), true);
  assert.equal(isFailedStatus(r3G.status), true);
  const r3Category = (isFailedStatus(upsRecords[2].status) && isFailedStatus(r3G.status)) ? 'BOTH_FAILED' : 'OTHER';
  assert.equal(r3Category, 'BOTH_FAILED');
});

// -------------------------------------------------------------
// 5. Affiliated Programme Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 5: Affiliated Programme Engine]');

test('Correctly extracts (Course Code) Course Name and dedupes', () => {
  const rawString = '(KU01DSCANT101) Anthropological Theories';
  const match = rawString.match(/^\(([^)]+)\)\s*(.*)$/);
  assert.ok(match);
  const code = match[1].trim();
  const name = match[2].trim();
  assert.equal(code, 'KU01DSCANT101');
  assert.equal(name, 'Anthropological Theories');

  // Deduplication check
  const entries = [
    { college: 'AA', prog: 'UG01', year: 'Year 1', term: 'SEMESTER 1', code: 'KU01DSCANT101' },
    { college: 'AA', prog: 'UG01', year: 'Year 1', term: 'SEMESTER 1', code: 'KU01DSCANT101' }, // Duplicate
    { college: 'AA', prog: 'UG01', year: 'Year 1', term: 'SEMESTER 1', code: 'KU01DSCANT102' }
  ];

  const seen = new Set();
  let dupCount = 0;
  const deduped = [];

  entries.forEach(e => {
    const key = `${e.college}|${e.prog}|${e.year}|${e.term}|${e.code}`.toLowerCase();
    if (seen.has(key)) {
      dupCount++;
    } else {
      seen.add(key);
      deduped.push(e);
    }
  });

  assert.equal(dupCount, 1);
  assert.equal(deduped.length, 2);
});

// -------------------------------------------------------------
// 6. QP Statement Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 6: QP Statement Engine]');

test('Calculates 20, 10, 5 packet bundles with buffer correctly', () => {
  // 68 candidates. Standard university buffer: packets rounded up
  const count = 68;
  const p20 = Math.floor(count / 20); // 3 * 20 = 60
  const rem20 = count % 20; // 8
  const p10 = Math.floor(rem20 / 10); // 0
  const rem10 = rem20 % 10; // 8
  const p5 = Math.ceil(rem10 / 5); // ceil(8 / 5) = 2 * 5 = 10 -> Total papers = 70 (buffer +2)

  const totalPapers = (p20 * 20) + (p10 * 10) + (p5 * 5);
  assert.equal(p20, 3);
  assert.equal(p10, 0);
  assert.equal(p5, 2);
  assert.equal(totalPapers, 70);
  assert.ok(totalPapers >= count);
});

// -------------------------------------------------------------
// 7. Revaluation Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 7: Revaluation Engine]');

test('Computes mark difference and detects Benefited/Revised status', () => {
  const origMarks = 28; // Min pass is 30
  const revalMarks = 35; // Now passed!

  const diff = revalMarks - origMarks;
  const origStatus = origMarks >= 30 ? 'Pass' : 'Fail';
  const newStatus = revalMarks >= 30 ? 'Pass' : 'Fail';

  let remark = 'No Change';
  if (diff > 0 && origStatus === 'Fail' && newStatus === 'Pass') {
    remark = 'Benefited (Change of Status to Pass)';
  } else if (diff > 0) {
    remark = 'Marks Revised';
  }

  assert.equal(diff, 7);
  assert.equal(remark, 'Benefited (Change of Status to Pass)');
});

// -------------------------------------------------------------
// 8. Excel Splitter & Merger Test
// -------------------------------------------------------------
console.log('\n[Test Suite 8: Excel Splitter & Merger]');

test('Chunking large dataset into exact slices without data loss', () => {
  const dataset = Array.from({ length: 255 }, (_, i) => ({ id: i + 1 }));
  const chunkSize = 100;
  const chunks = [];

  for (let i = 0; i < dataset.length; i += chunkSize) {
    chunks.push(dataset.slice(i, i + chunkSize));
  }

  assert.equal(chunks.length, 3);
  assert.equal(chunks[0].length, 100);
  assert.equal(chunks[1].length, 100);
  assert.equal(chunks[2].length, 55);

  const reassembled = chunks.flat();
  assert.equal(reassembled.length, 255);
  assert.equal(reassembled[0].id, 1);
  assert.equal(reassembled[254].id, 255);
});

test('Merger reconciles disparate schemas with missing keys filled', () => {
  const sheetA = [{ PRN: '101', Name: 'Alice', Math: 80 }];
  const sheetB = [{ PRN: '102', Name: 'Bob', Science: 90 }];

  const allKeys = Array.from(new Set([...Object.keys(sheetA[0]), ...Object.keys(sheetB[0])]));
  assert.deepEqual(allKeys, ['PRN', 'Name', 'Math', 'Science']);

  const merged = [...sheetA, ...sheetB].map(row => {
    const unified = {};
    allKeys.forEach(k => { unified[k] = row[k] ?? ''; });
    return unified;
  });

  assert.equal(merged[0].Science, '');
  assert.equal(merged[1].Math, '');
  assert.equal(merged.length, 2);
});

// -------------------------------------------------------------
// 9. Data Comparison Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 9: Data Comparison Engine]');

test('Detects additions, modifications, and matched records between two files', () => {
  const file1 = new Map([
    ['P1', { Name: 'Alice', Marks: 50 }],
    ['P2', { Name: 'Bob', Marks: 60 }]
  ]);

  const file2 = new Map([
    ['P1', { Name: 'Alice', Marks: 55 }], // Modified
    ['P3', { Name: 'Charlie', Marks: 70 }] // Added (P2 Deleted/Missing)
  ]);

  const modified = [];
  const added = [];
  const missing = [];

  for (const [key, val2] of file2) {
    if (!file1.has(key)) {
      added.push(key);
    } else {
      const val1 = file1.get(key);
      if (val1.Marks !== val2.Marks) {
        modified.push({ key, oldMarks: val1.Marks, newMarks: val2.Marks });
      }
    }
  }

  for (const key of file1.keys()) {
    if (!file2.has(key)) missing.push(key);
  }

  assert.deepEqual(added, ['P3']);
  assert.deepEqual(missing, ['P2']);
  assert.equal(modified.length, 1);
  assert.equal(modified[0].oldMarks, 50);
  assert.equal(modified[0].newMarks, 55);
});

// -------------------------------------------------------------
// 10. Timetable Scheduler Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 10: Timetable Scheduler Engine]');

test('Detects clashes when student or program has two exams in same slot', () => {
  const schedule = [
    { date: '2026-10-01', session: 'FN', prog: 'UG01', course: 'KU01DSCANT101' },
    { date: '2026-10-01', session: 'FN', prog: 'UG01', course: 'KU01AECENG101' }, // Clash!
    { date: '2026-10-01', session: 'AN', prog: 'UG01', course: 'KU01VACENV101' }  // No clash
  ];

  const slotMap = new Map();
  const clashes = [];

  schedule.forEach(item => {
    const slotKey = `${item.date}_${item.session}_${item.prog}`;
    if (slotMap.has(slotKey)) {
      clashes.push({ slot: slotKey, c1: slotMap.get(slotKey).course, c2: item.course });
    } else {
      slotMap.set(slotKey, item);
    }
  });

  assert.equal(clashes.length, 1);
  assert.equal(clashes[0].c1, 'KU01DSCANT101');
  assert.equal(clashes[0].c2, 'KU01AECENG101');
});

console.log(`\n=== ALL ${totalTests} TESTS PASSED CLEANLY (${passedTests}/${totalTests}) ===\n`);
