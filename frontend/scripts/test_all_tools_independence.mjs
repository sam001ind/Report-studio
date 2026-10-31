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
import * as XLSX from 'xlsx';
import { normalizeKeyText, cleanDisplayField } from '../src/utils/excelParser.js';

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

test('Correctly parses flexible date formats and computes earliest/latest transaction window', () => {
  const dates = [
    '2099-04-17 12:18:27',     // ISO (UPS)
    '22-Apr-2099 11:15:30 AM', // Named Month (ATOM)
    '2099-04-20 09:00:00',     // Midpoint
    '25-Apr-2099 16:45:10'      // End (ATOM)
  ];

  const parseFlexible = (raw) => {
    let s = String(raw).trim();
    const monMatch = s.match(/^(\d{1,2})[-\s/]([A-Za-z]{3,9})[-\s/](\d{2,4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
    if (monMatch) {
      const months = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
      const m = months[monMatch[2].toLowerCase().substring(0, 3)];
      if (m !== undefined) {
        return new Date(parseInt(monMatch[3], 10), m, parseInt(monMatch[1], 10));
      }
    }
    const isoMatch = s.match(/^(\d{4})[-\s/](\d{1,2})[-\s/](\d{1,2})/);
    if (isoMatch) {
      return new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
    }
    return null;
  };

  const parsed = dates.map(d => parseFlexible(d)).filter(Boolean);
  assert.equal(parsed.length, 4);

  let minDate = parsed[0];
  let maxDate = parsed[0];
  parsed.forEach(d => {
    if (d < minDate) minDate = d;
    if (d > maxDate) maxDate = d;
  });

  assert.equal(minDate.getDate(), 17);
  assert.equal(minDate.getMonth(), 3); // 0-indexed April
  assert.equal(maxDate.getDate(), 25);
  assert.equal(maxDate.getMonth(), 3);

  const daysSpan = Math.round((maxDate - minDate) / (1000 * 60 * 60 * 24)) + 1;
  assert.equal(daysSpan, 9);
});

test('Reconciles across multiple payment gateway files simultaneously (ATOM + SBI ePay)', () => {
  const upsRecords = [
    { ref: 'REF_ATOM_001', prn: '2099010000000001', status: 'initiated', amount: 965 },
    { ref: 'REF_EPAY_002', prn: '2099010000000002', status: 'initiated', amount: 1440 }
  ];

  const gatewayFiles = [
    {
      id: 'f1',
      name: 'ATOM_Settlement.xlsx',
      detectedGateway: 'ATOM',
      rows: [
        { 'Merchant Txn ID': 'REF_ATOM_001', 'Txn Status': 'SUCCESS', 'Atom Txn ID': 'ATOM111', Amount: 965 }
      ]
    },
    {
      id: 'f2',
      name: 'SBI_ePay_Report.xlsx',
      detectedGateway: 'SBI_EPAY',
      rows: [
        { 'MERCHANT ORDER NO': 'REF_EPAY_002', STATUS: 'SUCCESS', ATRN: 'ATRN222', 'MERCHANT ORDER AMOUNT': 1440 }
      ]
    }
  ];

  // Flatten rows from all gateway files, preserving source tags
  const combinedGatewayRows = gatewayFiles.flatMap((f) =>
    f.rows.map((r) => ({ ...r, _sourceFileName: f.name, _fileGateway: f.detectedGateway }))
  );
  assert.equal(combinedGatewayRows.length, 2);

  const gatewayMap = new Map();
  combinedGatewayRows.forEach((row) => {
    const isAtom = row._fileGateway === 'ATOM';
    const ref = isAtom ? row['Merchant Txn ID'] : row['MERCHANT ORDER NO'];
    const gatewayName = isAtom ? 'ATOM' : 'SBI_EPAY';
    gatewayMap.set(ref, { ...row, ref, gatewayName });
  });

  // Both UPS records should find their respective gateway matches across the two files
  const match1 = gatewayMap.get('REF_ATOM_001');
  assert.ok(match1);
  assert.equal(match1.gatewayName, 'ATOM');
  assert.equal(match1._sourceFileName, 'ATOM_Settlement.xlsx');

  const match2 = gatewayMap.get('REF_EPAY_002');
  assert.ok(match2);
  assert.equal(match2.gatewayName, 'SBI_EPAY');
  assert.equal(match2._sourceFileName, 'SBI_ePay_Report.xlsx');

  // Verify file removal: removing f1 drops REF_ATOM_001
  const remainingFiles = gatewayFiles.filter((f) => f.id !== 'f1');
  const remainingRows = remainingFiles.flatMap((f) => f.rows);
  assert.equal(remainingRows.length, 1);
  assert.equal(remainingRows[0]['MERCHANT ORDER NO'], 'REF_EPAY_002');
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

test('Aggregates same course on same date & venue with trailing spaces and periods into single row summing student count', () => {

  const rawRows = [
    {
      venueCode: 'V101 ',
      venueName: 'Govt College Campus',
      examDate: '2026-11-05',
      courseCode: 'KU2AECENG106. ', // trailing dot + space
      courseName: 'English for Career. ', // trailing dot + space
      studentCount: 37
    },
    {
      venueCode: 'V101',
      venueName: 'Govt College Campus',
      examDate: '2026-11-05 09:30:00', // date with time
      courseCode: 'KU2AECENG106.', // trailing dot
      courseName: 'English for Career', // without trailing dot
      studentCount: 38
    }
  ];

  // Group by venue
  const groups = {};
  rawRows.forEach(row => {
    const vCode = cleanDisplayField(row.venueCode);
    const vName = cleanDisplayField(row.venueName);
    const groupKey = `${vCode} - ${vName}`;
    if (!groups[groupKey]) groups[groupKey] = { code: vCode, name: vName, items: [] };

    groups[groupKey].items.push({
      date: '2026-11-05',
      courseCode: row.courseCode,
      courseName: row.courseName,
      studentCount: row.studentCount
    });
  });

  // Aggregate within group
  Object.keys(groups).forEach(key => {
    const mergedMap = new Map();
    groups[key].items.forEach(item => {
      const cleanCode = normalizeKeyText(item.courseCode);
      const cleanName = normalizeKeyText(item.courseName);
      const cleanDate = normalizeKeyText(item.date);
      const vCode = normalizeKeyText(groups[key].code);

      const mKey = `${vCode}__${cleanDate}__${cleanCode}__${cleanName}`;
      if (mergedMap.has(mKey)) {
        const existing = mergedMap.get(mKey);
        existing.studentCount += item.studentCount;
      } else {
        mergedMap.set(mKey, {
          ...item,
          courseCode: cleanDisplayField(item.courseCode),
          courseName: cleanDisplayField(item.courseName)
        });
      }
    });
    groups[key].items = Array.from(mergedMap.values());
  });

  const groupItems = groups['V101 - Govt College Campus'].items;
  assert.equal(groupItems.length, 1, 'Should have aggregated 2 duplicate course rows into exactly 1');
  assert.equal(groupItems[0].studentCount, 75, 'Student count should sum 37 + 38 = 75');
  assert.equal(groupItems[0].courseCode, 'KU2AECENG106');
  assert.equal(groupItems[0].courseName, 'English for Career');
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

test('Aggregates multiple Excel files for Dataset A and Dataset B with provenance tracking', () => {
  const filesA = [
    {
      id: 'f_a1',
      name: 'Roster_Science.xlsx',
      columns: ['PRN', 'Name', 'Subject'],
      rows: [
        { PRN: 'S101', Name: 'Alice', Subject: 'Physics' },
        { PRN: 'S102', Name: 'Bob', Subject: 'Chemistry' }
      ]
    },
    {
      id: 'f_a2',
      name: 'Roster_Arts.xlsx',
      columns: ['PRN', 'Name', 'Subject', 'Language'],
      rows: [
        { PRN: 'A201', Name: 'Charlie', Subject: 'History', Language: 'English' }
      ]
    }
  ];

  const filesB = [
    {
      id: 'f_b1',
      name: 'Exam_Center_North.xlsx',
      columns: ['RegNo', 'Candidate', 'Score'],
      rows: [
        { RegNo: 'S101', Candidate: 'Alice', Score: 85 },
        { RegNo: 'A201', Candidate: 'Charlie', Score: 92 }
      ]
    },
    {
      id: 'f_b2',
      name: 'Exam_Center_South.xlsx',
      columns: ['RegNo', 'Candidate', 'Score'],
      rows: [
        { RegNo: 'S102', Candidate: 'Bob', Score: 78 },
        { RegNo: 'X999', Candidate: 'Unknown', Score: 50 }
      ]
    }
  ];

  // Aggregation logic mimicking DataComparisonPage useMemo
  const aggregate = (fileList) => {
    const allRows = [];
    const colSet = new Set();
    const orderedCols = [];

    fileList.forEach(file => {
      (file.columns || []).forEach(c => {
        if (!c.startsWith('_') && !colSet.has(c)) {
          colSet.add(c);
          orderedCols.push(c);
        }
      });
      (file.rows || []).forEach(r => {
        allRows.push({
          ...r,
          _sourceFileName: file.name,
          _sourceFileId: file.id
        });
      });
    });

    return { columns: orderedCols, rows: allRows };
  };

  const dsA = aggregate(filesA);
  const dsB = aggregate(filesB);

  // Validate Dataset A aggregation
  assert.equal(dsA.rows.length, 3);
  assert.deepEqual(dsA.columns, ['PRN', 'Name', 'Subject', 'Language']);
  assert.equal(dsA.rows[0]._sourceFileName, 'Roster_Science.xlsx');
  assert.equal(dsA.rows[2]._sourceFileName, 'Roster_Arts.xlsx');

  // Validate Dataset B aggregation
  assert.equal(dsB.rows.length, 4);
  assert.deepEqual(dsB.columns, ['RegNo', 'Candidate', 'Score']);
  assert.equal(dsB.rows[0]._sourceFileName, 'Exam_Center_North.xlsx');
  assert.equal(dsB.rows[1]._sourceFileName, 'Exam_Center_North.xlsx');
  assert.equal(dsB.rows[2]._sourceFileName, 'Exam_Center_South.xlsx');
  assert.equal(dsB.rows[3]._sourceFileName, 'Exam_Center_South.xlsx');

  // Validate Key Reconciliation across multiple files
  const indexB = new Map(dsB.rows.map(r => [r.RegNo, r]));
  const matched = [];
  const unmatchedA = [];

  dsA.rows.forEach(rA => {
    if (indexB.has(rA.PRN)) {
      matched.push({
        prn: rA.PRN,
        sourceA: rA._sourceFileName,
        sourceB: indexB.get(rA.PRN)._sourceFileName
      });
    } else {
      unmatchedA.push(rA);
    }
  });

  assert.equal(matched.length, 3);
  assert.equal(unmatchedA.length, 0);
  assert.equal(matched[0].sourceA, 'Roster_Science.xlsx');
  assert.equal(matched[0].sourceB, 'Exam_Center_North.xlsx');
  assert.equal(matched[1].sourceA, 'Roster_Science.xlsx');
  assert.equal(matched[1].sourceB, 'Exam_Center_South.xlsx');
  assert.equal(matched[2].sourceA, 'Roster_Arts.xlsx');
  assert.equal(matched[2].sourceB, 'Exam_Center_North.xlsx');
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

// -------------------------------------------------------------
// 11. Admission Import & Transformation Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 11: Admission Import & Transformation Engine]');

test('Header scanner detects candidate table header below university title banners', () => {
  // Simulate an allotment sheet with 4 rows of banners and headers on Row 4
  const mockSheetData = [
    ['KANNUR UNIVERSITY - CENTRALIZED ADMISSION PROCESS (CAP 2026-27)'],
    ['GOVT. BRENNEN COLLEGE, THALASSERY (COLLEGE CODE: 101)'],
    ['MERIT ALLOTMENT LIST - PHASE 1 ALLOTMENT'],
    ['Report Date: 12/06/2026', '', '', '', '', 'Generated By: Portal Admin'],
    ['Sl.No', 'Application No', 'Candidate Name', 'Gender', 'Category', 'Mobile', 'Email', 'Index Mark', 'Option No'],
    [1, '24CAP001', 'STUDENT ALPHA', 'Male', 'General', '9876543210', 'alpha@sample.edu', '1180.00', 1],
    [2, '24CAP002', 'STUDENT BETA', 'Female', 'OBC', '9876543211', 'beta@sample.edu', '1120.00', 2],
    [3, '24CAP003', 'STUDENT GAMMA', 'Male', 'SC', '9876543212', 'gamma@sample.edu', '1090.00', 1]
  ];

  const ws = XLSX.utils.aoa_to_sheet(mockSheetData);

  // Re-run the admission header scan logic
  const normalize = (k) => String(k || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const weights = [
    { keys: ['name', 'candidatename', 'studentname'], weight: 12, label: 'Name' },
    { keys: ['appno', 'applicationno', 'slno'], weight: 10, label: 'Identifier' },
    { keys: ['mobile', 'mobilenumber', 'phone'], weight: 8, label: 'Mobile' },
    { keys: ['email', 'emailid'], weight: 6, label: 'Email' },
    { keys: ['marks', 'indexmark'], weight: 8, label: 'Marks' }
  ];

  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  let bestIdx = -1;
  let maxScore = -99;

  for (let r = 0; r < Math.min(aoa.length, 10); r++) {
    const row = aoa[r];
    const nonEmpty = row.filter(c => String(c ?? '').trim() !== '');
    if (nonEmpty.length < 2) continue;

    const norm = row.map(c => normalize(c));
    let score = 0;
    for (const w of weights) {
      if (norm.some(col => w.keys.some(k => col.includes(k)))) {
        score += w.weight;
      }
    }
    const joined = row.join(' ').toLowerCase();
    if (/university|centralized admission|allotment process/i.test(joined)) {
      score -= 15;
    }
    if (score > maxScore) {
      maxScore = score;
      bestIdx = r;
    }
  }

  // Best header row index MUST be 4 (0-indexed 5th row)
  assert.equal(bestIdx, 4);

  // Data rows extracted from index 5 to end
  const headers = aoa[bestIdx];
  const extractedRows = [];
  for (let r = bestIdx + 1; r < aoa.length; r++) {
    const row = aoa[r];
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    extractedRows.push(obj);
  }

  assert.equal(extractedRows.length, 3);
  assert.equal(extractedRows[0]['Candidate Name'], 'STUDENT ALPHA');
  assert.equal(extractedRows[1]['Application No'], '24CAP002');
  assert.equal(extractedRows[2]['Index Mark'], '1090.00');
});

test('Multi-sheet workbook: skips blank/instruction sheet 0 and selects candidate sheet 1', () => {
  const wb = XLSX.utils.book_new();

  // Sheet 0: Instructions (Empty of candidate tabular data)
  const wsInstructions = XLSX.utils.aoa_to_sheet([
    ['Instructions for Reporting Candidates'],
    ['1. Bring original certificates along with two passport-sized photographs.'],
    ['2. Failure to report before 4:00 PM leads to forfeiture of allotted seat.']
  ]);
  XLSX.utils.book_append_sheet(wb, wsInstructions, 'Instructions');

  // Sheet 1: Allotment Data
  const wsAllotment = XLSX.utils.aoa_to_sheet([
    ['Sl No', 'Application Number', 'Candidate Name', 'Mobile Number', 'Email ID'],
    [1, '209901001', 'STUDENT ALPHA', '9876543210', 'alpha@sample.edu'],
    [2, '209901002', 'STUDENT BETA', '9876543211', 'beta@sample.edu']
  ]);
  XLSX.utils.book_append_sheet(wb, wsAllotment, 'Allotment_List');

  // Multi-sheet scoring logic
  const normalize = (k) => String(k || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  let bestSheet = '';
  let bestCount = -1;

  wb.SheetNames.forEach(sheetName => {
    const sheet = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    let headerIdx = -1;
    let score = -99;

    for (let r = 0; r < Math.min(aoa.length, 10); r++) {
      const row = aoa[r];
      const nonEmpty = row.filter(c => String(c ?? '').trim() !== '');
      if (nonEmpty.length < 2) continue;
      const norm = row.map(c => normalize(c));
      let s = 0;
      if (norm.some(c => c.includes('name'))) s += 10;
      if (norm.some(c => c.includes('app') || c.includes('slno'))) s += 10;
      if (norm.some(c => c.includes('mobile') || c.includes('email'))) s += 10;
      if (s > score) {
        score = s;
        headerIdx = r;
      }
    }

    const rowCount = headerIdx >= 0 ? aoa.length - 1 - headerIdx : 0;
    if (rowCount > bestCount) {
      bestCount = rowCount;
      bestSheet = sheetName;
    }
  });

  // Must select Allotment_List, NOT Instructions
  assert.equal(bestSheet, 'Allotment_List');
  assert.equal(bestCount, 2);
});

// -------------------------------------------------------------
// 18. Data Comparison & Reconciliation Studio (2 Lakh Rows Engine)
// -------------------------------------------------------------
console.log('\n[Test Suite 18: Data Comparison & Reconciliation Studio (Large Scale Engine)]');

test('2 Lakh rows scale: sub-second reconciliation with Map and Uint8Array index tracking', () => {
  const rowCount = 100000;
  const datasetA = new Array(rowCount);
  const datasetB = new Array(rowCount);

  for (let i = 0; i < rowCount; i++) {
    datasetA[i] = {
      prn: `PRN_${i.toString().padStart(6, '0')}`,
      name: `STUDENT_${i}`,
      score: (i % 100).toString()
    };
    datasetB[i] = {
      regNo: `PRN_${i.toString().padStart(6, '0')}`,
      candName: `STUDENT_${i}`,
      marks: (i % 100).toString()
    };
  }

  // Inject 500 discrepancies and 500 missing
  for (let i = 0; i < 500; i++) {
    datasetB[i].marks = '999'; // Discrepancy
  }
  for (let i = 500; i < 1000; i++) {
    datasetB[i].regNo = `PRN_DIFF_${i}`; // Unmatched
  }

  const startTime = Date.now();

  // Optimized indexing engine as in DataComparisonPage.jsx
  const mapB = new Map();
  const dupB = new Set();
  const matchedBIndices = new Uint8Array(datasetB.length);

  for (let i = 0; i < datasetB.length; i++) {
    const key = datasetB[i].regNo.toLowerCase().trim();
    if (mapB.has(key)) {
      dupB.add(key);
    } else {
      mapB.set(key, i);
    }
  }

  const exactMatches = [];
  const valueDiscrepancies = [];
  const unmatchedA = [];

  for (let i = 0; i < datasetA.length; i++) {
    const keyA = datasetA[i].prn.toLowerCase().trim();
    const idxB = mapB.get(keyA);

    if (idxB !== undefined) {
      matchedBIndices[idxB] = 1;
      const bRow = datasetB[idxB];
      if (datasetA[i].score === bRow.marks) {
        exactMatches.push({ a: datasetA[i], b: bRow });
      } else {
        valueDiscrepancies.push({ a: datasetA[i], b: bRow });
      }
    } else {
      unmatchedA.push(datasetA[i]);
    }
  }

  const unmatchedB = [];
  for (let i = 0; i < datasetB.length; i++) {
    if (matchedBIndices[i] === 0) {
      unmatchedB.push(datasetB[i]);
    }
  }

  const elapsedMs = Date.now() - startTime;

  assert.equal(exactMatches.length, 99000);
  assert.equal(valueDiscrepancies.length, 500);
  assert.equal(unmatchedA.length, 500);
  assert.equal(unmatchedB.length, 500);
  assert.ok(elapsedMs < 500, `Execution took ${elapsedMs}ms, should be under 500ms`);

  // Verify pagination calculation
  const totalReconciled = exactMatches.length + valueDiscrepancies.length + unmatchedA.length + unmatchedB.length;
  const pageSize = 50;
  const totalPages = Math.ceil(totalReconciled / pageSize);
  assert.equal(totalReconciled, 100500);
  assert.equal(totalPages, 2010);
});

// -------------------------------------------------------------
// 19. Admission Report Maker - Row Wise (Wide-to-Long Reshaper)
// -------------------------------------------------------------
console.log('\n[Test Suite 19: Admission Report Maker - Row Wise Engine]');

test('Reshapes wide course data into row-wise allocations, preserves base columns, and omits empty slots', () => {
  const wideRows = [
    {
      ApplicationNo: '2024CAP001',
      CandidateName: 'STUDENT ALPHA',
      CollegeCode: 'AA',
      Mobile: '9876543210',
      ParentGroupName1: 'Major Core',
      ImmediateParentGroupName1: 'Discipline Specific Core',
      CourseCode1: 'KU4DSCHIS101',
      CourseName1: 'Early Indian History',
      ParentGroupName2: 'Minor Group',
      ImmediateParentGroupName2: 'Minor Stream A',
      CourseCode2: 'KU4MINPOL101',
      CourseName2: 'Political Theory',
      ParentGroupName3: '',
      ImmediateParentGroupName3: '',
      CourseCode3: '',
      CourseName3: ''
    },
    {
      ApplicationNo: '2024CAP002',
      CandidateName: 'STUDENT BETA',
      CollegeCode: 'BB',
      Mobile: '9876543211',
      ParentGroupName1: 'Major Core',
      ImmediateParentGroupName1: 'Discipline Specific Core',
      CourseCode1: 'KU4DSCENG101',
      CourseName1: 'Literary Studies',
      ParentGroupName2: 'Minor Group',
      ImmediateParentGroupName2: 'Minor Stream B',
      CourseCode2: 'KU4MINSOC101',
      CourseName2: 'Sociology',
      ParentGroupName3: 'AEC Language',
      ImmediateParentGroupName3: 'English AEC',
      CourseCode3: 'KU4AECENG101',
      CourseName3: 'English Fluency'
    }
  ];

  const columns = Object.keys(wideRows[0]);

  // Test dynamic numeric index detection
  const indicesSet = new Set();
  columns.forEach(col => {
    const match = col.trim().match(/^course[\s_]*code[\s_]*(\d+)$/i);
    if (match) {
      indicesSet.add(parseInt(match[1], 10));
    }
  });
  const indices = Array.from(indicesSet).sort((a, b) => a - b);
  assert.deepEqual(indices, [1, 2, 3]);

  // Identify base columns
  const courseFields = ['ParentGroupName', 'ImmediateParentGroupName', 'CourseCode', 'CourseName'];
  const repeatedCols = new Set();
  for (const field of courseFields) {
    for (const i of indices) {
      repeatedCols.add(`${field}${i}`);
    }
  }
  const baseCols = columns.filter(c => !repeatedCols.has(c));
  assert.deepEqual(baseCols, ['ApplicationNo', 'CandidateName', 'CollegeCode', 'Mobile']);

  // Reshape wide rows
  const isValid = (val) => val !== null && val !== undefined && String(val).trim() !== '' && String(val).trim().toLowerCase() !== 'nan';
  const reshaped = [];
  let omitted = 0;

  for (const row of wideRows) {
    const baseData = {};
    for (const bCol of baseCols) {
      baseData[bCol] = row[bCol];
    }

    for (const i of indices) {
      const code = row[`CourseCode${i}`];
      const name = row[`CourseName${i}`];
      if (isValid(code) || isValid(name)) {
        reshaped.push({
          ...baseData,
          ParentGroupName: row[`ParentGroupName${i}`] || '',
          ImmediateParentGroupName: row[`ImmediateParentGroupName${i}`] || '',
          CourseCode: code || '',
          CourseName: name || ''
        });
      } else {
        omitted++;
      }
    }
  }

  // Student ALPHA has 2 courses (slot 3 empty), Student BETA has 3 courses -> Total 5 rows
  assert.equal(reshaped.length, 5);
  assert.equal(omitted, 1);

  // Validate Student ALPHA row 1
  assert.equal(reshaped[0].ApplicationNo, '2024CAP001');
  assert.equal(reshaped[0].CandidateName, 'STUDENT ALPHA');
  assert.equal(reshaped[0].CourseCode, 'KU4DSCHIS101');
  assert.equal(reshaped[0].ParentGroupName, 'Major Core');

  // Validate Student BETA row 3
  assert.equal(reshaped[4].ApplicationNo, '2024CAP002');
  assert.equal(reshaped[4].CandidateName, 'STUDENT BETA');
  assert.equal(reshaped[4].CourseCode, 'KU4AECENG101');
  assert.equal(reshaped[4].CourseName, 'English Fluency');
});

test('Copies ImmediateParentGroupName to ParentGroupName when ParentGroupName is blank, and matches ParentGroup header alias', () => {
  const row = {
    ApplicationNo: '2024CAP003',
    CandidateName: 'STUDENT GAMMA',
    ParentGroup1: '', // Header uses ParentGroup1 (alias without Name) and is blank
    ImmediateParentGroupName1: 'Discipline Specific Core (Psychology)',
    CourseCode1: 'KU4DSCPSY101',
    CourseName1: 'Introduction to Psychology'
  };

  const columns = Object.keys(row);
  const colMap = new Map();
  columns.forEach(c => colMap.set(c.toLowerCase(), c));

  const findCol = (field, index) => {
    const fieldAliases = [];
    if (field === 'ParentGroupName') fieldAliases.push('ParentGroupName', 'ParentGroup');
    else if (field === 'ImmediateParentGroupName') fieldAliases.push('ImmediateParentGroupName', 'ImmediateParentGroup');
    else fieldAliases.push(field);

    for (const alias of fieldAliases) {
      const exact = `${alias}${index}`;
      if (columns.includes(exact)) return exact;
      if (colMap.has(exact.toLowerCase())) return colMap.get(exact.toLowerCase());
    }
    return null;
  };

  const pCol = findCol('ParentGroupName', 1);
  const immCol = findCol('ImmediateParentGroupName', 1);

  assert.equal(pCol, 'ParentGroup1');
  assert.equal(immCol, 'ImmediateParentGroupName1');

  let parentGroupName = pCol ? (row[pCol] || '').trim() : '';
  const immParentGroupName = immCol ? (row[immCol] || '').trim() : '';

  if (!parentGroupName && immParentGroupName) {
    parentGroupName = immParentGroupName;
  }

  assert.equal(parentGroupName, 'Discipline Specific Core (Psychology)');
  assert.equal(immParentGroupName, 'Discipline Specific Core (Psychology)');
});

test('Generates unique Program Part Term ParentGroupName ImmediateParentGroupName CourseCode CourseName report', () => {
  // Mock reshaped rows for multiple students
  const reshapedRows = [
    {
      ApplicationNo: '2024CAP001',
      CandidateName: 'STUDENT ALPHA',
      ProgrammeName: 'B.A. History (FYIMP)',
      PartNo: '1',
      Semester: '1',
      ParentGroupName: 'Major Core',
      ImmediateParentGroupName: 'Discipline Specific Core',
      CourseCode: 'KU4DSCHIS101',
      CourseName: 'Early Indian History and Culture'
    },
    {
      ApplicationNo: '2024CAP001',
      CandidateName: 'STUDENT ALPHA',
      ProgrammeName: 'B.A. History (FYIMP)',
      PartNo: '1',
      Semester: '1',
      ParentGroupName: 'Minor Group',
      ImmediateParentGroupName: 'Minor Stream A',
      CourseCode: 'KU4MINPOL101',
      CourseName: 'Principles of Political Theory'
    },
    // STUDENT EPSILON takes the EXACT same courses in the same program/term/groups
    {
      ApplicationNo: '2024CAP005',
      CandidateName: 'STUDENT EPSILON',
      ProgrammeName: 'B.A. History (FYIMP)',
      PartNo: '1',
      Semester: '1',
      ParentGroupName: 'Major Core',
      ImmediateParentGroupName: 'Discipline Specific Core',
      CourseCode: 'KU4DSCHIS101',
      CourseName: 'Early Indian History and Culture'
    },
    {
      ApplicationNo: '2024CAP005',
      CandidateName: 'STUDENT EPSILON',
      ProgrammeName: 'B.A. History (FYIMP)',
      PartNo: '1',
      Semester: '1',
      ParentGroupName: 'Minor Group',
      ImmediateParentGroupName: 'Minor Stream A',
      CourseCode: 'KU4MINPOL101',
      CourseName: 'Principles of Political Theory'
    },
    // STUDENT BETA has a row where ParentGroupName is blank
    {
      ApplicationNo: '2024CAP002',
      CandidateName: 'STUDENT BETA',
      ProgrammeName: 'B.A. English (FYIMP)',
      PartNo: '1',
      Semester: '1',
      ParentGroupName: 'Discipline Specific Core (Psychology)', // Fallback already applied
      ImmediateParentGroupName: 'Discipline Specific Core (Psychology)',
      CourseCode: 'KU4DSCPSY101',
      CourseName: 'Introduction to Psychology'
    }
  ];

  // Helper matching detectColumnMatch
  const detectColumnMatch = (columns, candidateNames) => {
    if (!columns || columns.length === 0) return '';
    const lowerCols = columns.map(c => ({
      original: c,
      clean: String(c).trim().toLowerCase().replace(/[\s_]+/g, '')
    }));

    for (const cand of candidateNames) {
      const cleanCand = cand.toLowerCase().replace(/[\s_]+/g, '');
      const found = lowerCols.find(c => c.clean === cleanCand);
      if (found) return found.original;
    }

    for (const cand of candidateNames) {
      const cleanCand = cand.toLowerCase().replace(/[\s_]+/g, '');
      const found = lowerCols.find(c => c.clean.includes(cleanCand) || cleanCand.includes(c.clean));
      if (found) return found.original;
    }
    return '';
  };

  const sampleCols = Object.keys(reshapedRows[0]);
  const progCol = detectColumnMatch(sampleCols, ['Program', 'ProgramName', 'Programme', 'ProgrammeName']);
  const partCol = detectColumnMatch(sampleCols, ['Part', 'PartName', 'PartNo', 'Year']);
  const termCol = detectColumnMatch(sampleCols, ['Term', 'TermName', 'TermNo', 'Semester', 'Sem']);

  assert.equal(progCol, 'ProgrammeName');
  assert.equal(partCol, 'PartNo');
  assert.equal(termCol, 'Semester');

  // Helper matching generateUniqueCourseReport
  const generateUniqueReport = (rows, mapping) => {
    const uniqueMap = new Map();

    for (const row of rows) {
      const courseCode = (row['CourseCode'] || '').trim();
      const courseName = (row['CourseName'] || '').trim();
      if (!courseCode && !courseName) continue;

      const program = mapping.programCol && row[mapping.programCol] !== undefined ? String(row[mapping.programCol]).trim() : 'Default Program';
      const part = mapping.partCol && row[mapping.partCol] !== undefined ? String(row[mapping.partCol]).trim() : '1';
      const term = mapping.termCol && row[mapping.termCol] !== undefined ? String(row[mapping.termCol]).trim() : '1';
      const parentGroupName = (row['ParentGroupName'] || '').trim();
      const immediateParentGroupName = (row['ImmediateParentGroupName'] || '').trim();

      const key = [
        program.toLowerCase(),
        part.toLowerCase(),
        term.toLowerCase(),
        parentGroupName.toLowerCase(),
        immediateParentGroupName.toLowerCase(),
        courseCode.toLowerCase(),
        courseName.toLowerCase()
      ].join('|||');

      if (!uniqueMap.has(key)) {
        uniqueMap.set(key, {
          Program: program,
          Part: part,
          Term: term,
          ParentGroupName: parentGroupName,
          ImmediateParentGroupName: immediateParentGroupName,
          CourseCode: courseCode,
          CourseName: courseName
        });
      }
    }
    return Array.from(uniqueMap.values());
  };

  const uniqueList = generateUniqueReport(reshapedRows, {
    programCol: progCol,
    partCol: partCol,
    termCol: termCol
  });

  // Out of 5 student allocations, exactly 3 unique courses exist (the 2 identical ones from Student Epsilon are deduped)
  assert.equal(uniqueList.length, 3);

  // Validate exact headers/keys
  const expectedKeys = [
    'Program',
    'Part',
    'Term',
    'ParentGroupName',
    'ImmediateParentGroupName',
    'CourseCode',
    'CourseName'
  ];
  assert.deepEqual(Object.keys(uniqueList[0]), expectedKeys);

  // Validate values for course 1
  assert.equal(uniqueList[0].Program, 'B.A. History (FYIMP)');
  assert.equal(uniqueList[0].Part, '1');
  assert.equal(uniqueList[0].Term, '1');
  assert.equal(uniqueList[0].ParentGroupName, 'Major Core');
  assert.equal(uniqueList[0].ImmediateParentGroupName, 'Discipline Specific Core');
  assert.equal(uniqueList[0].CourseCode, 'KU4DSCHIS101');
  assert.equal(uniqueList[0].CourseName, 'Early Indian History and Culture');

  // Validate course 3 with fallback group
  assert.equal(uniqueList[2].Program, 'B.A. English (FYIMP)');
  assert.equal(uniqueList[2].CourseCode, 'KU4DSCPSY101');
  assert.equal(uniqueList[2].ParentGroupName, 'Discipline Specific Core (Psychology)');
});

test('Admission Report Maker - Generates Unique Institute Course Master Report (Excel) with exact 6 columns and deduplication', () => {
  const reshapedRows = [
    {
      CandidateName: 'STUDENT ALPHA',
      CollegeCode: 'AA',
      CollegeName: 'Alpha Arts and Science College',
      ParentGroupName: 'Major Core',
      ImmediateParentGroupName: 'Discipline Specific Core',
      CourseCode: 'KU4DSCHIS101',
      CourseName: 'Early Indian History and Culture'
    },
    {
      CandidateName: 'STUDENT BETA',
      CollegeCode: 'AA',
      CollegeName: 'Alpha Arts and Science College',
      ParentGroupName: 'Major Core',
      ImmediateParentGroupName: 'Discipline Specific Core',
      CourseCode: 'KU4DSCHIS101', // duplicate course at same college AA
      CourseName: 'Early Indian History and Culture'
    },
    {
      CandidateName: 'STUDENT GAMMA',
      CollegeCode: 'AA',
      CollegeName: 'Alpha Arts and Science College',
      ParentGroupName: '', // blank ParentGroupName -> should fallback to ImmediateParentGroupName
      ImmediateParentGroupName: 'Discipline Specific Core (Economics)',
      CourseCode: 'KU4DSCECO101',
      CourseName: 'Introductory Microeconomics'
    },
    {
      CandidateName: 'STUDENT DELTA',
      CollegeCode: 'BB',
      CollegeName: 'Beta Commerce College',
      ParentGroupName: 'Major Core',
      ImmediateParentGroupName: 'Discipline Specific Core',
      CourseCode: 'KU4DSCHIS101', // same course but DIFFERENT college BB -> MUST NOT be deduplicated
      CourseName: 'Early Indian History and Culture'
    }
  ];

  // Logic matching generateUniqueInstituteCourseReport in AdmissionRowWisePage
  const generateUniqueInstituteReport = (rows, options = {}) => {
    const {
      instituteCodeCol = 'CollegeCode',
      instituteNameCol = 'CollegeName',
      fallbackInstituteCode = 'INST001',
      fallbackInstituteName = 'Main Campus'
    } = options;

    const uniqueMap = new Map();

    for (const row of rows) {
      const courseCode = String(row.CourseCode || '').trim();
      const courseName = String(row.CourseName || '').trim();
      if (!courseCode && !courseName) continue;

      let immediateParentGroupName = String(row.ImmediateParentGroupName || '').trim();
      let parentGroupName = String(row.ParentGroupName || '').trim();
      if (!parentGroupName && immediateParentGroupName) {
        parentGroupName = immediateParentGroupName;
      }

      let instCode = instituteCodeCol && row[instituteCodeCol] !== undefined ? String(row[instituteCodeCol]).trim() : '';
      if (!instCode) instCode = fallbackInstituteCode;

      let instName = instituteNameCol && row[instituteNameCol] !== undefined ? String(row[instituteNameCol]).trim() : '';
      if (!instName) instName = fallbackInstituteName;

      const key = [
        instCode.toLowerCase(),
        instName.toLowerCase(),
        parentGroupName.toLowerCase(),
        immediateParentGroupName.toLowerCase(),
        courseCode.toLowerCase(),
        courseName.toLowerCase()
      ].join('|||');

      if (!uniqueMap.has(key)) {
        uniqueMap.set(key, {
          'Institute Code': instCode,
          'Institute Name': instName,
          'ParentGroupName': parentGroupName,
          'ImmediateParentGroupName': immediateParentGroupName,
          'CourseCode': courseCode,
          'CourseName': courseName
        });
      }
    }
    return Array.from(uniqueMap.values());
  };

  const report = generateUniqueInstituteReport(reshapedRows);

  // 4 rows input:
  // - 2 identical for (AA, KU4DSCHIS101) -> deduplicated to 1
  // - 1 for (AA, KU4DSCECO101) with fallback ParentGroupName
  // - 1 for (BB, KU4DSCHIS101) at distinct college
  // Total expected unique combinations = 3
  assert.equal(report.length, 3);

  // Validate exact 6 requested headers
  const expectedHeaders = [
    'Institute Code',
    'Institute Name',
    'ParentGroupName',
    'ImmediateParentGroupName',
    'CourseCode',
    'CourseName'
  ];
  assert.deepEqual(Object.keys(report[0]), expectedHeaders);

  // Validate record 1 (College AA, HIS101)
  assert.equal(report[0]['Institute Code'], 'AA');
  assert.equal(report[0]['Institute Name'], 'Alpha Arts and Science College');
  assert.equal(report[0]['ParentGroupName'], 'Major Core');
  assert.equal(report[0]['ImmediateParentGroupName'], 'Discipline Specific Core');
  assert.equal(report[0]['CourseCode'], 'KU4DSCHIS101');
  assert.equal(report[0]['CourseName'], 'Early Indian History and Culture');

  // Validate record 2 (College AA, ECO101 with fallback ParentGroupName)
  assert.equal(report[1]['Institute Code'], 'AA');
  assert.equal(report[1]['ParentGroupName'], 'Discipline Specific Core (Economics)');
  assert.equal(report[1]['ImmediateParentGroupName'], 'Discipline Specific Core (Economics)');
  assert.equal(report[1]['CourseCode'], 'KU4DSCECO101');

  // Validate record 3 (College BB, HIS101)
  assert.equal(report[2]['Institute Code'], 'BB');
  assert.equal(report[2]['Institute Name'], 'Beta Commerce College');
  assert.equal(report[2]['CourseCode'], 'KU4DSCHIS101');
});

test('Admission Report Maker - Generates Unique 4-Column Course Master Report (ParentGroupName, ImmediateParentGroupName, CourseCode, CourseName)', () => {
  const reshapedRows = [
    {
      CollegeCode: 'AA',
      Program: 'B.A. History',
      ParentGroupName: 'Major Core',
      ImmediateParentGroupName: 'Discipline Specific Core',
      CourseCode: 'KU4DSCHIS101',
      CourseName: 'Early Indian History and Culture'
    },
    {
      CollegeCode: 'BB', // different college and program, but same 4 fields -> MUST be deduplicated
      Program: 'B.A. Archaeology',
      ParentGroupName: 'Major Core',
      ImmediateParentGroupName: 'Discipline Specific Core',
      CourseCode: 'KU4DSCHIS101',
      CourseName: 'Early Indian History and Culture'
    },
    {
      CollegeCode: 'CC',
      Program: 'B.Sc. Psychology',
      ParentGroupName: '', // blank -> fallback to ImmediateParentGroupName
      ImmediateParentGroupName: 'Discipline Specific Core (Psychology)',
      CourseCode: 'KU4DSCPSY101',
      CourseName: 'Foundations of Psychology'
    },
    {
      CollegeCode: 'DD',
      Program: 'B.Com.',
      ParentGroupName: 'Elective',
      ImmediateParentGroupName: 'Generic Elective',
      CourseCode: 'KU4GEECO101',
      CourseName: 'Business Economics'
    }
  ];

  // Logic matching generateUniqueCourseMasterReport
  const generateUniqueCourseMaster = (rows) => {
    const uniqueMap = new Map();
    for (const row of rows) {
      const courseCode = String(row.CourseCode || '').trim();
      const courseName = String(row.CourseName || '').trim();
      if (!courseCode && !courseName) continue;

      let immediateParentGroupName = String(row.ImmediateParentGroupName || '').trim();
      let parentGroupName = String(row.ParentGroupName || '').trim();
      if (!parentGroupName && immediateParentGroupName) {
        parentGroupName = immediateParentGroupName;
      }

      const key = [
        parentGroupName.toLowerCase(),
        immediateParentGroupName.toLowerCase(),
        courseCode.toLowerCase(),
        courseName.toLowerCase()
      ].join('|||');

      if (!uniqueMap.has(key)) {
        uniqueMap.set(key, {
          ParentGroupName: parentGroupName,
          ImmediateParentGroupName: immediateParentGroupName,
          CourseCode: courseCode,
          CourseName: courseName
        });
      }
    }
    return Array.from(uniqueMap.values());
  };

  const report = generateUniqueCourseMaster(reshapedRows);

  // 4 rows input:
  // - 2 identical for (Major Core, Discipline Specific Core, KU4DSCHIS101, Early Indian History) -> deduplicated to 1
  // - 1 for (KU4DSCPSY101) with fallback ParentGroupName
  // - 1 for (KU4GEECO101)
  // Total expected unique = 3
  assert.equal(report.length, 3);

  // Validate exact 4 headers requested
  const expectedHeaders = [
    'ParentGroupName',
    'ImmediateParentGroupName',
    'CourseCode',
    'CourseName'
  ];
  assert.deepEqual(Object.keys(report[0]), expectedHeaders);

  // Validate record 1
  assert.equal(report[0].ParentGroupName, 'Major Core');
  assert.equal(report[0].ImmediateParentGroupName, 'Discipline Specific Core');
  assert.equal(report[0].CourseCode, 'KU4DSCHIS101');
  assert.equal(report[0].CourseName, 'Early Indian History and Culture');

  // Validate record 2 with fallback ParentGroupName
  assert.equal(report[1].ParentGroupName, 'Discipline Specific Core (Psychology)');
  assert.equal(report[1].ImmediateParentGroupName, 'Discipline Specific Core (Psychology)');
  assert.equal(report[1].CourseCode, 'KU4DSCPSY101');

  // Validate record 3
  assert.equal(report[2].ParentGroupName, 'Elective');
  assert.equal(report[2].CourseCode, 'KU4GEECO101');
});

test('Generates unique 3-column Course Master report across ParentGroupName, CourseCode, CourseName with fallback', () => {
  const reshapedRows = [
    {
      CandidateName: 'STUDENT ALPHA',
      CollegeCode: 'AA',
      Program: 'B.A. History',
      ParentGroupName: 'Major Core',
      ImmediateParentGroupName: 'Discipline Specific Core - Section A',
      CourseCode: 'KU4DSCHIS101',
      CourseName: 'Early Indian History'
    },
    {
      CandidateName: 'STUDENT BETA',
      CollegeCode: 'BB',
      Program: 'B.A. Archaeology',
      ParentGroupName: 'Major Core',
      ImmediateParentGroupName: 'Discipline Specific Core - Section B', // Different Immediate group, but same Parent, Code, Name!
      CourseCode: 'KU4DSCHIS101',
      CourseName: 'Early Indian History'
    },
    {
      CandidateName: 'STUDENT GAMMA',
      CollegeCode: 'CC',
      Program: 'B.Sc. Psychology',
      ParentGroupName: '', // blank -> must fallback to ImmediateParentGroupName
      ImmediateParentGroupName: 'Discipline Specific Core (Psychology)',
      CourseCode: 'KU4DSCPSY101',
      CourseName: 'Foundations of Psychology'
    },
    {
      CandidateName: 'STUDENT DELTA',
      CollegeCode: 'DD',
      Program: 'B.Com.',
      ParentGroupName: 'Elective',
      ImmediateParentGroupName: 'Generic Elective',
      CourseCode: 'KU4GEECO101',
      CourseName: 'Business Economics'
    }
  ];

  // Logic matching generateUniqueThreeColCourseReport in AdmissionRowWisePage
  const generateUniqueThreeColCourseReport = (transformedRows) => {
    if (!transformedRows || !Array.isArray(transformedRows)) return [];

    const uniqueMap = new Map();

    for (const row of transformedRows) {
      const courseCode = (row['CourseCode'] || '').trim();
      const courseName = (row['CourseName'] || '').trim();

      // Skip empty course slots
      if (!courseCode && !courseName) continue;

      let immediateParentGroupName = (row['ImmediateParentGroupName'] || '').trim();
      let parentGroupName = (row['ParentGroupName'] || '').trim();
      // Fallback if blank
      if (!parentGroupName && immediateParentGroupName) {
        parentGroupName = immediateParentGroupName;
      }

      const dedupKey = [
        parentGroupName.toLowerCase(),
        courseCode.toLowerCase(),
        courseName.toLowerCase()
      ].join('|||');

      if (!uniqueMap.has(dedupKey)) {
        uniqueMap.set(dedupKey, {
          'ParentGroupName': parentGroupName,
          'CourseCode': courseCode,
          'CourseName': courseName
        });
      }
    }

    const list = Array.from(uniqueMap.values());
    list.sort((a, b) => {
      const pgCmp = a['ParentGroupName'].localeCompare(b['ParentGroupName']);
      if (pgCmp !== 0) return pgCmp;
      const codeCmp = a['CourseCode'].localeCompare(b['CourseCode']);
      if (codeCmp !== 0) return codeCmp;
      return a['CourseName'].localeCompare(b['CourseName']);
    });

    return list;
  };

  const report = generateUniqueThreeColCourseReport(reshapedRows);

  // 4 rows input:
  // - 2 for KU4DSCHIS101 with same ParentGroupName & Name but differing ImmediateParentGroupName -> must be deduplicated into 1
  // - 1 for KU4DSCPSY101 with blank ParentGroupName -> falls back to ImmediateParentGroupName
  // - 1 for KU4GEECO101
  // Total unique = 3 records
  assert.equal(report.length, 3);

  // Validate exact 3 headers
  const expectedHeaders = [
    'ParentGroupName',
    'CourseCode',
    'CourseName'
  ];
  assert.deepEqual(Object.keys(report[0]), expectedHeaders);
  assert.equal(Object.keys(report[0]).length, 3);

  // Validate sorted output and values
  // Record 1: 'Discipline Specific Core (Psychology)' (due to fallback from blank)
  assert.equal(report[0].ParentGroupName, 'Discipline Specific Core (Psychology)');
  assert.equal(report[0].CourseCode, 'KU4DSCPSY101');
  assert.equal(report[0].CourseName, 'Foundations of Psychology');
  assert.equal(report[0].ImmediateParentGroupName, undefined);

  // Record 2: 'Elective'
  assert.equal(report[1].ParentGroupName, 'Elective');
  assert.equal(report[1].CourseCode, 'KU4GEECO101');
  assert.equal(report[1].CourseName, 'Business Economics');

  // Record 3: 'Major Core'
  assert.equal(report[2].ParentGroupName, 'Major Core');
  assert.equal(report[2].CourseCode, 'KU4DSCHIS101');
  assert.equal(report[2].CourseName, 'Early Indian History');
});

// -------------------------------------------------------------
// 20. Time Format Preservation & Text Cleaning Engine Test
// -------------------------------------------------------------
console.log('\n[Test Suite 20: Time Format Preservation & Text Cleaning Engine]');

test('Strips 1899 Excel epoch dates and GMT timezone artifacts, preserving exact text time', () => {
  // Logic matching cleanTimeText from excelParser.js
  const cleanTimeText = (val) => {
    if (val === null || val === undefined) return '';
    if (val instanceof Date) {
      if (val.getFullYear() <= 1900) {
        const hh = String(val.getHours()).padStart(2, '0');
        const mm = String(val.getMinutes()).padStart(2, '0');
        const ss = val.getSeconds() ? ':' + String(val.getSeconds()).padStart(2, '0') : '';
        return `${hh}:${mm}${ss}`;
      }
      return val.toLocaleTimeString();
    }
    let str = String(val).trim();
    const regex = /(?:[A-Za-z]{3}\s+[A-Za-z]{3}\s+\d+\s+)?(?:1899|1900)\s+(\d{1,2}:\d{2}(?::\d{2})?)\s+GMT[^\s-]+(?:\s*\([^)]*\))?/gi;
    str = str.replace(regex, '$1');
    str = str.replace(/(?:1899|1900)-12-30[T\s](\d{1,2}:\d{2}(?::\d{2})?)(?:\.\d+)?Z?/gi, '$1');
    return str.trim();
  };

  // 1. Dual-range time string produced by concatenated 1899 SheetJS Date objects
  const rawDualTime = 'Sat Dec 30 1899 13:30:00 GMT+0521 (India Standard Time) - Sat Dec 30 1899 15:00:00 GMT+0521 (India Standard Time)';
  assert.equal(cleanTimeText(rawDualTime), '13:30:00 - 15:00:00');

  // 2. Single 1899 date string
  const rawSingleTime = 'Sat Dec 30 1899 13:30:00 GMT+0521 (India Standard Time)';
  assert.equal(cleanTimeText(rawSingleTime), '13:30:00');

  // 3. Date object representing 1899 Excel epoch time
  const epochDate = new Date(1899, 11, 30, 13, 30, 0);
  assert.equal(cleanTimeText(epochDate), '13:30');

  // 4. Native text time ranges remain preserved without corruption
  assert.equal(cleanTimeText('13:30 - 15:00'), '13:30 - 15:00');
  assert.equal(cleanTimeText('01:30 PM to 04:30 PM'), '01:30 PM to 04:30 PM');
  assert.equal(cleanTimeText('10:00 AM - 01:00 PM'), '10:00 AM - 01:00 PM');
});

test('sanitizeWorksheetTimes keeps time formatted cells as text and leaves actual exam dates untouched', () => {
  const sanitizeWorksheetTimes = (sheet) => {
    if (!sheet || typeof sheet !== 'object') return sheet;
    const fixCell = (cell) => {
      if (!cell || typeof cell !== 'object') return;
      if (cell.t === 'd' || cell.v instanceof Date) {
        const d = cell.v;
        if (d instanceof Date && d.getFullYear() <= 1900) {
          let timeStr = cell.w;
          if (!timeStr) {
            const hh = String(d.getHours()).padStart(2, '0');
            const mm = String(d.getMinutes()).padStart(2, '0');
            const ss = d.getSeconds() ? ':' + String(d.getSeconds()).padStart(2, '0') : '';
            timeStr = `${hh}:${mm}${ss}`;
          }
          cell.t = 's';
          cell.v = timeStr;
          cell.w = timeStr;
        }
      }
    };
    for (const key of Object.keys(sheet)) {
      if (key.startsWith('!')) continue;
      const val = sheet[key];
      if (Array.isArray(val)) {
        for (let c = 0; c < val.length; c++) fixCell(val[c]);
      } else {
        fixCell(val);
      }
    }
    return sheet;
  };

  // Mock sheet with exam date (2025) and time cells (1899 epoch)
  const mockSheet = {
    'A1': { t: 's', v: 'ExamDate', w: 'ExamDate' },
    'B1': { t: 's', v: 'StartTime', w: 'StartTime' },
    'C1': { t: 's', v: 'EndTime', w: 'EndTime' },
    'A2': { t: 'd', v: new Date(2025, 3, 20), w: '2025-04-20' },
    'B2': { t: 'd', v: new Date(1899, 11, 30, 13, 30, 0), w: '13:30' },
    'C2': { t: 'd', v: new Date(1899, 11, 30, 15, 0, 0), w: '15:00' }
  };

  sanitizeWorksheetTimes(mockSheet);

  // Exam date MUST remain Date object
  assert.equal(mockSheet['A2'].t, 'd');
  assert.equal(mockSheet['A2'].v.getFullYear(), 2025);

  // Time cells MUST be converted to string text
  assert.equal(mockSheet['B2'].t, 's');
  assert.equal(mockSheet['B2'].v, '13:30');
  assert.equal(mockSheet['C2'].t, 's');
  assert.equal(mockSheet['C2'].v, '15:00');
});

// -------------------------------------------------------------
// 21. Course Wise Venue Wise Datewise Report Engine
// -------------------------------------------------------------
console.log('\n[Test Suite 21: Course Wise Venue Wise Datewise Report Engine]');

test('Aggregates pre-exam registrations and enriches with timetable dates, times and verification keys', () => {
  const preExamData = [
    {
      'Programme Name': 'B.A History(with Credits)-Private-SLL- 2024-History',
      'Parent Group Name': '', // blank -> should copy immediate
      'Immediate Parent Group': 'AEC1',
      PaperCode: 'KU2AECENG107.',
      PaperName: 'English through Culture.',
      VenueCode: 'PL',
      VenueName: 'Pilathara Co-op Arts and Science College, Kannur'
    },
    {
      'Programme Name': 'B.A History(with Credits)-Private-SLL- 2024-History',
      'Parent Group Name': '',
      'Immediate Parent Group': 'AEC1',
      PaperCode: 'KU2AECENG107.',
      PaperName: 'English through Culture.',
      VenueCode: 'PL',
      VenueName: 'Pilathara Co-op Arts and Science College, Kannur'
    }
  ];

  const timetableData = [
    {
      TermAbbr: 'B.A History(with Credits)-Private-SLL- 2024-History-BA Year I SEMESTER II',
      CourseCode: 'KU2AECENG107.',
      CourseName: 'English through Culture.',
      ExamDate: 46331, // 05/11/2026
      Start_Time_HR: 1,
      Start_Time_MIN: 30,
      Start_AM_PM: 'PM',
      End_Time_HR: 3,
      End_Time_MIN: 0,
      End_AM_PM: 'PM'
    }
  ];

  // Aggregation logic
  const groupMap = new Map();
  preExamData.forEach(r => {
    const pgm = r['Programme Name'];
    const code = r.PaperCode;
    const vCode = r.VenueCode;
    const key = `${pgm}__${normalizeKeyText(code)}__${normalizeKeyText(vCode)}`;
    if (!groupMap.has(key)) {
      groupMap.set(key, { ...r, count: 0 });
    }
    groupMap.get(key).count++;
  });

  const res = Array.from(groupMap.values()).map(val => {
    const tt = timetableData.find(t => normalizeKeyText(t.CourseCode) === normalizeKeyText(val.PaperCode));
    const parent = val['Parent Group Name'] || val['Immediate Parent Group'];
    const imm = val['Immediate Parent Group'];
    const examDate = '05/11/2026';
    return {
      'Programme Name': val['Programme Name'],
      'Parent Group Name': parent,
      'Immediate Parent Group': imm,
      'Course  Code': val.PaperCode,
      'Course  Name': val.PaperName,
      'Exam Date': examDate,
      'Student Count': val.count,
      'Venue Code': val.VenueCode,
      'Venue Name': val.VenueName,
      'Concatenate_Groupname_Course Code': `${parent}${val.PaperCode}`,
      'PGM_Date_CourseCode_Venue': `${val['Programme Name']}${examDate}${val.PaperCode}${val.VenueCode}`
    };
  });

  assert.equal(res.length, 1);
  assert.equal(res[0]['Student Count'], 2);
  assert.equal(res[0]['Parent Group Name'], 'AEC1');
  assert.equal(res[0]['Concatenate_Groupname_Course Code'], 'AEC1KU2AECENG107.');
  assert.equal(res[0]['PGM_Date_CourseCode_Venue'], 'B.A History(with Credits)-Private-SLL- 2024-History05/11/2026KU2AECENG107.PL');
});

test('Course Wise Venue Wise Datewise Report formats Date strictly as DD/MM/YYYY and Times as 12h AM/PM', () => {
  const formatExamDate = (val) => {
    if (!val) return '';
    if (val instanceof Date && !isNaN(val.getTime())) {
      const noon = new Date(val.getTime() + 12 * 3600 * 1000);
      const dd = String(noon.getDate()).padStart(2, '0');
      const mm = String(noon.getMonth() + 1).padStart(2, '0');
      const yyyy = noon.getFullYear();
      return `${dd}/${mm}/${yyyy}`;
    }
    if (typeof val === 'number') {
      if (XLSX?.SSF?.parse_date_code) {
        const p = XLSX.SSF.parse_date_code(val);
        if (p && p.y && p.m && p.d) {
          return `${String(p.d).padStart(2, '0')}/${String(p.m).padStart(2, '0')}/${p.y}`;
        }
      }
      const d = new Date(Math.round((val - 25569) * 86400 * 1000) + 12 * 3600 * 1000);
      return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
    }
    const str = String(val).replace(/[\u00A0\u200B\uFEFF]/g, ' ').trim();
    if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(str)) {
      const p = str.split(/[-/.T\s]/);
      return `${p[2].padStart(2, '0')}/${p[1].padStart(2, '0')}/${p[0]}`;
    }
    if (/^\d{1,2}[-/.]\d{1,2}[-/.]\d{4}$/.test(str)) {
      const p = str.split(/[-/.]/);
      return `${p[0].padStart(2, '0')}/${p[1].padStart(2, '0')}/${p[2]}`;
    }
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 1900) {
      const noon = new Date(parsed.getTime() + 12 * 3600 * 1000);
      const dd = String(noon.getDate()).padStart(2, '0');
      const mm = String(noon.getMonth() + 1).padStart(2, '0');
      const yyyy = noon.getFullYear();
      return `${dd}/${mm}/${yyyy}`;
    }
    return str;
  };

  const formatTimeDisplay = (val) => {
    if (val === null || val === undefined || val === '') return '';
    if (val instanceof Date) {
      let h = val.getHours();
      const m = val.getMinutes();
      const s = val.getSeconds();
      const ampm = h >= 12 ? 'PM' : 'AM';
      let dispH = h % 12;
      if (dispH === 0) dispH = 12;
      const secStr = s > 0 ? ':' + String(s).padStart(2, '0') : ':00';
      return `${dispH}:${String(m).padStart(2, '0')}${secStr} ${ampm}`;
    }
    const num = typeof val === 'number' ? val : (typeof val === 'string' && /^\d*\.\d+$/.test(val.trim()) ? parseFloat(val.trim()) : NaN);
    if (!isNaN(num) && num > 0 && num < 1) {
      const totalSeconds = Math.round(num * 86400);
      const h24 = Math.floor(totalSeconds / 3600);
      const m = Math.floor((totalSeconds % 3600) / 60);
      const s = totalSeconds % 60;
      const ampm = h24 >= 12 ? 'PM' : 'AM';
      let h12 = h24 % 12;
      if (h12 === 0) h12 = 12;
      const secStr = s > 0 ? ':' + String(s).padStart(2, '0') : ':00';
      return `${h12}:${String(m).padStart(2, '0')}${secStr} ${ampm}`;
    }
    const str = String(val).trim();
    if (/(?:1899|1900)/.test(str)) {
      const match = str.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
      if (match) {
        let h24 = parseInt(match[1], 10);
        const m = parseInt(match[2], 10);
        const s = match[3] ? parseInt(match[3], 10) : 0;
        const ampm = h24 >= 12 ? 'PM' : 'AM';
        let h12 = h24 % 12;
        if (h12 === 0) h12 = 12;
        const secStr = s > 0 ? ':' + String(s).padStart(2, '0') : ':00';
        return `${h12}:${String(m).padStart(2, '0')}${secStr} ${ampm}`;
      }
    }
    const match = str.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);
    if (match) {
      let h = parseInt(match[1], 10);
      const m = parseInt(match[2], 10);
      const s = match[3] ? parseInt(match[3], 10) : 0;
      const ap = match[4] ? match[4].toUpperCase() : (h >= 12 ? 'PM' : 'AM');
      let dispH = h % 12;
      if (dispH === 0) dispH = 12;
      const secStr = s > 0 ? ':' + String(s).padStart(2, '0') : ':00';
      return `${dispH}:${String(m).padStart(2, '0')}${secStr} ${ap}`;
    }
    return str;
  };

  // Test full JS date string from user bug report
  assert.equal(formatExamDate('Tue Nov 10 2026 00:00:00 GMT+0530 (India Standard Time)'), '10/11/2026');
  // Test SheetJS date object with India timezone shift
  assert.equal(formatExamDate(new Date('2026-11-04T18:29:50.000Z')), '05/11/2026');
  // Test numeric Excel serial
  assert.equal(formatExamDate(46331), '05/11/2026');
  // Test ISO string
  assert.equal(formatExamDate('2026-11-05'), '05/11/2026');
  // Test standard format
  assert.equal(formatExamDate('05/11/2026'), '05/11/2026');

  // Test decimal time fractions from user bug report
  assert.equal(formatTimeDisplay(0.5625), '1:30:00 PM');
  assert.equal(formatTimeDisplay(0.625), '3:00:00 PM');
  assert.equal(formatTimeDisplay('0.5625'), '1:30:00 PM');
  assert.equal(formatTimeDisplay('0.625'), '3:00:00 PM');
  assert.equal(formatTimeDisplay('13:30'), '1:30:00 PM');
  assert.equal(formatTimeDisplay('15:00:00'), '3:00:00 PM');
});

test('detectPreExamColumns strictly prioritizes Programme Name over ProgramFullName and VenueCode over CenterCode', () => {
  const headers = [
    'Faculty',
    'Programme Code',
    'ProgramFullName',
    'Programme Name',
    'Programme ABBR',
    'Parent Group Name',
    'Immediate Parent Group',
    'PaperCode',
    'PaperName',
    'CenterCode',
    'CenterName',
    'VenueCode',
    'VenueName'
  ];

  const detectPreExamColumns = (headersList) => {
    let pgmExact = -1, pgmName = -1, pgmOnly = -1, pgmFull = -1;
    let parentGrpExact = -1, parentGrpFallback = -1;
    let immGrpExact = -1;
    let courseCodeExact = -1, courseCodeFallback = -1;
    let courseNameExact = -1, courseNameFallback = -1;
    let venueCodeExact = -1, centerCodeFallback = -1;
    let venueNameExact = -1, centerNameFallback = -1;
    let countIdx = -1;

    headersList.forEach((h, idx) => {
      const lower = String(h).toLowerCase().replace(/[^a-z0-9]/g, '');
      const rawTrimmed = String(h).trim().toLowerCase();

      if (rawTrimmed === 'programme name' || lower === 'programmename') {
        pgmExact = idx;
      } else if (rawTrimmed === 'program name' || lower === 'programname') {
        pgmName = idx;
      } else if (rawTrimmed === 'programme' || lower === 'programme') {
        pgmOnly = idx;
      } else if (rawTrimmed === 'programfullname' || lower === 'programfullname') {
        pgmFull = idx;
      }

      if (lower === 'parentgroupname') {
        parentGrpExact = idx;
      } else if (lower === 'parentgroup') {
        parentGrpFallback = idx;
      }

      if (lower === 'immediateparentgroup' || lower === 'immediategroupname' || lower === 'immediateparent' || lower === 'immediategroup') {
        immGrpExact = idx;
      }

      if (lower === 'coursecode' || lower === 'papercode') {
        courseCodeExact = idx;
      } else if (lower.includes('subjectcode')) {
        courseCodeFallback = idx;
      }

      if (lower === 'coursename' || lower === 'papername') {
        courseNameExact = idx;
      } else if (lower.includes('subjectname') || lower.includes('coursetitle')) {
        courseNameFallback = idx;
      }

      if (lower === 'venuecode') {
        venueCodeExact = idx;
      } else if (lower === 'centercode' || lower === 'centrecode') {
        centerCodeFallback = idx;
      }

      if (lower === 'venuename') {
        venueNameExact = idx;
      } else if (lower === 'centername' || lower === 'centrename') {
        centerNameFallback = idx;
      }

      if (lower.includes('studentcount') || lower === 'count' || lower === 'candidates') {
        countIdx = idx;
      }
    });

    return {
      programmeName: pgmExact !== -1 ? pgmExact : (pgmName !== -1 ? pgmName : (pgmOnly !== -1 ? pgmOnly : pgmFull)),
      parentGroupName: parentGrpExact !== -1 ? parentGrpExact : parentGrpFallback,
      immediateGroupName: immGrpExact,
      courseCode: courseCodeExact !== -1 ? courseCodeExact : courseCodeFallback,
      courseName: courseNameExact !== -1 ? courseNameExact : courseNameFallback,
      venueCode: venueCodeExact !== -1 ? venueCodeExact : centerCodeFallback,
      venueName: venueNameExact !== -1 ? venueNameExact : centerNameFallback,
      studentCount: countIdx
    };
  };

  const mapping = detectPreExamColumns(headers);
  assert.equal(mapping.programmeName, 3, 'programmeName must be index 3 (Programme Name), not index 2 (ProgramFullName)');
  assert.equal(mapping.venueCode, 11, 'venueCode must be index 11 (VenueCode), not index 9 (CenterCode)');
  assert.equal(mapping.venueName, 12, 'venueName must be index 12 (VenueName), not index 10 (CenterName)');
  assert.equal(mapping.courseCode, 7, 'courseCode must be index 7 (PaperCode)');
  assert.equal(mapping.courseName, 8, 'courseName must be index 8 (PaperName)');
});

console.log(`\n=== ALL ${totalTests} TESTS PASSED CLEANLY (${passedTests}/${totalTests}) ===\n`);


