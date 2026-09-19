// Test suite specifically for Affiliated Programme Details Engine
import assert from 'node:assert/strict';

console.log('=== TESTING AFFILIATED PROGRAMME DETAILS ENGINE ===\n');

// 1. Course Details splitting and regex parsing matching AffiliatedProgrammePage.jsx
function cleanCourseCode(code) {
  return String(code || '').trim().replace(/[\r\n\t]/g, '').toUpperCase();
}

function extractCourseItems(rawCourseDetails) {
  if (!rawCourseDetails) return [];
  let s = String(rawCourseDetails).trim();
  if (!s) return [];

  // Normalize all line endings
  s = s.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  // Any newline that is NOT immediately followed by optional whitespace and an opening bracket/paren is a wrapped line in the course title!
  // Join it with a space so titles like 'Fundamentals of \nChemistry - II' stay intact:
  s = s.replace(/\n(?!\s*[\(\[\{])/g, ' ');

  // Split by commas, semicolons, pipes, or newlines that precede an opening bracket
  let items = s.split(/[,;\n|]+\s*(?=[\(\[\{])/).map(x => x.replace(/^[,;\s]+|[,;\s]+$/g, '').trim()).filter(Boolean);

  // Fallback: if single item or multiple bracket codes in one chunk
  if (items.length <= 1) {
    const multiMatch = s.split(/(?<=[^\s])\s+(?=[\(\[][A-Z0-9_-]+[\)\]])/i).map(x => x.replace(/^[,;\s]+|[,;\s]+$/g, '').trim()).filter(Boolean);
    if (multiMatch.length > 1) items = multiMatch;
  }

  return items.length > 0 ? items : [s];
}

function parseCourseCodeAndName(courseStr) {
  let s = String(courseStr || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
  s = s.replace(/^[,;\s]+|[,;\s]+$/g, '').trim();
  if (!s) return { courseCode: '', courseName: '' };

  // Format 1: (CODE) Course Name OR [CODE] Course Name OR {CODE} Course Name
  const parenMatch = s.match(/^[\(\[\{]([^\)\]\}]+)[\)\]\}]\s*[-–—:]?\s*(.*)$/);
  if (parenMatch) {
    return {
      courseCode: cleanCourseCode(parenMatch[1] || ''),
      courseName: (parenMatch[2] || parenMatch[1] || '').trim()
    };
  }

  // Format 2: CODE - Course Name OR CODE : Course Name
  const dashMatch = s.match(/^([A-Z0-9_-]{4,25})\s*[-–—:]\s*(.*)$/i);
  if (dashMatch) {
    return {
      courseCode: cleanCourseCode(dashMatch[1] || ''),
      courseName: (dashMatch[2] || '').trim()
    };
  }

  // Format 3: CODE Course Name (e.g. KU02DSCCHE102 Fundamentals of Chemistry - II)
  const spaceMatch = s.match(/^([A-Z]{2,4}\d{1,2}[A-Z]{2,6}\d{2,4})\s+(.*)$/i);
  if (spaceMatch) {
    return {
      courseCode: cleanCourseCode(spaceMatch[1] || ''),
      courseName: (spaceMatch[2] || '').trim()
    };
  }

  return {
    courseCode: '',
    courseName: s
  };
}

function parseCourseDetails(rawDetails) {
  const items = extractCourseItems(rawDetails);
  return items.map(item => {
    const parsed = parseCourseCodeAndName(item);
    return {
      raw: item,
      code: parsed.courseCode,
      name: parsed.courseName
    };
  });
}

// 2. Term and Year parsing
function parseProgramTerm(programTerm) {
  const yearMatch = (programTerm || '').match(/(Year\s+[IVXLCDM\d]+)/i);
  const semMatch = (programTerm || '').match(/(SEMESTER\s+[IVXLCDM\d]+)/i);
  let year = yearMatch ? yearMatch[1] : '';
  const termName = semMatch ? semMatch[1] : '';

  if (!year && termName) {
    const semValMatch = termName.match(/(?:SEMESTER|SEM|S)\s*([IVXLCDM\d]+)/i);
    if (semValMatch) {
      const semVal = semValMatch[1].toUpperCase();
      const semYearMap = { '1': 1, 'I': 1, '2': 1, 'II': 1, '3': 2, 'III': 2, '4': 2, 'IV': 2, '5': 3, 'V': 3, '6': 3, 'VI': 3, '7': 4, 'VII': 4, '8': 4, 'VIII': 4 };
      const y = semYearMap[semVal];
      if (y) year = `Year ${y}`;
    }
  }

  return {
    year,
    termName
  };
}

// 3. Process records like AffiliatedProgrammePage
function processAffiliated(rows) {
  const allExploded = [];
  const dedupeList = [];
  const seen = new Set();
  let dupCount = 0;

  rows.forEach(row => {
    const collegeCode = row['College Code'] || row['ADEC Code'] || '';
    const collegeName = row['College Name'] || row['ADEC Name'] || '';
    const programCode = row['Program Code'] || '';
    const programTerm = row['Program Term'] || '';
    const { year, termName } = parseProgramTerm(programTerm);
    const parsedCourses = parseCourseDetails(row['Course Details']);

    parsedCourses.forEach(c => {
      // 9-column All Rows
      allExploded.push({
        'College Code': collegeCode,
        'College Name': collegeName,
        'Program Code': programCode,
        'Program Term': programTerm,
        'Programme Year': year,
        'Program Term Name': termName,
        'Course Details': c.raw,
        'Course Code': c.code,
        'Course Name': c.name
      });

      // 7-column Deduplicated
      const dedupeKey = `${collegeCode.trim().toLowerCase()}|${programCode.trim().toLowerCase()}|${year.trim().toLowerCase()}|${termName.trim().toLowerCase()}|${(c.code || c.name).trim().toLowerCase()}`;
      if (seen.has(dedupeKey)) {
        dupCount++;
      } else {
        seen.add(dedupeKey);
        dedupeList.push({
          'College Code': collegeCode,
          'College Name': collegeName,
          'Programme Year': year,
          'Program Term Name': termName,
          'Course Details': c.raw,
          'Course Code': c.code,
          'Course Name': c.name
        });
      }
    });
  });

  return { allExploded, dedupeList, dupCount };
}

// Test Case 1: Standard multi-course cell with parens
console.log('[Test 1] Multi-course cell explosion');
const sampleRow1 = {
  'ADEC Code': 'KU01',
  'ADEC Name': 'Kannur University Campus',
  'Program Code': 'UGANT',
  'Program Term': 'Year 1 / SEMESTER 1',
  'Course Details': '(KU01DSCANT101) Anthropological Theories, (KU01AECENG101) Communicative English, (KU01VACENV101) Environmental Studies'
};

const res1 = processAffiliated([sampleRow1]);
assert.equal(res1.allExploded.length, 3, 'Should explode into 3 rows');
assert.equal(res1.allExploded[0]['Course Code'], 'KU01DSCANT101');
assert.equal(res1.allExploded[0]['Course Name'], 'Anthropological Theories');
assert.equal(res1.allExploded[0]['Programme Year'], 'Year 1');
assert.equal(res1.allExploded[0]['Program Term Name'], 'SEMESTER 1');
assert.equal(res1.allExploded[1]['Course Code'], 'KU01AECENG101');
assert.equal(res1.allExploded[2]['Course Code'], 'KU01VACENV101');
console.log('  ✓ Multi-course cell parsed into 3 individual courses correctly');

// Test Case 2: Deduplication across multiple rows for the same college & program
console.log('\n[Test 2] Deduplication logic across duplicate rows');
const sampleRow2 = {
  'ADEC Code': 'KU01',
  'ADEC Name': 'Kannur University Campus',
  'Program Code': 'UGANT',
  'Program Term': 'Year 1 / SEMESTER 1',
  'Course Details': '(KU01DSCANT101) Anthropological Theories' // Duplicate of row 1 course
};

const res2 = processAffiliated([sampleRow1, sampleRow2]);
assert.equal(res2.allExploded.length, 4, 'All Exploded should have 4 rows');
assert.equal(res2.dedupeList.length, 3, 'Deduplicated should still have 3 unique courses');
assert.equal(res2.dupCount, 1, 'Should record 1 duplicate');
console.log('  ✓ Duplicate course correctly filtered in Deduplicated (7 Col) list');

// Test Case 3: Roman numeral semester parsing (e.g. Year I / SEMESTER II)
console.log('\n[Test 3] Roman numeral term parsing');
const termTest = parseProgramTerm('Year II / SEMESTER IV');
assert.equal(termTest.year, 'Year II');
assert.equal(termTest.termName, 'SEMESTER IV');
console.log('  ✓ Roman numeral term parsed correctly');

// Test Case 4: Course without parentheses fallback
console.log('\n[Test 4] Course without parentheses');
const fallbackRes = parseCourseDetails('General Human Physiology');
assert.equal(fallbackRes.length, 1);
assert.equal(fallbackRes[0].code, '');
assert.equal(fallbackRes[0].name, 'General Human Physiology');
console.log('  ✓ Course without parens safely kept as Course Name with blank code');

// Test Case 5: User Query Case - Tab-prefixed course "(KU02DSCCHE102) Fundamentals of Chemistry - II"
console.log('\n[Test 5] Tab-prefixed course string parsing');
const chemRes = parseCourseCodeAndName('\t(KU02DSCCHE102) Fundamentals of Chemistry - II');
assert.equal(chemRes.courseCode, 'KU02DSCCHE102', 'Course code should be KU02DSCCHE102');
assert.equal(chemRes.courseName, 'Fundamentals of Chemistry - II', 'Course name should be Fundamentals of Chemistry - II');
console.log('  ✓ Tab-prefixed chemistry course correctly split into code and name');

// Test Case 6: Multi-line Excel cell (Alt+Enter) with semicolon and trailing commas
console.log('\n[Test 6] Multi-line Excel cell with mixed delimiters');
const multiLineCell = `\t(KU02DSCCHE102) Fundamentals of Chemistry - II\r\n(KU02DSCCHE103) Chemistry Practical - I;\n[KU02AECENG101] General English`;
const multiLineParsed = parseCourseDetails(multiLineCell);
assert.equal(multiLineParsed.length, 3, 'Should split 3 courses across lines');
assert.equal(multiLineParsed[0].code, 'KU02DSCCHE102');
assert.equal(multiLineParsed[0].name, 'Fundamentals of Chemistry - II');
assert.equal(multiLineParsed[1].code, 'KU02DSCCHE103');
assert.equal(multiLineParsed[1].name, 'Chemistry Practical - I');
assert.equal(multiLineParsed[2].code, 'KU02AECENG101');
assert.equal(multiLineParsed[2].name, 'General English');
console.log('  ✓ Multi-line Excel cell successfully exploded and parsed');

// Test Case 7: Wrapped newline inside course title (e.g. "Fundamentals of \nChemistry - II")
console.log('\n[Test 7] Wrapped newline inside course name');
const wrappedCell = '(KU02DSCCHE102) Fundamentals of \nChemistry - II, (KU02DSCCHE103) Coordination \nChemistry - I';
const wrappedParsed = parseCourseDetails(wrappedCell);
assert.equal(wrappedParsed.length, 2, 'Should be 2 courses, not 4 broken lines');
assert.equal(wrappedParsed[0].code, 'KU02DSCCHE102');
assert.equal(wrappedParsed[0].name, 'Fundamentals of Chemistry - II');
assert.equal(wrappedParsed[1].code, 'KU02DSCCHE103');
assert.equal(wrappedParsed[1].name, 'Coordination Chemistry - I');
console.log('  ✓ Wrapped newline inside course name preserved and normalized');

// Test Case 8: Semester-to-Year fallback inference
console.log('\n[Test 8] Semester-to-Year fallback inference');
const noYearTerm = parseProgramTerm('Bachelor of Arts in Political Science-BA SEMESTER II');
assert.equal(noYearTerm.year, 'Year 1');
assert.equal(noYearTerm.termName, 'SEMESTER II');
console.log('  ✓ Year inferred as "Year 1" from "SEMESTER II"');

console.log('\n=== ALL AFFILIATED PROGRAMME ENGINE TESTS PASSED CLEANLY! ===\n');
