import { DEFAULT_COLLEGE_EMAILS, formatExamDate, normalizeKey } from './venueEmailHelpers.js';

export { DEFAULT_COLLEGE_EMAILS, formatExamDate, normalizeKey };

export const CONFIDENTIAL_SECTION_COLUMNS = [
  'Event Name',
  'Center Code',
  'Center Name',
  'Venue Code',
  'Venue Name',
  'Program Name',
  'Exam Date',
  'Course  Start Time',
  'Course  End Time',
  'Course  Code',
  'Course  Name',
  'Student Count',
  'venueId'
];

export const QP_EMAIL_DISPATCH_COLUMNS = [
  'Exam Date',
  'Course  Code',
  'Course  Name',
  'Venue Code',
  'Venue Name',
  'Email ID'
];

// Helper to generate the exact QP Email Dispatch Report records
export const generateQpEmailDispatchRows = (consolidatedRows) => {
  if (!Array.isArray(consolidatedRows)) return [];
  return consolidatedRows.map(r => ({
    examDate: r.examDate,
    courseCode: r.courseCode,
    courseName: r.courseName,
    venueCode: r.venueCode,
    venueName: r.venueName,
    emailId: r.collegeEmail || '',
    hasEmail: Boolean(r.collegeEmail && r.collegeEmail.includes('@')),
    // Details
    session: r.session,
    startTime: r.startTime,
    endTime: r.endTime,
    studentCount: r.studentCount,
    p20: r.p20,
    p10: r.p10,
    p5: r.p5,
    singles: r.singles,
    buffer: r.buffer,
    totalQp: r.totalQp,
    programsFormatted: r.programsFormatted || ''
  }));
};

// Robust column auto-detector prioritizing exact matches for Course Code, Course Name, Venue Code, Venue Name
export const detectColumns = (headers) => {
  const norm = (headers || []).map(h => normalizeKey(h));

  const findExact = (candidates) => {
    for (const cand of candidates) {
      const cNorm = normalizeKey(cand);
      const idx = norm.indexOf(cNorm);
      if (idx !== -1) return idx;
    }
    return -1;
  };

  // 1. Course Code: strictly prioritize 'Course Code', 'Course  Code', 'Subject Code'
  let courseCodeIdx = findExact(['Course Code', 'Course  Code', 'CourseCode', 'Subject Code', 'Paper Code', 'Course No']);
  if (courseCodeIdx === -1) {
    courseCodeIdx = norm.findIndex(h => (h.includes('course') || h.includes('subject') || h.includes('paper')) && h.includes('code') && !h.includes('center') && !h.includes('venue'));
  }
  if (courseCodeIdx === -1) {
    courseCodeIdx = norm.findIndex(h => h === 'code');
  }

  // 2. Course Name: strictly prioritize 'Course Name', 'Course  Name', 'Subject Name'
  let courseNameIdx = findExact(['Course Name', 'Course  Name', 'CourseName', 'Subject Name', 'Paper Name', 'Course Title']);
  if (courseNameIdx === -1) {
    courseNameIdx = norm.findIndex(h => (h.includes('course') || h.includes('subject') || h.includes('paper')) && (h.includes('name') || h.includes('title')) && !h.includes('center') && !h.includes('venue') && !h.includes('time'));
  }
  if (courseNameIdx === -1) {
    courseNameIdx = norm.findIndex(h => h === 'course' || h === 'subject');
  }

  // 3. Venue Code: strictly prioritize 'Venue Code', 'Venue  Code', 'College Code'
  let venueCodeIdx = findExact(['Venue Code', 'Venue  Code', 'VenueCode', 'College Code']);
  if (venueCodeIdx === -1) {
    venueCodeIdx = norm.findIndex(h => h.includes('venue') && h.includes('code'));
  }
  if (venueCodeIdx === -1) {
    venueCodeIdx = norm.findIndex(h => h.includes('venue') || h.includes('collegecode'));
  }

  // 4. Venue Name: strictly prioritize 'Venue Name', 'Venue  Name', 'College Name'
  let venueNameIdx = findExact(['Venue Name', 'Venue  Name', 'VenueName', 'College Name', 'Institution']);
  if (venueNameIdx === -1) {
    venueNameIdx = norm.findIndex(h => (h.includes('venue') || h.includes('college')) && h.includes('name'));
  }
  if (venueNameIdx === -1) {
    venueNameIdx = norm.findIndex(h => h.includes('venue') || h.includes('college') || h.includes('institution'));
  }

  // 5. Center Code & Center Name
  let centerCodeIdx = findExact(['Center Code', 'CenterCode', 'Centre Code', 'CentreCode']);
  if (centerCodeIdx === -1) {
    centerCodeIdx = norm.findIndex(h => (h.includes('center') || h.includes('centre')) && h.includes('code'));
  }

  let centerNameIdx = findExact(['Center Name', 'CenterName', 'Centre Name', 'CentreName']);
  if (centerNameIdx === -1) {
    centerNameIdx = norm.findIndex(h => (h.includes('center') || h.includes('centre')) && h.includes('name'));
  }

  // 6. Exam Date
  let examDateIdx = findExact(['Exam Date', 'ExamDate', 'Date of Exam', 'Date']);
  if (examDateIdx === -1) {
    examDateIdx = norm.findIndex(h => h.includes('date'));
  }

  // 7. Start Time & End Time
  let startTimeIdx = findExact(['Course Start Time', 'Course  Start Time', 'CourseStartTime', 'Start Time', 'StartTime']);
  if (startTimeIdx === -1) {
    startTimeIdx = norm.findIndex(h => h.includes('start') && (h.includes('time') || h.includes('course')));
  }
  if (startTimeIdx === -1) {
    startTimeIdx = norm.findIndex(h => h.includes('time') && !h.includes('end'));
  }

  let endTimeIdx = findExact(['Course End Time', 'Course  End Time', 'CourseEndTime', 'End Time', 'EndTime']);
  if (endTimeIdx === -1) {
    endTimeIdx = norm.findIndex(h => h.includes('end') && (h.includes('time') || h.includes('course')));
  }

  // 8. Event Name
  let eventNameIdx = findExact(['Event Name', 'EventName', 'Event', 'Examination']);
  if (eventNameIdx === -1) {
    eventNameIdx = norm.findIndex(h => h.includes('event'));
  }

  // 9. Program Name
  let programNameIdx = findExact(['Program Name', 'ProgramName', 'Programme Name', 'ProgrammeName', 'Program', 'Programme']);
  if (programNameIdx === -1) {
    programNameIdx = norm.findIndex(h => h.includes('prog') && h.includes('name'));
  }
  if (programNameIdx === -1) {
    programNameIdx = norm.findIndex(h => h.includes('prog'));
  }

  // 10. Student Count
  let studentCountIdx = findExact(['Student Count', 'StudentCount', 'Candidate Count', 'Count of Takers', 'Students', 'Takers', 'Count']);
  if (studentCountIdx === -1) {
    studentCountIdx = norm.findIndex(h => (h.includes('student') || h.includes('candidate') || h.includes('taker')) && h.includes('count'));
  }
  if (studentCountIdx === -1) {
    studentCountIdx = norm.findIndex(h => h.includes('count') || h.includes('students') || h.includes('takers'));
  }

  // 11. venueId
  let venueIdIdx = findExact(['venueId', 'venue_id', 'vid']);
  if (venueIdIdx === -1) {
    venueIdIdx = norm.findIndex(h => h === 'venueid' || h === 'vid');
  }

  // 12. Email
  let emailIdx = findExact(['Email ID', 'Email', 'College Email', 'Venue Email', 'Mail']);
  if (emailIdx === -1) {
    emailIdx = norm.findIndex(h => h.includes('email') || h.includes('mail'));
  }

  return {
    eventName: eventNameIdx,
    centerCode: centerCodeIdx,
    centerName: centerNameIdx,
    venueCode: venueCodeIdx,
    venueName: venueNameIdx,
    programName: programNameIdx,
    examDate: examDateIdx,
    startTime: startTimeIdx,
    endTime: endTimeIdx,
    courseCode: courseCodeIdx,
    courseName: courseNameIdx,
    studentCount: studentCountIdx,
    venueId: venueIdIdx,
    email: emailIdx
  };
};

// Resolves official college email address using directory lookups, code aliases, and fuzzy name matching
export const resolveCollegeEmail = (code, name, emailDirectory = DEFAULT_COLLEGE_EMAILS, fallbackEmail = '') => {
  if (fallbackEmail && String(fallbackEmail).includes('@')) {
    return String(fallbackEmail).trim();
  }
  const cleanCode = String(code || '').trim().toLowerCase();
  const cleanName = String(name || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');

  if (!emailDirectory || !Array.isArray(emailDirectory)) return '';

  // 1. Direct code match (e.g. 'PP', 'GB', 'ST', 'SS')
  const byCode = emailDirectory.find(e => String(e.code || '').trim().toLowerCase() === cleanCode);
  if (byCode && byCode.email) return byCode.email;

  // 2. Numeric center code to 2-letter alpha code aliases
  const numericToAlpha = {
    '209': 'pp',
    '201': 'gb',
    '205': 'gk',
    '211': 'ss',
    '212': 'sn',
    '214': 'ng',
    '215': 'mg',
    '216': 'nm',
    '241': 'na',
    '251': 'sp',
    '265': 'se',
    '266': 'mm',
    '200': 'gm',
    '207': 'km'
  };
  if (numericToAlpha[cleanCode]) {
    const targetAlpha = numericToAlpha[cleanCode];
    const byNumericAlias = emailDirectory.find(e => String(e.code || '').trim().toLowerCase() === targetAlpha);
    if (byNumericAlias && byNumericAlias.email) return byNumericAlias.email;
  }

  // 3. Exact clean name match
  if (cleanName) {
    const byExactName = emailDirectory.find(e => {
      const eClean = String(e.name || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      return eClean === cleanName;
    });
    if (byExactName && byExactName.email) return byExactName.email;

    // 4. Substring match
    const bySubName = emailDirectory.find(e => {
      const eClean = String(e.name || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
      return (eClean.length > 5 && cleanName.includes(eClean)) || (cleanName.length > 5 && eClean.includes(cleanName));
    });
    if (bySubName && bySubName.email) return bySubName.email;
  }

  return '';
};

// Helper to calculate question paper packets
export const calculateQpPackets = (studentCount, bufferRule = 'plus1') => {
  const count = parseInt(studentCount, 10) || 0;
  if (count <= 0) {
    return { p20: 0, p10: 0, p5: 0, singles: 0, buffer: 0, totalQp: 0 };
  }

  let buffer = 0;
  if (bufferRule === 'plus1') {
    buffer = 1;
  } else if (bufferRule === 'plus2') {
    buffer = 2;
  } else if (bufferRule === 'five_percent') {
    buffer = Math.max(1, Math.ceil(count * 0.05));
  } else if (bufferRule === 'none') {
    buffer = 0;
  }

  const totalQp = count + buffer;

  const p20 = Math.floor(totalQp / 20);
  let rem = totalQp % 20;

  const p10 = Math.floor(rem / 10);
  rem = rem % 10;

  const p5 = Math.floor(rem / 5);
  const singles = rem % 5;

  return {
    p20,
    p10,
    p5,
    singles,
    buffer,
    totalQp
  };
};

// Helper for Session detection
export const getExamSession = (startTime, endTime) => {
  const s = `${startTime || ''} ${endTime || ''}`.toUpperCase();
  if (s.includes('AM')) return 'FN';
  if (s.includes('PM')) {
    if (s.includes('09:') || s.includes('10:') || s.includes('11:')) return 'FN';
    return 'AN';
  }
  return 'FN';
};

export const SAMPLE_CONFISECTION_ROWS = [
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'PP',
    venueName: 'Payyannur College, Payyannur',
    collegeEmail: 'principal_pc@kannuruniv.ac.in',
    programName: 'Bachelor of Arts in Functional Hindi(with Credits)-Regular-FYUGP-2025-Functional Hindi-BA Year II SEMESTER III',
    examDate: '25/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCECO203',
    courseName: 'Economics of Banking and Finance',
    studentCount: 4,
    venueId: '68'
  },
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'PP',
    venueName: 'Payyannur College, Payyannur',
    collegeEmail: 'principal_pc@kannuruniv.ac.in',
    programName: 'Bachelor of Commerce(Finance)(with Credits)-Regular-FYUGP-2025-Commerce-BCom Year II SEMESTER III',
    examDate: '25/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCECO203',
    courseName: 'Economics of Banking and Finance',
    studentCount: 11,
    venueId: '68'
  },
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'PP',
    venueName: 'Payyannur College, Payyannur',
    collegeEmail: 'principal_pc@kannuruniv.ac.in',
    programName: 'Bachelor of Science in Mathematics(with Credits)-Regular-FYUGP-2025-Mathematics- SEMESTER III',
    examDate: '25/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCECO203',
    courseName: 'Economics of Banking and Finance',
    studentCount: 1,
    venueId: '68'
  },
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'PP',
    venueName: 'Payyannur College, Payyannur',
    collegeEmail: 'principal_pc@kannuruniv.ac.in',
    programName: 'Bachelor of Arts in Functional Hindi(with Credits)-Regular-FYUGP-2025-Functional Hindi-BA Year II SEMESTER III',
    examDate: '23/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCECO205',
    courseName: 'Economics of Care and Ageing',
    studentCount: 1,
    venueId: '68'
  },
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'PP',
    venueName: 'Payyannur College, Payyannur',
    collegeEmail: 'principal_pc@kannuruniv.ac.in',
    programName: 'Bachelor of Arts in Functional Hindi(with Credits)-Regular-FYUGP-2025-Functional Hindi-BA Year II SEMESTER III',
    examDate: '25/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCENG205',
    courseName: 'Migration and Border Studies',
    studentCount: 5,
    venueId: '68'
  },
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'PP',
    venueName: 'Payyannur College, Payyannur',
    collegeEmail: 'principal_pc@kannuruniv.ac.in',
    programName: 'Bachelor of Arts in English(with Credits)-Regular-FYUGP-2025-English-BA Year II SEMESTER III',
    examDate: '25/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCENG205',
    courseName: 'Migration and Border Studies',
    studentCount: 22,
    venueId: '68'
  },
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'PP',
    venueName: 'Payyannur College, Payyannur',
    collegeEmail: 'principal_pc@kannuruniv.ac.in',
    programName: 'Bachelor of Arts in Functional Hindi(with Credits)-Regular-FYUGP-2025-Functional Hindi-BA Year II SEMESTER III',
    examDate: '23/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCENG204',
    courseName: 'Crime and Detective Narratives',
    studentCount: 4,
    venueId: '68'
  },
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'ST',
    venueName: 'Sir Syed College, Taliparamba',
    collegeEmail: 'principal_ssitt@kannuruniv.ac.in',
    programName: 'Bachelor of Science in Chemistry(with Credits)-Regular-FYUGP-2025-Chemistry-BSc Year II SEMESTER III',
    examDate: '25/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCECO203',
    courseName: 'Economics of Banking and Finance',
    studentCount: 18,
    venueId: '72'
  },
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'ST',
    venueName: 'Sir Syed College, Taliparamba',
    collegeEmail: 'principal_ssitt@kannuruniv.ac.in',
    programName: 'Bachelor of Arts in Economics(with Credits)-Regular-FYUGP-2025-Economics-BA Year II SEMESTER III',
    examDate: '25/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCECO203',
    courseName: 'Economics of Banking and Finance',
    studentCount: 35,
    venueId: '72'
  },
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'SN',
    venueName: 'Sree Narayana College, Kannur',
    collegeEmail: 'principal_snct@kannuruniv.ac.in',
    programName: 'Bachelor of Arts in English(with Credits)-Regular-FYUGP-2025-English-BA Year II SEMESTER III',
    examDate: '25/11/2026',
    startTime: '01:30 PM',
    endTime: '03:30 PM',
    courseCode: 'KU3DSCENG205',
    courseName: 'Migration and Border Studies',
    studentCount: 40,
    venueId: '75'
  }
];
