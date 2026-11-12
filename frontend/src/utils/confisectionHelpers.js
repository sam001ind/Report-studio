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

// Resolves official college email address using directory lookups, code aliases, and fuzzy name matching
export const resolveCollegeEmail = (code, name, emailDirectory = DEFAULT_COLLEGE_EMAILS, fallbackEmail = '') => {
  if (fallbackEmail && String(fallbackEmail).includes('@')) {
    return String(fallbackEmail).trim();
  }
  const cleanCode = String(code || '').trim().toLowerCase();
  const cleanName = String(name || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');

  if (!emailDirectory || !Array.isArray(emailDirectory)) return '';

  // 1. Direct code match (e.g. '209', '201')
  const byCode = emailDirectory.find(e => String(e.code || '').trim().toLowerCase() === cleanCode);
  if (byCode && byCode.email) return byCode.email;

  // 2. Known Kannur Alpha Venue Code aliases
  const aliasMap = {
    pp: '209', // Payyannur College
    st: '211', // Sir Syed College
    sn: '212', // Sree Narayana College
    bc: '201', // Brennen College
    mc: '200', // Mananthavady College
    kc: '205', // Kasaragod College
    nc: '214', // Nirmalagiri College
    mg: '215', // MG College Iritty
    nam: '216', // NAM College Kallikandy
    nas: '241', // Nehru Arts & Science
    sp: '251', // St. Pius X
    ses: '265', // SES College
    mm: '266'  // Mary Matha College
  };
  if (aliasMap[cleanCode]) {
    const byAlias = emailDirectory.find(e => String(e.code || '').trim() === aliasMap[cleanCode]);
    if (byAlias && byAlias.email) return byAlias.email;
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
    collegeEmail: 'payyannurcollege@gmail.com',
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
    collegeEmail: 'payyannurcollege@gmail.com',
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
    collegeEmail: 'payyannurcollege@gmail.com',
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
    collegeEmail: 'payyannurcollege@gmail.com',
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
    collegeEmail: 'payyannurcollege@gmail.com',
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
    collegeEmail: 'payyannurcollege@gmail.com',
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
    collegeEmail: 'payyannurcollege@gmail.com',
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
    collegeEmail: 'sirsyedcollege@gmail.com',
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
    collegeEmail: 'sirsyedcollege@gmail.com',
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
    collegeEmail: 'sncollegekannur@gmail.com',
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
