export const TEMPLATE_ARCHETYPES = {
  NOMINAL_ROLL: {
    id: 'NOMINAL_ROLL',
    name: 'Venue-Wise Nominal Roll',
    description: 'Multi-course nested sub-row table grouped by Programme + Venue with candidate registration merging (Image 1 style).',
    category: 'Roster / Student Roll',
    icon: 'FileText',
    defaultConfig: {
      universityName: 'Kannur University',
      branchName: '(Examination Branch)',
      examTitle: 'IV Semester Private Registration 2024 -2027 Admission',
      sessionName: 'April 2026',
      groupBy: 'programme_venue', // 'programme_venue' | 'venue' | 'programme' | 'none'
      showGroupBox: true,
      formatCodeDotNameDot: true,
      headers: {
        slNo: 'Sl No',
        regNo: 'Register Number',
        name: 'Candidate Name',
        courses: 'Courses',
        remarks: 'Remarks'
      },
      columnStyles: {
        slNo: 10,
        regNo: 30,
        name: 42,
        courses: 76,
        remarks: 24
      }
    }
  },
  QP_STATEMENT: {
    id: 'QP_STATEMENT',
    name: 'QP Statement Report (Venue-Wise)',
    description: 'Daily exam schedule and question paper requirement matrix table per venue with blank QP/LP columns (Image 2 style).',
    category: 'Examination Matrix',
    icon: 'CalendarRange',
    defaultConfig: {
      universityName: 'Kannur University',
      branchName: '(Examination Branch)',
      examTitle: 'QP Statement for 4th Semester Degree Private Registration Regular Examination',
      sessionName: 'April 2026',
      centerPrefix: 'Center Name :',
      groupBy: 'venue', // 'venue' | 'center' | 'none'
      formatCodeDotNameDot: false,
      headers: {
        slNo: 'SL\nNo',
        date: 'Date',
        course: 'Course',
        count: 'NC',
        qp: 'QP',
        lp: 'LP'
      },
      columnStyles: {
        slNo: 12,
        date: 35,
        course: 85,
        count: 16,
        qp: 17,
        lp: 17
      }
    }
  },
  QP_COVER_LABEL: {
    id: 'QP_COVER_LABEL',
    name: 'QP Cover Label & Certificate',
    description: 'Bilingual header with university crest, key-value exam details box, certificate statement, and signature lines (Image 3 style).',
    category: 'Packaging Envelope / Certificate',
    icon: 'Tag',
    defaultConfig: {
      universityName: 'Kannur University',
      malayalamTitle: 'കണ്ണൂർ സർവകലാശാല',
      addressLine: 'Thavakkara,Civil Station P.O, Kannur',
      naacRating: "Reaccredited by NAAC with 'B++' Grade",
      branchName: '(Examination Branch)',
      examTitle: 'Fourth Semester Degree (Private Registration) Regular Examinations April 2026',
      showLogo: true,
      centerPrefix: 'CENTRE CODE AND NAME',
      certificateText: 'We hereby certify that we have examined this cover and satisfied ourselves that the seals are intact and that it was opened at ____________________________________ A.M/P.M in our presence.',
      signatures: {
        invigilatorLabel: 'INVIGILATOR',
        invigilatorCount: 2,
        addlChiefLabel: 'ADDL.CHIEF SUPERINTENDENT',
        chiefLabel: 'CHIEF SUPERINTENDENT',
        showPlaceDate: true
      }
    }
  },
  CUSTOM_TABULAR: {
    id: 'CUSTOM_TABULAR',
    name: 'Custom Tabular Data Report',
    description: 'Dynamic multi-column grid report with customizable headers, groupings, subtotals, and column styles.',
    category: 'Custom General Report',
    icon: 'TableProperties',
    defaultConfig: {
      universityName: 'Kannur University',
      branchName: 'Institutional Report',
      examTitle: 'Data Summary Statement',
      sessionName: '',
      groupBy: 'none',
      showTotalCount: true,
      selectedColumns: [],
      columnLabels: {},
      columnWidths: {}
    }
  }
};

/**
 * Intelligent column detector for uploaded Excel columns
 */
export function autoDetectDatasetColumns(headers = []) {
  const norm = headers.map(h => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));
  
  const findBest = (patterns) => {
    for (const pat of patterns) {
      if (typeof pat === 'string') {
        const idx = norm.findIndex(n => n === pat);
        if (idx !== -1) return headers[idx];
      } else if (typeof pat === 'function') {
        const idx = norm.findIndex(pat);
        if (idx !== -1) return headers[idx];
      }
    }
    return '';
  };

  return {
    seatNo: findBest(['seatnumber', 'seatno', 'registerno', 'registernumber', 'regno', 'candidatecode', 'prn', n => n.includes('seat') || n.includes('reg')]),
    name: findBest(['fullname', 'studentname', 'candidatename', n => (n.includes('name') || n.includes('candidate')) && !n.includes('father') && !n.includes('venue') && !n.includes('college') && !n.includes('program') && !n.includes('course') && !n.includes('paper') && !n.includes('regional'), 'name']),
    venueCode: findBest(['venuecode', 'examcentercode', 'centercode', 'centrecode', n => (n.includes('venue') || n.includes('center') || n.includes('centre')) && n.includes('code') && !n.includes('regional')]),
    venueName: findBest(['venuename', 'examcentername', 'centername', 'centrename', n => (n.includes('venue') || n.includes('center') || n.includes('centre')) && (n.includes('name') || n.includes('title')) && !n.includes('regional')]),
    courseCode: findBest(['papercode', 'coursecode', 'subjectcode', 'qpcode', n => (n.includes('paper') || n.includes('course') || n.includes('subject') || n.includes('qp')) && n.includes('code')]),
    courseTitle: findBest(['papername', 'coursetitle', 'subjectname', 'coursename', 'papertitle', n => (n.includes('paper') || n.includes('course') || n.includes('subject')) && (n.includes('name') || n.includes('title')) && !n.includes('code')]),
    programme: findBest(['programmename', 'programname', 'branch', 'programfullname', 'programmefullname', n => n.includes('program') && n.includes('name') && !n.includes('term') && !n.includes('code')]),
    examDate: findBest(['examdate', 'date', 'examinationdate', n => n.includes('date') && !n.includes('birth')]),
    startTime: findBest(['coursestarttime', 'starttime', 'start', n => n.includes('start') && n.includes('time')]),
    endTime: findBest(['courseendtime', 'endtime', 'end', n => n.includes('end') && n.includes('time')]),
    count: findBest(['studentcount', 'count', 'nc', 'candidatecount', 'candidates', n => (n.includes('student') || n.includes('candidate') || n.includes('total')) && n.includes('count')]),
    session: findBest(['eventname', 'programterm', 'examsession', 'session', n => n.includes('event') && n.includes('name')])
  };
}

/**
 * Suggest best matching archetype based on uploaded columns
 */
export function suggestArchetype(detectedCols) {
  if (detectedCols.seatNo && (detectedCols.courseCode || detectedCols.courseTitle)) {
    return 'NOMINAL_ROLL';
  }
  if (detectedCols.examDate && (detectedCols.courseCode || detectedCols.courseTitle)) {
    return 'QP_STATEMENT';
  }
  if (detectedCols.venueName && (detectedCols.courseCode || detectedCols.courseTitle) && detectedCols.examDate) {
    return 'QP_COVER_LABEL';
  }
  return 'CUSTOM_TABULAR';
}
