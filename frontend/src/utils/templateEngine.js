import { logoBase64 } from '../assets/logoBase64';

export const DEFAULT_REPORT_TEMPLATE = {
  pageSize: 'A4',
  orientation: 'portrait',
  logo: {
    show: false,
    src: logoBase64,
    position: 'center',
    width: 18
  },
  headersList: [
    { id: 'h1', text: 'Kannur University', size: 15, bold: true, italic: false, align: 'center', color: '#000000', font: 'helvetica' },
    { id: 'h2', text: '(Examination Branch)', size: 11, bold: true, italic: false, align: 'center', color: '#111111', font: 'helvetica' },
    { id: 'h3', text: 'Custom Examination Consolidated Report', size: 10, bold: true, italic: false, align: 'center', color: '#111111', font: 'helvetica' },
    { id: 'h4', text: 'Academic Session 2024 - 2026', size: 9.5, bold: false, italic: false, align: 'center', color: '#222222', font: 'helvetica' }
  ],
  tableColumns: [
    { id: 'c1', label: 'Sl No', field: 'slNo', width: 12, align: 'center' },
    { id: 'c2', label: 'Register Number', field: 'seatNo', width: 35, align: 'center', bold: true },
    { id: 'c3', label: 'Candidate Name', field: 'name', width: 45, align: 'left', bold: true },
    { id: 'c4', label: 'Course / Paper', field: 'course', width: 65, align: 'left' },
    { id: 'c5', label: 'Remarks / Status', field: 'remarks', width: 25, align: 'center' }
  ],
  tableTheme: {
    headerBg: '#f1f5f9',
    headerColor: '#000000',
    fontSize: 8.5,
    borderColor: '#64748b'
  },
  groupBy: 'none',
  showTotalCount: true,
  formatCodeDotNameDot: false
};

export const TEMPLATE_ARCHETYPES = {
  CUSTOM_TABULAR: {
    id: 'CUSTOM_TABULAR',
    name: 'Custom Tabular Data Report',
    description: 'Dynamic multi-column report builder with customizable headers, logo, table columns, grouping, and multi-format exports.',
    category: 'Custom General Report',
    icon: 'TableProperties',
    defaultConfig: DEFAULT_REPORT_TEMPLATE
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

export function suggestArchetype(_detectedCols) {
  return 'CUSTOM_TABULAR';
}
