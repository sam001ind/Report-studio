import React, { useState, useMemo, useRef } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Upload,
  FileSpreadsheet,
  Download,
  Search,
  Sliders,
  CheckCircle2,
  AlertCircle,
  HelpCircle,
  Building2,
  Calendar,
  BookOpen,
  Users,
  RefreshCw,
  Printer,
  ChevronDown,
  Layers,
  Sparkles,
  PackageCheck,
  ShieldCheck,
  X,
  FileText,
  CalendarDays,
  Clock,
  Send,
  SlidersHorizontal,
  ChevronLeft,
  ChevronRight,
  Filter
} from 'lucide-react';
import { readSpreadsheetAsAoa, cleanTimeText, normalizeKeyText, cleanDisplayField } from '../utils/excelParser';
import { formatExamDate } from '../utils/venueEmailHelpers';
import { 
  CONFIDENTIAL_SECTION_COLUMNS, 
  calculateQpPackets, 
  getExamSession 
} from '../utils/confisectionHelpers';

export { CONFIDENTIAL_SECTION_COLUMNS, calculateQpPackets, getExamSession };

export const SAMPLE_CONFISECTION_ROWS = [
  {
    eventName: 'III Sem(R) November 2026',
    centerCode: '1',
    centerName: 'Kannur',
    venueCode: 'PP',
    venueName: 'Payyannur College, Payyannur',
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

export default function ConfisectionExamReportPage() {
  const [rawRows, setRawRows] = useState([]);
  const [fileName, setFileName] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusAlert, setStatusAlert] = useState(null);

  // Configuration options
  const [activeTab, setActiveTab] = useState('consolidated'); // 'consolidated' | 'press_print' | 'venue_packing' | 'course_summary' | 'raw'
  const [bufferRule, setBufferRule] = useState('plus1'); // 'none' | 'plus1' | 'plus2' | 'five_percent'
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedDate, setSelectedDate] = useState('ALL');
  const [selectedVenue, setSelectedVenue] = useState('ALL');
  const [selectedSession, setSelectedSession] = useState('ALL'); // 'ALL' | 'FN' | 'AN'
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(0);
  const [showGuideModal, setShowGuideModal] = useState(false);

  const fileInputRef = useRef(null);

  // Column header auto-detector
  const detectColumns = (headers) => {
    const norm = headers.map(h => normalizeKeyText(h));
    const findIndex = (patterns) => {
      return norm.findIndex(h => patterns.some(p => h === p || h.includes(p)));
    };

    return {
      eventName: findIndex(['eventname', 'event']),
      centerCode: findIndex(['centercode', 'center']),
      centerName: findIndex(['centername', 'centre']),
      venueCode: findIndex(['venuecode', 'venue']),
      venueName: findIndex(['venuename', 'college']),
      programName: findIndex(['programname', 'programme', 'prog']),
      examDate: findIndex(['examdate', 'date']),
      startTime: findIndex(['coursestarttime', 'starttime', 'start']),
      endTime: findIndex(['courseendtime', 'endtime', 'end']),
      courseCode: findIndex(['coursecode', 'code']),
      courseName: findIndex(['coursename', 'course']),
      studentCount: findIndex(['studentcount', 'students', 'count', 'takers', 'candidates']),
      venueId: findIndex(['venueid', 'vid'])
    };
  };

  // Parse uploaded file
  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setFileName(file.name);
    setStatusAlert({ type: 'info', message: `Processing "${file.name}"...` });

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const aoa = readSpreadsheetAsAoa(data);

        if (!aoa || aoa.length < 2) {
          throw new Error('Spreadsheet appears to be empty or has no data rows.');
        }

        // Find header row in first 5 rows
        let headerRowIdx = 0;
        for (let i = 0; i < Math.min(5, aoa.length); i++) {
          const rowText = (aoa[i] || []).map(c => normalizeKeyText(c)).join(' ');
          if (rowText.includes('course') && (rowText.includes('venue') || rowText.includes('date'))) {
            headerRowIdx = i;
            break;
          }
        }

        const headers = aoa[headerRowIdx].map(c => String(c || '').trim());
        const mapping = detectColumns(headers);

        if (mapping.courseCode === -1 && mapping.examDate === -1) {
          throw new Error('Could not identify Course Code or Exam Date columns. Please ensure headers match standard schedule files.');
        }

        const parsedRecords = [];
        for (let r = headerRowIdx + 1; r < aoa.length; r++) {
          const row = aoa[r];
          if (!row || !Array.isArray(row)) continue;

          const rawCourseCode = mapping.courseCode !== -1 ? row[mapping.courseCode] : '';
          const rawExamDate = mapping.examDate !== -1 ? row[mapping.examDate] : '';
          if (!rawCourseCode && !rawExamDate) continue;

          const eventName = mapping.eventName !== -1 ? cleanDisplayField(row[mapping.eventName]) : '';
          const centerCode = mapping.centerCode !== -1 ? cleanDisplayField(row[mapping.centerCode]) : '';
          const centerName = mapping.centerName !== -1 ? cleanDisplayField(row[mapping.centerName]) : '';
          const venueCode = mapping.venueCode !== -1 ? cleanDisplayField(row[mapping.venueCode]) : '';
          const venueName = mapping.venueName !== -1 ? cleanDisplayField(row[mapping.venueName]) : '';
          const programName = mapping.programName !== -1 ? cleanDisplayField(row[mapping.programName]) : '';
          const examDate = formatExamDate(rawExamDate);
          const startTime = mapping.startTime !== -1 ? cleanTimeText(row[mapping.startTime]) : '';
          const endTime = mapping.endTime !== -1 ? cleanTimeText(row[mapping.endTime]) : '';
          const courseCode = cleanDisplayField(rawCourseCode);
          const courseName = mapping.courseName !== -1 ? cleanDisplayField(row[mapping.courseName]) : '';

          let studentCount = 0;
          if (mapping.studentCount !== -1) {
            const rawCount = String(row[mapping.studentCount] || '').replace(/,/g, '').trim();
            const p = parseInt(rawCount, 10);
            if (!isNaN(p)) studentCount = p;
          }

          const venueId = mapping.venueId !== -1 ? cleanDisplayField(row[mapping.venueId]) : '';

          parsedRecords.push({
            eventName,
            centerCode,
            centerName,
            venueCode,
            venueName,
            programName,
            examDate,
            startTime,
            endTime,
            courseCode,
            courseName,
            studentCount,
            venueId
          });
        }

        setRawRows(parsedRecords);
        setCurrentPage(0);
        setStatusAlert({
          type: 'success',
          message: `Successfully loaded ${parsedRecords.length} schedule rows from "${file.name}". Ready for Confisection reporting.`
        });
      } catch (err) {
        console.error('File parsing error:', err);
        setStatusAlert({ type: 'error', message: `Failed to load file: ${err.message}` });
      } finally {
        setIsProcessing(false);
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };

    reader.onerror = () => {
      setStatusAlert({ type: 'error', message: 'Failed to read file from disk.' });
      setIsProcessing(false);
    };

    reader.readAsArrayBuffer(file);
  };

  // Load sample dataset
  const handleLoadSample = () => {
    setRawRows(SAMPLE_CONFISECTION_ROWS);
    setFileName('Sample_Confisection_Data.xlsx');
    setCurrentPage(0);
    setStatusAlert({
      type: 'success',
      message: `Loaded ${SAMPLE_CONFISECTION_ROWS.length} sample Confisection schedule records with Payyannur College and Sir Syed College data.`
    });
  };

  // Helper for Session detection
  const getSession = (startTime, endTime) => {
    const s = `${startTime || ''} ${endTime || ''}`.toUpperCase();
    if (s.includes('AM')) return 'FN';
    if (s.includes('PM')) {
      // 12:00 PM - 01:00 PM is typically Forenoon/Noon, after 1:00 PM is AN
      if (s.includes('09:') || s.includes('10:') || s.includes('11:')) return 'FN';
      return 'AN';
    }
    return 'FN';
  };

  // Consolidated Master Report (Unique Date + Venue + Course Code with summed student counts)
  const consolidatedRows = useMemo(() => {
    if (rawRows.length === 0) return [];

    const groupMap = new Map();

    rawRows.forEach((row, idx) => {
      const vCode = row.venueCode || 'UNKNOWN';
      const cCode = row.courseCode || 'UNKNOWN';
      const date = row.examDate || 'UNKNOWN';
      const time = `${row.startTime}-${row.endTime}`;

      // Composite Key: Date + Venue Code + Course Code
      const key = `${date}|||${vCode}|||${cCode}`;

      const session = getSession(row.startTime, row.endTime);

      if (groupMap.has(key)) {
        const item = groupMap.get(key);
        item.studentCount += row.studentCount;
        if (row.programName && !item.programs.includes(row.programName)) {
          item.programs.push(row.programName);
        }
        if (!item.startTime && row.startTime) item.startTime = row.startTime;
        if (!item.endTime && row.endTime) item.endTime = row.endTime;
      } else {
        groupMap.set(key, {
          id: idx + 1,
          eventName: row.eventName,
          centerCode: row.centerCode,
          centerName: row.centerName,
          venueCode: row.venueCode,
          venueName: row.venueName,
          venueId: row.venueId,
          examDate: row.examDate,
          startTime: row.startTime,
          endTime: row.endTime,
          session,
          courseCode: row.courseCode,
          courseName: row.courseName,
          studentCount: row.studentCount,
          programs: row.programName ? [row.programName] : []
        });
      }
    });

    const list = Array.from(groupMap.values()).map(item => {
      const packets = calculateQpPackets(item.studentCount, bufferRule);
      return {
        ...item,
        programsFormatted: item.programs.join('; '),
        programsCount: item.programs.length,
        ...packets
      };
    });

    // Sort chronologically by date, then venue, then course
    list.sort((a, b) => {
      if (a.examDate !== b.examDate) return a.examDate.localeCompare(b.examDate);
      if (a.venueCode !== b.venueCode) return a.venueCode.localeCompare(b.venueCode);
      return a.courseCode.localeCompare(b.courseCode);
    });

    return list;
  }, [rawRows, bufferRule]);

  // Date-wise & Session-wise Print Order for Confidential Press
  const pressPrintOrderRows = useMemo(() => {
    if (consolidatedRows.length === 0) return [];

    const map = new Map();

    consolidatedRows.forEach(row => {
      // Group by Exam Date + Session + Course Code
      const key = `${row.examDate}|||${row.session}|||${row.courseCode}`;

      if (map.has(key)) {
        const entry = map.get(key);
        entry.totalStudents += row.studentCount;
        entry.totalQp += row.totalQp;
        entry.venuesCount += 1;
        if (!entry.venueCodes.includes(row.venueCode)) {
          entry.venueCodes.push(row.venueCode);
        }
      } else {
        map.set(key, {
          examDate: row.examDate,
          session: row.session,
          startTime: row.startTime,
          endTime: row.endTime,
          courseCode: row.courseCode,
          courseName: row.courseName,
          totalStudents: row.studentCount,
          totalQp: row.totalQp,
          venuesCount: 1,
          venueCodes: [row.venueCode]
        });
      }
    });

    const list = Array.from(map.values());
    list.sort((a, b) => {
      if (a.examDate !== b.examDate) return a.examDate.localeCompare(b.examDate);
      if (a.session !== b.session) return a.session.localeCompare(b.session);
      return a.courseCode.localeCompare(b.courseCode);
    });

    return list;
  }, [consolidatedRows]);

  // Venue-wise Packing Slips
  const venuePackingSlips = useMemo(() => {
    if (consolidatedRows.length === 0) return {};

    const groups = {};

    consolidatedRows.forEach(row => {
      const vKey = row.venueCode || 'UNKNOWN';
      if (!groups[vKey]) {
        groups[vKey] = {
          venueCode: row.venueCode,
          venueName: row.venueName,
          centerCode: row.centerCode,
          centerName: row.centerName,
          venueId: row.venueId,
          totalStudents: 0,
          totalQp: 0,
          items: []
        };
      }

      groups[vKey].totalStudents += row.studentCount;
      groups[vKey].totalQp += row.totalQp;
      groups[vKey].items.push(row);
    });

    return groups;
  }, [consolidatedRows]);

  // Course-wise University Summary
  const courseSummaryRows = useMemo(() => {
    if (consolidatedRows.length === 0) return [];

    const map = new Map();

    consolidatedRows.forEach(row => {
      const key = row.courseCode;

      if (map.has(key)) {
        const item = map.get(key);
        item.totalStudents += row.studentCount;
        item.totalQp += row.totalQp;
        item.venueCount += 1;
        if (!item.dates.includes(row.examDate)) item.dates.push(row.examDate);
      } else {
        map.set(key, {
          courseCode: row.courseCode,
          courseName: row.courseName,
          dates: [row.examDate],
          session: row.session,
          totalStudents: row.studentCount,
          totalQp: row.totalQp,
          venueCount: 1
        });
      }
    });

    const list = Array.from(map.values()).map(x => ({
      ...x,
      datesFormatted: x.dates.join(', ')
    }));

    list.sort((a, b) => b.totalStudents - a.totalStudents);
    return list;
  }, [consolidatedRows]);

  // Distinct filter options
  const filterOptions = useMemo(() => {
    const dates = new Set();
    const venues = new Map();

    consolidatedRows.forEach(r => {
      if (r.examDate) dates.add(r.examDate);
      if (r.venueCode) venues.set(r.venueCode, r.venueName || r.venueCode);
    });

    return {
      dates: Array.from(dates).sort(),
      venues: Array.from(venues.entries()).map(([code, name]) => ({ code, name }))
    };
  }, [consolidatedRows]);

  // Active list according to selected Tab and Filters
  const currentFilteredRows = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();

    const applyCommonFilter = (item) => {
      if (selectedDate !== 'ALL' && item.examDate !== selectedDate) return false;
      if (selectedSession !== 'ALL' && item.session !== selectedSession) return false;
      if (selectedVenue !== 'ALL' && item.venueCode !== selectedVenue) return false;

      if (q) {
        const code = String(item.courseCode || '').toLowerCase();
        const name = String(item.courseName || '').toLowerCase();
        const vCode = String(item.venueCode || '').toLowerCase();
        const vName = String(item.venueName || '').toLowerCase();
        const date = String(item.examDate || '').toLowerCase();
        const prog = String(item.programsFormatted || item.programName || '').toLowerCase();
        return code.includes(q) || name.includes(q) || vCode.includes(q) || vName.includes(q) || date.includes(q) || prog.includes(q);
      }
      return true;
    };

    if (activeTab === 'consolidated') {
      return consolidatedRows.filter(applyCommonFilter);
    }
    if (activeTab === 'press_print') {
      return pressPrintOrderRows.filter(item => {
        if (selectedDate !== 'ALL' && item.examDate !== selectedDate) return false;
        if (selectedSession !== 'ALL' && item.session !== selectedSession) return false;
        if (q) {
          const code = String(item.courseCode || '').toLowerCase();
          const name = String(item.courseName || '').toLowerCase();
          const date = String(item.examDate || '').toLowerCase();
          return code.includes(q) || name.includes(q) || date.includes(q);
        }
        return true;
      });
    }
    if (activeTab === 'course_summary') {
      return courseSummaryRows.filter(item => {
        if (selectedSession !== 'ALL' && item.session !== selectedSession) return false;
        if (q) {
          const code = String(item.courseCode || '').toLowerCase();
          const name = String(item.courseName || '').toLowerCase();
          return code.includes(q) || name.includes(q);
        }
        return true;
      });
    }
    if (activeTab === 'raw') {
      return rawRows.filter(item => {
        if (selectedDate !== 'ALL' && item.examDate !== selectedDate) return false;
        if (selectedVenue !== 'ALL' && item.venueCode !== selectedVenue) return false;
        if (q) {
          const code = String(item.courseCode || '').toLowerCase();
          const name = String(item.courseName || '').toLowerCase();
          const vCode = String(item.venueCode || '').toLowerCase();
          const vName = String(item.venueName || '').toLowerCase();
          return code.includes(q) || name.includes(q) || vCode.includes(q) || vName.includes(q);
        }
        return true;
      });
    }
    return [];
  }, [activeTab, consolidatedRows, pressPrintOrderRows, courseSummaryRows, rawRows, searchQuery, selectedDate, selectedSession, selectedVenue]);

  const paginatedRows = useMemo(() => {
    if (pageSize === 0) return currentFilteredRows;
    const start = currentPage * pageSize;
    return currentFilteredRows.slice(start, start + pageSize);
  }, [currentFilteredRows, currentPage, pageSize]);

  const totalPages = pageSize === 0 ? 1 : Math.ceil(currentFilteredRows.length / pageSize) || 1;

  // Key summary statistics
  const summaryStats = useMemo(() => {
    const totalCandidates = consolidatedRows.reduce((sum, r) => sum + r.studentCount, 0);
    const totalQp = consolidatedRows.reduce((sum, r) => sum + r.totalQp, 0);
    const uniqueExams = consolidatedRows.length;
    const uniqueVenues = new Set(consolidatedRows.map(r => r.venueCode)).size;
    const uniqueCourses = new Set(consolidatedRows.map(r => r.courseCode)).size;
    const uniqueDates = new Set(consolidatedRows.map(r => r.examDate)).size;

    return {
      totalCandidates,
      totalQp,
      uniqueExams,
      uniqueVenues,
      uniqueCourses,
      uniqueDates
    };
  }, [consolidatedRows]);

  // Export to Excel with 4 dedicated sheets
  const handleExportExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Consolidated Confisection Master Schedule
      const sheet1Headers = [
        'Event Name',
        'Center Code',
        'Center Name',
        'Venue Code',
        'Venue Name',
        'venueId',
        'Exam Date',
        'Session',
        'Start Time',
        'End Time',
        'Course Code',
        'Course Name',
        'Combined Programs',
        'Student Count',
        'Packets of 20',
        'Packets of 10',
        'Packets of 5',
        'Loose / Singles',
        'Buffer Copies',
        'Total QP Required'
      ];

      const sheet1Data = consolidatedRows.map(r => [
        r.eventName,
        r.centerCode,
        r.centerName,
        r.venueCode,
        r.venueName,
        r.venueId,
        r.examDate,
        r.session,
        r.startTime,
        r.endTime,
        r.courseCode,
        r.courseName,
        r.programsFormatted,
        r.studentCount,
        r.p20,
        r.p10,
        r.p5,
        r.singles,
        r.buffer,
        r.totalQp
      ]);

      const ws1 = XLSX.utils.aoa_to_sheet([sheet1Headers, ...sheet1Data]);
      ws1['!cols'] = [
        { wch: 25 }, { wch: 12 }, { wch: 20 }, { wch: 12 }, { wch: 35 },
        { wch: 10 }, { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 12 },
        { wch: 16 }, { wch: 38 }, { wch: 45 }, { wch: 14 }, { wch: 14 },
        { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 16 }
      ];
      XLSX.utils.book_append_sheet(wb, ws1, 'Confi_Master_Schedule');

      // Sheet 2: Daily Press Print Order
      const sheet2Headers = [
        'Exam Date',
        'Session',
        'Start Time',
        'End Time',
        'Course Code',
        'Course Name',
        'Total Venues',
        'Total Candidates',
        'Total QP Print Order',
        'Venues List'
      ];
      const sheet2Data = pressPrintOrderRows.map(r => [
        r.examDate,
        r.session,
        r.startTime,
        r.endTime,
        r.courseCode,
        r.courseName,
        r.venuesCount,
        r.totalStudents,
        r.totalQp,
        r.venueCodes.join(', ')
      ]);
      const ws2 = XLSX.utils.aoa_to_sheet([sheet2Headers, ...sheet2Data]);
      ws2['!cols'] = [
        { wch: 14 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 16 },
        { wch: 38 }, { wch: 14 }, { wch: 16 }, { wch: 20 }, { wch: 30 }
      ];
      XLSX.utils.book_append_sheet(wb, ws2, 'Daily_Press_Print_Order');

      // Sheet 3: Venue Packing Slip Summary
      const sheet3Headers = [
        'Venue Code',
        'Venue Name',
        'Center Code',
        'Exam Date',
        'Session',
        'Course Code',
        'Course Name',
        'Student Count',
        '20s',
        '10s',
        '5s',
        'Singles',
        'Total QP Packed'
      ];
      const sheet3Data = [];
      Object.values(venuePackingSlips).forEach(v => {
        v.items.forEach(item => {
          sheet3Data.push([
            v.venueCode,
            v.venueName,
            v.centerCode,
            item.examDate,
            item.session,
            item.courseCode,
            item.courseName,
            item.studentCount,
            item.p20,
            item.p10,
            item.p5,
            item.singles,
            item.totalQp
          ]);
        });
      });
      const ws3 = XLSX.utils.aoa_to_sheet([sheet3Headers, ...sheet3Data]);
      ws3['!cols'] = [
        { wch: 12 }, { wch: 35 }, { wch: 12 }, { wch: 14 }, { wch: 10 },
        { wch: 16 }, { wch: 38 }, { wch: 14 }, { wch: 8 }, { wch: 8 },
        { wch: 8 }, { wch: 10 }, { wch: 16 }
      ];
      XLSX.utils.book_append_sheet(wb, ws3, 'Venue_Packing_Slips');

      // Sheet 4: Course Master Summary
      const sheet4Headers = [
        'Course Code',
        'Course Name',
        'Scheduled Dates',
        'Session',
        'Total Venues',
        'Total Candidates',
        'Total QP Needed'
      ];
      const sheet4Data = courseSummaryRows.map(r => [
        r.courseCode,
        r.courseName,
        r.datesFormatted,
        r.session,
        r.venueCount,
        r.totalStudents,
        r.totalQp
      ]);
      const ws4 = XLSX.utils.aoa_to_sheet([sheet4Headers, ...sheet4Data]);
      ws4['!cols'] = [
        { wch: 16 }, { wch: 40 }, { wch: 25 }, { wch: 10 }, { wch: 14 }, { wch: 16 }, { wch: 16 }
      ];
      XLSX.utils.book_append_sheet(wb, ws4, 'Course_Master_Summary');

      XLSX.writeFile(wb, 'Confisection_Exam_Report.xlsx');
      setStatusAlert({
        type: 'success',
        message: 'Downloaded Confisection_Exam_Report.xlsx with 4 structured confidential reports.'
      });
    } catch (err) {
      console.error('Excel export error:', err);
      setStatusAlert({ type: 'error', message: `Export failed: ${err.message}` });
    }
  };

  // Download template
  const handleDownloadTemplate = () => {
    const headers = [
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
    const sampleRows = SAMPLE_CONFISECTION_ROWS.map(r => [
      r.eventName,
      r.centerCode,
      r.centerName,
      r.venueCode,
      r.venueName,
      r.programName,
      r.examDate,
      r.startTime,
      r.endTime,
      r.courseCode,
      r.courseName,
      r.studentCount,
      r.venueId
    ]);

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = headers.map(() => ({ wch: 22 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Confisection_Source_Template');
    XLSX.writeFile(wb, 'Confisection_Exam_Source_Template.xlsx');

    setStatusAlert({
      type: 'success',
      message: 'Downloaded Confisection_Exam_Source_Template.xlsx with all 13 standard headers.'
    });
  };

  // Print Dispatch / Packing Slips
  const handlePrintSlips = () => {
    window.print();
  };

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-100 flex flex-col font-sans transition-colors duration-150">
      {/* Top Header */}
      <header className="bg-white dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 sticky top-0 z-20 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <Link
              to="/"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700/60 text-xs font-medium transition-colors"
            >
              <ArrowLeft className="w-3.5 h-3.5" /> Back to Dashboard
            </Link>
            <div className="h-4 w-px bg-slate-300 dark:bg-slate-600 hidden sm:block" />
            <div>
              <div className="flex items-center gap-2">
                <span className="p-1 rounded-md bg-rose-50 dark:bg-rose-950/60 border border-rose-200 dark:border-rose-800 text-rose-600 dark:text-rose-400">
                  <ShieldCheck className="w-4 h-4" />
                </span>
                <h1 className="text-base sm:text-lg font-bold text-slate-900 dark:text-white">
                  Confisection Exam Report
                </h1>
                <span className="text-[10px] font-semibold tracking-wider uppercase px-2 py-0.5 rounded-full bg-rose-100 text-rose-800 dark:bg-rose-900/60 dark:text-rose-300 border border-rose-300 dark:border-rose-700">
                  Confidential Section
                </span>
              </div>
              <p className="text-xs text-slate-500 dark:text-slate-400 hidden md:block">
                Aggregate schedules, calculate question paper printing requirements, and generate venue packing slips.
              </p>
            </div>
          </div>

          <div className="flex items-center flex-wrap gap-2">
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
            />

            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={isProcessing}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold shadow-xs transition-colors disabled:opacity-50"
              title="Upload Center Venue Date Wise Schedule spreadsheet"
            >
              <Upload className="w-3.5 h-3.5" />
              Upload Schedule (.xlsx)
            </button>

            <button
              onClick={handleLoadSample}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-medium transition-colors"
              title="Load sample dataset for III Sem(R) November 2026"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-500" />
              Load Sample
            </button>

            <button
              onClick={handleDownloadTemplate}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-medium transition-colors"
              title="Download 13-Column Template"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
              Template
            </button>

            <button
              onClick={handleExportExcel}
              disabled={consolidatedRows.length === 0}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-xs transition-colors disabled:opacity-40"
              title="Export all 4 confidential reports to Excel"
            >
              <Download className="w-3.5 h-3.5" />
              Export Report (.xlsx)
            </button>

            <button
              onClick={handlePrintSlips}
              disabled={consolidatedRows.length === 0}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-medium transition-colors disabled:opacity-40"
              title="Print Packing Slips"
            >
              <Printer className="w-3.5 h-3.5 text-slate-500" />
              Print
            </button>

            <button
              onClick={() => setShowGuideModal(true)}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-medium transition-colors"
            >
              <HelpCircle className="w-3.5 h-3.5 text-sky-500" />
              Guide
            </button>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 w-full flex-1 flex flex-col gap-6">
        {/* Status Alert */}
        {statusAlert && (
          <div
            className={`p-3.5 rounded-xl border flex items-start justify-between gap-3 text-sm shadow-xs ${
              statusAlert.type === 'success'
                ? 'bg-emerald-50 dark:bg-emerald-950/40 border-emerald-200 dark:border-emerald-800 text-emerald-900 dark:text-emerald-200'
                : statusAlert.type === 'warning'
                ? 'bg-amber-50 dark:bg-amber-950/40 border-amber-200 dark:border-amber-800 text-amber-900 dark:text-amber-200'
                : statusAlert.type === 'error'
                ? 'bg-rose-50 dark:bg-rose-950/40 border-rose-200 dark:border-rose-800 text-rose-900 dark:text-rose-200'
                : 'bg-sky-50 dark:bg-sky-950/40 border-sky-200 dark:border-sky-800 text-sky-900 dark:text-sky-200'
            }`}
          >
            <div className="flex items-center gap-2">
              {statusAlert.type === 'success' && <CheckCircle2 className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0" />}
              {statusAlert.type === 'warning' && <AlertCircle className="w-4 h-4 text-amber-600 dark:text-amber-400 shrink-0" />}
              {statusAlert.type === 'error' && <AlertCircle className="w-4 h-4 text-rose-600 dark:text-rose-400 shrink-0" />}
              {statusAlert.type === 'info' && <Sparkles className="w-4 h-4 text-sky-600 dark:text-sky-400 shrink-0" />}
              <span>{statusAlert.message}</span>
            </div>
            <button
              onClick={() => setStatusAlert(null)}
              className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 p-0.5"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* Hero Statistics */}
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3 sm:gap-4">
          <div className="p-3.5 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3">
            <div className="p-2 rounded-lg bg-rose-50 dark:bg-rose-950/50 text-rose-600 dark:text-rose-400 border border-rose-200 dark:border-rose-800">
              <Users className="w-4 h-4" />
            </div>
            <div>
              <p className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Candidates</p>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white">{summaryStats.totalCandidates.toLocaleString()}</h3>
            </div>
          </div>

          <div className="p-3.5 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3">
            <div className="p-2 rounded-lg bg-emerald-50 dark:bg-emerald-950/50 text-emerald-600 dark:text-emerald-400 border border-emerald-200 dark:border-emerald-800">
              <PackageCheck className="w-4 h-4" />
            </div>
            <div>
              <p className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Total QP Print</p>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white">{summaryStats.totalQp.toLocaleString()}</h3>
            </div>
          </div>

          <div className="p-3.5 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3">
            <div className="p-2 rounded-lg bg-teal-50 dark:bg-teal-950/50 text-teal-600 dark:text-teal-400 border border-teal-200 dark:border-teal-800">
              <Layers className="w-4 h-4" />
            </div>
            <div>
              <p className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Unique Exams</p>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white">{summaryStats.uniqueExams.toLocaleString()}</h3>
            </div>
          </div>

          <div className="p-3.5 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3">
            <div className="p-2 rounded-lg bg-indigo-50 dark:bg-indigo-950/50 text-indigo-600 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-800">
              <Building2 className="w-4 h-4" />
            </div>
            <div>
              <p className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Venues</p>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white">{summaryStats.uniqueVenues}</h3>
            </div>
          </div>

          <div className="p-3.5 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3">
            <div className="p-2 rounded-lg bg-sky-50 dark:bg-sky-950/50 text-sky-600 dark:text-sky-400 border border-sky-200 dark:border-sky-800">
              <BookOpen className="w-4 h-4" />
            </div>
            <div>
              <p className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Courses</p>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white">{summaryStats.uniqueCourses}</h3>
            </div>
          </div>

          <div className="p-3.5 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3">
            <div className="p-2 rounded-lg bg-amber-50 dark:bg-amber-950/50 text-amber-600 dark:text-amber-400 border border-amber-200 dark:border-amber-800">
              <CalendarDays className="w-4 h-4" />
            </div>
            <div>
              <p className="text-[11px] font-medium text-slate-400 uppercase tracking-wider">Exam Dates</p>
              <h3 className="text-lg font-bold text-slate-900 dark:text-white">{summaryStats.uniqueDates}</h3>
            </div>
          </div>
        </div>

        {/* View Mode Tabs & Controls Container */}
        <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-xs overflow-hidden flex flex-col">
          {/* Main Navigation Tabs */}
          <div className="border-b border-slate-200 dark:border-slate-700 bg-slate-50/70 dark:bg-slate-800/80 px-4 pt-3 flex items-center gap-1.5 overflow-x-auto scrollbar-thin">
            <button
              onClick={() => { setActiveTab('consolidated'); setCurrentPage(0); }}
              className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-semibold rounded-t-lg transition-all border-b-2 whitespace-nowrap shrink-0 ${
                activeTab === 'consolidated'
                  ? 'border-rose-600 text-rose-600 dark:text-rose-400 bg-white dark:bg-slate-800 shadow-2xs'
                  : 'border-transparent text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200 hover:bg-slate-100/60'
              }`}
            >
              <PackageCheck className="w-3.5 h-3.5" />
              <span>Consolidated Master Schedule</span>
              <span className="px-1.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300">
                {consolidatedRows.length}
              </span>
            </button>

            <button
              onClick={() => { setActiveTab('press_print'); setCurrentPage(0); }}
              className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-semibold rounded-t-lg transition-all border-b-2 whitespace-nowrap shrink-0 ${
                activeTab === 'press_print'
                  ? 'border-rose-600 text-rose-600 dark:text-rose-400 bg-white dark:bg-slate-800 shadow-2xs'
                  : 'border-transparent text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200 hover:bg-slate-100/60'
              }`}
            >
              <Printer className="w-3.5 h-3.5" />
              <span>Daily Press Print Order</span>
              <span className="px-1.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300">
                {pressPrintOrderRows.length}
              </span>
            </button>

            <button
              onClick={() => { setActiveTab('venue_packing'); setCurrentPage(0); }}
              className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-semibold rounded-t-lg transition-all border-b-2 whitespace-nowrap shrink-0 ${
                activeTab === 'venue_packing'
                  ? 'border-rose-600 text-rose-600 dark:text-rose-400 bg-white dark:bg-slate-800 shadow-2xs'
                  : 'border-transparent text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200 hover:bg-slate-100/60'
              }`}
            >
              <Building2 className="w-3.5 h-3.5" />
              <span>Venue Packing Slips</span>
              <span className="px-1.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300">
                {Object.keys(venuePackingSlips).length} Venues
              </span>
            </button>

            <button
              onClick={() => { setActiveTab('course_summary'); setCurrentPage(0); }}
              className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-semibold rounded-t-lg transition-all border-b-2 whitespace-nowrap shrink-0 ${
                activeTab === 'course_summary'
                  ? 'border-rose-600 text-rose-600 dark:text-rose-400 bg-white dark:bg-slate-800 shadow-2xs'
                  : 'border-transparent text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200 hover:bg-slate-100/60'
              }`}
            >
              <BookOpen className="w-3.5 h-3.5" />
              <span>Course Master Summary</span>
              <span className="px-1.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300">
                {courseSummaryRows.length}
              </span>
            </button>

            <button
              onClick={() => { setActiveTab('raw'); setCurrentPage(0); }}
              className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-semibold rounded-t-lg transition-all border-b-2 whitespace-nowrap shrink-0 ${
                activeTab === 'raw'
                  ? 'border-rose-600 text-rose-600 dark:text-rose-400 bg-white dark:bg-slate-800 shadow-2xs'
                  : 'border-transparent text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200 hover:bg-slate-100/60'
              }`}
            >
              <FileText className="w-3.5 h-3.5" />
              <span>Raw Source Rows</span>
              <span className="px-1.5 py-0.5 rounded-full text-[10px] font-bold bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300">
                {rawRows.length}
              </span>
            </button>
          </div>

          {/* Filtering and Configuration Toolbar */}
          <div className="p-4 bg-slate-50/50 dark:bg-slate-800/50 border-b border-slate-200 dark:border-slate-700 flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-center flex-wrap gap-2.5 flex-1 min-w-[280px]">
              {/* Live Search */}
              <div className="relative w-full sm:w-64">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(0); }}
                  placeholder="Search code, course, venue, date..."
                  className="w-full pl-9 pr-8 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-rose-500"
                />
                {searchQuery && (
                  <button
                    onClick={() => setSearchQuery('')}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>

              {/* Date Filter */}
              <select
                value={selectedDate}
                onChange={(e) => { setSelectedDate(e.target.value); setCurrentPage(0); }}
                className="px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-medium focus:outline-none"
              >
                <option value="ALL">All Exam Dates ({filterOptions.dates.length})</option>
                {filterOptions.dates.map(d => (
                  <option key={d} value={d}>{d}</option>
                ))}
              </select>

              {/* Session Filter */}
              <select
                value={selectedSession}
                onChange={(e) => { setSelectedSession(e.target.value); setCurrentPage(0); }}
                className="px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-medium focus:outline-none"
              >
                <option value="ALL">All Sessions</option>
                <option value="FN">Forenoon (FN)</option>
                <option value="AN">Afternoon (AN)</option>
              </select>

              {/* Venue Filter */}
              {activeTab !== 'press_print' && activeTab !== 'course_summary' && (
                <select
                  value={selectedVenue}
                  onChange={(e) => { setSelectedVenue(e.target.value); setCurrentPage(0); }}
                  className="px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-medium focus:outline-none max-w-[220px] truncate"
                >
                  <option value="ALL">All Venues ({filterOptions.venues.length})</option>
                  {filterOptions.venues.map(v => (
                    <option key={v.code} value={v.code}>{v.code} - {v.name}</option>
                  ))}
                </select>
              )}
            </div>

            {/* Buffer policy and Page Size Controls */}
            <div className="flex items-center gap-2">
              <div className="flex items-center gap-1.5 bg-rose-50 dark:bg-rose-950/40 px-2.5 py-1 rounded-lg border border-rose-200 dark:border-rose-900/60 text-xs">
                <span className="font-semibold text-rose-800 dark:text-rose-300 text-[11px] shrink-0">Buffer Rule:</span>
                <select
                  value={bufferRule}
                  onChange={(e) => setBufferRule(e.target.value)}
                  className="bg-transparent text-rose-900 dark:text-rose-200 font-semibold focus:outline-none text-[11px] cursor-pointer"
                >
                  <option value="plus1">+1 Copy / Venue (Standard)</option>
                  <option value="plus2">+2 Copies / Venue (High Reserve)</option>
                  <option value="five_percent">+5% Round Up</option>
                  <option value="none">Exact Takers (No Buffer)</option>
                </select>
              </div>

              {activeTab !== 'venue_packing' && (
                <select
                  value={pageSize}
                  onChange={(e) => { setPageSize(Number(e.target.value)); setCurrentPage(0); }}
                  className="px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-medium focus:outline-none"
                >
                  <option value={20}>20 rows</option>
                  <option value={50}>50 rows</option>
                  <option value={100}>100 rows</option>
                  <option value={0}>Show All</option>
                </select>
              )}
            </div>
          </div>

          {/* Table / Content View */}
          {activeTab === 'venue_packing' ? (
            /* Venue-Wise Packing Slips View */
            <div className="p-4 sm:p-6 space-y-6 overflow-y-auto max-h-[70vh]">
              {Object.keys(venuePackingSlips).length === 0 ? (
                <div className="text-center py-12 text-slate-400">
                  <PackageCheck className="w-8 h-8 mx-auto mb-2 opacity-40" />
                  <p className="text-sm font-medium">No venue packing data available. Upload an Excel file or load sample data.</p>
                </div>
              ) : (
                Object.values(venuePackingSlips)
                  .filter(v => selectedVenue === 'ALL' || v.venueCode === selectedVenue)
                  .map((venueGroup, vIdx) => (
                    <div
                      key={vIdx}
                      className="border border-slate-200 dark:border-slate-700 rounded-xl overflow-hidden shadow-xs bg-white dark:bg-slate-800"
                    >
                      <div className="p-3.5 bg-slate-100/80 dark:bg-slate-800/90 border-b border-slate-200 dark:border-slate-700 flex flex-wrap items-center justify-between gap-2">
                        <div className="flex items-center gap-2.5">
                          <span className="px-2.5 py-1 rounded-md bg-indigo-100 text-indigo-800 dark:bg-indigo-950 dark:text-indigo-300 font-mono font-bold text-xs">
                            {venueGroup.venueCode}
                          </span>
                          <div>
                            <h3 className="text-sm font-bold text-slate-900 dark:text-white">
                              {venueGroup.venueName}
                            </h3>
                            <p className="text-[11px] text-slate-500">
                              Center: {venueGroup.centerCode} ({venueGroup.centerName || 'Main'}) • Venue ID: {venueGroup.venueId || 'N/A'}
                            </p>
                          </div>
                        </div>

                        <div className="flex items-center gap-3 text-xs">
                          <span className="font-medium text-slate-600 dark:text-slate-300">
                            Exams: <strong>{venueGroup.items.length}</strong>
                          </span>
                          <span className="font-medium text-slate-600 dark:text-slate-300">
                            Takers: <strong>{venueGroup.totalStudents}</strong>
                          </span>
                          <span className="px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-800 dark:bg-emerald-950 dark:text-emerald-300 font-bold">
                            Total QP Packed: {venueGroup.totalQp}
                          </span>
                        </div>
                      </div>

                      <div className="overflow-x-auto">
                        <table className="w-full text-left text-xs border-collapse">
                          <thead>
                            <tr className="bg-slate-50/70 dark:bg-slate-800/50 text-slate-500 font-semibold uppercase tracking-wider text-[10px] border-b border-slate-200 dark:border-slate-700">
                              <th className="py-2 px-3">Date</th>
                              <th className="py-2 px-3">Session</th>
                              <th className="py-2 px-3">Course Code</th>
                              <th className="py-2 px-3">Course Name</th>
                              <th className="py-2 px-3 text-center">Takers</th>
                              <th className="py-2 px-3 text-center">20s</th>
                              <th className="py-2 px-3 text-center">10s</th>
                              <th className="py-2 px-3 text-center">5s</th>
                              <th className="py-2 px-3 text-center">Loose</th>
                              <th className="py-2 px-3 text-right">Total QP</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60">
                            {venueGroup.items.map((item, idx) => (
                              <tr key={idx} className="hover:bg-slate-50/50 dark:hover:bg-slate-700/30">
                                <td className="py-2 px-3 font-medium text-slate-800 dark:text-slate-200 whitespace-nowrap">
                                  {item.examDate}
                                </td>
                                <td className="py-2 px-3 whitespace-nowrap">
                                  <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                    item.session === 'FN'
                                      ? 'bg-sky-100 text-sky-800 dark:bg-sky-950 dark:text-sky-300'
                                      : 'bg-indigo-100 text-indigo-800 dark:bg-indigo-950 dark:text-indigo-300'
                                  }`}>
                                    {item.session} ({item.startTime})
                                  </span>
                                </td>
                                <td className="py-2 px-3 font-mono font-semibold text-emerald-600 dark:text-emerald-400">
                                  {item.courseCode}
                                </td>
                                <td className="py-2 px-3 font-medium text-slate-800 dark:text-slate-100">
                                  {item.courseName}
                                </td>
                                <td className="py-2 px-3 text-center font-bold text-slate-900 dark:text-white">
                                  {item.studentCount}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-600 dark:text-slate-300">
                                  {item.p20 > 0 ? <span className="font-bold text-emerald-600">{item.p20}</span> : '-'}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-600 dark:text-slate-300">
                                  {item.p10 > 0 ? <span className="font-bold text-teal-600">{item.p10}</span> : '-'}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-600 dark:text-slate-300">
                                  {item.p5 > 0 ? <span className="font-bold text-sky-600">{item.p5}</span> : '-'}
                                </td>
                                <td className="py-2 px-3 text-center font-mono text-slate-600 dark:text-slate-300">
                                  {item.singles > 0 ? <span className="font-bold text-amber-600">{item.singles}</span> : '-'}
                                </td>
                                <td className="py-2 px-3 text-right font-bold text-emerald-600 dark:text-emerald-400 font-mono">
                                  {item.totalQp}
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  ))
              )}
            </div>
          ) : (
            /* Tabular Views: Consolidated, Press Print, Course Summary, Raw */
            <div className="overflow-x-auto flex-1">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-100/80 dark:bg-slate-800/80 text-slate-600 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700 font-semibold uppercase tracking-wider text-[10px]">
                    <th className="py-2.5 px-3 text-center w-12">#</th>
                    
                    {activeTab === 'consolidated' && (
                      <>
                        <th className="py-2.5 px-3 whitespace-nowrap">Exam Date</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Session</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Venue</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Course Code</th>
                        <th className="py-2.5 px-3 min-w-[200px]">Course Name</th>
                        <th className="py-2.5 px-3 min-w-[220px]">Program(s)</th>
                        <th className="py-2.5 px-3 text-center whitespace-nowrap">Takers</th>
                        <th className="py-2.5 px-3 text-center whitespace-nowrap">Packets (20/10/5/1)</th>
                        <th className="py-2.5 px-3 text-right whitespace-nowrap">Total QP</th>
                      </>
                    )}

                    {activeTab === 'press_print' && (
                      <>
                        <th className="py-2.5 px-3 whitespace-nowrap">Exam Date</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Session</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Timing</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Course Code</th>
                        <th className="py-2.5 px-3 min-w-[260px]">Course Name</th>
                        <th className="py-2.5 px-3 text-center whitespace-nowrap">Venues</th>
                        <th className="py-2.5 px-3 text-center whitespace-nowrap">Total Candidates</th>
                        <th className="py-2.5 px-3 text-right whitespace-nowrap">QP Print Requirement</th>
                      </>
                    )}

                    {activeTab === 'course_summary' && (
                      <>
                        <th className="py-2.5 px-3 whitespace-nowrap">Course Code</th>
                        <th className="py-2.5 px-3 min-w-[280px]">Course Name</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Scheduled Dates</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Session</th>
                        <th className="py-2.5 px-3 text-center whitespace-nowrap">Venues Count</th>
                        <th className="py-2.5 px-3 text-center whitespace-nowrap">Total Students</th>
                        <th className="py-2.5 px-3 text-right whitespace-nowrap">Total QP</th>
                      </>
                    )}

                    {activeTab === 'raw' && (
                      <>
                        <th className="py-2.5 px-3 whitespace-nowrap">Event Name</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Venue</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Program Name</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Date</th>
                        <th className="py-2.5 px-3 whitespace-nowrap">Course Code</th>
                        <th className="py-2.5 px-3 min-w-[200px]">Course Name</th>
                        <th className="py-2.5 px-3 text-right whitespace-nowrap">Count</th>
                      </>
                    )}
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60">
                  {paginatedRows.length === 0 ? (
                    <tr>
                      <td colSpan={10} className="py-12 text-center text-slate-400">
                        <div className="flex flex-col items-center justify-center gap-2">
                          <PackageCheck className="w-8 h-8 opacity-40 text-slate-400" />
                          <p className="text-sm font-medium">
                            {searchQuery ? `No records matching "${searchQuery}"` : 'No examination schedule loaded yet.'}
                          </p>
                          <p className="text-xs text-slate-400">
                            Upload your Center Venue Date Wise Schedule file (.xlsx) or click "Load Sample" above.
                          </p>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    paginatedRows.map((row, idx) => {
                      const rowNum = (pageSize === 0 ? 0 : currentPage * pageSize) + idx + 1;

                      return (
                        <tr key={idx} className="hover:bg-slate-50/80 dark:hover:bg-slate-700/40 transition-colors">
                          <td className="py-2 px-3 text-center font-mono text-[11px] text-slate-400">
                            {rowNum}
                          </td>

                          {activeTab === 'consolidated' && (
                            <>
                              <td className="py-2 px-3 font-semibold text-slate-900 dark:text-white whitespace-nowrap">
                                {row.examDate}
                              </td>
                              <td className="py-2 px-3 whitespace-nowrap">
                                <span className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                  row.session === 'FN'
                                    ? 'bg-sky-100 text-sky-800 dark:bg-sky-950 dark:text-sky-300'
                                    : 'bg-indigo-100 text-indigo-800 dark:bg-indigo-950 dark:text-indigo-300'
                                }`}>
                                  {row.session} ({row.startTime})
                                </span>
                              </td>
                              <td className="py-2 px-3 whitespace-nowrap">
                                <span className="font-semibold text-slate-800 dark:text-slate-100">
                                  {row.venueCode}
                                </span>
                                <span className="text-[11px] text-slate-500 block max-w-[180px] truncate" title={row.venueName}>
                                  {row.venueName}
                                </span>
                              </td>
                              <td className="py-2 px-3 font-mono font-bold text-emerald-600 dark:text-emerald-400 whitespace-nowrap">
                                {row.courseCode}
                              </td>
                              <td className="py-2 px-3 font-medium text-slate-800 dark:text-slate-100">
                                {row.courseName}
                              </td>
                              <td className="py-2 px-3 text-[11px] text-slate-500 leading-relaxed">
                                {row.programsFormatted || 'General'}
                              </td>
                              <td className="py-2 px-3 text-center font-bold text-slate-900 dark:text-white">
                                {row.studentCount}
                              </td>
                              <td className="py-2 px-3 text-center whitespace-nowrap">
                                <div className="inline-flex items-center gap-1 font-mono text-[10px]">
                                  <span className="px-1 py-0.5 rounded bg-emerald-50 dark:bg-emerald-950/60 border border-emerald-200 text-emerald-700 dark:text-emerald-300" title="Packets of 20">
                                    20s: <strong>{row.p20}</strong>
                                  </span>
                                  <span className="px-1 py-0.5 rounded bg-teal-50 dark:bg-teal-950/60 border border-teal-200 text-teal-700 dark:text-teal-300" title="Packets of 10">
                                    10s: <strong>{row.p10}</strong>
                                  </span>
                                  <span className="px-1 py-0.5 rounded bg-sky-50 dark:bg-sky-950/60 border border-sky-200 text-sky-700 dark:text-sky-300" title="Packets of 5">
                                    5s: <strong>{row.p5}</strong>
                                  </span>
                                  {row.singles > 0 && (
                                    <span className="px-1 py-0.5 rounded bg-amber-50 dark:bg-amber-950/60 border border-amber-200 text-amber-700 dark:text-amber-300" title="Loose singles">
                                      L: <strong>{row.singles}</strong>
                                    </span>
                                  )}
                                </div>
                              </td>
                              <td className="py-2 px-3 text-right font-mono font-bold text-emerald-600 dark:text-emerald-400">
                                {row.totalQp}
                              </td>
                            </>
                          )}

                          {activeTab === 'press_print' && (
                            <>
                              <td className="py-2 px-3 font-semibold text-slate-900 dark:text-white whitespace-nowrap">
                                {row.examDate}
                              </td>
                              <td className="py-2 px-3 whitespace-nowrap">
                                <span className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                  row.session === 'FN'
                                    ? 'bg-sky-100 text-sky-800 dark:bg-sky-950 dark:text-sky-300'
                                    : 'bg-indigo-100 text-indigo-800 dark:bg-indigo-950 dark:text-indigo-300'
                                }`}>
                                  {row.session}
                                </span>
                              </td>
                              <td className="py-2 px-3 text-slate-600 dark:text-slate-300 whitespace-nowrap">
                                {row.startTime} - {row.endTime}
                              </td>
                              <td className="py-2 px-3 font-mono font-bold text-emerald-600 dark:text-emerald-400 whitespace-nowrap">
                                {row.courseCode}
                              </td>
                              <td className="py-2 px-3 font-medium text-slate-800 dark:text-slate-100">
                                {row.courseName}
                              </td>
                              <td className="py-2 px-3 text-center font-bold text-slate-700 dark:text-slate-300">
                                {row.venuesCount}
                              </td>
                              <td className="py-2 px-3 text-center font-bold text-slate-900 dark:text-white">
                                {row.totalStudents}
                              </td>
                              <td className="py-2 px-3 text-right font-mono font-bold text-emerald-600 dark:text-emerald-400">
                                {row.totalQp}
                              </td>
                            </>
                          )}

                          {activeTab === 'course_summary' && (
                            <>
                              <td className="py-2 px-3 font-mono font-bold text-emerald-600 dark:text-emerald-400 whitespace-nowrap">
                                {row.courseCode}
                              </td>
                              <td className="py-2 px-3 font-medium text-slate-800 dark:text-slate-100">
                                {row.courseName}
                              </td>
                              <td className="py-2 px-3 font-semibold text-slate-700 dark:text-slate-300 whitespace-nowrap">
                                {row.datesFormatted}
                              </td>
                              <td className="py-2 px-3 whitespace-nowrap">
                                <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-100 dark:bg-slate-700">
                                  {row.session}
                                </span>
                              </td>
                              <td className="py-2 px-3 text-center font-bold text-slate-700 dark:text-slate-300">
                                {row.venueCount}
                              </td>
                              <td className="py-2 px-3 text-center font-bold text-slate-900 dark:text-white">
                                {row.totalStudents}
                              </td>
                              <td className="py-2 px-3 text-right font-mono font-bold text-emerald-600 dark:text-emerald-400">
                                {row.totalQp}
                              </td>
                            </>
                          )}

                          {activeTab === 'raw' && (
                            <>
                              <td className="py-2 px-3 text-slate-600 dark:text-slate-300 whitespace-nowrap">
                                {row.eventName}
                              </td>
                              <td className="py-2 px-3 whitespace-nowrap">
                                <strong>{row.venueCode}</strong> - {row.venueName}
                              </td>
                              <td className="py-2 px-3 text-[11px] text-slate-500 max-w-xs truncate" title={row.programName}>
                                {row.programName}
                              </td>
                              <td className="py-2 px-3 whitespace-nowrap">
                                {row.examDate}
                              </td>
                              <td className="py-2 px-3 font-mono font-bold text-emerald-600 whitespace-nowrap">
                                {row.courseCode}
                              </td>
                              <td className="py-2 px-3 font-medium text-slate-800 dark:text-slate-100">
                                {row.courseName}
                              </td>
                              <td className="py-2 px-3 text-right font-bold text-slate-900 dark:text-white">
                                {row.studentCount}
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          )}

          {/* Pagination Footer */}
          {activeTab !== 'venue_packing' && pageSize > 0 && totalPages > 1 && (
            <div className="p-3 bg-slate-50/70 dark:bg-slate-800/70 border-t border-slate-200 dark:border-slate-700 flex items-center justify-between text-xs text-slate-500">
              <span>
                Showing <strong>{paginatedRows.length}</strong> of <strong>{currentFilteredRows.length}</strong> records (Page {currentPage + 1} of {totalPages})
              </span>
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setCurrentPage(p => Math.max(0, p - 1))}
                  disabled={currentPage === 0}
                  className="px-2.5 py-1 rounded border border-slate-300 dark:border-slate-600 disabled:opacity-40 hover:bg-white dark:hover:bg-slate-700 flex items-center gap-1"
                >
                  <ChevronLeft className="w-3.5 h-3.5" /> Prev
                </button>
                <button
                  onClick={() => setCurrentPage(p => Math.min(totalPages - 1, p + 1))}
                  disabled={currentPage >= totalPages - 1}
                  className="px-2.5 py-1 rounded border border-slate-300 dark:border-slate-600 disabled:opacity-40 hover:bg-white dark:hover:bg-slate-700 flex items-center gap-1"
                >
                  Next <ChevronRight className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* Guide Modal */}
      {showGuideModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-800 rounded-2xl max-w-3xl w-full border border-slate-200 dark:border-slate-700 shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
            <div className="p-4 sm:p-5 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="p-1.5 rounded-lg bg-rose-50 dark:bg-rose-950/60 text-rose-600 dark:text-rose-400">
                  <HelpCircle className="w-5 h-5" />
                </span>
                <h3 className="text-base font-bold text-slate-900 dark:text-white">
                  Confisection Exam Report - Extraction Logic &amp; Guide
                </h3>
              </div>
              <button
                onClick={() => setShowGuideModal(false)}
                className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 sm:p-6 overflow-y-auto space-y-5 text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
              <div>
                <h4 className="text-sm font-bold text-slate-800 dark:text-slate-100 mb-1">
                  1. The 13 Source File Headers
                </h4>
                <p>The tool ingests examination schedule files containing the exact 13 university standard headers:</p>
                <div className="p-2.5 rounded-lg bg-slate-100 dark:bg-slate-900 font-mono text-[11px] overflow-x-auto mt-2 text-slate-700 dark:text-slate-300">
                  Event Name | Center Code | Center Name | Venue Code | Venue Name | Program Name | Exam Date | Course Start Time | Course End Time | Course Code | Course Name | Student Count | venueId
                </div>
              </div>

              <div>
                <h4 className="text-sm font-bold text-slate-800 dark:text-slate-100 mb-1">
                  2. Cross-Programme Candidate Aggregation
                </h4>
                <p>
                  In raw pre-exam exports, multiple academic programmes (e.g., BA Functional Hindi, BCom Finance, BSc Mathematics) taking the same examination paper at the same venue appear on separate rows. 
                  The Confisection engine consolidates them by composite key <code>Exam Date + Venue Code + Course Code</code> and sums candidate counts into a single consolidated QP packing record.
                </p>
              </div>

              <div>
                <h4 className="text-sm font-bold text-slate-800 dark:text-slate-100 mb-1">
                  3. Question Paper Packing Formula
                </h4>
                <p>
                  Standard confidential section packet bundle rules are calculated automatically:
                </p>
                <ul className="list-disc pl-5 mt-1.5 space-y-1">
                  <li><strong>Packets of 20</strong>: <code>Math.floor(Total QP / 20)</code></li>
                  <li><strong>Packets of 10</strong>: <code>Math.floor((Total QP % 20) / 10)</code></li>
                  <li><strong>Packets of 5</strong>: <code>Math.floor(((Total QP % 20) % 10) / 5)</code></li>
                  <li><strong>Loose / Singles</strong>: Remainder single copies</li>
                  <li><strong>Buffer Reserve</strong>: Selectable between +1 copy per venue, +2 copies per venue, +5% reserve, or exact count</li>
                </ul>
              </div>

              <div>
                <h4 className="text-sm font-bold text-slate-800 dark:text-slate-100 mb-1">
                  4. Four-Sheet Master Excel Export
                </h4>
                <p>
                  Clicking <strong>Export Report (.xlsx)</strong> generates a comprehensive workbook containing:
                </p>
                <ol className="list-decimal pl-5 mt-1.5 space-y-1">
                  <li><code>Confi_Master_Schedule</code>: Full master schedule with combined programs, counts, and packet allocations.</li>
                  <li><code>Daily_Press_Print_Order</code>: Daily date &amp; session printing totals for the confidential press.</li>
                  <li><code>Venue_Packing_Slips</code>: Center &amp; venue checklists ready for sealing into dispatch envelopes.</li>
                  <li><code>Course_Master_Summary</code>: Subject-level enrollment and venue coverage.</li>
                </ol>
              </div>
            </div>

            <div className="p-4 border-t border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/80 flex justify-end">
              <button
                onClick={() => setShowGuideModal(false)}
                className="px-4 py-2 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold"
              >
                Got It
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
