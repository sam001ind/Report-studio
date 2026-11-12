import React, { useState, useMemo, useRef, useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Upload,
  FileSpreadsheet,
  Download,
  Search,
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
  SlidersHorizontal,
  ChevronLeft,
  ChevronRight,
  Filter,
  Check,
  ClipboardList,
  Mail,
  Send,
  Copy,
  ExternalLink,
  Plus,
  Trash2,
  Edit2
} from 'lucide-react';
import { readSpreadsheetAsAoa, cleanTimeText, normalizeKeyText, cleanDisplayField } from '../utils/excelParser';
import { 
  CONFIDENTIAL_SECTION_COLUMNS, 
  QP_EMAIL_DISPATCH_COLUMNS,
  generateQpEmailDispatchRows,
  calculateQpPackets, 
  getExamSession,
  SAMPLE_CONFISECTION_ROWS,
  DEFAULT_COLLEGE_EMAILS,
  formatExamDate,
  resolveCollegeEmail
} from '../utils/confisectionHelpers';

export default function ConfisectionExamReportPage() {
  const navigate = useNavigate();

  // State
  const [rawRows, setRawRows] = useState([]);
  const [fileMeta, setFileMeta] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusAlert, setStatusAlert] = useState(null);
  const [copiedNoticeToast, setCopiedNoticeToast] = useState(null);

  // College Email Directory (persisted to localStorage)
  const [emailDirectory, setEmailDirectory] = useState(() => {
    try {
      const saved = localStorage.getItem('kannur_college_email_directory');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch (e) {
      console.warn('Failed reading college email directory from storage:', e);
    }
    return DEFAULT_COLLEGE_EMAILS;
  });

  // Modals state
  const [showGuideModal, setShowGuideModal] = useState(false);
  const [showEmailDirectoryModal, setShowEmailDirectoryModal] = useState(false);
  const [selectedVenueForEmailDraft, setSelectedVenueForEmailDraft] = useState(null);
  const [directorySearchQuery, setDirectorySearchQuery] = useState('');
  const [newCollegeForm, setNewCollegeForm] = useState({ code: '', name: '', email: '' });
  const [editingCollegeIndex, setEditingCollegeIndex] = useState(null);

  // Configuration options
  const [activeTab, setActiveTab] = useState('qp_email_dispatch'); // 'qp_email_dispatch' | 'consolidated' | 'venue_email' | 'press_print' | 'venue_packing' | 'course_summary' | 'raw'
  const [bufferRule, setBufferRule] = useState('plus1'); // 'none' | 'plus1' | 'plus2' | 'five_percent'
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedDate, setSelectedDate] = useState('ALL');
  const [selectedVenue, setSelectedVenue] = useState('ALL');
  const [selectedSession, setSelectedSession] = useState('ALL'); // 'ALL' | 'FN' | 'AN'
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(0);
  const [packedChecklist, setPackedChecklist] = useState({});

  const fileInputRef = useRef(null);
  const directoryFileInputRef = useRef(null);

  // Save email directory changes to localStorage
  const saveEmailDirectory = (updatedList) => {
    setEmailDirectory(updatedList);
    try {
      localStorage.setItem('kannur_college_email_directory', JSON.stringify(updatedList));
    } catch (e) {
      console.error('Failed to save email directory to localStorage', e);
    }
  };

  // Column auto-detection supporting both 13 Confi headers & 5-8 Venue-Date-Course-Email headers
  const detectColumns = (headers) => {
    const norm = headers.map(h => normalizeKeyText(h));
    const findIndex = (patterns) => {
      return norm.findIndex(h => patterns.some(p => h === p || h.includes(p)));
    };

    return {
      eventName: findIndex(['eventname', 'event']),
      centerCode: findIndex(['centercode', 'center']),
      centerName: findIndex(['centername', 'centre']),
      venueCode: findIndex(['venuecode', 'venue', 'collegecode']),
      venueName: findIndex(['venuename', 'college', 'institution']),
      programName: findIndex(['programname', 'programme', 'prog']),
      examDate: findIndex(['examdate', 'date']),
      startTime: findIndex(['coursestarttime', 'starttime', 'start', 'session', 'time']),
      endTime: findIndex(['courseendtime', 'endtime', 'end']),
      courseCode: findIndex(['coursecode', 'code']),
      courseName: findIndex(['coursename', 'course']),
      studentCount: findIndex(['studentcount', 'students', 'count', 'takers', 'candidates']),
      venueId: findIndex(['venueid', 'vid']),
      email: findIndex(['collegeemail', 'email', 'emailid', 'mail', 'venueemail', 'principalemail'])
    };
  };

  // Upload parser supporting both async spreadsheet reading and automatic directory detection
  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setStatusAlert({ type: 'info', message: `Reading "${file.name}"...` });

    try {
      const aoa = await readSpreadsheetAsAoa(file);

      if (!aoa || !Array.isArray(aoa) || aoa.length < 2) {
        throw new Error('Spreadsheet appears to be empty or has no data rows.');
      }

      // Check if uploaded file is a College Email Directory (e.g. College Code, College Name, Email ID)
      let isEmailDirectoryFile = false;
      let emailHeaderIdx = -1;
      for (let i = 0; i < Math.min(10, aoa.length); i++) {
        if (!Array.isArray(aoa[i])) continue;
        const rowText = aoa[i].map(c => normalizeKeyText(c)).join(' ');
        if (rowText.includes('email') && (rowText.includes('college') || rowText.includes('code') || rowText.includes('venue'))) {
          if (!rowText.includes('course') && !rowText.includes('examdate') && !rowText.includes('subject')) {
            isEmailDirectoryFile = true;
            emailHeaderIdx = i;
            break;
          }
        }
      }

      if (isEmailDirectoryFile && emailHeaderIdx !== -1) {
        const rawHeaders = aoa[emailHeaderIdx] || [];
        const dHeaders = rawHeaders.map(c => normalizeKeyText(c));
        const codeIdx = dHeaders.findIndex(h => h.includes('code') || h.includes('id'));
        const nameIdx = dHeaders.findIndex(h => h.includes('name') || h.includes('college'));
        const emailIdx = dHeaders.findIndex(h => h.includes('email') || h.includes('mail'));

        if (emailIdx !== -1) {
          const importedColleges = [];
          for (let r = emailHeaderIdx + 1; r < aoa.length; r++) {
            const row = aoa[r];
            if (!Array.isArray(row)) continue;
            const code = codeIdx !== -1 ? cleanDisplayField(row[codeIdx]) : '';
            const name = nameIdx !== -1 && nameIdx !== codeIdx ? cleanDisplayField(row[nameIdx]) : '';
            const email = cleanDisplayField(row[emailIdx]);
            if (code && email && email.includes('@')) {
              importedColleges.push({ code, name: name || code, email });
            }
          }

          if (importedColleges.length > 0) {
            saveEmailDirectory(importedColleges);
            setStatusAlert({
              type: 'success',
              message: `College Email Directory Detected: Successfully imported ${importedColleges.length} official college email records from "${file.name}".`
            });
            setIsProcessing(false);
            if (fileInputRef.current) fileInputRef.current.value = '';
            return;
          }
        }
      }

      // Otherwise, process as Examination Schedule
      let headerRowIdx = 0;
      for (let i = 0; i < Math.min(15, aoa.length); i++) {
        if (!Array.isArray(aoa[i])) continue;
        const rowText = aoa[i].map(c => normalizeKeyText(c)).join(' ');
        if (rowText.includes('course') && (rowText.includes('venue') || rowText.includes('date') || rowText.includes('code'))) {
          headerRowIdx = i;
          break;
        }
      }

      const rawHeaderRow = aoa[headerRowIdx] || aoa[0] || [];
      const headers = Array.isArray(rawHeaderRow) ? rawHeaderRow.map(c => String(c || '').trim()) : [];
      const mapping = detectColumns(headers);

      if (mapping.courseCode === -1 && mapping.examDate === -1) {
        throw new Error('Could not identify Course Code or Exam Date columns. Expected standard Examination Schedule headers.');
      }

      const parsedRecords = [];
      let eventDetected = '';

      for (let r = headerRowIdx + 1; r < aoa.length; r++) {
        const row = aoa[r];
        if (!row || !Array.isArray(row)) continue;

        const rawCourseCode = mapping.courseCode !== -1 ? row[mapping.courseCode] : '';
        const rawExamDate = mapping.examDate !== -1 ? row[mapping.examDate] : '';
        if (!rawCourseCode && !rawExamDate) continue;

        const eventName = mapping.eventName !== -1 ? cleanDisplayField(row[mapping.eventName]) : '';
        if (!eventDetected && eventName) eventDetected = eventName;

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
        const collegeEmail = mapping.email !== -1 ? cleanDisplayField(row[mapping.email]) : '';

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
          venueId,
          collegeEmail
        });
      }

      setRawRows(parsedRecords);
      setFileMeta({
        name: file.name,
        size: `${(file.size / 1024).toFixed(1)} KB`,
        rowsCount: parsedRecords.length,
        event: eventDetected || 'Examination Schedule'
      });
      setCurrentPage(0);
      setStatusAlert({
        type: 'success',
        message: `Loaded ${parsedRecords.length.toLocaleString()} schedule entries from "${file.name}". Ready for Confisection & Email analysis.`
      });
    } catch (err) {
      console.error('File parsing error:', err);
      setStatusAlert({ type: 'error', message: `Failed to load file: ${err.message}` });
    } finally {
      setIsProcessing(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  // Dedicated College Email Directory File Upload
  const handleDirectoryFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const aoa = await readSpreadsheetAsAoa(file);
      if (!aoa || !Array.isArray(aoa) || aoa.length < 2) {
        throw new Error('Directory file appears to be empty or has no data rows.');
      }

      let headerRowIdx = 0;
      for (let i = 0; i < Math.min(10, aoa.length); i++) {
        if (!Array.isArray(aoa[i])) continue;
        const rowText = aoa[i].map(c => normalizeKeyText(c)).join(' ');
        if (rowText.includes('code') || rowText.includes('email') || rowText.includes('college')) {
          headerRowIdx = i;
          break;
        }
      }

      const rawHeaders = aoa[headerRowIdx] || aoa[0] || [];
      const dHeaders = rawHeaders.map(c => normalizeKeyText(c));
      const codeIdx = dHeaders.findIndex(h => h.includes('code') || h.includes('id') || h === 'college');
      const nameIdx = dHeaders.findIndex(h => h.includes('name') || h.includes('college'));
      const emailIdx = dHeaders.findIndex(h => h.includes('email') || h.includes('mail'));

      if (emailIdx === -1) {
        throw new Error('Could not identify an Email ID column in the directory file.');
      }

      const imported = [];
      for (let r = headerRowIdx + 1; r < aoa.length; r++) {
        const row = aoa[r];
        if (!Array.isArray(row)) continue;
        const code = codeIdx !== -1 ? cleanDisplayField(row[codeIdx]) : '';
        const name = nameIdx !== -1 && nameIdx !== codeIdx ? cleanDisplayField(row[nameIdx]) : '';
        const email = cleanDisplayField(row[emailIdx]);
        if (code && email && email.includes('@')) {
          imported.push({ code: code || name || 'COLLEGE', name: name || code, email });
        }
      }

      if (imported.length === 0) {
        throw new Error('No valid college email rows found in the file.');
      }

      saveEmailDirectory(imported);
      setStatusAlert({
        type: 'success',
        message: `Imported ${imported.length} college email records from "${file.name}".`
      });
    } catch (err) {
      console.error('Directory upload error:', err);
      setStatusAlert({ type: 'error', message: `Directory import failed: ${err.message}` });
    } finally {
      if (directoryFileInputRef.current) directoryFileInputRef.current.value = '';
    }
  };

  // Load sample dataset
  const handleLoadSample = () => {
    setRawRows(SAMPLE_CONFISECTION_ROWS);
    setFileMeta({
      name: 'Kannur_UG_III_Sem_Exam_Schedule_Sample.xlsx',
      size: '26.8 KB',
      rowsCount: SAMPLE_CONFISECTION_ROWS.length,
      event: 'III Sem(R) November 2026'
    });
    setCurrentPage(0);
    setStatusAlert({
      type: 'success',
      message: `Loaded ${SAMPLE_CONFISECTION_ROWS.length} sample records from Kannur University III Sem (Regular) schedule with official college emails.`
    });
  };

  const handleClearData = () => {
    setRawRows([]);
    setFileMeta(null);
    setCurrentPage(0);
    setStatusAlert(null);
  };

  // Consolidated Master Report (Unique Date + Venue + Course Code with summed student counts & resolved email)
  const consolidatedRows = useMemo(() => {
    if (rawRows.length === 0) return [];

    const groupMap = new Map();

    rawRows.forEach((row, idx) => {
      const vCode = row.venueCode || 'UNKNOWN';
      const cCode = row.courseCode || 'UNKNOWN';
      const date = row.examDate || 'UNKNOWN';

      const key = `${date}|||${vCode}|||${cCode}`;
      const session = getExamSession(row.startTime, row.endTime);

      if (groupMap.has(key)) {
        const item = groupMap.get(key);
        item.studentCount += row.studentCount;
        if (row.programName && !item.programs.includes(row.programName)) {
          item.programs.push(row.programName);
        }
        if (!item.startTime && row.startTime) item.startTime = row.startTime;
        if (!item.endTime && row.endTime) item.endTime = row.endTime;
        if (!item.collegeEmail && row.collegeEmail) item.collegeEmail = row.collegeEmail;
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
          collegeEmail: row.collegeEmail,
          programs: row.programName ? [row.programName] : []
        });
      }
    });

    const list = Array.from(groupMap.values()).map(item => {
      const packets = calculateQpPackets(item.studentCount, bufferRule);
      const email = resolveCollegeEmail(item.venueCode, item.venueName, emailDirectory, item.collegeEmail);
      return {
        ...item,
        collegeEmail: email,
        hasEmail: Boolean(email && email.includes('@')),
        programsFormatted: item.programs.join('; '),
        programsCount: item.programs.length,
        ...packets
      };
    });

    list.sort((a, b) => {
      if (a.examDate !== b.examDate) return a.examDate.localeCompare(b.examDate);
      if (a.venueCode !== b.venueCode) return a.venueCode.localeCompare(b.venueCode);
      return a.courseCode.localeCompare(b.courseCode);
    });

    return list;
  }, [rawRows, bufferRule, emailDirectory]);

  // Dedicated QP Email Dispatch Report rows (Exam Date, Course Code, Course Name, Venue Code, Venue Name, Email ID + exam details)
  const qpEmailDispatchRows = useMemo(() => {
    return generateQpEmailDispatchRows(consolidatedRows);
  }, [consolidatedRows]);

  // Venue College Email Dispatch Summary
  const venueEmailSummaryRows = useMemo(() => {
    if (consolidatedRows.length === 0) return [];
    const map = new Map();

    consolidatedRows.forEach(row => {
      const vKey = row.venueCode || 'UNKNOWN';
      if (!map.has(vKey)) {
        map.set(vKey, {
          venueCode: row.venueCode,
          venueName: row.venueName,
          centerCode: row.centerCode,
          centerName: row.centerName,
          collegeEmail: row.collegeEmail,
          hasEmail: row.hasEmail,
          totalStudents: 0,
          totalQp: 0,
          exams: []
        });
      }
      const entry = map.get(vKey);
      entry.totalStudents += row.studentCount;
      entry.totalQp += row.totalQp;
      entry.exams.push(row);
    });

    const list = Array.from(map.values()).map(v => ({
      ...v,
      examsCount: v.exams.length
    }));

    list.sort((a, b) => a.venueName.localeCompare(b.venueName));
    return list;
  }, [consolidatedRows]);

  // Daily Press Print Order
  const pressPrintOrderRows = useMemo(() => {
    if (consolidatedRows.length === 0) return [];

    const map = new Map();

    consolidatedRows.forEach(row => {
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
          collegeEmail: row.collegeEmail,
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

  // Filter options
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

  // Active filtered records
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
        const email = String(item.emailId || item.collegeEmail || '').toLowerCase();
        return code.includes(q) || name.includes(q) || vCode.includes(q) || vName.includes(q) || date.includes(q) || prog.includes(q) || email.includes(q);
      }
      return true;
    };

    if (activeTab === 'qp_email_dispatch') {
      return qpEmailDispatchRows.filter(applyCommonFilter);
    }
    if (activeTab === 'consolidated') {
      return consolidatedRows.filter(applyCommonFilter);
    }
    if (activeTab === 'venue_email') {
      return venueEmailSummaryRows.filter(item => {
        if (selectedVenue !== 'ALL' && item.venueCode !== selectedVenue) return false;
        if (q) {
          const vCode = String(item.venueCode || '').toLowerCase();
          const vName = String(item.venueName || '').toLowerCase();
          const email = String(item.collegeEmail || '').toLowerCase();
          return vCode.includes(q) || vName.includes(q) || email.includes(q);
        }
        return true;
      });
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
        if (selectedDate !== 'ALL' && !item.dates.includes(selectedDate)) return false;
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
          const prog = String(item.programName || '').toLowerCase();
          const email = String(item.collegeEmail || '').toLowerCase();
          return code.includes(q) || name.includes(q) || vCode.includes(q) || vName.includes(q) || prog.includes(q) || email.includes(q);
        }
        return true;
      });
    }

    return [];
  }, [activeTab, qpEmailDispatchRows, consolidatedRows, venueEmailSummaryRows, pressPrintOrderRows, courseSummaryRows, rawRows, searchQuery, selectedDate, selectedSession, selectedVenue]);

  // Paginated view
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
    const uniqueVenues = venueEmailSummaryRows.length;
    const mappedVenues = venueEmailSummaryRows.filter(v => v.hasEmail).length;
    const uniqueCourses = new Set(consolidatedRows.map(r => r.courseCode)).size;
    const uniqueDates = new Set(consolidatedRows.map(r => r.examDate)).size;

    return {
      totalCandidates,
      totalQp,
      uniqueExams,
      uniqueVenues,
      mappedVenues,
      uniqueCourses,
      uniqueDates
    };
  }, [consolidatedRows, venueEmailSummaryRows]);

  // Toggle packing checklist item
  const handleToggleChecklist = (itemKey) => {
    setPackedChecklist(prev => ({
      ...prev,
      [itemKey]: !prev[itemKey]
    }));
  };

  // Copy all unique college emails to clipboard
  const handleCopyAllEmails = () => {
    const emails = venueEmailSummaryRows
      .map(v => v.collegeEmail)
      .filter(e => e && e.includes('@'));
    const uniqueEmails = Array.from(new Set(emails));

    if (uniqueEmails.length === 0) {
      setStatusAlert({ type: 'error', message: 'No valid college email addresses found to copy.' });
      return;
    }

    const textToCopy = uniqueEmails.join('; ');
    navigator.clipboard.writeText(textToCopy);
    setCopiedNoticeToast(`Copied ${uniqueEmails.length} unique college email addresses! Ready to paste into Gmail / Outlook BCC.`);
    setTimeout(() => setCopiedNoticeToast(null), 4000);
  };

  // Dedicated Export for QP Email Dispatch Report (.xlsx)
  const handleExportQpEmailReportOnly = () => {
    try {
      const wb = XLSX.utils.book_new();
      const headers = [
        ...QP_EMAIL_DISPATCH_COLUMNS,
        'Session',
        'Course  Start Time',
        'Course  End Time',
        'Student Count',
        'Packets of 20',
        'Packets of 10',
        'Packets of 5',
        'Singles',
        'Buffer',
        'Total QP Copies',
        'Combined Programs'
      ];
      const data = qpEmailDispatchRows.map(r => [
        r.examDate,
        r.courseCode,
        r.courseName,
        r.venueCode,
        r.venueName,
        r.emailId,
        r.session,
        r.startTime,
        r.endTime,
        r.studentCount,
        r.p20,
        r.p10,
        r.p5,
        r.singles,
        r.buffer,
        r.totalQp,
        r.programsFormatted
      ]);
      const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
      ws['!cols'] = headers.map(() => ({ wch: 22 }));
      XLSX.utils.book_append_sheet(wb, ws, 'QP_Email_Dispatch_Report');
      const outName = `QP_Email_Dispatch_Report_${fileMeta?.event ? fileMeta.event.replace(/[^a-zA-Z0-9]/g, '_') : 'Schedule'}.xlsx`;
      XLSX.writeFile(wb, outName);
      setStatusAlert({ type: 'success', message: `Downloaded ${outName} with all QP Email Dispatch records!` });
    } catch (err) {
      console.error('Export error:', err);
      setStatusAlert({ type: 'error', message: `Export failed: ${err.message}` });
    }
  };

  // Export to Excel with all 6 confidential & email dispatch sheets
  const handleExportExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: QP Email Dispatch Report
      const sheetQpHeaders = [
        ...QP_EMAIL_DISPATCH_COLUMNS,
        'Session',
        'Course  Start Time',
        'Course  End Time',
        'Student Count',
        'Packets of 20',
        'Packets of 10',
        'Packets of 5',
        'Singles',
        'Buffer',
        'Total QP Copies',
        'Combined Programs'
      ];
      const sheetQpData = qpEmailDispatchRows.map(r => [
        r.examDate,
        r.courseCode,
        r.courseName,
        r.venueCode,
        r.venueName,
        r.emailId,
        r.session,
        r.startTime,
        r.endTime,
        r.studentCount,
        r.p20,
        r.p10,
        r.p5,
        r.singles,
        r.buffer,
        r.totalQp,
        r.programsFormatted
      ]);
      const wsQp = XLSX.utils.aoa_to_sheet([sheetQpHeaders, ...sheetQpData]);
      wsQp['!cols'] = sheetQpHeaders.map(() => ({ wch: 22 }));
      XLSX.utils.book_append_sheet(wb, wsQp, 'QP_Email_Dispatch_Report');

      // Sheet 2: Master Schedule
      const sheet1Headers = [
        'Event Name', 'Center Code', 'Center Name', 'Venue Code', 'Venue Name', 'College Email', 'venueId',
        'Exam Date', 'Session', 'Start Time', 'End Time', 'Course Code', 'Course Name',
        'Combined Programs', 'Student Count', 'Packets of 20', 'Packets of 10', 'Packets of 5',
        'Loose / Singles', 'Buffer Copies', 'Total QP Required'
      ];
      const sheet1Data = consolidatedRows.map(r => [
        r.eventName, r.centerCode, r.centerName, r.venueCode, r.venueName, r.collegeEmail, r.venueId,
        r.examDate, r.session, r.startTime, r.endTime, r.courseCode, r.courseName,
        r.programsFormatted, r.studentCount, r.p20, r.p10, r.p5, r.singles, r.buffer, r.totalQp
      ]);
      const ws1 = XLSX.utils.aoa_to_sheet([sheet1Headers, ...sheet1Data]);
      ws1['!cols'] = sheet1Headers.map(() => ({ wch: 22 }));
      XLSX.utils.book_append_sheet(wb, ws1, 'Confi_Master_Schedule');

      // Sheet 3: Venue College Email Dispatch Schedule
      const sheet2Headers = [
        'Venue Code', 'Venue Name', 'College Official Email', 'Email Status', 'Scheduled Exams',
        'Total Candidates', 'Total Question Papers'
      ];
      const sheet2Data = venueEmailSummaryRows.map(v => [
        v.venueCode, v.venueName, v.collegeEmail, v.hasEmail ? 'Configured' : 'Missing',
        v.examsCount, v.totalStudents, v.totalQp
      ]);
      const ws2 = XLSX.utils.aoa_to_sheet([sheet2Headers, ...sheet2Data]);
      ws2['!cols'] = sheet2Headers.map(() => ({ wch: 24 }));
      XLSX.utils.book_append_sheet(wb, ws2, 'Venue_Email_Dispatch');

      // Sheet 4: Daily Press Print Order
      const sheet3Headers = [
        'Exam Date', 'Session', 'Start Time', 'End Time', 'Course Code', 'Course Name',
        'Total Venues', 'Total Candidates', 'Total QP Print Order', 'Venues List'
      ];
      const sheet3Data = pressPrintOrderRows.map(r => [
        r.examDate, r.session, r.startTime, r.endTime, r.courseCode, r.courseName,
        r.venuesCount, r.totalStudents, r.totalQp, r.venueCodes.join(', ')
      ]);
      const ws3 = XLSX.utils.aoa_to_sheet([sheet3Headers, ...sheet3Data]);
      ws3['!cols'] = sheet3Headers.map(() => ({ wch: 22 }));
      XLSX.utils.book_append_sheet(wb, ws3, 'Daily_Press_Print_Order');

      // Sheet 5: Venue Packing Slips
      const sheet4Headers = [
        'Venue Code', 'Venue Name', 'College Email', 'Center Code', 'Exam Date', 'Session',
        'Course Code', 'Course Name', 'Student Count', '20s', '10s', '5s', 'Singles', 'Total QP Packed'
      ];
      const sheet4Data = [];
      Object.values(venuePackingSlips).forEach(v => {
        v.items.forEach(item => {
          sheet4Data.push([
            v.venueCode, v.venueName, v.collegeEmail, v.centerCode, item.examDate, item.session,
            item.courseCode, item.courseName, item.studentCount, item.p20, item.p10, item.p5, item.singles, item.totalQp
          ]);
        });
      });
      const ws4 = XLSX.utils.aoa_to_sheet([sheet4Headers, ...sheet4Data]);
      ws4['!cols'] = sheet4Headers.map(() => ({ wch: 20 }));
      XLSX.utils.book_append_sheet(wb, ws4, 'Venue_Packing_Slips');

      // Sheet 6: Course Master Summary
      const sheet5Headers = [
        'Course Code', 'Course Name', 'Scheduled Dates', 'Session',
        'Total Venues', 'Total Candidates', 'Total QP Needed'
      ];
      const sheet5Data = courseSummaryRows.map(r => [
        r.courseCode, r.courseName, r.datesFormatted, r.session,
        r.venueCount, r.totalStudents, r.totalQp
      ]);
      const ws5 = XLSX.utils.aoa_to_sheet([sheet5Headers, ...sheet5Data]);
      ws5['!cols'] = sheet5Headers.map(() => ({ wch: 22 }));
      XLSX.utils.book_append_sheet(wb, ws5, 'Course_Master_Summary');

      const outName = `Confisection_Exam_Report_${fileMeta?.event ? fileMeta.event.replace(/[^a-zA-Z0-9]/g, '_') : 'Schedule'}.xlsx`;
      XLSX.writeFile(wb, outName);
      setStatusAlert({ type: 'success', message: `Downloaded ${outName} with all 6 confidential section & email dispatch reports!` });
    } catch (err) {
      console.error('Export error:', err);
      setStatusAlert({ type: 'error', message: `Export failed: ${err.message}` });
    }
  };

  // Download 13-Column Template
  const handleDownloadTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      CONFIDENTIAL_SECTION_COLUMNS,
      [
        'III Sem(R) November 2026', '1', 'Kannur', 'PP', 'Payyannur College, Payyannur',
        'Bachelor of Arts in English(with Credits)', '25/11/2026', '01:30 PM', '03:30 PM',
        'KU3DSCENG205', 'Migration and Border Studies', 25, '68'
      ]
    ]);
    ws['!cols'] = CONFIDENTIAL_SECTION_COLUMNS.map(() => ({ wch: 22 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Confi_Template');
    XLSX.writeFile(wb, 'Confisection_Exam_Source_Template.xlsx');
  };

  // College Directory Management Handlers
  const handleAddCollegeToDirectory = (e) => {
    e.preventDefault();
    if (!newCollegeForm.code || !newCollegeForm.email) {
      setStatusAlert({ type: 'error', message: 'College Code and Email are required.' });
      return;
    }

    if (editingCollegeIndex !== null) {
      const updated = [...emailDirectory];
      updated[editingCollegeIndex] = { ...newCollegeForm };
      saveEmailDirectory(updated);
      setEditingCollegeIndex(null);
      setStatusAlert({ type: 'success', message: `Updated email for ${newCollegeForm.name || newCollegeForm.code}` });
    } else {
      const updated = [...emailDirectory, { ...newCollegeForm }];
      saveEmailDirectory(updated);
      setStatusAlert({ type: 'success', message: `Added ${newCollegeForm.name || newCollegeForm.code} to directory` });
    }

    setNewCollegeForm({ code: '', name: '', email: '' });
  };

  const handleEditCollege = (idx) => {
    setEditingCollegeIndex(idx);
    setNewCollegeForm({ ...emailDirectory[idx] });
  };

  const handleDeleteCollege = (idx) => {
    const item = emailDirectory[idx];
    if (window.confirm(`Remove ${item.name || item.code} from directory?`)) {
      const updated = emailDirectory.filter((_, i) => i !== idx);
      saveEmailDirectory(updated);
      setStatusAlert({ type: 'info', message: `Removed college from directory.` });
    }
  };

  const handleResetDirectoryToDefault = () => {
    if (window.confirm('Reset College Email Directory to official Kannur University defaults?')) {
      saveEmailDirectory(DEFAULT_COLLEGE_EMAILS);
      setStatusAlert({ type: 'success', message: 'Reset email directory to official Kannur University defaults.' });
    }
  };

  return (
    <div style={{ maxWidth: '1440px', margin: '0 auto', padding: '24px 20px', minHeight: '100vh', background: 'var(--bg)', color: 'var(--ink)' }}>
      {/* Toast Alert for Copied text */}
      {copiedNoticeToast && (
        <div style={{
          position: 'fixed',
          top: '20px',
          right: '24px',
          zIndex: 9999,
          background: 'var(--ink)',
          color: 'var(--bg)',
          padding: '12px 20px',
          borderRadius: '8px',
          boxShadow: '0 8px 24px rgba(0,0,0,0.25)',
          fontSize: '13px',
          fontWeight: 600,
          display: 'flex',
          alignItems: 'center',
          gap: '10px'
        }}>
          <CheckCircle2 size={18} color="#10b981" />
          <span>{copiedNoticeToast}</span>
        </div>
      )}

      {/* Top Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '12px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <button
            type="button"
            onClick={() => navigate('/')}
            className="secondary"
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 12px', fontSize: '13px' }}
          >
            <ArrowLeft size={16} /> Home
          </button>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <div style={{
                background: 'linear-gradient(135deg, #e11d48, #be123c)',
                color: 'white',
                padding: '6px',
                borderRadius: '8px',
                display: 'inline-flex'
              }}>
                <ShieldCheck size={20} />
              </div>
              <h1 style={{ fontSize: '22px', fontWeight: 800, margin: 0, letterSpacing: '-0.3px', color: 'var(--ink)' }}>
                Confisection Exam &amp; Venue College Dispatch Report
              </h1>
            </div>
            <p style={{ margin: '3px 0 0', fontSize: '12.5px', color: 'var(--muted)' }}>
              All-in-one Confidential Section Studio: Question paper packet math (20s/10s/5s), daily press orders, cross-programme student aggregates, and college email dispatch schedules.
            </p>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
          <button
            type="button"
            onClick={() => setShowEmailDirectoryModal(true)}
            className="secondary"
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 14px', fontSize: '13px' }}
            title="Manage official college emails"
          >
            <Mail size={15} color="var(--accent)" /> College Email Directory ({emailDirectory.length})
          </button>

          <button
            type="button"
            onClick={() => setShowGuideModal(true)}
            className="secondary"
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 14px', fontSize: '13px' }}
          >
            <HelpCircle size={15} /> Logic &amp; Guide
          </button>

          <button
            type="button"
            onClick={handleDownloadTemplate}
            className="secondary"
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 14px', fontSize: '13px' }}
          >
            <FileSpreadsheet size={15} /> Template (.xlsx)
          </button>

          {consolidatedRows.length > 0 && (
            <>
              <button
                type="button"
                onClick={handleExportQpEmailReportOnly}
                className="secondary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 14px', fontSize: '13px', color: '#059669', borderColor: 'rgba(5, 150, 105, 0.3)' }}
                title="Export QP Email Dispatch Report (.xlsx)"
              >
                <Download size={15} /> Export QP Dispatch (.xlsx)
              </button>
              <button
                type="button"
                onClick={() => window.print()}
                className="secondary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 14px', fontSize: '13px' }}
              >
                <Printer size={15} /> Print View
              </button>
              <button
                type="button"
                onClick={handleExportExcel}
                className="primary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 16px', fontSize: '13px', background: '#e11d48', borderColor: '#e11d48' }}
              >
                <Download size={15} /> Export All Sheets (.xlsx)
              </button>
            </>
          )}
        </div>
      </div>

      {/* Status Notifications */}
      {statusAlert && (
        <div style={{
          padding: '12px 16px',
          borderRadius: '8px',
          marginBottom: '20px',
          fontSize: '13px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          background: statusAlert.type === 'error' ? 'rgba(239, 68, 68, 0.1)' : statusAlert.type === 'success' ? 'rgba(16, 185, 129, 0.1)' : 'var(--panel)',
          border: `1px solid ${statusAlert.type === 'error' ? 'rgba(239, 68, 68, 0.3)' : statusAlert.type === 'success' ? 'rgba(16, 185, 129, 0.3)' : 'var(--line)'}`,
          color: statusAlert.type === 'error' ? '#dc2626' : statusAlert.type === 'success' ? '#059669' : 'var(--ink)'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            {statusAlert.type === 'error' ? <AlertCircle size={16} /> : <CheckCircle2 size={16} />}
            <span>{statusAlert.message}</span>
          </div>
          <button
            type="button"
            onClick={() => setStatusAlert(null)}
            style={{ background: 'transparent', border: 'none', cursor: 'pointer', padding: 0, color: 'inherit' }}
          >
            <X size={14} />
          </button>
        </div>
      )}

      {/* Ingestion & Config Card */}
      <div className="card" style={{ padding: '20px', marginBottom: '20px', border: '1px solid var(--line)', background: 'var(--panel)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px' }}>
          {/* File Upload Zone */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              style={{ display: 'none' }}
            />
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              disabled={isProcessing}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '8px',
                padding: '10px 18px',
                fontSize: '13px',
                fontWeight: 700,
                borderRadius: '8px',
                background: 'var(--accent)',
                color: 'white',
                border: 'none',
                cursor: 'pointer'
              }}
            >
              <Upload size={16} />
              {isProcessing ? 'Processing File...' : 'Upload Examination Schedule'}
            </button>

            <span style={{ fontSize: '13px', color: 'var(--muted)' }}>or</span>

            <button
              type="button"
              onClick={handleLoadSample}
              className="secondary"
              style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '9px 14px', fontSize: '13px' }}
            >
              <Sparkles size={15} color="#e11d48" /> Load Kannur Sample Data
            </button>

            {rawRows.length > 0 && (
              <button
                type="button"
                onClick={handleClearData}
                className="secondary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '9px 14px', fontSize: '13px', color: 'var(--muted)' }}
              >
                <RefreshCw size={14} /> Clear
              </button>
            )}
          </div>

          {/* Buffer Rule Policy Selector */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <span style={{ fontSize: '12.5px', fontWeight: 600, color: 'var(--muted)', display: 'inline-flex', alignItems: 'center', gap: '6px' }}>
              <SlidersHorizontal size={14} /> Buffer Reserve Policy:
            </span>
            <select
              value={bufferRule}
              onChange={(e) => setBufferRule(e.target.value)}
              style={{
                fontSize: '13px',
                padding: '8px 12px',
                borderRadius: '6px',
                border: '1px solid var(--line)',
                background: 'var(--bg)',
                color: 'var(--ink)',
                fontWeight: 600,
                cursor: 'pointer'
              }}
            >
              <option value="plus1">+1 Copy / Venue (+1 Buffer)</option>
              <option value="plus2">+2 Copies / Venue (+2 Buffer)</option>
              <option value="five_percent">+5% Reserve (Min 1)</option>
              <option value="none">Exact Count (No Buffer)</option>
            </select>
          </div>
        </div>

        {/* File Metadata Chip */}
        {fileMeta && (
          <div style={{
            marginTop: '16px',
            paddingTop: '14px',
            borderTop: '1px solid var(--line)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flexWrap: 'wrap',
            gap: '12px',
            fontSize: '12.5px'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <FileSpreadsheet size={16} color="var(--accent)" />
              <strong style={{ color: 'var(--ink)' }}>{fileMeta.name}</strong>
              <span style={{ color: 'var(--muted)' }}>({fileMeta.size})</span>
              <span style={{
                background: 'var(--accent-soft)',
                color: 'var(--accent)',
                padding: '2px 8px',
                borderRadius: '4px',
                fontWeight: 700,
                fontSize: '11px'
              }}>
                {fileMeta.event}
              </span>
            </div>
            <div style={{ color: 'var(--muted)' }}>
              Source records: <strong style={{ color: 'var(--ink)' }}>{fileMeta.rowsCount.toLocaleString()}</strong> rows
            </div>
          </div>
        )}
      </div>

      {/* KPI Stat Cards */}
      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))',
        gap: '14px',
        marginBottom: '20px'
      }}>
        {/* Candidates */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid var(--accent)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Total Candidates
            </span>
            <Users size={18} color="var(--accent)" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>
            {summaryStats.totalCandidates.toLocaleString()}
          </div>
          <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Sum of all registrations</span>
        </div>

        {/* Question Papers */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid #e11d48' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Total QP Required
            </span>
            <ShieldCheck size={18} color="#e11d48" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: '#e11d48' }}>
            {summaryStats.totalQp.toLocaleString()}
          </div>
          <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Includes {bufferRule} buffer</span>
        </div>

        {/* Unique Exams */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid #7c3aed' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Unique Exams
            </span>
            <Calendar size={18} color="#7c3aed" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>
            {summaryStats.uniqueExams.toLocaleString()}
          </div>
          <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Date + Venue + Course</span>
        </div>

        {/* Colleges / Venues & Emails */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid #059669' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Colleges &amp; Emails
            </span>
            <Mail size={18} color="#059669" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>
            {summaryStats.uniqueVenues}
          </div>
          <span style={{ fontSize: '11px', color: '#059669', fontWeight: 700 }}>
            {summaryStats.mappedVenues} of {summaryStats.uniqueVenues} mapped with email ({summaryStats.uniqueVenues > 0 ? Math.round((summaryStats.mappedVenues / summaryStats.uniqueVenues) * 100) : 0}%)
          </span>
        </div>

        {/* Distinct Courses */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid #d97706' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Question Papers
            </span>
            <BookOpen size={18} color="#d97706" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>
            {summaryStats.uniqueCourses}
          </div>
          <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Distinct subject papers</span>
        </div>

        {/* Scheduled Dates */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid #0284c7' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Exam Days
            </span>
            <CalendarDays size={18} color="#0284c7" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>
            {summaryStats.uniqueDates}
          </div>
          <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Scheduled dates</span>
        </div>
      </div>

      {/* Main Container with Tabs & Data View */}
      <div className="card" style={{ padding: 0, overflow: 'hidden', borderRadius: '12px' }}>
        {/* Tab Navigation */}
        <div style={{ 
          display: 'flex', 
          borderBottom: '1px solid var(--line)', 
          background: 'var(--bg)', 
          padding: '4px 16px 0', 
          overflowX: 'auto',
          gap: '6px'
        }}>
          {/* 1. QP Email Dispatch Report */}
          <button
            type="button"
            onClick={() => { setActiveTab('qp_email_dispatch'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'qp_email_dispatch' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'qp_email_dispatch' ? '#059669' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'qp_email_dispatch' ? '3px solid #059669' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <Send size={16} color={activeTab === 'qp_email_dispatch' ? '#059669' : 'var(--muted)'} />
            QP Email Dispatch Report
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'qp_email_dispatch' ? 'rgba(5, 150, 105, 0.15)' : 'var(--line)',
              color: activeTab === 'qp_email_dispatch' ? '#059669' : 'var(--muted)'
            }}>
              {qpEmailDispatchRows.length}
            </span>
          </button>

          {/* 2. Master Schedule */}
          <button
            type="button"
            onClick={() => { setActiveTab('consolidated'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'consolidated' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'consolidated' ? 'var(--accent)' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'consolidated' ? '3px solid var(--accent)' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <PackageCheck size={16} />
            Consolidated Master Schedule
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'consolidated' ? 'var(--accent-soft)' : 'var(--line)',
              color: activeTab === 'consolidated' ? 'var(--accent)' : 'var(--muted)'
            }}>
              {consolidatedRows.length}
            </span>
          </button>

          {/* 2. Venue College Email Dispatch */}
          <button
            type="button"
            onClick={() => { setActiveTab('venue_email'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'venue_email' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'venue_email' ? '#059669' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'venue_email' ? '3px solid #059669' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <Mail size={16} color={activeTab === 'venue_email' ? '#059669' : 'var(--muted)'} />
            College Email Dispatch
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'venue_email' ? 'rgba(5, 150, 105, 0.15)' : 'var(--line)',
              color: activeTab === 'venue_email' ? '#059669' : 'var(--muted)'
            }}>
              {venueEmailSummaryRows.length}
            </span>
          </button>

          {/* 3. Daily Press Print Order */}
          <button
            type="button"
            onClick={() => { setActiveTab('press_print'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'press_print' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'press_print' ? '#e11d48' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'press_print' ? '3px solid #e11d48' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <Printer size={16} color={activeTab === 'press_print' ? '#e11d48' : 'var(--muted)'} />
            Daily Press Print Order
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'press_print' ? 'rgba(225, 29, 72, 0.15)' : 'var(--line)',
              color: activeTab === 'press_print' ? '#e11d48' : 'var(--muted)'
            }}>
              {pressPrintOrderRows.length}
            </span>
          </button>

          {/* 4. Venue Packing Slips */}
          <button
            type="button"
            onClick={() => { setActiveTab('venue_packing'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'venue_packing' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'venue_packing' ? '#7c3aed' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'venue_packing' ? '3px solid #7c3aed' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <ClipboardList size={16} color={activeTab === 'venue_packing' ? '#7c3aed' : 'var(--muted)'} />
            Venue Packing Slips &amp; Checklist
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'venue_packing' ? 'rgba(124, 58, 237, 0.15)' : 'var(--line)',
              color: activeTab === 'venue_packing' ? '#7c3aed' : 'var(--muted)'
            }}>
              {Object.keys(venuePackingSlips).length}
            </span>
          </button>

          {/* 5. Course Master Summary */}
          <button
            type="button"
            onClick={() => { setActiveTab('course_summary'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'course_summary' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'course_summary' ? '#d97706' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'course_summary' ? '3px solid #d97706' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <BookOpen size={16} color={activeTab === 'course_summary' ? '#d97706' : 'var(--muted)'} />
            Course Master Summary
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'course_summary' ? 'rgba(217, 119, 6, 0.15)' : 'var(--line)',
              color: activeTab === 'course_summary' ? '#d97706' : 'var(--muted)'
            }}>
              {courseSummaryRows.length}
            </span>
          </button>

          {/* 6. Raw Source Rows */}
          <button
            type="button"
            onClick={() => { setActiveTab('raw'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'raw' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'raw' ? 'var(--accent)' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'raw' ? '3px solid var(--accent)' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <FileSpreadsheet size={16} />
            Raw Source Rows
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'raw' ? 'var(--accent-soft)' : 'var(--line)',
              color: activeTab === 'raw' ? 'var(--accent)' : 'var(--muted)'
            }}>
              {rawRows.length}
            </span>
          </button>
        </div>

        {/* Filter & Search Bar */}
        <div style={{
          padding: '14px 20px',
          background: 'var(--panel)',
          borderBottom: '1px solid var(--line)',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: '12px'
        }}>
          {/* Search Box */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flex: '1 1 280px', maxWidth: '420px' }}>
            <div style={{ position: 'relative', width: '100%' }}>
              <Search size={15} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
              <input
                type="text"
                placeholder={activeTab === 'venue_email' ? 'Search venue code, name, or college email...' : activeTab === 'qp_email_dispatch' ? 'Search date, course code, venue, or college email...' : 'Search course code, subject, college, email...'}
                value={searchQuery}
                onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(0); }}
                style={{
                  width: '100%',
                  padding: '7px 12px 7px 32px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  fontSize: '13px'
                }}
              />
            </div>
            {searchQuery && (
              <button
                type="button"
                onClick={() => setSearchQuery('')}
                style={{ background: 'transparent', border: 'none', cursor: 'pointer', padding: '4px', color: 'var(--muted)' }}
              >
                <X size={14} />
              </button>
            )}
          </div>

          {/* Filters & Actions */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
            {/* Batch Copy Emails in qp_email_dispatch tab */}
            {activeTab === 'qp_email_dispatch' && (
              <button
                type="button"
                onClick={() => {
                  const emails = currentFilteredRows
                    .map(r => r.emailId)
                    .filter(e => e && e.includes('@'));
                  const uniqueEmails = Array.from(new Set(emails));
                  if (uniqueEmails.length === 0) {
                    setStatusAlert({ type: 'error', message: 'No valid college email addresses found to copy in current filter.' });
                    return;
                  }
                  navigator.clipboard.writeText(uniqueEmails.join('; '));
                  setCopiedNoticeToast(`Copied ${uniqueEmails.length} unique dispatch email addresses!`);
                  setTimeout(() => setCopiedNoticeToast(null), 3000);
                }}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '6px 12px',
                  borderRadius: '6px',
                  background: '#059669',
                  color: 'white',
                  border: 'none',
                  fontSize: '12.5px',
                  fontWeight: 700,
                  cursor: 'pointer'
                }}
                title="Copy all college emails in current dispatch report"
              >
                <Copy size={14} /> Copy Dispatch Emails
              </button>
            )}

            {/* Batch Copy All Emails in venue_email tab */}
            {activeTab === 'venue_email' && (
              <button
                type="button"
                onClick={handleCopyAllEmails}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '6px 12px',
                  borderRadius: '6px',
                  background: '#059669',
                  color: 'white',
                  border: 'none',
                  fontSize: '12.5px',
                  fontWeight: 700,
                  cursor: 'pointer'
                }}
                title="Copy all college emails separated by semicolons"
              >
                <Copy size={14} /> Copy All College Emails ({summaryStats.mappedVenues})
              </button>
            )}

            {/* Date Filter */}
            {activeTab !== 'course_summary' && filterOptions.dates.length > 0 && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Date:</span>
                <select
                  value={selectedDate}
                  onChange={(e) => { setSelectedDate(e.target.value); setCurrentPage(0); }}
                  style={{ fontSize: '12.5px', padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)', color: 'var(--ink)' }}
                >
                  <option value="ALL">All Dates ({filterOptions.dates.length})</option>
                  {filterOptions.dates.map((d, idx) => (
                    <option key={idx} value={d}>{d}</option>
                  ))}
                </select>
              </div>
            )}

            {/* Session Filter */}
            {activeTab !== 'venue_email' && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Session:</span>
                <select
                  value={selectedSession}
                  onChange={(e) => { setSelectedSession(e.target.value); setCurrentPage(0); }}
                  style={{ fontSize: '12.5px', padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)', color: 'var(--ink)' }}
                >
                  <option value="ALL">All Sessions</option>
                  <option value="FN">FN (Forenoon / Morning)</option>
                  <option value="AN">AN (Afternoon / Evening)</option>
                </select>
              </div>
            )}

            {/* Venue Filter */}
            {activeTab !== 'press_print' && activeTab !== 'course_summary' && filterOptions.venues.length > 0 && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ fontSize: '12px', color: 'var(--muted)' }}>College:</span>
                <select
                  value={selectedVenue}
                  onChange={(e) => { setSelectedVenue(e.target.value); setCurrentPage(0); }}
                  style={{ fontSize: '12.5px', padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)', color: 'var(--ink)', maxWidth: '180px' }}
                >
                  <option value="ALL">All Colleges ({filterOptions.venues.length})</option>
                  {filterOptions.venues.map((v, idx) => (
                    <option key={idx} value={v.code}>{v.name} ({v.code})</option>
                  ))}
                </select>
              </div>
            )}

            {/* Page Size */}
            {activeTab !== 'venue_packing' && (
              <select
                value={pageSize}
                onChange={(e) => { setPageSize(Number(e.target.value)); setCurrentPage(0); }}
                style={{ fontSize: '12.5px', padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)', color: 'var(--ink)' }}
              >
                <option value={25}>25 rows</option>
                <option value={50}>50 rows</option>
                <option value={100}>100 rows</option>
                <option value={0}>Show All</option>
              </select>
            )}
          </div>
        </div>

        {/* Content Body: Depending on Active Tab */}

        {/* Tab 1: Venue Packing Slips View */}
        {activeTab === 'venue_packing' && (
          <div style={{ padding: '24px', maxHeight: '70vh', overflowY: 'auto' }}>
            {Object.keys(venuePackingSlips).length === 0 ? (
              <div style={{ textAlign: 'center', padding: '60px 20px', color: 'var(--muted)' }}>
                <PackageCheck size={36} style={{ margin: '0 auto 10px', opacity: 0.4 }} />
                <p style={{ margin: 0, fontWeight: 600, fontSize: '14px' }}>No venue records found. Upload a schedule or load sample data.</p>
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
                {Object.values(venuePackingSlips)
                  .filter(v => selectedVenue === 'ALL' || v.venueCode === selectedVenue)
                  .map((venueGroup, vIdx) => (
                    <div 
                      key={vIdx} 
                      className="card" 
                      style={{ 
                        margin: 0, 
                        padding: 0, 
                        overflow: 'hidden', 
                        border: '1px solid var(--line)',
                        background: 'var(--panel)',
                        borderRadius: '10px'
                      }}
                    >
                      {/* Venue Header Card */}
                      <div style={{ 
                        padding: '14px 20px', 
                        background: 'var(--bg)', 
                        borderBottom: '1px solid var(--line)',
                        display: 'flex',
                        justifyContent: 'space-between',
                        alignItems: 'center',
                        flexWrap: 'wrap',
                        gap: '10px'
                      }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                          <span style={{ 
                            padding: '4px 10px', 
                            borderRadius: '6px', 
                            background: 'var(--accent)', 
                            color: 'white', 
                            fontWeight: 800, 
                            fontFamily: 'monospace',
                            fontSize: '13px'
                          }}>
                            {venueGroup.venueCode}
                          </span>
                          <div>
                            <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 800, color: 'var(--ink)' }}>
                              {venueGroup.venueName}
                            </h3>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginTop: '2px', fontSize: '12px', color: 'var(--muted)' }}>
                              <span>Center: {venueGroup.centerCode} ({venueGroup.centerName || 'Main'})</span>
                              <span>•</span>
                              {venueGroup.collegeEmail ? (
                                <a 
                                  href={`mailto:${venueGroup.collegeEmail}`}
                                  style={{ color: '#059669', textDecoration: 'none', display: 'inline-flex', alignItems: 'center', gap: '4px' }}
                                >
                                  <Mail size={12} /> {venueGroup.collegeEmail}
                                </a>
                              ) : (
                                <span style={{ color: '#dc2626' }}>No Email Configured</span>
                              )}
                            </div>
                          </div>
                        </div>

                        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', fontSize: '13px' }}>
                          <span style={{ color: 'var(--muted)' }}>
                            Exams: <strong style={{ color: 'var(--ink)' }}>{venueGroup.items.length}</strong>
                          </span>
                          <span style={{ color: 'var(--muted)' }}>
                            Candidates: <strong style={{ color: 'var(--ink)' }}>{venueGroup.totalStudents.toLocaleString()}</strong>
                          </span>
                          <span style={{ 
                            padding: '4px 12px', 
                            borderRadius: '6px', 
                            background: 'rgba(5, 150, 105, 0.1)', 
                            color: '#059669', 
                            fontWeight: 800,
                            border: '1px solid rgba(5, 150, 105, 0.2)'
                          }}>
                            Total QP Packed: {venueGroup.totalQp.toLocaleString()}
                          </span>
                        </div>
                      </div>

                      {/* Checklist Table */}
                      <div style={{ overflowX: 'auto' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px', textAlign: 'left' }}>
                          <thead>
                            <tr style={{ background: 'var(--panel)', borderBottom: '1px solid var(--line)', color: 'var(--muted)', fontSize: '11px', textTransform: 'uppercase' }}>
                              <th style={{ padding: '10px 14px', width: '40px', textAlign: 'center' }}>Packed</th>
                              <th style={{ padding: '10px 14px' }}>Date</th>
                              <th style={{ padding: '10px 14px' }}>Session / Time</th>
                              <th style={{ padding: '10px 14px' }}>Course Code</th>
                              <th style={{ padding: '10px 14px' }}>Course Name</th>
                              <th style={{ padding: '10px 14px', textAlign: 'center' }}>Takers</th>
                              <th style={{ padding: '10px 14px', textAlign: 'center' }}>Packets of 20</th>
                              <th style={{ padding: '10px 14px', textAlign: 'center' }}>Packets of 10</th>
                              <th style={{ padding: '10px 14px', textAlign: 'center' }}>Packets of 5</th>
                              <th style={{ padding: '10px 14px', textAlign: 'center' }}>Loose / Buffer</th>
                              <th style={{ padding: '10px 14px', textAlign: 'right' }}>Total QP</th>
                            </tr>
                          </thead>
                          <tbody>
                            {venueGroup.items.map((item, idx) => {
                              const itemKey = `${venueGroup.venueCode}_${item.examDate}_${item.courseCode}`;
                              const isPacked = packedChecklist[itemKey];

                              return (
                                <tr 
                                  key={idx} 
                                  style={{ 
                                    borderBottom: '1px solid var(--line)', 
                                    background: isPacked ? 'rgba(5, 150, 105, 0.04)' : idx % 2 === 0 ? 'var(--panel)' : 'var(--bg)',
                                    transition: 'background 0.15s ease'
                                  }}
                                >
                                  <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                                    <input
                                      type="checkbox"
                                      checked={!!isPacked}
                                      onChange={() => handleToggleChecklist(itemKey)}
                                      style={{ cursor: 'pointer', transform: 'scale(1.15)' }}
                                      title="Mark as sealed and packed in venue parcel"
                                    />
                                  </td>
                                  <td style={{ padding: '10px 14px', fontWeight: 600, color: 'var(--ink)' }}>
                                    {item.examDate}
                                  </td>
                                  <td style={{ padding: '10px 14px' }}>
                                    <span style={{
                                      padding: '2px 6px',
                                      borderRadius: '4px',
                                      fontSize: '11px',
                                      fontWeight: 700,
                                      background: item.session === 'FN' ? 'rgba(2, 132, 199, 0.1)' : 'rgba(217, 119, 6, 0.1)',
                                      color: item.session === 'FN' ? '#0284c7' : '#d97706',
                                      marginRight: '6px'
                                    }}>
                                      {item.session}
                                    </span>
                                    <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                                      {item.startTime || '—'} - {item.endTime || '—'}
                                    </span>
                                  </td>
                                  <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontWeight: 700, color: 'var(--ink)' }}>
                                    {item.courseCode}
                                  </td>
                                  <td style={{ padding: '10px 14px', maxWidth: '280px', color: 'var(--ink)' }}>
                                    {item.courseName}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 700 }}>
                                    {item.studentCount}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                                    {item.p20 > 0 ? (
                                      <span style={{ padding: '2px 8px', borderRadius: '4px', background: 'rgba(23, 107, 135, 0.1)', color: 'var(--accent)', fontWeight: 700 }}>
                                        {item.p20} pkts
                                      </span>
                                    ) : '—'}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                                    {item.p10 > 0 ? (
                                      <span style={{ padding: '2px 8px', borderRadius: '4px', background: 'rgba(124, 58, 237, 0.1)', color: '#7c3aed', fontWeight: 700 }}>
                                        {item.p10} pkts
                                      </span>
                                    ) : '—'}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                                    {item.p5 > 0 ? (
                                      <span style={{ padding: '2px 8px', borderRadius: '4px', background: 'rgba(217, 119, 6, 0.1)', color: '#d97706', fontWeight: 700 }}>
                                        {item.p5} pkts
                                      </span>
                                    ) : '—'}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                                    {item.singles > 0 || item.buffer > 0 ? (
                                      <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                                        {item.singles} loose {item.buffer > 0 && `(+${item.buffer} buf)`}
                                      </span>
                                    ) : '—'}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 800, color: isPacked ? '#059669' : 'var(--ink)' }}>
                                    {item.totalQp}
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  ))}
              </div>
            )}
          </div>
        )}

        {/* Tab 2: College Email Dispatch View */}
        {activeTab === 'venue_email' && (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', textAlign: 'left' }}>
              <thead>
                <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)', color: 'var(--muted)', fontSize: '11.5px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  <th style={{ padding: '12px 16px' }}>Venue Code</th>
                  <th style={{ padding: '12px 16px' }}>College / Center Name</th>
                  <th style={{ padding: '12px 16px' }}>Official College Email</th>
                  <th style={{ padding: '12px 16px', textAlign: 'center' }}>Scheduled Exams</th>
                  <th style={{ padding: '12px 16px', textAlign: 'center' }}>Total Candidates</th>
                  <th style={{ padding: '12px 16px', textAlign: 'center' }}>Total QP Copies</th>
                  <th style={{ padding: '12px 16px', textAlign: 'right' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {paginatedRows.length === 0 ? (
                  <tr>
                    <td colSpan={7} style={{ textAlign: 'center', padding: '60px 20px', color: 'var(--muted)' }}>
                      No college venues match your search criteria.
                    </td>
                  </tr>
                ) : (
                  paginatedRows.map((venue, idx) => (
                    <tr 
                      key={idx} 
                      style={{ 
                        borderBottom: '1px solid var(--line)',
                        background: idx % 2 === 0 ? 'var(--panel)' : 'var(--bg)'
                      }}
                    >
                      <td style={{ padding: '12px 16px', fontFamily: 'monospace', fontWeight: 700, color: 'var(--accent)' }}>
                        {venue.venueCode}
                      </td>
                      <td style={{ padding: '12px 16px' }}>
                        <div style={{ fontWeight: 700, color: 'var(--ink)' }}>{venue.venueName}</div>
                        <div style={{ fontSize: '11.5px', color: 'var(--muted)' }}>Center Code: {venue.centerCode || 'Main'}</div>
                      </td>
                      <td style={{ padding: '12px 16px' }}>
                        {venue.hasEmail ? (
                          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                            <a 
                              href={`mailto:${venue.collegeEmail}?subject=${encodeURIComponent(`Kannur University Examinations - ${fileMeta?.event || 'Question Paper Schedule'}`)}`}
                              style={{ 
                                color: '#059669', 
                                textDecoration: 'none', 
                                fontFamily: 'monospace', 
                                fontSize: '12.5px',
                                fontWeight: 600,
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '4px'
                              }}
                            >
                              <Mail size={13} /> {venue.collegeEmail}
                            </a>
                            <button
                              type="button"
                              onClick={() => {
                                navigator.clipboard.writeText(venue.collegeEmail);
                                setCopiedNoticeToast(`Copied ${venue.collegeEmail}`);
                                setTimeout(() => setCopiedNoticeToast(null), 2500);
                              }}
                              style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '2px', color: 'var(--muted)' }}
                              title="Copy email address"
                            >
                              <Copy size={13} />
                            </button>
                          </div>
                        ) : (
                          <button
                            type="button"
                            onClick={() => {
                              setNewCollegeForm({ code: venue.venueCode, name: venue.venueName, email: '' });
                              setShowEmailDirectoryModal(true);
                            }}
                            style={{
                              padding: '2px 8px',
                              borderRadius: '4px',
                              background: 'rgba(239, 68, 68, 0.1)',
                              color: '#dc2626',
                              border: '1px solid rgba(239, 68, 68, 0.2)',
                              fontSize: '11.5px',
                              fontWeight: 600,
                              cursor: 'pointer'
                            }}
                          >
                            + Assign Email
                          </button>
                        )}
                      </td>
                      <td style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 700 }}>
                        {venue.examsCount} slots
                      </td>
                      <td style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 700 }}>
                        {venue.totalStudents.toLocaleString()}
                      </td>
                      <td style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 800, color: '#e11d48' }}>
                        {venue.totalQp.toLocaleString()}
                      </td>
                      <td style={{ padding: '12px 16px', textAlign: 'right' }}>
                        <div style={{ display: 'inline-flex', gap: '8px' }}>
                          <button
                            type="button"
                            onClick={() => setSelectedVenueForEmailDraft(venue)}
                            className="secondary"
                            style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '5px 10px', fontSize: '12px' }}
                            title="Preview formal circular & email text"
                          >
                            <FileText size={13} color="var(--accent)" /> Circular Draft
                          </button>
                          {venue.hasEmail && (
                            <a
                              href={`mailto:${venue.collegeEmail}?subject=${encodeURIComponent(`Kannur University Examinations Schedule - ${fileMeta?.event || 'Intimation'}`)}`}
                              className="secondary"
                              style={{ display: 'inline-flex', alignItems: 'center', gap: '5px', padding: '5px 10px', fontSize: '12px', textDecoration: 'none' }}
                              title="Open default email application"
                            >
                              <Send size={13} color="#059669" /> Mail
                            </a>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* Tab 3, 4, 5, 6: Master Schedule, Daily Press, Course Summary, Raw Data */}
        {activeTab !== 'venue_packing' && activeTab !== 'venue_email' && (
          <div style={{ overflowX: 'auto' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', textAlign: 'left' }}>
              <thead>
                <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)', color: 'var(--muted)', fontSize: '11.5px', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  {/* QP Email Dispatch Report Tab Headers */}
                  {activeTab === 'qp_email_dispatch' && (
                    <>
                      <th style={{ padding: '12px 16px' }}>Exam Date</th>
                      <th style={{ padding: '12px 16px' }}>Course  Code</th>
                      <th style={{ padding: '12px 16px' }}>Course  Name</th>
                      <th style={{ padding: '12px 16px' }}>Venue Code</th>
                      <th style={{ padding: '12px 16px' }}>Venue Name</th>
                      <th style={{ padding: '12px 16px' }}>Email ID</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Session</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Candidates</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Packets (20/10/5)</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Buffer</th>
                      <th style={{ padding: '12px 16px', textAlign: 'right' }}>Total QP Copies</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Action</th>
                    </>
                  )}

                  {/* Master Schedule Tab Headers */}
                  {activeTab === 'consolidated' && (
                    <>
                      <th style={{ padding: '12px 16px' }}>Date</th>
                      <th style={{ padding: '12px 16px' }}>Session / Time</th>
                      <th style={{ padding: '12px 16px' }}>Venue College</th>
                      <th style={{ padding: '12px 16px' }}>College Email</th>
                      <th style={{ padding: '12px 16px' }}>Course Code</th>
                      <th style={{ padding: '12px 16px' }}>Course Name</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Candidates</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Sealed Packets</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Buffer</th>
                      <th style={{ padding: '12px 16px', textAlign: 'right' }}>Total QP</th>
                    </>
                  )}

                  {/* Daily Press Print Order Tab Headers */}
                  {activeTab === 'press_print' && (
                    <>
                      <th style={{ padding: '12px 16px' }}>Exam Date</th>
                      <th style={{ padding: '12px 16px' }}>Session / Timing</th>
                      <th style={{ padding: '12px 16px' }}>Course Code</th>
                      <th style={{ padding: '12px 16px' }}>Course Name</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Venues</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Total Takers</th>
                      <th style={{ padding: '12px 16px', textAlign: 'right' }}>Press Print Order</th>
                    </>
                  )}

                  {/* Course Master Summary Tab Headers */}
                  {activeTab === 'course_summary' && (
                    <>
                      <th style={{ padding: '12px 16px' }}>Course Code</th>
                      <th style={{ padding: '12px 16px' }}>Course Name</th>
                      <th style={{ padding: '12px 16px' }}>Exam Dates</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Session</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Venues</th>
                      <th style={{ padding: '12px 16px', textAlign: 'center' }}>Total Takers</th>
                      <th style={{ padding: '12px 16px', textAlign: 'right' }}>Total QP Copies</th>
                    </>
                  )}

                  {/* Raw Source Rows Tab Headers */}
                  {activeTab === 'raw' && (
                    <>
                      <th style={{ padding: '12px 16px' }}>#</th>
                      <th style={{ padding: '12px 16px' }}>Event Name</th>
                      <th style={{ padding: '12px 16px' }}>Venue</th>
                      <th style={{ padding: '12px 16px' }}>College Email</th>
                      <th style={{ padding: '12px 16px' }}>Date</th>
                      <th style={{ padding: '12px 16px' }}>Time</th>
                      <th style={{ padding: '12px 16px' }}>Course</th>
                      <th style={{ padding: '12px 16px' }}>Academic Programme</th>
                      <th style={{ padding: '12px 16px', textAlign: 'right' }}>Count</th>
                    </>
                  )}
                </tr>
              </thead>
              <tbody>
                {paginatedRows.length === 0 ? (
                  <tr>
                    <td colSpan={12} style={{ textAlign: 'center', padding: '60px 20px', color: 'var(--muted)' }}>
                      No records found. Upload an examination schedule or click "Load Kannur Sample Data".
                    </td>
                  </tr>
                ) : (
                  paginatedRows.map((row, idx) => (
                    <tr 
                      key={idx} 
                      style={{ 
                        borderBottom: '1px solid var(--line)',
                        background: idx % 2 === 0 ? 'var(--panel)' : 'var(--bg)'
                      }}
                    >
                      {/* QP Email Dispatch Cells */}
                      {activeTab === 'qp_email_dispatch' && (
                        <>
                          <td style={{ padding: '12px 16px', fontWeight: 600, color: 'var(--ink)' }}>
                            {row.examDate}
                          </td>
                          <td style={{ padding: '12px 16px', fontFamily: 'monospace', fontWeight: 700, color: 'var(--accent)' }}>
                            {row.courseCode}
                          </td>
                          <td style={{ padding: '12px 16px', maxWidth: '240px' }}>
                            <div style={{ color: 'var(--ink)', fontWeight: 600 }}>{row.courseName}</div>
                            {row.programsFormatted && (
                              <div style={{ fontSize: '11px', color: 'var(--muted)', marginTop: '2px' }} title={row.programsFormatted}>
                                {row.programsFormatted.length > 40 ? `${row.programsFormatted.substring(0, 40)}...` : row.programsFormatted}
                              </div>
                            )}
                          </td>
                          <td style={{ padding: '12px 16px', fontFamily: 'monospace', fontWeight: 700, color: 'var(--ink)' }}>
                            {row.venueCode}
                          </td>
                          <td style={{ padding: '12px 16px', maxWidth: '220px' }}>
                            <span style={{ fontWeight: 600, color: 'var(--ink)' }}>{row.venueName}</span>
                          </td>
                          <td style={{ padding: '12px 16px' }}>
                            {row.hasEmail ? (
                              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                <a 
                                  href={`mailto:${row.emailId}?subject=${encodeURIComponent(`Kannur University Question Paper Dispatch - ${row.courseCode} (${row.examDate} ${row.session})`)}`}
                                  style={{ color: '#059669', textDecoration: 'none', fontSize: '12px', fontFamily: 'monospace', fontWeight: 600 }}
                                  title="Dispatch QP Email"
                                >
                                  {row.emailId}
                                </a>
                                <button
                                  type="button"
                                  onClick={() => {
                                    navigator.clipboard.writeText(row.emailId);
                                    setCopiedNoticeToast(`Copied ${row.emailId}`);
                                    setTimeout(() => setCopiedNoticeToast(null), 2500);
                                  }}
                                  style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '2px', color: 'var(--muted)' }}
                                  title="Copy Email ID"
                                >
                                  <Copy size={12} />
                                </button>
                              </div>
                            ) : (
                              <span 
                                onClick={() => setShowEmailDirectoryModal(true)}
                                style={{ fontSize: '11px', color: '#dc2626', cursor: 'pointer', background: 'rgba(239, 68, 68, 0.08)', padding: '2px 6px', borderRadius: '4px' }}
                                title="Click to map college email"
                              >
                                Unmapped
                              </span>
                            )}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <span style={{
                              padding: '2px 6px',
                              borderRadius: '4px',
                              fontSize: '11px',
                              fontWeight: 700,
                              background: row.session === 'FN' ? 'rgba(2, 132, 199, 0.1)' : 'rgba(217, 119, 6, 0.1)',
                              color: row.session === 'FN' ? '#0284c7' : '#d97706'
                            }}>
                              {row.session}
                            </span>
                            {row.startTime && (
                              <div style={{ fontSize: '10.5px', color: 'var(--muted)', marginTop: '2px' }}>
                                {row.startTime}
                              </div>
                            )}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 700 }}>
                            {row.studentCount}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <div style={{ display: 'flex', gap: '4px', justifyContent: 'center', flexWrap: 'wrap' }}>
                              {row.p20 > 0 && (
                                <span style={{ padding: '1px 6px', borderRadius: '4px', background: 'rgba(23, 107, 135, 0.1)', color: 'var(--accent)', fontSize: '11px', fontWeight: 700 }}>
                                  {row.p20}×20
                                </span>
                              )}
                              {row.p10 > 0 && (
                                <span style={{ padding: '1px 6px', borderRadius: '4px', background: 'rgba(124, 58, 237, 0.1)', color: '#7c3aed', fontSize: '11px', fontWeight: 700 }}>
                                  {row.p10}×10
                                </span>
                              )}
                              {row.p5 > 0 && (
                                <span style={{ padding: '1px 6px', borderRadius: '4px', background: 'rgba(217, 119, 6, 0.1)', color: '#d97706', fontSize: '11px', fontWeight: 700 }}>
                                  {row.p5}×5
                                </span>
                              )}
                              {row.singles > 0 && (
                                <span style={{ padding: '1px 6px', borderRadius: '4px', background: 'var(--line)', color: 'var(--muted)', fontSize: '11px', fontWeight: 600 }}>
                                  {row.singles} loose
                                </span>
                              )}
                            </div>
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
                            +{row.buffer}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'right', fontWeight: 800, color: '#e11d48' }}>
                            {row.totalQp}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            {row.hasEmail ? (
                              <a
                                href={`mailto:${row.emailId}?subject=${encodeURIComponent(`Kannur University Question Paper Dispatch - ${row.courseCode} (${row.examDate} ${row.session})`)}`}
                                className="secondary"
                                style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', padding: '4px 8px', fontSize: '11.5px', textDecoration: 'none' }}
                                title="Open mail client"
                              >
                                <Send size={12} color="#059669" /> Mail
                              </a>
                            ) : (
                              <span style={{ fontSize: '11px', color: 'var(--muted)' }}>—</span>
                            )}
                          </td>
                        </>
                      )}

                      {/* Master Schedule Cells */}
                      {activeTab === 'consolidated' && (
                        <>
                          <td style={{ padding: '12px 16px', fontWeight: 600, color: 'var(--ink)' }}>
                            {row.examDate}
                          </td>
                          <td style={{ padding: '12px 16px' }}>
                            <span style={{
                              padding: '2px 6px',
                              borderRadius: '4px',
                              fontSize: '11px',
                              fontWeight: 700,
                              background: row.session === 'FN' ? 'rgba(2, 132, 199, 0.1)' : 'rgba(217, 119, 6, 0.1)',
                              color: row.session === 'FN' ? '#0284c7' : '#d97706',
                              marginRight: '6px'
                            }}>
                              {row.session}
                            </span>
                            <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                              {row.startTime || '—'}
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px' }}>
                            <span style={{ fontWeight: 700, color: 'var(--ink)' }}>{row.venueName}</span>
                            <span style={{ marginLeft: '6px', fontSize: '11px', fontFamily: 'monospace', color: 'var(--muted)' }}>
                              ({row.venueCode})
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px' }}>
                            {row.hasEmail ? (
                              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                <a 
                                  href={`mailto:${row.collegeEmail}`}
                                  style={{ color: '#059669', textDecoration: 'none', fontSize: '12px', fontFamily: 'monospace' }}
                                  title="Send email"
                                >
                                  {row.collegeEmail}
                                </a>
                                <button
                                  type="button"
                                  onClick={() => {
                                    navigator.clipboard.writeText(row.collegeEmail);
                                    setCopiedNoticeToast(`Copied ${row.collegeEmail}`);
                                    setTimeout(() => setCopiedNoticeToast(null), 2500);
                                  }}
                                  style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '2px', color: 'var(--muted)' }}
                                  title="Copy email"
                                >
                                  <Copy size={12} />
                                </button>
                              </div>
                            ) : (
                              <span 
                                onClick={() => setShowEmailDirectoryModal(true)}
                                style={{ fontSize: '11px', color: '#dc2626', cursor: 'pointer', background: 'rgba(239, 68, 68, 0.08)', padding: '2px 6px', borderRadius: '4px' }}
                                title="Click to configure email in directory"
                              >
                                Unmapped
                              </span>
                            )}
                          </td>
                          <td style={{ padding: '12px 16px', fontFamily: 'monospace', fontWeight: 700, color: 'var(--ink)' }}>
                            {row.courseCode}
                          </td>
                          <td style={{ padding: '12px 16px', maxWidth: '240px' }}>
                            <div style={{ color: 'var(--ink)', fontWeight: 500 }}>{row.courseName}</div>
                            {row.programsCount > 1 && (
                              <div style={{ fontSize: '11px', color: 'var(--accent)', marginTop: '2px' }}>
                                Sum of {row.programsCount} academic programmes
                              </div>
                            )}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 700 }}>
                            {row.studentCount}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <div style={{ display: 'flex', gap: '4px', justifyContent: 'center', flexWrap: 'wrap' }}>
                              {row.p20 > 0 && (
                                <span style={{ padding: '1px 6px', borderRadius: '4px', background: 'rgba(23, 107, 135, 0.1)', color: 'var(--accent)', fontSize: '11px', fontWeight: 700 }}>
                                  {row.p20}×20
                                </span>
                              )}
                              {row.p10 > 0 && (
                                <span style={{ padding: '1px 6px', borderRadius: '4px', background: 'rgba(124, 58, 237, 0.1)', color: '#7c3aed', fontSize: '11px', fontWeight: 700 }}>
                                  {row.p10}×10
                                </span>
                              )}
                              {row.p5 > 0 && (
                                <span style={{ padding: '1px 6px', borderRadius: '4px', background: 'rgba(217, 119, 6, 0.1)', color: '#d97706', fontSize: '11px', fontWeight: 700 }}>
                                  {row.p5}×5
                                </span>
                              )}
                              {row.singles > 0 && (
                                <span style={{ padding: '1px 6px', borderRadius: '4px', background: 'var(--line)', color: 'var(--muted)', fontSize: '11px', fontWeight: 600 }}>
                                  {row.singles} loose
                                </span>
                              )}
                            </div>
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
                            +{row.buffer}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'right', fontWeight: 800, color: '#e11d48' }}>
                            {row.totalQp}
                          </td>
                        </>
                      )}

                      {/* Daily Press Print Order Cells */}
                      {activeTab === 'press_print' && (
                        <>
                          <td style={{ padding: '12px 16px', fontWeight: 600, color: 'var(--ink)' }}>
                            {row.examDate}
                          </td>
                          <td style={{ padding: '12px 16px' }}>
                            <span style={{
                              padding: '2px 6px',
                              borderRadius: '4px',
                              fontSize: '11px',
                              fontWeight: 700,
                              background: row.session === 'FN' ? 'rgba(2, 132, 199, 0.1)' : 'rgba(217, 119, 6, 0.1)',
                              color: row.session === 'FN' ? '#0284c7' : '#d97706',
                              marginRight: '6px'
                            }}>
                              {row.session}
                            </span>
                            <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                              {row.startTime || '—'} - {row.endTime || '—'}
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px', fontFamily: 'monospace', fontWeight: 700, color: 'var(--ink)' }}>
                            {row.courseCode}
                          </td>
                          <td style={{ padding: '12px 16px', color: 'var(--ink)' }}>
                            {row.courseName}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <span style={{ padding: '2px 8px', borderRadius: '10px', background: 'var(--line)', fontSize: '11.5px', fontWeight: 600 }}>
                              {row.venuesCount} venues
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 700 }}>
                            {row.totalStudents}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'right', fontWeight: 800, color: '#e11d48', fontSize: '14px' }}>
                            {row.totalQp} copies
                          </td>
                        </>
                      )}

                      {/* Course Master Summary Cells */}
                      {activeTab === 'course_summary' && (
                        <>
                          <td style={{ padding: '12px 16px', fontFamily: 'monospace', fontWeight: 700, color: 'var(--ink)' }}>
                            {row.courseCode}
                          </td>
                          <td style={{ padding: '12px 16px', color: 'var(--ink)', fontWeight: 600 }}>
                            {row.courseName}
                          </td>
                          <td style={{ padding: '12px 16px', color: 'var(--muted)', fontSize: '12px' }}>
                            {row.datesFormatted}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            <span style={{
                              padding: '2px 6px',
                              borderRadius: '4px',
                              fontSize: '11px',
                              fontWeight: 700,
                              background: row.session === 'FN' ? 'rgba(2, 132, 199, 0.1)' : 'rgba(217, 119, 6, 0.1)',
                              color: row.session === 'FN' ? '#0284c7' : '#d97706'
                            }}>
                              {row.session}
                            </span>
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                            {row.venueCount}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'center', fontWeight: 700 }}>
                            {row.totalStudents}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'right', fontWeight: 800, color: '#e11d48' }}>
                            {row.totalQp}
                          </td>
                        </>
                      )}

                      {/* Raw Source Rows Cells */}
                      {activeTab === 'raw' && (
                        <>
                          <td style={{ padding: '12px 16px', color: 'var(--muted)', fontSize: '11.5px' }}>
                            {idx + 1}
                          </td>
                          <td style={{ padding: '12px 16px', fontSize: '12px' }}>
                            {row.eventName || '—'}
                          </td>
                          <td style={{ padding: '12px 16px' }}>
                            <span style={{ fontWeight: 600 }}>{row.venueName}</span>
                            <span style={{ marginLeft: '6px', color: 'var(--muted)', fontSize: '11px' }}>({row.venueCode})</span>
                          </td>
                          <td style={{ padding: '12px 16px', fontSize: '12px', fontFamily: 'monospace', color: '#059669' }}>
                            {row.collegeEmail || '—'}
                          </td>
                          <td style={{ padding: '12px 16px', fontSize: '12px' }}>
                            {row.examDate}
                          </td>
                          <td style={{ padding: '12px 16px', fontSize: '11.5px', color: 'var(--muted)' }}>
                            {row.startTime}
                          </td>
                          <td style={{ padding: '12px 16px' }}>
                            <div style={{ fontFamily: 'monospace', fontWeight: 700 }}>{row.courseCode}</div>
                            <div style={{ fontSize: '11.5px', color: 'var(--muted)' }}>{row.courseName}</div>
                          </td>
                          <td style={{ padding: '12px 16px', fontSize: '11.5px', maxWidth: '280px', color: 'var(--muted)' }}>
                            {row.programName || '—'}
                          </td>
                          <td style={{ padding: '12px 16px', textAlign: 'right', fontWeight: 700 }}>
                            {row.studentCount}
                          </td>
                        </>
                      )}
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        )}

        {/* Pagination Bar */}
        {activeTab !== 'venue_packing' && pageSize > 0 && totalPages > 1 && (
          <div style={{
            padding: '12px 20px',
            borderTop: '1px solid var(--line)',
            background: 'var(--bg)',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            fontSize: '12.5px',
            color: 'var(--muted)',
            flexWrap: 'wrap',
            gap: '10px'
          }}>
            <div>
              Showing {currentPage * pageSize + 1} to {Math.min((currentPage + 1) * pageSize, currentFilteredRows.length)} of {currentFilteredRows.length} entries
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <button
                type="button"
                onClick={() => setCurrentPage(p => Math.max(0, p - 1))}
                disabled={currentPage === 0}
                className="secondary"
                style={{ padding: '4px 8px', fontSize: '12px', display: 'inline-flex', alignItems: 'center' }}
              >
                <ChevronLeft size={14} /> Prev
              </button>
              <span>Page {currentPage + 1} of {totalPages}</span>
              <button
                type="button"
                onClick={() => setCurrentPage(p => Math.min(totalPages - 1, p + 1))}
                disabled={currentPage >= totalPages - 1}
                className="secondary"
                style={{ padding: '4px 8px', fontSize: '12px', display: 'inline-flex', alignItems: 'center' }}
              >
                Next <ChevronRight size={14} />
              </button>
            </div>
          </div>
        )}
      </div>

      {/* MODAL 1: Official Email Circular Notification Draft */}
      {selectedVenueForEmailDraft && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0,0,0,0.6)',
          backdropFilter: 'blur(3px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          padding: '20px'
        }}>
          <div className="card" style={{
            maxWidth: '750px',
            width: '100%',
            maxHeight: '90vh',
            overflowY: 'auto',
            background: 'var(--panel)',
            padding: '24px',
            borderRadius: '12px',
            boxShadow: '0 20px 40px rgba(0,0,0,0.3)',
            border: '1px solid var(--line)'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', borderBottom: '1px solid var(--line)', paddingBottom: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <Mail size={22} color="#059669" />
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 800 }}>
                  Official Intimation Circular Draft
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setSelectedVenueForEmailDraft(null)}
                style={{ background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer', padding: 0 }}
              >
                <X size={18} />
              </button>
            </div>

            <div style={{ background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px', padding: '16px', fontSize: '13px', lineHeight: '1.6', fontFamily: 'monospace', whiteSpace: 'pre-wrap' }}>
{`KANNUR UNIVERSITY - EXAMINATION BRANCH (CONFIDENTIAL SECTION)
======================================================================
CONFIDENTIAL & MOST URGENT

To,
The Principal / Chief Superintendent,
${selectedVenueForEmailDraft.venueName} (Venue Code: ${selectedVenueForEmailDraft.venueCode})
Official Email: ${selectedVenueForEmailDraft.collegeEmail || '[EMAIL NOT CONFIGURED]'}

Subject: Conduct of ${fileMeta?.event || 'University Examinations'} - Examination Schedule & Question Paper Allotment - Intimation - Reg.

Sir / Madam,
I am directed to inform you that the following examinations are scheduled to be conducted at your center. Sealed Question Paper parcels are dispatched as per the allotment schedule below:

${selectedVenueForEmailDraft.exams.map(e => `• Date: ${e.examDate} | Session: ${e.session} (${e.startTime || '—'} to ${e.endTime || '—'})
  Course: [${e.courseCode}] ${e.courseName}
  Registered Candidates: ${e.studentCount}
  QP Packets: ${e.p20 > 0 ? `${e.p20}x20 ` : ''}${e.p10 > 0 ? `${e.p10}x10 ` : ''}${e.p5 > 0 ? `${e.p5}x5 ` : ''}${e.singles > 0 ? `${e.singles} loose` : ''} (+${e.buffer} reserve)
  Total QP Dispatched: ${e.totalQp} copies
`).join('\n')}
----------------------------------------------------------------------
TOTAL CANDIDATES: ${selectedVenueForEmailDraft.totalStudents}
TOTAL QUESTION PAPERS ALLOTTED: ${selectedVenueForEmailDraft.totalQp} (including buffer)
----------------------------------------------------------------------

You are requested to verify the sealed question paper packets upon arrival against this schedule and confirm safe receipt to the Confidential Section immediately.

Yours faithfully,
Controller of Examinations / Confidential Section
Kannur University`}
            </div>

            <div style={{ marginTop: '20px', display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
              <button
                type="button"
                onClick={() => {
                  const draftText = `KANNUR UNIVERSITY - EXAMINATION BRANCH (CONFIDENTIAL SECTION)\nTo: The Principal, ${selectedVenueForEmailDraft.venueName}\n\nSchedule of Examinations:\n${selectedVenueForEmailDraft.exams.map(e => `${e.examDate} (${e.session}): [${e.courseCode}] ${e.courseName} - Takers: ${e.studentCount}, QP: ${e.totalQp}`).join('\n')}\n\nTotal QP: ${selectedVenueForEmailDraft.totalQp}`;
                  navigator.clipboard.writeText(draftText);
                  setCopiedNoticeToast(`Copied circular notification for ${selectedVenueForEmailDraft.venueName}`);
                  setTimeout(() => setCopiedNoticeToast(null), 3000);
                }}
                className="secondary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 16px', fontSize: '13px' }}
              >
                <Copy size={14} /> Copy Letter Text
              </button>

              {selectedVenueForEmailDraft.hasEmail && (
                <a
                  href={`mailto:${selectedVenueForEmailDraft.collegeEmail}?subject=${encodeURIComponent(`Kannur University Examinations Schedule - ${fileMeta?.event || 'Intimation'}`)}`}
                  className="primary"
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 16px', fontSize: '13px', textDecoration: 'none', background: '#059669', borderColor: '#059669' }}
                >
                  <Send size={14} /> Open in Mail App
                </a>
              )}

              <button
                type="button"
                onClick={() => setSelectedVenueForEmailDraft(null)}
                style={{ padding: '8px 16px', fontSize: '13px' }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: College Email Directory Manager */}
      {showEmailDirectoryModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0,0,0,0.6)',
          backdropFilter: 'blur(3px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          padding: '20px'
        }}>
          <div className="card" style={{
            maxWidth: '840px',
            width: '100%',
            maxHeight: '90vh',
            overflowY: 'auto',
            background: 'var(--panel)',
            padding: '24px',
            borderRadius: '12px',
            boxShadow: '0 20px 40px rgba(0,0,0,0.3)',
            border: '1px solid var(--line)'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', borderBottom: '1px solid var(--line)', paddingBottom: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <Building2 size={22} color="var(--accent)" />
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 800 }}>
                  Kannur University College Email Directory ({emailDirectory.length})
                </h3>
              </div>
              <button
                type="button"
                onClick={() => { setShowEmailDirectoryModal(false); setEditingCollegeIndex(null); }}
                style={{ background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer', padding: 0 }}
              >
                <X size={18} />
              </button>
            </div>

            {/* Add / Edit College Form */}
            <form onSubmit={handleAddCollegeToDirectory} style={{
              background: 'var(--bg)',
              padding: '14px',
              borderRadius: '8px',
              border: '1px solid var(--line)',
              marginBottom: '16px',
              display: 'flex',
              gap: '10px',
              flexWrap: 'wrap',
              alignItems: 'flex-end'
            }}>
              <div style={{ flex: '1 1 90px' }}>
                <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, color: 'var(--muted)', marginBottom: '4px' }}>
                  College / Venue Code *
                </label>
                <input
                  type="text"
                  placeholder="e.g. 209 or PP"
                  value={newCollegeForm.code}
                  onChange={(e) => setNewCollegeForm(prev => ({ ...prev, code: e.target.value }))}
                  required
                  style={{ width: '100%', padding: '6px 10px', fontSize: '12.5px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', color: 'var(--ink)' }}
                />
              </div>

              <div style={{ flex: '2 1 200px' }}>
                <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, color: 'var(--muted)', marginBottom: '4px' }}>
                  College Name
                </label>
                <input
                  type="text"
                  placeholder="e.g. Payyannur College, Payyannur"
                  value={newCollegeForm.name}
                  onChange={(e) => setNewCollegeForm(prev => ({ ...prev, name: e.target.value }))}
                  style={{ width: '100%', padding: '6px 10px', fontSize: '12.5px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', color: 'var(--ink)' }}
                />
              </div>

              <div style={{ flex: '2 1 220px' }}>
                <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, color: 'var(--muted)', marginBottom: '4px' }}>
                  Official Email Address *
                </label>
                <input
                  type="email"
                  placeholder="e.g. payyannurcollege@gmail.com"
                  value={newCollegeForm.email}
                  onChange={(e) => setNewCollegeForm(prev => ({ ...prev, email: e.target.value }))}
                  required
                  style={{ width: '100%', padding: '6px 10px', fontSize: '12.5px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', color: 'var(--ink)' }}
                />
              </div>

              <div style={{ display: 'flex', gap: '6px' }}>
                <button
                  type="submit"
                  style={{
                    padding: '7px 16px',
                    borderRadius: '4px',
                    background: 'var(--accent)',
                    color: 'white',
                    border: 'none',
                    fontSize: '12.5px',
                    fontWeight: 700,
                    cursor: 'pointer'
                  }}
                >
                  {editingCollegeIndex !== null ? 'Update College' : '+ Add to Directory'}
                </button>
                {editingCollegeIndex !== null && (
                  <button
                    type="button"
                    onClick={() => { setEditingCollegeIndex(null); setNewCollegeForm({ code: '', name: '', email: '' }); }}
                    className="secondary"
                    style={{ padding: '7px 12px', fontSize: '12px' }}
                  >
                    Cancel
                  </button>
                )}
              </div>
            </form>

            {/* Search within directory & Upload/Reset Options */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', flexWrap: 'wrap', gap: '10px' }}>
              <input
                type="text"
                placeholder="Search college directory by code, name, or email..."
                value={directorySearchQuery}
                onChange={(e) => setDirectorySearchQuery(e.target.value)}
                style={{
                  flex: '1 1 280px',
                  padding: '6px 12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  fontSize: '12.5px'
                }}
              />
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <input
                  type="file"
                  ref={directoryFileInputRef}
                  onChange={handleDirectoryFileUpload}
                  accept=".xlsx,.xls,.csv"
                  style={{ display: 'none' }}
                />
                <button
                  type="button"
                  onClick={() => directoryFileInputRef.current?.click()}
                  className="secondary"
                  style={{ fontSize: '12px', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: '5px' }}
                >
                  <Upload size={13} /> Import Directory (.xlsx / .csv)
                </button>
                <button
                  type="button"
                  onClick={handleResetDirectoryToDefault}
                  className="secondary"
                  style={{ fontSize: '12px', padding: '6px 12px', color: 'var(--muted)' }}
                >
                  Reset Defaults
                </button>
              </div>
            </div>

            {/* Directory Table */}
            <div style={{ border: '1px solid var(--line)', borderRadius: '8px', overflow: 'hidden', maxHeight: '45vh', overflowY: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px', textAlign: 'left' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)', color: 'var(--muted)', fontSize: '11px', textTransform: 'uppercase' }}>
                    <th style={{ padding: '8px 12px' }}>Code</th>
                    <th style={{ padding: '8px 12px' }}>College / Department Name</th>
                    <th style={{ padding: '8px 12px' }}>Official Email</th>
                    <th style={{ padding: '8px 12px', textAlign: 'right' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {emailDirectory
                    .filter(c => {
                      if (!directorySearchQuery) return true;
                      const q = directorySearchQuery.toLowerCase();
                      return String(c.code).toLowerCase().includes(q) ||
                             String(c.name).toLowerCase().includes(q) ||
                             String(c.email).toLowerCase().includes(q);
                    })
                    .map((item, idx) => (
                      <tr key={idx} style={{ borderBottom: '1px solid var(--line)', background: idx % 2 === 0 ? 'var(--panel)' : 'var(--bg)' }}>
                        <td style={{ padding: '8px 12px', fontFamily: 'monospace', fontWeight: 700, color: 'var(--accent)' }}>
                          {item.code}
                        </td>
                        <td style={{ padding: '8px 12px', fontWeight: 600, color: 'var(--ink)' }}>
                          {item.name || '—'}
                        </td>
                        <td style={{ padding: '8px 12px', fontFamily: 'monospace', color: '#059669' }}>
                          {item.email}
                        </td>
                        <td style={{ padding: '8px 12px', textAlign: 'right' }}>
                          <button
                            type="button"
                            onClick={() => handleEditCollege(idx)}
                            style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '2px 6px', color: 'var(--accent)' }}
                            title="Edit"
                          >
                            <Edit2 size={13} />
                          </button>
                          <button
                            type="button"
                            onClick={() => handleDeleteCollege(idx)}
                            style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '2px 6px', color: '#dc2626' }}
                            title="Remove"
                          >
                            <Trash2 size={13} />
                          </button>
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>

            <div style={{ marginTop: '16px', textAlign: 'right' }}>
              <button
                type="button"
                onClick={() => { setShowEmailDirectoryModal(false); setEditingCollegeIndex(null); }}
                style={{ padding: '8px 20px', fontSize: '13px' }}
              >
                Done
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 3: Global Extraction Logic & Guide Modal */}
      {showGuideModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0,0,0,0.6)',
          backdropFilter: 'blur(3px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          padding: '20px'
        }}>
          <div className="card" style={{
            maxWidth: '680px',
            width: '100%',
            maxHeight: '85vh',
            overflowY: 'auto',
            background: 'var(--panel)',
            padding: '24px',
            borderRadius: '12px',
            boxShadow: '0 20px 40px rgba(0,0,0,0.3)',
            border: '1px solid var(--line)'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', borderBottom: '1px solid var(--line)', paddingBottom: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <ShieldCheck size={22} color="#e11d48" />
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 800 }}>Confisection &amp; College Dispatch Guide</h3>
              </div>
              <button
                type="button"
                onClick={() => setShowGuideModal(false)}
                style={{ background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer', padding: 0 }}
              >
                <X size={18} />
              </button>
            </div>

            <div style={{ fontSize: '13px', lineHeight: '1.6', color: 'var(--ink)', display: 'flex', flexDirection: 'column', gap: '14px' }}>
              <div>
                <strong style={{ color: 'var(--accent)' }}>1. Universal Ingestion (13 Confi Headers or Exam Schedule):</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--muted)' }}>
                  Accepts standard university schedule files containing: <code>Event Name</code>, <code>Center Code</code>, <code>Center Name</code>, <code>Venue Code</code>, <code>Venue Name</code>, <code>Program Name</code>, <code>Exam Date</code>, <code>Course Start Time</code>, <code>Course End Time</code>, <code>Course Code</code>, <code>Course Name</code>, <code>Student Count</code>, and <code>venueId</code>, as well as Venue Date Course files with in-file <code>College Email</code>.
                </p>
              </div>

              <div>
                <strong style={{ color: 'var(--accent)' }}>2. Cross-Programme Candidate Aggregation:</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--muted)' }}>
                  Multiple academic programmes taking the same subject at the same venue on the same day are consolidated by <code>Date + Venue + Course Code</code>, summing student counts into a single Question Paper dispatch record.
                </p>
              </div>

              <div>
                <strong style={{ color: 'var(--accent)' }}>3. Official College Email Directory &amp; Circulars:</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--muted)' }}>
                  Automatically maps college venue codes to verified university email addresses. Provides one-click <strong>"Copy All College Emails"</strong> (semicolon-separated for Gmail/Outlook BCC), single-click <strong>"Mail"</strong> links, and formal notification letters for Principals and Chief Superintendents.
                </p>
              </div>

              <div>
                <strong style={{ color: 'var(--accent)' }}>4. Question Paper Bundle Calculations:</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--muted)' }}>
                  Calculates sealed confidential bundles: <strong>Packets of 20</strong>, <strong>Packets of 10</strong>, <strong>Packets of 5</strong>, and <strong>Loose singles</strong>, with selectable buffer reserve (+1 per venue, +2 per venue, or +5%).
                </p>
              </div>

              <div>
                <strong style={{ color: 'var(--accent)' }}>5. Comprehensive 5-Sheet Excel Export:</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--muted)' }}>
                  Clicking <strong>Export Report (.xlsx)</strong> generates:
                  (1) <code>Confi_Master_Schedule</code> (with College Emails), (2) <code>Venue_Email_Dispatch</code>, (3) <code>Daily_Press_Print_Order</code>, (4) <code>Venue_Packing_Slips</code>, and (5) <code>Course_Master_Summary</code>.
                </p>
              </div>
            </div>

            <div style={{ marginTop: '20px', paddingTop: '12px', borderTop: '1px solid var(--line)', textAlign: 'right' }}>
              <button
                type="button"
                onClick={() => setShowGuideModal(false)}
                style={{ padding: '8px 20px', fontSize: '13px' }}
              >
                Close Guide
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
