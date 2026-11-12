import React, { useState, useMemo, useRef } from 'react';
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
  ClipboardList
} from 'lucide-react';
import { readSpreadsheetAsAoa, cleanTimeText, normalizeKeyText, cleanDisplayField } from '../utils/excelParser';
import { formatExamDate } from '../utils/venueEmailHelpers';
import { 
  CONFIDENTIAL_SECTION_COLUMNS, 
  calculateQpPackets, 
  getExamSession,
  SAMPLE_CONFISECTION_ROWS
} from '../utils/confisectionHelpers';

export default function ConfisectionExamReportPage() {
  const navigate = useNavigate();

  // State
  const [rawRows, setRawRows] = useState([]);
  const [fileMeta, setFileMeta] = useState(null);
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
  const [packedChecklist, setPackedChecklist] = useState({});

  const fileInputRef = useRef(null);

  // Column auto-detection
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

  // Upload parser
  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setStatusAlert({ type: 'info', message: `Reading "${file.name}"...` });

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const aoa = readSpreadsheetAsAoa(data);

        if (!aoa || aoa.length < 2) {
          throw new Error('Spreadsheet appears to be empty or has no data rows.');
        }

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
          throw new Error('Could not identify Course Code or Exam Date columns. Expected 13 standard Confi Section headers.');
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
        setFileMeta({
          name: file.name,
          size: `${(file.size / 1024).toFixed(1)} KB`,
          rowsCount: parsedRecords.length,
          event: eventDetected || 'Examination Schedule'
        });
        setCurrentPage(0);
        setStatusAlert({
          type: 'success',
          message: `Loaded ${parsedRecords.length.toLocaleString()} schedule entries from "${file.name}". Ready for Confisection analysis.`
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
    setFileMeta({
      name: 'Center Venue Date Wise Report (Sample).xlsx',
      size: '24.5 KB',
      rowsCount: SAMPLE_CONFISECTION_ROWS.length,
      event: 'III Sem(R) November 2026'
    });
    setCurrentPage(0);
    setStatusAlert({
      type: 'success',
      message: `Loaded ${SAMPLE_CONFISECTION_ROWS.length} sample records from Kannur University III Sem (Regular) schedule.`
    });
  };

  const handleClearData = () => {
    setRawRows([]);
    setFileMeta(null);
    setCurrentPage(0);
    setStatusAlert(null);
  };

  // Consolidated Master Report (Unique Date + Venue + Course Code with summed student counts)
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

    list.sort((a, b) => {
      if (a.examDate !== b.examDate) return a.examDate.localeCompare(b.examDate);
      if (a.venueCode !== b.venueCode) return a.venueCode.localeCompare(b.venueCode);
      return a.courseCode.localeCompare(b.courseCode);
    });

    return list;
  }, [rawRows, bufferRule]);

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

  // Toggle packing checklist item
  const handleToggleChecklist = (itemKey) => {
    setPackedChecklist(prev => ({
      ...prev,
      [itemKey]: !prev[itemKey]
    }));
  };

  // Export to Excel with 4 sheets
  const handleExportExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Master Schedule
      const sheet1Headers = [
        'Event Name', 'Center Code', 'Center Name', 'Venue Code', 'Venue Name', 'venueId',
        'Exam Date', 'Session', 'Start Time', 'End Time', 'Course Code', 'Course Name',
        'Combined Programs', 'Student Count', 'Packets of 20', 'Packets of 10', 'Packets of 5',
        'Loose / Singles', 'Buffer Copies', 'Total QP Required'
      ];
      const sheet1Data = consolidatedRows.map(r => [
        r.eventName, r.centerCode, r.centerName, r.venueCode, r.venueName, r.venueId,
        r.examDate, r.session, r.startTime, r.endTime, r.courseCode, r.courseName,
        r.programsFormatted, r.studentCount, r.p20, r.p10, r.p5, r.singles, r.buffer, r.totalQp
      ]);
      const ws1 = XLSX.utils.aoa_to_sheet([sheet1Headers, ...sheet1Data]);
      ws1['!cols'] = sheet1Headers.map(() => ({ wch: 20 }));
      XLSX.utils.book_append_sheet(wb, ws1, 'Confi_Master_Schedule');

      // Sheet 2: Daily Press Print Order
      const sheet2Headers = [
        'Exam Date', 'Session', 'Start Time', 'End Time', 'Course Code', 'Course Name',
        'Total Venues', 'Total Candidates', 'Total QP Print Order', 'Venues List'
      ];
      const sheet2Data = pressPrintOrderRows.map(r => [
        r.examDate, r.session, r.startTime, r.endTime, r.courseCode, r.courseName,
        r.venuesCount, r.totalStudents, r.totalQp, r.venueCodes.join(', ')
      ]);
      const ws2 = XLSX.utils.aoa_to_sheet([sheet2Headers, ...sheet2Data]);
      ws2['!cols'] = sheet2Headers.map(() => ({ wch: 22 }));
      XLSX.utils.book_append_sheet(wb, ws2, 'Daily_Press_Print_Order');

      // Sheet 3: Venue Packing Slips
      const sheet3Headers = [
        'Venue Code', 'Venue Name', 'Center Code', 'Exam Date', 'Session',
        'Course Code', 'Course Name', 'Student Count', '20s', '10s', '5s', 'Singles', 'Total QP Packed'
      ];
      const sheet3Data = [];
      Object.values(venuePackingSlips).forEach(v => {
        v.items.forEach(item => {
          sheet3Data.push([
            v.venueCode, v.venueName, v.centerCode, item.examDate, item.session,
            item.courseCode, item.courseName, item.studentCount, item.p20, item.p10, item.p5, item.singles, item.totalQp
          ]);
        });
      });
      const ws3 = XLSX.utils.aoa_to_sheet([sheet3Headers, ...sheet3Data]);
      ws3['!cols'] = sheet3Headers.map(() => ({ wch: 18 }));
      XLSX.utils.book_append_sheet(wb, ws3, 'Venue_Packing_Slips');

      // Sheet 4: Course Master Summary
      const sheet4Headers = [
        'Course Code', 'Course Name', 'Scheduled Dates', 'Session',
        'Total Venues', 'Total Candidates', 'Total QP Needed'
      ];
      const sheet4Data = courseSummaryRows.map(r => [
        r.courseCode, r.courseName, r.datesFormatted, r.session,
        r.venueCount, r.totalStudents, r.totalQp
      ]);
      const ws4 = XLSX.utils.aoa_to_sheet([sheet4Headers, ...sheet4Data]);
      ws4['!cols'] = sheet4Headers.map(() => ({ wch: 22 }));
      XLSX.utils.book_append_sheet(wb, ws4, 'Course_Master_Summary');

      const outName = `Confisection_Exam_Report_${fileMeta?.event ? fileMeta.event.replace(/[^a-zA-Z0-9]/g, '_') : 'Schedule'}.xlsx`;
      XLSX.writeFile(wb, outName);
      setStatusAlert({ type: 'success', message: `Downloaded ${outName} with all 4 confidential reports!` });
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
    XLSX.utils.book_append_sheet(wb, ws, 'Confisection_Template');
    XLSX.writeFile(wb, 'Confisection_Exam_Source_Template.xlsx');
  };

  return (
    <div style={{ maxWidth: '1440px', margin: '0 auto', padding: '24px 32px 80px', color: 'var(--ink)' }}>
      {/* Top Navigation Bar */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '12px' }}>
        <button
          onClick={() => navigate('/')}
          className="secondary"
          style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '8px 14px', fontSize: '13px' }}
        >
          <ArrowLeft size={16} /> Back to Dashboard
        </button>

        <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
          <button
            type="button"
            className="secondary"
            onClick={handleLoadSample}
            style={{ 
              display: 'inline-flex', 
              alignItems: 'center', 
              gap: '6px', 
              fontSize: '13px', 
              background: 'var(--accent-soft)', 
              color: 'var(--accent)', 
              borderColor: 'var(--accent)' 
            }}
          >
            <Sparkles size={14} /> Load Kannur Sample Data
          </button>

          <button
            type="button"
            className="secondary"
            onClick={handleDownloadTemplate}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
          >
            <FileSpreadsheet size={14} /> 13-Col Template
          </button>

          <button
            type="button"
            className="secondary"
            onClick={() => window.print()}
            disabled={consolidatedRows.length === 0}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
          >
            <Printer size={14} /> Print Slips
          </button>

          <button
            type="button"
            onClick={handleExportExcel}
            disabled={consolidatedRows.length === 0}
            style={{ 
              display: 'inline-flex', 
              alignItems: 'center', 
              gap: '6px', 
              fontSize: '13px',
              background: '#059669',
              color: 'white'
            }}
          >
            <Download size={14} /> Export Report (.xlsx)
          </button>

          <button
            type="button"
            className="secondary"
            onClick={() => setShowGuideModal(true)}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
          >
            <HelpCircle size={14} /> Logic &amp; Guide
          </button>
        </div>
      </div>

      {/* Page Header */}
      <div style={{ marginBottom: '24px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '8px' }}>
          <div style={{ 
            background: 'rgba(225, 29, 72, 0.12)', 
            padding: '10px', 
            borderRadius: '10px', 
            color: '#e11d48',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
          }}>
            <ShieldCheck size={28} />
          </div>
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <h1 style={{ fontSize: '26px', fontWeight: 800, margin: 0, color: 'var(--ink)' }}>
                Confisection Exam Report
              </h1>
              <span style={{ 
                fontSize: '11px', 
                fontWeight: 700, 
                textTransform: 'uppercase', 
                letterSpacing: '0.5px',
                padding: '3px 8px', 
                borderRadius: '6px', 
                background: 'rgba(225, 29, 72, 0.1)', 
                color: '#e11d48',
                border: '1px solid rgba(225, 29, 72, 0.25)'
              }}>
                Confidential Section
              </span>
            </div>
            <p style={{ margin: '4px 0 0 0', fontSize: '13.5px', color: 'var(--muted)' }}>
              Ingest 13-column examination schedules to aggregate cross-programme enrollment, calculate confidential QP printing press quotas, and prepare sealed venue dispatch checklists.
            </p>
          </div>
        </div>
      </div>

      {/* File Upload Zone */}
      <div className="card" style={{ padding: '22px', border: '1.5px dashed var(--line)', background: 'var(--panel)', marginBottom: '22px', borderRadius: '12px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px', flexWrap: 'wrap', gap: '10px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <div style={{ background: 'var(--accent-soft)', padding: '8px', borderRadius: '8px', color: 'var(--accent)' }}>
              <FileSpreadsheet size={20} />
            </div>
            <div>
              <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700 }}>
                Upload Center Venue Date Wise Schedule
              </h3>
              <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
                Expects standard 13 headers: Event Name, Center Code/Name, Venue Code/Name, Program Name, Exam Date, Times, Course Code/Name, Student Count, venueId
              </span>
            </div>
          </div>

          {fileMeta && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ 
                padding: '4px 10px', 
                borderRadius: '6px', 
                fontSize: '12px', 
                fontWeight: 600, 
                background: 'rgba(5, 150, 105, 0.1)', 
                color: '#059669',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px'
              }}>
                <CheckCircle2 size={13} /> {fileMeta.name} ({fileMeta.rowsCount.toLocaleString()} rows)
              </span>
              <button
                type="button"
                onClick={handleClearData}
                style={{ background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer', padding: '4px' }}
                title="Clear current file"
              >
                <X size={15} />
              </button>
            </div>
          )}
        </div>

        <input
          type="file"
          ref={fileInputRef}
          onChange={handleFileUpload}
          accept=".xlsx,.xls,.csv"
          style={{ display: 'none' }}
        />

        <div
          onClick={() => fileInputRef.current?.click()}
          style={{
            border: '2px dashed var(--accent)',
            borderRadius: '10px',
            padding: '28px 20px',
            textAlign: 'center',
            cursor: 'pointer',
            background: 'var(--bg)',
            transition: 'all 0.2s ease'
          }}
        >
          <Upload size={28} color="var(--accent)" style={{ margin: '0 auto 8px' }} />
          <p style={{ margin: '0 0 4px', fontSize: '14px', fontWeight: 700, color: 'var(--ink)' }}>
            {fileMeta ? 'Click to replace or upload another schedule file' : 'Click to Upload Examination Schedule (.xlsx, .xls, .csv)'}
          </p>
          <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
            Drop in your "Center Venue Date Wise Report" export directly from the university examination portal
          </span>
        </div>
      </div>

      {/* Status Feedback Banner */}
      {statusAlert && (
        <div style={{
          padding: '12px 18px',
          borderRadius: '8px',
          fontSize: '13px',
          marginBottom: '22px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: '10px',
          background: statusAlert.type === 'error' ? 'rgba(239, 68, 68, 0.08)' : statusAlert.type === 'success' ? 'rgba(16, 185, 129, 0.08)' : 'var(--accent-soft)',
          color: statusAlert.type === 'error' ? '#ef4444' : statusAlert.type === 'success' ? '#059669' : 'var(--accent)',
          border: `1px solid ${statusAlert.type === 'error' ? '#ef444433' : statusAlert.type === 'success' ? '#10b98133' : 'var(--accent)'}`
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            {statusAlert.type === 'error' ? <AlertCircle size={16} /> : <CheckCircle2 size={16} />}
            <span>{statusAlert.message}</span>
          </div>
          <button
            onClick={() => setStatusAlert(null)}
            style={{ background: 'transparent', border: 'none', color: 'inherit', cursor: 'pointer', padding: 0 }}
          >
            <X size={15} />
          </button>
        </div>
      )}

      {/* KPI Metric Cards */}
      <div style={{ 
        display: 'grid', 
        gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', 
        gap: '16px', 
        marginBottom: '24px' 
      }}>
        {/* Total Candidates */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid #e11d48' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Total Candidates
            </span>
            <Users size={18} color="#e11d48" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>
            {summaryStats.totalCandidates.toLocaleString()}
          </div>
          <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Sum of registered students</span>
        </div>

        {/* Total QP Print Order */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid #059669' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Total QP Required
            </span>
            <PackageCheck size={18} color="#059669" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>
            {summaryStats.totalQp.toLocaleString()}
          </div>
          <span style={{ fontSize: '11px', color: '#059669', fontWeight: 600 }}>
            Includes buffer reserve ({bufferRule === 'none' ? 'Exact' : bufferRule})
          </span>
        </div>

        {/* Unique Exams */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid var(--accent)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Unique Exams
            </span>
            <Layers size={18} color="var(--accent)" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>
            {summaryStats.uniqueExams.toLocaleString()}
          </div>
          <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Distinct Date-Venue-Course slots</span>
        </div>

        {/* Venues */}
        <div className="card" style={{ padding: '16px 20px', margin: 0, borderLeft: '4px solid #6366f1' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '4px' }}>
            <span style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
              Colleges / Venues
            </span>
            <Building2 size={18} color="#6366f1" />
          </div>
          <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>
            {summaryStats.uniqueVenues}
          </div>
          <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Examination dispatch points</span>
        </div>

        {/* Courses */}
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
          <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Unique course subjects</span>
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

          <button
            type="button"
            onClick={() => { setActiveTab('press_print'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'press_print' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'press_print' ? 'var(--accent)' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'press_print' ? '3px solid var(--accent)' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <Printer size={16} />
            Daily Press Print Order
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'press_print' ? 'var(--accent-soft)' : 'var(--line)',
              color: activeTab === 'press_print' ? 'var(--accent)' : 'var(--muted)'
            }}>
              {pressPrintOrderRows.length}
            </span>
          </button>

          <button
            type="button"
            onClick={() => { setActiveTab('venue_packing'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'venue_packing' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'venue_packing' ? 'var(--accent)' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'venue_packing' ? '3px solid var(--accent)' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <Building2 size={16} />
            Venue Packing Slips
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'venue_packing' ? 'var(--accent-soft)' : 'var(--line)',
              color: activeTab === 'venue_packing' ? 'var(--accent)' : 'var(--muted)'
            }}>
              {Object.keys(venuePackingSlips).length} Venues
            </span>
          </button>

          <button
            type="button"
            onClick={() => { setActiveTab('course_summary'); setCurrentPage(0); }}
            style={{
              padding: '10px 16px',
              fontSize: '13px',
              fontWeight: 700,
              background: activeTab === 'course_summary' ? 'var(--panel)' : 'transparent',
              color: activeTab === 'course_summary' ? 'var(--accent)' : 'var(--muted)',
              border: 'none',
              borderBottom: activeTab === 'course_summary' ? '3px solid var(--accent)' : '3px solid transparent',
              borderRadius: '6px 6px 0 0',
              cursor: 'pointer',
              display: 'inline-flex',
              alignItems: 'center',
              gap: '8px',
              whiteSpace: 'nowrap'
            }}
          >
            <BookOpen size={16} />
            Course Master Summary
            <span style={{ 
              padding: '2px 7px', 
              borderRadius: '10px', 
              fontSize: '11px', 
              fontWeight: 700,
              background: activeTab === 'course_summary' ? 'var(--accent-soft)' : 'var(--line)',
              color: activeTab === 'course_summary' ? 'var(--accent)' : 'var(--muted)'
            }}>
              {courseSummaryRows.length}
            </span>
          </button>

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
            <FileText size={16} />
            Raw Rows
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

        {/* Filter Controls Toolbar */}
        <div style={{ 
          padding: '14px 20px', 
          background: 'var(--panel)', 
          borderBottom: '1px solid var(--line)', 
          display: 'flex', 
          alignItems: 'center', 
          justifyContent: 'space-between', 
          flexWrap: 'wrap', 
          gap: '12px' 
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap', flex: 1 }}>
            {/* Live Search */}
            <div style={{ position: 'relative', width: '280px' }}>
              <Search size={15} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(0); }}
                placeholder="Search code, course, venue, date..."
                style={{
                  width: '100%',
                  padding: '7px 28px 7px 32px',
                  fontSize: '13px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)'
                }}
              />
              {searchQuery && (
                <button
                  type="button"
                  onClick={() => setSearchQuery('')}
                  style={{
                    position: 'absolute',
                    right: '8px',
                    top: '50%',
                    transform: 'translateY(-50%)',
                    background: 'transparent',
                    border: 'none',
                    color: 'var(--muted)',
                    cursor: 'pointer',
                    padding: '2px'
                  }}
                >
                  <X size={14} />
                </button>
              )}
            </div>

            {/* Date Filter */}
            <select
              value={selectedDate}
              onChange={(e) => { setSelectedDate(e.target.value); setCurrentPage(0); }}
              style={{ fontSize: '13px', padding: '7px 10px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)', color: 'var(--ink)' }}
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
              style={{ fontSize: '13px', padding: '7px 10px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)', color: 'var(--ink)' }}
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
                style={{ fontSize: '13px', padding: '7px 10px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)', color: 'var(--ink)', maxWidth: '240px' }}
              >
                <option value="ALL">All Venues ({filterOptions.venues.length})</option>
                {filterOptions.venues.map(v => (
                  <option key={v.code} value={v.code}>{v.code} - {v.name}</option>
                ))}
              </select>
            )}
          </div>

          {/* Buffer Policy Selector & Page Size */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <div style={{ 
              display: 'inline-flex', 
              alignItems: 'center', 
              gap: '6px', 
              padding: '5px 10px', 
              borderRadius: '6px', 
              background: 'rgba(225, 29, 72, 0.08)', 
              border: '1px solid rgba(225, 29, 72, 0.25)',
              fontSize: '12px'
            }}>
              <span style={{ fontWeight: 700, color: '#e11d48' }}>Buffer Policy:</span>
              <select
                value={bufferRule}
                onChange={(e) => setBufferRule(e.target.value)}
                style={{ 
                  background: 'transparent', 
                  border: 'none', 
                  fontSize: '12px', 
                  fontWeight: 600, 
                  color: 'var(--ink)', 
                  cursor: 'pointer',
                  padding: 0
                }}
              >
                <option value="plus1">+1 Copy / Venue (Standard Confi)</option>
                <option value="plus2">+2 Copies / Venue (High Reserve)</option>
                <option value="five_percent">+5% Reserve</option>
                <option value="none">Exact Count (0 Buffer)</option>
              </select>
            </div>

            {activeTab !== 'venue_packing' && (
              <select
                value={pageSize}
                onChange={(e) => { setPageSize(Number(e.target.value)); setCurrentPage(0); }}
                style={{ fontSize: '13px', padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)', color: 'var(--ink)' }}
              >
                <option value={25}>25 rows</option>
                <option value={50}>50 rows</option>
                <option value={100}>100 rows</option>
                <option value={0}>Show All</option>
              </select>
            )}
          </div>
        </div>

        {/* Content Body */}
        {activeTab === 'venue_packing' ? (
          /* Card-based Venue Packing Slips */
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
                            <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
                              Center Code: {venueGroup.centerCode} ({venueGroup.centerName || 'Main'}) • Venue ID: {venueGroup.venueId || '—'}
                            </span>
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
                                  <td style={{ padding: '10px 14px', fontWeight: 600, whiteSpace: 'nowrap' }}>
                                    {item.examDate}
                                  </td>
                                  <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                                    <span style={{ 
                                      padding: '3px 8px', 
                                      borderRadius: '4px', 
                                      fontSize: '11px', 
                                      fontWeight: 700,
                                      background: item.session === 'FN' ? 'rgba(2, 132, 199, 0.1)' : 'rgba(99, 102, 241, 0.1)',
                                      color: item.session === 'FN' ? '#0284c7' : '#6366f1'
                                    }}>
                                      {item.session} ({item.startTime})
                                    </span>
                                  </td>
                                  <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontWeight: 800, color: 'var(--accent)' }}>
                                    {item.courseCode}
                                  </td>
                                  <td style={{ padding: '10px 14px', fontWeight: 500, color: 'var(--ink)' }}>
                                    {item.courseName}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 800 }}>
                                    {item.studentCount}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center', fontFamily: 'monospace' }}>
                                    {item.p20 > 0 ? <strong style={{ color: '#059669' }}>{item.p20}</strong> : <span style={{ color: 'var(--muted)' }}>—</span>}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center', fontFamily: 'monospace' }}>
                                    {item.p10 > 0 ? <strong style={{ color: '#0d9488' }}>{item.p10}</strong> : <span style={{ color: 'var(--muted)' }}>—</span>}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center', fontFamily: 'monospace' }}>
                                    {item.p5 > 0 ? <strong style={{ color: '#0284c7' }}>{item.p5}</strong> : <span style={{ color: 'var(--muted)' }}>—</span>}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'center', fontFamily: 'monospace' }}>
                                    {item.singles > 0 ? <strong style={{ color: '#d97706' }}>{item.singles}</strong> : <span style={{ color: 'var(--muted)' }}>—</span>}
                                  </td>
                                  <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 800, fontFamily: 'monospace', color: '#059669', fontSize: '13px' }}>
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
        ) : (
          /* Standard High-Performance Data Tables */
          <div style={{ overflowX: 'auto', maxHeight: '68vh' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px', textAlign: 'left' }}>
              <thead>
                <tr style={{ 
                  background: 'var(--bg)', 
                  borderBottom: '1.5px solid var(--line)', 
                  color: 'var(--muted)', 
                  fontSize: '11px', 
                  textTransform: 'uppercase',
                  position: 'sticky',
                  top: 0,
                  zIndex: 10
                }}>
                  <th style={{ padding: '12px 14px', width: '45px', textAlign: 'center' }}>#</th>

                  {activeTab === 'consolidated' && (
                    <>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Exam Date</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Session</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Venue</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Course Code</th>
                      <th style={{ padding: '12px 14px', minWidth: '220px' }}>Course Name</th>
                      <th style={{ padding: '12px 14px', minWidth: '240px' }}>Programme(s) Included</th>
                      <th style={{ padding: '12px 14px', textAlign: 'center', whiteSpace: 'nowrap' }}>Takers</th>
                      <th style={{ padding: '12px 14px', textAlign: 'center', whiteSpace: 'nowrap' }}>Sealed Packets (20/10/5/L)</th>
                      <th style={{ padding: '12px 14px', textAlign: 'right', whiteSpace: 'nowrap' }}>Total QP</th>
                    </>
                  )}

                  {activeTab === 'press_print' && (
                    <>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Exam Date</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Session</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Scheduled Timing</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Course Code</th>
                      <th style={{ padding: '12px 14px', minWidth: '260px' }}>Course Name</th>
                      <th style={{ padding: '12px 14px', textAlign: 'center', whiteSpace: 'nowrap' }}>Venues</th>
                      <th style={{ padding: '12px 14px', textAlign: 'center', whiteSpace: 'nowrap' }}>Total Candidates</th>
                      <th style={{ padding: '12px 14px', textAlign: 'right', whiteSpace: 'nowrap' }}>QP Print Order</th>
                    </>
                  )}

                  {activeTab === 'course_summary' && (
                    <>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Course Code</th>
                      <th style={{ padding: '12px 14px', minWidth: '280px' }}>Course Name</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Scheduled Dates</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Session</th>
                      <th style={{ padding: '12px 14px', textAlign: 'center', whiteSpace: 'nowrap' }}>Venue Coverage</th>
                      <th style={{ padding: '12px 14px', textAlign: 'center', whiteSpace: 'nowrap' }}>Total Students</th>
                      <th style={{ padding: '12px 14px', textAlign: 'right', whiteSpace: 'nowrap' }}>Total QP</th>
                    </>
                  )}

                  {activeTab === 'raw' && (
                    <>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Event Name</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Venue Code &amp; Name</th>
                      <th style={{ padding: '12px 14px', minWidth: '220px' }}>Programme</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Exam Date</th>
                      <th style={{ padding: '12px 14px', whiteSpace: 'nowrap' }}>Course Code</th>
                      <th style={{ padding: '12px 14px', minWidth: '220px' }}>Course Name</th>
                      <th style={{ padding: '12px 14px', textAlign: 'right', whiteSpace: 'nowrap' }}>Count</th>
                    </>
                  )}
                </tr>
              </thead>
              <tbody>
                {paginatedRows.length === 0 ? (
                  <tr>
                    <td colSpan={10} style={{ padding: '60px 20px', textAlign: 'center', color: 'var(--muted)' }}>
                      <PackageCheck size={36} style={{ margin: '0 auto 10px', opacity: 0.4 }} />
                      <p style={{ margin: 0, fontWeight: 600, fontSize: '14px' }}>
                        {searchQuery ? `No matching records found for "${searchQuery}"` : 'No examination schedule loaded.'}
                      </p>
                      <span style={{ fontSize: '12px' }}>
                        Upload your Schedule (.xlsx) or click "Load Kannur Sample Data" above to preview.
                      </span>
                    </td>
                  </tr>
                ) : (
                  paginatedRows.map((row, idx) => {
                    const rowNumber = (pageSize === 0 ? 0 : currentPage * pageSize) + idx + 1;

                    return (
                      <tr
                        key={idx}
                        style={{
                          borderBottom: '1px solid var(--line)',
                          background: idx % 2 === 0 ? 'var(--panel)' : 'var(--bg)',
                          transition: 'background 0.15s ease'
                        }}
                      >
                        <td style={{ padding: '10px 14px', textAlign: 'center', fontFamily: 'monospace', color: 'var(--muted)' }}>
                          {rowNumber}
                        </td>

                        {activeTab === 'consolidated' && (
                          <>
                            <td style={{ padding: '10px 14px', fontWeight: 700, whiteSpace: 'nowrap' }}>
                              {row.examDate}
                            </td>
                            <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                              <span style={{ 
                                padding: '3px 8px', 
                                borderRadius: '4px', 
                                fontSize: '11px', 
                                fontWeight: 700,
                                background: row.session === 'FN' ? 'rgba(2, 132, 199, 0.1)' : 'rgba(99, 102, 241, 0.1)',
                                color: row.session === 'FN' ? '#0284c7' : '#6366f1'
                              }}>
                                {row.session} ({row.startTime})
                              </span>
                            </td>
                            <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                              <div style={{ display: 'flex', flexDirection: 'column' }}>
                                <span style={{ fontWeight: 800, color: 'var(--ink)' }}>{row.venueCode}</span>
                                <span style={{ fontSize: '11px', color: 'var(--muted)', maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis' }} title={row.venueName}>
                                  {row.venueName}
                                </span>
                              </div>
                            </td>
                            <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontWeight: 800, color: 'var(--accent)', whiteSpace: 'nowrap' }}>
                              {row.courseCode}
                            </td>
                            <td style={{ padding: '10px 14px', fontWeight: 600, color: 'var(--ink)' }}>
                              {row.courseName}
                            </td>
                            <td style={{ padding: '10px 14px', fontSize: '11.5px', color: 'var(--muted)', maxWidth: '280px' }} title={row.programsFormatted}>
                              <span style={{ 
                                display: '-webkit-box', 
                                WebkitLineClamp: 2, 
                                WebkitBoxOrient: 'vertical', 
                                overflow: 'hidden' 
                              }}>
                                {row.programsFormatted || 'General Examination'}
                              </span>
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 800, fontSize: '13px' }}>
                              {row.studentCount}
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                              <div style={{ display: 'inline-flex', gap: '4px', fontFamily: 'monospace', fontSize: '11px' }}>
                                <span style={{ padding: '2px 5px', borderRadius: '4px', background: 'rgba(5, 150, 105, 0.1)', color: '#059669', fontWeight: 700 }} title="Packets of 20">
                                  20s: {row.p20}
                                </span>
                                <span style={{ padding: '2px 5px', borderRadius: '4px', background: 'rgba(13, 148, 136, 0.1)', color: '#0d9488', fontWeight: 700 }} title="Packets of 10">
                                  10s: {row.p10}
                                </span>
                                <span style={{ padding: '2px 5px', borderRadius: '4px', background: 'rgba(2, 132, 199, 0.1)', color: '#0284c7', fontWeight: 700 }} title="Packets of 5">
                                  5s: {row.p5}
                                </span>
                                {row.singles > 0 && (
                                  <span style={{ padding: '2px 5px', borderRadius: '4px', background: 'rgba(217, 119, 6, 0.1)', color: '#d97706', fontWeight: 700 }} title="Loose single copies">
                                    L: {row.singles}
                                  </span>
                                )}
                              </div>
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 800, fontFamily: 'monospace', color: '#059669', fontSize: '13px' }}>
                              {row.totalQp}
                            </td>
                          </>
                        )}

                        {activeTab === 'press_print' && (
                          <>
                            <td style={{ padding: '10px 14px', fontWeight: 700, whiteSpace: 'nowrap' }}>
                              {row.examDate}
                            </td>
                            <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                              <span style={{ 
                                padding: '3px 8px', 
                                borderRadius: '4px', 
                                fontSize: '11px', 
                                fontWeight: 700,
                                background: row.session === 'FN' ? 'rgba(2, 132, 199, 0.1)' : 'rgba(99, 102, 241, 0.1)',
                                color: row.session === 'FN' ? '#0284c7' : '#6366f1'
                              }}>
                                {row.session}
                              </span>
                            </td>
                            <td style={{ padding: '10px 14px', color: 'var(--muted)', whiteSpace: 'nowrap' }}>
                              {row.startTime} – {row.endTime}
                            </td>
                            <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontWeight: 800, color: 'var(--accent)', whiteSpace: 'nowrap' }}>
                              {row.courseCode}
                            </td>
                            <td style={{ padding: '10px 14px', fontWeight: 600, color: 'var(--ink)' }}>
                              {row.courseName}
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 700 }}>
                              {row.venuesCount}
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 800 }}>
                              {row.totalStudents.toLocaleString()}
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 800, fontFamily: 'monospace', color: '#059669', fontSize: '13.5px' }}>
                              {row.totalQp.toLocaleString()}
                            </td>
                          </>
                        )}

                        {activeTab === 'course_summary' && (
                          <>
                            <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontWeight: 800, color: 'var(--accent)', whiteSpace: 'nowrap' }}>
                              {row.courseCode}
                            </td>
                            <td style={{ padding: '10px 14px', fontWeight: 600, color: 'var(--ink)' }}>
                              {row.courseName}
                            </td>
                            <td style={{ padding: '10px 14px', fontWeight: 600, color: 'var(--ink)', whiteSpace: 'nowrap' }}>
                              {row.datesFormatted}
                            </td>
                            <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                              <span style={{ padding: '3px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 700, background: 'var(--bg)', color: 'var(--ink)' }}>
                                {row.session}
                              </span>
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 700 }}>
                              {row.venueCount}
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 800 }}>
                              {row.totalStudents.toLocaleString()}
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 800, fontFamily: 'monospace', color: '#059669' }}>
                              {row.totalQp.toLocaleString()}
                            </td>
                          </>
                        )}

                        {activeTab === 'raw' && (
                          <>
                            <td style={{ padding: '10px 14px', whiteSpace: 'nowrap', color: 'var(--muted)' }}>
                              {row.eventName}
                            </td>
                            <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                              <strong>{row.venueCode}</strong> – {row.venueName}
                            </td>
                            <td style={{ padding: '10px 14px', fontSize: '11.5px', color: 'var(--muted)', maxWidth: '260px', overflow: 'hidden', textOverflow: 'ellipsis' }} title={row.programName}>
                              {row.programName}
                            </td>
                            <td style={{ padding: '10px 14px', whiteSpace: 'nowrap', fontWeight: 600 }}>
                              {row.examDate}
                            </td>
                            <td style={{ padding: '10px 14px', fontFamily: 'monospace', fontWeight: 800, color: 'var(--accent)', whiteSpace: 'nowrap' }}>
                              {row.courseCode}
                            </td>
                            <td style={{ padding: '10px 14px', fontWeight: 500 }}>
                              {row.courseName}
                            </td>
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 800 }}>
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
          <div style={{ 
            padding: '12px 20px', 
            background: 'var(--bg)', 
            borderTop: '1px solid var(--line)', 
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'space-between', 
            fontSize: '12px', 
            color: 'var(--muted)' 
          }}>
            <span>
              Showing <strong>{paginatedRows.length}</strong> of <strong>{currentFilteredRows.length}</strong> records (Page {currentPage + 1} of {totalPages})
            </span>
            <div style={{ display: 'flex', gap: '6px' }}>
              <button
                type="button"
                className="secondary"
                onClick={() => setCurrentPage(p => Math.max(0, p - 1))}
                disabled={currentPage === 0}
                style={{ padding: '4px 10px', fontSize: '12px' }}
              >
                <ChevronLeft size={14} /> Prev
              </button>
              <button
                type="button"
                className="secondary"
                onClick={() => setCurrentPage(p => Math.min(totalPages - 1, p + 1))}
                disabled={currentPage >= totalPages - 1}
                style={{ padding: '4px 10px', fontSize: '12px' }}
              >
                Next <ChevronRight size={14} />
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Guide Modal */}
      {showGuideModal && (
        <div style={{
          position: 'fixed',
          inset: 0,
          background: 'rgba(0, 0, 0, 0.55)',
          backdropFilter: 'blur(3px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999,
          padding: '20px'
        }}>
          <div className="card" style={{
            maxWidth: '680px',
            width: '100%',
            maxHeight: '85vh',
            overflowY: 'auto',
            background: 'var(--panel)',
            padding: '24px',
            margin: 0,
            borderRadius: '12px',
            boxShadow: 'var(--shadow)'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', borderBottom: '1px solid var(--line)', paddingBottom: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <ShieldCheck size={22} color="#e11d48" />
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 800 }}>Confisection Exam Report Guide</h3>
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
                <strong style={{ color: 'var(--accent)' }}>1. 13 Source File Headers:</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--muted)' }}>
                  Accepts standard university schedule files containing: <code>Event Name</code>, <code>Center Code</code>, <code>Center Name</code>, <code>Venue Code</code>, <code>Venue Name</code>, <code>Program Name</code>, <code>Exam Date</code>, <code>Course Start Time</code>, <code>Course End Time</code>, <code>Course Code</code>, <code>Course Name</code>, <code>Student Count</code>, and <code>venueId</code>.
                </p>
              </div>

              <div>
                <strong style={{ color: 'var(--accent)' }}>2. Cross-Programme Candidate Aggregation:</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--muted)' }}>
                  Multiple academic programmes taking the same subject at the same venue on the same day are consolidated by <code>Date + Venue + Course Code</code>, summing student counts into a single Question Paper dispatch record.
                </p>
              </div>

              <div>
                <strong style={{ color: 'var(--accent)' }}>3. Question Paper Bundle Calculations:</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--muted)' }}>
                  Calculates sealed confidential bundles: <strong>Packets of 20</strong>, <strong>Packets of 10</strong>, <strong>Packets of 5</strong>, and <strong>Loose singles</strong>, with selectable buffer reserve (+1 per venue, +2 per venue, or +5%).
                </p>
              </div>

              <div>
                <strong style={{ color: 'var(--accent)' }}>4. Multi-Sheet Excel Export:</strong>
                <p style={{ margin: '4px 0 0', color: 'var(--muted)' }}>
                  Clicking <strong>Export Report (.xlsx)</strong> generates 4 distinct sheets:
                  (1) <code>Confi_Master_Schedule</code>, (2) <code>Daily_Press_Print_Order</code>, (3) <code>Venue_Packing_Slips</code>, and (4) <code>Course_Master_Summary</code>.
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
