import React, { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
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
  ChevronDown
} from 'lucide-react';
import { readSpreadsheetAsAoa, cleanTimeText, normalizeKeyText, cleanDisplayField } from '../utils/excelParser';

const CourseWiseVenueDatePage = () => {
  const navigate = useNavigate();

  // File States
  const [preExamFile, setPreExamFile] = useState(null);
  const [timetableFile, setTimetableFile] = useState(null);

  const [preExamHeaders, setPreExamHeaders] = useState([]);
  const [preExamRows, setPreExamRows] = useState([]);

  const [timetableHeaders, setTimetableHeaders] = useState([]);
  const [timetableRows, setTimetableRows] = useState([]);

  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Upload both Pre-Exam Student Data and Timetable files to generate the report.');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Column Mappings for Source 1 (Pre-Exam)
  const [preExamMapping, setPreExamMapping] = useState({
    programmeName: -1,
    parentGroupName: -1,
    immediateGroupName: -1,
    courseCode: -1,
    courseName: -1,
    venueCode: -1,
    venueName: -1,
    studentCount: -1
  });

  // Column Mappings for Source 2 (Timetable)
  const [timetableMapping, setTimetableMapping] = useState({
    termAbbr: -1,
    parentGroupName: -1,
    immediateGroupName: -1,
    courseCode: -1,
    courseName: -1,
    examDate: -1,
    startHr: -1,
    startMin: -1,
    startAmPm: -1,
    endHr: -1,
    endMin: -1,
    endAmPm: -1,
    courseStartTime: -1,
    courseEndTime: -1
  });

  const [showConfig, setShowConfig] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedVenue, setSelectedVenue] = useState('ALL');
  const [selectedDate, setSelectedDate] = useState('ALL');
  const [selectedGroup, setSelectedGroup] = useState('ALL');
  const [sortBy, setSortBy] = useState('date'); // 'date' | 'course' | 'venue' | 'count'

  // Auto-detect columns in Pre-Exam file
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

      // Programme Name priorities:
      // 1. Strict exact match "Programme Name" / "programmename"
      // 2. "Program Name" / "programname"
      // 3. "Programme"
      // 4. "ProgramFullName"
      if (rawTrimmed === 'programme name' || lower === 'programmename') {
        pgmExact = idx;
      } else if (rawTrimmed === 'program name' || lower === 'programname') {
        pgmName = idx;
      } else if (rawTrimmed === 'programme' || lower === 'programme') {
        pgmOnly = idx;
      } else if (rawTrimmed === 'programfullname' || lower === 'programfullname') {
        pgmFull = idx;
      }

      // Parent Group Name
      if (lower === 'parentgroupname') {
        parentGrpExact = idx;
      } else if (lower === 'parentgroup') {
        parentGrpFallback = idx;
      }

      // Immediate Parent Group
      if (lower === 'immediateparentgroup' || lower === 'immediategroupname' || lower === 'immediateparent' || lower === 'immediategroup') {
        immGrpExact = idx;
      }

      // Course / Paper Code
      if (lower === 'papercode' || lower === 'coursecode') {
        courseCodeExact = idx;
      } else if (lower.includes('subjectcode')) {
        courseCodeFallback = idx;
      }

      // Course / Paper Name
      if (lower === 'papername' || lower === 'coursename') {
        courseNameExact = idx;
      } else if (lower.includes('subjectname') || lower.includes('coursetitle')) {
        courseNameFallback = idx;
      }

      // Venue Code: strictly prefer VenueCode over CenterCode
      if (lower === 'venuecode') {
        venueCodeExact = idx;
      } else if (lower === 'centercode' || lower === 'centrecode') {
        centerCodeFallback = idx;
      }

      // Venue Name: strictly prefer VenueName over CenterName
      if (lower === 'venuename') {
        venueNameExact = idx;
      } else if (lower === 'centername' || lower === 'centrename') {
        centerNameFallback = idx;
      }

      // Student count
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

  // Auto-detect columns in Timetable file
  const detectTimetableColumns = (headersList) => {
    const mapping = {
      termAbbr: -1,
      parentGroupName: -1,
      immediateGroupName: -1,
      courseCode: -1,
      courseName: -1,
      examDate: -1,
      startHr: -1,
      startMin: -1,
      startAmPm: -1,
      endHr: -1,
      endMin: -1,
      endAmPm: -1,
      courseStartTime: -1,
      courseEndTime: -1
    };

    headersList.forEach((h, idx) => {
      const lower = String(h).toLowerCase().replace(/[^a-z0-9]/g, '');

      if (lower === 'termabbr' || lower.includes('programme') || lower === 'term') {
        mapping.termAbbr = idx;
      }
      if (lower === 'parentgroupname' || lower === 'parentgroup') {
        mapping.parentGroupName = idx;
      }
      if (lower === 'immediategroupname' || lower === 'immediateparentgroup' || lower === 'immediategroup') {
        mapping.immediateGroupName = idx;
      }
      if (lower === 'coursecode' || lower === 'papercode' || lower.includes('subjectcode')) {
        mapping.courseCode = idx;
      }
      if (lower === 'coursename' || lower === 'papername' || lower.includes('subjectname')) {
        mapping.courseName = idx;
      }
      if (lower === 'examdate' || lower === 'date') {
        mapping.examDate = idx;
      }
      if (lower === 'starttimehr' || lower === 'starthr') {
        mapping.startHr = idx;
      }
      if (lower === 'starttimemin' || lower === 'startmin') {
        mapping.startMin = idx;
      }
      if (lower === 'startampm' || lower === 'startmeridiem') {
        mapping.startAmPm = idx;
      }
      if (lower === 'endtimehr' || lower === 'endhr') {
        mapping.endHr = idx;
      }
      if (lower === 'endtimemin' || lower === 'endmin') {
        mapping.endMin = idx;
      }
      if (lower === 'endampm' || lower === 'endmeridiem') {
        mapping.endAmPm = idx;
      }
      if (lower === 'coursestarttime' || lower === 'starttime' || lower === 'timefrom') {
        mapping.courseStartTime = idx;
      }
      if (lower === 'courseendtime' || lower === 'endtime' || lower === 'timeto') {
        mapping.courseEndTime = idx;
      }
    });

    return mapping;
  };

  // Handle uploading Source 1 (Pre-Exam)
  const handlePreExamUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setPreExamFile(file);
    setIsProcessing(true);
    setStatusMsg(`Loading Pre-Exam File: ${file.name}...`);
    setStatusType('normal');

    try {
      const aoa = await readSpreadsheetAsAoa(file);
      if (aoa && aoa.length > 0) {
        const detectedHeaders = aoa[0].map((h, i) => h ? String(h).trim() : `Column ${i + 1}`);
        const rows = aoa.slice(1);
        setPreExamHeaders(detectedHeaders);
        setPreExamRows(rows);

        const mapping = detectPreExamColumns(detectedHeaders);
        setPreExamMapping(mapping);

        setStatusMsg(`Loaded ${rows.length} records from ${file.name}.`);
        setStatusType('success');
      } else {
        setStatusMsg('No data found in uploaded Pre-Exam file.');
        setStatusType('error');
      }
    } catch (err) {
      console.error(err);
      setStatusMsg(`Error reading Pre-Exam file: ${err.message}`);
      setStatusType('error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Handle uploading Source 2 (Timetable)
  const handleTimetableUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setTimetableFile(file);
    setIsProcessing(true);
    setStatusMsg(`Loading Timetable File: ${file.name}...`);
    setStatusType('normal');

    try {
      const aoa = await readSpreadsheetAsAoa(file);
      if (aoa && aoa.length > 0) {
        const detectedHeaders = aoa[0].map((h, i) => h ? String(h).trim() : `Column ${i + 1}`);
        const rows = aoa.slice(1);
        setTimetableHeaders(detectedHeaders);
        setTimetableRows(rows);

        const mapping = detectTimetableColumns(detectedHeaders);
        setTimetableMapping(mapping);

        setStatusMsg(`Loaded ${rows.length} exam slots from Timetable ${file.name}.`);
        setStatusType('success');
      } else {
        setStatusMsg('No data found in uploaded Timetable file.');
        setStatusType('error');
      }
    } catch (err) {
      console.error(err);
      setStatusMsg(`Error reading Timetable file: ${err.message}`);
      setStatusType('error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Format date strictly to DD/MM/YYYY
  const formatExamDate = (val) => {
    if (!val) return '';
    if (val instanceof Date && !isNaN(val.getTime())) {
      // Add 12 hours to neutralize LMT/historic timezone offset (23:59:50 -> 11:59:50 next day)
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
      const dd = String(d.getUTCDate()).padStart(2, '0');
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
      const yyyy = d.getUTCFullYear();
      return `${dd}/${mm}/${yyyy}`;
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
    // Also parse date strings like "Tue Nov 10 2026 00:00:00 GMT+0530..." or "05 Nov 2026"
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

  // Convert time values to Excel day fraction
  const calcTimeFraction = (hr, min, ampm, directTimeVal) => {
    if (directTimeVal !== undefined && directTimeVal !== null && directTimeVal !== '') {
      if (typeof directTimeVal === 'number' && directTimeVal < 1) return directTimeVal;
      const cleaned = cleanTimeText(directTimeVal);
      if (cleaned) {
        const match = cleaned.match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?/i);
        if (match) {
          let h = parseInt(match[1], 10);
          const m = parseInt(match[2], 10);
          const ap = match[3] ? match[3].toUpperCase() : '';
          if (ap === 'PM' && h !== 12) h += 12;
          if (ap === 'AM' && h === 12) h = 0;
          return ((h * 60) + m) / 1440;
        }
      }
    }

    if (hr !== undefined && hr !== null && hr !== '') {
      let h = parseInt(hr, 10);
      const m = parseInt(min || 0, 10);
      const ap = String(ampm || '').toUpperCase().trim();
      if (ap === 'PM' && h !== 12) h += 12;
      if (ap === 'AM' && h === 12) h = 0;
      return ((h * 60) + m) / 1440;
    }
    return '';
  };

  // Format fraction or time to display string e.g. "1:30:00 PM"
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
      const cleaned = cleanTimeText(str);
      const match = cleaned.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/);
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
      return cleaned;
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
    return cleanTimeText(str);
  };

  // Main Report Generation Engine
  const reportData = useMemo(() => {
    if (preExamRows.length === 0 || timetableRows.length === 0) return { rows: [], unmapped: [] };

    // 1. Build Timetable Indexes
    const ttPgmMap = new Map();
    const ttCodeMap = new Map();

    timetableRows.forEach(row => {
      const cCodeRaw = timetableMapping.courseCode !== -1 && row[timetableMapping.courseCode] ? cleanDisplayField(row[timetableMapping.courseCode]) : '';
      if (!cCodeRaw) return;
      const cleanCode = normalizeKeyText(cCodeRaw);

      const termAbbr = timetableMapping.termAbbr !== -1 && row[timetableMapping.termAbbr] ? String(row[timetableMapping.termAbbr]).trim().toLowerCase() : '';
      const parentGroup = timetableMapping.parentGroupName !== -1 && row[timetableMapping.parentGroupName] ? cleanDisplayField(row[timetableMapping.parentGroupName]) : '';
      const immGroup = timetableMapping.immediateGroupName !== -1 && row[timetableMapping.immediateGroupName] ? cleanDisplayField(row[timetableMapping.immediateGroupName]) : '';
      const courseName = timetableMapping.courseName !== -1 && row[timetableMapping.courseName] ? cleanDisplayField(row[timetableMapping.courseName]) : '';

      const dateRaw = timetableMapping.examDate !== -1 ? row[timetableMapping.examDate] : '';
      const examDate = formatExamDate(dateRaw);

      const hrStart = timetableMapping.startHr !== -1 ? row[timetableMapping.startHr] : '';
      const minStart = timetableMapping.startMin !== -1 ? row[timetableMapping.startMin] : '';
      const ampmStart = timetableMapping.startAmPm !== -1 ? row[timetableMapping.startAmPm] : '';
      const directStart = timetableMapping.courseStartTime !== -1 ? row[timetableMapping.courseStartTime] : '';
      const startTimeFrac = calcTimeFraction(hrStart, minStart, ampmStart, directStart);

      const hrEnd = timetableMapping.endHr !== -1 ? row[timetableMapping.endHr] : '';
      const minEnd = timetableMapping.endMin !== -1 ? row[timetableMapping.endMin] : '';
      const ampmEnd = timetableMapping.endAmPm !== -1 ? row[timetableMapping.endAmPm] : '';
      const directEnd = timetableMapping.courseEndTime !== -1 ? row[timetableMapping.courseEndTime] : '';
      const endTimeFrac = calcTimeFraction(hrEnd, minEnd, ampmEnd, directEnd);

      const entry = {
        cCodeRaw,
        cleanCode,
        courseName,
        examDate,
        startTime: startTimeFrac,
        endTime: endTimeFrac,
        parentGroup,
        immediateGroup: immGroup || parentGroup
      };

      if (termAbbr) {
        const fullKey = `${termAbbr}__${cleanCode}`;
        ttPgmMap.set(fullKey, entry);
      }
      if (!ttCodeMap.has(cleanCode)) {
        ttCodeMap.set(cleanCode, entry);
      }
    });

    // 2. Group Pre-Exam Candidates by (Programme Name + Course Code + Venue Code)
    const groupMap = new Map();
    const unmappedCourses = new Set();

    preExamRows.forEach(row => {
      const pgm = preExamMapping.programmeName !== -1 && row[preExamMapping.programmeName] ? cleanDisplayField(row[preExamMapping.programmeName]) : '';
      const code = preExamMapping.courseCode !== -1 && row[preExamMapping.courseCode] ? cleanDisplayField(row[preExamMapping.courseCode]) : '';
      const name = preExamMapping.courseName !== -1 && row[preExamMapping.courseName] ? cleanDisplayField(row[preExamMapping.courseName]) : '';
      const vCode = preExamMapping.venueCode !== -1 && row[preExamMapping.venueCode] ? cleanDisplayField(row[preExamMapping.venueCode]) : 'DEFAULT';
      const vName = preExamMapping.venueName !== -1 && row[preExamMapping.venueName] ? cleanDisplayField(row[preExamMapping.venueName]) : 'General Examination Venue';

      if (!code) return;

      const parent = preExamMapping.parentGroupName !== -1 && row[preExamMapping.parentGroupName] ? cleanDisplayField(row[preExamMapping.parentGroupName]) : '';
      const imm = preExamMapping.immediateGroupName !== -1 && row[preExamMapping.immediateGroupName] ? cleanDisplayField(row[preExamMapping.immediateGroupName]) : '';

      // Count candidate registrations
      let countToAdd = 1;
      if (preExamMapping.studentCount !== -1 && row[preExamMapping.studentCount] !== undefined && row[preExamMapping.studentCount] !== null) {
        const parsed = parseInt(String(row[preExamMapping.studentCount]).replace(/,/g, ''), 10);
        if (!isNaN(parsed) && parsed > 0) countToAdd = parsed;
      }

      const groupKey = `${pgm}__${normalizeKeyText(code)}__${normalizeKeyText(vCode)}`;

      if (!groupMap.has(groupKey)) {
        groupMap.set(groupKey, {
          pgm,
          code,
          name,
          vCode,
          vName,
          parent,
          imm,
          count: 0
        });
      }

      const item = groupMap.get(groupKey);
      item.count += countToAdd;
      if (!item.parent && parent) item.parent = parent;
      if (!item.imm && imm) item.imm = imm;
      if (!item.name && name) item.name = name;
    });

    // 3. Assemble Output Rows with Exact 13 Columns
    const outputList = [];

    groupMap.forEach((val) => {
      const cleanCode = normalizeKeyText(val.code);
      const pgmLower = val.pgm.toLowerCase();

      // Find matching Timetable slot
      let matched = null;
      for (const [ttKey, entry] of ttPgmMap.entries()) {
        if (ttKey.startsWith(pgmLower) && ttKey.endsWith('__' + cleanCode)) {
          matched = entry;
          break;
        }
      }
      if (!matched) {
        matched = ttCodeMap.get(cleanCode);
      }

      if (!matched) {
        unmappedCourses.add(val.code);
      }

      // Group resolution: copy ImmediateParentGroupName if ParentGroupName is blank
      let parentGroup = val.parent || val.imm || matched?.parentGroup || matched?.immediateGroup || '';
      let immediateGroup = val.imm || matched?.immediateGroup || parentGroup;

      const examDate = formatExamDate(matched?.examDate || '');
      const startTime = formatTimeDisplay(matched?.startTime);
      const endTime = formatTimeDisplay(matched?.endTime);
      const courseName = val.name || matched?.courseName || '';

      // Concat columns
      const concat1 = `${parentGroup}${val.code}`;
      const concat2 = `${val.pgm}${examDate}${val.code}${val.vCode}`;

      outputList.push({
        programmeName: val.pgm,
        parentGroupName: parentGroup,
        immediateParentGroup: immediateGroup,
        courseCode: val.code,
        courseName,
        examDate,
        courseStartTime: startTime,
        courseEndTime: endTime,
        studentCount: val.count,
        venueCode: val.vCode,
        venueName: val.vName,
        concatGroupCourse: concat1,
        concatPgmDateCourseVenue: concat2,
        sortDate: examDate
      });
    });

    return { rows: outputList, unmapped: Array.from(unmappedCourses) };
  }, [preExamRows, timetableRows, preExamMapping, timetableMapping]);

  // Distinct Values for Filter Dropdowns
  const venuesList = useMemo(() => {
    const s = new Set();
    reportData.rows.forEach(r => { if (r.venueCode) s.add(r.venueCode); });
    return Array.from(s).sort();
  }, [reportData.rows]);

  const datesList = useMemo(() => {
    const s = new Set();
    reportData.rows.forEach(r => { if (r.examDate) s.add(r.examDate); });
    return Array.from(s).sort();
  }, [reportData.rows]);

  const groupsList = useMemo(() => {
    const s = new Set();
    reportData.rows.forEach(r => { if (r.parentGroupName) s.add(r.parentGroupName); });
    return Array.from(s).sort();
  }, [reportData.rows]);

  // Filtered & Sorted Display Rows
  const filteredRows = useMemo(() => {
    let result = reportData.rows;

    if (selectedVenue !== 'ALL') {
      result = result.filter(r => r.venueCode === selectedVenue);
    }
    if (selectedDate !== 'ALL') {
      result = result.filter(r => r.examDate === selectedDate);
    }
    if (selectedGroup !== 'ALL') {
      result = result.filter(r => r.parentGroupName === selectedGroup);
    }

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      result = result.filter(r =>
        r.programmeName.toLowerCase().includes(q) ||
        r.courseCode.toLowerCase().includes(q) ||
        r.courseName.toLowerCase().includes(q) ||
        r.venueCode.toLowerCase().includes(q) ||
        r.venueName.toLowerCase().includes(q) ||
        r.examDate.toLowerCase().includes(q) ||
        r.parentGroupName.toLowerCase().includes(q) ||
        String(r.studentCount).includes(q)
      );
    }

    // Sort order
    return [...result].sort((a, b) => {
      if (sortBy === 'date') {
        if (a.examDate !== b.examDate) return a.examDate.localeCompare(b.examDate);
        return a.courseCode.localeCompare(b.courseCode);
      }
      if (sortBy === 'course') {
        return a.courseCode.localeCompare(b.courseCode);
      }
      if (sortBy === 'venue') {
        if (a.venueCode !== b.venueCode) return a.venueCode.localeCompare(b.venueCode);
        return a.examDate.localeCompare(b.examDate);
      }
      if (sortBy === 'count') {
        return b.studentCount - a.studentCount;
      }
      return 0;
    });
  }, [reportData.rows, selectedVenue, selectedDate, selectedGroup, searchQuery, sortBy]);

  // Overall Statistics
  const stats = useMemo(() => {
    const totalAllocations = reportData.rows.length;
    const totalStudents = reportData.rows.reduce((sum, r) => sum + r.studentCount, 0);
    const uniqueCourses = new Set(reportData.rows.map(r => normalizeKeyText(r.courseCode))).size;
    const uniqueVenues = new Set(reportData.rows.map(r => normalizeKeyText(r.venueCode))).size;
    const uniqueDates = new Set(reportData.rows.map(r => r.examDate).filter(Boolean)).size;

    return { totalAllocations, totalStudents, uniqueCourses, uniqueVenues, uniqueDates };
  }, [reportData.rows]);

  // Export to Excel with Exact 13 Columns
  const exportToExcel = () => {
    if (reportData.rows.length === 0) return;

    const exportRows = reportData.rows.map(r => ({
      'Programme Name': r.programmeName,
      'Parent Group Name': r.parentGroupName,
      'Immediate Parent Group': r.immediateParentGroup,
      'Course  Code': r.courseCode,
      'Course  Name': r.courseName,
      'Exam Date': formatExamDate(r.examDate),
      'Course  Start Time': formatTimeDisplay(r.courseStartTime),
      'Course  End Time': formatTimeDisplay(r.courseEndTime),
      'Student Count': r.studentCount,
      'Venue Code': r.venueCode,
      'Venue Name': r.venueName,
      'Concatenate_Groupname_Course Code': r.concatGroupCourse,
      'PGM_Date_CourseCode_Venue': r.concatPgmDateCourseVenue
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);

    // Apply standard column widths
    ws['!cols'] = [
      { wch: 45 }, // Programme Name
      { wch: 18 }, // Parent Group Name
      { wch: 22 }, // Immediate Parent Group
      { wch: 16 }, // Course Code
      { wch: 35 }, // Course Name
      { wch: 14 }, // Exam Date
      { wch: 16 }, // Course Start Time
      { wch: 16 }, // Course End Time
      { wch: 14 }, // Student Count
      { wch: 12 }, // Venue Code
      { wch: 38 }, // Venue Name
      { wch: 30 }, // Concat 1
      { wch: 55 }  // Concat 2
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Course_Venue_Date_Report');

    const fileName = `Course_Wise_Venue_Wise_Datewise_Report_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  // Export to CSV
  const exportToCsv = () => {
    if (reportData.rows.length === 0) return;

    const exportRows = reportData.rows.map(r => ({
      'Programme Name': r.programmeName,
      'Parent Group Name': r.parentGroupName,
      'Immediate Parent Group': r.immediateParentGroup,
      'Course  Code': r.courseCode,
      'Course  Name': r.courseName,
      'Exam Date': r.examDate,
      'Course  Start Time': formatTimeDisplay(r.courseStartTime),
      'Course  End Time': formatTimeDisplay(r.courseEndTime),
      'Student Count': r.studentCount,
      'Venue Code': r.venueCode,
      'Venue Name': r.venueName,
      'Concatenate_Groupname_Course Code': r.concatGroupCourse,
      'PGM_Date_CourseCode_Venue': r.concatPgmDateCourseVenue
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const csvContent = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Course_Wise_Venue_Wise_Datewise_Report_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="studio-container" style={{ padding: '24px 32px', maxWidth: '1440px', margin: '0 auto' }}>
      {/* Top Navigation */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
        <button
          onClick={() => navigate('/')}
          className="btn-secondary"
          style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '8px 14px', fontSize: '13px' }}
        >
          <ArrowLeft size={16} /> Back to Dashboard
        </button>

        <div style={{ display: 'flex', gap: '10px' }}>
          <button
            onClick={() => setShowConfig(!showConfig)}
            className="btn-secondary"
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
          >
            <Sliders size={15} /> {showConfig ? 'Hide Column Mapping' : 'Column Mapping'}
          </button>
        </div>
      </div>

      {/* Page Title & Subtitle */}
      <div style={{ marginBottom: '24px' }}>
        <h1 style={{ fontSize: '26px', fontWeight: 800, margin: '0 0 6px 0', color: 'var(--ink)' }}>
          Course Wise Venue Wise Datewise Report
        </h1>
        <p style={{ margin: 0, fontSize: '14px', color: 'var(--muted)' }}>
          Upload your <strong>Pre-Exam Student Data</strong> and <strong>Timetable</strong> to compile complete course allocations with session timing, student counts, and validation keys.
        </p>
      </div>

      {/* Dual File Upload Cards */}
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginBottom: '24px' }}>
        {/* Source 1: Pre-Exam / Student Data */}
        <div className="card" style={{ padding: '22px', border: '1.5px dashed var(--border)', background: 'var(--panel)', borderRadius: '12px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '12px' }}>
            <div style={{ background: 'rgba(23,107,135,0.1)', padding: '10px', borderRadius: '8px', color: 'var(--accent)' }}>
              <FileSpreadsheet size={22} />
            </div>
            <div>
              <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>Source 1: Pre-Exam / Student Data</h3>
              <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Student candidates, courses, and examination venues (.xlsx, .csv)</span>
            </div>
          </div>

          <label
            style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '24px',
              borderRadius: '8px',
              background: 'var(--bg)',
              border: '1px dashed var(--accent)',
              cursor: 'pointer',
              textAlign: 'center'
            }}
          >
            <Upload size={24} style={{ color: 'var(--accent)', marginBottom: '8px' }} />
            <span style={{ fontWeight: 600, fontSize: '13.5px' }}>
              {preExamFile ? preExamFile.name : 'Click to Upload Pre-Exam File'}
            </span>
            <span style={{ fontSize: '12px', color: 'var(--muted)', marginTop: '4px' }}>
              {preExamRows.length > 0 ? `${preExamRows.length} rows loaded` : 'Supports standard ERP pre-exam export'}
            </span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handlePreExamUpload} style={{ display: 'none' }} />
          </label>
        </div>

        {/* Source 2: Timetable Data */}
        <div className="card" style={{ padding: '22px', border: '1.5px dashed var(--border)', background: 'var(--panel)', borderRadius: '12px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '12px' }}>
            <div style={{ background: 'rgba(23,107,135,0.1)', padding: '10px', borderRadius: '8px', color: 'var(--accent)' }}>
              <Calendar size={22} />
            </div>
            <div>
              <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>Source 2: Examination Timetable</h3>
              <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Exam dates, session timings, and course codes (.xlsx, .csv)</span>
            </div>
          </div>

          <label
            style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '24px',
              borderRadius: '8px',
              background: 'var(--bg)',
              border: '1px dashed var(--accent)',
              cursor: 'pointer',
              textAlign: 'center'
            }}
          >
            <Upload size={24} style={{ color: 'var(--accent)', marginBottom: '8px' }} />
            <span style={{ fontWeight: 600, fontSize: '13.5px' }}>
              {timetableFile ? timetableFile.name : 'Click to Upload Timetable File'}
            </span>
            <span style={{ fontSize: '12px', color: 'var(--muted)', marginTop: '4px' }}>
              {timetableRows.length > 0 ? `${timetableRows.length} exam slots loaded` : 'Supports official timetable schedules'}
            </span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleTimetableUpload} style={{ display: 'none' }} />
          </label>
        </div>
      </div>

      {/* Advanced Column Mapping Accordion */}
      {showConfig && (
        <div className="card" style={{ padding: '20px 24px', marginBottom: '24px', background: 'var(--panel)', border: '1px solid var(--accent)', borderRadius: '10px' }}>
          <h4 style={{ margin: '0 0 14px 0', fontSize: '15px', color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Sliders size={16} /> Column Header Mappings
          </h4>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '24px' }}>
            {/* Pre-Exam Column Mapping */}
            <div>
              <h5 style={{ margin: '0 0 10px 0', fontSize: '13.5px', fontWeight: 700 }}>Source 1: Pre-Exam Columns</h5>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '10px' }}>
                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>Programme Name</label>
                  <select
                    value={preExamMapping.programmeName}
                    onChange={(e) => setPreExamMapping({ ...preExamMapping, programmeName: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {preExamHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>

                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>Course Code / Paper Code</label>
                  <select
                    value={preExamMapping.courseCode}
                    onChange={(e) => setPreExamMapping({ ...preExamMapping, courseCode: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {preExamHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>

                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>Course Name / Paper Name</label>
                  <select
                    value={preExamMapping.courseName}
                    onChange={(e) => setPreExamMapping({ ...preExamMapping, courseName: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {preExamHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>

                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>Venue Code</label>
                  <select
                    value={preExamMapping.venueCode}
                    onChange={(e) => setPreExamMapping({ ...preExamMapping, venueCode: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {preExamHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>

                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>Venue Name</label>
                  <select
                    value={preExamMapping.venueName}
                    onChange={(e) => setPreExamMapping({ ...preExamMapping, venueName: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {preExamHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>

                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>Parent Group Name</label>
                  <select
                    value={preExamMapping.parentGroupName}
                    onChange={(e) => setPreExamMapping({ ...preExamMapping, parentGroupName: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {preExamHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>
              </div>
            </div>

            {/* Timetable Column Mapping */}
            <div>
              <h5 style={{ margin: '0 0 10px 0', fontSize: '13.5px', fontWeight: 700 }}>Source 2: Timetable Columns</h5>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '10px' }}>
                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>Course Code</label>
                  <select
                    value={timetableMapping.courseCode}
                    onChange={(e) => setTimetableMapping({ ...timetableMapping, courseCode: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {timetableHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>

                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>Exam Date</label>
                  <select
                    value={timetableMapping.examDate}
                    onChange={(e) => setTimetableMapping({ ...timetableMapping, examDate: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {timetableHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>

                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>TermAbbr / Programme</label>
                  <select
                    value={timetableMapping.termAbbr}
                    onChange={(e) => setTimetableMapping({ ...timetableMapping, termAbbr: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {timetableHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>

                <div>
                  <label style={{ fontSize: '12px', color: 'var(--muted)' }}>Parent Group Name</label>
                  <select
                    value={timetableMapping.parentGroupName}
                    onChange={(e) => setTimetableMapping({ ...timetableMapping, parentGroupName: parseInt(e.target.value, 10) })}
                    style={{ width: '100%', padding: '6px', fontSize: '12.5px', marginTop: '4px' }}
                  >
                    <option value="-1">-- Unmapped --</option>
                    {timetableHeaders.map((h, i) => <option key={i} value={i}>{h}</option>)}
                  </select>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Status Bar */}
      {statusMsg && (
        <div
          style={{
            padding: '12px 18px',
            marginBottom: '24px',
            borderRadius: '8px',
            fontSize: '13.5px',
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            background: statusType === 'error' ? 'rgba(239, 68, 68, 0.1)' : statusType === 'success' ? 'rgba(34, 197, 94, 0.1)' : 'var(--panel)',
            color: statusType === 'error' ? '#ef4444' : statusType === 'success' ? '#16a34a' : 'var(--ink)',
            border: `1px solid ${statusType === 'error' ? '#ef4444' : statusType === 'success' ? '#22c55e' : 'var(--border)'}`
          }}
        >
          {statusType === 'success' && <CheckCircle2 size={18} />}
          {statusType === 'error' && <AlertCircle size={18} />}
          <span>{statusMsg}</span>
        </div>
      )}

      {/* KPI Statistic Cards */}
      {reportData.rows.length > 0 && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(5, 1fr)', gap: '14px', marginBottom: '24px' }}>
          <div className="card" style={{ padding: '16px', borderRadius: '10px', background: 'var(--panel)' }}>
            <span style={{ fontSize: '12px', color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Total Allocations</span>
            <strong style={{ fontSize: '22px', color: 'var(--accent)' }}>{stats.totalAllocations}</strong>
          </div>

          <div className="card" style={{ padding: '16px', borderRadius: '10px', background: 'var(--panel)' }}>
            <span style={{ fontSize: '12px', color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Total Candidates</span>
            <strong style={{ fontSize: '22px', color: 'var(--ink)' }}>{stats.totalStudents}</strong>
          </div>

          <div className="card" style={{ padding: '16px', borderRadius: '10px', background: 'var(--panel)' }}>
            <span style={{ fontSize: '12px', color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Unique Courses</span>
            <strong style={{ fontSize: '22px', color: 'var(--ink)' }}>{stats.uniqueCourses}</strong>
          </div>

          <div className="card" style={{ padding: '16px', borderRadius: '10px', background: 'var(--panel)' }}>
            <span style={{ fontSize: '12px', color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Exam Venues</span>
            <strong style={{ fontSize: '22px', color: 'var(--ink)' }}>{stats.uniqueVenues}</strong>
          </div>

          <div className="card" style={{ padding: '16px', borderRadius: '10px', background: 'var(--panel)' }}>
            <span style={{ fontSize: '12px', color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Exam Dates</span>
            <strong style={{ fontSize: '22px', color: 'var(--ink)' }}>{stats.uniqueDates}</strong>
          </div>
        </div>
      )}

      {/* Unmapped Warning Box */}
      {reportData.unmapped && reportData.unmapped.length > 0 && (
        <div
          style={{
            padding: '14px 18px',
            marginBottom: '20px',
            borderRadius: '8px',
            background: 'rgba(234, 179, 8, 0.1)',
            border: '1px solid #eab308',
            color: '#854d0e',
            fontSize: '13px'
          }}
        >
          <strong>Notice:</strong> {reportData.unmapped.length} course(s) from Pre-Exam could not be found in the Timetable:
          <span style={{ marginLeft: '6px', fontFamily: 'monospace' }}>{reportData.unmapped.join(', ')}</span>
        </div>
      )}

      {/* Main Results Container */}
      {reportData.rows.length > 0 && (
        <div className="card" style={{ padding: '24px', borderRadius: '12px', background: 'var(--panel)' }}>
          {/* Action Toolbar */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px', marginBottom: '20px' }}>
            {/* Filters */}
            <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap', alignItems: 'center' }}>
              <div style={{ position: 'relative', width: '240px' }}>
                <Search size={15} style={{ position: 'absolute', left: '10px', top: '10px', color: 'var(--muted)' }} />
                <input
                  type="text"
                  placeholder="Search courses, venues..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  style={{ width: '100%', padding: '8px 10px 8px 32px', fontSize: '13px', borderRadius: '6px', border: '1px solid var(--border)' }}
                />
              </div>

              {/* Venue Filter */}
              <select
                value={selectedVenue}
                onChange={(e) => setSelectedVenue(e.target.value)}
                style={{ padding: '8px 12px', fontSize: '13px', borderRadius: '6px', border: '1px solid var(--border)' }}
              >
                <option value="ALL">All Venues ({venuesList.length})</option>
                {venuesList.map(v => <option key={v} value={v}>Venue: {v}</option>)}
              </select>

              {/* Date Filter */}
              <select
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                style={{ padding: '8px 12px', fontSize: '13px', borderRadius: '6px', border: '1px solid var(--border)' }}
              >
                <option value="ALL">All Dates ({datesList.length})</option>
                {datesList.map(d => <option key={d} value={d}>Date: {d}</option>)}
              </select>

              {/* Group Filter */}
              <select
                value={selectedGroup}
                onChange={(e) => setSelectedGroup(e.target.value)}
                style={{ padding: '8px 12px', fontSize: '13px', borderRadius: '6px', border: '1px solid var(--border)' }}
              >
                <option value="ALL">All Parent Groups</option>
                {groupsList.map(g => <option key={g} value={g}>Group: {g}</option>)}
              </select>

              {/* Sort Dropdown */}
              <select
                value={sortBy}
                onChange={(e) => setSortBy(e.target.value)}
                style={{ padding: '8px 12px', fontSize: '13px', borderRadius: '6px', border: '1px solid var(--border)' }}
              >
                <option value="date">Sort by Exam Date</option>
                <option value="course">Sort by Course Code</option>
                <option value="venue">Sort by Venue Code</option>
                <option value="count">Sort by Student Count</option>
              </select>
            </div>

            {/* Export Buttons */}
            <div style={{ display: 'flex', gap: '10px' }}>
              <button
                onClick={exportToCsv}
                className="btn-secondary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 14px', fontSize: '13px' }}
              >
                <Download size={15} /> CSV
              </button>

              <button
                onClick={exportToExcel}
                className="btn-primary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 16px', fontSize: '13px', fontWeight: 600 }}
              >
                <Download size={15} /> Download Output Excel Report
              </button>
            </div>
          </div>

          {/* Table Container */}
          <div style={{ overflowX: 'auto', border: '1px solid var(--border)', borderRadius: '8px' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px', textAlign: 'left' }}>
              <thead>
                <tr style={{ background: 'var(--bg)', borderBottom: '2px solid var(--border)', color: 'var(--ink)' }}>
                  <th style={{ padding: '10px 8px', textAlign: 'center', width: '40px' }}>#</th>
                  <th style={{ padding: '10px 10px', minWidth: '180px' }}>Programme Name</th>
                  <th style={{ padding: '10px 8px', minWidth: '90px' }}>Parent Group</th>
                  <th style={{ padding: '10px 8px', minWidth: '110px' }}>Course Code</th>
                  <th style={{ padding: '10px 10px', minWidth: '180px' }}>Course Name</th>
                  <th style={{ padding: '10px 8px', minWidth: '95px', textAlign: 'center' }}>Exam Date</th>
                  <th style={{ padding: '10px 8px', minWidth: '90px', textAlign: 'center' }}>Start Time</th>
                  <th style={{ padding: '10px 8px', minWidth: '90px', textAlign: 'center' }}>End Time</th>
                  <th style={{ padding: '10px 8px', textAlign: 'center', width: '70px' }}>Students</th>
                  <th style={{ padding: '10px 8px', textAlign: 'center', width: '60px' }}>Venue</th>
                  <th style={{ padding: '10px 10px', minWidth: '180px' }}>Venue Name</th>
                  <th style={{ padding: '10px 8px', minWidth: '150px' }}>Verification Key</th>
                </tr>
              </thead>
              <tbody>
                {filteredRows.slice(0, 500).map((row, idx) => (
                  <tr
                    key={idx}
                    style={{
                      borderBottom: '1px solid var(--border)',
                      background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.015)'
                    }}
                  >
                    <td style={{ padding: '8px', textAlign: 'center', color: 'var(--muted)', fontWeight: 600 }}>{idx + 1}</td>
                    <td style={{ padding: '8px 10px', fontWeight: 500 }}>{row.programmeName}</td>
                    <td style={{ padding: '8px' }}>
                      <span style={{ background: 'rgba(23,107,135,0.08)', padding: '2px 6px', borderRadius: '4px', fontSize: '11.5px', fontWeight: 600, color: 'var(--accent)' }}>
                        {row.parentGroupName}
                      </span>
                    </td>
                    <td style={{ padding: '8px', fontFamily: 'monospace', fontWeight: 600 }}>{row.courseCode}</td>
                    <td style={{ padding: '8px 10px' }}>{row.courseName}</td>
                    <td style={{ padding: '8px', textAlign: 'center', whiteSpace: 'nowrap', fontWeight: 600 }}>{formatExamDate(row.examDate)}</td>
                    <td style={{ padding: '8px', textAlign: 'center', whiteSpace: 'nowrap', color: 'var(--muted)' }}>
                      {formatTimeDisplay(row.courseStartTime)}
                    </td>
                    <td style={{ padding: '8px', textAlign: 'center', whiteSpace: 'nowrap', color: 'var(--muted)' }}>
                      {formatTimeDisplay(row.courseEndTime)}
                    </td>
                    <td style={{ padding: '8px', textAlign: 'center' }}>
                      <span style={{ background: '#e0f2fe', color: '#0369a1', padding: '2px 8px', borderRadius: '10px', fontWeight: 700 }}>
                        {row.studentCount}
                      </span>
                    </td>
                    <td style={{ padding: '8px', textAlign: 'center', fontWeight: 700 }}>{row.venueCode}</td>
                    <td style={{ padding: '8px 10px', color: 'var(--muted)' }}>{row.venueName}</td>
                    <td style={{ padding: '8px', fontFamily: 'monospace', fontSize: '11px', color: 'var(--muted)', maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={row.concatPgmDateCourseVenue}>
                      {row.concatPgmDateCourseVenue}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div style={{ marginTop: '14px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '12.5px', color: 'var(--muted)' }}>
            <div>
              Showing <strong>{filteredRows.length}</strong> of <strong>{reportData.rows.length}</strong> allocation rows
            </div>
            <div>
              Total Filtered Students: <strong>{filteredRows.reduce((sum, r) => sum + r.studentCount, 0)}</strong>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default CourseWiseVenueDatePage;
