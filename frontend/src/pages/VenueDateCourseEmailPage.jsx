import React, { useState, useMemo, useEffect } from 'react';
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
  Mail,
  Copy,
  Check,
  Plus,
  Trash2,
  Edit2,
  FileText,
  Send,
  ExternalLink,
  ChevronDown
} from 'lucide-react';
import { readSpreadsheetAsAoa, cleanTimeText, normalizeKeyText, cleanDisplayField } from '../utils/excelParser';
import { DEFAULT_COLLEGE_EMAILS } from '../utils/venueEmailHelpers';

export { DEFAULT_COLLEGE_EMAILS };

// Sample Dataset for instant demonstration
export const SAMPLE_EXAM_SCHEDULE_ROWS = [
  {
    'Exam Date': '15/11/2026',
    'Venue Code': '201',
    'Venue Name': 'Govt. Brennen College, Thalassery',
    'Course Code': 'KU4DSCHIS101',
    'Course Name': 'Early Indian History and Culture',
    'Student Count': 45,
    'Session': '09:30 AM - 12:30 PM',
    'College Email': 'brennencollege@gmail.com'
  },
  {
    'Exam Date': '15/11/2026',
    'Venue Code': '209',
    'Venue Name': 'Payyannur College, Payyannur',
    'Course Code': 'KU4DSCHIS101',
    'Course Name': 'Early Indian History and Culture',
    'Student Count': 38,
    'Session': '09:30 AM - 12:30 PM',
    'College Email': 'payyannurcollege@gmail.com'
  },
  {
    'Exam Date': '15/11/2026',
    'Venue Code': '201',
    'Venue Name': 'Govt. Brennen College, Thalassery',
    'Course Code': 'KU4DSCENG101',
    'Course Name': 'Introduction to Literary Studies',
    'Student Count': 52,
    'Session': '01:30 PM - 04:30 PM',
    'College Email': 'brennencollege@gmail.com'
  },
  {
    'Exam Date': '17/11/2026',
    'Venue Code': '212',
    'Venue Name': 'Sree Narayana College, Kannur',
    'Course Code': 'KU4DSCCOM101',
    'Course Name': 'Financial Accounting Fundamentals',
    'Student Count': 64,
    'Session': '09:30 AM - 12:30 PM',
    'College Email': 'sncollegekannur@gmail.com'
  },
  {
    'Exam Date': '17/11/2026',
    'Venue Code': '201',
    'Venue Name': 'Govt. Brennen College, Thalassery',
    'Course Code': 'KU4DSCCOM101',
    'Course Name': 'Financial Accounting Fundamentals',
    'Student Count': 48,
    'Session': '09:30 AM - 12:30 PM',
    'College Email': 'brennencollege@gmail.com'
  },
  {
    'Exam Date': '19/11/2026',
    'Venue Code': '435',
    'Venue Name': 'Department of Anthropology',
    'Course Code': 'PFPSYGS25',
    'Course Name': 'Clinical Psychology (FYIMP)',
    'Student Count': 30,
    'Session': '09:30 AM - 12:30 PM',
    'College Email': 'anthropology@kannuruniv.ac.in'
  },
  {
    'Exam Date': '19/11/2026',
    'Venue Code': '432',
    'Venue Name': 'School of Behavioural Sciences, Mangattuparamba',
    'Course Code': 'PFCOMGS25',
    'Course Name': 'Commerce (FYIMP)',
    'Student Count': 35,
    'Session': '09:30 AM - 12:30 PM',
    'College Email': 'behaviouralsciences@kannuruniv.ac.in'
  },
  {
    'Exam Date': '22/11/2026',
    'Venue Code': '200',
    'Venue Name': 'Govt. College Mananthavady',
    'Course Code': 'KU4MINPOL101',
    'Course Name': 'Principles of Political Theory',
    'Student Count': 27,
    'Session': '09:30 AM - 12:30 PM',
    'College Email': 'gcmananthavady@gmail.com'
  },
  {
    'Exam Date': '22/11/2026',
    'Venue Code': '209',
    'Venue Name': 'Payyannur College, Payyannur',
    'Course Code': 'KU4MINPOL101',
    'Course Name': 'Principles of Political Theory',
    'Student Count': 41,
    'Session': '09:30 AM - 12:30 PM',
    'College Email': 'payyannurcollege@gmail.com'
  }
];

/**
 * Strict date formatting to DD/MM/YYYY
 */
export const formatExamDate = (val) => {
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

/**
 * Intelligent helper to auto-detect columns from headers
 */
export const detectExamColumns = (headersList) => {
  const mapping = {
    examDate: -1,
    venueCode: -1,
    venueName: -1,
    courseCode: -1,
    courseName: -1,
    email: -1,
    studentCount: -1,
    session: -1
  };

  headersList.forEach((h, idx) => {
    if (!h) return;
    const raw = String(h).trim();
    const clean = raw.toLowerCase().replace(/[^a-z0-9]/g, '');

    // Date
    if (mapping.examDate === -1) {
      if (['examdate', 'date', 'dateofexam', 'examinationdate'].includes(clean)) {
        mapping.examDate = idx;
      } else if (clean.includes('date') && !clean.includes('birth') && !clean.includes('result')) {
        mapping.examDate = idx;
      }
    }

    // Venue Code
    if (mapping.venueCode === -1) {
      if (['venuecode', 'examcentercode', 'centercode', 'centrecode', 'collegecode', 'colcode', 'collegeid', 'colid'].includes(clean)) {
        mapping.venueCode = idx;
      } else if ((clean.includes('venue') || clean.includes('center') || clean.includes('centre') || clean.includes('college')) && clean.includes('code')) {
        mapping.venueCode = idx;
      }
    }

    // Venue Name
    if (mapping.venueName === -1) {
      if (['venuename', 'examcentername', 'centername', 'centrename', 'collegename', 'clgname', 'college'].includes(clean)) {
        mapping.venueName = idx;
      } else if ((clean.includes('venue') || clean.includes('center') || clean.includes('centre') || clean.includes('college')) && clean.includes('name')) {
        mapping.venueName = idx;
      }
    }

    // Course Code
    if (mapping.courseCode === -1) {
      if (['coursecode', 'papercode', 'subjectcode', 'c_code'].includes(clean)) {
        mapping.courseCode = idx;
      } else if (clean.includes('course') && clean.includes('code')) {
        mapping.courseCode = idx;
      }
    }

    // Course Name
    if (mapping.courseName === -1) {
      if (['coursename', 'papername', 'subjectname', 'subject', 'c_name'].includes(clean)) {
        mapping.courseName = idx;
      } else if (clean.includes('course') && clean.includes('name')) {
        mapping.courseName = idx;
      }
    }

    // Email
    if (mapping.email === -1) {
      if (['email', 'emailid', 'collegeemail', 'venueemail', 'principalemail', 'examemail', 'mail', 'mailid'].includes(clean)) {
        mapping.email = idx;
      } else if (clean.includes('email') || clean.includes('mail')) {
        mapping.email = idx;
      }
    }

    // Student Count
    if (mapping.studentCount === -1) {
      if (['studentcount', 'count', 'strength', 'candidatecount', 'total', 'noofstudents'].includes(clean)) {
        mapping.studentCount = idx;
      } else if (clean.includes('count') || clean.includes('strength')) {
        mapping.studentCount = idx;
      }
    }

    // Session / Time
    if (mapping.session === -1) {
      if (['session', 'time', 'examtime', 'coursetime', 'timing', 'slot'].includes(clean)) {
        mapping.session = idx;
      }
    }
  });

  return mapping;
};

/**
 * Intelligent helper to auto-detect columns in College Email Directory
 */
export const detectEmailDirectoryColumns = (headersList) => {
  const mapping = {
    code: -1,
    name: -1,
    email: -1
  };

  headersList.forEach((h, idx) => {
    if (!h) return;
    const clean = String(h).trim().toLowerCase().replace(/[^a-z0-9]/g, '');

    // Code
    if (mapping.code === -1) {
      if (['collegecode', 'venuecode', 'centercode', 'code', 'collegeid', 'id'].includes(clean)) {
        mapping.code = idx;
      } else if (clean.includes('code') || clean.includes('id')) {
        mapping.code = idx;
      }
    }

    // Name
    if (mapping.name === -1) {
      if (['collegename', 'venuename', 'centername', 'name', 'collegeref', 'institution'].includes(clean)) {
        mapping.name = idx;
      } else if (clean.includes('name') || clean.includes('college')) {
        mapping.name = idx;
      }
    }

    // Email
    if (mapping.email === -1) {
      if (['email', 'emailid', 'collegeemail', 'venueemail', 'principalemail', 'mail', 'mailid'].includes(clean)) {
        mapping.email = idx;
      } else if (clean.includes('email') || clean.includes('mail')) {
        mapping.email = idx;
      }
    }
  });

  return mapping;
};

const VenueDateCourseEmailPage = () => {
  const navigate = useNavigate();

  // Primary Data File
  const [scheduleFile, setScheduleFile] = useState(null);
  const [scheduleHeaders, setScheduleHeaders] = useState([]);
  const [scheduleRows, setScheduleRows] = useState([]);
  const [scheduleMapping, setScheduleMapping] = useState({
    examDate: -1,
    venueCode: -1,
    venueName: -1,
    courseCode: -1,
    courseName: -1,
    email: -1,
    studentCount: -1,
    session: -1
  });

  // Optional Email Directory File
  const [directoryFile, setDirectoryFile] = useState(null);
  const [directoryHeaders, setDirectoryHeaders] = useState([]);
  const [directoryRows, setDirectoryRows] = useState([]);

  // In-Memory & LocalStorage College Email Directory Map
  const [collegeEmailMap, setCollegeEmailMap] = useState(() => {
    try {
      const saved = localStorage.getItem('reportstudio_college_email_directory');
      if (saved) {
        return JSON.parse(saved);
      }
    } catch {
      // fallback
    }
    const initialMap = {};
    DEFAULT_COLLEGE_EMAILS.forEach(c => {
      initialMap[c.code.toLowerCase().trim()] = { code: c.code, name: c.name, email: c.email };
      initialMap[c.name.toLowerCase().trim()] = { code: c.code, name: c.name, email: c.email };
    });
    return initialMap;
  });

  // UI States
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Upload your Examination Schedule or Pre-Exam file to generate the Venue Date Course Email report.');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  const [showConfig, setShowConfig] = useState(false);
  const [showEmailModal, setShowEmailModal] = useState(false);
  const [showGuideModal, setShowGuideModal] = useState(false);

  // Filters & Search
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedDate, setSelectedDate] = useState('ALL');
  const [selectedVenue, setSelectedVenue] = useState('ALL');
  const [emailFilter, setEmailFilter] = useState('ALL'); // 'ALL' | 'HAS_EMAIL' | 'MISSING_EMAIL'
  const [activeTab, setActiveTab] = useState('all'); // 'all' | 'byVenue' | 'byDate'

  // Clipboard Copied State
  const [copiedEmails, setCopiedEmails] = useState(false);

  // New Directory Entry in Modal
  const [newDirCode, setNewDirCode] = useState('');
  const [newDirName, setNewDirName] = useState('');
  const [newDirEmail, setNewDirEmail] = useState('');
  const [editingEmailKey, setEditingEmailKey] = useState(null);
  const [editingEmailVal, setEditingEmailVal] = useState('');

  // Persist Email Directory Changes
  const saveEmailDirectory = (newMap) => {
    setCollegeEmailMap(newMap);
    try {
      localStorage.setItem('reportstudio_college_email_directory', JSON.stringify(newMap));
    } catch {
      // ignore
    }
  };

  // Helper to resolve email for a venue
  const resolveCollegeEmail = (vCode, vName, inlineEmail) => {
    if (inlineEmail && String(inlineEmail).includes('@')) {
      return String(inlineEmail).trim();
    }
    const cleanCode = String(vCode || '').toLowerCase().trim();
    if (cleanCode && collegeEmailMap[cleanCode]?.email) {
      return collegeEmailMap[cleanCode].email;
    }
    const cleanName = String(vName || '').toLowerCase().trim();
    if (cleanName && collegeEmailMap[cleanName]?.email) {
      return collegeEmailMap[cleanName].email;
    }
    // Substring search in directory names
    for (const [k, v] of Object.entries(collegeEmailMap)) {
      if (cleanName && (cleanName.includes(k) || k.includes(cleanName)) && v.email) {
        return v.email;
      }
    }
    return '';
  };

  // Handle Primary Schedule Upload
  const handleScheduleUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setScheduleFile(file);
    setIsProcessing(true);
    setStatusMsg(`Loading "${file.name}"...`);
    setStatusType('normal');

    try {
      const aoa = await readSpreadsheetAsAoa(file);
      if (aoa && aoa.length > 0) {
        const detectedHeaders = aoa[0].map((h, i) => h ? String(h).trim() : `Column ${i + 1}`);
        const rows = aoa.slice(1);
        setScheduleHeaders(detectedHeaders);
        setScheduleRows(rows);

        const mapping = detectExamColumns(detectedHeaders);
        setScheduleMapping(mapping);

        setStatusMsg(`Successfully loaded ${rows.length.toLocaleString()} rows from "${file.name}".`);
        setStatusType('success');
      } else {
        setStatusMsg('No data found in uploaded schedule file.');
        setStatusType('error');
      }
    } catch (err) {
      console.error(err);
      setStatusMsg(`Error reading schedule spreadsheet: ${err.message}`);
      setStatusType('error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Handle Secondary College Email Directory Upload
  const handleDirectoryUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setDirectoryFile(file);
    setIsProcessing(true);
    setStatusMsg(`Loading College Email Directory "${file.name}"...`);

    try {
      const aoa = await readSpreadsheetAsAoa(file);
      if (aoa && aoa.length > 1) {
        const headers = aoa[0];
        const rows = aoa.slice(1);
        setDirectoryHeaders(headers);
        setDirectoryRows(rows);

        const mapCols = detectEmailDirectoryColumns(headers);
        const updated = { ...collegeEmailMap };
        let addedCount = 0;

        rows.forEach(r => {
          const code = mapCols.code !== -1 && r[mapCols.code] ? String(r[mapCols.code]).trim() : '';
          const name = mapCols.name !== -1 && r[mapCols.name] ? String(r[mapCols.name]).trim() : '';
          const email = mapCols.email !== -1 && r[mapCols.email] ? String(r[mapCols.email]).trim() : '';

          if (email && email.includes('@')) {
            if (code) updated[code.toLowerCase()] = { code, name, email };
            if (name) updated[name.toLowerCase()] = { code, name, email };
            addedCount++;
          }
        });

        saveEmailDirectory(updated);
        setStatusMsg(`Imported ${addedCount} college email entries from "${file.name}".`);
        setStatusType('success');
      } else {
        setStatusMsg('Invalid or empty directory spreadsheet.');
        setStatusType('error');
      }
    } catch (err) {
      console.error(err);
      setStatusMsg(`Error reading directory spreadsheet: ${err.message}`);
      setStatusType('error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Load Built-in Realistic Sample Data
  const handleLoadSampleData = () => {
    const headers = ['Exam Date', 'Venue Code', 'Venue Name', 'Course Code', 'Course Name', 'Student Count', 'Session', 'College Email'];
    const rows = SAMPLE_EXAM_SCHEDULE_ROWS.map(r => [
      r['Exam Date'],
      r['Venue Code'],
      r['Venue Name'],
      r['Course Code'],
      r['Course Name'],
      r['Student Count'],
      r['Session'],
      r['College Email']
    ]);

    setScheduleFile({ name: 'Kannur_University_Exam_Schedule_Sample.xlsx' });
    setScheduleHeaders(headers);
    setScheduleRows(rows);
    setScheduleMapping({
      examDate: 0,
      venueCode: 1,
      venueName: 2,
      courseCode: 3,
      courseName: 4,
      studentCount: 5,
      session: 6,
      email: 7
    });

    setStatusMsg('Sample university examination dataset loaded successfully.');
    setStatusType('success');
  };

  // Main Report Aggregation Engine: Unique (Date + Venue + Course)
  const reportData = useMemo(() => {
    if (scheduleRows.length === 0) return { items: [], venues: [], dates: [], stats: { totalCombinations: 0, totalVenues: 0, totalDates: 0, totalCourses: 0, totalStudents: 0, emailCount: 0 } };

    const uniqueMap = new Map();

    scheduleRows.forEach((row, rIdx) => {
      const dateRaw = scheduleMapping.examDate !== -1 ? row[scheduleMapping.examDate] : '';
      const examDate = formatExamDate(dateRaw) || 'Date TBA';

      const vCode = scheduleMapping.venueCode !== -1 && row[scheduleMapping.venueCode] !== undefined
        ? cleanDisplayField(row[scheduleMapping.venueCode])
        : 'VENUE-01';

      const vName = scheduleMapping.venueName !== -1 && row[scheduleMapping.venueName] !== undefined
        ? cleanDisplayField(row[scheduleMapping.venueName])
        : 'Examination Centre';

      const cCode = scheduleMapping.courseCode !== -1 && row[scheduleMapping.courseCode] !== undefined
        ? cleanDisplayField(row[scheduleMapping.courseCode])
        : '';

      const cName = scheduleMapping.courseName !== -1 && row[scheduleMapping.courseName] !== undefined
        ? cleanDisplayField(row[scheduleMapping.courseName])
        : '';

      // Skip empty course slots
      if (!cCode && !cName) return;

      const inlineEmail = scheduleMapping.email !== -1 && row[scheduleMapping.email]
        ? String(row[scheduleMapping.email]).trim()
        : '';

      const email = resolveCollegeEmail(vCode, vName, inlineEmail);

      const session = scheduleMapping.session !== -1 && row[scheduleMapping.session]
        ? cleanDisplayField(row[scheduleMapping.session])
        : '';

      let studentCount = 1;
      if (scheduleMapping.studentCount !== -1 && row[scheduleMapping.studentCount] !== undefined && row[scheduleMapping.studentCount] !== null) {
        const parsed = parseInt(String(row[scheduleMapping.studentCount]).replace(/,/g, ''), 10);
        if (!isNaN(parsed) && parsed > 0) studentCount = parsed;
      }

      // Unique Tuple Key: Date + Venue Code + Course Code
      const compositeKey = [
        normalizeKeyText(examDate),
        normalizeKeyText(vCode),
        normalizeKeyText(cCode)
      ].join('|||');

      if (!uniqueMap.has(compositeKey)) {
        uniqueMap.set(compositeKey, {
          examDate,
          venueCode: vCode,
          venueName: vName,
          courseCode: cCode,
          courseName: cName,
          email,
          studentCount,
          session,
          sourceRowIndices: [rIdx]
        });
      } else {
        const existing = uniqueMap.get(compositeKey);
        existing.studentCount += studentCount;
        existing.sourceRowIndices.push(rIdx);
        if (!existing.email && email) existing.email = email;
        if (!existing.session && session) existing.session = session;
        if (!existing.venueName && vName) existing.venueName = vName;
        if (!existing.courseName && cName) existing.courseName = cName;
      }
    });

    const items = Array.from(uniqueMap.values());

    // Sort: Date (chronological) -> Venue Code -> Course Code
    items.sort((a, b) => {
      // Compare dates
      const parseD = (s) => {
        const parts = s.split('/');
        return parts.length === 3 ? new Date(parts[2], parts[1] - 1, parts[0]).getTime() : 0;
      };
      const dDiff = parseD(a.examDate) - parseD(b.examDate);
      if (dDiff !== 0) return dDiff;

      const vDiff = a.venueCode.localeCompare(b.venueCode, undefined, { numeric: true });
      if (vDiff !== 0) return vDiff;

      return a.courseCode.localeCompare(b.courseCode);
    });

    const venueSet = new Set(items.map(i => i.venueCode));
    const dateSet = new Set(items.map(i => i.examDate));
    const courseSet = new Set(items.map(i => i.courseCode));
    const totalStudents = items.reduce((acc, i) => acc + (i.studentCount || 0), 0);
    const emailCount = items.filter(i => i.email && i.email.includes('@')).length;

    return {
      items,
      venues: Array.from(venueSet).sort(),
      dates: Array.from(dateSet),
      stats: {
        totalCombinations: items.length,
        totalVenues: venueSet.size,
        totalDates: dateSet.size,
        totalCourses: courseSet.size,
        totalStudents,
        emailCount
      }
    };
  }, [scheduleRows, scheduleMapping, collegeEmailMap]);

  // Filtered Items
  const filteredItems = useMemo(() => {
    return reportData.items.filter(item => {
      if (selectedDate !== 'ALL' && item.examDate !== selectedDate) return false;
      if (selectedVenue !== 'ALL' && item.venueCode !== selectedVenue) return false;
      if (emailFilter === 'HAS_EMAIL' && (!item.email || !item.email.includes('@'))) return false;
      if (emailFilter === 'MISSING_EMAIL' && item.email && item.email.includes('@')) return false;

      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        return (
          item.examDate.toLowerCase().includes(q) ||
          item.venueCode.toLowerCase().includes(q) ||
          item.venueName.toLowerCase().includes(q) ||
          item.courseCode.toLowerCase().includes(q) ||
          item.courseName.toLowerCase().includes(q) ||
          item.email.toLowerCase().includes(q)
        );
      }
      return true;
    });
  }, [reportData.items, selectedDate, selectedVenue, emailFilter, searchQuery]);

  // Grouped by Venue for Tab 2
  const groupedByVenue = useMemo(() => {
    const map = new Map();
    filteredItems.forEach(item => {
      const vKey = item.venueCode;
      if (!map.has(vKey)) {
        map.set(vKey, {
          venueCode: item.venueCode,
          venueName: item.venueName,
          email: item.email,
          totalStudents: 0,
          offerings: []
        });
      }
      const grp = map.get(vKey);
      grp.totalStudents += item.studentCount;
      grp.offerings.push(item);
    });
    return Array.from(map.values()).sort((a, b) => a.venueCode.localeCompare(b.venueCode, undefined, { numeric: true }));
  }, [filteredItems]);

  // Grouped by Date for Tab 3
  const groupedByDate = useMemo(() => {
    const map = new Map();
    filteredItems.forEach(item => {
      const dKey = item.examDate;
      if (!map.has(dKey)) {
        map.set(dKey, {
          examDate: item.examDate,
          totalStudents: 0,
          offerings: []
        });
      }
      const grp = map.get(dKey);
      grp.totalStudents += item.studentCount;
      grp.offerings.push(item);
    });
    return Array.from(map.values());
  }, [filteredItems]);

  // Copy All Unique Emails to Clipboard
  const handleCopyEmails = () => {
    const uniqueEmails = Array.from(
      new Set(
        filteredItems
          .map(i => i.email)
          .filter(e => e && e.includes('@'))
      )
    );

    if (uniqueEmails.length === 0) {
      alert('No valid college email addresses to copy.');
      return;
    }

    const emailStr = uniqueEmails.join(', ');
    navigator.clipboard.writeText(emailStr);
    setCopiedEmails(true);
    setTimeout(() => setCopiedEmails(false), 2500);
  };

  // Inline Email Update
  const handleSaveInlineEmail = (venueCode, venueName, newEmail) => {
    const updated = { ...collegeEmailMap };
    if (venueCode) updated[venueCode.toLowerCase()] = { code: venueCode, name: venueName, email: newEmail.trim() };
    if (venueName) updated[venueName.toLowerCase()] = { code: venueCode, name: venueName, email: newEmail.trim() };
    saveEmailDirectory(updated);
    setEditingEmailKey(null);
  };

  // Export to Excel (.xlsx)
  const handleExportExcel = () => {
    if (filteredItems.length === 0) {
      alert('No records available to export.');
      return;
    }

    const wb = XLSX.utils.book_new();

    // Sheet 1: Combinations Table
    const exportRows = filteredItems.map((item, idx) => ({
      'Sl No': idx + 1,
      'Exam Date': item.examDate,
      'Venue Code': item.venueCode,
      'Venue Name': item.venueName,
      'Course Code': item.courseCode,
      'Course Name': item.courseName,
      'College Email ID': item.email || '—',
      'Student Count': item.studentCount,
      'Session / Time': item.session || '—'
    }));

    const ws1 = XLSX.utils.json_to_sheet(exportRows);
    ws1['!cols'] = [
      { wch: 8 },  // Sl No
      { wch: 14 }, // Exam Date
      { wch: 14 }, // Venue Code
      { wch: 40 }, // Venue Name
      { wch: 18 }, // Course Code
      { wch: 42 }, // Course Name
      { wch: 34 }, // Email ID
      { wch: 14 }, // Student Count
      { wch: 22 }  // Session
    ];
    XLSX.utils.book_append_sheet(wb, ws1, 'Venue_Date_Course_Email');

    // Sheet 2: Venue Email Directory & Summary
    const venueSummary = groupedByVenue.map(v => ({
      'Venue Code': v.venueCode,
      'Venue / College Name': v.venueName,
      'College Email ID': v.email || 'NOT AVAILABLE',
      'Total Exam Slots': v.offerings.length,
      'Total Candidates': v.totalStudents
    }));
    const ws2 = XLSX.utils.json_to_sheet(venueSummary);
    ws2['!cols'] = [
      { wch: 14 },
      { wch: 42 },
      { wch: 34 },
      { wch: 18 },
      { wch: 18 }
    ];
    XLSX.utils.book_append_sheet(wb, ws2, 'Venue_Email_Summary');

    const fileName = `Venue_Date_Course_College_Email_Report_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  // Export to CSV (.csv)
  const handleExportCsv = () => {
    if (filteredItems.length === 0) {
      alert('No records available to export.');
      return;
    }

    const exportRows = filteredItems.map((item, idx) => ({
      'Sl No': idx + 1,
      'Exam Date': item.examDate,
      'Venue Code': item.venueCode,
      'Venue Name': item.venueName,
      'Course Code': item.courseCode,
      'Course Name': item.courseName,
      'College Email ID': item.email || '',
      'Student Count': item.studentCount,
      'Session': item.session || ''
    }));

    const ws = XLSX.utils.json_to_sheet(exportRows);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `Venue_Date_Course_Email_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  // Download Blank / Sample College Email Directory Template (.xlsx)
  const handleDownloadDirectoryTemplate = () => {
    const wb = XLSX.utils.book_new();
    const rows = DEFAULT_COLLEGE_EMAILS.map(c => ({
      'College Code': c.code,
      'College Name': c.name,
      'Email ID': c.email
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws['!cols'] = [
      { wch: 16 },
      { wch: 42 },
      { wch: 36 }
    ];
    XLSX.utils.book_append_sheet(wb, ws, 'College_Email_Directory');
    XLSX.writeFile(wb, 'College_Email_Directory_Template.xlsx');
  };

  return (
    <div style={{ padding: '24px 32px', maxWidth: '1440px', margin: '0 auto', color: 'var(--ink)' }}>
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
            onClick={handleLoadSampleData}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', background: 'rgba(23,107,135,0.08)', color: 'var(--accent)', borderColor: 'var(--accent)' }}
          >
            <RefreshCw size={14} /> Load Sample Data
          </button>

          <button
            type="button"
            className="secondary"
            onClick={() => setShowEmailModal(true)}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
          >
            <Mail size={14} /> College Email Directory ({Object.keys(collegeEmailMap).length / 2})
          </button>

          <button
            type="button"
            className="secondary"
            onClick={handleDownloadDirectoryTemplate}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
          >
            <Download size={14} /> Email Template
          </button>

          <button
            type="button"
            className="secondary"
            onClick={() => setShowConfig(!showConfig)}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
          >
            <Sliders size={14} /> {showConfig ? 'Hide Columns' : 'Configure Columns'}
          </button>

          <button
            type="button"
            className="secondary"
            onClick={() => setShowGuideModal(true)}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
          >
            <HelpCircle size={14} /> Logic Guide
          </button>
        </div>
      </div>

      {/* Page Title & Mission */}
      <div style={{ marginBottom: '24px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '6px' }}>
          <div style={{ background: 'rgba(23,107,135,0.1)', padding: '8px', borderRadius: '8px', color: 'var(--accent)' }}>
            <Building2 size={24} />
          </div>
          <h1 style={{ fontSize: '24px', fontWeight: 800, margin: 0, color: 'var(--ink)' }}>
            Venue Date Course College Email Report
          </h1>
        </div>
        <p style={{ margin: 0, fontSize: '14px', color: 'var(--muted)' }}>
          Compile deduplicated, unique combinations of <strong>Date</strong>, <strong>Venue</strong>, <strong>Course Code</strong>, and <strong>Course Name</strong> along with the official <strong>College Email ID</strong> of each examination venue for seamless question paper dispatch and circulars.
        </p>
      </div>

      {/* Dual File Ingestion Zone */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(340px, 1fr))', gap: '20px', marginBottom: '24px' }}>
        
        {/* Card 1: Exam Schedule / Pre-Exam Allotments */}
        <div className="card" style={{ padding: '20px', border: '1.5px dashed var(--line)', background: 'var(--panel)', borderRadius: '12px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '12px' }}>
            <div style={{ background: 'rgba(23,107,135,0.1)', padding: '10px', borderRadius: '8px', color: 'var(--accent)' }}>
              <FileSpreadsheet size={20} />
            </div>
            <div>
              <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700 }}>Primary: Exam Schedule / Pre-Exam File</h3>
              <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Contains Date, Venue, Course Code & Name (.xlsx, .xls, .csv)</span>
            </div>
          </div>

          <label
            style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '20px',
              borderRadius: '8px',
              background: 'var(--bg)',
              border: '1px dashed var(--accent)',
              cursor: 'pointer',
              textAlign: 'center'
            }}
          >
            <Upload size={22} style={{ color: 'var(--accent)', marginBottom: '6px' }} />
            <span style={{ fontWeight: 600, fontSize: '13px' }}>
              {scheduleFile ? scheduleFile.name : 'Click to Upload Exam Schedule'}
            </span>
            <span style={{ fontSize: '11.5px', color: 'var(--muted)', marginTop: '2px' }}>
              {scheduleRows.length > 0 ? `${scheduleRows.length.toLocaleString()} rows loaded` : 'Supports QP Statements & Allotment Spreadsheets'}
            </span>
            <input type="file" accept=".xlsx,.xls,.csv,.zip" onChange={handleScheduleUpload} style={{ display: 'none' }} />
          </label>
        </div>

        {/* Card 2: Supplementary College Email Directory (Optional) */}
        <div className="card" style={{ padding: '20px', border: '1.5px dashed var(--line)', background: 'var(--panel)', borderRadius: '12px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '12px' }}>
            <div style={{ background: 'rgba(16,185,129,0.1)', padding: '10px', borderRadius: '8px', color: '#059669' }}>
              <Mail size={20} />
            </div>
            <div>
              <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700 }}>Optional: College Email Directory</h3>
              <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Override or supplement college email addresses (.xlsx, .csv)</span>
            </div>
          </div>

          <label
            style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '20px',
              borderRadius: '8px',
              background: 'var(--bg)',
              border: '1px dashed #059669',
              cursor: 'pointer',
              textAlign: 'center'
            }}
          >
            <Upload size={22} style={{ color: '#059669', marginBottom: '6px' }} />
            <span style={{ fontWeight: 600, fontSize: '13px' }}>
              {directoryFile ? directoryFile.name : 'Upload Custom Email Directory (Optional)'}
            </span>
            <span style={{ fontSize: '11.5px', color: 'var(--muted)', marginTop: '2px' }}>
              {directoryRows.length > 0 ? `${directoryRows.length} custom college emails active` : 'Preloaded with official Kannur University college emails'}
            </span>
            <input type="file" accept=".xlsx,.xls,.csv" onChange={handleDirectoryUpload} style={{ display: 'none' }} />
          </label>
        </div>

      </div>

      {/* Status Bar */}
      <div style={{ 
        padding: '10px 16px', 
        borderRadius: '8px', 
        fontSize: '13px', 
        marginBottom: '24px', 
        display: 'flex', 
        alignItems: 'center', 
        gap: '8px',
        background: statusType === 'error' ? 'rgba(239, 68, 68, 0.08)' : statusType === 'success' ? 'rgba(16, 185, 129, 0.08)' : 'rgba(23,107,135,0.06)',
        color: statusType === 'error' ? '#ef4444' : statusType === 'success' ? '#059669' : 'var(--ink)'
      }}>
        {statusType === 'error' ? <AlertCircle size={16} /> : <CheckCircle2 size={16} />}
        <span>{statusMsg}</span>
      </div>

      {/* Column Mapping Drawer / Accordion */}
      {showConfig && scheduleHeaders.length > 0 && (
        <div className="card" style={{ padding: '20px', marginBottom: '24px', background: 'var(--panel)' }}>
          <h3 style={{ fontSize: '15px', fontWeight: 700, margin: '0 0 14px 0', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Sliders size={16} color="var(--accent)" /> Column Detection & Mapping (Schedule File)
          </h3>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '14px' }}>
            <div>
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Exam Date</label>
              <select
                value={scheduleMapping.examDate}
                onChange={(e) => setScheduleMapping({ ...scheduleMapping, examDate: parseInt(e.target.value, 10) })}
                style={{ width: '100%', fontSize: '12.5px' }}
              >
                <option value={-1}>-- Not Mapped --</option>
                {scheduleHeaders.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Venue Code</label>
              <select
                value={scheduleMapping.venueCode}
                onChange={(e) => setScheduleMapping({ ...scheduleMapping, venueCode: parseInt(e.target.value, 10) })}
                style={{ width: '100%', fontSize: '12.5px' }}
              >
                <option value={-1}>-- Not Mapped --</option>
                {scheduleHeaders.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Venue Name</label>
              <select
                value={scheduleMapping.venueName}
                onChange={(e) => setScheduleMapping({ ...scheduleMapping, venueName: parseInt(e.target.value, 10) })}
                style={{ width: '100%', fontSize: '12.5px' }}
              >
                <option value={-1}>-- Not Mapped --</option>
                {scheduleHeaders.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Course Code</label>
              <select
                value={scheduleMapping.courseCode}
                onChange={(e) => setScheduleMapping({ ...scheduleMapping, courseCode: parseInt(e.target.value, 10) })}
                style={{ width: '100%', fontSize: '12.5px' }}
              >
                <option value={-1}>-- Not Mapped --</option>
                {scheduleHeaders.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Course Name</label>
              <select
                value={scheduleMapping.courseName}
                onChange={(e) => setScheduleMapping({ ...scheduleMapping, courseName: parseInt(e.target.value, 10) })}
                style={{ width: '100%', fontSize: '12.5px' }}
              >
                <option value={-1}>-- Not Mapped --</option>
                {scheduleHeaders.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Email (Optional from file)</label>
              <select
                value={scheduleMapping.email}
                onChange={(e) => setScheduleMapping({ ...scheduleMapping, email: parseInt(e.target.value, 10) })}
                style={{ width: '100%', fontSize: '12.5px' }}
              >
                <option value={-1}>-- Use Directory --</option>
                {scheduleHeaders.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Student Count</label>
              <select
                value={scheduleMapping.studentCount}
                onChange={(e) => setScheduleMapping({ ...scheduleMapping, studentCount: parseInt(e.target.value, 10) })}
                style={{ width: '100%', fontSize: '12.5px' }}
              >
                <option value={-1}>-- Count 1 per row --</option>
                {scheduleHeaders.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '4px' }}>Session / Timing</label>
              <select
                value={scheduleMapping.session}
                onChange={(e) => setScheduleMapping({ ...scheduleMapping, session: parseInt(e.target.value, 10) })}
                style={{ width: '100%', fontSize: '12.5px' }}
              >
                <option value={-1}>-- Not Mapped --</option>
                {scheduleHeaders.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>
          </div>
        </div>
      )}

      {/* Summary Statistics Cards */}
      {reportData.items.length > 0 && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(170px, 1fr))', gap: '14px', marginBottom: '24px' }}>
          <div className="card" style={{ padding: '16px', borderLeft: '4px solid var(--accent)' }}>
            <span style={{ fontSize: '11.5px', color: 'var(--muted)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>Unique Combinations</span>
            <strong style={{ fontSize: '22px', color: 'var(--ink)' }}>{reportData.stats.totalCombinations.toLocaleString()}</strong>
            <span style={{ fontSize: '11px', color: 'var(--muted)', display: 'block', marginTop: '2px' }}>Date + Venue + Course</span>
          </div>

          <div className="card" style={{ padding: '16px', borderLeft: '4px solid #8b5cf6' }}>
            <span style={{ fontSize: '11.5px', color: 'var(--muted)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>Examination Venues</span>
            <strong style={{ fontSize: '22px', color: '#7c3aed' }}>{reportData.stats.totalVenues.toLocaleString()}</strong>
            <span style={{ fontSize: '11px', color: 'var(--muted)', display: 'block', marginTop: '2px' }}>Colleges / Centres</span>
          </div>

          <div className="card" style={{ padding: '16px', borderLeft: '4px solid #059669' }}>
            <span style={{ fontSize: '11.5px', color: 'var(--muted)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>Distinct Exam Dates</span>
            <strong style={{ fontSize: '22px', color: '#059669' }}>{reportData.stats.totalDates.toLocaleString()}</strong>
            <span style={{ fontSize: '11px', color: 'var(--muted)', display: 'block', marginTop: '2px' }}>Exam sessions</span>
          </div>

          <div className="card" style={{ padding: '16px', borderLeft: '4px solid #f59e0b' }}>
            <span style={{ fontSize: '11.5px', color: 'var(--muted)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>Distinct Courses</span>
            <strong style={{ fontSize: '22px', color: '#d97706' }}>{reportData.stats.totalCourses.toLocaleString()}</strong>
            <span style={{ fontSize: '11px', color: 'var(--muted)', display: 'block', marginTop: '2px' }}>Papers / subjects</span>
          </div>

          <div className="card" style={{ padding: '16px', borderLeft: '4px solid #3b82f6' }}>
            <span style={{ fontSize: '11.5px', color: 'var(--muted)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>Total Candidates</span>
            <strong style={{ fontSize: '22px', color: '#2563eb' }}>{reportData.stats.totalStudents.toLocaleString()}</strong>
            <span style={{ fontSize: '11px', color: 'var(--muted)', display: 'block', marginTop: '2px' }}>Sum of registrants</span>
          </div>

          <div className="card" style={{ padding: '16px', borderLeft: '4px solid #ec4899' }}>
            <span style={{ fontSize: '11.5px', color: 'var(--muted)', display: 'block', textTransform: 'uppercase', fontWeight: 700 }}>Email Coverage</span>
            <strong style={{ fontSize: '22px', color: '#db2777' }}>
              {Math.round((reportData.stats.emailCount / (reportData.stats.totalCombinations || 1)) * 100)}%
            </strong>
            <span style={{ fontSize: '11px', color: 'var(--muted)', display: 'block', marginTop: '2px' }}>
              {reportData.stats.emailCount} of {reportData.stats.totalCombinations} with email
            </span>
          </div>
        </div>
      )}

      {/* Main Report Table & Actions Deck */}
      {reportData.items.length > 0 && (
        <div className="card" style={{ padding: 0, overflow: 'hidden', marginBottom: '32px' }}>
          
          {/* View Mode Tabs & Quick Copy Bar */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid var(--line)', padding: '0 16px', flexWrap: 'wrap', gap: '12px' }}>
            <div style={{ display: 'flex', gap: '4px' }}>
              <button
                type="button"
                onClick={() => setActiveTab('all')}
                style={{
                  padding: '12px 18px',
                  fontSize: '13px',
                  fontWeight: 700,
                  background: activeTab === 'all' ? 'var(--panel)' : 'transparent',
                  border: 'none',
                  borderBottom: activeTab === 'all' ? '2px solid var(--accent)' : '2px solid transparent',
                  color: activeTab === 'all' ? 'var(--accent)' : 'var(--muted)',
                  cursor: 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px'
                }}
              >
                <FileSpreadsheet size={15} /> All Combinations ({filteredItems.length})
              </button>

              <button
                type="button"
                onClick={() => setActiveTab('byVenue')}
                style={{
                  padding: '12px 18px',
                  fontSize: '13px',
                  fontWeight: 700,
                  background: activeTab === 'byVenue' ? 'var(--panel)' : 'transparent',
                  border: 'none',
                  borderBottom: activeTab === 'byVenue' ? '2px solid #8b5cf6' : '2px solid transparent',
                  color: activeTab === 'byVenue' ? '#7c3aed' : 'var(--muted)',
                  cursor: 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px'
                }}
              >
                <Building2 size={15} /> Grouped by College / Venue ({groupedByVenue.length})
              </button>

              <button
                type="button"
                onClick={() => setActiveTab('byDate')}
                style={{
                  padding: '12px 18px',
                  fontSize: '13px',
                  fontWeight: 700,
                  background: activeTab === 'byDate' ? 'var(--panel)' : 'transparent',
                  border: 'none',
                  borderBottom: activeTab === 'byDate' ? '2px solid #059669' : '2px solid transparent',
                  color: activeTab === 'byDate' ? '#059669' : 'var(--muted)',
                  cursor: 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px'
                }}
              >
                <Calendar size={15} /> Grouped by Exam Date ({groupedByDate.length})
              </button>
            </div>

            {/* Quick Actions (Copy Emails & Exports) */}
            <div style={{ display: 'flex', gap: '8px', padding: '8px 0', flexWrap: 'wrap' }}>
              <button
                type="button"
                onClick={handleCopyEmails}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '7px 12px',
                  fontSize: '12.5px',
                  fontWeight: 600,
                  background: copiedEmails ? 'rgba(16,185,129,0.1)' : 'rgba(23,107,135,0.08)',
                  color: copiedEmails ? '#059669' : 'var(--accent)',
                  border: '1px solid currentColor',
                  borderRadius: '6px',
                  cursor: 'pointer'
                }}
                title="Copy comma-separated unique email addresses for mail merge or bulk email client"
              >
                {copiedEmails ? <Check size={14} /> : <Copy size={14} />}
                {copiedEmails ? 'Emails Copied!' : 'Copy Venue Emails'}
              </button>

              <button
                type="button"
                onClick={handleExportExcel}
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '7px 14px',
                  fontSize: '12.5px',
                  fontWeight: 700,
                  background: '#059669',
                  color: 'white',
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer'
                }}
              >
                <Download size={14} /> Export Excel (.xlsx)
              </button>

              <button
                type="button"
                className="secondary"
                onClick={handleExportCsv}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '7px 12px', fontSize: '12.5px' }}
              >
                <FileText size={14} /> Export CSV
              </button>
            </div>
          </div>

          {/* Search & Dynamic Filter Controls */}
          <div style={{ padding: '14px 16px', background: 'var(--bg)', borderBottom: '1px solid var(--line)', display: 'flex', gap: '12px', alignItems: 'center', flexWrap: 'wrap' }}>
            <div style={{ position: 'relative', flex: 1, minWidth: '240px' }}>
              <Search size={15} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
              <input
                type="text"
                placeholder="Search across Date, Venue Code, College Name, Course, Email..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                style={{ width: '100%', paddingLeft: '32px', fontSize: '12.5px' }}
              />
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600 }}>Date:</span>
              <select
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                style={{ fontSize: '12.5px', padding: '5px 8px' }}
              >
                <option value="ALL">All Dates ({reportData.dates.length})</option>
                {reportData.dates.map(d => <option key={d} value={d}>{d}</option>)}
              </select>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600 }}>Venue:</span>
              <select
                value={selectedVenue}
                onChange={(e) => setSelectedVenue(e.target.value)}
                style={{ fontSize: '12.5px', padding: '5px 8px' }}
              >
                <option value="ALL">All Venues ({reportData.venues.length})</option>
                {reportData.venues.map(v => <option key={v} value={v}>Venue: {v}</option>)}
              </select>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
              <span style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600 }}>Email Status:</span>
              <select
                value={emailFilter}
                onChange={(e) => setEmailFilter(e.target.value)}
                style={{ fontSize: '12.5px', padding: '5px 8px' }}
              >
                <option value="ALL">All Records</option>
                <option value="HAS_EMAIL">With Valid Email ({reportData.stats.emailCount})</option>
                <option value="MISSING_EMAIL">Missing Email ({reportData.stats.totalCombinations - reportData.stats.emailCount})</option>
              </select>
            </div>
          </div>

          {/* TAB 1: ALL UNIQUE COMBINATIONS TABLE */}
          {activeTab === 'all' && (
            <div style={{ overflowX: 'auto', maxHeight: '600px' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                <thead>
                  <tr style={{ background: 'var(--panel)', borderBottom: '2px solid var(--line)', position: 'sticky', top: 0, zIndex: 1 }}>
                    <th style={{ padding: '10px 12px', width: '45px', textAlign: 'center' }}>#</th>
                    <th style={{ padding: '10px 12px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--accent)' }}>Exam Date</th>
                    <th style={{ padding: '10px 12px', textAlign: 'center', width: '90px', whiteSpace: 'nowrap' }}>Venue Code</th>
                    <th style={{ padding: '10px 12px', textAlign: 'left', minWidth: '180px' }}>Venue / College Name</th>
                    <th style={{ padding: '10px 12px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--accent)' }}>Course Code</th>
                    <th style={{ padding: '10px 12px', textAlign: 'left', minWidth: '190px' }}>Course Name</th>
                    <th style={{ padding: '10px 12px', textAlign: 'left', minWidth: '220px', color: '#059669' }}>College Email ID</th>
                    <th style={{ padding: '10px 12px', textAlign: 'center', width: '90px', whiteSpace: 'nowrap' }}>Student Count</th>
                    <th style={{ padding: '10px 12px', textAlign: 'left', width: '130px', whiteSpace: 'nowrap' }}>Session</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredItems.length > 0 ? (
                    filteredItems.map((item, idx) => {
                      const isEditing = editingEmailKey === `${item.venueCode}_${item.venueName}`;
                      return (
                        <tr 
                          key={idx}
                          style={{
                            borderBottom: '1px solid var(--line)',
                            background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.015)'
                          }}
                        >
                          <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 600, color: 'var(--muted)' }}>
                            {idx + 1}
                          </td>
                          <td style={{ padding: '8px 12px', whiteSpace: 'nowrap', fontWeight: 600, color: 'var(--ink)' }}>
                            <span style={{ display: 'inline-flex', alignItems: 'center', gap: '5px' }}>
                              <Calendar size={13} color="var(--accent)" />
                              {item.examDate}
                            </span>
                          </td>
                          <td style={{ padding: '8px 12px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                            <span style={{
                              fontFamily: 'monospace',
                              fontWeight: 700,
                              padding: '2px 7px',
                              borderRadius: '4px',
                              background: 'rgba(139,92,246,0.1)',
                              color: '#7c3aed',
                              fontSize: '11.5px'
                            }}>
                              {item.venueCode}
                            </span>
                          </td>
                          <td style={{ padding: '8px 12px', fontWeight: 600 }}>
                            {item.venueName}
                          </td>
                          <td style={{ padding: '8px 12px', whiteSpace: 'nowrap' }}>
                            <span style={{
                              fontFamily: 'monospace',
                              fontWeight: 700,
                              padding: '2px 7px',
                              borderRadius: '4px',
                              background: 'rgba(23,107,135,0.1)',
                              color: 'var(--accent)',
                              fontSize: '11.5px'
                            }}>
                              {item.courseCode}
                            </span>
                          </td>
                          <td style={{ padding: '8px 12px', color: 'var(--ink)' }}>
                            {item.courseName || '—'}
                          </td>

                          {/* Email Cell with Quick Inline Editing */}
                          <td style={{ padding: '8px 12px', whiteSpace: 'nowrap' }}>
                            {isEditing ? (
                              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                <input
                                  type="email"
                                  value={editingEmailVal}
                                  onChange={(e) => setEditingEmailVal(e.target.value)}
                                  placeholder="college@univ.edu"
                                  autoFocus
                                  style={{ padding: '4px 8px', fontSize: '11.5px', width: '180px' }}
                                />
                                <button
                                  type="button"
                                  onClick={() => handleSaveInlineEmail(item.venueCode, item.venueName, editingEmailVal)}
                                  style={{ padding: '4px 8px', fontSize: '11px', background: '#059669', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
                                >
                                  Save
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setEditingEmailKey(null)}
                                  className="secondary"
                                  style={{ padding: '4px 8px', fontSize: '11px' }}
                                >
                                  Cancel
                                </button>
                              </div>
                            ) : item.email && item.email.includes('@') ? (
                              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '6px' }}>
                                <a
                                  href={`mailto:${item.email}`}
                                  style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', color: '#059669', textDecoration: 'none', fontWeight: 600 }}
                                >
                                  <Mail size={13} />
                                  {item.email}
                                </a>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setEditingEmailKey(`${item.venueCode}_${item.venueName}`);
                                    setEditingEmailVal(item.email);
                                  }}
                                  style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--muted)', padding: '2px' }}
                                  title="Edit email address for this college"
                                >
                                  <Edit2 size={12} />
                                </button>
                              </div>
                            ) : (
                              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                                <span style={{ color: '#ef4444', fontStyle: 'italic', fontSize: '11.5px' }}>
                                  Missing Email
                                </span>
                                <button
                                  type="button"
                                  onClick={() => {
                                    setEditingEmailKey(`${item.venueCode}_${item.venueName}`);
                                    setEditingEmailVal('');
                                  }}
                                  style={{ padding: '2px 6px', fontSize: '10.5px', background: 'rgba(239,68,68,0.1)', color: '#ef4444', border: '1px solid #ef4444', borderRadius: '4px', cursor: 'pointer' }}
                                >
                                  + Add Email
                                </button>
                              </div>
                            )}
                          </td>

                          <td style={{ padding: '8px 12px', textAlign: 'center', whiteSpace: 'nowrap' }}>
                            <span style={{
                              display: 'inline-flex',
                              alignItems: 'center',
                              gap: '4px',
                              fontFamily: 'monospace',
                              fontWeight: 700,
                              padding: '2px 8px',
                              borderRadius: '10px',
                              background: 'rgba(59,130,246,0.1)',
                              color: '#2563eb',
                              fontSize: '11.5px'
                            }}>
                              <Users size={12} />
                              {item.studentCount.toLocaleString()}
                            </span>
                          </td>

                          <td style={{ padding: '8px 12px', whiteSpace: 'nowrap', fontSize: '11.5px', color: 'var(--muted)' }}>
                            {item.session || '—'}
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={9} style={{ textAlign: 'center', padding: '40px', color: 'var(--muted)' }}>
                        No records match the active search or filters.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          )}

          {/* TAB 2: GROUPED BY COLLEGE / VENUE */}
          {activeTab === 'byVenue' && (
            <div style={{ padding: '16px', display: 'flex', flexDirection: 'column', gap: '16px', maxHeight: '600px', overflowY: 'auto' }}>
              {groupedByVenue.map((venue, vIdx) => (
                <div key={vIdx} style={{ border: '1px solid var(--line)', borderRadius: '8px', background: 'var(--panel)', padding: '16px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px', marginBottom: '12px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <span style={{
                        fontFamily: 'monospace',
                        fontWeight: 700,
                        padding: '3px 8px',
                        borderRadius: '4px',
                        background: 'rgba(139,92,246,0.1)',
                        color: '#7c3aed',
                        fontSize: '12px'
                      }}>
                        {venue.venueCode}
                      </span>
                      <h4 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: 'var(--ink)' }}>{venue.venueName}</h4>
                    </div>

                    <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                      {venue.email ? (
                        <a href={`mailto:${venue.email}`} style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', color: '#059669', fontSize: '13px', textDecoration: 'none', fontWeight: 600 }}>
                          <Mail size={14} /> {venue.email}
                        </a>
                      ) : (
                        <span style={{ color: '#ef4444', fontSize: '12px', fontStyle: 'italic' }}>Missing Email</span>
                      )}

                      <span style={{ fontSize: '12px', fontWeight: 700, padding: '2px 8px', borderRadius: '12px', background: 'rgba(59,130,246,0.1)', color: '#2563eb' }}>
                        {venue.totalStudents.toLocaleString()} Candidates
                      </span>
                    </div>
                  </div>

                  {/* Scheduled Courses at this venue */}
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                    <thead>
                      <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)' }}>
                        <th style={{ padding: '6px 10px', textAlign: 'left', width: '110px' }}>Exam Date</th>
                        <th style={{ padding: '6px 10px', textAlign: 'left', width: '130px' }}>Course Code</th>
                        <th style={{ padding: '6px 10px', textAlign: 'left' }}>Course Name</th>
                        <th style={{ padding: '6px 10px', textAlign: 'center', width: '80px' }}>Count</th>
                        <th style={{ padding: '6px 10px', textAlign: 'left', width: '140px' }}>Timing</th>
                      </tr>
                    </thead>
                    <tbody>
                      {venue.offerings.map((off, oIdx) => (
                        <tr key={oIdx} style={{ borderBottom: '1px solid var(--line)' }}>
                          <td style={{ padding: '6px 10px', fontWeight: 600 }}>{off.examDate}</td>
                          <td style={{ padding: '6px 10px', fontFamily: 'monospace', color: 'var(--accent)', fontWeight: 700 }}>{off.courseCode}</td>
                          <td style={{ padding: '6px 10px' }}>{off.courseName}</td>
                          <td style={{ padding: '6px 10px', textAlign: 'center', fontWeight: 600 }}>{off.studentCount}</td>
                          <td style={{ padding: '6px 10px', color: 'var(--muted)' }}>{off.session || '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ))}
            </div>
          )}

          {/* TAB 3: GROUPED BY EXAM DATE */}
          {activeTab === 'byDate' && (
            <div style={{ padding: '16px', display: 'flex', flexDirection: 'column', gap: '16px', maxHeight: '600px', overflowY: 'auto' }}>
              {groupedByDate.map((dateGrp, dIdx) => (
                <div key={dIdx} style={{ border: '1px solid var(--line)', borderRadius: '8px', background: 'var(--panel)', padding: '16px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      <Calendar size={18} color="var(--accent)" />
                      <h4 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>{dateGrp.examDate}</h4>
                      <span style={{ fontSize: '12px', color: 'var(--muted)' }}>({dateGrp.offerings.length} course slots)</span>
                    </div>
                    <span style={{ fontSize: '12px', fontWeight: 700, padding: '2px 8px', borderRadius: '12px', background: 'rgba(16,185,129,0.1)', color: '#059669' }}>
                      {dateGrp.totalStudents.toLocaleString()} Candidates Scheduled
                    </span>
                  </div>

                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                    <thead>
                      <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)' }}>
                        <th style={{ padding: '6px 10px', textAlign: 'center', width: '80px' }}>Venue</th>
                        <th style={{ padding: '6px 10px', textAlign: 'left', minWidth: '180px' }}>College / Centre</th>
                        <th style={{ padding: '6px 10px', textAlign: 'left', minWidth: '180px' }}>College Email</th>
                        <th style={{ padding: '6px 10px', textAlign: 'left', width: '120px' }}>Course Code</th>
                        <th style={{ padding: '6px 10px', textAlign: 'left' }}>Course Name</th>
                        <th style={{ padding: '6px 10px', textAlign: 'center', width: '70px' }}>Count</th>
                      </tr>
                    </thead>
                    <tbody>
                      {dateGrp.offerings.map((off, oIdx) => (
                        <tr key={oIdx} style={{ borderBottom: '1px solid var(--line)' }}>
                          <td style={{ padding: '6px 10px', textAlign: 'center', fontFamily: 'monospace', fontWeight: 700 }}>{off.venueCode}</td>
                          <td style={{ padding: '6px 10px', fontWeight: 600 }}>{off.venueName}</td>
                          <td style={{ padding: '6px 10px', color: '#059669' }}>{off.email || '—'}</td>
                          <td style={{ padding: '6px 10px', fontFamily: 'monospace', color: 'var(--accent)', fontWeight: 700 }}>{off.courseCode}</td>
                          <td style={{ padding: '6px 10px' }}>{off.courseName}</td>
                          <td style={{ padding: '6px 10px', textAlign: 'center', fontWeight: 600 }}>{off.studentCount}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ))}
            </div>
          )}

          {/* Footer Bar */}
          <div style={{ padding: '12px 16px', background: 'var(--bg)', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '12.5px', color: 'var(--muted)' }}>
            <span>Showing <strong>{filteredItems.length.toLocaleString()}</strong> unique Date-Venue-Course allocation rows</span>
            <span>All records verified & formatted strictly for university dispatch</span>
          </div>

        </div>
      )}

      {/* College Email Directory Modal */}
      {showEmailModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          padding: '20px'
        }}>
          <div style={{
            background: 'var(--panel)',
            borderRadius: '12px',
            width: '100%',
            maxWidth: '820px',
            maxHeight: '90vh',
            display: 'flex',
            flexDirection: 'column',
            boxShadow: '0 20px 40px rgba(0,0,0,0.2)'
          }}>
            <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Mail size={18} color="var(--accent)" />
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>College Email Directory Manager</h3>
              </div>
              <button
                type="button"
                onClick={() => setShowEmailModal(false)}
                style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontSize: '16px', color: 'var(--muted)' }}
              >
                ✕
              </button>
            </div>

            <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--line)', background: 'var(--bg)' }}>
              <h4 style={{ margin: '0 0 10px 0', fontSize: '13px', fontWeight: 700 }}>+ Add or Update College Email</h4>
              <div style={{ display: 'grid', gridTemplateColumns: '120px 1fr 1fr auto', gap: '10px' }}>
                <input
                  type="text"
                  placeholder="Code (e.g. 201)"
                  value={newDirCode}
                  onChange={(e) => setNewDirCode(e.target.value)}
                  style={{ fontSize: '12.5px', padding: '6px 8px' }}
                />
                <input
                  type="text"
                  placeholder="College / Venue Name"
                  value={newDirName}
                  onChange={(e) => setNewDirName(e.target.value)}
                  style={{ fontSize: '12.5px', padding: '6px 8px' }}
                />
                <input
                  type="email"
                  placeholder="Official Email ID"
                  value={newDirEmail}
                  onChange={(e) => setNewDirEmail(e.target.value)}
                  style={{ fontSize: '12.5px', padding: '6px 8px' }}
                />
                <button
                  type="button"
                  onClick={() => {
                    if (!newDirEmail || !newDirEmail.includes('@')) {
                      alert('Please provide a valid email address.');
                      return;
                    }
                    const updated = { ...collegeEmailMap };
                    if (newDirCode) updated[newDirCode.toLowerCase()] = { code: newDirCode, name: newDirName, email: newDirEmail.trim() };
                    if (newDirName) updated[newDirName.toLowerCase()] = { code: newDirCode, name: newDirName, email: newDirEmail.trim() };
                    saveEmailDirectory(updated);
                    setNewDirCode('');
                    setNewDirName('');
                    setNewDirEmail('');
                  }}
                  style={{ padding: '6px 14px', background: 'var(--accent)', color: 'white', border: 'none', borderRadius: '6px', fontWeight: 600, fontSize: '12.5px', cursor: 'pointer' }}
                >
                  Save Entry
                </button>
              </div>
            </div>

            <div style={{ overflowY: 'auto', flex: 1, padding: '16px 20px' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)' }}>
                    <th style={{ padding: '8px', textAlign: 'center', width: '80px' }}>Code</th>
                    <th style={{ padding: '8px', textAlign: 'left' }}>College / Centre Name</th>
                    <th style={{ padding: '8px', textAlign: 'left', minWidth: '220px' }}>Official Email Address</th>
                    <th style={{ padding: '8px', textAlign: 'center', width: '60px' }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {DEFAULT_COLLEGE_EMAILS.map((c, idx) => {
                    const currentEmail = collegeEmailMap[c.code.toLowerCase()]?.email || c.email;
                    return (
                      <tr key={idx} style={{ borderBottom: '1px solid var(--line)' }}>
                        <td style={{ padding: '8px', textAlign: 'center', fontFamily: 'monospace', fontWeight: 700 }}>{c.code}</td>
                        <td style={{ padding: '8px', fontWeight: 600 }}>{c.name}</td>
                        <td style={{ padding: '8px', color: '#059669', fontWeight: 600 }}>{currentEmail}</td>
                        <td style={{ padding: '8px', textAlign: 'center' }}>
                          <button
                            type="button"
                            onClick={() => {
                              setNewDirCode(c.code);
                              setNewDirName(c.name);
                              setNewDirEmail(currentEmail);
                            }}
                            style={{ background: 'transparent', border: 'none', cursor: 'pointer', color: 'var(--accent)' }}
                            title="Edit this college"
                          >
                            <Edit2 size={13} />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div style={{ padding: '12px 20px', borderTop: '1px solid var(--line)', background: 'var(--bg)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ fontSize: '12px', color: 'var(--muted)' }}>All changes are persisted to your browser session automatically.</span>
              <button
                type="button"
                className="secondary"
                onClick={() => setShowEmailModal(false)}
                style={{ padding: '6px 14px', fontSize: '12.5px' }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Logic Guide Modal */}
      {showGuideModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0,0,0,0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          padding: '20px'
        }}>
          <div style={{
            background: 'var(--panel)',
            borderRadius: '12px',
            width: '100%',
            maxWidth: '680px',
            maxHeight: '85vh',
            overflowY: 'auto',
            padding: '24px',
            boxShadow: '0 20px 40px rgba(0,0,0,0.2)'
          }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <HelpCircle size={20} color="var(--accent)" />
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 800 }}>Venue Date Course College Email Guide</h3>
              </div>
              <button
                type="button"
                onClick={() => setShowGuideModal(false)}
                style={{ background: 'transparent', border: 'none', cursor: 'pointer', fontSize: '18px', color: 'var(--muted)' }}
              >
                ✕
              </button>
            </div>

            <div style={{ fontSize: '13px', lineHeight: 1.6, display: 'flex', flexDirection: 'column', gap: '14px', color: 'var(--ink)' }}>
              <div>
                <h4 style={{ margin: '0 0 4px 0', color: 'var(--accent)', fontSize: '14px' }}>1. Purpose of the Tool</h4>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  During university examination administration, question paper branch officers and controllers must coordinate with venues and colleges. This tool creates an unduplicated, structured operational report showing each unique combination of <strong>Exam Date</strong>, <strong>Venue</strong>, <strong>Course Code</strong>, and <strong>Course Name</strong>, enriched with the official <strong>College Email ID</strong>.
                </p>
              </div>

              <div>
                <h4 style={{ margin: '0 0 4px 0', color: 'var(--accent)', fontSize: '14px' }}>2. Deduplication & Unique Combinations</h4>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  Student registrations or slot-level records are grouped strictly by <code>Date</code> + <code>Venue Code</code> + <code>Course Code</code>. If multiple candidates or classes take the same examination at the same venue on the same date, they are combined into a single operational row with their student count summed up.
                </p>
              </div>

              <div>
                <h4 style={{ margin: '0 0 4px 0', color: 'var(--accent)', fontSize: '14px' }}>3. How College Email IDs are Resolved</h4>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  The tool resolves email addresses using a 3-tier cascade:
                </p>
                <ol style={{ margin: '6px 0 0 18px', padding: 0, color: 'var(--muted)' }}>
                  <li><strong>In-File Column:</strong> If the uploaded schedule already contains an email column, it uses that email directly.</li>
                  <li><strong>Custom Email Directory:</strong> Matches against any uploaded or edited college email directory by Venue Code or College Name.</li>
                  <li><strong>University Defaults:</strong> Automatically falls back to official Kannur University affiliated colleges and teaching departments directory.</li>
                </ol>
              </div>

              <div>
                <h4 style={{ margin: '0 0 4px 0', color: 'var(--accent)', fontSize: '14px' }}>4. One-Click Bulk Email Dispatch</h4>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  Use the <strong>"Copy Venue Emails"</strong> button to copy all unique recipient emails formatted as a comma-separated list, ready to paste directly into your university email client (Gmail / Outlook Bcc) for instant exam circulars.
                </p>
              </div>
            </div>

            <div style={{ marginTop: '20px', textAlign: 'right' }}>
              <button
                type="button"
                className="primary"
                onClick={() => setShowGuideModal(false)}
                style={{ padding: '8px 18px', fontSize: '13px' }}
              >
                Got It
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};

export default VenueDateCourseEmailPage;
