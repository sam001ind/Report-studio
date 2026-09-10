import React, { useState, useMemo, useRef } from "react";
import { Link } from "react-router-dom";
import * as XLSX from "xlsx";
import { 
  ArrowLeft, 
  Upload, 
  Download, 
  FileSpreadsheet, 
  CheckCircle2, 
  XCircle, 
  AlertTriangle, 
  Search, 
  Filter, 
  RefreshCw, 
  Table, 
  Sparkles, 
  HelpCircle, 
  Layers, 
  GraduationCap, 
  Calculator, 
  BookOpen, 
  ArrowUp, 
  ArrowDown, 
  ArrowUpDown, 
  X, 
  Percent, 
  Award,
  UserCheck,
  UserX,
  Sliders,
  FileUp,
  FileDown,
  Zap,
  Check,
  ToggleLeft,
  ToggleRight,
  TrendingUp,
  BarChart3,
  Users,
  ChevronDown,
  ChevronRight,
  ShieldAlert,
  AlertCircle,
  Lock
} from "lucide-react";

export const ADES_OUTPUT_HEADERS = [
  "Faculty",
  "Program Term Name",
  "Course Code",
  "Course Name",
  "Seat Number",
  "PRN",
  "ESE - PR Max",
  "ESE - PR Min",
  "ESE - PR Obtained",
  "ESE - TH Max",
  "ESE - TH Min",
  "ESE - TH Obtained",
  "ESE - Max",
  "ESE - Min",
  "ESE Overall",
  "CE - PR Max",
  "CE - PR Min",
  "CE - PR Obtained",
  "CE - TH Max",
  "CE - TH Min",
  "CE - TH Obtained",
  "CE - Max",
  "CE - Min",
  "CE Overall Marks ",
  "Overall Maximum",
  "Overall Minimum",
  "Course Overall Marks ",
  "ESE Pass",
  "Overall pass",
  "Course Pass/Fail",
  "Moderation Marks"
];

const normalizeKey = (key) => String(key || "").toLowerCase().replace(/[^a-z0-9]/g, "");

const cleanString = (str) => {
  if (!str) return "";
  return String(str).replace(/\s+/g, " ").trim();
};

export const getCell = (row, hMap = {}, ...aliases) => {
  if (!row || !hMap) return "";
  for (const alias of aliases) {
    const norm = normalizeKey(alias);
    const actualKey = hMap[norm];
    if (actualKey && row[actualKey] !== undefined && row[actualKey] !== null) {
      return row[actualKey];
    }
  }
  return "";
};

export const parseNumber = (val) => {
  if (val === null || val === undefined) return null;
  const s = String(val).trim();
  if (s === "") return null;
  const lower = s.toLowerCase();
  if (lower.includes("absent") || lower.includes("(ab)")) return null;
  if (lower.includes("malpractice") || lower.includes("(mp)") || lower.includes("smp") || lower.includes("ehb")) return null;
  const num = Number(s);
  return isNaN(num) ? null : num;
};

// Cleans and canonicalizes course codes: strips ., #, *, _, -, spaces, wrapping brackets, and converts to uppercase
export const cleanCourseCode = (rawCode) => {
  if (!rawCode) return "";
  let code = String(rawCode).trim();
  code = code.replace(/^\s*[\(\[\{]\s*/, "").replace(/\s*[\)\]\}]\s*$/, "");
  code = code.replace(/[\.\#\*\_\-\:\;\,]+/g, " ");
  code = code.replace(/\s+/g, "");
  code = code.replace(/^[^a-zA-Z0-9]+|[^a-zA-Z0-9]+$/g, "");
  return code.toUpperCase();
};

// Cleans course title removing trailing special markers like # or *
export const cleanCourseName = (rawName) => {
  if (!rawName) return "";
  let name = String(rawName).trim();
  name = name.replace(/[\#\*\_\-]+$/g, "").trim();
  return name;
};

// Cleans college name: strips leading code prefixes e.g. "[101] - Name", "101 - Name", "101. Name", "101: Name"
export const cleanCollegeName = (rawName) => {
  if (!rawName) return "";
  let name = cleanString(rawName);
  // Pattern 1: "[101] - Govt College" or "101 - Govt College" or "101: Govt College" or "101. Govt College"
  const prefixMatch = name.match(/^\[?[A-Za-z0-9_]+\]?\s*[-:–—.]\s*(.+)$/);
  if (prefixMatch) {
    name = prefixMatch[1].trim();
  } else {
    // Pattern 2: "101 Govt College"
    const numStartMatch = name.match(/^\[?\d{2,}\]?\s*[-:–—.]?\s*([A-Za-z].+)$/);
    if (numStartMatch) {
      name = numStartMatch[1].trim();
    }
  }
  return name;
};

// Parses raw college code and name
export const parseCollegeRaw = (rawCodeInput, rawNameInput) => {
  let code = cleanString(rawCodeInput);
  let name = cleanCollegeName(rawNameInput);

  if (!code && rawNameInput) {
    const rawN = cleanString(rawNameInput);
    const prefixMatch = rawN.match(/^\[?([A-Za-z0-9_]+)\]?\s*[-:–—.]\s*(.+)$/);
    if (prefixMatch) {
      code = prefixMatch[1].trim();
    } else {
      const suffixMatch = rawN.match(/^(.+?)\s*[\(\[]\s*([A-Za-z0-9_]+)\s*[\)\]]$/);
      if (suffixMatch && !suffixMatch[1].toLowerCase().includes("autonomous") && !suffixMatch[1].toLowerCase().includes("aided")) {
        code = suffixMatch[2].trim();
      }
    }
  }

  return { code, name };
};

// Builds a dataset-wide canonical resolver using Seat Number and ADEC Name connection
export const buildCollegeCanonicalRegistry = (rows = [], currentHeaderMap = {}, currentAbsentMap = null, currentMalpracticeMap = null, currentHeldbackMap = null) => {
  const seatToCollege = new Map();
  const prnToCollege = new Map();
  const codeToNames = new Map();
  const nameToCodes = new Map();
  const allCollegeNames = new Set();

  const registerEntry = (seat, prn, rawCode, rawName) => {
    const { code, name } = parseCollegeRaw(rawCode, rawName);
    const cleanSeat = cleanString(seat);
    const cleanPrn = cleanString(prn);
    const normC = normalizeKey(code);
    const normN = normalizeKey(name);

    if (name) {
      allCollegeNames.add(name);
      if (normC) {
        if (!codeToNames.has(normC)) codeToNames.set(normC, new Set());
        codeToNames.get(normC).add(name);
      }
    }
    if (code && normN) {
      if (!nameToCodes.has(normN)) nameToCodes.set(normN, new Set());
      nameToCodes.get(normN).add(code);
    }

    if (cleanSeat && (name || code)) {
      if (!seatToCollege.has(cleanSeat)) {
        seatToCollege.set(cleanSeat, { code, name: name || code });
      } else {
        const cur = seatToCollege.get(cleanSeat);
        if (!cur.name && name) cur.name = name;
        if (!cur.code && code) cur.code = code;
        if (name && cur.name && name.length > cur.name.length) cur.name = name;
      }
    }

    if (cleanPrn && (name || code)) {
      if (!prnToCollege.has(cleanPrn)) {
        prnToCollege.set(cleanPrn, { code, name: name || code });
      } else {
        const cur = prnToCollege.get(cleanPrn);
        if (!cur.name && name) cur.name = name;
        if (!cur.code && code) cur.code = code;
        if (name && cur.name && name.length > cur.name.length) cur.name = name;
      }
    }
  };

  // 1. Scan primary rows
  if (Array.isArray(rows)) {
    rows.forEach(row => {
      const seat = getCell(row, currentHeaderMap, "Seat Number", "SeatNumber", "SeatNo", "Seat_Number", "RollNo", "Roll Number");
      const prn = getCell(row, currentHeaderMap, "PRN", "PRN Number", "PRNNo", "RegisterNo", "RegNo", "StudentID");
      const cCode = getCell(row, currentHeaderMap, "ADEC Code", "ADECCode", "ADEC_Code", "ADEC", "College Code", "CollegeCode", "College_Code", "Center Code", "CenterCode", "InstCode");
      const cName = getCell(row, currentHeaderMap, "ADEC Name", "ADECName", "ADEC_Name", "ADEC", "College Name", "CollegeName", "College_Name", "College", "Center Name", "CenterName", "Institute", "Institute Name", "College / Department");
      registerEntry(seat, prn, cCode, cName);
    });
  }

  // 2. Scan external report entries
  [currentHeldbackMap, currentMalpracticeMap, currentAbsentMap].forEach(m => {
    if (m) {
      m.forEach(e => {
        registerEntry(e.seat, e.prn, e.collegeCode, e.collegeName);
      });
    }
  });

  const canonicalCodeMap = new Map();
  codeToNames.forEach((nameSet, normC) => {
    const names = Array.from(nameSet);
    names.sort((a, b) => {
      const aIsUpper = a === a.toUpperCase();
      const bIsUpper = b === b.toUpperCase();
      if (aIsUpper !== bIsUpper) return aIsUpper ? 1 : -1;
      return b.length - a.length;
    });
    canonicalCodeMap.set(normC, names[0]);
  });

  const resolve = (seatInput, prnInput, rawCodeInput, rawNameInput) => {
    const cleanSeat = cleanString(seatInput);
    const cleanPrn = cleanString(prnInput);

    // 1. Connection via Seat Number (ADEC Name & Seat combination)
    if (cleanSeat && seatToCollege.has(cleanSeat)) {
      const entry = seatToCollege.get(cleanSeat);
      if (entry.name) {
        return {
          collegeCode: entry.code || "",
          collegeName: entry.name,
          college: entry.name
        };
      }
    }

    // 2. Connection via PRN
    if (cleanPrn && prnToCollege.has(cleanPrn)) {
      const entry = prnToCollege.get(cleanPrn);
      if (entry.name) {
        return {
          collegeCode: entry.code || "",
          collegeName: entry.name,
          college: entry.name
        };
      }
    }

    // 3. Fallback to direct raw code/name resolution
    const { code, name } = parseCollegeRaw(rawCodeInput, rawNameInput);
    const normC = normalizeKey(code);
    let finalName = name || (normC ? canonicalCodeMap.get(normC) || "" : "") || code || "";

    if (normC && canonicalCodeMap.has(normC)) {
      finalName = canonicalCodeMap.get(normC);
    }

    return {
      collegeCode: code || "",
      collegeName: finalName,
      college: finalName
    };
  };

  return { resolve, allColleges: Array.from(allCollegeNames).sort() };
};

export default function AdesResultCalculatorPage() {
  const [sourceFile, setSourceFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [workbook, setWorkbook] = useState(null);
  const [rawRows, setRawRows] = useState([]);
  const [groupedRecords, setGroupedRecords] = useState([]);
  const [headerMap, setHeaderMap] = useState({});
  const [statusMsg, setStatusMsg] = useState("Ready");
  const [statusType, setStatusType] = useState("info");
  const [isProcessing, setIsProcessing] = useState(false);
  
  // Moderation, Simulation & Student Summary State
  const [courseModerationMap, setCourseModerationMap] = useState({});
  const [allowPrOnlyModeration, setAllowPrOnlyModeration] = useState(false); // Moderation on solely ESE-PR courses toggle (default: false)
  const [activeTab, setActiveTab] = useState("results"); // "results" | "students" | "moderation" | "simulation"
  const [moderationSearch, setModerationSearch] = useState("");
  const [modSortConfig, setModSortConfig] = useState({ column: "rawFailed", direction: "desc" });
  const [modFilterType, setModFilterType] = useState("ALL"); // "ALL" | "TH" | "PR_ONLY"
  const [modFilterStatus, setModFilterStatus] = useState("ALL"); // "ALL" | "FAILED" | "NEAR_PASS" | "ACTIVE_MOD" | "RESCUED"
  const [simSearchQuery, setSimSearchQuery] = useState("");
  const [studentSearchQuery, setStudentSearchQuery] = useState("");
  const [studentFilterStatus, setStudentFilterStatus] = useState("ALL"); // "ALL" | "PASS" | "FAIL" | "RESCUED"
  const [studentProgramFilter, setStudentProgramFilter] = useState("ALL");
  const [studentCourseFilter, setStudentCourseFilter] = useState("ALL");
  const [studentCollegeFilter, setStudentCollegeFilter] = useState("ALL");
  const [expandedStudents, setExpandedStudents] = useState({});
  const [bulkModValue, setBulkModValue] = useState(4);
  const modFileInputRef = useRef(null);

  // Absent Students Report State
  const [absentRecordsMap, setAbsentRecordsMap] = useState(new Map());
  const [absentFileName, setAbsentFileName] = useState("");
  const [absentList, setAbsentList] = useState([]);
  const absentFileInputRef = useRef(null);

  // Malpractice Students Report State
  const [malpracticeRecordsMap, setMalpracticeRecordsMap] = useState(new Map());
  const [malpracticeFileName, setMalpracticeFileName] = useState("");
  const [malpracticeList, setMalpracticeList] = useState([]);
  const malpracticeFileInputRef = useRef(null);

  // Heldback Students Report State (Term-level & Course-level)
  const [heldbackRecordsMap, setHeldbackRecordsMap] = useState(new Map());
  const [heldbackFileName, setHeldbackFileName] = useState("");
  const [heldbackList, setHeldbackList] = useState([]);
  const heldbackFileInputRef = useRef(null);

  // Table Filters & Pagination
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedFacultyFilter, setSelectedFacultyFilter] = useState("ALL");
  const [selectedCollegeFilter, setSelectedCollegeFilter] = useState("ALL");
  const [selectedProgramFilter, setSelectedProgramFilter] = useState("ALL");
  const [selectedCourseFilter, setSelectedCourseFilter] = useState("ALL");
  const [selectedResultFilter, setSelectedResultFilter] = useState("ALL");
  const [columnFilters, setColumnFilters] = useState({});
  const [sortConfig, setSortConfig] = useState({ column: null, direction: null });
  const [page, setPage] = useState(0);
  const pageSize = 50;
  const [showHelpModal, setShowHelpModal] = useState(false);
  const [sheetMetadata, setSheetMetadata] = useState({});

  const setStatus = (msg, type = "info") => {
    setStatusMsg(msg);
    setStatusType(type);
  };


  // Helper to match student course record against loaded Absent Report
  const getAbsentEntry = (prn, seat, code, absentMap) => {
    if (!absentMap || absentMap.size === 0) return null;
    const normCode = normalizeKey(code);
    if (!normCode) return null;

    const normPrn = normalizeKey(prn);
    if (normPrn) {
      const entry = absentMap.get(`${normPrn}___${normCode}`);
      if (entry) return entry;
    }

    const normSeat = normalizeKey(seat);
    if (normSeat) {
      const entry = absentMap.get(`${normSeat}___${normCode}`);
      if (entry) return entry;
    }

    return null;
  };

  // Build lookup map and formatted list from uploaded Absent Report rows
  const buildAbsentLookup = (rows, headerMap = {}) => {
    const map = new Map();
    const list = [];

    rows.forEach(r => {
      const prn = String(getCell(r, headerMap, "PRN", "PRN Number", "PRNNo", "RegisterNo", "RegNo", "StudentID") || r["PRN"] || "").trim();
      const seat = String(getCell(r, headerMap, "Seat Number", "SeatNumber", "SeatNo", "Seat_Number", "RollNo", "Roll Number") || r["Seat Number"] || "").trim();
      const rawCode = String(getCell(r, headerMap, "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || r["Course Code"] || "").trim();
      const code = cleanCourseCode(rawCode);
      const courseName = String(getCell(r, headerMap, "Course Name", "CourseName", "PaperName", "SubjectName") || r["Course Name"] || "").trim();
      const studentName = String(getCell(r, headerMap, "Student Name", "StudentName", "Name", "CandidateName") || r["Student Name"] || "").trim();
      const rawCollegeCode = String(getCell(r, headerMap, "ADEC Code", "ADECCode", "ADEC_Code", "ADEC", "College Code", "CollegeCode", "College_Code") || r["College Code"] || r["ADEC Code"] || "").trim();
      const rawCollegeName = String(getCell(r, headerMap, "ADEC Name", "ADECName", "ADEC_Name", "ADEC", "College Name", "CollegeName", "College_Name", "College", "Center Name", "Institute") || r["College Name"] || r["ADEC Name"] || "").trim();
      const { code: collegeCode, name: collegeName } = parseCollegeRaw(rawCollegeCode, rawCollegeName);

      const am = String(getCell(r, headerMap, "AM", "Assessment Method", "AssessmentMethod") || r["AM"] || "").trim().toUpperCase();
      const at = String(getCell(r, headerMap, "AT", "Assessment Type", "AssessmentType") || r["AT"] || "").trim().toUpperCase();

      const status = String(getCell(r, headerMap, "Absent Status (categorized as)", "Absent Status", "AbsentStatus", "Status", "Remarks") || r["Absent Status (categorized as)"] || "").trim() || "Marked Absent During Mark Entry";

      const normCode = normalizeKey(code);
      const normPrn = normalizeKey(prn);
      const normSeat = normalizeKey(seat);

      const isEseTh = (am.includes("ESE") || !am) && (at.includes("TH") || !at);
      const isEsePr = am.includes("ESE") && at.includes("PR");
      const isCeTh = am.includes("CE") && at.includes("TH");
      const isCePr = am.includes("CE") && at.includes("PR");
      const isWholeEse = am.includes("ESE") && !at;
      const isWholeCourse = !am && !at;

      const details = {
        studentName,
        prn,
        seat,
        code,
        courseName,
        collegeCode,
        collegeName,
        am: am || "ESE",
        at: at || "TH",
        status,
        isEseTh: isEseTh || isWholeEse || isWholeCourse,
        isEsePr: isEsePr || isWholeEse || isWholeCourse,
        isCeTh: isCeTh || isWholeCourse,
        isCePr: isCePr || isWholeCourse,
        isAbsentOverall: isEseTh || isEsePr || isWholeEse || isWholeCourse
      };

      if (normPrn && normCode) {
        map.set(`${normPrn}___${normCode}`, details);
      }
      if (normSeat && normCode) {
        map.set(`${normSeat}___${normCode}`, details);
      }

      if (normCode && (normPrn || normSeat)) {
        list.push(details);
      }
    });

    return { map, list };
  };

  // Helper to match student course record against loaded Malpractice Report
  const getMalpracticeEntry = (prn, seat, code, malpracticeMap) => {
    if (!malpracticeMap || malpracticeMap.size === 0) return null;
    const normCode = normalizeKey(code);
    if (!normCode) return null;

    const normPrn = normalizeKey(prn);
    if (normPrn) {
      const entry = malpracticeMap.get(`${normPrn}___${normCode}`);
      if (entry) return entry;
    }

    const normSeat = normalizeKey(seat);
    if (normSeat) {
      const entry = malpracticeMap.get(`${normSeat}___${normCode}`);
      if (entry) return entry;
    }

    return null;
  };

  // Build lookup map and formatted list from uploaded Malpractice Report rows
  const buildMalpracticeLookup = (rows, headerMap = {}) => {
    const map = new Map();
    const list = [];

    rows.forEach(r => {
      const prn = String(getCell(r, headerMap, "PRN", "PRN Number", "PRNNo", "RegisterNo", "RegNo", "StudentID") || r["PRN"] || "").trim();
      const seat = String(getCell(r, headerMap, "Seat Number", "SeatNumber", "SeatNo", "Seat_Number", "RollNo", "Roll Number") || r["Seat Number"] || "").trim();
      const rawCode = String(getCell(r, headerMap, "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || r["Course Code"] || "").trim();
      const code = cleanCourseCode(rawCode);
      const courseName = String(getCell(r, headerMap, "Course Name", "CourseName", "PaperName", "SubjectName") || r["Course Name"] || "").trim();
      const studentName = String(getCell(r, headerMap, "Student Name", "StudentName", "Name", "CandidateName") || r["Student Name"] || "").trim();
      const rawCollegeCode = String(getCell(r, headerMap, "ADEC Code", "ADECCode", "ADEC_Code", "ADEC", "College Code", "CollegeCode", "College_Code") || r["College Code"] || r["ADEC Code"] || "").trim();
      const rawCollegeName = String(getCell(r, headerMap, "ADEC Name", "ADECName", "ADEC_Name", "ADEC", "College Name", "CollegeName", "College_Name", "College", "Center Name", "Institute") || r["College Name"] || r["ADEC Name"] || "").trim();
      const { code: collegeCode, name: collegeName } = parseCollegeRaw(rawCollegeCode, rawCollegeName);

      const am = String(getCell(r, headerMap, "AM", "Assessment Method", "AssessmentMethod") || r["AM"] || "").trim().toUpperCase();
      const at = String(getCell(r, headerMap, "AT", "Assessment Type", "AssessmentType") || r["AT"] || "").trim().toUpperCase();

      const status = String(getCell(r, headerMap, "Status (e.g., Unresolved, Resolved, SPC, CPC,NP)", "Status", "UM Status", "Malpractice Status") || r["Status (e.g., Unresolved, Resolved, SPC, CPC,NP)"] || r["Status"] || "").trim() || "EHB";
      const remarks = String(getCell(r, headerMap, "Description / Remarks", "Description", "Remarks", "Description/Remarks") || r["Description / Remarks"] || r["Remarks"] || "").trim();
      const umMarkedDate = String(getCell(r, headerMap, "UM Marked Date", "UMMarkedDate", "Marked Date", "Date") || r["UM Marked Date"] || "").trim();

      const normCode = normalizeKey(code);
      const normPrn = normalizeKey(prn);
      const normSeat = normalizeKey(seat);

      const isEseTh = (am.includes("ESE") || !am) && (at.includes("TH") || !at);
      const isEsePr = am.includes("ESE") && at.includes("PR");
      const isCeTh = am.includes("CE") && at.includes("TH");
      const isCePr = am.includes("CE") && at.includes("PR");
      const isWholeEse = am.includes("ESE") && !at;
      const isWholeCourse = !am && !at;

      const details = {
        studentName,
        prn,
        seat,
        code,
        courseName,
        collegeCode,
        collegeName,
        am: am || "ESE",
        at: at || "TH",
        status,
        remarks,
        umMarkedDate,
        isEseTh: isEseTh || isWholeEse || isWholeCourse,
        isEsePr: isEsePr || isWholeEse || isWholeCourse,
        isCeTh: isCeTh || isWholeCourse,
        isCePr: isCePr || isWholeCourse,
        isMalpracticeOverall: isEseTh || isEsePr || isWholeEse || isWholeCourse
      };

      if (normPrn && normCode) {
        map.set(`${normPrn}___${normCode}`, details);
      }
      if (normSeat && normCode) {
        map.set(`${normSeat}___${normCode}`, details);
      }

      if (normCode && (normPrn || normSeat)) {
        list.push(details);
      }
    });

    return { map, list };
  };

  // Helper to match student course record against loaded Heldback Report
  const getHeldbackEntry = (prn, seat, code, heldbackMap) => {
    if (!heldbackMap || heldbackMap.size === 0) return null;
    const normPrn = normalizeKey(prn);
    const normSeat = normalizeKey(seat);
    const normCode = normalizeKey(code);

    // 1. Check term-level heldback (applies to ALL papers for student)
    if (normPrn) {
      const termEntry = heldbackMap.get(`${normPrn}___ALL`);
      if (termEntry) return termEntry;
    }
    if (normSeat) {
      const termEntry = heldbackMap.get(`${normSeat}___ALL`);
      if (termEntry) return termEntry;
    }

    // 2. Check specific paper heldback
    if (normCode) {
      if (normPrn) {
        const paperEntry = heldbackMap.get(`${normPrn}___${normCode}`);
        if (paperEntry) return paperEntry;
      }
      if (normSeat) {
        const paperEntry = heldbackMap.get(`${normSeat}___${normCode}`);
        if (paperEntry) return paperEntry;
      }
    }

    return null;
  };

  // Build lookup map and formatted list from uploaded Heldback Report rows
  const buildHeldbackLookup = (rows, headerMap = {}) => {
    const map = new Map();
    const list = [];

    rows.forEach(r => {
      const prn = String(getCell(r, headerMap, "PRN", "PRN Number", "PRNNo", "RegisterNo", "RegNo", "StudentID") || r["PRN"] || "").trim();
      const seat = String(getCell(r, headerMap, "Seat Number", "SeatNumber", "SeatNo", "Seat_Number", "RollNo", "Roll Number") || r["Seat Number"] || "").trim();
      const reason = String(getCell(r, headerMap, "Reason", "Heldback Reason", "HeldbackReason", "Remarks", "Description") || r["Reason"] || "").trim() || "APC Heldback";
      const rawCollegeCode = String(getCell(r, headerMap, "ADEC Code", "ADECCode", "ADEC_Code", "ADEC", "College Code", "CollegeCode", "College_Code") || r["College Code"] || r["ADEC Code"] || "").trim();
      const rawCollegeName = String(getCell(r, headerMap, "ADEC Name", "ADECName", "ADEC_Name", "ADEC", "College Name", "CollegeName", "College_Name", "College", "Center Name", "Institute") || r["College Name"] || r["ADEC Name"] || "").trim();
      const { code: collegeCode, name: collegeName } = parseCollegeRaw(rawCollegeCode, rawCollegeName);
      const studentName = String(getCell(r, headerMap, "Student Name", "StudentName", "Name", "CandidateName") || r["Student Name"] || "").trim();
      const rawPaper = String(getCell(r, headerMap, "Paper", "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || r["Paper"] || "").trim();
      const assessmentType = String(getCell(r, headerMap, "Assessment Type", "AssessmentType", "AT") || r["Assessment Type"] || "").trim();
      const tlm = String(getCell(r, headerMap, "Teaching Learning Method", "TeachingLearningMethod", "TLM") || r["Teaching Learning Method"] || "").trim();
      const am = String(getCell(r, headerMap, "Assessment Method", "AssessmentMethod", "AM") || r["Assessment Method"] || "").trim();

      const paperLower = rawPaper.toLowerCase();
      const atLower = assessmentType.toLowerCase();
      const isTermLevel = !rawPaper || 
        paperLower.includes("term-level") || 
        paperLower.includes("term level") || 
        atLower.includes("term-level") || 
        atLower.includes("term level") || 
        paperLower === "heldback" ||
        paperLower === "all";

      const paper = isTermLevel ? rawPaper : cleanCourseCode(rawPaper);

      const normPrn = normalizeKey(prn);
      const normSeat = normalizeKey(seat);
      const normPaper = normalizeKey(paper);

      const details = {
        studentName,
        prn,
        seat,
        reason,
        assessmentType,
        collegeCode,
        collegeName,
        paper,
        tlm,
        am,
        isTermLevel
      };

      if (isTermLevel) {
        if (normPrn) map.set(`${normPrn}___ALL`, details);
        if (normSeat) map.set(`${normSeat}___ALL`, details);
      } else {
        if (normPrn && normPaper) map.set(`${normPrn}___${normPaper}`, details);
        if (normSeat && normPaper) map.set(`${normSeat}___${normPaper}`, details);
      }

      if (normPrn || normSeat) {
        list.push(details);
      }
    });

    return { map, list };
  };

  const isAlreadyAggregatedSheet = (hMap) => {
    const keys = Object.keys(hMap);
    const hasEseOverall = keys.some(k => k.includes("eseoverall") || k.includes("esemax") || k.includes("esethobtained") || k.includes("esethmax"));
    const hasCourseOverall = keys.some(k => k.includes("courseoverall") || k.includes("overallmax") || k.includes("overallmin"));
    const hasRawAssessment = keys.some(k => k === "assessmentmethod" || k === "am" || k.includes("assessmenttype") || k === "at");
    return (hasEseOverall || hasCourseOverall) && !hasRawAssessment;
  };

  // Helper to detect if Teaching Learning Method (TLM) requires both Theory and Practical
  const isLecLabTlm = (tlmStr) => {
    if (!tlmStr) return false;
    const s = String(tlmStr).toLowerCase();
    if (s.includes("lec-lab") || s.includes("lec - lab") || s.includes("lecture-lab") || s.includes("lecture - lab")) return true;
    if (s.includes("lec-prac") || s.includes("lecture-practical") || s.includes("theory-practical") || s.includes("th-pr")) return true;
    if ((s.includes("lec") || s.includes("theory") || s.includes("th")) && (s.includes("lab") || s.includes("prac") || s.includes("pr"))) {
      const hasTh = s.includes("lec") || s.includes("theory");
      const hasPr = s.includes("lab") || s.includes("prac");
      if (hasTh && hasPr) return true;
    }
    return false;
  };

  // Group raw assessment rows OR parse pre-aggregated rows into Student-Course Base Aggregates
  const buildGroupedRecordsFromRows = (rows, currentHeaderMap, currentAbsentMap = absentRecordsMap, currentMalpracticeMap = malpracticeRecordsMap, currentHeldbackMap = heldbackRecordsMap) => {
    if (!rows || rows.length === 0) return [];
    const isAgg = isAlreadyAggregatedSheet(currentHeaderMap);
    // Build dataset-wide canonical college registry from all rows and external reports connecting Seat Number and ADEC Name
    const collegeRegistry = buildCollegeCanonicalRegistry(rows, currentHeaderMap, currentAbsentMap, currentMalpracticeMap, currentHeldbackMap);

    // Pass 1: Build expected component profiles for all courses present in the dataset
    const courseExpectedComponentsMap = new Map();

    rows.forEach(row => {
      const rawCode = String(getCell(row, currentHeaderMap, "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || "").trim();
      const code = cleanCourseCode(rawCode);
      const norm = normalizeKey(code);
      if (!norm) return;

      if (!courseExpectedComponentsMap.has(norm)) {
        courseExpectedComponentsMap.set(norm, {
          courseCode: code,
          tlm: "",
          eseMax: 0,
          eseThMax: 0,
          esePrMax: 0,
          ceMax: 0,
          ceThMax: 0,
          cePrMax: 0,
          courseMax: 0,
          hasSeenEseTh: false,
          hasSeenEsePr: false,
          hasSeenCeTh: false,
          hasSeenCePr: false,
          requiresEseTh: false,
          requiresEsePr: false,
          requiresCeTh: false,
          requiresCePr: false,
          hasEseTh: false,
          hasEsePr: false,
          hasCeTh: false,
          hasCePr: false,
          isPrOnly: false,
          maxMarks: {
            ESE_TH: 0,
            ESE_PR: 0,
            CE_TH: 0,
            CE_PR: 0
          }
        });
      }

      const prof = courseExpectedComponentsMap.get(norm);
      const tlmRaw = String(getCell(row, currentHeaderMap, "TeachingLearningMethod", "TLM", "Teaching Learning Method", "Teaching_Learning_Method", "MethodType", "Course Type") || "").trim();
      if (tlmRaw) prof.tlm = tlmRaw;

      if (isAgg) {
        const ese_max_raw = parseNumber(getCell(row, currentHeaderMap, "ESE - Max", "ESEMax", "ESE Max", "ESE_Max"));
        const ese_th_m = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Max", "ESETHMax", "ESE TH Max", "ESE_TH_Max"));
        const ese_pr_m = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Max", "ESEPRMax", "ESE PR Max", "ESE_PR_Max"));
        const ce_max_raw = parseNumber(getCell(row, currentHeaderMap, "CE - Max", "CEMax", "CE Max", "CE_Max"));
        const ce_th_m = parseNumber(getCell(row, currentHeaderMap, "CE - TH Max", "CETHMax", "CE TH Max", "CE_TH_Max"));
        const ce_pr_m = parseNumber(getCell(row, currentHeaderMap, "CE - PR Max", "CEPRMax", "CE PR Max", "CE_PR_Max"));
        const overall_max_raw = parseNumber(getCell(row, currentHeaderMap, "Overall Maximum", "OverallMaximum", "OverallMax", "Course Max", "CourseMax", "Total Max"));

        if (ese_max_raw !== null && ese_max_raw > 0) prof.eseMax = Math.max(prof.eseMax, ese_max_raw);
        if (ese_th_m !== null && ese_th_m > 0) {
          prof.hasSeenEseTh = true;
          prof.eseThMax = Math.max(prof.eseThMax, ese_th_m);
          prof.maxMarks.ESE_TH = Math.max(prof.maxMarks.ESE_TH, ese_th_m);
        }
        if (ese_pr_m !== null && ese_pr_m > 0) {
          prof.hasSeenEsePr = true;
          prof.esePrMax = Math.max(prof.esePrMax, ese_pr_m);
          prof.maxMarks.ESE_PR = Math.max(prof.maxMarks.ESE_PR, ese_pr_m);
        }
        if (ce_max_raw !== null && ce_max_raw > 0) prof.ceMax = Math.max(prof.ceMax, ce_max_raw);
        if (ce_th_m !== null && ce_th_m > 0) {
          prof.hasSeenCeTh = true;
          prof.ceThMax = Math.max(prof.ceThMax, ce_th_m);
          prof.maxMarks.CE_TH = Math.max(prof.maxMarks.CE_TH, ce_th_m);
        }
        if (ce_pr_m !== null && ce_pr_m > 0) {
          prof.hasSeenCePr = true;
          prof.cePrMax = Math.max(prof.cePrMax, ce_pr_m);
          prof.maxMarks.CE_PR = Math.max(prof.maxMarks.CE_PR, ce_pr_m);
        }
        if (overall_max_raw !== null && overall_max_raw > 0) prof.courseMax = Math.max(prof.courseMax, overall_max_raw);
      } else {
        const methodRaw = String(getCell(row, currentHeaderMap, "Assessment Method", "AssessmentMethod", "AM", "Method", "Assessment_Method") || "").trim().toUpperCase();
        const typeRaw = String(getCell(row, currentHeaderMap, "Assessment Type", "AssessmentType", "AT", "Type", "Assessment_Type") || "").trim().toUpperCase();
        const amMaxRaw = parseNumber(getCell(row, currentHeaderMap, "AM Max Marks", "AMMaxMarks", "AM Max", "AMMax", "AM_Max_Marks", "Method Max Marks", "Method Max"));
        const atMaxRaw = parseNumber(getCell(row, currentHeaderMap, "AT Max Marks", "ATMaxMarks", "MaxMarks", "Max Marks", "Max", "AT Max", "AT_Max_Marks"));
        const courseMaxRaw = parseNumber(getCell(row, currentHeaderMap, "Course Max", "CourseMax", "Overall Maximum", "OverallMax", "Total Max"));

        let method = "ESE";
        if (methodRaw.includes("CE") || methodRaw.includes("CA") || methodRaw.includes("IA") || methodRaw.includes("CCA") || methodRaw.includes("INTERNAL")) {
          method = "CE";
        } else if (methodRaw.includes("ESE") || methodRaw.includes("EXT") || methodRaw.includes("EXTERNAL") || methodRaw.includes("THEORY")) {
          method = "ESE";
        }

        let type = "TH";
        if (typeRaw.includes("PR") || typeRaw.includes("PRACTICAL") || typeRaw.includes("VIVA") || typeRaw.includes("LAB")) {
          type = "PR";
        } else if (typeRaw.includes("TH") || typeRaw.includes("THEORY")) {
          type = "TH";
        }

        if (method === "ESE") {
          if (amMaxRaw !== null && amMaxRaw > 0) prof.eseMax = Math.max(prof.eseMax, amMaxRaw);
          if (type === "TH") {
            prof.hasSeenEseTh = true;
            if (atMaxRaw !== null && atMaxRaw > 0) {
              prof.eseThMax = Math.max(prof.eseThMax, atMaxRaw);
              prof.maxMarks.ESE_TH = Math.max(prof.maxMarks.ESE_TH, atMaxRaw);
            }
          } else if (type === "PR") {
            prof.hasSeenEsePr = true;
            if (atMaxRaw !== null && atMaxRaw > 0) {
              prof.esePrMax = Math.max(prof.esePrMax, atMaxRaw);
              prof.maxMarks.ESE_PR = Math.max(prof.maxMarks.ESE_PR, atMaxRaw);
            }
          }
        } else if (method === "CE") {
          if (amMaxRaw !== null && amMaxRaw > 0) prof.ceMax = Math.max(prof.ceMax, amMaxRaw);
          if (type === "TH") {
            prof.hasSeenCeTh = true;
            if (atMaxRaw !== null && atMaxRaw > 0) {
              prof.ceThMax = Math.max(prof.ceThMax, atMaxRaw);
              prof.maxMarks.CE_TH = Math.max(prof.maxMarks.CE_TH, atMaxRaw);
            }
          } else if (type === "PR") {
            prof.hasSeenCePr = true;
            if (atMaxRaw !== null && atMaxRaw > 0) {
              prof.cePrMax = Math.max(prof.cePrMax, atMaxRaw);
              prof.maxMarks.CE_PR = Math.max(prof.maxMarks.CE_PR, atMaxRaw);
            }
          }
        }

        if (courseMaxRaw !== null && courseMaxRaw > 0) {
          prof.courseMax = Math.max(prof.courseMax, courseMaxRaw);
        }
      }
    });

    // Finalize expected components deduction for each course based on mathematical marks rules
    courseExpectedComponentsMap.forEach(prof => {
      // 1. Cross-check with Course Max if available
      if (prof.courseMax > 0) {
        if (prof.eseMax > 0 && prof.ceMax === 0 && prof.courseMax > prof.eseMax) {
          prof.ceMax = prof.courseMax - prof.eseMax;
        } else if (prof.ceMax > 0 && prof.eseMax === 0 && prof.courseMax > prof.ceMax) {
          prof.eseMax = prof.courseMax - prof.ceMax;
        }
      } else if (prof.eseMax > 0 || prof.ceMax > 0) {
        prof.courseMax = prof.eseMax + prof.ceMax;
      }

      // 2. Deduce ESE Component Requirements
      if (prof.eseMax > 0) {
        if (prof.eseThMax > 0 && prof.eseMax > prof.eseThMax) {
          // ESE TH accounts for less than ESE Max -> Must have ESE PR to make up the difference!
          // Example: KU4DSCAFZ206 (ESE Max 65 > ESE TH Max 50, PR Max 15)
          prof.requiresEseTh = true;
          prof.requiresEsePr = true;
          if (prof.esePrMax === 0) prof.esePrMax = prof.eseMax - prof.eseThMax;
          prof.maxMarks.ESE_PR = Math.max(prof.maxMarks.ESE_PR, prof.esePrMax);
        } else if (prof.eseThMax > 0 && prof.eseMax === prof.eseThMax) {
          // ESE TH accounts for the entire ESE Max -> Theory Only, NO ESE PR!
          // Example: KU4DSCCOM208 (ESE Max 70 === ESE TH Max 70)
          prof.requiresEseTh = true;
          prof.requiresEsePr = false;
        } else if (prof.esePrMax > 0 && prof.eseMax === prof.esePrMax) {
          // ESE PR accounts for entire ESE Max -> Practical Only!
          prof.requiresEseTh = false;
          prof.requiresEsePr = true;
        } else if (prof.hasSeenEsePr && prof.hasSeenEseTh) {
          prof.requiresEseTh = true;
          prof.requiresEsePr = true;
        } else if (prof.hasSeenEsePr) {
          prof.requiresEsePr = true;
        } else {
          prof.requiresEseTh = true;
        }
      } else {
        // ESE Max was not explicitly stated
        if (prof.eseThMax > 0 && prof.esePrMax > 0) {
          prof.requiresEseTh = true;
          prof.requiresEsePr = true;
          prof.eseMax = prof.eseThMax + prof.esePrMax;
        } else if (prof.esePrMax > 0 && prof.eseThMax === 0) {
          prof.requiresEseTh = false;
          prof.requiresEsePr = true;
          prof.eseMax = prof.esePrMax;
        } else if (prof.eseThMax > 0) {
          prof.requiresEseTh = true;
          prof.requiresEsePr = false;
          prof.eseMax = prof.eseThMax;
        } else if (prof.hasSeenEsePr && !prof.hasSeenEseTh) {
          prof.requiresEsePr = true;
        } else if (prof.hasSeenEseTh) {
          prof.requiresEseTh = true;
        }
      }

      // 3. Deduce CE Component Requirements
      if (prof.ceMax > 0) {
        if (prof.ceThMax > 0 && prof.ceMax > prof.ceThMax) {
          // CE TH accounts for less than CE Max -> Must have CE PR!
          prof.requiresCeTh = true;
          prof.requiresCePr = true;
          if (prof.cePrMax === 0) prof.cePrMax = prof.ceMax - prof.ceThMax;
          prof.maxMarks.CE_PR = Math.max(prof.maxMarks.CE_PR, prof.cePrMax);
        } else if (prof.ceThMax > 0 && prof.ceMax === prof.ceThMax) {
          // CE TH accounts for entire CE Max -> Theory Only, NO CE PR!
          prof.requiresCeTh = true;
          prof.requiresCePr = false;
        } else if (prof.cePrMax > 0 && prof.ceMax === prof.cePrMax) {
          prof.requiresCeTh = false;
          prof.requiresCePr = true;
        } else if (prof.hasSeenCePr && prof.hasSeenCeTh) {
          prof.requiresCeTh = true;
          prof.requiresCePr = true;
        } else if (prof.hasSeenCePr) {
          prof.requiresCePr = true;
        } else {
          prof.requiresCeTh = true;
        }
      } else {
        if (prof.ceThMax > 0 && prof.cePrMax > 0) {
          prof.requiresCeTh = true;
          prof.requiresCePr = true;
          prof.ceMax = prof.ceThMax + prof.cePrMax;
        } else if (prof.cePrMax > 0 && prof.ceThMax === 0) {
          prof.requiresCeTh = false;
          prof.requiresCePr = true;
          prof.ceMax = prof.cePrMax;
        } else if (prof.ceThMax > 0) {
          prof.requiresCeTh = true;
          prof.requiresCePr = false;
          prof.ceMax = prof.ceThMax;
        } else if (prof.hasSeenCePr && !prof.hasSeenCeTh) {
          prof.requiresCePr = true;
        } else if (prof.hasSeenCeTh) {
          prof.requiresCeTh = true;
        }
      }

      prof.hasEseTh = prof.requiresEseTh;
      prof.hasEsePr = prof.requiresEsePr;
      prof.hasCeTh = prof.requiresCeTh;
      prof.hasCePr = prof.requiresCePr;
      prof.isPrOnly = prof.requiresEsePr && !prof.requiresEseTh;
    });

    if (isAgg) {
      const baseRecords = [];

      rows.forEach((row) => {
        const faculty = String(getCell(row, currentHeaderMap, "Faculty", "Fac", "FacultyName", "Department") || "").trim();
        const program = String(getCell(row, currentHeaderMap, "Program Term Name", "ProgramTermName", "ProgramTerm", "Program Term", "Degree", "Term", "Semester") || "").trim();
        const rawCollegeCode = String(getCell(row, currentHeaderMap, "ADEC Code", "ADECCode", "ADEC_Code", "ADEC", "College Code", "CollegeCode", "College_Code", "Center Code", "CenterCode", "InstCode") || "").trim();
        const rawCollegeName = String(getCell(row, currentHeaderMap, "ADEC Name", "ADECName", "ADEC_Name", "ADEC", "College Name", "CollegeName", "College_Name", "College", "Center Name", "CenterName", "Institute", "Institute Name", "College / Department") || "").trim();
        const seat = String(getCell(row, currentHeaderMap, "Seat Number", "SeatNumber", "SeatNo", "Seat_Number", "RollNo", "Roll Number") || "").trim();
        const prn = String(getCell(row, currentHeaderMap, "PRN", "PRN Number", "PRNNo", "RegisterNo", "RegNo", "StudentID") || "").trim();
        const rawCode = String(getCell(row, currentHeaderMap, "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || "").trim();
        const code = cleanCourseCode(rawCode);
        const name = cleanCourseName(String(getCell(row, currentHeaderMap, "Course Name", "CourseName", "PaperName", "SubjectName", "CourseTitle") || "").trim());

        const normCode = normalizeKey(code);
        const prof = courseExpectedComponentsMap.get(normCode);

        // Check if this student-course has a heldback, malpractice or absent record
        const heldbackEntry = getHeldbackEntry(prn, seat, code, currentHeldbackMap);
        const malpracticeEntry = !heldbackEntry && getMalpracticeEntry(prn, seat, code, currentMalpracticeMap);
        const absentEntry = !heldbackEntry && !malpracticeEntry && getAbsentEntry(prn, seat, code, currentAbsentMap);

        const rawC = rawCollegeCode || heldbackEntry?.collegeCode || malpracticeEntry?.collegeCode || absentEntry?.collegeCode || "";
        const rawN = rawCollegeName || heldbackEntry?.collegeName || malpracticeEntry?.collegeName || absentEntry?.collegeName || "";
        const { collegeCode, collegeName, college } = collegeRegistry.resolve(seat, prn, rawC, rawN);

        const ese_pr_max = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Max", "ESEPRMax")) ?? (prof?.requiresEsePr ? (prof?.maxMarks?.ESE_PR || "") : "");
        const ese_pr_min = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Min", "ESEPRMin")) ?? "";
        let ese_pr_obtained = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Obtained", "ESEPRObtained")) ?? "";

        const ese_th_max = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Max", "ESETHMax")) ?? (prof?.requiresEseTh ? (prof?.maxMarks?.ESE_TH || "") : "");
        const ese_th_min = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Min", "ESETHMin")) ?? (ese_th_max !== "" ? 0 : "");
        let ese_th_obtained = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Obtained", "ESETHObtained")) ?? "";

        let ese_max = parseNumber(getCell(row, currentHeaderMap, "ESE - Max", "ESEMax"));
        if (ese_max === null) {
          ese_max = (parseNumber(ese_pr_max) || 0) + (parseNumber(ese_th_max) || 0);
        }
        if ((ese_max === null || ese_max === 0) && prof?.eseMax > 0) {
          ese_max = prof.eseMax;
        }
        let ese_min = parseNumber(getCell(row, currentHeaderMap, "ESE - Min", "ESEMin"));
        if (ese_min === null) {
          ese_min = Math.ceil(0.30 * ese_max);
        }

        const has_ese_pr = (parseNumber(ese_pr_max) || 0) > 0 || (prof?.requiresEsePr ?? false);
        const has_ese_th = (parseNumber(ese_th_max) || 0) > 0 || (prof?.requiresEseTh ?? false) || !has_ese_pr;
        const is_pr_only = has_ese_pr && !has_ese_th;

        let is_heldback = false;
        let heldback_reason = "";

        let is_malpractice = false;
        let malpractice_status = "";
        let malpractice_remarks = "";
        let malpractice_date = "";

        let is_absent = false;
        let absent_status = "";

        if (heldbackEntry) {
          is_heldback = true;
          heldback_reason = heldbackEntry.reason || "Heldback at term-level";
          ese_th_obtained = "Held";
          ese_pr_obtained = "Held";
        } else if (malpracticeEntry) {
          is_malpractice = true;
          malpractice_status = malpracticeEntry.status;
          malpractice_remarks = malpracticeEntry.remarks;
          malpractice_date = malpracticeEntry.umMarkedDate;
          if (malpracticeEntry.isEseTh) {
            ese_th_obtained = "Malpractice (MP)";
          }
          if (malpracticeEntry.isEsePr) {
            ese_pr_obtained = "Malpractice (MP)";
          }
        } else if (absentEntry) {
          is_absent = true;
          absent_status = absentEntry.status;
          if (absentEntry.isEseTh) {
            ese_th_obtained = "Absent (Ab)";
          }
          if (absentEntry.isEsePr) {
            ese_pr_obtained = "Absent (Ab)";
          }
        }

        // Check for Missing Component(s)
        const missingComponents = [];
        if (!is_heldback && !is_malpractice && !is_absent && prof) {
          const isBlank = (val) => val === undefined || val === null || String(val).trim() === "";
          const rawEsePr = getCell(row, currentHeaderMap, "ESE - PR Obtained", "ESEPRObtained");
          const rawEseTh = getCell(row, currentHeaderMap, "ESE - TH Obtained", "ESETHObtained");
          const rawCePr = getCell(row, currentHeaderMap, "CE - PR Obtained", "CEPRObtained");
          const rawCeTh = getCell(row, currentHeaderMap, "CE - TH Obtained", "CETHObtained");

          if (prof.requiresEsePr && isBlank(rawEsePr)) {
            missingComponents.push("ESE-PR");
            ese_pr_obtained = "Missing";
          }
          if (prof.requiresEseTh && isBlank(rawEseTh)) {
            missingComponents.push("ESE-TH");
            ese_th_obtained = "Missing";
          }
          if (prof.requiresCePr && isBlank(rawCePr)) {
            missingComponents.push("CE-PR");
          }
          if (prof.requiresCeTh && isBlank(rawCeTh)) {
            missingComponents.push("CE-TH");
          }
        }
        const is_held = missingComponents.length > 0;

        let ese_obtained;
        if (is_held) {
          ese_obtained = "Held";
        } else if (is_malpractice && ese_th_obtained === "Malpractice (MP)" && (!has_ese_pr || ese_pr_obtained === "Malpractice (MP)")) {
          ese_obtained = "Malpractice (MP)";
        } else if (is_malpractice && ese_pr_obtained === "Malpractice (MP)" && !has_ese_th) {
          ese_obtained = "Malpractice (MP)";
        } else if (is_absent && ese_th_obtained === "Absent (Ab)" && (!has_ese_pr || ese_pr_obtained === "Absent (Ab)")) {
          ese_obtained = "Absent (Ab)";
        } else if (is_absent && ese_pr_obtained === "Absent (Ab)" && !has_ese_th) {
          ese_obtained = "Absent (Ab)";
        } else {
          ese_obtained = parseNumber(getCell(row, currentHeaderMap, "ESE Overall", "ESEOverall"));
          if (ese_obtained === null) {
            const prNum = parseNumber(ese_pr_obtained) || 0;
            const thNum = parseNumber(ese_th_obtained) || 0;
            ese_obtained = prNum + thNum;
          }
        }

        const ce_pr_max = parseNumber(getCell(row, currentHeaderMap, "CE - PR Max", "CEPRMax")) ?? (prof?.requiresCePr ? (prof?.maxMarks?.CE_PR || "") : "");
        const ce_pr_min = parseNumber(getCell(row, currentHeaderMap, "CE - PR Min", "CEPRMin")) ?? (ce_pr_max !== "" ? 0 : "");
        const ce_pr_obtained = missingComponents.includes("CE-PR") ? "Missing" : (parseNumber(getCell(row, currentHeaderMap, "CE - PR Obtained", "CEPRObtained")) ?? (prof?.requiresCePr ? "" : ""));

        const ce_th_max = parseNumber(getCell(row, currentHeaderMap, "CE - TH Max", "CETHMax")) ?? (prof?.requiresCeTh ? (prof?.maxMarks?.CE_TH || "") : "");
        const ce_th_min = parseNumber(getCell(row, currentHeaderMap, "CE - TH Min", "CETHMin")) ?? (ce_th_max !== "" ? 0 : "");
        const ce_th_obtained = missingComponents.includes("CE-TH") ? "Missing" : (parseNumber(getCell(row, currentHeaderMap, "CE - TH Obtained", "CETHObtained")) ?? (prof?.requiresCeTh ? "" : ""));

        let ce_max = parseNumber(getCell(row, currentHeaderMap, "CE - Max", "CEMax"));
        if (ce_max === null) {
          ce_max = (parseNumber(ce_pr_max) || 0) + (parseNumber(ce_th_max) || 0);
        }
        if ((ce_max === null || ce_max === 0) && prof?.ceMax > 0) {
          ce_max = prof.ceMax;
        }
        let ce_min = 0;
        let ce_obtained;
        if (missingComponents.includes("CE-PR") || missingComponents.includes("CE-TH")) {
          ce_obtained = "Held";
        } else {
          ce_obtained = parseNumber(getCell(row, currentHeaderMap, "CE Overall Marks ", "CE Overall Marks", "CEOverallMarks", "CEOverall"));
          if (ce_obtained === null) {
            ce_obtained = (parseNumber(ce_pr_obtained) || 0) + (parseNumber(ce_th_obtained) || 0);
          }
        }

        let overall_max = parseNumber(getCell(row, currentHeaderMap, "Overall Maximum", "OverallMaximum", "OverallMax", "Course Max"));
        if (overall_max === null || overall_max === 0) {
          overall_max = ese_max + ce_max;
        }
        let overall_min = parseNumber(getCell(row, currentHeaderMap, "Overall Minimum", "OverallMinimum", "OverallMin"));
        if (overall_min === null) {
          overall_min = Math.ceil(0.35 * overall_max);
        }
        let course_overall;
        if (is_held) {
          course_overall = "Held";
        } else {
          course_overall = parseNumber(getCell(row, currentHeaderMap, "Course Overall Marks ", "Course Overall Marks", "CourseOverallMarks"));
          if (course_overall === null) {
            course_overall = (parseNumber(ese_obtained) || 0) + (parseNumber(ce_obtained) || 0);
          }
        }

        let ese_deficit = 0;
        let overall_deficit = 0;
        let raw_ese_pass = false;
        let raw_overall_pass = false;
        let raw_course_pass = false;

        if (is_held) {
          raw_ese_pass = false;
          raw_overall_pass = false;
          raw_course_pass = false;
          ese_deficit = 999;
          overall_deficit = 999;
        } else if (is_malpractice) {
          raw_ese_pass = false;
          raw_overall_pass = false;
          raw_course_pass = false;
          ese_deficit = 999;
          overall_deficit = 999;
        } else if (is_absent) {
          raw_ese_pass = false;
          raw_overall_pass = false;
          raw_course_pass = false;
          ese_deficit = 999;
          overall_deficit = 999;
        } else {
          const numEse = parseNumber(ese_obtained) || 0;
          ese_deficit = Math.max(0, ese_min - numEse);
          overall_deficit = Math.max(0, overall_min - (parseNumber(course_overall) || 0));
          raw_ese_pass = ese_deficit === 0;
          raw_overall_pass = overall_deficit === 0;
          raw_course_pass = raw_ese_pass && raw_overall_pass;
        }

        baseRecords.push({
          identifiers: { faculty, program, seat, prn, code, name, college, collegeCode, collegeName },
          _college: college,
          _collegeCode: collegeCode,
          _collegeName: collegeName,
          raw: {
            "Faculty": faculty,
            "Program Term Name": program,
            "Course Code": code,
            "Course Name": name,
            "Seat Number": seat,
            "PRN": prn,
            "ESE - PR Max": ese_pr_max,
            "ESE - PR Min": ese_pr_min,
            "ESE - PR Obtained": ese_pr_obtained,
            "ESE - TH Max": ese_th_max,
            "ESE - TH Min": ese_th_min,
            "ESE - TH Obtained": ese_th_obtained,
            "ESE - Max": ese_max,
            "ESE - Min": ese_min,
            "ESE Overall": ese_obtained,
            "CE - PR Max": ce_pr_max,
            "CE - PR Min": ce_pr_min,
            "CE - PR Obtained": ce_pr_obtained,
            "CE - TH Max": ce_th_max,
            "CE - TH Min": ce_th_min,
            "CE - TH Obtained": ce_th_obtained,
            "CE - Max": ce_max,
            "CE - Min": ce_min,
            "CE Overall Marks ": ce_obtained,
            "Overall Maximum": overall_max,
            "Overall Minimum": overall_min,
            "Course Overall Marks ": course_overall,
          },
          has_ese_th,
          has_ese_pr,
          is_pr_only,
          ese_deficit,
          overall_deficit,
          raw_ese_pass,
          raw_overall_pass,
          raw_course_pass,
          is_absent,
          absent_status,
          is_malpractice,
          malpractice_status,
          malpractice_remarks,
          malpractice_date,
          is_held,
          is_heldback,
          heldback_reason,
          is_missing_component: is_missing,
          missing_components: missingComponents
        });
      });

      return baseRecords;
    }

    // Standard Raw Assessment Grouping
    const groups = new Map();

    rows.forEach((row) => {
      const faculty = String(getCell(row, currentHeaderMap, "Faculty", "Fac", "FacultyName", "Department") || "").trim();
      const program = String(getCell(row, currentHeaderMap, "Program Term Name", "ProgramTermName", "ProgramTerm", "Program Term", "Degree", "Term", "Semester") || "").trim();
      const rawCollegeCode = String(getCell(row, currentHeaderMap, "ADEC Code", "ADECCode", "ADEC_Code", "ADEC", "College Code", "CollegeCode", "College_Code", "Center Code", "CenterCode", "InstCode") || "").trim();
      const rawCollegeName = String(getCell(row, currentHeaderMap, "ADEC Name", "ADECName", "ADEC_Name", "ADEC", "College Name", "CollegeName", "College_Name", "College", "Center Name", "CenterName", "Institute", "Institute Name", "College / Department") || "").trim();
      const seat = String(getCell(row, currentHeaderMap, "Seat Number", "SeatNumber", "SeatNo", "Seat_Number", "RollNo", "Roll Number") || "").trim();
      const prn = String(getCell(row, currentHeaderMap, "PRN", "PRN Number", "PRNNo", "RegisterNo", "RegNo", "StudentID") || "").trim();
      const rawCode = String(getCell(row, currentHeaderMap, "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || "").trim();
      const code = cleanCourseCode(rawCode);
      const name = cleanCourseName(String(getCell(row, currentHeaderMap, "Course Name", "CourseName", "PaperName", "SubjectName", "CourseTitle") || "").trim());

      const methodRaw = String(getCell(row, currentHeaderMap, "Assessment Method", "AssessmentMethod", "AM", "Method", "Assessment_Method") || "").trim().toUpperCase();
      const typeRaw = String(getCell(row, currentHeaderMap, "Assessment Type", "AssessmentType", "AT", "Type", "Assessment_Type") || "").trim().toUpperCase();
      const marksRaw = parseNumber(getCell(row, currentHeaderMap, "Marks", "ObtainedMarks", "Mark", "Obtained"));
      const atMaxRaw = parseNumber(getCell(row, currentHeaderMap, "AT Max Marks", "ATMaxMarks", "MaxMarks", "Max Marks", "Max", "AT Max", "AT_Max_Marks"));
      const rawMarksVal = getCell(row, currentHeaderMap, "Marks", "ObtainedMarks", "Mark", "Obtained");

      let method = "ESE";
      if (methodRaw.includes("CE") || methodRaw.includes("CA") || methodRaw.includes("IA") || methodRaw.includes("CCA") || methodRaw.includes("INTERNAL")) {
        method = "CE";
      } else if (methodRaw.includes("ESE") || methodRaw.includes("EXT") || methodRaw.includes("EXTERNAL") || methodRaw.includes("THEORY")) {
        method = "ESE";
      }

      let type = "TH";
      if (typeRaw.includes("PR") || typeRaw.includes("PRACTICAL") || typeRaw.includes("VIVA") || typeRaw.includes("LAB")) {
        type = "PR";
      } else if (typeRaw.includes("TH") || typeRaw.includes("THEORY")) {
        type = "TH";
      }

      const groupKey = faculty + "|||" + program + "|||" + seat + "|||" + prn + "|||" + code;

      if (!groups.has(groupKey)) {
        groups.set(groupKey, {
          identifiers: { faculty, program, seat, prn, code, name, rawCollegeCode, rawCollegeName },
          components: {},
          tlm: ""
        });
      } else {
        const grp = groups.get(groupKey);
        if (!grp.identifiers.rawCollegeCode && rawCollegeCode) grp.identifiers.rawCollegeCode = rawCollegeCode;
        if (!grp.identifiers.rawCollegeName && rawCollegeName) grp.identifiers.rawCollegeName = rawCollegeName;
        if ((!grp.identifiers.name || (name && name.length > grp.identifiers.name.length)) && name) grp.identifiers.name = name;
      }

      const group = groups.get(groupKey);
      const tlmVal = String(getCell(row, currentHeaderMap, "TeachingLearningMethod", "TLM", "Teaching Learning Method", "Teaching_Learning_Method", "MethodType", "Course Type") || "").trim();
      if (tlmVal) group.tlm = tlmVal;

      group.components[method + "_" + type] = {
        marks: marksRaw !== null ? marksRaw : null,
        max: atMaxRaw !== null ? atMaxRaw : null,
        present: true,
        rawMarksVal
      };
    });

    const baseRecords = [];

    groups.forEach(({ identifiers, components }) => {
      const normCode = normalizeKey(identifiers.code);
      const prof = courseExpectedComponentsMap.get(normCode);

      const heldbackEntry = getHeldbackEntry(identifiers.prn, identifiers.seat, identifiers.code, currentHeldbackMap);
      const is_heldback = !!heldbackEntry;
      const heldback_reason = heldbackEntry ? (heldbackEntry.reason || "Heldback at term-level") : "";

      const malpracticeEntry = !is_heldback && getMalpracticeEntry(identifiers.prn, identifiers.seat, identifiers.code, currentMalpracticeMap);
      const is_malpractice = !is_heldback && !!malpracticeEntry;
      const malpractice_status = malpracticeEntry ? malpracticeEntry.status : "";
      const malpractice_remarks = malpracticeEntry ? malpracticeEntry.remarks : "";
      const malpractice_date = malpracticeEntry ? malpracticeEntry.umMarkedDate : "";

      const absentEntry = !is_heldback && !is_malpractice && getAbsentEntry(identifiers.prn, identifiers.seat, identifiers.code, currentAbsentMap);
      const is_absent = !is_heldback && !is_malpractice && !!absentEntry;
      const absent_status = absentEntry ? absentEntry.status : "";

      // Check for Missing Component(s)
      const missingComponents = [];
      const isBlank = (val) => val === undefined || val === null || String(val).trim() === "";

      if (!is_heldback && !is_malpractice && !is_absent && prof) {
        // Check ESE_PR
        if (prof.requiresEsePr) {
          const c = components["ESE_PR"];
          if (!c || !c.present || isBlank(c.rawMarksVal)) {
            missingComponents.push("ESE-PR");
          }
        }
        // Check ESE_TH
        if (prof.requiresEseTh) {
          const c = components["ESE_TH"];
          if (!c || !c.present || isBlank(c.rawMarksVal)) {
            missingComponents.push("ESE-TH");
          }
        }
        // Check CE_PR
        if (prof.requiresCePr) {
          const c = components["CE_PR"];
          if (!c || !c.present || isBlank(c.rawMarksVal)) {
            missingComponents.push("CE-PR");
          }
        }
        // Check CE_TH
        if (prof.requiresCeTh) {
          const c = components["CE_TH"];
          if (!c || !c.present || isBlank(c.rawMarksVal)) {
            missingComponents.push("CE-TH");
          }
        }
      }
      const is_missing = missingComponents.length > 0;
      const is_held = is_heldback || is_missing;

      const ese_pr = components["ESE_PR"];
      const ese_pr_max = (ese_pr && ese_pr.max !== null) ? ese_pr.max : (prof?.requiresEsePr ? (prof?.maxMarks?.ESE_PR || 0) : 0);
      let ese_pr_obtained = (ese_pr && ese_pr.marks !== null) ? ese_pr.marks : (missingComponents.includes("ESE-PR") ? "Missing" : (prof?.requiresEsePr ? 0 : ""));

      const ese_th = components["ESE_TH"];
      const ese_th_max = (ese_th && ese_th.max !== null) ? ese_th.max : (prof?.requiresEseTh ? (prof?.maxMarks?.ESE_TH || 0) : 0);
      let ese_th_obtained = (ese_th && ese_th.marks !== null) ? ese_th.marks : (missingComponents.includes("ESE-TH") ? "Missing" : (prof?.requiresEseTh ? 0 : ""));

      if (is_heldback) {
        ese_th_obtained = "Held";
        ese_pr_obtained = "Held";
      } else if (is_malpractice && malpracticeEntry) {
        if (malpracticeEntry.isEseTh) {
          ese_th_obtained = "Malpractice (MP)";
        }
        if (malpracticeEntry.isEsePr) {
          ese_pr_obtained = "Malpractice (MP)";
        }
      } else if (is_absent && absentEntry) {
        if (absentEntry.isEseTh) {
          ese_th_obtained = "Absent (Ab)";
        }
        if (absentEntry.isEsePr) {
          ese_pr_obtained = "Absent (Ab)";
        }
      }

      const ese_max = Math.max(ese_pr_max + ese_th_max, prof?.eseMax || 0);
      const ese_min = Math.ceil(0.30 * ese_max);

      let ese_obtained;
      if (is_held) {
        ese_obtained = "Held";
      } else if (is_malpractice && ese_th_obtained === "Malpractice (MP)" && (ese_pr_max === 0 || ese_pr_obtained === "Malpractice (MP)")) {
        ese_obtained = "Malpractice (MP)";
      } else if (is_malpractice && ese_pr_obtained === "Malpractice (MP)" && ese_th_max === 0) {
        ese_obtained = "Malpractice (MP)";
      } else if (is_absent && ese_th_obtained === "Absent (Ab)" && (ese_pr_max === 0 || ese_pr_obtained === "Absent (Ab)")) {
        ese_obtained = "Absent (Ab)";
      } else if (is_absent && ese_pr_obtained === "Absent (Ab)" && ese_th_max === 0) {
        ese_obtained = "Absent (Ab)";
      } else {
        const numPr = parseNumber(ese_pr_obtained) || 0;
        const numTh = parseNumber(ese_th_obtained) || 0;
        ese_obtained = numPr + numTh;
      }

      const ce_pr = components["CE_PR"];
      const ce_pr_max = (ce_pr && ce_pr.max !== null) ? ce_pr.max : (prof?.requiresCePr ? (prof?.maxMarks?.CE_PR || 0) : 0);
      const ce_pr_obtained = (ce_pr && ce_pr.marks !== null) ? ce_pr.marks : (missingComponents.includes("CE-PR") ? "Missing" : (prof?.requiresCePr ? 0 : ""));

      const ce_th = components["CE_TH"];
      const ce_th_max = (ce_th && ce_th.max !== null) ? ce_th.max : (prof?.requiresCeTh ? (prof?.maxMarks?.CE_TH || 0) : 0);
      const ce_th_obtained = (ce_th && ce_th.marks !== null) ? ce_th.marks : (missingComponents.includes("CE-TH") ? "Missing" : (prof?.requiresCeTh ? 0 : ""));

      const ce_max = Math.max(ce_pr_max + ce_th_max, prof?.ceMax || 0);
      const ce_min = 0;
      let ce_obtained;
      if (missingComponents.includes("CE-PR") || missingComponents.includes("CE-TH")) {
        ce_obtained = "Held";
      } else {
        ce_obtained = (parseNumber(ce_pr_obtained) || 0) + (parseNumber(ce_th_obtained) || 0);
      }

      const overall_max = Math.max(ese_max + ce_max, prof?.courseMax || 0);
      const overall_min = Math.ceil(0.35 * overall_max);
      let course_overall;
      if (is_held) {
        course_overall = "Held";
      } else {
        course_overall = (parseNumber(ese_obtained) || 0) + (parseNumber(ce_obtained) || 0);
      }

      let ese_deficit = 0;
      let overall_deficit = 0;
      let raw_ese_pass = false;
      let raw_overall_pass = false;
      let raw_course_pass = false;

      if (is_held) {
        raw_ese_pass = false;
        raw_overall_pass = false;
        raw_course_pass = false;
        ese_deficit = 999;
        overall_deficit = 999;
      } else if (is_malpractice) {
        raw_ese_pass = false;
        raw_overall_pass = false;
        raw_course_pass = false;
        ese_deficit = 999;
        overall_deficit = 999;
      } else if (is_absent) {
        raw_ese_pass = false;
        raw_overall_pass = false;
        raw_course_pass = false;
        ese_deficit = 999;
        overall_deficit = 999;
      } else {
        const numEse = parseNumber(ese_obtained) || 0;
        ese_deficit = Math.max(0, ese_min - numEse);
        overall_deficit = Math.max(0, overall_min - (parseNumber(course_overall) || 0));
        raw_ese_pass = ese_deficit === 0;
        raw_overall_pass = overall_deficit === 0;
        raw_course_pass = raw_ese_pass && raw_overall_pass;
      }

      const has_ese_pr = (parseNumber(ese_pr_max) || 0) > 0 || (prof?.requiresEsePr ?? false);
      const has_ese_th = (parseNumber(ese_th_max) || 0) > 0 || (prof?.requiresEseTh ?? false) || !has_ese_pr;
      const is_pr_only = has_ese_pr && !has_ese_th;

      const rawC = identifiers.rawCollegeCode || heldbackEntry?.collegeCode || malpracticeEntry?.collegeCode || absentEntry?.collegeCode || "";
      const rawN = identifiers.rawCollegeName || heldbackEntry?.collegeName || malpracticeEntry?.collegeName || absentEntry?.collegeName || "";
      const { collegeCode, collegeName, college } = collegeRegistry.resolve(identifiers.seat, identifiers.prn, rawC, rawN);

      baseRecords.push({
        identifiers: { ...identifiers, college, collegeCode, collegeName },
        _college: college,
        _collegeCode: collegeCode,
        _collegeName: collegeName,
        raw: {
          "Faculty": identifiers.faculty,
          "Program Term Name": identifiers.program,
          "Course Code": identifiers.code,
          "Course Name": identifiers.name,
          "Seat Number": identifiers.seat,
          "PRN": identifiers.prn,
          "ESE - PR Max": ese_pr_max !== 0 ? ese_pr_max : "",
          "ESE - PR Min": "",
          "ESE - PR Obtained": (ese_pr && ese_pr_obtained === "Malpractice (MP)") ? "Malpractice (MP)" : ((ese_pr && ese_pr_obtained === "Absent (Ab)") ? "Absent (Ab)" : ese_pr_obtained),
          "ESE - TH Max": ese_th_max !== 0 ? ese_th_max : "",
          "ESE - TH Min": ese_th_max > 0 ? 0 : "",
          "ESE - TH Obtained": (ese_th && ese_th_obtained === "Malpractice (MP)") ? "Malpractice (MP)" : ((ese_th && ese_th_obtained === "Absent (Ab)") ? "Absent (Ab)" : ese_th_obtained),
          "ESE - Max": Math.round(ese_max),
          "ESE - Min": ese_min,
          "ESE Overall": is_held ? "Held" : (ese_obtained === "Malpractice (MP)" ? "Malpractice (MP)" : (ese_obtained === "Absent (Ab)" ? "Absent (Ab)" : Math.round(ese_obtained))),
          "CE - PR Max": ce_pr_max !== 0 ? ce_pr_max : "",
          "CE - PR Min": ce_pr_max > 0 ? 0 : "",
          "CE - PR Obtained": ce_pr_obtained,
          "CE - TH Max": ce_th_max !== 0 ? ce_th_max : "",
          "CE - TH Min": ce_th_max > 0 ? 0 : "",
          "CE - TH Obtained": ce_th_obtained,
          "CE - Max": Math.round(ce_max),
          "CE - Min": ce_min,
          "CE Overall Marks ": is_held && (ce_obtained === "Held") ? "Held" : Math.round(parseNumber(ce_obtained) || 0),
          "Overall Maximum": overall_max,
          "Overall Minimum": overall_min,
          "Course Overall Marks ": is_held ? "Held" : course_overall,
        },
        has_ese_th,
        has_ese_pr,
        is_pr_only,
        ese_deficit,
        overall_deficit,
        raw_ese_pass,
        raw_overall_pass,
        raw_course_pass,
        is_absent,
        absent_status,
        is_malpractice,
        malpractice_status,
        malpractice_remarks,
        malpractice_date,
        is_held,
        is_heldback,
        heldback_reason,
        is_missing_component: is_missing,
        missing_components: missingComponents
      });
    });

    return baseRecords;
  };

  // Distinct Courses extracted from loaded data
  const distinctCourses = useMemo(() => {
    const map = new Map();
    groupedRecords.forEach(rec => {
      const code = rec.identifiers.code;
      const name = rec.identifiers.name;
      const norm = normalizeKey(code);
      if (!norm) return;

      if (!map.has(norm)) {
        map.set(norm, {
          normCode: norm,
          courseCode: code,
          courseName: name,
          faculty: rec.identifiers.faculty,
          program: rec.identifiers.program,
          hasEseTh: rec.has_ese_th ?? !rec.has_ese_pr,
          hasEsePr: rec.has_ese_pr ?? false,
          isPrOnly: rec.is_pr_only ?? false,
          totalStudents: 0,
          rawPassed: 0,
          rawFailed: 0,
          absentCount: 0,
          malpracticeCount: 0,
          heldbackCount: 0,
          heldCount: 0,
          nearPassCount: 0
        });
      }

      const item = map.get(norm);
      if ((!item.courseName || (name && name.length > item.courseName.length)) && name) {
        item.courseName = name;
      }
      item.hasEseTh = item.hasEseTh || !!rec.has_ese_th;
      item.hasEsePr = item.hasEsePr || !!rec.has_ese_pr;
      item.isPrOnly = item.hasEsePr && !item.hasEseTh;
      item.totalStudents++;
      if (rec.is_heldback) {
        item.heldbackCount = (item.heldbackCount || 0) + 1;
        item.heldCount = (item.heldCount || 0) + 1;
        // Heldback papers are strictly reserved/uncalculated
      } else if (rec.is_held) {
        item.heldCount = (item.heldCount || 0) + 1;
        // Held papers are uncalculated: not rawPassed and not in nearPassCount
      } else if (rec.is_malpractice) {
        item.malpracticeCount = (item.malpracticeCount || 0) + 1;
        item.rawFailed++;
        // Malpractice students cannot be rescued by moderation; do not count in nearPassCount
      } else if (rec.is_absent) {
        item.absentCount++;
        item.rawFailed++;
        // Absent students cannot be rescued by moderation; do not count in nearPassCount
      } else if (rec.raw_course_pass) {
        item.rawPassed++;
      } else {
        item.rawFailed++;
        const maxDeficit = Math.max(rec.ese_deficit, rec.overall_deficit);
        if (maxDeficit > 0 && maxDeficit <= 5) {
          item.nearPassCount++;
        }
      }
    });

    return Array.from(map.values()).sort((a, b) => a.courseCode.localeCompare(b.courseCode));
  }, [groupedRecords]);

  // Apply Course Moderation Rules and compute processedRows
  const processedRows = useMemo(() => {
    return groupedRecords.map(rec => {
      const row = { ...rec.raw };
      row._college = rec._college || rec.identifiers?.college || "";
      row._collegeCode = rec._collegeCode || rec.identifiers?.collegeCode || "";
      row._collegeName = rec._collegeName || rec.identifiers?.collegeName || "";
      const normCode = normalizeKey(rec.identifiers.code);
      const modLimit = courseModerationMap[normCode] || 0;

      // Heldback Student Handling: Strictly locked from calculation, moderation, and pass simulation; Result Held
      if (rec.is_heldback) {
        row["ESE Pass"] = "Held";
        row["Overall pass"] = "Held";
        row["Course Pass/Fail"] = "Held (Heldback)";
        row["Moderation Marks"] = 0;

        row._rawPass = false;
        row._isModeratedPass = false;
        row._isAbsent = false;
        row._absentStatus = "";
        row._isMalpractice = false;
        row._malpracticeStatus = "";
        row._malpracticeRemarks = "";
        row._malpracticeDate = "";
        row._isHeld = true;
        row._isHeldback = true;
        row._heldbackReason = rec.heldback_reason || "Heldback at term-level";
        row._isMissingComp = false;
        row._missingComponents = [];
        row._modLimit = modLimit;
        row._isEligibleForMod = false;
        row._isPrOnly = rec.is_pr_only;
        row._eseDeficit = 999;
        row._overallDeficit = 999;
        return row;
      }

      // Held / Missing Component Handling: Never eligible for moderation, Result Held
      if (rec.is_held) {
        row["ESE Pass"] = "Held";
        row["Overall pass"] = "Held";
        row["Course Pass/Fail"] = "Held (Missing Component)";
        row["Moderation Marks"] = 0;

        row._rawPass = false;
        row._isModeratedPass = false;
        row._isAbsent = false;
        row._absentStatus = "";
        row._isMalpractice = false;
        row._malpracticeStatus = "";
        row._malpracticeRemarks = "";
        row._malpracticeDate = "";
        row._isHeld = true;
        row._isHeldback = false;
        row._heldbackReason = "";
        row._isMissingComp = true;
        row._missingComponents = rec.missing_components || [];
        row._modLimit = modLimit;
        row._isEligibleForMod = false;
        row._isPrOnly = rec.is_pr_only;
        row._eseDeficit = 999;
        row._overallDeficit = 999;
        return row;
      }

      // Malpractice Student Handling: Never eligible for moderation, automatically Fail
      if (rec.is_malpractice) {
        row["ESE Pass"] = "Fail";
        row["Overall pass"] = "Fail";
        row["Course Pass/Fail"] = "Fail";
        row["Moderation Marks"] = 0;

        row._rawPass = false;
        row._isModeratedPass = false;
        row._isAbsent = false;
        row._absentStatus = "";
        row._isMalpractice = true;
        row._malpracticeStatus = rec.malpractice_status || "EHB";
        row._malpracticeRemarks = rec.malpractice_remarks || "";
        row._malpracticeDate = rec.malpractice_date || "";
        row._isHeld = false;
        row._isHeldback = false;
        row._heldbackReason = "";
        row._isMissingComp = false;
        row._missingComponents = [];
        row._modLimit = modLimit;
        row._isEligibleForMod = false;
        row._isPrOnly = rec.is_pr_only;
        row._eseDeficit = rec.ese_deficit;
        row._overallDeficit = rec.overall_deficit;
        return row;
      }

      // Absent Student Handling: Never eligible for moderation, automatically Fail
      if (rec.is_absent) {
        row["ESE Pass"] = "Fail";
        row["Overall pass"] = "Fail";
        row["Course Pass/Fail"] = "Fail";
        row["Moderation Marks"] = 0;

        row._rawPass = false;
        row._isModeratedPass = false;
        row._isAbsent = true;
        row._absentStatus = rec.absent_status || "Marked Absent";
        row._isMalpractice = false;
        row._malpracticeStatus = "";
        row._malpracticeRemarks = "";
        row._malpracticeDate = "";
        row._isHeld = false;
        row._isHeldback = false;
        row._heldbackReason = "";
        row._isMissingComp = false;
        row._missingComponents = [];
        row._modLimit = modLimit;
        row._isEligibleForMod = false;
        row._isPrOnly = rec.is_pr_only;
        row._eseDeficit = rec.ese_deficit;
        row._overallDeficit = rec.overall_deficit;
        return row;
      }

      // Moderation eligibility: Effected if course has ESE-TH, OR if it is ESE-PR only AND user enabled allowPrOnlyModeration
      const isEligibleForModeration = rec.has_ese_th || (rec.is_pr_only && allowPrOnlyModeration);

      let moderation_awarded = 0;
      let final_ese_pass = rec.raw_ese_pass ? "Pass" : "Fail";
      let final_overall_pass = rec.raw_overall_pass ? "Pass" : "Fail";
      let final_course_pass = rec.raw_course_pass ? "Pass" : "Fail";
      let is_moderated_pass = false;

      if (!rec.raw_course_pass && modLimit > 0 && isEligibleForModeration) {
        const marks_needed = Math.max(rec.ese_deficit, rec.overall_deficit);
        if (marks_needed <= modLimit) {
          moderation_awarded = marks_needed;
          final_ese_pass = "Pass";
          final_overall_pass = "Pass";
          final_course_pass = "Pass";
          is_moderated_pass = true;
        }
      }

      row["ESE Pass"] = final_ese_pass;
      row["Overall pass"] = final_overall_pass;
      row["Course Pass/Fail"] = final_course_pass;
      row["Moderation Marks"] = moderation_awarded;

      // Internal flags for UI rendering & statistics
      row._rawPass = rec.raw_course_pass;
      row._isModeratedPass = is_moderated_pass;
      row._isAbsent = false;
      row._absentStatus = "";
      row._isMalpractice = false;
      row._malpracticeStatus = "";
      row._malpracticeRemarks = "";
      row._malpracticeDate = "";
      row._isHeld = false;
      row._isHeldback = false;
      row._heldbackReason = "";
      row._isMissingComp = false;
      row._missingComponents = [];
      row._modLimit = modLimit;
      row._isEligibleForMod = isEligibleForModeration;
      row._isPrOnly = rec.is_pr_only;
      row._eseDeficit = rec.ese_deficit;
      row._overallDeficit = rec.overall_deficit;

      return row;
    });
  }, [groupedRecords, courseModerationMap, allowPrOnlyModeration]);

  // Course-Wise Pass Simulation (0 to +10 Moderation Marks) - Enforcing both 30% ESE & 35% Overall Pass Conditions
  const courseSimulationData = useMemo(() => {
    const map = new Map();

    groupedRecords.forEach(rec => {
      const code = rec.identifiers.code;
      const name = rec.identifiers.name;
      const faculty = rec.identifiers.faculty;
      const program = rec.identifiers.program;
      const norm = normalizeKey(code);
      if (!norm) return;

      if (!map.has(norm)) {
        map.set(norm, {
          normCode: norm,
          courseCode: code,
          courseName: name,
          faculty: faculty || "",
          program: program || "",
          hasEseTh: rec.has_ese_th ?? !rec.has_ese_pr,
          hasEsePr: rec.has_ese_pr ?? false,
          isPrOnly: rec.is_pr_only ?? false,
          isEligible: (rec.has_ese_th ?? !rec.has_ese_pr) || (rec.is_pr_only && allowPrOnlyModeration),
          totalStudents: 0,
          heldCount: 0,
          heldbackCount: 0,
          absentCount: 0,
          malpracticeCount: 0,
          rawEsePassCount: 0,       // 30% ESE Rule Pass Count
          rawOverallPassCount: 0,   // 35% Overall Course Rule Pass Count
          rawPassCount: 0,          // Combined (Both 30% ESE & 35% Overall Met)
          esePassAtMod: Array(11).fill(0),
          overallPassAtMod: Array(11).fill(0),
          passCountAtMod: Array(11).fill(0), // Combined Dual-Condition Pass at +0..+10
        });
      }

      const item = map.get(norm);
      if ((!item.courseName || (name && name.length > item.courseName.length)) && name) {
        item.courseName = name;
      }
      item.hasEseTh = item.hasEseTh || !!rec.has_ese_th;
      item.hasEsePr = item.hasEsePr || !!rec.has_ese_pr;
      item.isPrOnly = item.hasEsePr && !item.hasEseTh;
      item.isEligible = item.hasEseTh || (item.isPrOnly && allowPrOnlyModeration);
      item.totalStudents++;

      // Heldback students are strictly locked and excluded from pass calculations at all moderation levels
      if (rec.is_heldback) {
        item.heldbackCount = (item.heldbackCount || 0) + 1;
        item.heldCount = (item.heldCount || 0) + 1;
        return;
      }

      // Held (Missing Component) records are strictly excluded from pass calculations at all moderation levels
      if (rec.is_held) {
        item.heldCount = (item.heldCount || 0) + 1;
        return;
      }

      // Absent or Malpractice students are strictly excluded from pass counts at all moderation levels
      if (rec.is_malpractice) {
        item.malpracticeCount = (item.malpracticeCount || 0) + 1;
        return;
      }
      if (rec.is_absent) {
        item.absentCount++;
        return;
      }

      const ese_deficit = rec.ese_deficit;
      const overall_deficit = rec.overall_deficit;
      const is_raw_ese_pass = rec.raw_ese_pass;         // 30% ESE Rule
      const is_raw_overall_pass = rec.raw_overall_pass; // 35% Overall Rule
      const is_raw_pass = rec.raw_course_pass;          // Both 30% and 35% Rules

      if (is_raw_ese_pass) {
        item.rawEsePassCount++;
      }
      if (is_raw_overall_pass) {
        item.rawOverallPassCount++;
      }
      if (is_raw_pass) {
        item.rawPassCount++;
      }

      const isEligible = rec.has_ese_th || (rec.is_pr_only && allowPrOnlyModeration);

      // Evaluate simulated pass for each moderation mark level from 0 to 10
      // Student is awarded pass only if BOTH 30% ESE and 35% Overall conditions are met
      for (let m = 0; m <= 10; m++) {
        const meetsEse = is_raw_ese_pass || (isEligible && ese_deficit <= m);
        const meetsOverall = is_raw_overall_pass || (isEligible && overall_deficit <= m);
        const meetsBoth = meetsEse && meetsOverall; // Precisely is_raw_pass || (isEligible && marks_needed <= m)

        if (meetsEse) item.esePassAtMod[m]++;
        if (meetsOverall) item.overallPassAtMod[m]++;
        if (meetsBoth) item.passCountAtMod[m]++;
      }
    });

    return Array.from(map.values()).sort((a, b) => a.courseCode.localeCompare(b.courseCode));
  }, [groupedRecords, allowPrOnlyModeration]);

  // Filtered Course Simulation Data (based on simSearchQuery)
  const filteredSimulationCourses = useMemo(() => {
    if (!simSearchQuery.trim()) return courseSimulationData;
    const q = simSearchQuery.toLowerCase().trim();
    return courseSimulationData.filter(c => 
      c.courseCode.toLowerCase().includes(q) ||
      c.courseName.toLowerCase().includes(q) ||
      c.faculty.toLowerCase().includes(q) ||
      c.program.toLowerCase().includes(q)
    );
  }, [courseSimulationData, simSearchQuery]);

  // Course Moderation Manager enriched dataset
  const moderationCoursesList = useMemo(() => {
    return distinctCourses.map(c => {
      const currentMod = courseModerationMap[c.normCode] || 0;
      const courseRows = processedRows.filter(r => normalizeKey(r["Course Code"]) === c.normCode);
      const rescuedInCourse = courseRows.filter(r => r._isModeratedPass).length;
      const componentTypeStr = c.isPrOnly 
        ? "ESE-PR Only" 
        : c.hasEsePr 
          ? "ESE-TH + PR" 
          : "ESE-TH Only";

      return {
        ...c,
        currentMod,
        rescuedInCourse,
        componentTypeStr
      };
    });
  }, [distinctCourses, courseModerationMap, processedRows]);

  // Counts for filter pills in Course Moderation tab
  const modFilterCounts = useMemo(() => {
    const total = moderationCoursesList.length;
    let thCount = 0;
    let prOnlyCount = 0;
    let failedCount = 0;
    let nearPassCount = 0;
    let activeModCount = 0;
    let rescuedCount = 0;

    moderationCoursesList.forEach(c => {
      if (c.hasEseTh) thCount++;
      if (c.isPrOnly) prOnlyCount++;
      if (c.rawFailed > 0) failedCount++;
      if (c.nearPassCount > 0) nearPassCount++;
      if (c.currentMod > 0) activeModCount++;
      if (c.rescuedInCourse > 0) rescuedCount++;
    });

    return {
      total,
      thCount,
      prOnlyCount,
      failedCount,
      nearPassCount,
      activeModCount,
      rescuedCount
    };
  }, [moderationCoursesList]);

  // Filtered & Sorted Course Moderation List
  const filteredAndSortedModCourses = useMemo(() => {
    let list = moderationCoursesList;

    // Component Type filter
    if (modFilterType === "TH") {
      list = list.filter(c => c.hasEseTh);
    } else if (modFilterType === "PR_ONLY") {
      list = list.filter(c => c.isPrOnly);
    }

    // Status filter
    if (modFilterStatus === "FAILED") {
      list = list.filter(c => c.rawFailed > 0);
    } else if (modFilterStatus === "NEAR_PASS") {
      list = list.filter(c => c.nearPassCount > 0);
    } else if (modFilterStatus === "ACTIVE_MOD") {
      list = list.filter(c => c.currentMod > 0);
    } else if (modFilterStatus === "RESCUED") {
      list = list.filter(c => c.rescuedInCourse > 0);
    }

    // Search query
    if (moderationSearch.trim()) {
      const q = moderationSearch.toLowerCase().trim();
      list = list.filter(c => 
        (c.courseCode && c.courseCode.toLowerCase().includes(q)) ||
        (c.courseName && c.courseName.toLowerCase().includes(q)) ||
        (c.program && c.program.toLowerCase().includes(q)) ||
        (c.faculty && c.faculty.toLowerCase().includes(q))
      );
    }

    // Sort
    if (modSortConfig.column && modSortConfig.direction) {
      const col = modSortConfig.column;
      const dir = modSortConfig.direction === "asc" ? 1 : -1;

      list = [...list].sort((a, b) => {
        let valA, valB;
        switch (col) {
          case "courseCode":
            valA = a.courseCode || "";
            valB = b.courseCode || "";
            return dir * String(valA).localeCompare(String(valB), undefined, { numeric: true });
          case "courseName":
            valA = a.courseName || "";
            valB = b.courseName || "";
            return dir * String(valA).localeCompare(String(valB));
          case "componentType":
            valA = a.componentTypeStr || "";
            valB = b.componentTypeStr || "";
            return dir * String(valA).localeCompare(String(valB));
          case "totalStudents":
            valA = a.totalStudents || 0;
            valB = b.totalStudents || 0;
            return dir * (valA - valB);
          case "rawFailed":
            valA = a.rawFailed || 0;
            valB = b.rawFailed || 0;
            return dir * (valA - valB);
          case "nearPassCount":
            valA = a.nearPassCount || 0;
            valB = b.nearPassCount || 0;
            return dir * (valA - valB);
          case "currentMod":
            valA = a.currentMod || 0;
            valB = b.currentMod || 0;
            return dir * (valA - valB);
          case "rescuedInCourse":
            valA = a.rescuedInCourse || 0;
            valB = b.rescuedInCourse || 0;
            return dir * (valA - valB);
          default:
            return 0;
        }
      });
    }

    return list;
  }, [moderationCoursesList, modFilterType, modFilterStatus, moderationSearch, modSortConfig]);

  // Handler for sorting moderation columns
  const handleModSort = (column) => {
    setModSortConfig(prev => {
      if (prev.column === column) {
        return { column, direction: prev.direction === "asc" ? "desc" : "asc" };
      }
      const defaultDir = ["totalStudents", "rawFailed", "nearPassCount", "currentMod", "rescuedInCourse"].includes(column) ? "desc" : "asc";
      return { column, direction: defaultDir };
    });
  };

  // Summary Totals for the currently filtered courses in the Moderation Manager table
  const modTableTotals = useMemo(() => {
    let totalStudents = 0;
    let rawFailed = 0;
    let nearPassCount = 0;
    let rescuedCount = 0;

    filteredAndSortedModCourses.forEach(c => {
      totalStudents += c.totalStudents;
      rawFailed += c.rawFailed;
      nearPassCount += c.nearPassCount;
      rescuedCount += c.rescuedInCourse;
    });

    return {
      totalStudents,
      rawFailed,
      nearPassCount,
      rescuedCount
    };
  }, [filteredAndSortedModCourses]);

  // Overall Simulation Totals across filtered courses
  const simTotals = useMemo(() => {
    let totalStudents = 0;
    let totalAbsent = 0;
    let totalMalpractice = 0;
    let totalHeldback = 0;
    let totalHeld = 0;
    let rawEsePass = 0;
    let rawOverallPass = 0;
    let rawPass = 0;
    const modPass = Array(11).fill(0);
    const eseModPass = Array(11).fill(0);
    const overallModPass = Array(11).fill(0);

    filteredSimulationCourses.forEach(c => {
      totalStudents += c.totalStudents;
      totalAbsent += (c.absentCount || 0);
      totalMalpractice += (c.malpracticeCount || 0);
      totalHeldback += (c.heldbackCount || 0);
      totalHeld += (c.heldCount || 0);
      rawEsePass += c.rawEsePassCount;
      rawOverallPass += c.rawOverallPassCount;
      rawPass += c.rawPassCount;
      for (let m = 0; m <= 10; m++) {
        modPass[m] += c.passCountAtMod[m];
        eseModPass[m] += c.esePassAtMod[m];
        overallModPass[m] += c.overallPassAtMod[m];
      }
    });

    return {
      totalCourses: filteredSimulationCourses.length,
      totalStudents,
      totalAbsent,
      totalMalpractice,
      totalHeldback,
      totalHeld,
      rawEsePass,
      rawEsePassPct: totalStudents > 0 ? ((rawEsePass / totalStudents) * 100).toFixed(1) : "0.0",
      rawOverallPass,
      rawOverallPassPct: totalStudents > 0 ? ((rawOverallPass / totalStudents) * 100).toFixed(1) : "0.0",
      rawPass,
      rawPassPct: totalStudents > 0 ? ((rawPass / totalStudents) * 100).toFixed(1) : "0.0",
      modPass,
      modPassPct: (m) => totalStudents > 0 ? ((modPass[m] / totalStudents) * 100).toFixed(1) : "0.0",
      rescuedAtMod: (m) => modPass[m] - rawPass,
      eseModPass,
      overallModPass
    };
  }, [filteredSimulationCourses]);

  // Student-Level Semester Pass/Fail Evaluation
  // Rule: A student is "Pass" in the Semester ONLY IF they pass ALL attempted papers in that semester/term.
  // If any paper is Held (missing component), the student's semester result is Held.
  const studentSemesterData = useMemo(() => {
    const map = new Map();

    processedRows.forEach(row => {
      const prn = String(row["PRN"] || "").trim();
      const seat = String(row["Seat Number"] || "").trim();
      const program = String(row["Program Term Name"] || "").trim();
      const faculty = String(row["Faculty"] || "").trim();
      const college = String(row._college || row["College Name"] || row["College Code"] || row["College"] || "").trim();
      const collegeCode = String(row._collegeCode || row["College Code"] || "").trim();
      const collegeName = String(row._collegeName || row["College Name"] || row["College"] || "").trim();
      
      const studentId = prn || seat || "UNKNOWN";
      const key = `${studentId}__${program}`;

      if (!map.has(key)) {
        map.set(key, {
          key,
          studentId,
          prn,
          seatNumber: seat,
          faculty,
          program,
          college,
          collegeCode,
          collegeName,
          totalCourses: 0,
          rawPassedCourses: 0,
          rawFailedCourses: 0,
          finalPassedCourses: 0,
          finalFailedCourses: 0,
          heldCourses: 0,
          heldbackCourses: 0,
          heldbackReason: "",
          totalModerationMarks: 0,
          courses: []
        });
      }

      const st = map.get(key);
      st.totalCourses++;

      const isRawCoursePass = row._rawPass;
      const isFinalCoursePass = row["Course Pass/Fail"] === "Pass";
      const modMarks = row["Moderation Marks"] || 0;

      if (row._isHeldback) {
        st.heldbackCourses = (st.heldbackCourses || 0) + 1;
        st.heldCourses = (st.heldCourses || 0) + 1;
        if (!st.heldbackReason && row._heldbackReason) {
          st.heldbackReason = row._heldbackReason;
        }
      } else if (row._isHeld) {
        st.heldCourses = (st.heldCourses || 0) + 1;
      }

      if (isRawCoursePass) {
        st.rawPassedCourses++;
      } else {
        st.rawFailedCourses++;
      }

      if (isFinalCoursePass) {
        st.finalPassedCourses++;
      } else {
        st.finalFailedCourses++;
      }

      if (row._isMalpractice) {
        st.malpracticeCourses = (st.malpracticeCourses || 0) + 1;
      }

      if (row._isAbsent) {
        st.absentCourses = (st.absentCourses || 0) + 1;
      }

      st.totalModerationMarks += modMarks;

      st.courses.push({
        courseCode: row["Course Code"],
        courseName: row["Course Name"],
        esePass: row["ESE Pass"],
        overallPass: row["Overall pass"],
        coursePass: row["Course Pass/Fail"],
        rawPass: isRawCoursePass,
        isModeratedPass: row._isModeratedPass,
        isAbsent: !!row._isAbsent,
        absentStatus: row._absentStatus || "",
        isMalpractice: !!row._isMalpractice,
        malpracticeStatus: row._malpracticeStatus || "",
        malpracticeRemarks: row._malpracticeRemarks || "",
        malpracticeDate: row._malpracticeDate || "",
        isHeld: !!row._isHeld,
        isHeldback: !!row._isHeldback,
        heldbackReason: row._heldbackReason || "",
        missingComponents: row._missingComponents || [],
        modMarks,
        eseOverall: row["ESE Overall"],
        eseMin: row["ESE - Min"],
        courseOverall: row["Course Overall Marks "],
        overallMin: row["Overall Minimum"]
      });
    });

    const list = Array.from(map.values()).map(st => {
      const hasHeldback = (st.heldbackCourses || 0) > 0;
      const hasHeld = (st.heldCourses || 0) > 0;
      let semesterResult = "Fail";
      let rawSemesterResult = "Fail";
      let finalSemesterPass = false;
      let rawSemesterPass = false;

      if (hasHeldback) {
        semesterResult = "Held (Heldback)";
        rawSemesterResult = "Held (Heldback)";
        finalSemesterPass = false;
        rawSemesterPass = false;
      } else if (hasHeld) {
        semesterResult = "Held (Missing Component)";
        rawSemesterResult = "Held (Missing Component)";
        finalSemesterPass = false;
        rawSemesterPass = false;
      } else {
        rawSemesterPass = st.rawFailedCourses === 0;
        finalSemesterPass = st.finalFailedCourses === 0;
        semesterResult = finalSemesterPass ? "Pass" : "Fail";
        rawSemesterResult = rawSemesterPass ? "Pass" : "Fail";
      }

      const isRescuedSemester = !hasHeld && !rawSemesterPass && finalSemesterPass;

      return {
        ...st,
        isHeld: hasHeld,
        isHeldback: hasHeldback,
        rawSemesterPass,
        finalSemesterPass,
        isRescuedSemester,
        semesterResult,
        rawSemesterResult
      };
    });

    return list.sort((a, b) => {
      if (a.seatNumber && b.seatNumber) {
        return String(a.seatNumber).localeCompare(String(b.seatNumber), undefined, { numeric: true });
      }
      return String(a.prn).localeCompare(String(b.prn), undefined, { numeric: true });
    });
  }, [processedRows]);

  // Context-Scoped Student Data (College, Program, Course, Search)
  const scopedStudents = useMemo(() => {
    let list = studentSemesterData;

    // 1. College filter
    if (studentCollegeFilter !== "ALL") {
      list = list.filter(st => st.college === studentCollegeFilter);
    }

    // 2. Program filter
    if (studentProgramFilter !== "ALL") {
      list = list.filter(st => st.program === studentProgramFilter);
    }

    // 3. Course filter
    if (studentCourseFilter !== "ALL") {
      const cleanFilter = cleanCourseCode(studentCourseFilter);
      list = list.filter(st => st.courses.some(c => cleanCourseCode(c.courseCode) === cleanFilter));
    }

    // 4. Search query
    if (studentSearchQuery.trim()) {
      const q = studentSearchQuery.toLowerCase().trim();
      list = list.filter(st => 
        st.prn.toLowerCase().includes(q) ||
        st.seatNumber.toLowerCase().includes(q) ||
        (st.college && st.college.toLowerCase().includes(q)) ||
        st.program.toLowerCase().includes(q) ||
        st.faculty.toLowerCase().includes(q) ||
        (q === "held" && st.isHeld) ||
        (q === "heldback" && st.isHeldback) ||
        (st.heldbackReason && st.heldbackReason.toLowerCase().includes(q)) ||
        st.courses.some(c => (c.courseCode && c.courseCode.toLowerCase().includes(q)) || (c.courseName && c.courseName.toLowerCase().includes(q)))
      );
    }

    return list;
  }, [studentSemesterData, studentCollegeFilter, studentProgramFilter, studentCourseFilter, studentSearchQuery]);

  // Consolidated / Global Overall Student Metrics (Unfiltered, for sidebar)
  const consolidatedStudentMetrics = useMemo(() => {
    const total = studentSemesterData.length;
    if (total === 0) {
      return {
        totalStudents: 0,
        rawPassedStudents: 0,
        rawPassedPct: "0.0",
        finalPassedStudents: 0,
        finalPassedPct: "0.0",
        failedStudents: 0,
        failedPct: "0.0",
        heldStudents: 0,
        heldPct: "0.0",
        heldbackStudents: 0,
        heldbackPct: "0.0",
        heldMissingStudents: 0,
        heldMissingPct: "0.0",
        rescuedStudents: 0,
        rescuedPct: "0.0",
        absentStudents: 0,
        absentPct: "0.0",
        malpracticeStudents: 0,
        malpracticePct: "0.0",
        totalPapersAttempted: 0,
        avgPapersPerStudent: "0.0"
      };
    }

    let rawPassedStudents = 0;
    let finalPassedStudents = 0;
    let rescuedStudents = 0;
    let absentStudents = 0;
    let malpracticeStudents = 0;
    let heldStudents = 0;
    let heldbackStudents = 0;
    let heldMissingStudents = 0;
    let totalPapersAttempted = 0;

    studentSemesterData.forEach(st => {
      totalPapersAttempted += st.totalCourses;
      if (st.isHeldback) {
        heldbackStudents++;
        heldStudents++;
      } else if (st.isHeld) {
        heldMissingStudents++;
        heldStudents++;
      } else {
        if (st.rawSemesterPass) rawPassedStudents++;
        if (st.finalSemesterPass) finalPassedStudents++;
        if (st.isRescuedSemester) rescuedStudents++;
      }
      if ((st.absentCourses || 0) > 0) absentStudents++;
      if ((st.malpracticeCourses || 0) > 0) malpracticeStudents++;
    });

    const failedStudents = total - finalPassedStudents - heldStudents;

    return {
      totalStudents: total,
      rawPassedStudents,
      rawPassedPct: total > 0 ? ((rawPassedStudents / total) * 100).toFixed(1) : "0.0",
      finalPassedStudents,
      finalPassedPct: total > 0 ? ((finalPassedStudents / total) * 100).toFixed(1) : "0.0",
      failedStudents,
      failedPct: total > 0 ? ((failedStudents / total) * 100).toFixed(1) : "0.0",
      heldStudents,
      heldPct: total > 0 ? ((heldStudents / total) * 100).toFixed(1) : "0.0",
      heldbackStudents,
      heldbackPct: total > 0 ? ((heldbackStudents / total) * 100).toFixed(1) : "0.0",
      heldMissingStudents,
      heldMissingPct: total > 0 ? ((heldMissingStudents / total) * 100).toFixed(1) : "0.0",
      rescuedStudents,
      rescuedPct: total > 0 ? ((rescuedStudents / total) * 100).toFixed(1) : "0.0",
      absentStudents,
      absentPct: total > 0 ? ((absentStudents / total) * 100).toFixed(1) : "0.0",
      malpracticeStudents,
      malpracticePct: total > 0 ? ((malpracticeStudents / total) * 100).toFixed(1) : "0.0",
      totalPapersAttempted,
      avgPapersPerStudent: total > 0 ? (totalPapersAttempted / total).toFixed(1) : "0.0"
    };
  }, [studentSemesterData]);

  // Student Metrics calculated directly from the active scoped filters (for main student view)
  const studentMetrics = useMemo(() => {
    const total = scopedStudents.length;
    if (total === 0) {
      return {
        totalStudents: 0,
        rawPassedStudents: 0,
        rawPassedPct: "0.0",
        finalPassedStudents: 0,
        finalPassedPct: "0.0",
        failedStudents: 0,
        failedPct: "0.0",
        heldStudents: 0,
        heldPct: "0.0",
        heldbackStudents: 0,
        heldbackPct: "0.0",
        heldMissingStudents: 0,
        heldMissingPct: "0.0",
        rescuedStudents: 0,
        rescuedPct: "0.0",
        absentStudents: 0,
        absentPct: "0.0",
        malpracticeStudents: 0,
        malpracticePct: "0.0",
        totalPapersAttempted: 0,
        avgPapersPerStudent: "0.0"
      };
    }

    let rawPassedStudents = 0;
    let finalPassedStudents = 0;
    let rescuedStudents = 0;
    let absentStudents = 0;
    let malpracticeStudents = 0;
    let heldStudents = 0;
    let heldbackStudents = 0;
    let heldMissingStudents = 0;
    let totalPapersAttempted = 0;

    scopedStudents.forEach(st => {
      totalPapersAttempted += st.totalCourses;
      if (st.isHeldback) {
        heldbackStudents++;
        heldStudents++;
      } else if (st.isHeld) {
        heldMissingStudents++;
        heldStudents++;
      } else {
        if (st.rawSemesterPass) rawPassedStudents++;
        if (st.finalSemesterPass) finalPassedStudents++;
        if (st.isRescuedSemester) rescuedStudents++;
      }
      if ((st.absentCourses || 0) > 0) absentStudents++;
      if ((st.malpracticeCourses || 0) > 0) malpracticeStudents++;
    });

    const failedStudents = total - finalPassedStudents - heldStudents;

    return {
      totalStudents: total,
      rawPassedStudents,
      rawPassedPct: total > 0 ? ((rawPassedStudents / total) * 100).toFixed(1) : "0.0",
      finalPassedStudents,
      finalPassedPct: total > 0 ? ((finalPassedStudents / total) * 100).toFixed(1) : "0.0",
      failedStudents,
      failedPct: total > 0 ? ((failedStudents / total) * 100).toFixed(1) : "0.0",
      heldStudents,
      heldPct: total > 0 ? ((heldStudents / total) * 100).toFixed(1) : "0.0",
      heldbackStudents,
      heldbackPct: total > 0 ? ((heldbackStudents / total) * 100).toFixed(1) : "0.0",
      heldMissingStudents,
      heldMissingPct: total > 0 ? ((heldMissingStudents / total) * 100).toFixed(1) : "0.0",
      rescuedStudents,
      rescuedPct: total > 0 ? ((rescuedStudents / total) * 100).toFixed(1) : "0.0",
      absentStudents,
      absentPct: total > 0 ? ((absentStudents / total) * 100).toFixed(1) : "0.0",
      malpracticeStudents,
      malpracticePct: total > 0 ? ((malpracticeStudents / total) * 100).toFixed(1) : "0.0",
      totalPapersAttempted,
      avgPapersPerStudent: total > 0 ? (totalPapersAttempted / total).toFixed(1) : "0.0"
    };
  }, [scopedStudents]);

  // Contextual Unique Lists for Student Dropdown Filters
  const uniqueStudentColleges = useMemo(() => {
    const set = new Set();
    studentSemesterData.forEach(st => {
      if (
        (studentProgramFilter === "ALL" || st.program === studentProgramFilter) &&
        (studentCourseFilter === "ALL" || st.courses.some(c => c.courseCode === studentCourseFilter))
      ) {
        if (st.college) set.add(st.college);
      }
    });
    return Array.from(set).sort();
  }, [studentSemesterData, studentProgramFilter, studentCourseFilter]);

  const uniqueStudentPrograms = useMemo(() => {
    const set = new Set();
    studentSemesterData.forEach(st => {
      if (
        (studentCollegeFilter === "ALL" || st.college === studentCollegeFilter) &&
        (studentCourseFilter === "ALL" || st.courses.some(c => c.courseCode === studentCourseFilter))
      ) {
        if (st.program) set.add(st.program);
      }
    });
    return Array.from(set).sort();
  }, [studentSemesterData, studentCollegeFilter, studentCourseFilter]);

  const uniqueStudentCourses = useMemo(() => {
    const courseMap = new Map();
    studentSemesterData.forEach(st => {
      if (
        (studentCollegeFilter === "ALL" || st.college === studentCollegeFilter) &&
        (studentProgramFilter === "ALL" || st.program === studentProgramFilter)
      ) {
        st.courses.forEach(c => {
          const code = cleanCourseCode(c.courseCode);
          const name = cleanCourseName(c.courseName);
          const norm = normalizeKey(code);
          if (norm && !courseMap.has(norm)) {
            courseMap.set(norm, { code, label: name ? `${code} - ${name}` : code });
          }
        });
      }
    });
    return Array.from(courseMap.values()).sort((a, b) => a.code.localeCompare(b.code));
  }, [studentSemesterData, studentCollegeFilter, studentProgramFilter]);

  // Final Filtered Student List (Filtered by Status)
  const filteredStudents = useMemo(() => {
    let list = scopedStudents;

    if (studentFilterStatus === "PASS") {
      list = list.filter(st => st.finalSemesterPass);
    } else if (studentFilterStatus === "FAIL") {
      list = list.filter(st => !st.finalSemesterPass && !st.isHeld);
    } else if (studentFilterStatus === "HELDBACK") {
      list = list.filter(st => st.isHeldback);
    } else if (studentFilterStatus === "HELD") {
      list = list.filter(st => st.isHeld);
    } else if (studentFilterStatus === "RESCUED") {
      list = list.filter(st => st.isRescuedSemester);
    } else if (studentFilterStatus === "ABSENT") {
      list = list.filter(st => (st.absentCourses || 0) > 0);
    } else if (studentFilterStatus === "MALPRACTICE") {
      list = list.filter(st => (st.malpracticeCourses || 0) > 0);
    }

    return list;
  }, [scopedStudents, studentFilterStatus]);

  const toggleStudentExpand = (key) => {
    setExpandedStudents(prev => ({
      ...prev,
      [key]: !prev[key]
    }));
  };

  const scoreHeaderRow = (rowArray) => {
    if (!Array.isArray(rowArray)) return { score: 0, matchedHeaders: [] };
    let score = 0;
    const matchedHeaders = [];
    const normCols = rowArray.map(c => normalizeKey(c));

    // Raw assessment indicators
    if (normCols.some(c => c.includes("assessmentmethod") || c === "am" || c === "method")) {
      score += 12;
      matchedHeaders.push("Assessment Method");
    }
    if (normCols.some(c => c.includes("assessmenttype") || c === "at" || c === "type")) {
      score += 12;
      matchedHeaders.push("Assessment Type");
    }
    if (normCols.some(c => c.includes("atmaxmarks") || c.includes("maxmarks") || c.includes("maxmark"))) {
      score += 8;
      matchedHeaders.push("AT Max Marks");
    }
    if (normCols.some(c => c === "marks" || c.includes("obtained") || c === "mark")) {
      score += 8;
      matchedHeaders.push("Marks");
    }

    // Pre-aggregated indicators
    if (normCols.some(c => c.includes("eseoverall") || c.includes("esemax") || c.includes("esethmax") || c.includes("esethobtained"))) {
      score += 12;
      matchedHeaders.push("ESE Component");
    }
    if (normCols.some(c => c.includes("ceoverall") || c.includes("cemax") || c.includes("cethmax") || c.includes("cethobtained"))) {
      score += 12;
      matchedHeaders.push("CE Component");
    }
    if (normCols.some(c => c.includes("courseoverall") || c.includes("overallmax") || c.includes("overallmin"))) {
      score += 10;
      matchedHeaders.push("Course Overall");
    }

    // Common identifiers
    if (normCols.some(c => c.includes("coursecode") || c.includes("papercode") || c.includes("subjectcode"))) {
      score += 6;
      matchedHeaders.push("Course Code");
    }
    if (normCols.some(c => c.includes("coursename") || c.includes("papername") || c.includes("subjectname"))) {
      score += 6;
      matchedHeaders.push("Course Name");
    }
    if (normCols.some(c => c.includes("prn") || c.includes("registerno") || c.includes("studentid"))) {
      score += 6;
      matchedHeaders.push("PRN");
    }
    if (normCols.some(c => c.includes("seat") || c.includes("rollno"))) {
      score += 4;
      matchedHeaders.push("Seat Number");
    }
    if (normCols.some(c => c.includes("program") || c.includes("semester") || c.includes("term"))) {
      score += 4;
      matchedHeaders.push("Program Term");
    }
    if (normCols.some(c => c.includes("faculty") || c.includes("department"))) {
      score += 4;
      matchedHeaders.push("Faculty");
    }
    if (normCols.some(c => c.includes("college") || c.includes("adec") || c.includes("institute") || c.includes("center"))) {
      score += 4;
      matchedHeaders.push("College / ADEC");
    }

    // Heldback indicators
    if (normCols.some(c => c.includes("heldback") || c.includes("reason"))) {
      score += 8;
      matchedHeaders.push("Heldback / Reason");
    }
    if (normCols.some(c => c === "paper" || c.includes("paper"))) {
      score += 6;
      matchedHeaders.push("Paper");
    }

    return { score, matchedHeaders };
  };

  const parseSheetWithHeaderScan = (ws) => {
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (!aoa || aoa.length === 0) return { rows: [], headerMap: {}, headerRowIdx: 0, score: 0, matchedHeaders: [] };

    let bestRowIdx = 0;
    let maxScore = -1;
    let bestMatchedHeaders = [];

    for (let r = 0; r < Math.min(aoa.length, 15); r++) {
      const rowArr = aoa[r];
      if (!Array.isArray(rowArr) || rowArr.length === 0) continue;
      const { score, matchedHeaders } = scoreHeaderRow(rowArr);
      if (score > maxScore) {
        maxScore = score;
        bestRowIdx = r;
        bestMatchedHeaders = matchedHeaders;
      }
    }

    const rawHeaders = aoa[bestRowIdx] || [];
    const headerMap = {};
    const validHeaderIndices = [];

    rawHeaders.forEach((colName, idx) => {
      const name = String(colName || "").trim();
      if (name) {
        const norm = normalizeKey(name);
        headerMap[norm] = name;
        validHeaderIndices.push({ idx, name, norm });
      }
    });

    const rows = [];
    for (let r = bestRowIdx + 1; r < aoa.length; r++) {
      const rowArr = aoa[r];
      if (!Array.isArray(rowArr) || rowArr.every(c => String(c || "").trim() === "")) continue;
      const rowObj = {};
      validHeaderIndices.forEach(({ idx, name }) => {
        rowObj[name] = rowArr[idx] !== undefined ? rowArr[idx] : "";
      });
      rows.push(rowObj);
    }

    return { rows, headerMap, headerRowIdx: bestRowIdx, score: maxScore, matchedHeaders: bestMatchedHeaders };
  };

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setStatus("Analyzing workbook " + file.name + "...", "info");
    setSourceFile(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        setWorkbook(wb);
        setSheetNames(wb.SheetNames);

        const meta = {};
        let bestSheet = wb.SheetNames[0];
        let highestTotalScore = -1;

        wb.SheetNames.forEach((sName) => {
          const ws = wb.Sheets[sName];
          const { score, matchedHeaders, headerRowIdx, rows } = parseSheetWithHeaderScan(ws);
          
          let nameBonus = 0;
          const sNorm = normalizeKey(sName);
          if (sNorm.includes("sourcefile") || sNorm.includes("source") || sNorm.includes("raw")) nameBonus += 10;
          else if (sNorm.includes("result") || sNorm.includes("marks") || sNorm.includes("exam") || sNorm.includes("output")) nameBonus += 6;

          const totalScore = score + nameBonus + (rows.length > 0 ? 4 : 0);
          meta[sName] = { score: totalScore, matchedHeaders, headerRowIdx, rowCount: rows.length };

          if (totalScore > highestTotalScore) {
            highestTotalScore = totalScore;
            bestSheet = sName;
          }
        });

        setSheetMetadata(meta);
        setSelectedSheet(bestSheet);

        const ws = wb.Sheets[bestSheet];
        const { rows, headerMap: hMap, headerRowIdx, matchedHeaders } = parseSheetWithHeaderScan(ws);

        if (!rows || rows.length === 0) {
          setStatus("No valid data rows found in sheet " + bestSheet + ".", "warning");
          setIsProcessing(false);
          return;
        }

        setHeaderMap(hMap);
        setRawRows(rows);

        const baseGrouped = buildGroupedRecordsFromRows(rows, hMap, absentRecordsMap, malpracticeRecordsMap, heldbackRecordsMap);
        setGroupedRecords(baseGrouped);
        setPage(0);

        const headerInfo = matchedHeaders.length > 0 ? " (Detected Headers: " + matchedHeaders.join(", ") + " on Row " + (headerRowIdx + 1) + ")" : "";
        setStatus("Successfully loaded " + baseGrouped.length + " courses from \"" + bestSheet + "\"!" + headerInfo, "success");
      } catch (err) {
        console.error("Error parsing sheet:", err);
        setStatus("Failed to read file: " + err.message, "error");
      } finally {
        setIsProcessing(false);
      }
    };

    reader.onerror = () => {
      setStatus("File reading error.", "error");
      setIsProcessing(false);
    };

    reader.readAsArrayBuffer(file);
  };

  const handleSheetChange = (sheetName) => {
    if (!workbook) return;
    setSelectedSheet(sheetName);
    setIsProcessing(true);
    setStatus("Loading sheet " + sheetName + "...", "info");

    try {
      const ws = workbook.Sheets[sheetName];
      const { rows, headerMap: hMap, headerRowIdx, matchedHeaders } = parseSheetWithHeaderScan(ws);

      if (!rows || rows.length === 0) {
        setStatus("Sheet " + sheetName + " is empty or has no valid rows.", "warning");
        setRawRows([]);
        setGroupedRecords([]);
        setIsProcessing(false);
        return;
      }

      setHeaderMap(hMap);
      setRawRows(rows);

      const baseGrouped = buildGroupedRecordsFromRows(rows, hMap, absentRecordsMap, malpracticeRecordsMap, heldbackRecordsMap);
      setGroupedRecords(baseGrouped);
      setPage(0);

      const headerInfo = matchedHeaders.length > 0 ? " (Headers: " + matchedHeaders.join(", ") + " on Row " + (headerRowIdx + 1) + ")" : "";
      setStatus("Loaded \"" + sheetName + "\": " + baseGrouped.length + " calculated course result rows." + headerInfo, "success");
    } catch (err) {
      console.error(err);
      setStatus("Failed to load sheet: " + err.message, "error");
    } finally {
      setIsProcessing(false);
    }
  };

  // Moderation Handlers: UI & Excel Ingestion
  const updateCourseModeration = (normCode, value) => {
    const num = Math.max(0, parseInt(value, 10) || 0);
    setCourseModerationMap(prev => ({
      ...prev,
      [normCode]: num
    }));
  };

  const applyBulkModeration = (marks) => {
    const val = Math.max(0, parseInt(marks, 10) || 0);
    const updated = {};
    distinctCourses.forEach(c => {
      updated[c.normCode] = val;
    });
    setCourseModerationMap(updated);
    setStatus("Applied " + val + " marks moderation to all " + distinctCourses.length + " courses.", "success");
  };

  const resetAllModeration = () => {
    setCourseModerationMap({});
    setStatus("Reset all course moderation marks to 0.", "info");
  };

  // Upload Moderation Excel File
  const handleModerationExcelUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: "array" });
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        const jsonRows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });

        if (!jsonRows || jsonRows.length === 0) {
          setStatus("Uploaded moderation sheet is empty.", "warning");
          return;
        }

        const newMap = { ...courseModerationMap };
        let countUpdated = 0;

        jsonRows.forEach(row => {
          let courseCode = "";
          let modMarks = 0;

          for (const [key, val] of Object.entries(row)) {
            const kNorm = normalizeKey(key);
            if (kNorm.includes("coursecode") || kNorm.includes("papercode") || kNorm.includes("subjectcode") || kNorm === "code") {
              courseCode = cleanCourseCode(val);
            } else if (kNorm.includes("currentmoderationmarks") || kNorm.includes("moderation") || kNorm.includes("modmarks") || kNorm.includes("gracemarks") || kNorm === "marks") {
              const parsed = parseInt(val, 10);
              if (!isNaN(parsed)) modMarks = Math.max(0, parsed);
            }
          }

          if (courseCode) {
            const norm = normalizeKey(courseCode);
            newMap[norm] = modMarks;
            countUpdated++;
          }
        });

        setCourseModerationMap(newMap);
        setStatus("Successfully imported moderation marks for " + countUpdated + " courses from " + file.name + "!", "success");
      } catch (err) {
        console.error("Error importing moderation sheet:", err);
        setStatus("Failed to read moderation sheet: " + err.message, "error");
      }
    };

    reader.readAsArrayBuffer(file);
    if (modFileInputRef.current) modFileInputRef.current.value = "";
  };

  // Download Pre-filled Moderation Template Excel (Course Code, Current Moderation Marks)
  const handleDownloadModerationTemplate = () => {
    if (distinctCourses.length === 0) {
      alert("Please upload a source marks sheet first to extract unique course codes.");
      return;
    }

    const templateData = [
      ["Course Code", "Current Moderation Marks"],
      ...distinctCourses.map(c => [
        c.courseCode,
        courseModerationMap[c.normCode] || 0
      ])
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(templateData);
    ws["!cols"] = [{ wch: 22 }, { wch: 28 }];
    XLSX.utils.book_append_sheet(wb, ws, "Course_Moderation_Template");

    XLSX.writeFile(wb, "course_moderation_template.xlsx");
    setStatus("Downloaded pre-filled Course Moderation Template (.xlsx) with Course Code & Current Moderation Marks.", "success");
  };

  // Helper to read multiple Excel files and extract all rows across all sheets
  const readMultipleExcelFiles = async (fileList) => {
    const files = Array.from(fileList || []);
    let allRows = [];
    const parsedFileNames = [];

    for (const file of files) {
      try {
        const buffer = await file.arrayBuffer();
        const data = new Uint8Array(buffer);
        const wb = XLSX.read(data, { type: "array" });
        let fileRowsCount = 0;

        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          if (!ws || !ws["!ref"]) continue;
          const { rows } = parseSheetWithHeaderScan(ws);
          const effectiveRows = (rows && rows.length > 0) ? rows : XLSX.utils.sheet_to_json(ws, { defval: "" });
          if (effectiveRows && effectiveRows.length > 0) {
            allRows = allRows.concat(effectiveRows);
            fileRowsCount += effectiveRows.length;
          }
        }

        if (fileRowsCount > 0) {
          parsedFileNames.push(file.name);
        }
      } catch (err) {
        console.error(`Error reading file ${file.name}:`, err);
      }
    }

    return { allRows, parsedFileNames };
  };

  // Handle Upload of Absent Students Report Excel (.xlsx, .xls, .csv) - Supports single or multi-file uploads
  const handleAbsentExcelUpload = async (e) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    try {
      const { allRows, parsedFileNames } = await readMultipleExcelFiles(files);

      if (allRows.length === 0) {
        setStatus("Uploaded absent report sheet(s) are empty.", "warning");
        if (absentFileInputRef.current) absentFileInputRef.current.value = "";
        return;
      }

      const { map, list } = buildAbsentLookup(allRows);

      if (list.length === 0) {
        setStatus("No valid student absent entries found across uploaded file(s). Ensure columns include PRN/Seat Number and Course Code.", "warning");
        if (absentFileInputRef.current) absentFileInputRef.current.value = "";
        return;
      }

      const fileDisplayName = parsedFileNames.length === 1 
        ? parsedFileNames[0] 
        : `${parsedFileNames.length} files (${parsedFileNames.join(", ")})`;

      setAbsentRecordsMap(map);
      setAbsentList(list);
      setAbsentFileName(fileDisplayName);

      // If source records are already loaded, re-group them immediately with this absent map!
      if (rawRows && rawRows.length > 0 && headerMap) {
        const updatedGrouped = buildGroupedRecordsFromRows(rawRows, headerMap, map, malpracticeRecordsMap, heldbackRecordsMap);
        setGroupedRecords(updatedGrouped);
      }

      const fileDetails = parsedFileNames.length === 1 ? `"${parsedFileNames[0]}"` : `${parsedFileNames.length} programme file(s) [${parsedFileNames.join(", ")}]`;
      setStatus(`Successfully imported ${list.length} absent record(s) from ${fileDetails}. Matching components are marked as "Absent (Ab)" with Fail status.`, "success");
    } catch (err) {
      console.error("Error importing absent sheet(s):", err);
      setStatus("Failed to read absent report: " + err.message, "error");
    } finally {
      if (absentFileInputRef.current) absentFileInputRef.current.value = "";
    }
  };

  // Clear loaded Absent Data and revert source marks
  const handleClearAbsentData = () => {
    setAbsentRecordsMap(new Map());
    setAbsentList([]);
    setAbsentFileName("");

    if (rawRows && rawRows.length > 0 && headerMap) {
      const resetGrouped = buildGroupedRecordsFromRows(rawRows, headerMap, new Map(), malpracticeRecordsMap, heldbackRecordsMap);
      setGroupedRecords(resetGrouped);
    }

    setStatus("Cleared absent records. Source marks reverted to original.", "info");
  };

  // Download Absent Report Template (.xlsx) with standard ADES format and sample rows
  const handleDownloadAbsentTemplate = () => {
    const headers = [
      "Student Name",
      "PRN",
      "Seat Number",
      "Program Name",
      "Program Branch",
      "Program Pattern",
      "Program Part Name",
      "Program Term Name",
      "Specialization",
      "College Code",
      "College Name",
      "Course Code",
      "Course Name",
      "TLM",
      "AM",
      "AT",
      "Absent Status (categorized as)"
    ];

    const sampleRows = [
      [
        "FATHIMATH SANAH",
        "2024012600144595",
        "KH24CCOR025",
        "Bachelor of Commerce(Co-operation)",
        "Commerce",
        "FYUGP-2024",
        "BCom Year II",
        "SEMESTER IV",
        "",
        "KH",
        "Khansa Women's College for Advanced Studies, Kumbla, Kasargod",
        "KU4VACCOM101",
        "Consumer Rights and Protection",
        "Lec-Lab",
        "ESE",
        "TH",
        "Marked Absent During Mark Entry"
      ],
      [
        "FATHIMATH SANAH",
        "2024012600144595",
        "KH24CCOR025",
        "Bachelor of Commerce(Co-operation)",
        "Commerce",
        "FYUGP-2024",
        "BCom Year II",
        "SEMESTER IV",
        "",
        "KH",
        "Khansa Women's College for Advanced Studies, Kumbla, Kasargod",
        "KU4SECCOM100",
        "Office Secretaryship and Practices",
        "Lec-Lab",
        "ESE",
        "TH",
        "Marked Absent During Mark Entry"
      ]
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws["!cols"] = [
      { wch: 22 }, // Student Name
      { wch: 20 }, // PRN
      { wch: 16 }, // Seat Number
      { wch: 34 }, // Program Name
      { wch: 18 }, // Program Branch
      { wch: 16 }, // Program Pattern
      { wch: 16 }, // Program Part Name
      { wch: 16 }, // Program Term Name
      { wch: 16 }, // Specialization
      { wch: 14 }, // College Code
      { wch: 38 }, // College Name
      { wch: 18 }, // Course Code
      { wch: 34 }, // Course Name
      { wch: 12 }, // TLM
      { wch: 10 }, // AM
      { wch: 10 }, // AT
      { wch: 34 }  // Absent Status
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Absent_Report_Template");
    XLSX.writeFile(wb, "absent_report_template.xlsx");
    setStatus("Downloaded Absent Report Template (.xlsx) with standard ADES columns and sample entries.", "success");
  };

  // Handle Upload of Malpractice Students Report Excel (.xlsx, .xls, .csv) - Supports single or multi-file uploads
  const handleMalpracticeExcelUpload = async (e) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    try {
      const { allRows, parsedFileNames } = await readMultipleExcelFiles(files);

      if (allRows.length === 0) {
        setStatus("Uploaded malpractice report sheet(s) are empty.", "warning");
        if (malpracticeFileInputRef.current) malpracticeFileInputRef.current.value = "";
        return;
      }

      const { map, list } = buildMalpracticeLookup(allRows);

      if (list.length === 0) {
        setStatus("No valid student malpractice entries found across uploaded file(s). Ensure columns include PRN/Seat Number and Course Code.", "warning");
        if (malpracticeFileInputRef.current) malpracticeFileInputRef.current.value = "";
        return;
      }

      const fileDisplayName = parsedFileNames.length === 1 
        ? parsedFileNames[0] 
        : `${parsedFileNames.length} files (${parsedFileNames.join(", ")})`;

      setMalpracticeRecordsMap(map);
      setMalpracticeList(list);
      setMalpracticeFileName(fileDisplayName);

      // If source records are already loaded, re-group them immediately with this malpractice map!
      if (rawRows && rawRows.length > 0 && headerMap) {
        const updatedGrouped = buildGroupedRecordsFromRows(rawRows, headerMap, absentRecordsMap, map, heldbackRecordsMap);
        setGroupedRecords(updatedGrouped);
      }

      const fileDetails = parsedFileNames.length === 1 ? `"${parsedFileNames[0]}"` : `${parsedFileNames.length} programme file(s) [${parsedFileNames.join(", ")}]`;
      setStatus(`Successfully imported ${list.length} malpractice record(s) from ${fileDetails}. Matching components are marked as "Malpractice (MP)" with Fail status.`, "success");
    } catch (err) {
      console.error("Error importing malpractice sheet(s):", err);
      setStatus("Failed to read malpractice report: " + err.message, "error");
    } finally {
      if (malpracticeFileInputRef.current) malpracticeFileInputRef.current.value = "";
    }
  };

  // Clear loaded Malpractice Data and revert source marks
  const handleClearMalpracticeData = () => {
    setMalpracticeRecordsMap(new Map());
    setMalpracticeList([]);
    setMalpracticeFileName("");

    if (rawRows && rawRows.length > 0 && headerMap) {
      const resetGrouped = buildGroupedRecordsFromRows(rawRows, headerMap, absentRecordsMap, new Map(), heldbackRecordsMap);
      setGroupedRecords(resetGrouped);
    }

    setStatus("Cleared malpractice records. Source marks reverted to original.", "info");
  };

  // Download Malpractice Report Template (.xlsx) with standard university UM columns and sample rows
  const handleDownloadMalpracticeTemplate = () => {
    const headers = [
      "Student Name",
      "PRN",
      "Seat Number",
      "Program Name",
      "Program Branch",
      "Program Pattern",
      "Program Part Name",
      "Program Term Name",
      "Specialization",
      "College Code",
      "College Name",
      "Course Code",
      "Center Name",
      "Venue Name",
      "Course Name",
      "TLM",
      "AM",
      "AT",
      "Status (e.g., Unresolved, Resolved, SPC, CPC,NP)",
      "Description / Remarks",
      "UM Marked Date"
    ];

    const sampleRows = [
      [
        "HISHAM ABDULLA V P",
        "2024012600145181",
        "GD24CFIR018",
        "Bachelor of Commerce(Finance)",
        "Commerce",
        "FYUGP-2024",
        "BCom Year II",
        "SEMESTER IV",
        "",
        "GD",
        "Gurudev Arts and Science College, Mathil",
        "KU4VACCOM102",
        "1",
        "Gurudev Arts and Science College, Mathil",
        "Environmental Studies and Disaster Management",
        "Lec-Lab",
        "ESE",
        "TH",
        "EHB",
        "malpractice entered",
        "01/09/2026"
      ],
      [
        "FATHIMA SAFA C P",
        "2024012600141337",
        "SS24ARBR019",
        "Bachelor of Arts in Arabic",
        "Arabic",
        "FYUGP-2024",
        "BA Year II",
        "SEMESTER IV",
        "",
        "SS",
        "Sir Syed College,Taliparamba",
        "KU4DSCARB207",
        "1",
        "Sir Syed College,Taliparamba",
        "Classical and Medieval Arabic Prose",
        "Lec-Lab",
        "ESE",
        "TH",
        "EHB",
        "",
        "13/08/2026"
      ],
      [
        "RIZA AMINA K",
        "2024012600136085",
        "SS24ECOR022",
        "Bachelor of Arts in Economics",
        "Economics",
        "FYUGP-2024",
        "BA Year II",
        "SEMESTER IV",
        "",
        "SS",
        "Sir Syed College,Taliparamba",
        "KU4VACHIS203",
        "1",
        "Sir Syed College,Taliparamba",
        "Gandhian Political Ideologies and Practices",
        "Lec-Lab",
        "ESE",
        "TH",
        "EHB",
        "",
        "13/08/2026"
      ],
      [
        "YADHUKRISHNAN K K",
        "2024012600132257",
        "SE24CCOR042",
        "Bachelor of Commerce(Co-operation)",
        "Commerce",
        "FYUGP-2024",
        "BCom Year II",
        "SEMESTER IV",
        "",
        "SE",
        "SES College, Thaliparamba",
        "KU4DSCCOM207",
        "1",
        "SES College, Thaliparamba",
        "Cost Accounting",
        "Lec-Lab",
        "ESE",
        "TH",
        "EHB",
        "SMP",
        "19/08/2026"
      ]
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws["!cols"] = [
      { wch: 22 }, // Student Name
      { wch: 20 }, // PRN
      { wch: 16 }, // Seat Number
      { wch: 32 }, // Program Name
      { wch: 18 }, // Program Branch
      { wch: 16 }, // Program Pattern
      { wch: 16 }, // Program Part Name
      { wch: 16 }, // Program Term Name
      { wch: 16 }, // Specialization
      { wch: 14 }, // College Code
      { wch: 38 }, // College Name
      { wch: 18 }, // Course Code
      { wch: 14 }, // Center Name
      { wch: 38 }, // Venue Name
      { wch: 36 }, // Course Name
      { wch: 12 }, // TLM
      { wch: 10 }, // AM
      { wch: 10 }, // AT
      { wch: 36 }, // Status
      { wch: 24 }, // Description / Remarks
      { wch: 16 }  // UM Marked Date
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Malpractice_Template");
    XLSX.writeFile(wb, "malpractice_report_template.xlsx");
    setStatus("Downloaded Malpractice Report Template (.xlsx) with standard 21 UM columns and sample entries.", "success");
  };

  // Handle Upload of Heldback Students Report Excel (.xlsx, .xls, .csv) - Supports single or multi-file uploads across programmes
  const handleHeldbackExcelUpload = async (e) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    try {
      const { allRows, parsedFileNames } = await readMultipleExcelFiles(files);

      if (allRows.length === 0) {
        setStatus("Uploaded heldback report sheet(s) are empty.", "warning");
        if (heldbackFileInputRef.current) heldbackFileInputRef.current.value = "";
        return;
      }

      const { map, list } = buildHeldbackLookup(allRows);

      if (list.length === 0) {
        setStatus("No valid student heldback entries found across uploaded file(s). Ensure columns include PRN / Seat Number and Reason.", "warning");
        if (heldbackFileInputRef.current) heldbackFileInputRef.current.value = "";
        return;
      }

      const fileDisplayName = parsedFileNames.length === 1 
        ? parsedFileNames[0] 
        : `${parsedFileNames.length} files (${parsedFileNames.join(", ")})`;

      setHeldbackRecordsMap(map);
      setHeldbackList(list);
      setHeldbackFileName(fileDisplayName);

      // If source records are already loaded, re-group them immediately with this heldback map!
      if (rawRows && rawRows.length > 0 && headerMap) {
        const updatedGrouped = buildGroupedRecordsFromRows(rawRows, headerMap, absentRecordsMap, malpracticeRecordsMap, map);
        setGroupedRecords(updatedGrouped);
      }

      const fileDetails = parsedFileNames.length === 1 ? `"${parsedFileNames[0]}"` : `${parsedFileNames.length} programme file(s) [${parsedFileNames.join(", ")}]`;
      setStatus(`Successfully imported ${list.length} heldback record(s) from ${fileDetails}. Matched students are locked from calculation, moderation, and pass simulation with "Held (Heldback)" status.`, "success");
    } catch (err) {
      console.error("Error importing heldback sheet(s):", err);
      setStatus("Failed to read heldback report: " + err.message, "error");
    } finally {
      if (heldbackFileInputRef.current) heldbackFileInputRef.current.value = "";
    }
  };

  // Clear loaded Heldback Data and revert source marks
  const handleClearHeldbackData = () => {
    setHeldbackRecordsMap(new Map());
    setHeldbackList([]);
    setHeldbackFileName("");

    if (rawRows && rawRows.length > 0 && headerMap) {
      const resetGrouped = buildGroupedRecordsFromRows(rawRows, headerMap, absentRecordsMap, malpracticeRecordsMap, new Map());
      setGroupedRecords(resetGrouped);
    }

    setStatus("Cleared heldback records. Student marks and evaluations restored.", "info");
  };

  // Download Heldback Report Template (.xlsx) with standard university APC / Heldback columns and sample rows
  const handleDownloadHeldbackTemplate = () => {
    const headers = [
      "Seat Number",
      "PRN",
      "Reason",
      "Assessment Type",
      "College Code",
      "Student Name",
      "Paper",
      "College Name",
      "Teaching Learning Method",
      "Assessment Method"
    ];

    const sampleRows = [
      [
        "NK24ECOR020",
        "2024012600000989",
        "APC WITHIN CONDONABLE LIMIT",
        "Heldback at term-level",
        "NK",
        "AHADAL JAMEEL. V.K",
        "Heldback at term-level",
        "Naher Arts and Science College, Kanhirode",
        "Heldback at term-level",
        "Heldback at term-level"
      ],
      [
        "EK24ECOR001",
        "2024012600011711",
        "APC WCL",
        "Heldback at term-level",
        "EK",
        "AFNA RAHIMAN CA",
        "Heldback at term-level",
        "E.K. Nayanar Memorial Govt. College, Elerithattu",
        "Heldback at term-level",
        "Heldback at term-level"
      ],
      [
        "EK24ECOR018",
        "2024012600011831",
        "APC WCL",
        "Heldback at term-level",
        "EK",
        "DEVIKA P C",
        "Heldback at term-level",
        "E.K. Nayanar Memorial Govt. College, Elerithattu",
        "Heldback at term-level",
        "Heldback at term-level"
      ],
      [
        "EK24ECOR029",
        "2024012600011909",
        "APC WCL",
        "Heldback at term-level",
        "EK",
        "ARJUN SAJI",
        "Heldback at term-level",
        "E.K. Nayanar Memorial Govt. College, Elerithattu",
        "Heldback at term-level",
        "Heldback at term-level"
      ]
    ];

    const aoa = [headers, ...sampleRows];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [
      { wch: 16 }, // Seat Number
      { wch: 20 }, // PRN
      { wch: 32 }, // Reason
      { wch: 24 }, // Assessment Type
      { wch: 14 }, // College Code
      { wch: 26 }, // Student Name
      { wch: 24 }, // Paper
      { wch: 45 }, // College Name
      { wch: 24 }, // TLM
      { wch: 24 }  // AM
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Heldback_Template");
    XLSX.writeFile(wb, "heldback_report_template.xlsx");
    setStatus("Downloaded Heldback Report Template (.xlsx) with standard term-level columns and sample entries.", "success");
  };

  // Export Course-Wise Pass Simulation Report (+0 to +10 Moderation)
  const handleExportSimulationExcel = () => {
    if (!courseSimulationData || courseSimulationData.length === 0) {
      setStatus("No course data available to export simulation. Please upload a source file first.", "error");
      return;
    }

    const headers = [
      "Faculty",
      "Program Term Name",
      "Course Code",
      "Course Name",
      "Component Type",
      "Total Students",
      "Absent (Ab)",
      "Malpractice (MP)",
      "Held (Heldback)",
      "Held (Missing Component)",
      "30% ESE Pass (0 Mod)",
      "30% ESE Pass %",
      "35% Overall Pass (0 Mod)",
      "35% Overall Pass %",
      "Normal Pass (Both 30% & 35% Met)",
      "Normal Pass %",
      "+1 Mod Pass (Both Met)",
      "+2 Mod Pass (Both Met)",
      "+3 Mod Pass (Both Met)",
      "+4 Mod Pass (Both Met)",
      "+5 Mod Pass (Both Met)",
      "+6 Mod Pass (Both Met)",
      "+7 Mod Pass (Both Met)",
      "+8 Mod Pass (Both Met)",
      "+9 Mod Pass (Both Met)",
      "+10 Mod Pass (Both Met)",
      "+10 Mod Pass %",
      "Max Rescued (+10)"
    ];

    let totalAllStudents = 0;
    let totalAllAbsent = 0;
    let totalAllMalpractice = 0;
    let totalAllHeldback = 0;
    let totalAllHeldMissing = 0;
    let totalAllRawEsePass = 0;
    let totalAllRawOverallPass = 0;
    let totalAllRawPass = 0;
    const totalAllModPass = Array(11).fill(0);

    const rows = courseSimulationData.map(c => {
      const absent = c.absentCount || 0;
      const malpractice = c.malpracticeCount || 0;
      const heldback = c.heldbackCount || 0;
      const heldMissing = Math.max(0, (c.heldCount || 0) - heldback);
      totalAllStudents += c.totalStudents;
      totalAllAbsent += absent;
      totalAllMalpractice += malpractice;
      totalAllHeldback += heldback;
      totalAllHeldMissing += heldMissing;
      totalAllRawEsePass += c.rawEsePassCount;
      totalAllRawOverallPass += c.rawOverallPassCount;
      totalAllRawPass += c.rawPassCount;
      for (let m = 1; m <= 10; m++) {
        totalAllModPass[m] += c.passCountAtMod[m];
      }

      const rawEsePct = c.totalStudents > 0 ? ((c.rawEsePassCount / c.totalStudents) * 100).toFixed(2) + "%" : "0.00%";
      const rawOverallPct = c.totalStudents > 0 ? ((c.rawOverallPassCount / c.totalStudents) * 100).toFixed(2) + "%" : "0.00%";
      const rawPct = c.totalStudents > 0 ? ((c.rawPassCount / c.totalStudents) * 100).toFixed(2) + "%" : "0.00%";
      const plus10Pct = c.totalStudents > 0 ? ((c.passCountAtMod[10] / c.totalStudents) * 100).toFixed(2) + "%" : "0.00%";
      const compType = (c.hasEseTh && c.hasEsePr)
        ? "Theory & Practical (TH + PR)"
        : (c.isPrOnly
            ? "Practical Only (PR)"
            : "Theory (TH)");

      return [
        c.faculty,
        c.program,
        c.courseCode,
        c.courseName,
        compType,
        c.totalStudents,
        absent,
        malpractice,
        heldback,
        heldMissing,
        c.rawEsePassCount,
        rawEsePct,
        c.rawOverallPassCount,
        rawOverallPct,
        c.rawPassCount,
        rawPct,
        c.passCountAtMod[1],
        c.passCountAtMod[2],
        c.passCountAtMod[3],
        c.passCountAtMod[4],
        c.passCountAtMod[5],
        c.passCountAtMod[6],
        c.passCountAtMod[7],
        c.passCountAtMod[8],
        c.passCountAtMod[9],
        c.passCountAtMod[10],
        plus10Pct,
        maxRescued
      ];
    });

    // Summary / Total Row
    const overallRawEsePct = totalAllStudents > 0 ? ((totalAllRawEsePass / totalAllStudents) * 100).toFixed(2) + "%" : "0.00%";
    const overallRawOverallPct = totalAllStudents > 0 ? ((totalAllRawOverallPass / totalAllStudents) * 100).toFixed(2) + "%" : "0.00%";
    const overallRawPct = totalAllStudents > 0 ? ((totalAllRawPass / totalAllStudents) * 100).toFixed(2) + "%" : "0.00%";
    const overallPlus10Pct = totalAllStudents > 0 ? ((totalAllModPass[10] / totalAllStudents) * 100).toFixed(2) + "%" : "0.00%";
    const overallMaxRescued = totalAllModPass[10] - totalAllRawPass;

    const summaryRow = [
      "TOTAL / ALL COURSES",
      "-",
      "-",
      "-",
      "-",
      totalAllStudents,
      totalAllAbsent,
      totalAllMalpractice,
      totalAllHeldback,
      totalAllHeldMissing,
      totalAllRawEsePass,
      overallRawEsePct,
      totalAllRawOverallPass,
      overallRawOverallPct,
      totalAllRawPass,
      overallRawPct,
      totalAllModPass[1],
      totalAllModPass[2],
      totalAllModPass[3],
      totalAllModPass[4],
      totalAllModPass[5],
      totalAllModPass[6],
      totalAllModPass[7],
      totalAllModPass[8],
      totalAllModPass[9],
      totalAllModPass[10],
      overallPlus10Pct,
      overallMaxRescued
    ];

    const aoa = [headers, ...rows, summaryRow];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    ws["!cols"] = [
      { wch: 24 }, // Faculty
      { wch: 24 }, // Program
      { wch: 18 }, // Course Code
      { wch: 32 }, // Course Name
      { wch: 16 }, // Component Type
      { wch: 14 }, // Total Students
      { wch: 14 }, // Absent (Ab)
      { wch: 16 }, // Malpractice (MP)
      { wch: 18 }, // Held (Heldback)
      { wch: 22 }, // Held (Missing Component)
      { wch: 20 }, // 30% ESE Pass
      { wch: 15 }, // 30% ESE Pass %
      { wch: 22 }, // 35% Overall Pass
      { wch: 16 }, // 35% Overall Pass %
      { wch: 26 }, // Normal Pass (Both Met)
      { wch: 15 }, // Normal Pass %
      { wch: 18 }, // +1 Mod Pass
      { wch: 18 }, // +2 Mod Pass
      { wch: 18 }, // +3 Mod Pass
      { wch: 18 }, // +4 Mod Pass
      { wch: 18 }, // +5 Mod Pass
      { wch: 18 }, // +6 Mod Pass
      { wch: 18 }, // +7 Mod Pass
      { wch: 18 }, // +8 Mod Pass
      { wch: 18 }, // +9 Mod Pass
      { wch: 18 }, // +10 Mod Pass
      { wch: 16 }, // +10 Mod Pass %
      { wch: 16 }  // Max Rescued (+10)
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Course_Pass_Simulation");
    XLSX.writeFile(wb, "course_wise_pass_simulation.xlsx");
    setStatus("Generated & downloaded Course-Wise Pass Simulation Report (+0 to +10 Moderation) with 30% ESE and 35% Aggregate checks.", "success");
  };

  // Export Student-Wise Semester Results Report (.xlsx) - Supports active live filtered view or all students
  const handleExportStudentSemesterExcel = (studentsToExport = filteredStudents, customFilename = null) => {
    if (!studentsToExport || studentsToExport.length === 0) {
      setStatus("No student records available to export for the current filter/view.", "warning");
      return;
    }

    const isFiltered = studentsToExport.length !== studentSemesterData.length;
    let defaultName = "student_semester_results.xlsx";
    if (isFiltered) {
      const parts = ["student_semester_results"];
      if (studentCollegeFilter !== "ALL") {
        parts.push(studentCollegeFilter.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 16));
      }
      if (studentProgramFilter !== "ALL") {
        parts.push(studentProgramFilter.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 16));
      }
      if (studentCourseFilter !== "ALL") {
        parts.push(studentCourseFilter.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 14));
      }
      if (studentFilterStatus !== "ALL") {
        parts.push(studentFilterStatus.toLowerCase());
      }
      parts.push(`${studentsToExport.length}students.xlsx`);
      defaultName = parts.join("_");
    }
    const filename = customFilename || defaultName;

    setIsProcessing(true);
    setStatus("Generating Student Semester Excel export...", "info");

    try {
      const headers = [
        "Faculty",
        "College",
        "Program Term Name",
        "Seat Number",
        "PRN",
        "Papers Attempted",
        "Papers Passed",
        "Papers Failed",
        "Heldback / Missing Papers",
        "Raw Semester Result (0 Mod)",
        "Final Semester Result",
        "Semester Rescued via Moderation",
        "Total Moderation Marks Awarded",
        "Failed Courses List",
        "All Attempted Courses Breakdown"
      ];

      // Dynamic summary stats based on the exported subset
      let totalAttempted = 0;
      let totalPassed = 0;
      let totalFailed = 0;
      let totalRawPass = 0;
      let totalFinalPass = 0;
      let totalRescued = 0;
      let totalHeld = 0;
      let totalHeldback = 0;

      const rows = studentsToExport.map(st => {
        totalAttempted += st.totalCourses;
        totalPassed += st.finalPassedCourses;
        totalFailed += st.finalFailedCourses;
        if (st.isHeldback) totalHeldback++;
        if (st.isHeld) totalHeld++;
        if (st.rawSemesterPass) totalRawPass++;
        if (st.finalSemesterPass) totalFinalPass++;
        if (st.isRescuedSemester) totalRescued++;

        const failedList = st.courses.filter(c => c.coursePass === "Fail").map(c => `${c.courseCode} (${c.courseName})`).join("; ");
        const coursesSummary = st.courses.map(c => `${c.courseCode}: ${c.coursePass}${c.modMarks > 0 ? ` (+${c.modMarks} Mod)` : ''}`).join("; ");
        const heldInfo = st.isHeldback 
          ? `Heldback (${st.heldbackReason || "Term-level"})`
          : st.heldCourses > 0 
            ? `${st.heldCourses} Missing Component`
            : "0";

        return [
          st.faculty,
          st.college || "-",
          st.program,
          st.seatNumber,
          st.prn,
          st.totalCourses,
          st.finalPassedCourses,
          st.finalFailedCourses,
          heldInfo,
          st.rawSemesterResult,
          st.semesterResult,
          st.isRescuedSemester ? "Yes (Rescued)" : "No",
          st.totalModerationMarks,
          failedList || "None (All Passed)",
          coursesSummary
        ];
      });

      const expCount = studentsToExport.length;
      const rawPct = ((totalRawPass / expCount) * 100).toFixed(1);
      const finalPct = ((totalFinalPass / expCount) * 100).toFixed(1);
      const resPct = ((totalRescued / expCount) * 100).toFixed(1);
      const failCount = expCount - totalFinalPass - totalHeld;
      const failPct = ((failCount / expCount) * 100).toFixed(1);

      // Summary / Total Row
      const summaryRow = [
        `TOTAL EXPORTED STUDENTS: ${expCount}`,
        "-",
        "-",
        "-",
        totalAttempted,
        totalPassed,
        totalFailed,
        `${totalHeld} Held (${totalHeldback} Heldback)`,
        `${totalRawPass} Passed (${rawPct}%)`,
        `${totalFinalPass} Passed (${finalPct}%)`,
        `+${totalRescued} Rescued (${resPct}%)`,
        "-",
        `${failCount} Failed (${failPct}%)`,
        "-"
      ];

      const aoa = [headers, ...rows, summaryRow];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(aoa);

      ws["!autofilter"] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: aoa.length - 2, c: headers.length - 1 }
        })
      };

      ws["!cols"] = [
        { wch: 24 }, // Faculty
        { wch: 28 }, // Program
        { wch: 16 }, // Seat Number
        { wch: 18 }, // PRN
        { wch: 16 }, // Papers Attempted
        { wch: 14 }, // Papers Passed
        { wch: 14 }, // Papers Failed
        { wch: 24 }, // Heldback / Missing Papers
        { wch: 24 }, // Raw Semester Result
        { wch: 24 }, // Final Semester Result
        { wch: 22 }, // Semester Rescued
        { wch: 18 }, // Total Moderation Marks
        { wch: 40 }, // Failed Courses List
        { wch: 60 }  // All Attempted Courses Breakdown
      ];

      XLSX.utils.book_append_sheet(wb, ws, "Student_Semester_Results");
      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);

      setStatus(`Successfully exported ${studentsToExport.length} student record(s) [${isFiltered ? 'Filtered Live View' : 'All Students'}] to ${filename}!`, "success");
    } catch (err) {
      console.error("Export Student Semester Error:", err);
      setStatus("Failed to export Student Semester summary: " + err.message, "error");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleSort = (column) => {
    setSortConfig(prev => {
      if (prev.column !== column) {
        return { column, direction: "asc" };
      }
      if (prev.direction === "asc") {
        return { column, direction: "desc" };
      }
      return { column: null, direction: null };
    });
    setPage(0);
  };

  const updateColumnFilter = (colKey, val) => {
    setColumnFilters(prev => {
      const next = { ...prev };
      if (val === undefined || val === null || String(val).trim() === "") {
        delete next[colKey];
      } else {
        next[colKey] = String(val);
      }
      return next;
    });
    setPage(0);
  };

  const clearAllFilters = () => {
    setColumnFilters({});
    setSortConfig({ column: null, direction: null });
    setSelectedFacultyFilter("ALL");
    setSelectedCollegeFilter("ALL");
    setSelectedProgramFilter("ALL");
    setSelectedCourseFilter("ALL");
    setSelectedResultFilter("ALL");
    setSearchQuery("");
    setPage(0);
  };

  // Context-Scoped Course Rows (Faculty, College, Program, Course, Search, Column Filters)
  const scopedCourseRows = useMemo(() => {
    let result = [...processedRows];

    // 1. Faculty filter
    if (selectedFacultyFilter !== "ALL") {
      result = result.filter(r => r["Faculty"] === selectedFacultyFilter);
    }

    // 2. College filter
    if (selectedCollegeFilter !== "ALL") {
      result = result.filter(r => (r._college || r["College Name"] || r["College Code"] || r["College"]) === selectedCollegeFilter);
    }

    // 3. Program filter
    if (selectedProgramFilter !== "ALL") {
      result = result.filter(r => r["Program Term Name"] === selectedProgramFilter);
    }

    // 4. Course filter
    if (selectedCourseFilter !== "ALL") {
      const cleanFilter = cleanCourseCode(selectedCourseFilter);
      result = result.filter(r => cleanCourseCode(r["Course Code"]) === cleanFilter);
    }

    // 5. Search Query
    if (searchQuery) {
      const q = searchQuery.toLowerCase().trim();
      result = result.filter(r =>
        (r._college && r._college.toLowerCase().includes(q)) ||
        Object.values(r).some(val => String(val || "").toLowerCase().includes(q))
      );
    }

    // 6. Column Filters
    const activeFilterEntries = Object.entries(columnFilters);
    if (activeFilterEntries.length > 0) {
      result = result.filter(r => {
        return activeFilterEntries.every(([colKey, filterVal]) => {
          if (!filterVal || String(filterVal).trim() === "") return true;
          const cellVal = String(r[colKey] !== undefined && r[colKey] !== null ? r[colKey] : "").toLowerCase();
          return cellVal.includes(String(filterVal).toLowerCase().trim());
        });
      });
    }

    return result;
  }, [processedRows, selectedFacultyFilter, selectedCollegeFilter, selectedProgramFilter, selectedCourseFilter, searchQuery, columnFilters]);

  // Unique Lists for Dropdown Filters (Cascading & Contextual)
  const uniqueFaculties = useMemo(() => {
    const set = new Set();
    processedRows.forEach(r => {
      const col = r._college || r["College Name"] || r["College Code"] || r["College"];
      const prog = r["Program Term Name"];
      const course = r["Course Code"];
      if (
        (selectedCollegeFilter === "ALL" || col === selectedCollegeFilter) &&
        (selectedProgramFilter === "ALL" || prog === selectedProgramFilter) &&
        (selectedCourseFilter === "ALL" || cleanCourseCode(course) === cleanCourseCode(selectedCourseFilter))
      ) {
        if (r["Faculty"]) set.add(r["Faculty"]);
      }
    });
    return Array.from(set).sort();
  }, [processedRows, selectedCollegeFilter, selectedProgramFilter, selectedCourseFilter]);

  const uniqueColleges = useMemo(() => {
    const set = new Set();
    processedRows.forEach(r => {
      const fac = r["Faculty"];
      const prog = r["Program Term Name"];
      const course = r["Course Code"];
      if (
        (selectedFacultyFilter === "ALL" || fac === selectedFacultyFilter) &&
        (selectedProgramFilter === "ALL" || prog === selectedProgramFilter) &&
        (selectedCourseFilter === "ALL" || cleanCourseCode(course) === cleanCourseCode(selectedCourseFilter))
      ) {
        const col = r._college || r["College Name"] || r["College Code"] || r["College"];
        if (col) set.add(col);
      }
    });
    return Array.from(set).sort();
  }, [processedRows, selectedFacultyFilter, selectedProgramFilter, selectedCourseFilter]);

  const uniquePrograms = useMemo(() => {
    const set = new Set();
    processedRows.forEach(r => {
      const fac = r["Faculty"];
      const col = r._college || r["College Name"] || r["College Code"] || r["College"];
      const course = r["Course Code"];
      if (
        (selectedFacultyFilter === "ALL" || fac === selectedFacultyFilter) &&
        (selectedCollegeFilter === "ALL" || col === selectedCollegeFilter) &&
        (selectedCourseFilter === "ALL" || cleanCourseCode(course) === cleanCourseCode(selectedCourseFilter))
      ) {
        if (r["Program Term Name"]) set.add(r["Program Term Name"]);
      }
    });
    return Array.from(set).sort();
  }, [processedRows, selectedFacultyFilter, selectedCollegeFilter, selectedCourseFilter]);

  const uniqueCourses = useMemo(() => {
    const courseMap = new Map();
    processedRows.forEach(r => {
      const fac = r["Faculty"];
      const col = r._college || r["College Name"] || r["College Code"] || r["College"];
      const prog = r["Program Term Name"];
      const code = cleanCourseCode(r["Course Code"]);
      const name = cleanCourseName(r["Course Name"]);
      const norm = normalizeKey(code);
      if (
        (selectedFacultyFilter === "ALL" || fac === selectedFacultyFilter) &&
        (selectedCollegeFilter === "ALL" || col === selectedCollegeFilter) &&
        (selectedProgramFilter === "ALL" || prog === selectedProgramFilter)
      ) {
        if (norm && !courseMap.has(norm)) {
          courseMap.set(norm, { code, label: name ? `${code} - ${name}` : code });
        }
      }
    });
    return Array.from(courseMap.values()).sort((a, b) => a.code.localeCompare(b.code));
  }, [processedRows, selectedFacultyFilter, selectedCollegeFilter, selectedProgramFilter]);

  // Consolidated / Global Overall Course Metrics (Unfiltered, for sidebar)
  const consolidatedCourseMetrics = useMemo(() => {
    const total = processedRows.length;
    if (total === 0) return { 
      total: 0, 
      uniqueStudents: 0, 
      rawPassed: 0, 
      rawPassPct: "0.0",
      moderatedPassed: 0, 
      moderatedPct: "0.0",
      totalPassed: 0, 
      passPct: "0.0", 
      failed: 0, 
      failedPct: "0.0",
      heldCount: 0, 
      heldPct: "0.0",
      heldbackCount: 0, 
      missingCompCount: 0, 
      eseFailed: 0, 
      overallFailed: 0, 
      absentCount: 0, 
      absentPct: "0.0",
      malpracticeCount: 0,
      malpracticePct: "0.0"
    };

    const prnSet = new Set();
    let rawPassed = 0;
    let moderatedPassed = 0;
    let totalPassed = 0;
    let failed = 0;
    let heldCount = 0;
    let heldbackCount = 0;
    let missingCompCount = 0;
    let eseFailed = 0;
    let overallFailed = 0;
    let absentCount = 0;
    let malpracticeCount = 0;

    const coursePassKey = "Course Pass/Fail";
    const esePassKey = "ESE Pass";
    const overallPassKey = "Overall pass";

    processedRows.forEach(r => {
      if (r["PRN"]) prnSet.add(r["PRN"]);
      
      if (r._isHeldback) {
        heldbackCount++;
        heldCount++;
        return;
      }

      if (r._isHeld) {
        missingCompCount++;
        heldCount++;
        return;
      }

      if (r._isMalpractice) {
        malpracticeCount++;
      } else if (r._isAbsent) {
        absentCount++;
      }

      if (r._rawPass) {
        rawPassed++;
      }
      if (r._isModeratedPass) {
        moderatedPassed++;
      }

      if (r[coursePassKey] === "Pass") {
        totalPassed++;
      } else {
        failed++;
        if (r[esePassKey] === "Fail") eseFailed++;
        if (r[overallPassKey] === "Fail") overallFailed++;
      }
    });

    const evaluatedTotal = total - heldCount;
    const passPct = evaluatedTotal > 0 ? ((totalPassed / evaluatedTotal) * 100).toFixed(1) : "0.0";
    const rawPassPct = evaluatedTotal > 0 ? ((rawPassed / evaluatedTotal) * 100).toFixed(1) : "0.0";
    const failedPct = evaluatedTotal > 0 ? ((failed / evaluatedTotal) * 100).toFixed(1) : "0.0";
    const heldPct = total > 0 ? ((heldCount / total) * 100).toFixed(1) : "0.0";

    return {
      total,
      uniqueStudents: prnSet.size,
      rawPassed,
      rawPassPct,
      moderatedPassed,
      moderatedPct: evaluatedTotal > 0 ? ((moderatedPassed / evaluatedTotal) * 100).toFixed(1) : "0.0",
      totalPassed,
      failed,
      failedPct,
      heldCount,
      heldPct,
      heldbackCount,
      missingCompCount,
      passPct,
      eseFailed,
      overallFailed,
      absentCount,
      absentPct: total > 0 ? ((absentCount / total) * 100).toFixed(1) : "0.0",
      malpracticeCount,
      malpracticePct: total > 0 ? ((malpracticeCount / total) * 100).toFixed(1) : "0.0"
    };
  }, [processedRows]);

  // Statistics Metrics (calculated from active scoped filters, for Course Results view)
  const metrics = useMemo(() => {
    const total = scopedCourseRows.length;
    if (total === 0) return { 
      total: 0, 
      uniqueStudents: 0, 
      rawPassed: 0, 
      rawPassPct: "0.0",
      moderatedPassed: 0, 
      moderatedPct: "0.0",
      totalPassed: 0, 
      passPct: "0.0", 
      failed: 0, 
      failedPct: "0.0",
      heldCount: 0, 
      heldPct: "0.0",
      heldbackCount: 0, 
      missingCompCount: 0, 
      eseFailed: 0, 
      overallFailed: 0, 
      absentCount: 0, 
      absentPct: "0.0",
      malpracticeCount: 0,
      malpracticePct: "0.0"
    };

    const prnSet = new Set();
    let rawPassed = 0;
    let moderatedPassed = 0;
    let totalPassed = 0;
    let failed = 0;
    let heldCount = 0;
    let heldbackCount = 0;
    let missingCompCount = 0;
    let eseFailed = 0;
    let overallFailed = 0;
    let absentCount = 0;
    let malpracticeCount = 0;

    const coursePassKey = "Course Pass/Fail";
    const esePassKey = "ESE Pass";
    const overallPassKey = "Overall pass";

    scopedCourseRows.forEach(r => {
      if (r["PRN"]) prnSet.add(r["PRN"]);
      
      if (r._isHeldback) {
        heldbackCount++;
        heldCount++;
        return;
      }

      if (r._isHeld) {
        missingCompCount++;
        heldCount++;
        return;
      }

      if (r._isMalpractice) {
        malpracticeCount++;
      } else if (r._isAbsent) {
        absentCount++;
      }

      if (r._rawPass) {
        rawPassed++;
      }
      if (r._isModeratedPass) {
        moderatedPassed++;
      }

      if (r[coursePassKey] === "Pass") {
        totalPassed++;
      } else {
        failed++;
        if (r[esePassKey] === "Fail") eseFailed++;
        if (r[overallPassKey] === "Fail") overallFailed++;
      }
    });

    const evaluatedTotal = total - heldCount;
    const passPct = evaluatedTotal > 0 ? ((totalPassed / evaluatedTotal) * 100).toFixed(1) : "0.0";
    const rawPassPct = evaluatedTotal > 0 ? ((rawPassed / evaluatedTotal) * 100).toFixed(1) : "0.0";
    const failedPct = evaluatedTotal > 0 ? ((failed / evaluatedTotal) * 100).toFixed(1) : "0.0";
    const heldPct = total > 0 ? ((heldCount / total) * 100).toFixed(1) : "0.0";

    return {
      total,
      uniqueStudents: prnSet.size,
      rawPassed,
      rawPassPct,
      moderatedPassed,
      moderatedPct: evaluatedTotal > 0 ? ((moderatedPassed / evaluatedTotal) * 100).toFixed(1) : "0.0",
      totalPassed,
      failed,
      failedPct,
      heldCount,
      heldPct,
      heldbackCount,
      missingCompCount,
      passPct,
      eseFailed,
      overallFailed,
      absentCount,
      absentPct: total > 0 ? ((absentCount / total) * 100).toFixed(1) : "0.0",
      malpracticeCount,
      malpracticePct: total > 0 ? ((malpracticeCount / total) * 100).toFixed(1) : "0.0"
    };
  }, [scopedCourseRows]);

  // Filtered & Sorted Rows
  const filteredRows = useMemo(() => {
    let result = [...scopedCourseRows];

    const coursePassKey = "Course Pass/Fail";
    const esePassKey = "ESE Pass";
    const overallPassKey = "Overall pass";

    // Result Status Filter
    if (selectedResultFilter === "PASS") {
      result = result.filter(r => r[coursePassKey] === "Pass");
    } else if (selectedResultFilter === "PASS_MOD") {
      result = result.filter(r => r._isModeratedPass);
    } else if (selectedResultFilter === "FAIL") {
      result = result.filter(r => r[coursePassKey] === "Fail");
    } else if (selectedResultFilter === "HELDBACK") {
      result = result.filter(r => r._isHeldback);
    } else if (selectedResultFilter === "HELD_MISSING") {
      result = result.filter(r => r._isMissingComp);
    } else if (selectedResultFilter === "HELD") {
      result = result.filter(r => r._isHeld);
    } else if (selectedResultFilter === "ESE_FAIL") {
      result = result.filter(r => r[esePassKey] === "Fail");
    } else if (selectedResultFilter === "OVERALL_FAIL") {
      result = result.filter(r => r[overallPassKey] === "Fail");
    } else if (selectedResultFilter === "ABSENT") {
      result = result.filter(r => r._isAbsent);
    } else if (selectedResultFilter === "MALPRACTICE") {
      result = result.filter(r => r._isMalpractice);
    }

    // Sorting
    if (sortConfig.column && sortConfig.direction) {
      const col = sortConfig.column;
      const dir = sortConfig.direction === "asc" ? 1 : -1;

      result.sort((a, b) => {
        const valA = a[col] !== undefined && a[col] !== null ? a[col] : "";
        const valB = b[col] !== undefined && b[col] !== null ? b[col] : "";

        const numA = Number(valA);
        const numB = Number(valB);
        const isNumA = typeof valA === "number" || (String(valA).trim() !== "" && !isNaN(numA));
        const isNumB = typeof valB === "number" || (String(valB).trim() !== "" && !isNaN(numB));

        if (isNumA && isNumB) {
          return (numA - numB) * dir;
        }

        return String(valA).localeCompare(String(valB), undefined, { numeric: true, sensitivity: "base" }) * dir;
      });
    }

    return result;
  }, [scopedCourseRows, selectedResultFilter, sortConfig]);

  const pagedRows = useMemo(() => {
    const start = page * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, page]);

  // Export to Excel Matching exact 31 columns structure with "Output file " sheet name
  const handleExportExcel = (rowsToExport = filteredRows, customFilename = null) => {
    if (!rowsToExport || rowsToExport.length === 0) {
      alert("No rows to export.");
      return;
    }

    const isFiltered = processedRows && rowsToExport.length !== processedRows.length;
    let filename = customFilename;
    if (!filename) {
      if (isFiltered) {
        const parts = ["ades_course_results"];
        if (selectedCollegeFilter !== "ALL") {
          parts.push(selectedCollegeFilter.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 16));
        }
        if (selectedProgramFilter !== "ALL") {
          parts.push(selectedProgramFilter.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 16));
        }
        if (selectedCourseFilter !== "ALL") {
          parts.push(selectedCourseFilter.replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 14));
        }
        if (selectedResultFilter !== "ALL") {
          parts.push(selectedResultFilter.toLowerCase());
        }
        parts.push(`${rowsToExport.length}rows.xlsx`);
        filename = parts.join("_");
      } else {
        filename = "converted_output_card.xlsx";
      }
    }

    setIsProcessing(true);
    setStatus("Generating ADES Result Excel export...", "info");

    try {
      const aoa = [
        ADES_OUTPUT_HEADERS,
        ...rowsToExport.map(r => ADES_OUTPUT_HEADERS.map(h => r[h] !== undefined ? r[h] : ""))
      ];

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });

      ws["!autofilter"] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: aoa.length - 1, c: ADES_OUTPUT_HEADERS.length - 1 }
        })
      };

      ws["!cols"] = ADES_OUTPUT_HEADERS.map(h => ({
        wch: Math.min(Math.max(h.length + 3, 14), 45)
      }));

      XLSX.utils.book_append_sheet(wb, ws, "Output file ");
      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);

      setStatus("Successfully exported " + rowsToExport.length + " calculated rows to " + filename + "!", "success");
    } catch (err) {
      console.error("Export Error:", err);
      setStatus("Export failed: " + err.message, "error");
    } finally {
      setIsProcessing(false);
    }
  };

  const coursePassKey = "Course Pass/Fail";

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", width: "100vw", overflow: "hidden", background: "var(--bg)", color: "var(--ink)" }}>
      
      {/* Top Navigation Bar */}
      <header style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 20px", borderBottom: "1px solid var(--line)", background: "var(--panel)", flexShrink: 0 }}>
        <div style={{ display: "flex", alignItems: "center", gap: "14px" }}>
          <Link to="/" style={{ display: "flex", alignItems: "center", gap: "6px", color: "var(--muted)", textDecoration: "none", fontSize: "13px", fontWeight: 600 }}>
            <ArrowLeft size={16} /> Back to Portal
          </Link>
          <div style={{ height: "18px", width: "1px", background: "var(--line)" }} />
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <Calculator size={20} color="var(--accent)" />
            <h2 style={{ fontSize: "16px", fontWeight: 700, margin: 0 }}>ADES Result Calculator</h2>
            <span style={{ fontSize: "11px", background: "var(--accent-soft)", color: "var(--accent)", padding: "2px 8px", borderRadius: "12px", fontWeight: 600 }}>
              ESE & CE Evaluator + Moderation
            </span>
          </div>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          {/* Tab Switcher */}
          <div style={{ display: "flex", background: "var(--bg)", padding: "3px", borderRadius: "8px", border: "1px solid var(--line)", gap: "4px" }}>
            <button
              type="button"
              onClick={() => setActiveTab("results")}
              style={{
                display: "flex",
                alignItems: "center",
                gap: "6px",
                padding: "5px 12px",
                fontSize: "12px",
                fontWeight: 600,
                borderRadius: "6px",
                border: "none",
                cursor: "pointer",
                background: activeTab === "results" ? "var(--accent)" : "transparent",
                color: activeTab === "results" ? "white" : "var(--muted)"
              }}
            >
              <Table size={13} /> Course Results (31 Cols)
            </button>
            <button
              type="button"
              onClick={() => setActiveTab("students")}
              style={{
                display: "flex",
                alignItems: "center",
                gap: "6px",
                padding: "5px 12px",
                fontSize: "12px",
                fontWeight: 600,
                borderRadius: "6px",
                border: "none",
                cursor: "pointer",
                background: activeTab === "students" ? "var(--accent)" : "transparent",
                color: activeTab === "students" ? "white" : "var(--muted)"
              }}
            >
              <Users size={13} /> Student Semester Results ({studentMetrics.totalStudents})
            </button>
            <button
              type="button"
              onClick={() => setActiveTab("moderation")}
              style={{
                display: "flex",
                alignItems: "center",
                gap: "6px",
                padding: "5px 12px",
                fontSize: "12px",
                fontWeight: 600,
                borderRadius: "6px",
                border: "none",
                cursor: "pointer",
                background: activeTab === "moderation" ? "var(--accent)" : "transparent",
                color: activeTab === "moderation" ? "white" : "var(--muted)"
              }}
            >
              <Sliders size={13} /> Moderation Matrix ({Object.values(courseModerationMap).filter(v => v > 0).length} Active)
            </button>
            <button
              type="button"
              onClick={() => setActiveTab("simulation")}
              style={{
                display: "flex",
                alignItems: "center",
                gap: "6px",
                padding: "5px 12px",
                fontSize: "12px",
                fontWeight: 600,
                borderRadius: "6px",
                border: "none",
                cursor: "pointer",
                background: activeTab === "simulation" ? "var(--accent)" : "transparent",
                color: activeTab === "simulation" ? "white" : "var(--muted)"
              }}
            >
              <TrendingUp size={13} /> Pass Simulation (+0 to +10)
            </button>
          </div>

          <button 
            type="button" 
            className="secondary" 
            onClick={() => setShowHelpModal(true)}
            style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", padding: "6px 12px" }}
          >
            <HelpCircle size={14} /> Extraction Logic & Guide
          </button>

          {processedRows.length > 0 && (
            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
              <button 
                type="button" 
                onClick={() => handleExportStudentSemesterExcel(filteredStudents)}
                style={{ 
                  display: "flex", 
                  alignItems: "center", 
                  gap: "6px", 
                  padding: "6px 12px", 
                  fontSize: "12px", 
                  background: "#6366f1", 
                  color: "white", 
                  border: "none", 
                  borderRadius: "6px", 
                  fontWeight: 600, 
                  cursor: "pointer" 
                }}
                title={filteredStudents.length !== studentSemesterData.length ? `Export current filtered view (${filteredStudents.length} students)` : "Export all student semester results"}
              >
                <Download size={14} /> Export Students ({filteredStudents.length !== studentSemesterData.length ? `${filteredStudents.length}/${studentMetrics.totalStudents}` : studentMetrics.totalStudents})
              </button>
              <button 
                type="button" 
                onClick={handleExportSimulationExcel}
                style={{ 
                  display: "flex", 
                  alignItems: "center", 
                  gap: "6px", 
                  padding: "6px 12px", 
                  fontSize: "12px", 
                  background: "#10b981", 
                  color: "white", 
                  border: "none", 
                  borderRadius: "6px", 
                  fontWeight: 600, 
                  cursor: "pointer" 
                }}
                title="Export course-wise pass count under 0 to +10 moderation marks"
              >
                <Download size={14} /> Export Simulation (+0..+10)
              </button>
              <button 
                type="button" 
                onClick={() => handleExportExcel(filteredRows)}
                style={{ 
                  display: "flex", 
                  alignItems: "center", 
                  gap: "6px", 
                  padding: "6px 14px", 
                  fontSize: "12.5px", 
                  background: "var(--accent)", 
                  color: "white", 
                  border: "none", 
                  borderRadius: "6px", 
                  fontWeight: 600, 
                  cursor: "pointer" 
                }}
                title={filteredRows.length !== processedRows.length ? `Export current filtered view (${filteredRows.length} rows)` : "Export all 31-column ADES results"}
              >
                <Download size={14} /> Export 31-Col XLSX ({filteredRows.length !== processedRows.length ? `${filteredRows.length}/${processedRows.length}` : `${filteredRows.length} Rows`})
              </button>
            </div>
          )}
        </div>
      </header>

      {/* Main Workspace Layout */}
      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        
        {/* Left Sidebar / Config Panel */}
        <aside style={{ width: "320px", borderRight: "1px solid var(--line)", background: "var(--panel)", display: "flex", flexDirection: "column", flexShrink: 0, overflowY: "auto", padding: "16px", gap: "16px" }}>
          
          {/* File Upload Box */}
          <div style={{ background: "var(--bg)", border: "1.5px dashed var(--line)", borderRadius: "8px", padding: "16px", textAlign: "center", position: "relative" }}>
            <input 
              type="file" 
              accept=".xlsx,.xls,.csv" 
              onChange={handleFileUpload}
              style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", opacity: 0, cursor: "pointer" }}
            />
            <FileSpreadsheet size={32} color="var(--accent)" style={{ margin: "0 auto 8px", opacity: 0.8 }} />
            <div style={{ fontSize: "13px", fontWeight: 600, color: "var(--ink)" }}>
              {sourceFile ? sourceFile : "Upload ADES Marks Excel"}
            </div>
            <div style={{ fontSize: "11px", color: "var(--muted)", marginTop: "4px" }}>
              Drop .xlsx / .xls file here (Raw Marks or Output Format)
            </div>
          </div>

          {/* Sheet Selector (if multiple sheets exist) */}
          {sheetNames.length > 1 && (
            <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
              <label style={{ fontSize: "11.5px", fontWeight: 600, color: "var(--muted)" }}>Select Source Sheet:</label>
              <select 
                value={selectedSheet} 
                onChange={(e) => handleSheetChange(e.target.value)}
                style={{ padding: "6px 10px", fontSize: "12px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)" }}
              >
                {sheetNames.map(s => {
                  const meta = sheetMetadata[s];
                  const hasHeaders = meta && meta.matchedHeaders && meta.matchedHeaders.length > 0;
                  return (
                    <option key={s} value={s}>
                      {s} {hasHeaders ? "(✓ " + meta.matchedHeaders.length + " headers)" : ""}
                    </option>
                  );
                })}
              </select>
            </div>
          )}

          {/* Absent Mark Entry Report Card */}
          <div style={{ 
            background: absentList.length > 0 ? "rgba(239, 68, 68, 0.05)" : "var(--bg)", 
            border: absentList.length > 0 ? "1.5px solid rgba(239, 68, 68, 0.35)" : "1px solid var(--line)", 
            borderRadius: "8px", 
            padding: "12px", 
            display: "flex", 
            flexDirection: "column", 
            gap: "8px" 
          }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ fontSize: "12px", fontWeight: 700, color: absentList.length > 0 ? "#ef4444" : "var(--ink)", display: "flex", alignItems: "center", gap: "6px" }}>
                <UserX size={14} color={absentList.length > 0 ? "#ef4444" : "var(--muted)"} /> Absent Mark Entry
              </div>
              {absentList.length > 0 ? (
                <span style={{ fontSize: "10px", background: "#ef4444", color: "white", padding: "1px 6px", borderRadius: "10px", fontWeight: 700 }}>
                  {absentList.length} Absent Records
                </span>
              ) : (
                <span style={{ fontSize: "10px", background: "var(--panel)", color: "var(--muted)", padding: "1px 6px", borderRadius: "10px", border: "1px solid var(--line)" }}>
                  Optional
                </span>
              )}
            </div>

            <div style={{ fontSize: "11px", color: "var(--muted)", lineHeight: "1.35" }}>
              {absentList.length > 0 ? (
                <span>
                  Active: <strong>{absentFileName}</strong> ({absentList.length} mapped). Mapped components show <strong style={{ color: "#ef4444" }}>Absent (Ab)</strong> &amp; <strong>Fail</strong>.
                </span>
              ) : (
                "Upload one or multiple absent reports across programmes (select multiple .xlsx/.xls files) to replace 0 marks with \"Absent (Ab)\" and mark courses as Fail."
              )}
            </div>

            <input 
              type="file" 
              ref={absentFileInputRef}
              accept=".xlsx,.xls,.csv" 
              multiple
              onChange={handleAbsentExcelUpload}
              style={{ display: "none" }}
            />

            <div style={{ display: "flex", gap: "6px", marginTop: "2px" }}>
              <button 
                type="button"
                onClick={() => absentFileInputRef.current?.click()}
                style={{
                  flex: 1,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: "4px",
                  padding: "5px 8px",
                  fontSize: "11px",
                  background: absentList.length > 0 ? "rgba(239, 68, 68, 0.12)" : "var(--panel)",
                  color: absentList.length > 0 ? "#ef4444" : "var(--ink)",
                  border: absentList.length > 0 ? "1px solid rgba(239, 68, 68, 0.3)" : "1px solid var(--line)",
                  borderRadius: "4px",
                  fontWeight: 600,
                  cursor: "pointer"
                }}
              >
                <FileUp size={12} /> {absentList.length > 0 ? "Replace / Upload Files" : "Upload Absent Excel"}
              </button>

              <button 
                type="button"
                onClick={handleDownloadAbsentTemplate}
                title="Download standard absent report template (.xlsx)"
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: "4px",
                  padding: "5px 8px",
                  fontSize: "11px",
                  background: "var(--panel)",
                  color: "var(--ink)",
                  border: "1px solid var(--line)",
                  borderRadius: "4px",
                  fontWeight: 600,
                  cursor: "pointer"
                }}
              >
                <FileDown size={12} /> Template
              </button>

              {absentList.length > 0 && (
                <button 
                  type="button"
                  onClick={handleClearAbsentData}
                  title="Clear absent data and revert to original marks"
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    padding: "5px 8px",
                    fontSize: "11px",
                    background: "transparent",
                    color: "#ef4444",
                    border: "1px solid #ef4444",
                    borderRadius: "4px",
                    fontWeight: 600,
                    cursor: "pointer"
                  }}
                >
                  Clear
                </button>
              )}
            </div>
          </div>

          {/* Malpractice Student Records Management Card */}
          <div style={{ background: "var(--bg)", border: "1px solid var(--line)", borderRadius: "8px", padding: "12px", display: "flex", flexDirection: "column", gap: "8px" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", fontWeight: 700, color: "var(--ink)" }}>
                <ShieldAlert size={14} color="#d97706" /> Malpractice Record Entry
              </div>
              {malpracticeList.length > 0 ? (
                <span style={{ fontSize: "10px", background: "rgba(245, 158, 11, 0.15)", color: "#b45309", padding: "1px 6px", borderRadius: "10px", fontWeight: 700 }}>
                  {malpracticeList.length} MP Records
                </span>
              ) : (
                <span style={{ fontSize: "10px", background: "var(--panel)", color: "var(--muted)", padding: "1px 6px", borderRadius: "10px", border: "1px solid var(--line)" }}>
                  Optional
                </span>
              )}
            </div>

            <div style={{ fontSize: "11px", color: "var(--muted)", lineHeight: "1.35" }}>
              {malpracticeList.length > 0 ? (
                <span>
                  Active: <strong>{malpracticeFileName}</strong> ({malpracticeList.length} mapped). Mapped components show <strong style={{ color: "#b45309" }}>Malpractice (MP)</strong> &amp; <strong>Fail</strong>.
                </span>
              ) : (
                "Upload one or multiple malpractice reports across programmes (select multiple .xlsx/.xls files) to replace marks with \"Malpractice (MP)\" and mark courses as Fail."
              )}
            </div>

            <input 
              type="file" 
              ref={malpracticeFileInputRef}
              accept=".xlsx,.xls,.csv" 
              multiple
              onChange={handleMalpracticeExcelUpload}
              style={{ display: "none" }}
            />

            <div style={{ display: "flex", gap: "6px", marginTop: "2px" }}>
              <button 
                type="button"
                onClick={() => malpracticeFileInputRef.current?.click()}
                style={{
                  flex: 1,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: "4px",
                  padding: "5px 8px",
                  fontSize: "11px",
                  background: malpracticeList.length > 0 ? "rgba(245, 158, 11, 0.12)" : "var(--panel)",
                  color: malpracticeList.length > 0 ? "#b45309" : "var(--ink)",
                  border: malpracticeList.length > 0 ? "1px solid rgba(245, 158, 11, 0.3)" : "1px solid var(--line)",
                  borderRadius: "4px",
                  fontWeight: 600,
                  cursor: "pointer"
                }}
              >
                <FileUp size={12} /> {malpracticeList.length > 0 ? "Replace / Upload Files" : "Upload MP Excel"}
              </button>

              <button 
                type="button"
                onClick={handleDownloadMalpracticeTemplate}
                title="Download standard malpractice report template (.xlsx)"
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: "4px",
                  padding: "5px 8px",
                  fontSize: "11px",
                  background: "var(--panel)",
                  color: "var(--ink)",
                  border: "1px solid var(--line)",
                  borderRadius: "4px",
                  fontWeight: 600,
                  cursor: "pointer"
                }}
              >
                <FileDown size={12} /> Template
              </button>

              {malpracticeList.length > 0 && (
                <button 
                  type="button"
                  onClick={handleClearMalpracticeData}
                  title="Clear malpractice data and revert to original marks"
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    padding: "5px 8px",
                    fontSize: "11px",
                    background: "transparent",
                    color: "#b45309",
                    border: "1px solid #b45309",
                    borderRadius: "4px",
                    fontWeight: 600,
                    cursor: "pointer"
                  }}
                >
                  Clear
                </button>
              )}
            </div>
          </div>

          {/* Heldback Students Records Management Card */}
          <div style={{ background: "var(--bg)", border: "1px solid var(--line)", borderRadius: "8px", padding: "12px", display: "flex", flexDirection: "column", gap: "8px" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", fontWeight: 700, color: "var(--ink)" }}>
                <Lock size={14} color="#c026d3" /> Heldback Record Entry
              </div>
              {heldbackList.length > 0 ? (
                <span style={{ fontSize: "10px", background: "rgba(192, 38, 211, 0.15)", color: "#c026d3", padding: "1px 6px", borderRadius: "10px", fontWeight: 700 }}>
                  {heldbackList.length} Heldback
                </span>
              ) : (
                <span style={{ fontSize: "10px", background: "var(--panel)", color: "var(--muted)", padding: "1px 6px", borderRadius: "10px", border: "1px solid var(--line)" }}>
                  Optional
                </span>
              )}
            </div>

            <div style={{ fontSize: "11px", color: "var(--muted)", lineHeight: "1.35" }}>
              {heldbackList.length > 0 ? (
                <span>
                  Active: <strong>{heldbackFileName}</strong> ({heldbackList.length} mapped). Matched students are locked with <strong style={{ color: "#c026d3" }}>Held (Heldback)</strong> &amp; excluded from moderation/simulations.
                </span>
              ) : (
                "Upload one or multiple heldback reports across programmes (select multiple .xlsx/.xls files together) to lock matched students across papers from pass calculation, moderation marks, and pass simulation."
              )}
            </div>

            <input 
              type="file" 
              ref={heldbackFileInputRef}
              accept=".xlsx,.xls,.csv" 
              multiple
              onChange={handleHeldbackExcelUpload}
              style={{ display: "none" }}
            />

            <div style={{ display: "flex", gap: "6px", marginTop: "2px" }}>
              <button 
                type="button"
                onClick={() => heldbackFileInputRef.current?.click()}
                style={{
                  flex: 1,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: "4px",
                  padding: "5px 8px",
                  fontSize: "11px",
                  background: heldbackList.length > 0 ? "rgba(192, 38, 211, 0.12)" : "var(--panel)",
                  color: heldbackList.length > 0 ? "#c026d3" : "var(--ink)",
                  border: heldbackList.length > 0 ? "1px solid rgba(192, 38, 211, 0.3)" : "1px solid var(--line)",
                  borderRadius: "4px",
                  fontWeight: 600,
                  cursor: "pointer"
                }}
              >
                <FileUp size={12} /> {heldbackList.length > 0 ? "Replace / Upload Files" : "Upload Heldback"}
              </button>

              <button 
                type="button"
                onClick={handleDownloadHeldbackTemplate}
                title="Download standard heldback / APC report template (.xlsx)"
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  gap: "4px",
                  padding: "5px 8px",
                  fontSize: "11px",
                  background: "var(--panel)",
                  color: "var(--ink)",
                  border: "1px solid var(--line)",
                  borderRadius: "4px",
                  fontWeight: 600,
                  cursor: "pointer"
                }}
              >
                <FileDown size={12} /> Template
              </button>

              {heldbackList.length > 0 && (
                <button 
                  type="button"
                  onClick={handleClearHeldbackData}
                  title="Clear heldback data and restore student results"
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    padding: "5px 8px",
                    fontSize: "11px",
                    background: "transparent",
                    color: "#c026d3",
                    border: "1px solid #c026d3",
                    borderRadius: "4px",
                    fontWeight: 600,
                    cursor: "pointer"
                  }}
                >
                  Clear
                </button>
              )}
            </div>
          </div>

          {/* Consolidated Overall Result Stat Card (Global / No Filter) */}
          {processedRows.length > 0 && (
            <div style={{ background: "var(--bg)", border: "1px solid var(--line)", borderRadius: "8px", padding: "12px", display: "flex", flexDirection: "column", gap: "10px" }}>
              <div style={{ fontSize: "12px", fontWeight: 700, color: "var(--ink)", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <span style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                  <Users size={14} color="#6366f1" /> Consolidated Overview
                </span>
                <span style={{ fontSize: "9.5px", background: "rgba(99, 102, 241, 0.12)", color: "#6366f1", padding: "1px 6px", borderRadius: "10px", fontWeight: 700 }}>
                  All Data (No Filter)
                </span>
              </div>
              
              {/* Student Semester Level Cards */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px", fontSize: "11px" }}>
                <div style={{ background: "var(--panel)", padding: "8px", borderRadius: "6px", border: "1px solid var(--line)" }}>
                  <div style={{ color: "var(--muted)" }}>Unique Students</div>
                  <strong style={{ fontSize: "15px", color: "var(--ink)" }}>{consolidatedStudentMetrics.totalStudents}</strong>
                  <div style={{ fontSize: "9.5px", color: "var(--muted)", marginTop: "1px" }}>
                    {consolidatedStudentMetrics.totalPapersAttempted} papers (~{consolidatedStudentMetrics.avgPapersPerStudent}/st)
                  </div>
                </div>

                <div style={{ background: "rgba(99, 102, 241, 0.08)", padding: "8px", borderRadius: "6px", border: "1px solid rgba(99, 102, 241, 0.25)" }}>
                  <div style={{ color: "#6366f1", fontWeight: 600 }}>Semester Pass %</div>
                  <strong style={{ fontSize: "15px", color: "#6366f1" }}>{consolidatedStudentMetrics.finalPassedPct}%</strong>
                  <div style={{ fontSize: "9.5px", color: "#6366f1" }}>All Papers Rule</div>
                </div>

                <div style={{ background: "rgba(16, 185, 129, 0.1)", padding: "8px", borderRadius: "6px", border: "1px solid rgba(16, 185, 129, 0.3)" }}>
                  <div style={{ color: "#10b981", fontWeight: 600 }}>Passed All Papers</div>
                  <strong style={{ fontSize: "15px", color: "#10b981" }}>{consolidatedStudentMetrics.finalPassedStudents}</strong>
                  <div style={{ fontSize: "9.5px", color: "#10b981" }}>({consolidatedStudentMetrics.finalPassedPct}%)</div>
                </div>

                <div style={{ background: "rgba(239, 68, 68, 0.1)", padding: "8px", borderRadius: "6px", border: "1px solid rgba(239, 68, 68, 0.3)" }}>
                  <div style={{ color: "#ef4444", fontWeight: 600 }}>Failed &ge; 1 Paper</div>
                  <strong style={{ fontSize: "15px", color: "#ef4444" }}>{consolidatedStudentMetrics.failedStudents}</strong>
                  <div style={{ fontSize: "9.5px", color: "#ef4444" }}>({consolidatedStudentMetrics.failedPct}%)</div>
                </div>
              </div>

              {/* Student Semester Moderation Impact */}
              <div style={{ background: "var(--panel)", padding: "8px", borderRadius: "6px", border: "1px solid var(--line)", fontSize: "11px", display: "flex", flexDirection: "column", gap: "4px" }}>
                <div style={{ fontWeight: 600, color: "var(--ink)", display: "flex", alignItems: "center", gap: "4px" }}>
                  <Zap size={12} color="#f59e0b" /> Semester Moderation Impact:
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)" }}>
                  <span>Raw Passed (0 Mod):</span>
                  <strong>{consolidatedStudentMetrics.rawPassedStudents} ({consolidatedStudentMetrics.rawPassedPct}%)</strong>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", color: "#10b981", fontWeight: 600 }}>
                  <span>Rescued to Semester Pass:</span>
                  <span>+{consolidatedStudentMetrics.rescuedStudents} students ({consolidatedStudentMetrics.rescuedPct}%)</span>
                </div>
              </div>

              {/* Student Cases & Special Statuses */}
              {(consolidatedStudentMetrics.heldStudents > 0 || consolidatedStudentMetrics.absentStudents > 0 || consolidatedStudentMetrics.malpracticeStudents > 0) && (
                <div style={{ background: "var(--panel)", padding: "8px", borderRadius: "6px", border: "1px solid var(--line)", fontSize: "11px", display: "flex", flexDirection: "column", gap: "4px" }}>
                  <div style={{ fontWeight: 600, color: "var(--ink)", display: "flex", alignItems: "center", gap: "4px" }}>
                    <AlertCircle size={12} color="#c026d3" /> Special Student Statuses:
                  </div>
                  {consolidatedStudentMetrics.heldStudents > 0 && (
                    <div style={{ display: "flex", justifyContent: "space-between", color: "#c026d3" }}>
                      <span>Held Students (Total):</span>
                      <strong>{consolidatedStudentMetrics.heldStudents} ({consolidatedStudentMetrics.heldPct}%)</strong>
                    </div>
                  )}
                  {consolidatedStudentMetrics.heldbackStudents > 0 && (
                    <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)", paddingLeft: "8px" }}>
                      <span>&bull; Heldback Report:</span>
                      <strong style={{ color: "#c026d3" }}>{consolidatedStudentMetrics.heldbackStudents}</strong>
                    </div>
                  )}
                  {consolidatedStudentMetrics.heldMissingStudents > 0 && (
                    <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)", paddingLeft: "8px" }}>
                      <span>&bull; Missing Component:</span>
                      <strong style={{ color: "#9333ea" }}>{consolidatedStudentMetrics.heldMissingStudents}</strong>
                    </div>
                  )}
                  {consolidatedStudentMetrics.absentStudents > 0 && (
                    <div style={{ display: "flex", justifyContent: "space-between", color: "#dc2626" }}>
                      <span>With Absences (&ge;1 Paper):</span>
                      <strong>{consolidatedStudentMetrics.absentStudents} ({consolidatedStudentMetrics.absentPct}%)</strong>
                    </div>
                  )}
                  {consolidatedStudentMetrics.malpracticeStudents > 0 && (
                    <div style={{ display: "flex", justifyContent: "space-between", color: "#d97706" }}>
                      <span>With Malpractice (&ge;1 Paper):</span>
                      <strong>{consolidatedStudentMetrics.malpracticeStudents} ({consolidatedStudentMetrics.malpracticePct}%)</strong>
                    </div>
                  )}
                </div>
              )}

              {/* Course-Level Statistics Summary */}
              <div style={{ background: "var(--panel)", padding: "8px", borderRadius: "6px", border: "1px solid var(--line)", fontSize: "11px", display: "flex", flexDirection: "column", gap: "4px" }}>
                <div style={{ fontWeight: 600, color: "var(--ink)", display: "flex", alignItems: "center", gap: "4px" }}>
                  <BookOpen size={12} color="var(--accent)" /> Course-Level Aggregate:
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)" }}>
                  <span>Total Course Papers:</span>
                  <strong>{consolidatedCourseMetrics.total}</strong>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", color: "#10b981" }}>
                  <span>Course Papers Passed:</span>
                  <strong>{consolidatedCourseMetrics.totalPassed} ({consolidatedCourseMetrics.passPct}%)</strong>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)", fontSize: "10.5px" }}>
                  <span>&bull; Raw / Via Mod:</span>
                  <span><strong>{consolidatedCourseMetrics.rawPassed}</strong> / <strong style={{ color: "#10b981" }}>+{consolidatedCourseMetrics.moderatedPassed}</strong></span>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", color: "#ef4444" }}>
                  <span>Course Papers Failed:</span>
                  <strong>{consolidatedCourseMetrics.failed} ({consolidatedCourseMetrics.failedPct}%)</strong>
                </div>
                {consolidatedCourseMetrics.heldbackCount > 0 && (
                  <div style={{ display: "flex", justifyContent: "space-between", color: "#c026d3" }}>
                    <span>Held (Heldback Report):</span>
                    <strong>{consolidatedCourseMetrics.heldbackCount}</strong>
                  </div>
                )}
                {consolidatedCourseMetrics.missingCompCount > 0 && (
                  <div style={{ display: "flex", justifyContent: "space-between", color: "#9333ea" }}>
                    <span>Held (Missing Component):</span>
                    <strong>{consolidatedCourseMetrics.missingCompCount}</strong>
                  </div>
                )}
                {consolidatedCourseMetrics.absentCount > 0 && (
                  <div style={{ display: "flex", justifyContent: "space-between", color: "#dc2626" }}>
                    <span>Absent Papers:</span>
                    <strong>{consolidatedCourseMetrics.absentCount}</strong>
                  </div>
                )}
                {consolidatedCourseMetrics.malpracticeCount > 0 && (
                  <div style={{ display: "flex", justifyContent: "space-between", color: "#d97706" }}>
                    <span>Malpractice Papers:</span>
                    <strong>{consolidatedCourseMetrics.malpracticeCount}</strong>
                  </div>
                )}
              </div>

              {/* Failure Breakdown */}
              {consolidatedCourseMetrics.failed > 0 && (
                <div style={{ background: "var(--panel)", padding: "8px", borderRadius: "6px", border: "1px solid var(--line)", fontSize: "11px", display: "flex", flexDirection: "column", gap: "4px" }}>
                  <div style={{ fontWeight: 600, color: "var(--ink)" }}>Failure Breakdown:</div>
                  <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)" }}>
                    <span>ESE Failed (&lt; 30%):</span>
                    <strong style={{ color: "#ef4444" }}>{consolidatedCourseMetrics.eseFailed}</strong>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)" }}>
                    <span>Aggregate Failed (&lt; 35%):</span>
                    <strong style={{ color: "#ef4444" }}>{consolidatedCourseMetrics.overallFailed}</strong>
                  </div>
                </div>
              )}

              {/* Course Pass Simulation (+0 to +10) Card */}
              <div style={{ background: "rgba(16, 185, 129, 0.08)", padding: "10px", borderRadius: "6px", border: "1px solid rgba(16, 185, 129, 0.25)", fontSize: "11px", display: "flex", flexDirection: "column", gap: "6px" }}>
                <div style={{ fontWeight: 700, color: "#10b981", display: "flex", alignItems: "center", gap: "5px" }}>
                  <TrendingUp size={13} /> Pass Simulation (+0 to +10):
                </div>
                <div style={{ color: "var(--muted)", lineHeight: "1.3" }}>
                  Pass counts calculated for every course across 0 to +10 moderation marks.
                </div>
                <div style={{ display: "flex", gap: "6px", marginTop: "2px" }}>
                  <button
                    type="button"
                    onClick={handleExportSimulationExcel}
                    style={{
                      flex: 1,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: "4px",
                      padding: "5px 8px",
                      fontSize: "11px",
                      background: "#10b981",
                      color: "white",
                      border: "none",
                      borderRadius: "4px",
                      fontWeight: 600,
                      cursor: "pointer"
                    }}
                  >
                    <Download size={12} /> Excel Report
                  </button>
                  <button
                    type="button"
                    onClick={() => setActiveTab("matrix")}
                    style={{
                      flex: 1,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: "4px",
                      padding: "5px 8px",
                      fontSize: "11px",
                      background: "var(--panel)",
                      color: "var(--ink)",
                      border: "1px solid var(--line)",
                      borderRadius: "4px",
                      fontWeight: 600,
                      cursor: "pointer"
                    }}
                  >
                    View Matrix
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* Quick Rules Summary */}
          <div style={{ background: "var(--bg)", border: "1px solid var(--line)", borderRadius: "8px", padding: "12px", fontSize: "11.5px", color: "var(--muted)", lineHeight: "1.4" }}>
            <strong style={{ color: "var(--ink)", display: "block", marginBottom: "6px" }}>Evaluation Rules:</strong>
            <ul style={{ margin: 0, paddingLeft: "16px", display: "flex", flexDirection: "column", gap: "4px" }}>
              <li><strong>ESE Min:</strong> <code>ceil(30% × ESE Max)</code></li>
              <li><strong>Overall Min:</strong> <code>ceil(35% × Overall Max)</code></li>
              <li><strong>Pass Condition:</strong> <code>ESE Pass AND Overall Pass</code></li>
              <li><strong>Moderation Rule:</strong> <code>marks_needed = max(ESE Deficit, Overall Deficit)</code></li>
              <li><strong>ESE-TH Only (Default):</strong> Moderation applies only to courses with ESE Theory (TH) unless PR-only option is enabled.</li>
            </ul>
          </div>

          {/* Status Message */}
          <div style={{ marginTop: "auto", padding: "8px 12px", borderRadius: "6px", fontSize: "11.5px", background: statusType === "error" ? "var(--danger-soft)" : statusType === "success" ? "var(--accent-soft)" : "var(--bg)", color: statusType === "error" ? "var(--danger)" : statusType === "success" ? "var(--accent)" : "var(--muted)", border: "1px solid var(--line)" }}>
            {statusMsg}
          </div>

        </aside>

        {/* Right Content / Dynamic View (Results Table or Moderation Matrix) */}
        <main style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden", background: "var(--bg)" }}>
          
          {activeTab === "students" ? (
            /* Student-Wise Semester Results Panel */
            <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden", padding: "16px", gap: "14px" }}>
              
              {/* Header & Actions */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "10px", background: "var(--panel)", padding: "14px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                <div>
                  <h3 style={{ fontSize: "15px", fontWeight: 700, margin: "0 0 4px", display: "flex", alignItems: "center", gap: "6px" }}>
                    <Users size={18} color="var(--accent)" /> Student Semester Results & Breakdown
                  </h3>
                  <p style={{ margin: 0, fontSize: "12px", color: "var(--muted)" }}>
                    Rule: A student passes the semester <strong>only if they pass ALL attempted papers</strong> in that term. Shows PRN/Seat, paper counts, moderation benefit, and course breakdown.
                  </p>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap" }}>
                  <button 
                    type="button"
                    onClick={() => handleExportStudentSemesterExcel(filteredStudents)}
                    style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", padding: "7px 14px", background: "var(--accent)", color: "white", border: "none", borderRadius: "6px", cursor: "pointer", fontWeight: 600 }}
                    title={filteredStudents.length !== studentSemesterData.length ? `Export current filtered view (${filteredStudents.length} students)` : "Export all student semester results"}
                  >
                    <Download size={14} /> Export Current View ({filteredStudents.length} Students)
                  </button>
                  {filteredStudents.length !== studentSemesterData.length && (
                    <button 
                      type="button"
                      onClick={() => handleExportStudentSemesterExcel(studentSemesterData, "student_semester_results_all.xlsx")}
                      style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", padding: "7px 12px", background: "transparent", color: "var(--ink)", border: "1px solid var(--line)", borderRadius: "6px", cursor: "pointer", fontWeight: 600 }}
                      title="Export all student records without filters (.xlsx)"
                    >
                      <Download size={14} /> Export All ({studentSemesterData.length})
                    </button>
                  )}
                </div>
              </div>

              {/* Student KPI Summary Cards */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: "10px" }}>
                <div style={{ background: "var(--panel)", padding: "12px 14px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                  <div style={{ fontSize: "11px", color: "var(--muted)", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                    <Users size={14} color="var(--accent)" /> Unique Students
                  </div>
                  <div style={{ fontSize: "20px", fontWeight: 700, color: "var(--ink)", marginTop: "3px" }}>
                    {studentMetrics.totalStudents}
                  </div>
                  <div style={{ fontSize: "10.5px", color: "var(--muted)" }}>
                    Total: {studentMetrics.totalPapersAttempted} papers (~{studentMetrics.avgPapersPerStudent}/student)
                  </div>
                </div>

                <div style={{ background: "rgba(16, 185, 129, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(16, 185, 129, 0.3)" }}>
                  <div style={{ fontSize: "11px", color: "#10b981", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                    <CheckCircle2 size={14} /> Passed Semester (All Papers)
                  </div>
                  <div style={{ fontSize: "20px", fontWeight: 700, color: "#10b981", marginTop: "3px" }}>
                    {studentMetrics.finalPassedStudents}
                  </div>
                  <div style={{ fontSize: "10.5px", color: "#10b981", fontWeight: 600 }}>
                    {studentMetrics.finalPassedPct}% of total students
                  </div>
                </div>

                <div style={{ background: "rgba(239, 68, 68, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(239, 68, 68, 0.3)" }}>
                  <div style={{ fontSize: "11px", color: "#ef4444", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                    <XCircle size={14} /> Failed Semester (≥1 Paper)
                  </div>
                  <div style={{ fontSize: "20px", fontWeight: 700, color: "#ef4444", marginTop: "3px" }}>
                    {studentMetrics.failedStudents}
                  </div>
                  <div style={{ fontSize: "10.5px", color: "#ef4444", fontWeight: 600 }}>
                    {studentMetrics.failedPct}% of total students
                  </div>
                </div>

                <div style={{ background: "rgba(245, 158, 11, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(245, 158, 11, 0.3)" }}>
                  <div style={{ fontSize: "11px", color: "#f59e0b", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                    <Sparkles size={14} /> Rescued to Pass with Mod
                  </div>
                  <div style={{ fontSize: "20px", fontWeight: 700, color: "#f59e0b", marginTop: "3px" }}>
                    +{studentMetrics.rescuedStudents}
                  </div>
                  <div style={{ fontSize: "10.5px", color: "#f59e0b", fontWeight: 600 }}>
                    {studentMetrics.rescuedPct}% students rescued to pass
                  </div>
                </div>

                <div style={{ background: "var(--panel)", padding: "12px 14px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                  <div style={{ fontSize: "11px", color: "var(--muted)", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                    <Calculator size={14} color="var(--muted)" /> Raw Passed (0 Mod)
                  </div>
                  <div style={{ fontSize: "20px", fontWeight: 700, color: "var(--ink)", marginTop: "3px" }}>
                    {studentMetrics.rawPassedStudents}
                  </div>
                  <div style={{ fontSize: "10.5px", color: "var(--muted)" }}>
                    {studentMetrics.rawPassedPct}% without moderation
                  </div>
                </div>

                {studentMetrics.heldStudents > 0 && (
                  <div style={{ background: "rgba(192, 38, 211, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(192, 38, 211, 0.3)" }}>
                    <div style={{ fontSize: "11px", color: "#c026d3", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                      <Lock size={14} /> Held Students (Total)
                    </div>
                    <div style={{ fontSize: "20px", fontWeight: 700, color: "#c026d3", marginTop: "3px" }}>
                      {studentMetrics.heldStudents}
                    </div>
                    <div style={{ fontSize: "10.5px", color: "#c026d3" }}>
                      {studentMetrics.heldbackStudents > 0 ? `${studentMetrics.heldbackStudents} HB` : ""}{studentMetrics.heldbackStudents > 0 && studentMetrics.heldMissingStudents > 0 ? " · " : ""}{studentMetrics.heldMissingStudents > 0 ? `${studentMetrics.heldMissingStudents} Missing` : ""} ({studentMetrics.heldPct}%)
                    </div>
                  </div>
                )}

                {studentMetrics.absentStudents > 0 && (
                  <div style={{ background: "rgba(220, 38, 38, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(220, 38, 38, 0.3)" }}>
                    <div style={{ fontSize: "11px", color: "#dc2626", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                      <AlertCircle size={14} /> With Absences
                    </div>
                    <div style={{ fontSize: "20px", fontWeight: 700, color: "#dc2626", marginTop: "3px" }}>
                      {studentMetrics.absentStudents}
                    </div>
                    <div style={{ fontSize: "10.5px", color: "#dc2626" }}>
                      {studentMetrics.absentPct}% in ≥1 paper
                    </div>
                  </div>
                )}

                {studentMetrics.malpracticeStudents > 0 && (
                  <div style={{ background: "rgba(217, 119, 6, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(217, 119, 6, 0.3)" }}>
                    <div style={{ fontSize: "11px", color: "#d97706", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                      <AlertCircle size={14} /> With Malpractice
                    </div>
                    <div style={{ fontSize: "20px", fontWeight: 700, color: "#d97706", marginTop: "3px" }}>
                      {studentMetrics.malpracticeStudents}
                    </div>
                    <div style={{ fontSize: "10.5px", color: "#d97706" }}>
                      {studentMetrics.malpracticePct}% in ≥1 paper
                    </div>
                  </div>
                )}
              </div>

              {/* Filters & Search Toolbar */}
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "10px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap", flex: 1 }}>
                  <div style={{ position: "relative", minWidth: "220px", flex: 1, maxWidth: "320px" }}>
                    <Search size={14} style={{ position: "absolute", left: "10px", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                    <input 
                      type="text"
                      placeholder="Search by PRN, Seat Number, Program, Course..."
                      value={studentSearchQuery}
                      onChange={(e) => setStudentSearchQuery(e.target.value)}
                      style={{ width: "100%", padding: "6px 10px 6px 30px", fontSize: "12px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--panel)" }}
                    />
                    {studentSearchQuery && (
                      <X 
                        size={13} 
                        onClick={() => setStudentSearchQuery("")}
                        style={{ position: "absolute", right: "10px", top: "50%", transform: "translateY(-50%)", cursor: "pointer", color: "var(--muted)" }}
                      />
                    )}
                  </div>

                  {/* College Filter */}
                  {uniqueStudentColleges.length > 0 && (
                    <select 
                      value={studentCollegeFilter} 
                      onChange={(e) => setStudentCollegeFilter(e.target.value)}
                      style={{ padding: "6px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--panel)", maxWidth: "200px" }}
                      title="Filter students by College"
                    >
                      <option value="ALL">All Colleges ({uniqueStudentColleges.length})</option>
                      {uniqueStudentColleges.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  )}

                  {/* Program Filter */}
                  {uniqueStudentPrograms.length > 0 && (
                    <select 
                      value={studentProgramFilter} 
                      onChange={(e) => { 
                        setStudentProgramFilter(e.target.value); 
                        setStudentCourseFilter("ALL"); 
                      }}
                      style={{ padding: "6px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--panel)", maxWidth: "220px" }}
                      title="Filter students by Programme"
                    >
                      <option value="ALL">All Programs ({uniqueStudentPrograms.length})</option>
                      {uniqueStudentPrograms.map(p => <option key={p} value={p}>{p}</option>)}
                    </select>
                  )}

                  {/* Course Filter */}
                  {uniqueStudentCourses.length > 0 && (
                    <select 
                      value={studentCourseFilter} 
                      onChange={(e) => setStudentCourseFilter(e.target.value)}
                      style={{ padding: "6px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--panel)", maxWidth: "260px" }}
                      title="Filter students who attempted specific Course"
                    >
                      <option value="ALL">All Courses ({uniqueStudentCourses.length})</option>
                      {uniqueStudentCourses.map(c => <option key={c.code} value={c.code}>{c.label}</option>)}
                    </select>
                  )}

                  {(studentCollegeFilter !== "ALL" || studentProgramFilter !== "ALL" || studentCourseFilter !== "ALL" || studentFilterStatus !== "ALL" || studentSearchQuery) && (
                    <button 
                      type="button" 
                      onClick={() => {
                        setStudentFilterStatus("ALL");
                        setStudentSearchQuery("");
                        setStudentCollegeFilter("ALL");
                        setStudentProgramFilter("ALL");
                        setStudentCourseFilter("ALL");
                      }}
                      style={{ display: "flex", alignItems: "center", gap: "4px", padding: "5px 9px", fontSize: "11px", color: "var(--danger)", background: "var(--danger-soft)", border: "none", borderRadius: "5px", cursor: "pointer", fontWeight: 600 }}
                    >
                      <RefreshCw size={11} /> Reset Filters
                    </button>
                  )}
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: "4px", background: "var(--panel)", padding: "3px", borderRadius: "6px", border: "1px solid var(--line)", flexWrap: "wrap" }}>
                  <button 
                    type="button"
                    onClick={() => setStudentFilterStatus("ALL")}
                    style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "ALL" ? "var(--accent)" : "transparent", color: studentFilterStatus === "ALL" ? "white" : "var(--muted)" }}
                  >
                    All ({studentMetrics.totalStudents})
                  </button>
                  <button 
                    type="button"
                    onClick={() => setStudentFilterStatus("PASS")}
                    style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "PASS" ? "#10b981" : "transparent", color: studentFilterStatus === "PASS" ? "white" : "var(--muted)" }}
                  >
                    Passed Semester ({studentMetrics.finalPassedStudents})
                  </button>
                  <button 
                    type="button"
                    onClick={() => setStudentFilterStatus("FAIL")}
                    style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "FAIL" ? "#ef4444" : "transparent", color: studentFilterStatus === "FAIL" ? "white" : "var(--muted)" }}
                  >
                    Failed Semester ({studentMetrics.failedStudents})
                  </button>
                  <button 
                    type="button"
                    onClick={() => setStudentFilterStatus("RESCUED")}
                    style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "RESCUED" ? "#f59e0b" : "transparent", color: studentFilterStatus === "RESCUED" ? "white" : "var(--muted)" }}
                  >
                    Rescued ({studentMetrics.rescuedStudents})
                  </button>
                  {studentMetrics.heldStudents > 0 && (
                    <button 
                      type="button"
                      onClick={() => setStudentFilterStatus("HELD")}
                      style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "HELD" ? "#9333ea" : "transparent", color: studentFilterStatus === "HELD" ? "white" : "#9333ea" }}
                    >
                      Held ({studentMetrics.heldStudents})
                    </button>
                  )}
                  {studentMetrics.heldbackStudents > 0 && (
                    <button 
                      type="button"
                      onClick={() => setStudentFilterStatus("HELDBACK")}
                      style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "HELDBACK" ? "#c026d3" : "transparent", color: studentFilterStatus === "HELDBACK" ? "white" : "#c026d3" }}
                    >
                      Heldback ({studentMetrics.heldbackStudents})
                    </button>
                  )}
                  {studentMetrics.absentStudents > 0 && (
                    <button 
                      type="button"
                      onClick={() => setStudentFilterStatus("ABSENT")}
                      style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "ABSENT" ? "#dc2626" : "transparent", color: studentFilterStatus === "ABSENT" ? "white" : "#dc2626" }}
                    >
                      With Absences ({studentMetrics.absentStudents})
                    </button>
                  )}
                  {studentMetrics.malpracticeStudents > 0 && (
                    <button 
                      type="button"
                      onClick={() => setStudentFilterStatus("MALPRACTICE")}
                      style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "MALPRACTICE" ? "#d97706" : "transparent", color: studentFilterStatus === "MALPRACTICE" ? "white" : "#d97706" }}
                    >
                      With Malpractice ({studentMetrics.malpracticeStudents})
                    </button>
                  )}
                </div>
              </div>

              {/* Student Semester Table */}
              <div style={{ flex: 1, overflow: "auto", background: "var(--panel)", border: "1px solid var(--line)", borderRadius: "8px" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px", textAlign: "left" }}>
                  <thead style={{ position: "sticky", top: 0, background: "var(--bg)", borderBottom: "2px solid var(--line)", zIndex: 10 }}>
                    <tr>
                      <th style={{ padding: "10px 12px", width: "40px" }}></th>
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600, width: "120px" }}>Seat Number</th>
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600, width: "140px" }}>PRN</th>
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600 }}>Program / Term</th>
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center", width: "100px" }}>Attempted</th>
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center", width: "90px" }}>Passed</th>
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center", width: "90px" }}>Failed</th>
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center", width: "110px" }}>Mod Marks</th>
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center", width: "110px" }}>Raw Result</th>
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center", width: "140px" }}>Semester Result</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredStudents.length === 0 ? (
                      <tr>
                        <td colSpan={10} style={{ padding: "36px", textAlign: "center", color: "var(--muted)" }}>
                          {studentSemesterData.length === 0 
                            ? "No student data loaded. Upload an ADES Marks Excel sheet to begin." 
                            : "No students match your filter criteria."}
                        </td>
                      </tr>
                    ) : (
                      filteredStudents.map((st) => {
                        const isExpanded = expandedStudents[st.key];
                        return (
                          <React.Fragment key={st.key}>
                            <tr 
                              onClick={() => toggleStudentExpand(st.key)}
                              style={{ 
                                borderBottom: isExpanded ? "none" : "1px solid var(--line)", 
                                cursor: "pointer",
                                background: isExpanded 
                                  ? "rgba(59, 130, 246, 0.05)" 
                                  : st.isHeldback
                                    ? "rgba(192, 38, 211, 0.05)"
                                    : st.isHeld 
                                      ? "rgba(147, 51, 234, 0.04)" 
                                      : "transparent"
                              }}
                            >
                              <td style={{ padding: "10px 12px", textAlign: "center", color: "var(--muted)" }}>
                                {isExpanded ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
                              </td>
                              <td style={{ padding: "10px 12px", fontWeight: 600, color: "var(--ink)" }}>
                                {st.seatNumber || "-"}
                              </td>
                              <td style={{ padding: "10px 12px", fontWeight: 600, color: "var(--accent)" }}>
                                {st.prn || "-"}
                              </td>
                              <td style={{ padding: "10px 12px", color: "var(--muted)" }}>
                                <div style={{ color: "var(--ink)", fontWeight: 500 }}>{st.program || "-"}</div>
                                {st.faculty && <div style={{ fontSize: "10.5px", color: "var(--muted)" }}>{st.faculty}</div>}
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center", fontWeight: 600 }}>
                                {st.totalCourses} Papers
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center", color: "#10b981", fontWeight: 600 }}>
                                {st.finalPassedCourses}
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center", color: st.finalFailedCourses > 0 ? "#ef4444" : "var(--muted)", fontWeight: st.finalFailedCourses > 0 ? 600 : 400 }}>
                                {st.finalFailedCourses}
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center" }}>
                                {st.totalModerationMarks > 0 ? (
                                  <span style={{ fontSize: "11px", fontWeight: 700, color: "#f59e0b", background: "rgba(245, 158, 11, 0.12)", padding: "2px 6px", borderRadius: "4px" }}>
                                    +{st.totalModerationMarks} Marks
                                  </span>
                                ) : (
                                  <span style={{ color: "var(--muted)", fontSize: "11px" }}>0</span>
                                )}
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center" }}>
                                <span style={{ 
                                  fontSize: "11px", 
                                  fontWeight: 600, 
                                  color: st.isHeldback ? "#c026d3" : st.isHeld ? "#7e22ce" : st.rawSemesterPass ? "#10b981" : "#ef4444" 
                                }}>
                                  {st.rawSemesterResult}
                                </span>
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center" }}>
                                {st.isHeldback ? (
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(192, 38, 211, 0.15)", color: "#c026d3", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }} title={st.heldbackReason || "Heldback at term-level"}>
                                    <Lock size={12} /> Held (Heldback)
                                  </span>
                                ) : st.isHeld ? (
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(147, 51, 234, 0.15)", color: "#7e22ce", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }}>
                                    <AlertTriangle size={12} /> Held ({st.heldCourses} {st.heldCourses === 1 ? "Paper" : "Papers"} Missing)
                                  </span>
                                ) : st.isRescuedSemester ? (
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(245, 158, 11, 0.15)", color: "#d97706", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }}>
                                    <Sparkles size={12} /> Pass (Rescued)
                                  </span>
                                ) : st.finalSemesterPass ? (
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(16, 185, 129, 0.15)", color: "#10b981", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }}>
                                    <CheckCircle2 size={12} /> Pass
                                  </span>
                                ) : st.absentCourses > 0 ? (
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(239, 68, 68, 0.15)", color: "#dc2626", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }}>
                                    <UserX size={12} /> Fail ({st.absentCourses} Absent{st.finalFailedCourses > st.absentCourses ? `, ${st.finalFailedCourses - st.absentCourses} failed` : ""})
                                  </span>
                                ) : (
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(239, 68, 68, 0.12)", color: "#ef4444", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }}>
                                    <XCircle size={12} /> Fail ({st.finalFailedCourses} Failed)
                                  </span>
                                )}
                              </td>
                            </tr>

                            {/* Nested Course Breakdown Row */}
                            {isExpanded && (
                              <tr style={{ background: "rgba(59, 130, 246, 0.03)", borderBottom: "1px solid var(--line)" }}>
                                <td colSpan={10} style={{ padding: "10px 14px 16px 40px" }}>
                                  <div style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: "6px", overflow: "hidden" }}>
                                    <div style={{ padding: "6px 12px", background: "var(--bg)", fontSize: "11px", fontWeight: 700, color: "var(--muted)", borderBottom: "1px solid var(--line)" }}>
                                      COURSE-BY-COURSE BREAKDOWN FOR {st.seatNumber || st.prn} ({st.courses.length} Attempted Papers)
                                    </div>
                                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "11.5px" }}>
                                      <thead>
                                        <tr style={{ borderBottom: "1px solid var(--line)", background: "rgba(0,0,0,0.02)" }}>
                                          <th style={{ padding: "6px 10px", textAlign: "left", color: "var(--muted)" }}>Course Code</th>
                                          <th style={{ padding: "6px 10px", textAlign: "left", color: "var(--muted)" }}>Course Name</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>ESE (Marks / Min)</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>Overall (Marks / Min)</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>ESE Pass</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>Overall Pass</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>Mod Awarded</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>Course Result</th>
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {st.courses.map((c, cIdx) => (
                                          <tr key={cIdx} style={{ borderBottom: "1px solid var(--line)", background: c.coursePass === "Pass" ? "transparent" : c.isHeldback ? "rgba(192, 38, 211, 0.05)" : c.isHeld ? "rgba(147, 51, 234, 0.04)" : "rgba(239, 68, 68, 0.03)" }}>
                                            <td style={{ padding: "6px 10px", fontWeight: 600 }}>{c.courseCode}</td>
                                            <td style={{ padding: "6px 10px" }}>{c.courseName}</td>
                                            <td style={{ padding: "6px 10px", textAlign: "center" }}>
                                              {c.isHeldback ? (
                                                <span style={{ color: "#c026d3", fontWeight: 700, background: "rgba(192, 38, 211, 0.12)", padding: "1px 6px", borderRadius: "4px" }}>
                                                  Held
                                                </span>
                                              ) : c.isHeld ? (
                                                <span style={{ color: "#7e22ce", fontWeight: 700, background: "rgba(147, 51, 234, 0.12)", padding: "1px 6px", borderRadius: "4px" }}>
                                                  Held
                                                </span>
                                              ) : c.isAbsent ? (
                                                <span style={{ color: "#dc2626", fontWeight: 700, background: "rgba(239, 68, 68, 0.12)", padding: "1px 6px", borderRadius: "4px" }}>
                                                  Absent (Ab)
                                                </span>
                                              ) : c.eseOverall !== null ? (
                                                `${c.eseOverall} / ${c.eseMin ?? "-"}`
                                              ) : "-"}
                                            </td>
                                            <td style={{ padding: "6px 10px", textAlign: "center" }}>
                                              {c.isHeldback ? (
                                                <span style={{ color: "#c026d3", fontWeight: 700, background: "rgba(192, 38, 211, 0.12)", padding: "1px 6px", borderRadius: "4px" }}>
                                                  Held
                                                </span>
                                              ) : c.isHeld ? (
                                                <span style={{ color: "#7e22ce", fontWeight: 700, background: "rgba(147, 51, 234, 0.12)", padding: "1px 6px", borderRadius: "4px" }}>
                                                  Held
                                                </span>
                                              ) : c.courseOverall !== null ? (
                                                `${c.courseOverall} / ${c.overallMin ?? "-"}`
                                              ) : "-"}
                                            </td>
                                            <td style={{ padding: "6px 10px", textAlign: "center", fontWeight: 600, color: c.isHeldback ? "#c026d3" : c.isHeld ? "#7e22ce" : c.esePass === "Pass" ? "#10b981" : "#ef4444" }}>
                                              {c.esePass || "-"}
                                            </td>
                                            <td style={{ padding: "6px 10px", textAlign: "center", fontWeight: 600, color: c.isHeldback ? "#c026d3" : c.isHeld ? "#7e22ce" : c.overallPass === "Pass" ? "#10b981" : "#ef4444" }}>
                                              {c.overallPass || "-"}
                                            </td>
                                            <td style={{ padding: "6px 10px", textAlign: "center" }}>
                                              {c.modMarks > 0 ? (
                                                <span style={{ color: "#f59e0b", fontWeight: 700 }}>+{c.modMarks}</span>
                                              ) : (
                                                <span style={{ color: "var(--muted)" }}>0</span>
                                              )}
                                            </td>
                                            <td style={{ padding: "6px 10px", textAlign: "center" }}>
                                              {c.isHeldback ? (
                                                <span style={{ color: "#c026d3", fontWeight: 700, background: "rgba(192, 38, 211, 0.15)", padding: "2px 6px", borderRadius: "4px", display: "inline-flex", alignItems: "center", gap: "3px" }} title={c.heldbackReason || "Heldback at term-level"}>
                                                  <Lock size={11} /> Held (Heldback)
                                                </span>
                                              ) : c.isHeld ? (
                                                <span style={{ color: "#7e22ce", fontWeight: 700, background: "rgba(147, 51, 234, 0.15)", padding: "2px 6px", borderRadius: "4px", display: "inline-flex", alignItems: "center", gap: "3px" }} title={`Missing: ${c.missingComponents && c.missingComponents.length > 0 ? c.missingComponents.join(", ") : "Required Component"}`}>
                                                  <AlertTriangle size={11} /> Held (Missing {c.missingComponents && c.missingComponents.length > 0 ? c.missingComponents.join(", ") : "Component"})
                                                </span>
                                              ) : c.isMalpractice ? (
                                                <span style={{ color: "#b45309", fontWeight: 700, background: "rgba(245, 158, 11, 0.15)", padding: "2px 6px", borderRadius: "4px", display: "inline-flex", alignItems: "center", gap: "3px" }} title={`Status: ${c.malpracticeStatus || "EHB"}${c.malpracticeRemarks ? ` | ${c.malpracticeRemarks}` : ""}${c.malpracticeDate ? ` | ${c.malpracticeDate}` : ""}`}>
                                                  <ShieldAlert size={11} /> Fail (MP)
                                                </span>
                                              ) : c.isAbsent ? (
                                                <span style={{ color: "#dc2626", fontWeight: 700, background: "rgba(239, 68, 68, 0.12)", padding: "2px 6px", borderRadius: "4px", display: "inline-flex", alignItems: "center", gap: "3px" }} title={c.absentStatus || "Marked Absent"}>
                                                  <UserX size={11} /> Fail (Absent)
                                                </span>
                                              ) : c.isModeratedPass ? (
                                                <span style={{ color: "#d97706", fontWeight: 700 }}>Pass (Mod)</span>
                                              ) : c.coursePass === "Pass" ? (
                                                <span style={{ color: "#10b981", fontWeight: 600 }}>Pass</span>
                                              ) : (
                                                <span style={{ color: "#ef4444", fontWeight: 600 }}>Fail</span>
                                              )}
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          ) : activeTab === "moderation" ? (
            /* Moderation Management Panel */
            <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden", padding: "16px", gap: "14px" }}>
              
              {/* Header & Controls */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "10px", background: "var(--panel)", padding: "14px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                <div>
                  <h3 style={{ fontSize: "15px", fontWeight: 700, margin: "0 0 4px", display: "flex", alignItems: "center", gap: "6px" }}>
                    <Sliders size={18} color="var(--accent)" /> Course Moderation Manager
                  </h3>
                  <p style={{ margin: 0, fontSize: "12px", color: "var(--muted)" }}>
                    Set course-specific moderation marks limit. Supports live interactive editing or uploading an Excel sheet.
                  </p>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap" }}>
                  {/* Bulk Set Input */}
                  <div style={{ display: "flex", alignItems: "center", gap: "6px", background: "var(--bg)", padding: "4px 8px", borderRadius: "6px", border: "1px solid var(--line)" }}>
                    <span style={{ fontSize: "11.5px", color: "var(--muted)" }}>Bulk Marks:</span>
                    <input 
                      type="number" 
                      min="0" 
                      max="20" 
                      value={bulkModValue} 
                      onChange={(e) => setBulkModValue(e.target.value)}
                      style={{ width: "48px", padding: "4px", fontSize: "12px", textAlign: "center", borderRadius: "4px", border: "1px solid var(--line)" }}
                    />
                    <button 
                      type="button"
                      onClick={() => applyBulkModeration(bulkModValue)}
                      style={{ fontSize: "11.5px", padding: "4px 8px", background: "var(--accent)", color: "white", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: 600 }}
                    >
                      Set to All Courses
                    </button>
                  </div>

                  {/* Excel Upload for Moderation */}
                  <input 
                    type="file" 
                    ref={modFileInputRef}
                    accept=".xlsx,.xls,.csv" 
                    onChange={handleModerationExcelUpload}
                    style={{ display: "none" }}
                  />
                  <button 
                    type="button"
                    onClick={() => modFileInputRef.current?.click()}
                    className="secondary"
                    style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", padding: "6px 12px" }}
                  >
                    <FileUp size={14} /> Upload Moderation Excel
                  </button>

                  {/* Download Template */}
                  <button 
                    type="button"
                    onClick={handleDownloadModerationTemplate}
                    className="secondary"
                    style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", padding: "6px 12px" }}
                  >
                    <FileDown size={14} /> Download Template (.xlsx)
                  </button>

                  {/* Download Simulation Report */}
                  <button 
                    type="button"
                    onClick={handleExportSimulationExcel}
                    style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", padding: "6px 12px", background: "#10b981", color: "white", border: "none", borderRadius: "6px", cursor: "pointer", fontWeight: 600 }}
                  >
                    <Download size={14} /> Simulation Report (+0..+10)
                  </button>

                  {/* Reset Button */}
                  <button 
                    type="button"
                    onClick={resetAllModeration}
                    style={{ fontSize: "12px", padding: "6px 12px", background: "transparent", color: "#ef4444", border: "1px solid #ef4444", borderRadius: "6px", cursor: "pointer", fontWeight: 600 }}
                  >
                    Reset All to 0
                  </button>
                </div>
              </div>

              {/* Policy Toggle Bar: ESE-PR Only Moderation Switch */}
              <div style={{ 
                display: "flex", 
                alignItems: "center", 
                justifyContent: "space-between", 
                background: allowPrOnlyModeration ? "rgba(16, 185, 129, 0.08)" : "var(--panel)", 
                padding: "12px 16px", 
                borderRadius: "8px", 
                border: allowPrOnlyModeration ? "1.5px solid rgba(16, 185, 129, 0.3)" : "1px solid var(--line)",
                transition: "all 0.2s ease"
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                  <div style={{ 
                    background: allowPrOnlyModeration ? "rgba(16, 185, 129, 0.2)" : "rgba(239, 68, 68, 0.12)", 
                    color: allowPrOnlyModeration ? "#10b981" : "#ef4444", 
                    padding: "6px", 
                    borderRadius: "6px" 
                  }}>
                    <Sliders size={18} />
                  </div>
                  <div>
                    <div style={{ fontSize: "13px", fontWeight: 700, color: "var(--ink)", display: "flex", alignItems: "center", gap: "6px" }}>
                      Allow Moderation on Solely ESE - PR (Practical-Only) Courses
                      {allowPrOnlyModeration && (
                        <span style={{ fontSize: "10.5px", background: "#10b981", color: "white", padding: "1px 6px", borderRadius: "10px", fontWeight: 700 }}>
                          ACTIVE
                        </span>
                      )}
                    </div>
                    <div style={{ fontSize: "11.5px", color: "var(--muted)", marginTop: "2px" }}>
                      {allowPrOnlyModeration 
                        ? "Enabled: Moderation marks will be applied to both Theory (ESE-TH) and Practical-Only (ESE-PR) courses." 
                        : "Default: Moderation is ONLY effected on courses with an ESE Theory (TH) component. Practical-Only courses will not receive moderation marks unless this option is enabled."}
                    </div>
                  </div>
                </div>

                <label style={{ display: "flex", alignItems: "center", gap: "8px", cursor: "pointer", fontSize: "12px", fontWeight: 600, userSelect: "none" }}>
                  <input 
                    type="checkbox" 
                    checked={allowPrOnlyModeration} 
                    onChange={(e) => {
                      setAllowPrOnlyModeration(e.target.checked);
                      setStatus(e.target.checked 
                        ? "Moderation enabled for Practical-Only (ESE - PR) courses." 
                        : "Moderation restricted to courses with ESE Theory (TH) component.", "info");
                    }}
                    style={{ width: "18px", height: "18px", accentColor: "var(--accent)", cursor: "pointer" }}
                  />
                  <span style={{ color: allowPrOnlyModeration ? "#10b981" : "var(--muted)" }}>
                    {allowPrOnlyModeration ? "Enabled" : "Disabled (Default)"}
                  </span>
                </label>
              </div>

              {/* Filters & Search Toolbar in Moderation Tab */}
              <div style={{ display: "flex", flexDirection: "column", gap: "10px", background: "var(--panel)", padding: "12px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "10px" }}>
                  
                  {/* Search Input */}
                  <div style={{ position: "relative", flex: 1, minWidth: "260px", maxWidth: "380px" }}>
                    <Search size={14} style={{ position: "absolute", left: "10px", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                    <input 
                      type="text"
                      placeholder="Search courses by code, title, program..."
                      value={moderationSearch}
                      onChange={(e) => setModerationSearch(e.target.value)}
                      style={{ width: "100%", padding: "6px 28px 6px 30px", fontSize: "12px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)" }}
                    />
                    {moderationSearch && (
                      <X 
                        size={13} 
                        onClick={() => setModerationSearch("")}
                        style={{ position: "absolute", right: "10px", top: "50%", transform: "translateY(-50%)", cursor: "pointer", color: "var(--muted)" }}
                      />
                    )}
                  </div>

                  {/* Component Type Filters */}
                  <div style={{ display: "flex", alignItems: "center", gap: "4px", background: "var(--bg)", padding: "3px", borderRadius: "6px", border: "1px solid var(--line)" }}>
                    <span style={{ fontSize: "11px", color: "var(--muted)", padding: "0 4px", fontWeight: 600 }}>Type:</span>
                    <button 
                      type="button"
                      onClick={() => setModFilterType("ALL")}
                      style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: modFilterType === "ALL" ? "var(--accent)" : "transparent", color: modFilterType === "ALL" ? "white" : "var(--muted)" }}
                    >
                      All ({modFilterCounts.total})
                    </button>
                    <button 
                      type="button"
                      onClick={() => setModFilterType("TH")}
                      style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: modFilterType === "TH" ? "var(--accent)" : "transparent", color: modFilterType === "TH" ? "white" : "var(--muted)" }}
                    >
                      Theory ({modFilterCounts.thCount})
                    </button>
                    <button 
                      type="button"
                      onClick={() => setModFilterType("PR_ONLY")}
                      style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: modFilterType === "PR_ONLY" ? "var(--accent)" : "transparent", color: modFilterType === "PR_ONLY" ? "white" : "var(--muted)" }}
                    >
                      PR-Only ({modFilterCounts.prOnlyCount})
                    </button>
                  </div>

                  {/* Status Filters */}
                  <div style={{ display: "flex", alignItems: "center", gap: "4px", background: "var(--bg)", padding: "3px", borderRadius: "6px", border: "1px solid var(--line)", flexWrap: "wrap" }}>
                    <span style={{ fontSize: "11px", color: "var(--muted)", padding: "0 4px", fontWeight: 600 }}>Status:</span>
                    <button 
                      type="button"
                      onClick={() => setModFilterStatus("ALL")}
                      style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: modFilterStatus === "ALL" ? "var(--accent)" : "transparent", color: modFilterStatus === "ALL" ? "white" : "var(--muted)" }}
                    >
                      All ({modFilterCounts.total})
                    </button>
                    <button 
                      type="button"
                      onClick={() => setModFilterStatus("FAILED")}
                      style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: modFilterStatus === "FAILED" ? "#ef4444" : "transparent", color: modFilterStatus === "FAILED" ? "white" : "var(--muted)" }}
                    >
                      Has Failures ({modFilterCounts.failedCount})
                    </button>
                    <button 
                      type="button"
                      onClick={() => setModFilterStatus("NEAR_PASS")}
                      style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: modFilterStatus === "NEAR_PASS" ? "#f59e0b" : "transparent", color: modFilterStatus === "NEAR_PASS" ? "white" : "var(--muted)" }}
                    >
                      Deficit ≤ 5 ({modFilterCounts.nearPassCount})
                    </button>
                    <button 
                      type="button"
                      onClick={() => setModFilterStatus("ACTIVE_MOD")}
                      style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: modFilterStatus === "ACTIVE_MOD" ? "#3b82f6" : "transparent", color: modFilterStatus === "ACTIVE_MOD" ? "white" : "var(--muted)" }}
                    >
                      Has Mod ({modFilterCounts.activeModCount})
                    </button>
                    <button 
                      type="button"
                      onClick={() => setModFilterStatus("RESCUED")}
                      style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: modFilterStatus === "RESCUED" ? "#10b981" : "transparent", color: modFilterStatus === "RESCUED" ? "white" : "var(--muted)" }}
                    >
                      Rescued ({modFilterCounts.rescuedCount})
                    </button>
                  </div>

                  {/* Reset Filters */}
                  {(moderationSearch || modFilterType !== "ALL" || modFilterStatus !== "ALL") && (
                    <button
                      type="button"
                      onClick={() => {
                        setModerationSearch("");
                        setModFilterType("ALL");
                        setModFilterStatus("ALL");
                      }}
                      style={{ padding: "4px 8px", fontSize: "11px", background: "transparent", border: "1px solid var(--line)", borderRadius: "4px", color: "var(--muted)", cursor: "pointer" }}
                    >
                      Reset Filters
                    </button>
                  )}

                  <span style={{ fontSize: "11.5px", color: "var(--muted)", marginLeft: "auto" }}>
                    Showing <strong>{filteredAndSortedModCourses.length}</strong> of {distinctCourses.length} courses
                  </span>
                </div>
              </div>

              {/* Moderation Course List Table with Interactive Sort Headers */}
              <div style={{ flex: 1, overflow: "auto", background: "var(--panel)", border: "1px solid var(--line)", borderRadius: "8px" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px", textAlign: "left" }}>
                  <thead style={{ position: "sticky", top: 0, background: "var(--bg)", borderBottom: "2px solid var(--line)", zIndex: 10 }}>
                    <tr>
                      {/* Course Code */}
                      <th 
                        onClick={() => handleModSort("courseCode")}
                        style={{ padding: "10px 12px", color: modSortConfig.column === "courseCode" ? "var(--ink)" : "var(--muted)", fontWeight: 600, cursor: "pointer", userSelect: "none" }}
                      >
                        <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
                          Course Code
                          {modSortConfig.column === "courseCode" ? (
                            modSortConfig.direction === "asc" ? <ArrowUp size={12} color="var(--accent)" /> : <ArrowDown size={12} color="var(--accent)" />
                          ) : (
                            <ArrowUpDown size={11} style={{ opacity: 0.35 }} />
                          )}
                        </div>
                      </th>

                      {/* Course Name */}
                      <th 
                        onClick={() => handleModSort("courseName")}
                        style={{ padding: "10px 12px", color: modSortConfig.column === "courseName" ? "var(--ink)" : "var(--muted)", fontWeight: 600, cursor: "pointer", userSelect: "none" }}
                      >
                        <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
                          Course Name
                          {modSortConfig.column === "courseName" ? (
                            modSortConfig.direction === "asc" ? <ArrowUp size={12} color="var(--accent)" /> : <ArrowDown size={12} color="var(--accent)" />
                          ) : (
                            <ArrowUpDown size={11} style={{ opacity: 0.35 }} />
                          )}
                        </div>
                      </th>

                      {/* Component Type */}
                      <th 
                        onClick={() => handleModSort("componentType")}
                        style={{ padding: "10px 12px", color: modSortConfig.column === "componentType" ? "var(--ink)" : "var(--muted)", fontWeight: 600, textAlign: "center", cursor: "pointer", userSelect: "none" }}
                      >
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "4px" }}>
                          Component Type
                          {modSortConfig.column === "componentType" ? (
                            modSortConfig.direction === "asc" ? <ArrowUp size={12} color="var(--accent)" /> : <ArrowDown size={12} color="var(--accent)" />
                          ) : (
                            <ArrowUpDown size={11} style={{ opacity: 0.35 }} />
                          )}
                        </div>
                      </th>

                      {/* Total Enrolled */}
                      <th 
                        onClick={() => handleModSort("totalStudents")}
                        style={{ padding: "10px 12px", color: modSortConfig.column === "totalStudents" ? "var(--ink)" : "var(--muted)", fontWeight: 600, textAlign: "center", cursor: "pointer", userSelect: "none" }}
                      >
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "4px" }}>
                          Total Enrolled
                          {modSortConfig.column === "totalStudents" ? (
                            modSortConfig.direction === "asc" ? <ArrowUp size={12} color="var(--accent)" /> : <ArrowDown size={12} color="var(--accent)" />
                          ) : (
                            <ArrowUpDown size={11} style={{ opacity: 0.35 }} />
                          )}
                        </div>
                      </th>

                      {/* Raw Failed */}
                      <th 
                        onClick={() => handleModSort("rawFailed")}
                        style={{ padding: "10px 12px", color: modSortConfig.column === "rawFailed" ? "var(--ink)" : "var(--muted)", fontWeight: 600, textAlign: "center", cursor: "pointer", userSelect: "none" }}
                      >
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "4px" }}>
                          Raw Failed
                          {modSortConfig.column === "rawFailed" ? (
                            modSortConfig.direction === "asc" ? <ArrowUp size={12} color="var(--accent)" /> : <ArrowDown size={12} color="var(--accent)" />
                          ) : (
                            <ArrowUpDown size={11} style={{ opacity: 0.35 }} />
                          )}
                        </div>
                      </th>

                      {/* Deficit <= 5 Marks */}
                      <th 
                        onClick={() => handleModSort("nearPassCount")}
                        style={{ padding: "10px 12px", color: modSortConfig.column === "nearPassCount" ? "var(--ink)" : "var(--muted)", fontWeight: 600, textAlign: "center", cursor: "pointer", userSelect: "none" }}
                      >
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "4px" }}>
                          Deficit ≤ 5 Marks
                          {modSortConfig.column === "nearPassCount" ? (
                            modSortConfig.direction === "asc" ? <ArrowUp size={12} color="var(--accent)" /> : <ArrowDown size={12} color="var(--accent)" />
                          ) : (
                            <ArrowUpDown size={11} style={{ opacity: 0.35 }} />
                          )}
                        </div>
                      </th>

                      {/* Current Moderation Marks */}
                      <th 
                        onClick={() => handleModSort("currentMod")}
                        style={{ padding: "10px 12px", color: modSortConfig.column === "currentMod" ? "var(--ink)" : "var(--muted)", fontWeight: 600, textAlign: "center", width: "160px", cursor: "pointer", userSelect: "none" }}
                      >
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "4px" }}>
                          Current Moderation Marks
                          {modSortConfig.column === "currentMod" ? (
                            modSortConfig.direction === "asc" ? <ArrowUp size={12} color="var(--accent)" /> : <ArrowDown size={12} color="var(--accent)" />
                          ) : (
                            <ArrowUpDown size={11} style={{ opacity: 0.35 }} />
                          )}
                        </div>
                      </th>

                      {/* Rescued with Mod */}
                      <th 
                        onClick={() => handleModSort("rescuedInCourse")}
                        style={{ padding: "10px 12px", color: modSortConfig.column === "rescuedInCourse" ? "var(--ink)" : "var(--muted)", fontWeight: 600, textAlign: "center", cursor: "pointer", userSelect: "none" }}
                      >
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "4px" }}>
                          Rescued with Mod
                          {modSortConfig.column === "rescuedInCourse" ? (
                            modSortConfig.direction === "asc" ? <ArrowUp size={12} color="var(--accent)" /> : <ArrowDown size={12} color="var(--accent)" />
                          ) : (
                            <ArrowUpDown size={11} style={{ opacity: 0.35 }} />
                          )}
                        </div>
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredAndSortedModCourses.length === 0 ? (
                      <tr>
                        <td colSpan={8} style={{ padding: "36px", textAlign: "center", color: "var(--muted)" }}>
                          {distinctCourses.length === 0 
                            ? "No courses loaded. Upload an ADES Marks Excel sheet to begin." 
                            : "No courses match your filter criteria."}
                        </td>
                      </tr>
                    ) : (
                      filteredAndSortedModCourses.map(c => {
                        const currentMod = c.currentMod;
                        const rescuedInCourse = c.rescuedInCourse;
                        const isPrOnly = c.isPrOnly;

                        return (
                          <tr key={c.normCode} style={{ borderBottom: "1px solid var(--line)", background: currentMod > 0 ? "rgba(245, 158, 11, 0.04)" : "transparent" }}>
                            <td style={{ padding: "8px 12px", fontWeight: 600, color: "var(--ink)" }}>{c.courseCode}</td>
                            <td style={{ padding: "8px 12px", color: "var(--ink)" }}>
                              <div>{c.courseName}</div>
                              {c.program && <div style={{ fontSize: "10.5px", color: "var(--muted)" }}>{c.program}</div>}
                            </td>
                            <td style={{ padding: "8px 12px", textAlign: "center" }}>
                              {isPrOnly ? (
                                <span style={{ 
                                  fontSize: "10.5px", 
                                  padding: "2px 7px", 
                                  borderRadius: "4px", 
                                  fontWeight: 600, 
                                  background: allowPrOnlyModeration ? "rgba(16, 185, 129, 0.12)" : "rgba(239, 68, 68, 0.12)", 
                                  color: allowPrOnlyModeration ? "#10b981" : "#ef4444" 
                                }}>
                                  ESE-PR Only {allowPrOnlyModeration ? "(Eligible)" : "(Excluded)"}
                                </span>
                              ) : (
                                <span style={{ fontSize: "10.5px", padding: "2px 7px", borderRadius: "4px", fontWeight: 600, background: "var(--accent-soft)", color: "var(--accent)" }}>
                                  {c.hasEsePr ? "ESE-TH + PR" : "ESE-TH Only"}
                                </span>
                              )}
                            </td>
                            <td style={{ padding: "8px 12px", textAlign: "center", fontWeight: 600 }}>
                              {c.totalStudents}
                              {c.heldCount > 0 && (
                                <div style={{ fontSize: "9.5px", color: "#9333ea", fontWeight: 600 }} title={`${c.heldCount} student paper(s) held due to missing component(s)`}>
                                  {c.heldCount} held
                                </div>
                              )}
                            </td>
                            <td style={{ padding: "8px 12px", textAlign: "center", color: c.rawFailed > 0 ? "#ef4444" : "var(--muted)", fontWeight: c.rawFailed > 0 ? 600 : 400 }}>
                              {c.rawFailed}
                            </td>
                            <td style={{ padding: "8px 12px", textAlign: "center", color: c.nearPassCount > 0 ? "#f59e0b" : "var(--muted)", fontWeight: c.nearPassCount > 0 ? 600 : 400 }}>
                              {c.nearPassCount}
                            </td>
                            <td style={{ padding: "8px 12px", textAlign: "center" }}>
                              <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "6px" }}>
                                <input 
                                  type="number" 
                                  min="0" 
                                  max="50" 
                                  value={currentMod} 
                                  onChange={(e) => updateCourseModeration(c.normCode, e.target.value)}
                                  style={{ 
                                    width: "55px", 
                                    padding: "4px 6px", 
                                    fontSize: "12.5px", 
                                    textAlign: "center", 
                                    borderRadius: "4px", 
                                    border: currentMod > 0 ? "1.5px solid #f59e0b" : "1px solid var(--line)", 
                                    background: currentMod > 0 ? "rgba(245, 158, 11, 0.1)" : "var(--bg)",
                                    fontWeight: currentMod > 0 ? 700 : 400
                                  }}
                                />
                                <div style={{ display: "flex", gap: "2px" }}>
                                  {[2, 4, 6].map(m => (
                                    <button
                                      key={m}
                                      type="button"
                                      onClick={() => updateCourseModeration(c.normCode, m)}
                                      style={{ fontSize: "10px", padding: "2px 5px", borderRadius: "3px", border: "1px solid var(--line)", background: currentMod === m ? "var(--accent)" : "var(--bg)", color: currentMod === m ? "white" : "var(--muted)", cursor: "pointer" }}
                                    >
                                      +{m}
                                    </button>
                                  ))}
                                </div>
                              </div>
                            </td>
                            <td style={{ padding: "8px 12px", textAlign: "center" }}>
                              {rescuedInCourse > 0 ? (
                                <span style={{ display: "inline-flex", alignItems: "center", gap: "3px", background: "rgba(16, 185, 129, 0.15)", color: "#10b981", padding: "2px 8px", borderRadius: "10px", fontWeight: 600, fontSize: "11px" }}>
                                  +{rescuedInCourse} Passed
                                </span>
                              ) : (
                                <span style={{ color: "var(--muted)", fontSize: "11px" }}>
                                  {isPrOnly && !allowPrOnlyModeration && currentMod > 0 ? "(PR-Only Excluded)" : "0"}
                                </span>
                              )}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                  {/* Summary Totals Footer Row */}
                  {filteredAndSortedModCourses.length > 0 && (
                    <tfoot style={{ position: "sticky", bottom: 0, background: "var(--bg)", borderTop: "2px solid var(--line)", fontWeight: 700 }}>
                      <tr>
                        <td colSpan={3} style={{ padding: "10px 12px", color: "var(--ink)" }}>
                          TOTAL ({filteredAndSortedModCourses.length} Courses)
                        </td>
                        <td style={{ padding: "10px 12px", textAlign: "center", color: "var(--ink)" }}>
                          {modTableTotals.totalStudents}
                        </td>
                        <td style={{ padding: "10px 12px", textAlign: "center", color: "#ef4444" }}>
                          {modTableTotals.rawFailed}
                        </td>
                        <td style={{ padding: "10px 12px", textAlign: "center", color: "#f59e0b" }}>
                          {modTableTotals.nearPassCount}
                        </td>
                        <td style={{ padding: "10px 12px", textAlign: "center", color: "var(--muted)", fontSize: "11px" }}>
                          -
                        </td>
                        <td style={{ padding: "10px 12px", textAlign: "center", color: "#10b981", background: "rgba(16, 185, 129, 0.08)" }}>
                          +{modTableTotals.rescuedCount} Passed
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            </div>
          ) : activeTab === "simulation" ? (
            /* Course Pass Simulation Panel (+0 to +10) */
            <div style={{ flex: 1, display: "flex", flexDirection: "column", overflow: "hidden", padding: "16px", gap: "14px" }}>
              
              {/* Header & Controls */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "10px", background: "var(--panel)", padding: "14px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                <div>
                  <h3 style={{ fontSize: "15px", fontWeight: 700, margin: "0 0 4px", display: "flex", alignItems: "center", gap: "6px" }}>
                    <TrendingUp size={18} color="#10b981" /> Course-Wise Pass Simulation (0 to +10 Moderation)
                  </h3>
                  <p style={{ margin: 0, fontSize: "12px", color: "var(--muted)" }}>
                    Calculates and predicts how many students pass each course with normal marks vs incremental moderation from +1 up to +10 marks.
                  </p>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap" }}>
                  {/* Search Filter */}
                  <div style={{ position: "relative", width: "220px" }}>
                    <Search size={14} style={{ position: "absolute", left: "8px", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                    <input 
                      type="text" 
                      placeholder="Filter courses..." 
                      value={simSearchQuery} 
                      onChange={(e) => setSimSearchQuery(e.target.value)}
                      style={{ width: "100%", padding: "5px 8px 5px 28px", fontSize: "12px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)" }}
                    />
                    {simSearchQuery && (
                      <X 
                        size={13} 
                        onClick={() => setSimSearchQuery("")}
                        style={{ position: "absolute", right: "8px", top: "50%", transform: "translateY(-50%)", cursor: "pointer", color: "var(--muted)" }}
                      />
                    )}
                  </div>

                  {/* Export Simulation Excel Button */}
                  <button 
                    type="button"
                    onClick={handleExportSimulationExcel}
                    style={{ 
                      display: "flex", 
                      alignItems: "center", 
                      gap: "6px", 
                      fontSize: "12.5px", 
                      padding: "6px 14px", 
                      background: "#10b981", 
                      color: "white", 
                      border: "none", 
                      borderRadius: "6px", 
                      fontWeight: 600, 
                      cursor: "pointer" 
                    }}
                  >
                    <Download size={14} /> Download Simulation Excel (.xlsx)
                  </button>
                </div>
              </div>

              {/* Dual Pass Condition Verification Banner */}
              <div style={{ background: "rgba(59, 130, 246, 0.08)", border: "1px solid rgba(59, 130, 246, 0.25)", borderRadius: "8px", padding: "10px 14px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "10px", fontSize: "11.5px" }}>
                <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                  <Sparkles size={16} color="#3b82f6" />
                  <div>
                    <strong style={{ color: "var(--ink)" }}>Dual Pass Conditions Enforced for All Simulation Levels (+0 to +10):</strong>
                    <span style={{ color: "var(--muted)", marginLeft: "6px" }}>
                      1. <strong>30% ESE Rule</strong> (<code>ESE Overall + Mod &ge; ESE Min</code>) &nbsp;|&nbsp; 
                      2. <strong>35% Overall Rule</strong> (<code>Course Overall + Mod &ge; Overall Min</code>) &nbsp;&rarr;&nbsp; 
                      <span style={{ color: "#10b981", fontWeight: 700 }}>Pass awarded ONLY if BOTH rules pass.</span>
                    </span>
                  </div>
                </div>
                <div style={{ fontSize: "11px", color: "var(--muted)", background: "var(--bg)", padding: "3px 8px", borderRadius: "4px", border: "1px solid var(--line)" }}>
                  Eligibility: {allowPrOnlyModeration ? "Applied to TH & PR-Only" : "Restricted to ESE-TH courses (Default)"}
                </div>
              </div>

              {/* KPI Summary Cards */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: "10px" }}>
                <div style={{ background: "var(--panel)", padding: "10px 12px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                  <div style={{ fontSize: "10.5px", color: "var(--muted)", fontWeight: 600 }}>Total Courses</div>
                  <div style={{ fontSize: "17px", fontWeight: 700, color: "var(--ink)", marginTop: "2px" }}>{simTotals.totalCourses}</div>
                  <div style={{ fontSize: "10.5px", color: "var(--muted)" }}>{simTotals.totalStudents} student entries</div>
                </div>

                <div style={{ background: "rgba(239, 68, 68, 0.08)", padding: "10px 12px", borderRadius: "8px", border: "1px solid rgba(239, 68, 68, 0.25)" }}>
                  <div style={{ fontSize: "10.5px", color: "#ef4444", fontWeight: 600 }}>Absentees (Ab)</div>
                  <div style={{ fontSize: "17px", fontWeight: 700, color: "#ef4444", marginTop: "2px" }}>{simTotals.totalAbsent}</div>
                  <div style={{ fontSize: "10.5px", color: "#ef4444" }}>Excluded from pass</div>
                </div>

                <div style={{ background: "rgba(249, 115, 22, 0.08)", padding: "10px 12px", borderRadius: "8px", border: "1px solid rgba(249, 115, 22, 0.25)" }}>
                  <div style={{ fontSize: "10.5px", color: "#f97316", fontWeight: 600 }}>Malpractice (MP)</div>
                  <div style={{ fontSize: "17px", fontWeight: 700, color: "#f97316", marginTop: "2px" }}>{simTotals.totalMalpractice}</div>
                  <div style={{ fontSize: "10.5px", color: "#f97316" }}>Excluded from pass</div>
                </div>

                <div style={{ background: "var(--panel)", padding: "10px 12px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                  <div style={{ fontSize: "10.5px", color: "var(--muted)", fontWeight: 600 }}>30% ESE Pass (0 Mod)</div>
                  <div style={{ fontSize: "17px", fontWeight: 700, color: "#6366f1", marginTop: "2px" }}>{simTotals.rawEsePass}</div>
                  <div style={{ fontSize: "10.5px", color: "#6366f1", fontWeight: 600 }}>{simTotals.rawEsePassPct}% meet ESE &ge; 30%</div>
                </div>

                <div style={{ background: "var(--panel)", padding: "10px 12px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                  <div style={{ fontSize: "10.5px", color: "var(--muted)", fontWeight: 600 }}>35% Overall Pass (0 Mod)</div>
                  <div style={{ fontSize: "17px", fontWeight: 700, color: "#8b5cf6", marginTop: "2px" }}>{simTotals.rawOverallPass}</div>
                  <div style={{ fontSize: "10.5px", color: "#8b5cf6", fontWeight: 600 }}>{simTotals.rawOverallPassPct}% meet Agg &ge; 35%</div>
                </div>

                <div style={{ background: "rgba(59, 130, 246, 0.08)", padding: "10px 12px", borderRadius: "8px", border: "1px solid rgba(59, 130, 246, 0.25)" }}>
                  <div style={{ fontSize: "10.5px", color: "#3b82f6", fontWeight: 600 }}>Normal Pass (Both Met)</div>
                  <div style={{ fontSize: "17px", fontWeight: 700, color: "#3b82f6", marginTop: "2px" }}>{simTotals.rawPass}</div>
                  <div style={{ fontSize: "10.5px", color: "#3b82f6", fontWeight: 600 }}>{simTotals.rawPassPct}% raw combined pass</div>
                </div>

                <div style={{ background: "rgba(245, 158, 11, 0.1)", padding: "10px 12px", borderRadius: "8px", border: "1px solid rgba(245, 158, 11, 0.3)" }}>
                  <div style={{ fontSize: "10.5px", color: "#d97706", fontWeight: 600 }}>Pass at +5 Mod</div>
                  <div style={{ fontSize: "17px", fontWeight: 700, color: "#d97706", marginTop: "2px" }}>{simTotals.modPass[5]}</div>
                  <div style={{ fontSize: "10.5px", color: "#d97706", fontWeight: 600 }}>{simTotals.modPassPct(5)}% (+{simTotals.rescuedAtMod(5)})</div>
                </div>

                <div style={{ background: "rgba(16, 185, 129, 0.12)", padding: "10px 12px", borderRadius: "8px", border: "1px solid rgba(16, 185, 129, 0.35)" }}>
                  <div style={{ fontSize: "10.5px", color: "#10b981", fontWeight: 600 }}>Pass at +10 Mod</div>
                  <div style={{ fontSize: "17px", fontWeight: 700, color: "#10b981", marginTop: "2px" }}>{simTotals.modPass[10]}</div>
                  <div style={{ fontSize: "10.5px", color: "#10b981", fontWeight: 600 }}>{simTotals.modPassPct(10)}% (+{simTotals.rescuedAtMod(10)})</div>
                </div>
              </div>

              {/* Simulation Table Container */}
              <div style={{ flex: 1, overflow: "auto", border: "1px solid var(--line)", borderRadius: "8px", background: "var(--panel)" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                  <thead style={{ position: "sticky", top: 0, background: "var(--bg)", zIndex: 10, borderBottom: "2px solid var(--line)" }}>
                    <tr>
                      <th style={{ padding: "10px 12px", textAlign: "left", color: "var(--muted)", fontWeight: 600, width: "110px" }}>Course Code</th>
                      <th style={{ padding: "10px 12px", textAlign: "left", color: "var(--muted)", fontWeight: 600, minWidth: "160px" }}>Course Name</th>
                      <th style={{ padding: "10px 12px", textAlign: "center", color: "var(--muted)", fontWeight: 600, width: "90px" }}>Type</th>
                      <th style={{ padding: "10px 12px", textAlign: "center", color: "var(--muted)", fontWeight: 600, width: "65px" }}>Total</th>
                      <th style={{ padding: "10px 8px", textAlign: "center", color: "#ef4444", fontWeight: 600, width: "65px", background: "rgba(239, 68, 68, 0.05)" }} title="Absent students count">Absent</th>
                      <th style={{ padding: "10px 8px", textAlign: "center", color: "#f97316", fontWeight: 600, width: "65px", background: "rgba(249, 115, 22, 0.05)" }} title="Malpractice students count">MP</th>
                      <th style={{ padding: "10px 8px", textAlign: "center", color: "#6366f1", fontWeight: 600, width: "75px", background: "rgba(99, 102, 241, 0.05)" }}>30% ESE</th>
                      <th style={{ padding: "10px 8px", textAlign: "center", color: "#8b5cf6", fontWeight: 600, width: "75px", background: "rgba(139, 92, 246, 0.05)" }}>35% Agg</th>
                      <th style={{ padding: "10px 10px", textAlign: "center", color: "var(--ink)", fontWeight: 700, width: "95px", background: "rgba(59, 130, 246, 0.1)" }}>Normal (Both)</th>
                      {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(m => (
                        <th 
                          key={m} 
                          style={{ 
                            padding: "10px 6px", 
                            textAlign: "center", 
                            color: m === 5 ? "#d97706" : m === 10 ? "#10b981" : "var(--muted)", 
                            fontWeight: m === 5 || m === 10 ? 700 : 600, 
                            width: "50px",
                            background: m === 5 ? "rgba(245, 158, 11, 0.1)" : m === 10 ? "rgba(16, 185, 129, 0.1)" : "transparent"
                          }}
                        >
                          +{m}
                        </th>
                      ))}
                      <th style={{ padding: "10px 12px", textAlign: "center", color: "#10b981", fontWeight: 700, width: "95px", background: "rgba(16, 185, 129, 0.08)" }}>Max Gain (+10)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredSimulationCourses.length === 0 ? (
                      <tr>
                        <td colSpan={20} style={{ padding: "32px", textAlign: "center", color: "var(--muted)" }}>
                          No matching courses found.
                        </td>
                      </tr>
                    ) : (
                      filteredSimulationCourses.map((c, idx) => {
                        const rawPct = c.totalStudents > 0 ? ((c.rawPassCount / c.totalStudents) * 100).toFixed(0) : "0";
                        const maxRescued = c.passCountAtMod[10] - c.rawPassCount;
                        const plus10Pct = c.totalStudents > 0 ? ((c.passCountAtMod[10] / c.totalStudents) * 100).toFixed(0) : "0";

                        return (
                          <tr key={c.normCode} style={{ borderBottom: "1px solid var(--line)", background: idx % 2 === 0 ? "transparent" : "rgba(255, 255, 255, 0.02)" }}>
                            <td style={{ padding: "8px 12px", fontWeight: 600, color: "var(--ink)" }}>{c.courseCode}</td>
                            <td style={{ padding: "8px 12px", color: "var(--muted)" }}>
                              <div style={{ fontWeight: 500, color: "var(--ink)" }}>{c.courseName}</div>
                              {c.program && <div style={{ fontSize: "10.5px", color: "var(--muted)" }}>{c.program}</div>}
                            </td>
                            <td style={{ padding: "8px 12px", textAlign: "center" }}>
                              {(c.hasEseTh && c.hasEsePr) ? (
                                <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(59, 130, 246, 0.12)", color: "#3b82f6", fontWeight: 600 }}>
                                  TH + PR
                                </span>
                              ) : c.isPrOnly ? (
                                <span style={{ 
                                  fontSize: "10.5px", 
                                  padding: "2px 6px", 
                                  borderRadius: "4px", 
                                  background: allowPrOnlyModeration ? "rgba(16, 185, 129, 0.12)" : "rgba(239, 68, 68, 0.12)", 
                                  color: allowPrOnlyModeration ? "#10b981" : "#ef4444", 
                                  fontWeight: 600 
                                }}>
                                  PR-Only {allowPrOnlyModeration ? "✓" : "✗"}
                                </span>
                              ) : (
                                <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(99, 102, 241, 0.12)", color: "#6366f1", fontWeight: 600 }}>
                                  Theory (TH)
                                </span>
                              )}
                            </td>
                            <td style={{ padding: "8px 12px", textAlign: "center", fontWeight: 600 }}>
                              {c.totalStudents}
                              {c.heldCount > 0 && (
                                <div style={{ fontSize: "9.5px", color: "#9333ea", fontWeight: 600 }} title={`${c.heldCount} student paper(s) held due to missing component(s)`}>
                                  {c.heldCount} held
                                </div>
                              )}
                            </td>
                            <td style={{ padding: "8px 8px", textAlign: "center", color: (c.absentCount || 0) > 0 ? "#ef4444" : "var(--muted)", fontWeight: (c.absentCount || 0) > 0 ? 600 : 400, background: "rgba(239, 68, 68, 0.03)" }}>
                              {c.absentCount || 0}
                            </td>
                            <td style={{ padding: "8px 8px", textAlign: "center", color: (c.malpracticeCount || 0) > 0 ? "#f97316" : "var(--muted)", fontWeight: (c.malpracticeCount || 0) > 0 ? 600 : 400, background: "rgba(249, 115, 22, 0.03)" }}>
                              {c.malpracticeCount || 0}
                            </td>
                            <td style={{ padding: "8px 8px", textAlign: "center", color: "#6366f1", background: "rgba(99, 102, 241, 0.03)" }}>
                              {c.rawEsePassCount}
                            </td>
                            <td style={{ padding: "8px 8px", textAlign: "center", color: "#8b5cf6", background: "rgba(139, 92, 246, 0.03)" }}>
                              {c.rawOverallPassCount}
                            </td>
                            <td style={{ padding: "8px 10px", textAlign: "center", background: "rgba(59, 130, 246, 0.06)" }}>
                              <span style={{ fontWeight: 700, color: "var(--ink)" }}>{c.rawPassCount}</span>
                              <span style={{ fontSize: "10.5px", color: "var(--muted)", marginLeft: "3px" }}>({rawPct}%)</span>
                            </td>
                            {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(m => {
                              const cnt = c.passCountAtMod[m];
                              const diff = cnt - c.rawPassCount;
                              const isHighlight = m === 5 || m === 10;
                              return (
                                <td 
                                  key={m} 
                                  style={{ 
                                    padding: "8px 5px", 
                                    textAlign: "center", 
                                    fontWeight: diff > 0 ? 600 : 400,
                                    color: diff > 0 ? (m === 10 ? "#10b981" : m >= 5 ? "#d97706" : "var(--ink)") : "var(--muted)",
                                    background: isHighlight ? (m === 10 ? "rgba(16, 185, 129, 0.05)" : "rgba(245, 158, 11, 0.05)") : "transparent"
                                  }}
                                >
                                  {cnt}
                                  {diff > 0 && (
                                    <div style={{ fontSize: "9.5px", color: m === 10 ? "#10b981" : "#f59e0b", fontWeight: 700 }}>
                                      +{diff}
                                    </div>
                                  )}
                                </td>
                              );
                            })}
                            <td style={{ padding: "8px 12px", textAlign: "center", background: "rgba(16, 185, 129, 0.05)" }}>
                              {maxRescued > 0 ? (
                                <span style={{ display: "inline-flex", alignItems: "center", gap: "2px", background: "rgba(16, 185, 129, 0.15)", color: "#10b981", padding: "2px 6px", borderRadius: "10px", fontWeight: 700, fontSize: "11px" }}>
                                  +{maxRescued} ({plus10Pct}%)
                                </span>
                              ) : (
                                <span style={{ color: "var(--muted)", fontSize: "11px" }}>0</span>
                              )}
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                  {/* Summary Totals Row */}
                  {filteredSimulationCourses.length > 0 && (
                    <tfoot style={{ position: "sticky", bottom: 0, background: "var(--bg)", borderTop: "2px solid var(--line)", fontWeight: 700 }}>
                      <tr>
                        <td colSpan={3} style={{ padding: "10px 12px", color: "var(--ink)" }}>TOTAL (All Filtered Courses)</td>
                        <td style={{ padding: "10px 12px", textAlign: "center", color: "var(--ink)" }}>{simTotals.totalStudents}</td>
                        <td style={{ padding: "10px 8px", textAlign: "center", color: "#ef4444" }}>{simTotals.totalAbsent}</td>
                        <td style={{ padding: "10px 8px", textAlign: "center", color: "#f97316" }}>{simTotals.totalMalpractice}</td>
                        <td style={{ padding: "10px 8px", textAlign: "center", color: "#6366f1" }}>{simTotals.rawEsePass}</td>
                        <td style={{ padding: "10px 8px", textAlign: "center", color: "#8b5cf6" }}>{simTotals.rawOverallPass}</td>
                        <td style={{ padding: "10px 10px", textAlign: "center", color: "#3b82f6", background: "rgba(59, 130, 246, 0.08)" }}>
                          {simTotals.rawPass} ({simTotals.rawPassPct}%)
                        </td>
                        {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(m => (
                          <td 
                            key={m} 
                            style={{ 
                              padding: "10px 5px", 
                              textAlign: "center", 
                              color: m === 10 ? "#10b981" : m === 5 ? "#d97706" : "var(--ink)",
                              background: m === 10 ? "rgba(16, 185, 129, 0.1)" : m === 5 ? "rgba(245, 158, 11, 0.1)" : "transparent"
                            }}
                          >
                            {simTotals.modPass[m]}
                          </td>
                        ))}
                        <td style={{ padding: "10px 12px", textAlign: "center", color: "#10b981", background: "rgba(16, 185, 129, 0.1)" }}>
                          +{simTotals.rescuedAtMod(10)} ({simTotals.modPassPct(10)}%)
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            </div>
          ) : (
            /* Results Table & Toolbar */
            <>
              {/* Interactive Toolbar */}
              <div style={{ display: "flex", flexDirection: "column", gap: "8px", padding: "10px 16px", borderBottom: "1px solid var(--line)", background: "var(--panel)" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: "8px" }}>
                  
                  <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
                    
                    {/* Search Box */}
                    <div style={{ position: "relative", width: "220px" }}>
                      <Search size={14} style={{ position: "absolute", left: "8px", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                      <input 
                        type="text" 
                        placeholder="Search student, seat, PRN, course..." 
                        value={searchQuery}
                        onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }}
                        style={{ width: "100%", padding: "5px 8px 5px 28px", fontSize: "12px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)" }}
                      />
                      {searchQuery && (
                        <X 
                          size={13} 
                          onClick={() => setSearchQuery("")}
                          style={{ position: "absolute", right: "8px", top: "50%", transform: "translateY(-50%)", cursor: "pointer", color: "var(--muted)" }}
                        />
                      )}
                    </div>

                    {/* Result Status Filter Pills */}
                    <div style={{ display: "flex", background: "var(--bg)", padding: "2px", borderRadius: "6px", border: "1px solid var(--line)", gap: "2px" }}>
                      <button 
                        type="button"
                        onClick={() => { setSelectedResultFilter("ALL"); setPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "ALL" ? "var(--accent)" : "transparent", color: selectedResultFilter === "ALL" ? "white" : "var(--muted)" }}
                      >
                        All ({processedRows.length})
                      </button>
                      <button 
                        type="button"
                        onClick={() => { setSelectedResultFilter("PASS"); setPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "PASS" ? "#10b981" : "transparent", color: selectedResultFilter === "PASS" ? "white" : "var(--muted)" }}
                      >
                        Passed ({metrics.totalPassed})
                      </button>
                      <button 
                        type="button"
                        onClick={() => { setSelectedResultFilter("PASS_MOD"); setPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "PASS_MOD" ? "#f59e0b" : "transparent", color: selectedResultFilter === "PASS_MOD" ? "white" : "var(--muted)" }}
                      >
                        Via Mod (+{metrics.moderatedPassed})
                      </button>
                      <button 
                        type="button"
                        onClick={() => { setSelectedResultFilter("FAIL"); setPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "FAIL" ? "#ef4444" : "transparent", color: selectedResultFilter === "FAIL" ? "white" : "var(--muted)" }}
                      >
                        Failed ({metrics.failed})
                      </button>
                      <button 
                        type="button"
                        onClick={() => { setSelectedResultFilter("ESE_FAIL"); setPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "ESE_FAIL" ? "#ef4444" : "transparent", color: selectedResultFilter === "ESE_FAIL" ? "white" : "var(--muted)" }}
                      >
                        ESE Fail ({metrics.eseFailed})
                      </button>
                      <button 
                        type="button"
                        onClick={() => { setSelectedResultFilter("OVERALL_FAIL"); setPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "OVERALL_FAIL" ? "#ef4444" : "transparent", color: selectedResultFilter === "OVERALL_FAIL" ? "white" : "var(--muted)" }}
                      >
                        Overall Fail ({metrics.overallFailed})
                      </button>
                      {metrics.absentCount > 0 && (
                        <button 
                          type="button"
                          onClick={() => { setSelectedResultFilter("ABSENT"); setPage(0); }}
                          style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "ABSENT" ? "#dc2626" : "transparent", color: selectedResultFilter === "ABSENT" ? "white" : "#dc2626" }}
                        >
                          Absent ({metrics.absentCount})
                        </button>
                      )}
                      {metrics.malpracticeCount > 0 && (
                        <button 
                          type="button"
                          onClick={() => { setSelectedResultFilter("MALPRACTICE"); setPage(0); }}
                          style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "MALPRACTICE" ? "#d97706" : "transparent", color: selectedResultFilter === "MALPRACTICE" ? "white" : "#d97706" }}
                        >
                          Malpractice ({metrics.malpracticeCount})
                        </button>
                      )}
                      {metrics.heldbackCount > 0 && (
                        <button 
                          type="button"
                          onClick={() => { setSelectedResultFilter("HELDBACK"); setPage(0); }}
                          style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "HELDBACK" ? "#c026d3" : "transparent", color: selectedResultFilter === "HELDBACK" ? "white" : "#c026d3" }}
                        >
                          Heldback ({metrics.heldbackCount})
                        </button>
                      )}
                      {metrics.missingCompCount > 0 && (
                        <button 
                          type="button"
                          onClick={() => { setSelectedResultFilter("HELD_MISSING"); setPage(0); }}
                          style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "HELD_MISSING" ? "#9333ea" : "transparent", color: selectedResultFilter === "HELD_MISSING" ? "white" : "#9333ea" }}
                        >
                          Missing Comp ({metrics.missingCompCount})
                        </button>
                      )}
                      {metrics.heldbackCount === 0 && metrics.missingCompCount === 0 && metrics.heldCount > 0 && (
                        <button 
                          type="button"
                          onClick={() => { setSelectedResultFilter("HELD"); setPage(0); }}
                          style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: selectedResultFilter === "HELD" ? "#9333ea" : "transparent", color: selectedResultFilter === "HELD" ? "white" : "#9333ea" }}
                        >
                          Held ({metrics.heldCount})
                        </button>
                      )}
                    </div>

                    {/* Faculty Filter */}
                    {uniqueFaculties.length > 0 && (
                      <select 
                        value={selectedFacultyFilter} 
                        onChange={(e) => { setSelectedFacultyFilter(e.target.value); setPage(0); }}
                        style={{ padding: "4px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)", maxWidth: "180px" }}
                        title="Filter by Faculty"
                      >
                        <option value="ALL">All Faculties ({uniqueFaculties.length})</option>
                        {uniqueFaculties.map(f => <option key={f} value={f}>{f}</option>)}
                      </select>
                    )}

                    {/* College Filter */}
                    {uniqueColleges.length > 0 && (
                      <select 
                        value={selectedCollegeFilter} 
                        onChange={(e) => { setSelectedCollegeFilter(e.target.value); setPage(0); }}
                        style={{ padding: "4px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)", maxWidth: "200px" }}
                        title="Filter by College"
                      >
                        <option value="ALL">All Colleges ({uniqueColleges.length})</option>
                        {uniqueColleges.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    )}

                    {/* Program Filter */}
                    {uniquePrograms.length > 0 && (
                      <select 
                        value={selectedProgramFilter} 
                        onChange={(e) => { 
                          setSelectedProgramFilter(e.target.value); 
                          setSelectedCourseFilter("ALL");
                          setPage(0); 
                        }}
                        style={{ padding: "4px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)", maxWidth: "220px" }}
                        title="Filter by Programme"
                      >
                        <option value="ALL">All Programs ({uniquePrograms.length})</option>
                        {uniquePrograms.map(p => <option key={p} value={p}>{p}</option>)}
                      </select>
                    )}

                    {/* Course Filter */}
                    {uniqueCourses.length > 0 && (
                      <select 
                        value={selectedCourseFilter} 
                        onChange={(e) => { setSelectedCourseFilter(e.target.value); setPage(0); }}
                        style={{ padding: "4px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)", maxWidth: "260px" }}
                        title="Filter by Course"
                      >
                        <option value="ALL">All Courses ({uniqueCourses.length})</option>
                        {uniqueCourses.map(c => <option key={c.code} value={c.code}>{c.label}</option>)}
                      </select>
                    )}

                    {(Object.keys(columnFilters).length > 0 || sortConfig.column || searchQuery || selectedResultFilter !== "ALL" || selectedFacultyFilter !== "ALL" || selectedCollegeFilter !== "ALL" || selectedProgramFilter !== "ALL" || selectedCourseFilter !== "ALL") && (
                      <button 
                        type="button" 
                        onClick={clearAllFilters}
                        style={{ display: "flex", alignItems: "center", gap: "4px", padding: "4px 8px", fontSize: "11px", color: "var(--danger)", background: "var(--danger-soft)", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: 600 }}
                      >
                        <RefreshCw size={11} /> Reset Filters
                      </button>
                    )}
                  </div>

                  <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                    <button 
                      type="button" 
                      onClick={() => handleExportExcel(filteredRows)}
                      style={{ display: "flex", alignItems: "center", gap: "4px", padding: "4px 9px", fontSize: "11px", borderRadius: "5px", border: "1px solid var(--accent)", background: "var(--accent-soft)", color: "var(--accent)", cursor: "pointer", fontWeight: 600 }}
                      title={filteredRows.length !== processedRows.length ? `Export current filtered view (${filteredRows.length} rows) to 31-Col ADES XLSX` : "Export all rows to 31-Col ADES XLSX"}
                    >
                      <Download size={12} /> Export Current View ({filteredRows.length})
                    </button>

                    <span style={{ fontSize: "12px", color: "var(--muted)" }}>
                      Showing {filteredRows.length > 0 ? page * pageSize + 1 : 0} - {Math.min((page + 1) * pageSize, filteredRows.length)} of {filteredRows.length} rows
                    </span>

                    {/* Pagination Controls */}
                    <div style={{ display: "flex", gap: "4px" }}>
                      <button 
                        type="button" 
                        disabled={page === 0}
                        onClick={() => setPage(p => Math.max(0, p - 1))}
                        style={{ padding: "3px 8px", fontSize: "11px", borderRadius: "4px", border: "1px solid var(--line)", background: "var(--bg)", cursor: page === 0 ? "not-allowed" : "pointer", opacity: page === 0 ? 0.5 : 1 }}
                      >
                        Prev
                      </button>
                      <button 
                        type="button" 
                        disabled={(page + 1) * pageSize >= filteredRows.length}
                        onClick={() => setPage(p => p + 1)}
                        style={{ padding: "3px 8px", fontSize: "11px", borderRadius: "4px", border: "1px solid var(--line)", background: "var(--bg)", cursor: (page + 1) * pageSize >= filteredRows.length ? "not-allowed" : "pointer", opacity: (page + 1) * pageSize >= filteredRows.length ? 0.5 : 1 }}
                      >
                        Next
                      </button>
                    </div>
                  </div>

                </div>
              </div>

              {/* Data Table Grid */}
              <div style={{ flex: 1, overflow: "auto", position: "relative" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "11.5px", whiteSpace: "nowrap" }}>
                  <thead style={{ position: "sticky", top: 0, background: "var(--panel)", zIndex: 10, borderBottom: "1px solid var(--line)", boxShadow: "0 1px 3px rgba(0,0,0,0.05)" }}>
                    <tr>
                      <th style={{ padding: "6px 8px", borderRight: "1px solid var(--line)", textAlign: "center", width: "40px", color: "var(--muted)" }}>#</th>
                      {ADES_OUTPUT_HEADERS.map((header) => {
                        const isSorted = sortConfig.column === header;
                        const isPassCol = header.includes("Pass") || header.includes("Pass/Fail");
                        const isModCol = header === "Moderation Marks";
                        
                        return (
                          <th 
                            key={header} 
                            style={{ 
                              padding: "6px 8px", 
                              borderRight: "1px solid var(--line)", 
                              textAlign: "left", 
                              color: isPassCol ? "var(--ink)" : isModCol ? "#f59e0b" : "var(--muted)", 
                              fontWeight: (isPassCol || isModCol) ? 700 : 600, 
                              background: isModCol ? "rgba(245, 158, 11, 0.08)" : isPassCol ? "var(--accent-soft)" : "transparent",
                              minWidth: header.length > 25 ? "180px" : header.length > 15 ? "130px" : "90px",
                              cursor: "pointer",
                              userSelect: "none"
                            }}
                            onClick={() => handleSort(header)}
                          >
                            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: "4px" }}>
                              <span>{header}</span>
                              <div style={{ display: "flex", alignItems: "center" }}>
                                {isSorted ? (
                                  sortConfig.direction === "asc" ? <ArrowUp size={12} color="var(--accent)" /> : <ArrowDown size={12} color="var(--accent)" />
                                ) : (
                                  <ArrowUpDown size={11} style={{ opacity: 0.3 }} />
                                )}
                              </div>
                            </div>
                            
                            {/* Column Search Filter Input */}
                            <div style={{ marginTop: "4px" }} onClick={(e) => e.stopPropagation()}>
                              <input 
                                type="text"
                                placeholder={"Filter..."}
                                value={columnFilters[header] || ""}
                                onChange={(e) => updateColumnFilter(header, e.target.value)}
                                style={{ width: "100%", fontSize: "10px", padding: "2px 4px", borderRadius: "3px", border: "1px solid var(--line)", background: "var(--bg)" }}
                              />
                            </div>
                          </th>
                        );
                      })}
                    </tr>
                  </thead>
                  <tbody>
                    {pagedRows.length === 0 ? (
                      <tr>
                        <td colSpan={ADES_OUTPUT_HEADERS.length + 1} style={{ padding: "40px", textAlign: "center", color: "var(--muted)" }}>
                          {processedRows.length === 0 
                            ? "Upload an ADES Marks Excel sheet (.xlsx, .xls) to view calculated student course results and apply moderation." 
                            : "No student records match the active search or column filters."}
                        </td>
                      </tr>
                    ) : (
                      pagedRows.map((row, idx) => {
                        const globalIdx = page * pageSize + idx + 1;
                        const isHeldback = !!row._isHeldback;
                        const isHeld = !!row._isHeld;
                        const isAbsent = !isHeldback && !!row._isAbsent;
                        const isMalpractice = !isHeldback && !isAbsent && !!row._isMalpractice;
                        const isPass = !isHeldback && !isHeld && !isAbsent && !isMalpractice && row[coursePassKey] === "Pass";
                        const isModPass = !isHeldback && !isHeld && !isAbsent && !isMalpractice && row._isModeratedPass;

                        return (
                          <tr 
                            key={globalIdx} 
                            style={{ 
                              borderBottom: "1px solid var(--line)", 
                              background: isHeldback
                                ? "rgba(192, 38, 211, 0.08)"
                                : isHeld
                                  ? "rgba(147, 51, 234, 0.08)"
                                  : isMalpractice
                                    ? "rgba(245, 158, 11, 0.08)"
                                    : isAbsent
                                      ? "rgba(239, 68, 68, 0.08)"
                                      : isModPass 
                                        ? "rgba(245, 158, 11, 0.05)" 
                                        : isPass 
                                          ? "transparent" 
                                          : "rgba(239, 68, 68, 0.04)" 
                            }}
                          >
                            <td style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center", color: "var(--muted)" }}>
                              {globalIdx}
                            </td>

                            {ADES_OUTPUT_HEADERS.map((col) => {
                              const val = row[col];
                              const isCoursePass = col === coursePassKey;
                              const isOtherPass = col === "ESE Pass" || col === "Overall pass";
                              const isModMarks = col === "Moderation Marks";
                              const isAbsentMark = val === "Absent (Ab)";
                              const isMalpracticeMark = val === "Malpractice (MP)";
                              const isMissingMark = val === "Missing";

                              if (isCoursePass) {
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center" }}>
                                    {isHeldback ? (
                                      <span 
                                        style={{ display: "inline-flex", alignItems: "center", gap: "3px", background: "rgba(192, 38, 211, 0.2)", color: "#c026d3", padding: "2px 6px", borderRadius: "4px", fontWeight: 700 }} 
                                        title={`Heldback: ${row._heldbackReason || "APC Within Condonable Limit"}`}
                                      >
                                        <Lock size={11} /> Held (Heldback)
                                      </span>
                                    ) : isHeld ? (
                                      <span 
                                        style={{ display: "inline-flex", alignItems: "center", gap: "3px", background: "rgba(147, 51, 234, 0.18)", color: "#7e22ce", padding: "2px 6px", borderRadius: "4px", fontWeight: 700 }} 
                                        title={`Held: Missing required component(s) [${(row._missingComponents || []).join(", ")}]`}
                                      >
                                        <AlertTriangle size={11} /> Held (Missing Component)
                                      </span>
                                    ) : isMalpractice ? (
                                      <span 
                                        style={{ display: "inline-flex", alignItems: "center", gap: "3px", background: "rgba(245, 158, 11, 0.2)", color: "#b45309", padding: "2px 6px", borderRadius: "4px", fontWeight: 700 }} 
                                        title={`Malpractice (${row._malpracticeStatus || "EHB"})${row._malpracticeRemarks ? ` - ${row._malpracticeRemarks}` : ""}${row._malpracticeDate ? ` [${row._malpracticeDate}]` : ""}`}
                                      >
                                        <ShieldAlert size={11} /> Fail (MP)
                                      </span>
                                    ) : isAbsent ? (
                                      <span style={{ display: "inline-flex", alignItems: "center", gap: "3px", background: "rgba(239, 68, 68, 0.2)", color: "#dc2626", padding: "2px 6px", borderRadius: "4px", fontWeight: 700 }} title={row._absentStatus || "Marked Absent"}>
                                        <UserX size={11} /> Fail (Absent)
                                      </span>
                                    ) : isModPass ? (
                                      <span style={{ display: "inline-flex", alignItems: "center", gap: "3px", background: "rgba(16, 185, 129, 0.15)", color: "#10b981", padding: "2px 6px", borderRadius: "4px", fontWeight: 700 }}>
                                        <CheckCircle2 size={11} /> Pass (Mod +{row["Moderation Marks"]})
                                      </span>
                                    ) : isPass ? (
                                      <span style={{ display: "inline-flex", alignItems: "center", gap: "3px", background: "rgba(16, 185, 129, 0.12)", color: "#10b981", padding: "2px 6px", borderRadius: "4px", fontWeight: 600 }}>
                                        <CheckCircle2 size={11} /> Pass
                                      </span>
                                    ) : (
                                      <span style={{ display: "inline-flex", alignItems: "center", gap: "3px", background: "rgba(239, 68, 68, 0.12)", color: "#ef4444", padding: "2px 6px", borderRadius: "4px", fontWeight: 600 }}>
                                        <XCircle size={11} /> Fail
                                      </span>
                                    )}
                                  </td>
                                );
                              }

                              if (isMissingMark) {
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center" }}>
                                    <span style={{ 
                                      display: "inline-block", 
                                      color: "#7e22ce", 
                                      background: "rgba(147, 51, 234, 0.15)", 
                                      padding: "1px 6px", 
                                      borderRadius: "4px", 
                                      fontWeight: 700, 
                                      fontSize: "11px" 
                                    }} title="Required component missing from source data">
                                      Missing
                                    </span>
                                  </td>
                                );
                              }

                              if (isMalpracticeMark) {
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center" }}>
                                    <span style={{ 
                                      display: "inline-block", 
                                      color: "#b45309", 
                                      background: "rgba(245, 158, 11, 0.2)", 
                                      padding: "1px 6px", 
                                      borderRadius: "4px", 
                                      fontWeight: 700, 
                                      fontSize: "11px" 
                                    }} title={`Status: ${row._malpracticeStatus || "EHB"}${row._malpracticeRemarks ? ` | Remarks: ${row._malpracticeRemarks}` : ""}${row._malpracticeDate ? ` | Date: ${row._malpracticeDate}` : ""}`}>
                                      Malpractice (MP)
                                    </span>
                                  </td>
                                );
                              }

                              if (isAbsentMark) {
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center" }}>
                                    <span style={{ 
                                      display: "inline-block", 
                                      color: "#dc2626", 
                                      background: "rgba(239, 68, 68, 0.16)", 
                                      padding: "1px 6px", 
                                      borderRadius: "4px", 
                                      fontWeight: 700, 
                                      fontSize: "11px" 
                                    }} title={row._absentStatus || "Marked Absent During Mark Entry"}>
                                      Absent (Ab)
                                    </span>
                                  </td>
                                );
                              }

                              if (isOtherPass) {
                                const passVal = String(val);
                                const isHeldbackOther = isHeldback || passVal === "Held (Heldback)";
                                const isHeldOther = passVal === "Held";
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center" }}>
                                    <span style={{ 
                                      color: isHeldbackOther ? "#c026d3" : isHeldOther ? "#7e22ce" : passVal === "Pass" ? "#10b981" : "#ef4444",
                                      fontWeight: 600,
                                      background: isHeldbackOther ? "rgba(192, 38, 211, 0.12)" : isHeldOther ? "rgba(147, 51, 234, 0.12)" : passVal === "Pass" ? "rgba(16, 185, 129, 0.08)" : "rgba(239, 68, 68, 0.08)",
                                      padding: "1px 5px",
                                      borderRadius: "3px"
                                    }}>
                                      {passVal}
                                    </span>
                                  </td>
                                );
                              }

                              if (isModMarks) {
                                const num = Number(val);
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center", fontWeight: num > 0 ? 700 : 400, color: num > 0 ? "#f59e0b" : "var(--muted)" }}>
                                    {num > 0 ? "+" + num : 0}
                                  </td>
                                );
                              }

                              return (
                                <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", color: "var(--ink)" }}>
                                  {val !== undefined && val !== null ? String(val) : ""}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </>
          )}

        </main>
      </div>

      {/* Extraction & Calculation Logic Modal */}
      {showHelpModal && (
        <div 
          style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.6)", backdropFilter: "blur(3px)", zIndex: 2000, display: "flex", alignItems: "center", justifyContent: "center", padding: "20px" }}
          onClick={() => setShowHelpModal(false)}
        >
          <div 
            style={{ background: "var(--panel)", border: "1px solid var(--line)", borderRadius: "12px", width: "680px", maxWidth: "92vw", maxHeight: "85vh", display: "flex", flexDirection: "column", boxShadow: "0 20px 40px rgba(0,0,0,0.3)", overflow: "hidden" }}
            onClick={(e) => e.stopPropagation()}
          >
            {/* Modal Header */}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "16px 20px", borderBottom: "1px solid var(--line)" }}>
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                <Calculator size={20} color="var(--accent)" />
                <h3 style={{ margin: 0, fontSize: "16px", fontWeight: 700 }}>ADES Result Calculator & Moderation Guide</h3>
              </div>
              <button 
                type="button" 
                onClick={() => setShowHelpModal(false)}
                style={{ background: "transparent", border: "none", color: "var(--muted)", cursor: "pointer", padding: "4px" }}
              >
                <X size={18} />
              </button>
            </div>

            {/* Modal Body */}
            <div style={{ padding: "20px", overflowY: "auto", display: "flex", flexDirection: "column", gap: "16px", fontSize: "13px", lineHeight: "1.5" }}>
              
              <div>
                <h4 style={{ margin: "0 0 6px", fontSize: "14px", color: "var(--accent)" }}>1. Student-Course Level Grouping</h4>
                <p style={{ margin: 0, color: "var(--muted)" }}>
                  Groups raw assessment rows by <code>(Faculty, Program Term Name, Seat Number, PRN, Course Code, Course Name)</code> to aggregate individual component marks across Practical (PR) and Theory (TH).
                </p>
              </div>

              <div>
                <h4 style={{ margin: "0 0 6px", fontSize: "14px", color: "var(--accent)" }}>2. End Semester Exam (ESE) 30% Pass Rule</h4>
                <p style={{ margin: 0, color: "var(--muted)" }}>
                  <code>ESE - Max = ESE - PR Max + ESE - TH Max</code><br />
                  <code>ESE - Min = ceil(0.30 × ESE - Max)</code> (30% ceiling minimum)<br />
                  <code>ESE Overall = ESE - PR Obtained + ESE - TH Obtained</code><br />
                  <code>ESE Pass = "Pass" if ESE Overall &ge; ESE - Min else "Fail"</code>
                </p>
              </div>

              <div>
                <h4 style={{ margin: "0 0 6px", fontSize: "14px", color: "var(--accent)" }}>3. Continuous Evaluation (CE) Component</h4>
                <p style={{ margin: 0, color: "var(--muted)" }}>
                  <code>CE - Max = CE - PR Max + CE - TH Max</code><br />
                  <code>CE - Min = 0</code><br />
                  <code>CE Overall Marks = CE - PR Obtained + CE - TH Obtained</code>
                </p>
              </div>

              <div>
                <h4 style={{ margin: "0 0 6px", fontSize: "14px", color: "var(--accent)" }}>4. Course Aggregate & 35% Overall Pass Rule</h4>
                <p style={{ margin: 0, color: "var(--muted)" }}>
                  <code>Overall Maximum = ESE - Max + CE - Max</code><br />
                  <code>Overall Minimum = ceil(0.35 × Overall Maximum)</code> (35% ceiling minimum)<br />
                  <code>Course Overall Marks = ESE Overall + CE Overall Marks</code><br />
                  <code>Overall pass = "Pass" if Course Overall Marks &ge; Overall Minimum else "Fail"</code><br />
                  <code>Course Pass/Fail = "Pass" ONLY if BOTH ESE Pass and Overall Pass are "Pass"</code>
                </p>
              </div>

              <div>
                <h4 style={{ margin: "0 0 6px", fontSize: "14px", color: "#f59e0b" }}>5. Course Moderation Engine (Theory Only vs Practical-Only Toggle)</h4>
                <p style={{ margin: 0, color: "var(--muted)" }}>
                  &bull; <strong>Theory Rule (Default):</strong> Moderation is effected only for courses containing an <code>ESE Theory (TH)</code> component.<br />
                  &bull; <strong>Practical-Only Option:</strong> Courses that have solely <code>ESE - PR</code> components are excluded by default, but can be included by enabling the <strong>"Allow Moderation on Solely ESE - PR Courses"</strong> toggle in the Moderation Matrix tab.<br />
                  &bull; <strong>Formula:</strong> <code>Marks Needed = max(ESE Deficit, Overall Deficit)</code>. Awarded if <code>Marks Needed &le; Course Moderation Limit</code>.
                </p>
              </div>

              <div>
                <h4 style={{ margin: "0 0 6px", fontSize: "14px", color: "var(--accent)" }}>6. 31 Standardized Master Output Columns</h4>
                <p style={{ margin: 0, color: "var(--muted)" }}>
                  Exports complete formatted 31-column calculation sheets to Excel with sheet name <code>"Output file "</code>, including <code>ESE Pass</code>, <code>Overall pass</code>, <code>Course Pass/Fail</code>, and <code>Moderation Marks</code>.
                </p>
              </div>

            </div>

            {/* Modal Footer */}
            <div style={{ padding: "12px 20px", borderTop: "1px solid var(--line)", display: "flex", justifyContent: "flex-end", background: "var(--bg)" }}>
              <button 
                type="button" 
                onClick={() => setShowHelpModal(false)}
                style={{ padding: "6px 16px", background: "var(--accent)", color: "white", border: "none", borderRadius: "6px", fontWeight: 600, cursor: "pointer", fontSize: "12px" }}
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
