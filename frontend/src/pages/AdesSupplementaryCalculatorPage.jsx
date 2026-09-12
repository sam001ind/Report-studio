import React, { useState, useEffect, useMemo, useRef } from "react";
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
  Lock,
  FileCheck,
  Scale,
  GitCompare,
  ArrowRightLeft,
  FileText,
  Trash2,
  PanelLeftClose,
  PanelLeftOpen,
  Columns,
  Maximize2,
  Minimize2,
  Eye,
  EyeOff
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

// Robust parser for moderation/ordinance marks across various gazette representations:
// e.g. 3, "3", "3*", "*3", "O.4 (3)", "O.5042 : 3", "3 Marks", "+3"
export const parseOrdMarks = (val) => {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const str = String(val).trim();
  if (!str) return 0;
  const directNum = Number(str);
  if (!isNaN(directNum)) return directNum;
  
  const colonMatch = str.match(/:\s*(\d+(\.\d+)?)/);
  if (colonMatch) return Number(colonMatch[1]) || 0;
  const parenMatch = str.match(/\((\d+(\.\d+)?)\)/);
  if (parenMatch) return Number(parenMatch[1]) || 0;
  const numMatch = str.match(/(\d+(\.\d+)?)/);
  if (numMatch) return Number(numMatch[1]) || 0;
  return 0;
};

// Cleans and canonicalizes course codes: strips component suffixes like (TH), (PR), (T), (P), [TH], -TH,
// strips ., #, *, _, -, spaces, wrapping brackets, and converts to uppercase
export const cleanCourseCode = (rawCode) => {
  if (!rawCode) return "";
  let code = String(rawCode).trim();
  code = code.replace(/\s*\((TH|PR|T|P|THEORY|PRACTICAL|ESE|CE)\)\s*/gi, "");
  code = code.replace(/\s*\[(TH|PR|T|P|THEORY|PRACTICAL|ESE|CE)\]\s*/gi, "");
  code = code.replace(/[-_](TH|PR|T|P|THEORY|PRACTICAL|ESE|CE)\b/gi, "");
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

// Checks whether a candidate's course (by code or name) matches a gazette reappear entry
export const isCourseReappearInGazetteEntry = (courseCode, courseName, entry) => {
  if (!entry) return false;
  const targetCode = cleanCourseCode(courseCode);
  const targetName = cleanCourseName(courseName);
  const normTargetName = normalizeKey(targetName);

  // 1. Direct code check in parsed codes
  if (targetCode && (entry.reappearCodes || []).includes(targetCode)) {
    return true;
  }

  // 2. Check if rawReappear text contains the course code
  const rawUpper = String(entry.rawReappear || "").toUpperCase();
  if (targetCode && rawUpper.includes(targetCode)) {
    return true;
  }

  // 3. Title check in parsed titles or rawReappear text
  if (normTargetName && normTargetName.length >= 4) {
    for (const t of (entry.reappearTitles || [])) {
      const normT = normalizeKey(t);
      if (normT === normTargetName || normT.includes(normTargetName) || normTargetName.includes(normT)) {
        return true;
      }
    }
    if (normalizeKey(rawUpper).includes(normTargetName)) {
      return true;
    }
  }

  return false;
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

    // 1. Connection via PRN (Constant permanent identifier across exam sessions)
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

    // 2. Fallback via Seat Number only if PRN did not match or is absent
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

export default function AdesSupplementaryCalculatorPage() {
  const [sourceFile, setSourceFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [workbook, setWorkbook] = useState(null);
  const [rawRows, setRawRows] = useState([]);
  // Note: groupedRecords is computed dynamically and reactively via useMemo below
  const setGroupedRecords = () => {};
  const sourceFileInputRef = useRef(null);
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

  // Historical / Previous Events ADES Reports State (Multi-Report Baseline for Carry Forward)
  const [previousReports, setPreviousReports] = useState([]); // array of { id, fileName, recordCount, courseCount, parsedRecords, courseProfiles, uploadTime }
  const [historicalRecordsMap, setHistoricalRecordsMap] = useState(new Map());
  const [historicalCourseProfilesMap, setHistoricalCourseProfilesMap] = useState(new Map()); // Map of confirmed course component structures & counts
  const [improvementScoringMode, setImprovementScoringMode] = useState("current"); // "current" | "best"
  const [attemptTypeFilter, setAttemptTypeFilter] = useState("ALL"); // "ALL" | "IMPROVEMENT" | "SUPPLEMENTARY" | "NO_BASELINE"
  const prevReportsFileInputRef = useRef(null);

  // Previous Semester University Result Gazette Reports State (Optional Reappear Assurance)
  const [gazetteReports, setGazetteReports] = useState([]); // array of { id, fileName, recordCount, reappearCount, parsedGazetteRecords, uploadTime }
  const [historicalGazetteMap, setHistoricalGazetteMap] = useState(new Map()); // Map of cleanIdKey(prn/seat) => Array of gazette entries
  const gazetteFileInputRef = useRef(null);

  // Result & Ordinance Reconciliation Comparison Tool State
  const [comparisonFile, setComparisonFile] = useState(null);
  const [comparisonFileName, setComparisonFileName] = useState("");
  const [comparisonSheetNames, setComparisonSheetNames] = useState([]);
  const [comparisonSelectedSheet, setComparisonSelectedSheet] = useState("");
  const [comparisonWorkbook, setComparisonWorkbook] = useState(null);
  const [rawComparisonRows, setRawComparisonRows] = useState([]);
  const [comparisonFilterStatus, setComparisonFilterStatus] = useState("ALL"); // "ALL" | "MISMATCH" | "MOD_DIFF" | "EXACT_MATCH" | "PASSED_HERE_FAILED_PUB" | "FAILED_HERE_PASSED_PUB" | "HELD_MISMATCH" | "NOT_IN_DATA"
  const [comparisonCollegeFilter, setComparisonCollegeFilter] = useState("ALL");
  const [comparisonSearch, setComparisonSearch] = useState("");
  const [comparisonPage, setComparisonPage] = useState(0);
  const comparisonPageSize = 50;
  const rowsPerPage = comparisonPageSize;
  const comparisonFileInputRef = useRef(null);

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
  const [pageSize, setPageSize] = useState(50);
  const [showHelpModal, setShowHelpModal] = useState(false);
  const [sheetMetadata, setSheetMetadata] = useState({});

  // View Options & Sidebar Customization
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [sidebarWidth, setSidebarWidth] = useState(250); // Default reduced from 320 to 250
  const [tableDensity, setTableDensity] = useState("normal"); // "compact" | "normal" | "comfortable"
  const [columnPreset, setColumnPreset] = useState("ALL"); // "ALL" | "ESSENTIAL" | "SCORES" | "CUSTOM"
  const [visibleColumns, setVisibleColumns] = useState(() => new Set(ADES_OUTPUT_HEADERS));
  const [showColumnPicker, setShowColumnPicker] = useState(false);
  const [isTableMaximized, setIsTableMaximized] = useState(false);

  // MacBook Air M2 Vertical Accordion Navigation
  const [sidebarAccordions, setSidebarAccordions] = useState({
    files: true,       // Marksheet file & sheet selector
    baseline: false,   // Earlier ADES Baseline
    gazette: false,    // Result Gazette Reappear Assurance
    special: false,    // Absent, Malpractice, Heldback
    overview: false,   // Consolidated KPIs & Simulation
    rules: false       // Evaluation Rules
  });

  const toggleAccordion = (key) => {
    setSidebarAccordions(prev => ({ ...prev, [key]: !prev[key] }));
  };

  const setAllAccordions = (open) => {
    setSidebarAccordions({
      files: open,
      baseline: open,
      gazette: open,
      special: open,
      overview: open,
      rules: open
    });
  };

  // Trackpad / Mouse Drag-to-Resize Sidebar
  const isDraggingSidebar = useRef(false);
  const startDragX = useRef(0);
  const startWidth = useRef(250);

  const handleSidebarMouseDown = (e) => {
    e.preventDefault();
    isDraggingSidebar.current = true;
    startDragX.current = e.clientX;
    startWidth.current = sidebarWidth;
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';

    const handleMouseMove = (moveEvent) => {
      if (!isDraggingSidebar.current) return;
      const delta = moveEvent.clientX - startDragX.current;
      const newWidth = Math.max(180, Math.min(460, startWidth.current + delta));
      setSidebarWidth(newWidth);
    };

    const handleMouseUp = () => {
      isDraggingSidebar.current = false;
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
  };

  // Dynamic Column List based on user selected preset or custom picker
  const displayedHeaders = useMemo(() => {
    if (columnPreset === "ALL") return ADES_OUTPUT_HEADERS;
    if (columnPreset === "ESSENTIAL") {
      const essentialSet = new Set([
        "Course Code",
        "Course Name",
        "Seat Number",
        "PRN",
        "ESE - TH Obtained",
        "CE - TH Obtained",
        "Course Overall Marks ",
        "ESE Pass",
        "Course Pass/Fail",
        "Moderation Marks"
      ]);
      return ADES_OUTPUT_HEADERS.filter(h => essentialSet.has(h));
    }
    if (columnPreset === "SCORES") {
      const scoresSet = new Set([
        "Seat Number",
        "PRN",
        "Course Code",
        "Course Name",
        "ESE - PR Obtained",
        "ESE - TH Obtained",
        "ESE Overall",
        "CE - PR Obtained",
        "CE - TH Obtained",
        "CE Overall Marks ",
        "Course Overall Marks ",
        "ESE Pass",
        "Course Pass/Fail",
        "Moderation Marks"
      ]);
      return ADES_OUTPUT_HEADERS.filter(h => scoresSet.has(h));
    }
    if (columnPreset === "CUSTOM") {
      return ADES_OUTPUT_HEADERS.filter(h => visibleColumns.has(h));
    }
    return ADES_OUTPUT_HEADERS;
  }, [columnPreset, visibleColumns]);

  const setStatus = (msg, type = "info") => {
    setStatusMsg(msg);
    setStatusType(type);
  };


  // Helper to match student course record against loaded Absent Report (PRN is constant)
  const getAbsentEntry = (prn, seat, code, absentMap) => {
    if (!absentMap || absentMap.size === 0) return null;
    const normCode = normalizeKey(code);
    if (!normCode) return null;

    const normPrn = normalizeKey(prn);
    if (normPrn) {
      const entry = absentMap.get(`${normPrn}___${normCode}`);
      if (entry) return entry;
      return null; // PRN is the constant authoritative identifier; do not fall back to potentially mismatched seat
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

  // Helper to match student course record against loaded Malpractice Report (PRN is constant)
  const getMalpracticeEntry = (prn, seat, code, malpracticeMap) => {
    if (!malpracticeMap || malpracticeMap.size === 0) return null;
    const normCode = normalizeKey(code);
    if (!normCode) return null;

    const normPrn = normalizeKey(prn);
    if (normPrn) {
      const entry = malpracticeMap.get(`${normPrn}___${normCode}`);
      if (entry) return entry;
      return null; // PRN is the constant authoritative identifier
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

  // Helper to match student course record against loaded Heldback Report (PRN is constant)
  const getHeldbackEntry = (prn, seat, code, heldbackMap) => {
    if (!heldbackMap || heldbackMap.size === 0) return null;
    const normPrn = normalizeKey(prn);
    const normSeat = normalizeKey(seat);
    const normCode = normalizeKey(code);

    // 1. If PRN is present, use PRN as the constant authoritative identifier
    if (normPrn) {
      const termEntry = heldbackMap.get(`${normPrn}___ALL`);
      if (termEntry) return termEntry;

      if (normCode) {
        const paperEntry = heldbackMap.get(`${normPrn}___${normCode}`);
        if (paperEntry) return paperEntry;
      }
      return null; // PRN is authoritative; do not fall back to potentially mismatched seat
    }

    // 2. Only if PRN is completely missing, fall back to seat number
    if (normSeat) {
      const termEntry = heldbackMap.get(`${normSeat}___ALL`);
      if (termEntry) return termEntry;

      if (normCode) {
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

  // Published University Result Summary Comparison Handlers
  const handleComparisonFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setStatus(`Loading comparison file "${file.name}"...`, "info");
      const data = await file.arrayBuffer();
      // Performance: skip formatting/style parsing — only need cell values
      const wb = XLSX.read(data, {
        type: "array",
        cellStyles: false,
        cellNF: false,
        cellHTML: false,
        sheetStubs: false
      });
      setComparisonWorkbook(wb);
      setComparisonFile(file);
      setComparisonFileName(file.name);
      setComparisonSheetNames(wb.SheetNames);
      const initialSheet = wb.SheetNames[0] || "";
      setComparisonSelectedSheet(initialSheet);
      parseComparisonSheet(wb, initialSheet);
      setStatus(`Loaded comparison file "${file.name}" successfully.`, "success");
    } catch (err) {
      console.error(err);
      setStatus(`Failed to read comparison file: ${err.message}`, "error");
    }
  };

  const parseComparisonSheet = (wb, sheetName) => {
    if (!wb || !sheetName) return;
    const ws = wb.Sheets[sheetName];
    if (!ws) return;
    const rawData = XLSX.utils.sheet_to_json(ws, { header: 1 });
    if (!rawData || rawData.length === 0) {
      setRawComparisonRows([]);
      return;
    }

    // Auto-detect header row
    let headerRowIdx = -1;
    for (let r = 0; r < Math.min(25, rawData.length); r++) {
      const row = rawData[r] || [];
      const str = row.map(c => String(c || "").toLowerCase()).join(" ");
      if (str.includes("result status") || (str.includes("ordinance") && str.includes("prn")) || (str.includes("seat") && str.includes("result"))) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx === -1) headerRowIdx = 0;

    const headers = rawData[headerRowIdx] || [];
    const hMap = {};
    headers.forEach((h, idx) => {
      if (h) hMap[normalizeKey(h)] = idx;
    });

    const getCol = (r, ...aliases) => {
      for (const a of aliases) {
        const idx = hMap[normalizeKey(a)];
        if (idx !== undefined && r[idx] !== undefined && r[idx] !== null) {
          return String(r[idx]).trim();
        }
      }
      return "";
    };

    const parsed = [];
    for (let i = headerRowIdx + 1; i < rawData.length; i++) {
      const row = rawData[i];
      if (!row || row.length === 0) continue;
      const prn = getCol(row, "PRN-Permanent Registration Number", "PRN", "PRN Number", "PRNNo", "StudentID", "RegisterNo");
      const seat = getCol(row, "Examination Seat Number", "Seat Number", "SeatNumber", "SeatNo", "RollNo", "Roll Number");
      if (!prn && !seat) continue;

      const name = getCol(row, "Name of Student", "Student Name", "StudentName", "Candidate Name");
      const college = getCol(row, "College Name", "CollegeName", "ADEC Name", "ADECName", "College", "Institute Name");
      const collegeCode = getCol(row, "College Code", "CollegeCode", "ADEC Code", "ADECCode");
      const result = getCol(row, "RESULT STATUS", "Result Status", "Result", "Status");
      const otherStatus = getCol(row, "OTHER STATUS", "Other Status", "Remarks");
      const ord = getCol(row, "Ordinance", "Ordinance Applied");
      const ordTotal = parseNumber(getCol(row, "Ord Total", "OrdTotal", "Ordinance Total", "Moderation Marks")) || 0;
      const reappear = getCol(row, "Reappear Paper Codes", "Reappear", "Failed Papers", "Failed Subjects", "Reappear Papers");
      const failCount = parseNumber(getCol(row, "No of Fail Subjects", "Fail Subjects", "Fail Count")) || 0;
      const grandTotal = getCol(row, "Grand Total", "GrandTotal", "Total Marks");
      const sgpa = getCol(row, "SGPA");
      const cgpa = getCol(row, "CGPA");

      parsed.push({
        prn,
        seat,
        name,
        college: cleanCollegeName(college) || college,
        collegeCode,
        result: result || (otherStatus ? otherStatus : "Unknown"),
        otherStatus,
        ord: ord ? ord.toUpperCase() : (ordTotal > 0 ? "YES" : "NO"),
        ordTotal,
        reappear,
        failCount,
        grandTotal,
        sgpa,
        cgpa,
        rawRow: row
      });
    }

    setRawComparisonRows(parsed);
    setComparisonPage(0);
  };

  const handleClearComparisonFile = () => {
    setComparisonFile(null);
    setComparisonFileName("");
    setComparisonSheetNames([]);
    setComparisonSelectedSheet("");
    setComparisonWorkbook(null);
    setRawComparisonRows([]);
    setComparisonFilterStatus("ALL");
    setComparisonCollegeFilter("ALL");
    setComparisonSearch("");
    if (comparisonFileInputRef.current) comparisonFileInputRef.current.value = "";
    setStatus("Cleared comparison data.", "info");
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

  // Helper to normalize Student ID / Seat / PRN immune to Excel .0 or whitespace or scientific notation
  const cleanIdKey = (val) => {
    if (val === null || val === undefined) return "";
    let s = String(val).trim();
    if (/^[0-9]+(\.[0-9]+)?e\+[0-9]+$/i.test(s)) {
      try {
        s = BigInt(Math.round(Number(s))).toString();
      } catch (e) {}
    }
    return s.replace(/\.0+$/, "").replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
  };

  // Lookup student-course record from historical baseline map (PRN is the constant identifier)
  const getHistoricalEntry = (prn, seat, code, histMap) => {
    if (!histMap || histMap.size === 0) return null;
    const cleanP = cleanIdKey(prn);
    const cleanS = cleanIdKey(seat);
    const cleanC = cleanIdKey(cleanCourseCode(code));
    if (!cleanC) return null;
    
    // PRN is constant across exam sessions; prioritize PRN lookup!
    if (cleanP) {
      const key1 = `${cleanP}_${cleanC}`;
      if (histMap.has(key1)) return histMap.get(key1);
      return null; // PRN is constant; do not fall back to a different student's seat
    }
    // Only fall back to seat number if PRN is completely missing
    if (cleanS) {
      const key2 = `${cleanS}_${cleanC}`;
      if (histMap.has(key2)) return histMap.get(key2);
    }
    return null;
  };

  // Helper to detect if an uploaded sheet is a University Semester Tabulation Gazette / Result Summary Register
  const isGazetteOrTabulationSheet = (hMap) => {
    if (!hMap) return false;
    const keys = Object.keys(hMap);
    const hasReappear = keys.some(k => k.includes("reappear") || k.includes("failsubject") || k.includes("nooffail") || k.includes("arrear") || k.includes("backlog"));
    const hasResultStatus = keys.some(k => k.includes("resultstatus") || k.includes("statementnumber") || k.includes("result") || k.includes("status"));
    const hasPrnOrSeat = keys.some(k => k.includes("prn") || k.includes("seat") || k.includes("reg") || k.includes("roll"));
    return hasPrnOrSeat && (hasReappear || hasResultStatus);
  };

  // Extract student-level result gazette and reappear records from a previous semester tabulation register
  const extractGazetteRecordsFromRows = (rows, hMap, fileName) => {
    const list = [];
    if (!rows || rows.length === 0 || !hMap) return list;

    // Discover column keys once per sheet for fast row iteration
    const hKeys = Object.keys(hMap);
    const fuzzyPrnKey = hKeys.find(k => k.includes("prn") || (k.includes("reg") && !k.includes("regular")));
    const fuzzySeatKey = hKeys.find(k => k.includes("seat") || k.includes("roll") || k.includes("hallticket") || k.includes("usn"));
    const fuzzyNameKey = hKeys.find(k => k.includes("studentname") || k.includes("candidatename") || (k.includes("name") && !k.includes("course") && !k.includes("program") && !k.includes("college")));
    const fuzzyReappearKey = hKeys.find(k => k.includes("reappear") || k.includes("backlog") || k.includes("fail") || k.includes("arrear") || k.includes("supplementary") || k.includes("duepaper"));
    const fuzzyStatusKey = hKeys.find(k => k.includes("resultstatus") || k === "result" || k.includes("status") || k.includes("remark"));
    const fuzzyTermKey = hKeys.find(k => k.includes("term") || k.includes("semester") || k.includes("sem"));
    const fuzzyOrdKey = hKeys.find(k => 
      k.includes("ordtotal") || 
      k.includes("ordinancetotal") || 
      k.includes("modmarks") || 
      k.includes("moderationmarks") || 
      k.includes("gracemarks") || 
      k.includes("moderation") ||
      k.includes("condonation") ||
      (k.includes("ord") && !k.includes("order") && !k.includes("accord") && !k.includes("record") && !k.includes("board") && !k.includes("coord") && !k.includes("word") && !k.includes("ford"))
    );

    rows.forEach(row => {
      let prn = String(getCell(row, hMap, 
        "PRN-Permanent Registration Number", 
        "PRN - Permanent Registration Number",
        "Permanent Registration Number",
        "Candidate Register Number",
        "Candidate Register No",
        "University Register Number",
        "Univ Reg No",
        "PRN Number", 
        "PRN No", 
        "PRN", 
        "PRNNo", 
        "Registration Number",
        "Register Number",
        "Register No",
        "RegisterNo",
        "Reg No",
        "RegNo",
        "Regd No",
        "Regd Number",
        "Enrollment No",
        "Enrollment Number",
        "Candidate Code",
        "StudentID", 
        "Student ID"
      ) || "").trim();

      if (!prn && fuzzyPrnKey && hMap[fuzzyPrnKey]) {
        prn = String(row[hMap[fuzzyPrnKey]] || "").trim();
      }

      let seat = String(getCell(row, hMap, 
        "Examination Seat Number", 
        "Exam Seat Number",
        "Exam Seat No",
        "Candidate Seat Number",
        "Seat Number", 
        "SeatNumber", 
        "Seat No", 
        "SeatNo", 
        "Seat",
        "Roll Number", 
        "Roll No", 
        "RollNo",
        "Hall Ticket Number",
        "Hall Ticket No",
        "HallTicketNo",
        "Hall Ticket",
        "USN"
      ) || "").trim();

      if (!seat && fuzzySeatKey && hMap[fuzzySeatKey]) {
        seat = String(row[hMap[fuzzySeatKey]] || "").trim();
      }

      if (!prn && !seat) return;

      let studentName = String(getCell(row, hMap, "Name of Student", "Student Name", "StudentName", "Candidate Name", "Name") || "").trim();
      if (!studentName && fuzzyNameKey && hMap[fuzzyNameKey]) {
        studentName = String(row[hMap[fuzzyNameKey]] || "").trim();
      }

      let term = String(getCell(row, hMap, "Program Part Term", "ProgramPartTerm", "Term", "Semester", "Sem") || "").trim();
      if (!term && fuzzyTermKey && hMap[fuzzyTermKey]) {
        term = String(row[hMap[fuzzyTermKey]] || "").trim();
      }

      const examEvent = String(getCell(row, hMap, "Exam Event", "ExamEvent", "Event", "Session") || "").trim();

      let program = String(getCell(row, hMap, "Program Name", "Program Full Name", "Program", "Programme", "Degree", "Course") || "").trim();
      if (!program && fuzzyProgKey && hMap[fuzzyProgKey]) {
        program = String(row[hMap[fuzzyProgKey]] || "").trim();
      }

      let resultStatus = String(getCell(row, hMap, "RESULT STATUS", "Result Status", "Result", "Status", "Remarks") || "").trim();
      if (!resultStatus && fuzzyStatusKey && hMap[fuzzyStatusKey]) {
        resultStatus = String(row[hMap[fuzzyStatusKey]] || "").trim();
      }

      let rawReappear = String(getCell(row, hMap, 
        "Reappear Paper Codes", 
        "Reappear Paper Code", 
        "Reappear Papers", 
        "Reappear Paper", 
        "Reappear", 
        "Failed Papers", 
        "Failed Paper", 
        "Failed Subjects", 
        "Failed Subject", 
        "Fail Subjects", 
        "Backlog Papers",
        "Backlogs",
        "Arrear Papers",
        "Arrears",
        "Supplementary Papers"
      ) || "").trim();
      if (!rawReappear && fuzzyReappearKey && hMap[fuzzyReappearKey]) {
        rawReappear = String(row[hMap[fuzzyReappearKey]] || "").trim();
      }

      let rawOrd = parseOrdMarks(getCell(row, hMap, 
        "Ord Total", 
        "OrdTotal", 
        "Ordinance Total", 
        "Ordinance Marks",
        "Ord Marks",
        "Moderation Marks", 
        "Moderation Total", 
        "Mod Total", 
        "Mod Marks", 
        "Total Moderation Marks", 
        "Total Mod Marks",
        "Ordinance (Total)",
        "Ordinance(Total)",
        "Grace Marks", 
        "Grace", 
        "Condonation Marks",
        "Ordinance", 
        "Ord"
      ));
      if ((!rawOrd || rawOrd === 0) && fuzzyOrdKey && hMap[fuzzyOrdKey]) {
        rawOrd = parseOrdMarks(row[hMap[fuzzyOrdKey]]);
      }
      if (!rawOrd || rawOrd === 0) {
        for (const k of Object.keys(row)) {
          const normK = normalizeKey(k);
          if (
            (normK.includes("ordinance") || normK.includes("ordtotal") || normK.includes("modmark") || normK.includes("moderation")) &&
            !normK.includes("record") && !normK.includes("order") && !normK.includes("coord")
          ) {
            const v = parseOrdMarks(row[k]);
            if (v > 0) {
              rawOrd = v;
              break;
            }
          }
        }
      }
      const ordTotal = rawOrd || 0;

      const rawFailCount = getCell(row, hMap, "No of Fail Subjects", "Fail Subjects", "Fail Count", "Backlog Count", "No of Arrears");
      const rawText = rawReappear.trim();
      const codes = new Set();
      const titles = new Set();

      if (rawText) {
        // 1. Extract alphanumeric course codes via regex (e.g. KU2DSCCSC110, 24CCAR038, etc.)
        const codeMatches = rawText.match(/[A-Za-z]{1,6}\d{1,2}[A-Za-z0-9]{3,}/g) || [];
        codeMatches.forEach(m => {
          const cleaned = cleanCourseCode(m);
          if (cleaned) codes.add(cleaned);
        });

        // 2. Split by comma, semicolon, newline, slash, pipe, or bullet
        const chunks = rawText.split(/[,;\n\r|•]+|\s*\/\s*/).map(s => s.trim()).filter(Boolean);
        chunks.forEach(chunk => {
          if (chunk.includes("-") || chunk.includes(":")) {
            const parts = chunk.split(/[-:]+/).map(p => p.trim());
            parts.forEach(p => {
              const cleaned = cleanCourseCode(p);
              if (cleaned && cleaned.length >= 5 && /\d/.test(cleaned)) {
                codes.add(cleaned);
              } else if (p.length > 2) {
                titles.add(p);
              }
            });
          } else {
            const cleaned = cleanCourseCode(chunk);
            if (cleaned && cleaned.length >= 5 && /\d/.test(cleaned)) {
              codes.add(cleaned);
            } else if (chunk.length > 2) {
              titles.add(chunk);
            }
          }
        });
      }

      const reappearCodes = Array.from(codes);
      const reappearTitles = Array.from(titles);
      const failCount = rawFailCount !== "" && rawFailCount !== null && rawFailCount !== undefined 
        ? (parseInt(rawFailCount, 10) || 0) 
        : (reappearCodes.length > 0 ? reappearCodes.length : reappearTitles.length);

      list.push({
        prn,
        seat,
        studentName,
        term,
        examEvent,
        program,
        resultStatus: resultStatus ? resultStatus.toUpperCase() : "UNKNOWN",
        reappearCodes,
        reappearTitles,
        rawReappear: rawText,
        failCount,
        ordTotal,
        grandTotal: String(getCell(row, hMap, "Grand Total", "Total Marks", "Total") || "").trim(),
        sgpa: String(getCell(row, hMap, "SGPA") || "").trim(),
        cgpa: String(getCell(row, hMap, "CGPA") || "").trim(),
        sourceFile: fileName
      });
    });

    return list;
  };

  // Rebuild the unified historical gazette composite map across all uploaded previous event reports
  const rebuildHistoricalGazetteMap = (reports) => {
    const map = new Map();
    (reports || []).forEach(rep => {
      const records = rep.parsedGazetteRecords || rep.records || [];
      if (Array.isArray(records)) {
        records.forEach(entry => {
          const cleanP = cleanIdKey(entry.prn);
          const cleanS = cleanIdKey(entry.seat);
          const cleanN = cleanIdKey(entry.studentName);
          if (cleanP) {
            if (!map.has(cleanP)) map.set(cleanP, []);
            map.get(cleanP).push(entry);
          }
          if (cleanS && cleanS !== cleanP) {
            if (!map.has(cleanS)) map.set(cleanS, []);
            map.get(cleanS).push(entry);
          }
          if (cleanN && cleanN !== cleanP && cleanN !== cleanS && cleanN.length > 3) {
            if (!map.has(cleanN)) map.set(cleanN, []);
            map.get(cleanN).push(entry);
          }
        });
      }
    });
    return map;
  };

  // Cross-verify a candidate's course attempt against uploaded Gazette/Tabulation records (PRN is constant)
  const verifyCourseAttempt = (prn, seat, courseCode, gazetteMap, courseName = "", studentName = "") => {
    if (!gazetteMap || gazetteMap.size === 0) return null;
    const cleanP = cleanIdKey(prn);
    const cleanS = cleanIdKey(seat);
    const cleanN = cleanIdKey(studentName);
    
    // PRN is constant across exam sessions; strictly prioritize PRN lookup!
    let entries = cleanP ? (gazetteMap.get(cleanP) || []) : [];
    
    // Fall back to seat number or name if PRN lookup yielded no entries
    if (entries.length === 0 && cleanS) {
      entries = gazetteMap.get(cleanS) || [];
    }
    if (entries.length === 0 && cleanN && cleanN.length > 3) {
      entries = gazetteMap.get(cleanN) || [];
    }
    if (entries.length === 0) return null;

    const targetCode = cleanCourseCode(courseCode);
    const targetName = cleanCourseName(courseName);
    const normTargetName = normalizeKey(targetName);
    let isVerifiedReappear = false;
    let matchingTerm = null;
    let matchingEvent = null;
    let allPendingReappears = [];
    let isPriorPass = false;

    for (const entry of entries) {
      const isReappear = isCourseReappearInGazetteEntry(courseCode, courseName, entry);
      if (isReappear) {
        isVerifiedReappear = true;
        matchingTerm = entry.term;
        matchingEvent = entry.examEvent;
      }

      // Prior pass check: ONLY if this entry does NOT list this course as a reappear,
      // and the entry explicitly has a PASS result status and 0 fails
      if (!isReappear && (entry.resultStatus === "PASS" || (entry.failCount === 0 && (entry.reappearCodes || []).length === 0 && entry.resultStatus !== "FAIL"))) {
        isPriorPass = true;
      }

      (entry.reappearCodes || []).forEach(code => {
        if (code !== targetCode && !allPendingReappears.includes(code)) {
          allPendingReappears.push(code);
        }
      });
      (entry.reappearTitles || []).forEach(title => {
        const normT = normalizeKey(title);
        if (normT && normT !== normTargetName && !allPendingReappears.includes(title)) {
          allPendingReappears.push(title);
        }
      });
    }

    // A course can NEVER be a Prior Pass (Improvement) if it is officially a Verified Reappear (Supplementary)!
    if (isVerifiedReappear) {
      isPriorPass = false;
    }

    return {
      hasGazetteRecord: true,
      isVerifiedReappear,
      isPriorPass,
      matchingTerm,
      matchingEvent,
      allPendingReappears,
      latestStatus: entries[0].resultStatus,
      failCount: entries[0].failCount,
      sourceFile: entries[0].sourceFile
    };
  };

  // Rebuild the unified historical composite map from all uploaded previous event reports
  const rebuildHistoricalMap = (reports) => {
    const map = new Map();
    reports.forEach(rep => {
      (rep.parsedRecords || []).forEach(rec => {
        const cleanP = cleanIdKey(rec.prn);
        const cleanS = cleanIdKey(rec.seat);
        const cleanC = cleanIdKey(cleanCourseCode(rec.code));
        if (cleanP && cleanC) {
          map.set(`${cleanP}_${cleanC}`, rec);
        }
        if (cleanS && cleanC) {
          map.set(`${cleanS}_${cleanC}`, rec);
        }
      });
    });
    return map;
  };

  // Extract verified course component structures & component counts from a previous event report
  const extractHistoricalCourseProfilesFromRows = (rows, hMap, fileName) => {
    const profiles = new Map();
    const isAgg = isAlreadyAggregatedSheet(hMap);

    rows.forEach(row => {
      const rawCode = String(getCell(row, hMap, "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || "").trim();
      const code = cleanCourseCode(rawCode);
      const norm = normalizeKey(code);
      if (!norm) return;

      const courseName = cleanCourseName(String(getCell(row, hMap, "Course Name", "CourseName", "PaperName", "SubjectName", "CourseTitle") || "").trim());

      if (!profiles.has(norm)) {
        profiles.set(norm, {
          courseCode: code,
          courseName,
          hasSeenEseTh: false,
          hasSeenEsePr: false,
          hasSeenCeTh: false,
          hasSeenCePr: false,
          eseThMax: 0,
          eseThMin: 0,
          esePrMax: 0,
          esePrMin: 0,
          ceThMax: 0,
          ceThMin: 0,
          cePrMax: 0,
          cePrMin: 0,
          eseMax: 0,
          eseMin: 0,
          ceMax: 0,
          ceMin: 0,
          courseMax: 0,
          courseMin: 0,
          sourceFiles: new Set([fileName])
        });
      }

      const p = profiles.get(norm);
      p.sourceFiles.add(fileName);
      if (courseName && !p.courseName) p.courseName = courseName;

      if (isAgg) {
        const ese_th_m = parseNumber(getCell(row, hMap, "ESE - TH Max", "ESETHMax", "ESE-TH Max", "ESE_TH Max", "ESE TH Max", "Theory Max", "TH Max"));
        const ese_th_min = parseNumber(getCell(row, hMap, "ESE - TH Min", "ESETHMin", "ESE-TH Min", "ESE_TH Min", "ESE TH Min", "Theory Min", "TH Min"));
        const ese_pr_m = parseNumber(getCell(row, hMap, "ESE - PR Max", "ESEPRMax", "ESE-PR Max", "ESE_PR Max", "ESE PR Max", "Practical Max", "PR Max"));
        const ese_pr_min = parseNumber(getCell(row, hMap, "ESE - PR Min", "ESEPRMin", "ESE-PR Min", "ESE_PR Min", "ESE PR Min", "Practical Min", "PR Min"));
        const raw_ce_th_m = parseNumber(getCell(row, hMap, "CE - TH Max", "CETHMax", "CE-TH Max", "CE_TH Max", "CE TH Max", "Internal Max", "CA Max", "IA Max"));
        const raw_ce_th_min = parseNumber(getCell(row, hMap, "CE - TH Min", "CETHMin", "CE-TH Min", "CE_TH Min", "CE TH Min", "Internal Min", "CA Min", "IA Min"));
        const ce_pr_m = parseNumber(getCell(row, hMap, "CE - PR Max", "CEPRMax", "CE-PR Max", "CE_PR Max", "CE PR Max", "Practical Internal Max"));
        const ce_pr_min = parseNumber(getCell(row, hMap, "CE - PR Min", "CEPRMin", "CE-PR Min", "CE_PR Min", "CE PR Min", "Practical Internal Min"));
        const ese_m = parseNumber(getCell(row, hMap, "ESE - Max", "ESEMax", "ESE Max", "ESE_Max", "ESE-Max", "External Max"));
        const ese_min = parseNumber(getCell(row, hMap, "ESE - Min", "ESEMin", "ESE Min", "ESE_Min", "ESE-Min", "External Min"));
        const ce_m_raw = parseNumber(getCell(row, hMap, "CE - Max", "CEMax", "CE Max", "CE_Max", "CE-Max", "Internal Max", "CA Max", "Continuous Evaluation Max"));
        const ce_min_raw = parseNumber(getCell(row, hMap, "CE - Min", "CEMin", "CE Min", "CE_Min", "CE-Min", "Internal Min", "CA Min"));
        const overall_m = parseNumber(getCell(row, hMap, "Overall Maximum", "OverallMax", "Overall Max", "Course Max", "CourseMax", "Total Max", "Max Marks"));
        const overall_min = parseNumber(getCell(row, hMap, "Overall Minimum", "OverallMin", "Overall Min", "Course Min", "CourseMin", "Total Min", "Min Marks"));

        const ce_th_m = raw_ce_th_m !== null ? raw_ce_th_m : ((ce_pr_m === null || ce_pr_m === 0) ? ce_m_raw : null);
        const ce_th_min = raw_ce_th_min !== null ? raw_ce_th_min : ((ce_pr_min === null || ce_pr_min === 0) ? ce_min_raw : null);
        const ce_m = ce_m_raw !== null ? ce_m_raw : ((ce_th_m || 0) + (ce_pr_m || 0) || null);
        const ce_min = ce_min_raw !== null ? ce_min_raw : null;

        if (ese_th_min !== null) p.eseThMin = ese_th_min;
        if (ese_pr_min !== null) p.esePrMin = ese_pr_min;
        if (ce_th_min !== null) p.ceThMin = ce_th_min;
        if (ce_pr_min !== null) p.cePrMin = ce_pr_min;
        if (ese_m !== null && ese_m > 0) {
          const calcEseMin = Math.ceil(0.30 * ese_m);
          p.eseMin = Math.max(ese_min || 0, calcEseMin);
        } else if (ese_min !== null) {
          p.eseMin = ese_min;
        }
        if (ce_min !== null) p.ceMin = ce_min;
        if (overall_m !== null && overall_m > 0) {
          const calcCourseMin = Math.ceil(0.35 * overall_m);
          p.courseMin = Math.max(overall_min || 0, calcCourseMin);
        } else if (overall_min !== null) {
          p.courseMin = overall_min;
        }

        const ese_th_obt = getCell(row, hMap, "ESE - TH Obtained", "ESETHObtained", "ESE-TH Obtained", "ESE_TH Obtained", "ESE TH Obtained", "ESE - TH Marks", "ESE-TH Marks", "ESE TH Marks", "ESE - TH", "ESE-TH", "ESE_TH", "ESE TH", "ESE Theory", "Theory Obtained", "Theory");
        const ese_pr_obt = getCell(row, hMap, "ESE - PR Obtained", "ESEPRObtained", "ESE-PR Obtained", "ESE_PR Obtained", "ESE PR Obtained", "ESE - PR Marks", "ESE-PR Marks", "ESE PR Marks", "ESE - PR", "ESE-PR", "ESE_PR", "ESE PR", "ESE Practical", "Practical Obtained", "Practical", "PR Obtained", "PR");
        const ce_th_obt = getCell(row, hMap, "CE - TH Obtained", "CETHObtained", "CE-TH Obtained", "CE_TH Obtained", "CE TH Obtained", "CE - TH Marks", "CE-TH Marks", "CE TH Marks", "CE - TH", "CE-TH", "CE_TH", "CE TH", "CE Obtained", "CE Marks", "CA", "IA");
        const ce_pr_obt = getCell(row, hMap, "CE - PR Obtained", "CEPRObtained", "CE-PR Obtained", "CE_PR Obtained", "CE PR Obtained", "CE - PR Marks", "CE-PR Marks", "CE PR Marks", "CE - PR", "CE-PR", "CE_PR", "CE PR", "CE Practical");

        if ((ese_th_m !== null && ese_th_m > 0) || (ese_th_obt !== undefined && ese_th_obt !== null && String(ese_th_obt).trim() !== "")) {
          p.hasSeenEseTh = true;
          if (ese_th_m > 0) p.eseThMax = Math.max(p.eseThMax, ese_th_m);
        }
        if ((ese_pr_m !== null && ese_pr_m > 0) || (ese_pr_obt !== undefined && ese_pr_obt !== null && String(ese_pr_obt).trim() !== "")) {
          p.hasSeenEsePr = true;
          if (ese_pr_m > 0) p.esePrMax = Math.max(p.esePrMax, ese_pr_m);
        }
        if ((ce_th_m !== null && ce_th_m > 0) || (ce_th_obt !== undefined && ce_th_obt !== null && String(ce_th_obt).trim() !== "")) {
          p.hasSeenCeTh = true;
          if (ce_th_m > 0) p.ceThMax = Math.max(p.ceThMax, ce_th_m);
        }
        if ((ce_pr_m !== null && ce_pr_m > 0) || (ce_pr_obt !== undefined && ce_pr_obt !== null && String(ce_pr_obt).trim() !== "")) {
          p.hasSeenCePr = true;
          if (ce_pr_m > 0) p.cePrMax = Math.max(p.cePrMax, ce_pr_m);
        }
        if (ese_m > 0) p.eseMax = Math.max(p.eseMax, ese_m);
        if (ce_m > 0) p.ceMax = Math.max(p.ceMax, ce_m);
        if (overall_m > 0) p.courseMax = Math.max(p.courseMax, overall_m);
      } else {
        const methodRaw = String(getCell(row, hMap, "Assessment Method", "AssessmentMethod", "AM", "Method") || "").trim().toUpperCase();
        const typeRaw = String(getCell(row, hMap, "Assessment Type", "AssessmentType", "AT", "Type") || "").trim().toUpperCase();
        const atMaxRaw = parseNumber(getCell(row, hMap, "AT Max Marks", "ATMaxMarks", "MaxMarks", "Max Marks", "Max", "AT Max"));
        const amMaxRaw = parseNumber(getCell(row, hMap, "AM Max Marks", "AMMaxMarks", "AM Max"));
        const courseMaxRaw = parseNumber(getCell(row, hMap, "Course Max", "CourseMax", "Overall Maximum", "OverallMax"));

        const isEse = methodRaw.includes("ESE") || methodRaw.includes("EXT") || methodRaw.includes("EXTERNAL") || methodRaw.includes("THEORY");
        const isCe = methodRaw.includes("CE") || methodRaw.includes("CA") || methodRaw.includes("IA") || methodRaw.includes("INTERNAL");
        const isPr = typeRaw.includes("PR") || typeRaw.includes("PRACTICAL") || typeRaw.includes("VIVA") || typeRaw.includes("LAB");
        const isTh = typeRaw.includes("TH") || typeRaw.includes("THEORY");

        if (isEse) {
          if (isTh || (!isPr && !isTh)) {
            p.hasSeenEseTh = true;
            if (atMaxRaw > 0) p.eseThMax = Math.max(p.eseThMax, atMaxRaw);
          }
          if (isPr) {
            p.hasSeenEsePr = true;
            if (atMaxRaw > 0) p.esePrMax = Math.max(p.esePrMax, atMaxRaw);
          }
          if (amMaxRaw > 0) p.eseMax = Math.max(p.eseMax, amMaxRaw);
        } else if (isCe) {
          if (isTh || (!isPr && !isTh)) {
            p.hasSeenCeTh = true;
            if (atMaxRaw > 0) p.ceThMax = Math.max(p.ceThMax, atMaxRaw);
          }
          if (isPr) {
            p.hasSeenCePr = true;
            if (atMaxRaw > 0) p.cePrMax = Math.max(p.cePrMax, atMaxRaw);
          }
          if (amMaxRaw > 0) p.ceMax = Math.max(p.ceMax, amMaxRaw);
        }
        if (courseMaxRaw > 0) p.courseMax = Math.max(p.courseMax, courseMaxRaw);
      }
    });

    profiles.forEach(p => {
      if (p.eseMax > 0 && p.eseThMax > 0 && p.eseMax > p.eseThMax) {
        p.requiresEseTh = true;
        p.requiresEsePr = true;
        if (p.esePrMax === 0) p.esePrMax = p.eseMax - p.eseThMax;
      } else if (p.hasSeenEseTh && p.hasSeenEsePr) {
        p.requiresEseTh = true;
        p.requiresEsePr = true;
      } else if (p.hasSeenEsePr && !p.hasSeenEseTh) {
        p.requiresEseTh = false;
        p.requiresEsePr = true;
      } else {
        p.requiresEseTh = true;
        p.requiresEsePr = false;
      }

      if (p.ceMax > 0 && p.ceThMax > 0 && p.ceMax > p.ceThMax) {
        p.requiresCeTh = true;
        p.requiresCePr = true;
        if (p.cePrMax === 0) p.cePrMax = p.ceMax - p.ceThMax;
      } else if (p.hasSeenCeTh && p.hasSeenCePr) {
        p.requiresCeTh = true;
        p.requiresCePr = true;
      } else if (p.hasSeenCePr && !p.hasSeenCeTh) {
        p.requiresCeTh = false;
        p.requiresCePr = true;
      } else {
        p.requiresCeTh = true;
        p.requiresCePr = false;
      }

      if (p.requiresCeTh && !p.requiresCePr && (p.ceThMax === 0 || !p.ceThMax) && p.ceMax > 0) {
        p.ceThMax = p.ceMax;
      }
      if (p.requiresCePr && !p.requiresCeTh && (p.cePrMax === 0 || !p.cePrMax) && p.ceMax > 0) {
        p.cePrMax = p.ceMax;
      }
      if (p.requiresEseTh && !p.requiresEsePr && (p.eseThMax === 0 || !p.eseThMax) && p.eseMax > 0) {
        p.eseThMax = p.eseMax;
      }
      if (p.requiresEsePr && !p.requiresEseTh && (p.esePrMax === 0 || !p.esePrMax) && p.eseMax > 0) {
        p.esePrMax = p.eseMax;
      }

      const expectedComps = [];
      if (p.requiresCeTh) expectedComps.push("CE-TH");
      if (p.requiresCePr) expectedComps.push("CE-PR");
      if (p.requiresEseTh) expectedComps.push("ESE-TH");
      if (p.requiresEsePr) expectedComps.push("ESE-PR");

      p.expectedComponents = expectedComps;
      p.componentCount = expectedComps.length;
      p.sourceFiles = Array.from(p.sourceFiles);
    });

    return profiles;
  };

  // Combine course profiles across all uploaded previous reports
  const rebuildHistoricalCourseProfiles = (reports) => {
    const map = new Map();
    reports.forEach(rep => {
      if (rep.courseProfiles) {
        rep.courseProfiles.forEach((prof, k) => {
          if (!map.has(k)) {
            map.set(k, { ...prof, sourceFiles: [...prof.sourceFiles] });
          } else {
            const existing = map.get(k);
            existing.requiresEseTh = existing.requiresEseTh || prof.requiresEseTh;
            existing.requiresEsePr = existing.requiresEsePr || prof.requiresEsePr;
            existing.requiresCeTh = existing.requiresCeTh || prof.requiresCeTh;
            existing.requiresCePr = existing.requiresCePr || prof.requiresCePr;
            existing.eseThMax = Math.max(existing.eseThMax, prof.eseThMax);
            existing.esePrMax = Math.max(existing.esePrMax, prof.esePrMax);
            existing.ceThMax = Math.max(existing.ceThMax, prof.ceThMax);
            existing.cePrMax = Math.max(existing.cePrMax, prof.cePrMax);
            existing.eseMax = Math.max(existing.eseMax, prof.eseMax);
            existing.ceMax = Math.max(existing.ceMax, prof.ceMax);
            existing.courseMax = Math.max(existing.courseMax, prof.courseMax);
            prof.sourceFiles.forEach(f => {
              if (!existing.sourceFiles.includes(f)) existing.sourceFiles.push(f);
            });
            const expectedComps = [];
            if (existing.requiresCeTh) expectedComps.push("CE-TH");
            if (existing.requiresCePr) expectedComps.push("CE-PR");
            if (existing.requiresEseTh) expectedComps.push("ESE-TH");
            if (existing.requiresEsePr) expectedComps.push("ESE-PR");
            existing.expectedComponents = expectedComps;
            existing.componentCount = expectedComps.length;
          }
        });
      }
    });
    return map;
  };

  // Extract historical course records from a previous event worksheet
  const extractHistoricalRecordsFromRows = (rows, hMap, fileName) => {
    const list = [];
    const isAgg = isAlreadyAggregatedSheet(hMap);
    
    if (isAgg) {
      rows.forEach(row => {
        const prn = String(getCell(row, hMap, "PRN", "PRN Number", "PRNNo", "RegisterNo", "RegNo", "StudentID") || "").trim();
        const seat = String(getCell(row, hMap, "Seat Number", "SeatNumber", "SeatNo", "Seat_Number", "RollNo", "Roll Number") || "").trim();
        const rawCode = String(getCell(row, hMap, "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || "").trim();
        const code = cleanCourseCode(rawCode);
        const name = cleanCourseName(String(getCell(row, hMap, "Course Name", "CourseName", "PaperName", "SubjectName", "CourseTitle") || "").trim());
        if (!code || (!prn && !seat)) return;

        const ese_th_obtained = parseNumber(getCell(row, hMap, 
          "ESE - TH Obtained", "ESETHObtained", "ESE-TH Obtained", "ESE_TH Obtained", "ESE TH Obtained", 
          "ESE - TH Marks", "ESE-TH Marks", "ESE_TH Marks", "ESE TH Marks", "ESE TH Mark", 
          "ESE - TH", "ESE-TH", "ESE_TH", "ESE TH", 
          "ESE Theory Obtained", "ESE Theory Marks", "ESE Theory", 
          "Theory Obtained", "Theory Marks", "Theory Mark", "Theory", 
          "TH Obtained", "TH Marks", "TH Mark", "TH"
        )) ?? "";

        let ese_pr_obtained = parseNumber(getCell(row, hMap, 
          "ESE - PR Obtained", "ESEPRObtained", "ESE-PR Obtained", "ESE_PR Obtained", "ESE PR Obtained", 
          "ESE - PR Marks", "ESE-PR Marks", "ESE_PR Marks", "ESE PR Marks", "ESE PR Mark", 
          "ESE - PR", "ESE-PR", "ESE_PR", "ESE PR", 
          "ESE Practical Obtained", "ESE Practical Marks", "ESE Practical", 
          "Practical Obtained", "Practical Marks", "Practical Mark", "Practical", "Practicals", 
          "PR Obtained", "PR Marks", "PR Mark", "PR", "Lab Obtained", "Lab Marks", "Lab"
        )) ?? "";

        let ese_obtained = parseNumber(getCell(row, hMap, "ESE Overall", "ESEOverall", "ESE Overall Marks ", "ESE Overall Marks", "ESE Total", "ESE", "ESE Marks"));
        if (ese_obtained === null && (ese_pr_obtained !== "" || ese_th_obtained !== "")) {
          ese_obtained = (parseNumber(ese_pr_obtained) || 0) + (parseNumber(ese_th_obtained) || 0);
        }
        // If sheet has ESE Overall and ESE-TH but ESE-PR column was omitted/blank, infer ESE-PR
        if (ese_pr_obtained === "" && ese_obtained !== null && ese_th_obtained !== "" && ese_obtained > parseNumber(ese_th_obtained)) {
          ese_pr_obtained = ese_obtained - parseNumber(ese_th_obtained);
        }
        let ce_th_obtained = parseNumber(getCell(row, hMap, "CE - TH Obtained", "CETHObtained", "CE-TH Obtained", "CE_TH Obtained", "CE TH Obtained", "CE-TH", "CE_TH", "CE TH", "CE Obtained", "CE Marks", "CE Total", "CA", "IA", "Internal", "Continuous Evaluation")) ?? "";
        let ce_pr_obtained = parseNumber(getCell(row, hMap, "CE - PR Obtained", "CEPRObtained", "CE-PR Obtained", "CE_PR Obtained", "CE PR Obtained", "CE-PR", "CE_PR", "CE PR", "CE Practical Obtained", "CE Practical")) ?? "";
        let ce_obtained = parseNumber(getCell(row, hMap, "CE Overall Marks ", "CE Overall Marks", "CEOverallMarks", "CEOverall", "CE Total", "CE Marks", "CE Obtained", "CE"));
        if (ce_obtained === null) {
          ce_obtained = (parseNumber(ce_pr_obtained) || 0) + (parseNumber(ce_th_obtained) || 0);
        }
        // If sheet has CE Overall and no separate CE-PR, infer CE-TH from CE Overall
        if (ce_th_obtained === "" && ce_obtained !== null && (ce_pr_obtained === "" || ce_pr_obtained === 0)) {
          ce_th_obtained = ce_obtained;
        }
        let course_overall = parseNumber(getCell(row, hMap, "Course Overall Marks ", "Course Overall Marks", "CourseOverallMarks"));
        if (course_overall === null) {
          course_overall = (parseNumber(ese_obtained) || 0) + (parseNumber(ce_obtained) || 0);
        }

        const rawEseThMax = parseNumber(getCell(row, hMap, "ESE - TH Max", "ESETHMax", "ESE-TH Max", "ESE_TH Max", "ESE TH Max", "Theory Max", "TH Max"));
        const rawEseThMin = parseNumber(getCell(row, hMap, "ESE - TH Min", "ESETHMin", "ESE-TH Min", "ESE_TH Min", "ESE TH Min", "Theory Min", "TH Min"));
        const rawEsePrMax = parseNumber(getCell(row, hMap, "ESE - PR Max", "ESEPRMax", "ESE-PR Max", "ESE_PR Max", "ESE PR Max", "Practical Max", "PR Max"));
        const rawEsePrMin = parseNumber(getCell(row, hMap, "ESE - PR Min", "ESEPRMin", "ESE-PR Min", "ESE_PR Min", "ESE PR Min", "Practical Min", "PR Min"));
        const rawEseMax = parseNumber(getCell(row, hMap, "ESE - Max", "ESEMax", "ESE Max", "ESE_Max", "ESE-Max", "External Max"));
        const rawEseMin = parseNumber(getCell(row, hMap, "ESE - Min", "ESEMin", "ESE Min", "ESE_Min", "ESE-Min", "External Min"));

        const ese_max = rawEseMax ?? ((rawEsePrMax || 0) + (rawEseThMax || 0) || 50);
        const calculatedEseMin = Math.ceil(0.30 * (ese_max || 0));
        const ese_min = Math.max(rawEseMin || 0, calculatedEseMin);

        const rawCeThMax = parseNumber(getCell(row, hMap, "CE - TH Max", "CETHMax", "CE-TH Max", "CE_TH Max", "CE TH Max", "Internal Max", "CA Max", "IA Max"));
        const rawCeThMin = parseNumber(getCell(row, hMap, "CE - TH Min", "CETHMin", "CE-TH Min", "CE_TH Min", "CE TH Min", "Internal Min", "CA Min", "IA Min"));
        const rawCePrMax = parseNumber(getCell(row, hMap, "CE - PR Max", "CEPRMax", "CE-PR Max", "CE_PR Max", "CE PR Max", "Practical Internal Max"));
        const rawCePrMin = parseNumber(getCell(row, hMap, "CE - PR Min", "CEPRMin", "CE-PR Min", "CE_PR Min", "CE PR Min", "Practical Internal Min"));
        const rawCeMax = parseNumber(getCell(row, hMap, "CE - Max", "CEMax", "CE Max", "CE_Max", "CE-Max", "Internal Max", "CA Max", "Continuous Evaluation Max"));
        const rawCeMin = parseNumber(getCell(row, hMap, "CE - Min", "CEMin", "CE Min", "CE_Min", "CE-Min", "Internal Min", "CA Min"));

        const ce_max = rawCeMax ?? ((rawCePrMax || 0) + (rawCeThMax || 0) || null);
        const ce_min = rawCeMin ?? 0;
        const effectiveCeThMax = rawCeThMax !== null ? rawCeThMax : ((rawCePrMax === null || rawCePrMax === 0) && ce_max !== null ? ce_max : null);
        const effectiveCeThMin = rawCeThMin !== null ? rawCeThMin : ((rawCePrMin === null || rawCePrMin === 0) && ce_min !== null ? ce_min : 0);
        const effectiveEseThMax = rawEseThMax !== null ? rawEseThMax : ((rawEsePrMax === null || rawEsePrMax === 0) && ese_max !== null ? ese_max : null);
        const effectiveEseThMin = rawEseThMin !== null ? rawEseThMin : ((rawEsePrMin === null || rawEsePrMin === 0) && ese_min !== null ? ese_min : 0);

        const rawOverallMax = parseNumber(getCell(row, hMap, "Overall Maximum", "OverallMax", "Overall Max", "Course Max", "CourseMax", "Total Max", "Max Marks"));
        const rawOverallMin = parseNumber(getCell(row, hMap, "Overall Minimum", "OverallMin", "Overall Min", "Course Min", "CourseMin", "Total Min", "Min Marks"));

        const overall_max = rawOverallMax ?? ((ese_max || 0) + (ce_max || 0) || 100);
        const calculatedOverallMin = Math.ceil(0.35 * (overall_max || 0));
        const overall_min = Math.max(rawOverallMin || 0, calculatedOverallMin);

        const numEse = parseNumber(ese_obtained) || 0;
        const numOverall = parseNumber(course_overall) || 0;
        const raw_ese_pass = numEse > 0 && numEse >= ese_min;
        const raw_overall_pass = numOverall >= overall_min;
        const raw_course_pass = raw_ese_pass && raw_overall_pass;

        list.push({
          prn,
          seat,
          code,
          name,
          ese_pr_obtained,
          ese_th_obtained,
          ese_obtained,
          ce_pr_obtained,
          ce_th_obtained,
          ce_obtained,
          course_overall,
          raw_ese_pass,
          raw_overall_pass,
          raw_course_pass,
          ese_pr_max: rawEsePrMax,
          ese_pr_min: rawEsePrMin,
          ese_th_max: effectiveEseThMax,
          ese_th_min: effectiveEseThMin,
          ese_max,
          ese_min,
          ce_pr_max: rawCePrMax,
          ce_pr_min: rawCePrMin,
          ce_th_max: effectiveCeThMax,
          ce_th_min: effectiveCeThMin,
          ce_max,
          ce_min,
          overall_max,
          overall_min,
          sourceFile: fileName
        });
      });
    } else {
      const tempGroup = new Map();
      rows.forEach(row => {
        const prn = String(getCell(row, hMap, "PRN", "PRN Number", "PRNNo", "RegisterNo") || "").trim();
        const seat = String(getCell(row, hMap, "Seat Number", "SeatNumber", "SeatNo") || "").trim();
        const rawCode = String(getCell(row, hMap, "Course Code", "CourseCode", "PaperCode") || "").trim();
        const code = cleanCourseCode(rawCode);
        if (!code || (!prn && !seat)) return;

        const k = `${prn}_${code}`;
        if (!tempGroup.has(k)) {
          tempGroup.set(k, {
            prn,
            seat,
            code,
            name: cleanCourseName(String(getCell(row, hMap, "Course Name", "CourseName") || "").trim()),
            ese_pr_obtained: "",
            ese_th_obtained: "",
            ce_pr_obtained: "",
            ce_th_obtained: "",
            sourceFile: fileName
          });
        }
        const item = tempGroup.get(k);
        const am = String(getCell(row, hMap, "Assessment Method", "AssessmentMethod", "AM", "Method", "Assessment_Method") || "").toUpperCase();
        const at = String(getCell(row, hMap, "Assessment Type", "AssessmentType", "AT", "Type", "Assessment_Type") || "").toUpperCase();
        const marks = parseNumber(getCell(row, hMap, "Assessment Marks", "Marks", "ObtainedMarks", "Mark", "Obtained"));
        const isEse = am.includes("ESE") || am.includes("EXT") || am.includes("THEORY");
        const isCe = am.includes("CE") || am.includes("CA") || am.includes("IA") || am.includes("INTERNAL");
        const isPr = at.includes("PR") || at.includes("PRACTICAL") || at.includes("VIVA") || at.includes("LAB");
        const isTh = at.includes("TH") || (!isPr && isEse);
        if (isEse && isTh) item.ese_th_obtained = marks ?? item.ese_th_obtained;
        if (isEse && isPr) item.ese_pr_obtained = marks ?? item.ese_pr_obtained;
        if (isCe && isTh) item.ce_th_obtained = marks ?? item.ce_th_obtained;
        if (isCe && isPr) item.ce_pr_obtained = marks ?? item.ce_pr_obtained;
      });

      tempGroup.forEach(item => {
        const ese_obtained = (parseNumber(item.ese_pr_obtained) || 0) + (parseNumber(item.ese_th_obtained) || 0);
        const ce_obtained = (parseNumber(item.ce_pr_obtained) || 0) + (parseNumber(item.ce_th_obtained) || 0);
        const course_overall = ese_obtained + ce_obtained;
        item.ese_obtained = ese_obtained;
        item.ce_obtained = ce_obtained;
        item.course_overall = course_overall;
        item.raw_ese_pass = ese_obtained >= 15;
        item.raw_overall_pass = course_overall >= 35;
        item.raw_course_pass = item.raw_ese_pass && item.raw_overall_pass;
        list.push(item);
      });
    }

    return list;
  };

  // Group raw assessment rows OR parse pre-aggregated rows into Student-Course Base Aggregates
  const buildGroupedRecordsFromRows = (rows, currentHeaderMap, currentAbsentMap = absentRecordsMap, currentMalpracticeMap = malpracticeRecordsMap, currentHeldbackMap = heldbackRecordsMap, currentHistoricalMap = historicalRecordsMap, currentImprovementMode = improvementScoringMode, currentHistoricalCourseProfiles = historicalCourseProfilesMap, currentGazetteMap = historicalGazetteMap) => {
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
        const ese_max_raw = parseNumber(getCell(row, currentHeaderMap, "ESE - Max", "ESEMax", "ESE Max", "ESE_Max", "External Max"));
        const ese_th_m = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Max", "ESETHMax", "ESE-TH Max", "ESE_TH Max", "ESE TH Max", "ESE_TH_Max", "Theory Max", "TH Max"));
        const ese_pr_m = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Max", "ESEPRMax", "ESE-PR Max", "ESE_PR Max", "ESE PR Max", "ESE_PR_Max", "Practical Max", "PR Max"));
        const ce_max_raw = parseNumber(getCell(row, currentHeaderMap, "CE - Max", "CEMax", "CE Max", "CE_Max", "CE-Max", "Internal Max", "CA Max", "Continuous Evaluation Max"));
        const raw_ce_th_m = parseNumber(getCell(row, currentHeaderMap, "CE - TH Max", "CETHMax", "CE-TH Max", "CE_TH Max", "CE TH Max", "CE_TH_Max", "Internal Max", "CA Max", "IA Max"));
        const ce_pr_m = parseNumber(getCell(row, currentHeaderMap, "CE - PR Max", "CEPRMax", "CE-PR Max", "CE_PR Max", "CE PR Max", "CE_PR_Max", "Practical Internal Max"));
        const ce_th_m = raw_ce_th_m !== null ? raw_ce_th_m : ((ce_pr_m === null || ce_pr_m === 0) ? ce_max_raw : null);
        const overall_max_raw = parseNumber(getCell(row, currentHeaderMap, "Overall Maximum", "OverallMaximum", "OverallMax", "Course Max", "CourseMax", "Total Max", "Max Marks"));

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
    courseExpectedComponentsMap.forEach((prof, norm) => {
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
          // Example: XT9DSCAFZ206 (ESE Max 65 > ESE TH Max 50, PR Max 15)
          prof.requiresEseTh = true;
          prof.requiresEsePr = true;
          if (prof.esePrMax === 0) prof.esePrMax = prof.eseMax - prof.eseThMax;
          prof.maxMarks.ESE_PR = Math.max(prof.maxMarks.ESE_PR, prof.esePrMax);
        } else if (prof.eseThMax > 0 && prof.eseMax === prof.eseThMax) {
          // ESE TH accounts for the entire ESE Max -> Theory Only, NO ESE PR!
          // Example: XT9DSCCOM208 (ESE Max 70 === ESE TH Max 70)
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

        // Propagate ceMax to ceThMax when course has no practical
        if (prof.requiresCeTh && !prof.requiresCePr && (prof.ceThMax === 0 || !prof.ceThMax) && prof.ceMax > 0) {
          prof.ceThMax = prof.ceMax;
          prof.maxMarks.CE_TH = prof.ceMax;
        }
        if (prof.requiresCePr && !prof.requiresCeTh && (prof.cePrMax === 0 || !prof.cePrMax) && prof.ceMax > 0) {
          prof.cePrMax = prof.ceMax;
          prof.maxMarks.CE_PR = prof.ceMax;
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
        } else {
          // Mandatory University Course Rule:
          // Every university degree course in CBCSS/FYUGP/ADES requires Continuous Evaluation (CE)!
          // In Supplementary & Improvement exams, CE is conducted in the regular semester and is NOT re-tested here.
          // Therefore, if CE marks/columns were omitted from the current supplementary upload,
          // the course STILL requires CE (CE-PR if practical-only, otherwise CE-TH).
          if (prof.isPrOnly || (prof.hasSeenEsePr && !prof.hasSeenEseTh)) {
            prof.requiresCePr = true;
            prof.requiresCeTh = false;
          } else {
            prof.requiresCeTh = true;
            prof.requiresCePr = false;
          }
          if (prof.ceMax === 0) {
            const courseStr = `${prof.courseCode || ""} ${norm || ""} ${prof.courseName || ""}`.toUpperCase();
            const isAec = courseStr.includes("AEC");
            const isVac = courseStr.includes("VAC");
            const isMdc = courseStr.includes("MDC");

            if (prof.courseMax > 0 && prof.courseMax > prof.eseMax) {
              prof.ceMax = prof.courseMax - prof.eseMax;
            } else if (isAec || isVac) {
              // Ability Enhancement Courses (AEC) and Value Added Courses (VAC) in Kerala FYUGP / ADES regulations have CE Max = 15!
              prof.ceMax = 15;
              prof.courseMax = (prof.eseMax || 0) + 15;
            } else if (isMdc) {
              // Multi-Disciplinary Courses (MDC): 3 credits, ESE 50, CE 25, Total 75
              prof.ceMax = 25;
              prof.courseMax = (prof.eseMax || 0) + 25;
            } else if (prof.eseMax === 75) {
              prof.ceMax = 25;
              prof.courseMax = 100;
            } else if (prof.eseMax === 70) {
              prof.ceMax = 30;
              prof.courseMax = 100;
            } else if (prof.eseMax === 80) {
              prof.ceMax = 20;
              prof.courseMax = 100;
            } else if (prof.eseMax === 60) {
              prof.ceMax = 40;
              prof.courseMax = 100;
            } else if (prof.eseMax === 50) {
              if (isAec || isVac) {
                prof.ceMax = 15;
                prof.courseMax = 65;
              } else {
                prof.ceMax = 25;
                prof.courseMax = 75;
              }
            } else if (prof.eseMax > 0) {
              prof.ceMax = Math.round(prof.eseMax / 3);
              prof.courseMax = prof.eseMax + prof.ceMax;
            }
          }
          if (prof.requiresCeTh && prof.ceThMax === 0) {
            prof.ceThMax = prof.ceMax;
            prof.maxMarks.CE_TH = prof.ceMax;
          }
          if (prof.requiresCePr && prof.cePrMax === 0) {
            prof.cePrMax = prof.ceMax;
            prof.maxMarks.CE_PR = prof.ceMax;
          }
        }
      }

      // Cross-check & Confirm course component requirements from uploaded previous event report baseline
      const normKey = norm || normalizeKey(prof?.courseCode);
      const histProf = currentHistoricalCourseProfiles?.get(normKey);
      if (histProf) {
        prof.requiresEseTh = histProf.requiresEseTh;
        prof.requiresEsePr = histProf.requiresEsePr;
        prof.requiresCeTh = histProf.requiresCeTh;
        prof.requiresCePr = histProf.requiresCePr;
        prof.expectedComponents = [...histProf.expectedComponents];
        prof.componentCount = histProf.componentCount;
        prof.confirmedFromPrevious = true;
        if (histProf.eseThMax > 0) prof.eseThMax = histProf.eseThMax;
        else if (histProf.eseMax > 0 && !histProf.requiresEsePr) prof.eseThMax = histProf.eseMax;
        if (histProf.eseThMin !== undefined) prof.eseThMin = histProf.eseThMin;
        if (histProf.esePrMax > 0) prof.esePrMax = histProf.esePrMax;
        else if (histProf.eseMax > 0 && histProf.requiresEsePr && !histProf.requiresEseTh) prof.esePrMax = histProf.eseMax;
        if (histProf.esePrMin !== undefined) prof.esePrMin = histProf.esePrMin;
        if (histProf.ceThMax > 0) prof.ceThMax = histProf.ceThMax;
        else if (histProf.ceMax > 0 && !histProf.requiresCePr) prof.ceThMax = histProf.ceMax;
        if (histProf.ceThMin !== undefined) prof.ceThMin = histProf.ceThMin;
        if (histProf.cePrMax > 0) prof.cePrMax = histProf.cePrMax;
        else if (histProf.ceMax > 0 && histProf.requiresCePr && !histProf.requiresCeTh) prof.cePrMax = histProf.ceMax;
        if (histProf.cePrMin !== undefined) prof.cePrMin = histProf.cePrMin;
        if (histProf.eseMax > 0) prof.eseMax = histProf.eseMax;
        if (histProf.eseMin !== undefined) prof.eseMin = histProf.eseMin;
        if (histProf.ceMax > 0) prof.ceMax = histProf.ceMax;
        if (histProf.ceMin !== undefined) prof.ceMin = histProf.ceMin;
        if (histProf.courseMax > 0) prof.courseMax = histProf.courseMax;
        if (histProf.courseMin !== undefined) prof.courseMin = histProf.courseMin;

        if (prof.ceThMax > 0) prof.maxMarks.CE_TH = prof.ceThMax;
        if (prof.cePrMax > 0) prof.maxMarks.CE_PR = prof.cePrMax;
        if (prof.eseThMax > 0) prof.maxMarks.ESE_TH = prof.eseThMax;
        if (prof.esePrMax > 0) prof.maxMarks.ESE_PR = prof.esePrMax;
      } else {
        const expectedComps = [];
        if (prof.requiresCeTh) expectedComps.push("CE-TH");
        if (prof.requiresCePr) expectedComps.push("CE-PR");
        if (prof.requiresEseTh) expectedComps.push("ESE-TH");
        if (prof.requiresEsePr) expectedComps.push("ESE-PR");
        prof.expectedComponents = expectedComps;
        prof.componentCount = expectedComps.length;
        prof.confirmedFromPrevious = false;
      }

      if (prof.eseMax > 0) {
        const calcEseMin = Math.ceil(0.30 * prof.eseMax);
        prof.eseMin = Math.max(prof.eseMin || 0, calcEseMin);
      }
      if (prof.courseMax > 0) {
        const calcCourseMin = Math.ceil(0.35 * prof.courseMax);
        prof.courseMin = Math.max(prof.courseMin || 0, calcCourseMin);
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

        // Historical baseline lookup & Attempt Type determination
        const rawCat = String(getCell(row, currentHeaderMap, "Exam Category", "ExamCategory", "Appearance Type", "AppearanceType", "Exam Type", "ExamType", "Appearance", "Exam_Category", "Exam_Type") || "").trim().toUpperCase();
        const hist = getHistoricalEntry(prn, seat, code, currentHistoricalMap);
        const gazetteAssurance = verifyCourseAttempt(prn, seat, code, currentGazetteMap, name);

        const hasImp = rawCat.includes("IMP") || rawCat.includes("IMPROVE");
        const hasSupp = rawCat.includes("SUPP") || rawCat.includes("BACKLOG") || rawCat.includes("REPEATER") || rawCat.includes("RE-APPEAR") || rawCat.includes("REAPPEAR");

        let is_improvement = false;
        let attempt_type = "SUPPLEMENTARY";
        if (gazetteAssurance && gazetteAssurance.isVerifiedReappear) {
          is_improvement = false;
          attempt_type = "SUPPLEMENTARY";
        } else if (hasImp && !hasSupp) {
          is_improvement = true;
          attempt_type = "IMPROVEMENT";
        } else if (hasSupp) {
          is_improvement = false;
          attempt_type = "SUPPLEMENTARY";
        } else if (gazetteAssurance && gazetteAssurance.isPriorPass && !hasSupp) {
          is_improvement = true;
          attempt_type = "IMPROVEMENT";
        } else if (hist && (hist.raw_course_pass === true || (hist.course_overall !== null && hist.course_overall >= 35 && hist.raw_ese_pass))) {
          is_improvement = true;
          attempt_type = "IMPROVEMENT";
        } else {
          is_improvement = false;
          attempt_type = "SUPPLEMENTARY";
        }

        // Check if this student-course has a heldback, malpractice or absent record
        const heldbackEntry = getHeldbackEntry(prn, seat, code, currentHeldbackMap);
        const malpracticeEntry = !heldbackEntry && getMalpracticeEntry(prn, seat, code, currentMalpracticeMap);
        const absentEntry = !heldbackEntry && !malpracticeEntry && getAbsentEntry(prn, seat, code, currentAbsentMap);

        const rawC = rawCollegeCode || heldbackEntry?.collegeCode || malpracticeEntry?.collegeCode || absentEntry?.collegeCode || "";
        const rawN = rawCollegeName || heldbackEntry?.collegeName || malpracticeEntry?.collegeName || absentEntry?.collegeName || "";
        const { collegeCode, collegeName, college } = collegeRegistry.resolve(seat, prn, rawC, rawN);

        const rowEsePrMax = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Max", "ESEPRMax", "ESE-PR Max", "ESE_PR Max", "ESE PR Max", "Practical Max", "PR Max"));
        const ese_pr_max = rowEsePrMax !== null ? rowEsePrMax : (hist?.ese_pr_max ?? (prof?.requiresEsePr ? (prof?.maxMarks?.ESE_PR || prof?.esePrMax || "") : ""));

        const rowEsePrMin = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Min", "ESEPRMin", "ESE-PR Min", "ESE_PR Min", "ESE PR Min", "Practical Min", "PR Min"));
        const ese_pr_min = rowEsePrMin !== null ? rowEsePrMin : (hist?.ese_pr_min ?? (prof?.esePrMin ?? (ese_pr_max !== "" ? 0 : "")));

        let ese_pr_obtained = parseNumber(getCell(row, currentHeaderMap, 
          "ESE - PR Obtained", "ESEPRObtained", "ESE-PR Obtained", "ESE_PR Obtained", "ESE PR Obtained", 
          "ESE - PR Marks", "ESE-PR Marks", "ESE_PR Marks", "ESE PR Marks", "ESE PR Mark", 
          "ESE - PR", "ESE-PR", "ESE_PR", "ESE PR", 
          "ESE Practical Obtained", "ESE Practical Marks", "ESE Practical", 
          "Practical Obtained", "Practical Marks", "Practical Mark", "Practical", "Practicals", 
          "PR Obtained", "PR Marks", "PR Mark", "PR", "Lab Obtained", "Lab Marks", "Lab"
        )) ?? "";

        const rowEseThMax = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Max", "ESETHMax", "ESE-TH Max", "ESE_TH Max", "ESE TH Max", "Theory Max", "TH Max"));
        const ese_th_max = rowEseThMax !== null ? rowEseThMax : (hist?.ese_th_max ?? (prof?.requiresEseTh ? (prof?.maxMarks?.ESE_TH || prof?.eseThMax || "") : ""));

        const rowEseThMin = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Min", "ESETHMin", "ESE-TH Min", "ESE_TH Min", "ESE TH Min", "Theory Min", "TH Min"));
        const ese_th_min = rowEseThMin !== null ? rowEseThMin : (hist?.ese_th_min ?? (prof?.eseThMin ?? (ese_th_max !== "" ? 0 : "")));

        let ese_th_obtained = parseNumber(getCell(row, currentHeaderMap, 
          "ESE - TH Obtained", "ESETHObtained", "ESE-TH Obtained", "ESE_TH Obtained", "ESE TH Obtained", 
          "ESE - TH Marks", "ESE-TH Marks", "ESE_TH Marks", "ESE TH Marks", "ESE TH Mark", 
          "ESE - TH", "ESE-TH", "ESE_TH", "ESE TH", 
          "ESE Theory Obtained", "ESE Theory Marks", "ESE Theory", 
          "Theory Obtained", "Theory Marks", "Theory Mark", "Theory", 
          "TH Obtained", "TH Marks", "TH Mark", "TH"
        )) ?? "";

        const rowCePrMax = parseNumber(getCell(row, currentHeaderMap, "CE - PR Max", "CEPRMax", "CE-PR Max", "CE_PR Max", "CE PR Max", "Practical Internal Max"));
        const ce_pr_max = rowCePrMax !== null ? rowCePrMax : (hist?.ce_pr_max ?? (prof?.requiresCePr ? (prof?.maxMarks?.CE_PR || prof?.cePrMax || "") : ""));

        const rowCePrMin = parseNumber(getCell(row, currentHeaderMap, "CE - PR Min", "CEPRMin", "CE-PR Min", "CE_PR Min", "CE PR Min", "Practical Internal Min"));
        const ce_pr_min = rowCePrMin !== null ? rowCePrMin : (hist?.ce_pr_min ?? (prof?.cePrMin ?? (ce_pr_max !== "" ? 0 : "")));

        const rowCePr = parseNumber(getCell(row, currentHeaderMap, "CE - PR Obtained", "CEPRObtained", "CE-PR Obtained", "CE_PR Obtained", "CE PR Obtained", "CE-PR", "CE_PR", "CE PR", "CE Practical Obtained", "CE Practical"));

        const rowCeThMax = parseNumber(getCell(row, currentHeaderMap, "CE - TH Max", "CETHMax", "CE-TH Max", "CE_TH Max", "CE TH Max", "CE - Max", "CEMax", "CE Max", "CE_Max", "Internal Max", "CA Max", "IA Max"));
        const ce_th_max = rowCeThMax !== null ? rowCeThMax : (hist?.ce_th_max ?? (prof?.requiresCeTh ? (prof?.maxMarks?.CE_TH || prof?.ceThMax || (!prof?.requiresCePr ? (prof?.ceMax || hist?.ce_max) : null) || "") : ""));

        const rowCeThMin = parseNumber(getCell(row, currentHeaderMap, "CE - TH Min", "CETHMin", "CE-TH Min", "CE_TH Min", "CE TH Min", "CE - Min", "CEMin", "CE Min", "Internal Min", "CA Min", "IA Min"));
        const ce_th_min = rowCeThMin !== null ? rowCeThMin : (hist?.ce_th_min ?? (prof?.ceThMin ?? (ce_th_max !== "" ? 0 : "")));

        const rowCeTh = parseNumber(getCell(row, currentHeaderMap, "CE - TH Obtained", "CETHObtained", "CE-TH Obtained", "CE_TH Obtained", "CE TH Obtained", "CE-TH", "CE_TH", "CE TH", "CE Obtained", "CE Marks", "CA", "IA", "Internal"));
        const rowCeOverall = parseNumber(getCell(row, currentHeaderMap, "CE Overall Marks ", "CE Overall Marks", "CEOverallMarks", "CEOverall", "CE Total", "CE Marks", "CE Obtained", "CE"));

        let ce_th_raw = rowCeTh !== null ? rowCeTh : ((rowCeOverall !== null && !prof?.requiresCePr) ? rowCeOverall : "");
        let ce_pr_raw = rowCePr !== null ? rowCePr : "";

        let carriedForwardComponents = [];
        let repeatedComponents = [];
        const missingComponents = [];
        const isBlank = (val) => val === undefined || val === null || String(val).trim() === "";

        // Verify Candidate Components against Confirmed Course Structure
        if (is_improvement) {
          // Improvement candidate: ONLY ESE-TH is taken from current attempt
          repeatedComponents.push("ESE-TH");
          if (prof?.requiresEseTh) {
            const rawTh = getCell(row, currentHeaderMap, "ESE - TH Obtained", "ESETHObtained", "ESE Overall", "ESEOverall");
            if (isBlank(rawTh)) {
              missingComponents.push("ESE-TH");
              ese_th_obtained = "Missing";
            } else {
              const currVal = parseNumber(rawTh) || 0;
              if (currentImprovementMode === "best" && hist && hist.ese_th_obtained !== "" && hist.ese_th_obtained !== null) {
                const histVal = parseNumber(hist.ese_th_obtained);
                ese_th_obtained = (histVal !== null && currVal < histVal) ? histVal : currVal;
              } else {
                ese_th_obtained = currVal;
              }
            }
          }
          // ESE-PR: strictly carried forward from previous baseline along with CE component
          if (prof?.requiresEsePr || (hist && !isBlank(hist.ese_pr_obtained))) {
            if (hist && !isBlank(hist.ese_pr_obtained)) {
              ese_pr_obtained = hist.ese_pr_obtained;
              carriedForwardComponents.push("ESE-PR");
            } else if (hist && hist.ese_obtained !== null && !isBlank(hist.ese_th_obtained) && hist.ese_obtained > parseNumber(hist.ese_th_obtained)) {
              ese_pr_obtained = hist.ese_obtained - parseNumber(hist.ese_th_obtained);
              carriedForwardComponents.push("ESE-PR");
            } else if (!isBlank(ese_pr_obtained)) {
              // already in current row
            } else {
              missingComponents.push("ESE-PR");
              ese_pr_obtained = "Missing";
            }
          }
          // CE-TH: check current row first, or carry forward from hist
          if (prof?.requiresCeTh) {
            if (!isBlank(ce_th_raw)) {
              // Present in current attempt row
            } else if (hist && !isBlank(hist.ce_th_obtained)) {
              ce_th_raw = hist.ce_th_obtained;
              carriedForwardComponents.push("CE-TH");
            } else if (hist && !isBlank(hist.ce_obtained) && !prof?.requiresCePr) {
              ce_th_raw = hist.ce_obtained;
              carriedForwardComponents.push("CE-TH");
            } else {
              missingComponents.push("CE-TH");
              ce_th_raw = "Missing";
            }
          }
          // CE-PR: check current row first, or carry forward from hist
          if (prof?.requiresCePr) {
            if (!isBlank(ce_pr_raw)) {
              // Present in current attempt row
            } else if (hist && !isBlank(hist.ce_pr_obtained)) {
              ce_pr_raw = hist.ce_pr_obtained;
              carriedForwardComponents.push("CE-PR");
            } else {
              missingComponents.push("CE-PR");
              ce_pr_raw = "Missing";
            }
          }
        } else {
          // Supplementary candidate: BOTH ESE-TH and ESE-PR are taken from current attempt!
          if (prof?.requiresEseTh) {
            const rawTh = getCell(row, currentHeaderMap, "ESE - TH Obtained", "ESETHObtained", "ESE Overall", "ESEOverall");
            if (isBlank(rawTh)) {
              missingComponents.push("ESE-TH");
              ese_th_obtained = "Missing";
            } else {
              ese_th_obtained = parseNumber(rawTh);
              repeatedComponents.push("ESE-TH");
            }
          }
          if (prof?.requiresEsePr) {
            const rawPr = getCell(row, currentHeaderMap, "ESE - PR Obtained", "ESEPRObtained");
            if (isBlank(rawPr)) {
              // If blank in current row, check if hist has it
              if (hist && !isBlank(hist.ese_pr_obtained)) {
                ese_pr_obtained = hist.ese_pr_obtained;
                carriedForwardComponents.push("ESE-PR");
              } else {
                missingComponents.push("ESE-PR");
                ese_pr_obtained = "Missing";
              }
            } else {
              ese_pr_obtained = parseNumber(rawPr);
              repeatedComponents.push("ESE-PR");
            }
          }
          // CE-TH: check current row first, or carry forward from hist
          if (prof?.requiresCeTh) {
            if (!isBlank(ce_th_raw)) {
              // Present in current attempt row
            } else if (hist && !isBlank(hist.ce_th_obtained)) {
              ce_th_raw = hist.ce_th_obtained;
              carriedForwardComponents.push("CE-TH");
            } else if (hist && !isBlank(hist.ce_obtained) && !prof?.requiresCePr) {
              ce_th_raw = hist.ce_obtained;
              carriedForwardComponents.push("CE-TH");
            } else {
              missingComponents.push("CE-TH");
              ce_th_raw = "Missing";
            }
          }
          // CE-PR: check current row first, or carry forward from hist
          if (prof?.requiresCePr) {
            if (!isBlank(ce_pr_raw)) {
              // Present in current attempt row
            } else if (hist && !isBlank(hist.ce_pr_obtained)) {
              ce_pr_raw = hist.ce_pr_obtained;
              carriedForwardComponents.push("CE-PR");
            } else {
              missingComponents.push("CE-PR");
              ce_pr_raw = "Missing";
            }
          }
        }

        let ese_max = parseNumber(getCell(row, currentHeaderMap, "ESE - Max", "ESEMax", "ESE Max", "ESE_Max", "ESE-Max", "External Max"));
        if (ese_max === null || ese_max === 0) {
          if (hist && hist.ese_max !== null && hist.ese_max !== undefined && hist.ese_max > 0) {
            ese_max = hist.ese_max;
          } else if (prof?.eseMax > 0) {
            ese_max = prof.eseMax;
          } else {
            ese_max = (parseNumber(ese_pr_max) || 0) + (parseNumber(ese_th_max) || 0);
          }
        }
        let ese_min = parseNumber(getCell(row, currentHeaderMap, "ESE - Min", "ESEMin", "ESE Min", "ESE_Min", "ESE-Min", "External Min"));
        const calculatedEseMin = Math.ceil(0.30 * (ese_max || 0));
        if (ese_min === null || ese_min === 0 || ese_min < calculatedEseMin) {
          if (hist && hist.ese_min !== null && hist.ese_min !== undefined && hist.ese_min >= calculatedEseMin) {
            ese_min = hist.ese_min;
          } else if (prof?.eseMin > 0 && prof.eseMin >= calculatedEseMin) {
            ese_min = prof.eseMin;
          } else {
            ese_min = calculatedEseMin;
          }
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
          // NOTE: Do NOT overwrite ese marks here — let missing-component detection decide
          // whether marks are actually present. Only override if marks are blank (handled below).
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

        // is_held = marks are actually missing (regardless of heldback status)
        // heldback students with ALL marks present are NOT treated as held — they get pass/fail calculated
        const is_missing = missingComponents.length > 0;
        const is_held = is_missing; // heldback alone does NOT make is_held true if all marks exist

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
          const prNum = parseNumber(ese_pr_obtained) || 0;
          const thNum = parseNumber(ese_th_obtained) || 0;
          if (parseNumber(ese_th_obtained) !== null || parseNumber(ese_pr_obtained) !== null) {
            ese_obtained = prNum + thNum;
          } else {
            const rawOverall = parseNumber(getCell(row, currentHeaderMap, "ESE Overall", "ESEOverall", "ESE Overall Marks ", "ESE Overall Marks", "ESE Total", "ESE"));
            ese_obtained = rawOverall !== null ? rawOverall : (prNum + thNum);
          }
        }

        // Component marks finalized


        const ce_pr_obtained = missingComponents.includes("CE-PR") ? "Missing" : ce_pr_raw;
        const ce_th_obtained = missingComponents.includes("CE-TH") ? "Missing" : ce_th_raw;

        let ce_max = parseNumber(getCell(row, currentHeaderMap, "CE - Max", "CEMax", "CE Max", "CE_Max", "CE-Max", "Internal Max", "CA Max", "Continuous Evaluation Max"));
        if (ce_max === null) {
          if (hist && hist.ce_max !== null && hist.ce_max !== undefined) {
            ce_max = hist.ce_max;
          } else if (prof?.ceMax > 0) {
            ce_max = prof.ceMax;
          } else {
            ce_max = (parseNumber(ce_pr_max) || 0) + (parseNumber(ce_th_max) || 0);
          }
        }
        let ce_min = parseNumber(getCell(row, currentHeaderMap, "CE - Min", "CEMin", "CE Min", "CE_Min", "CE-Min"));
        if (ce_min === null) {
          ce_min = hist?.ce_min ?? prof?.ceMin ?? 0;
        }
        let ce_obtained;
        if (missingComponents.includes("CE-PR") || missingComponents.includes("CE-TH")) {
          ce_obtained = "Held";
        } else {
          const numPr = parseNumber(ce_pr_obtained) || 0;
          const numTh = parseNumber(ce_th_obtained) || 0;
          if (ce_th_obtained !== "" || ce_pr_obtained !== "") {
            ce_obtained = numPr + numTh;
          } else if (rowCeOverall !== null) {
            ce_obtained = rowCeOverall;
          } else if (hist && hist.ce_obtained !== null && hist.ce_obtained !== "") {
            ce_obtained = hist.ce_obtained;
          } else {
            ce_obtained = parseNumber(getCell(row, currentHeaderMap, "CE Overall Marks ", "CE Overall Marks", "CEOverallMarks", "CEOverall", "CE Total", "CE Marks", "CE Obtained", "CE"));
            if (ce_obtained === null) ce_obtained = numPr + numTh;
          }
        }

        let overall_max = parseNumber(getCell(row, currentHeaderMap, "Overall Maximum", "OverallMaximum", "OverallMax", "Overall Max", "Course Max", "CourseMax", "Total Max"));
        if (overall_max === null || overall_max === 0) {
          if (hist && hist.overall_max) {
            overall_max = hist.overall_max;
          } else if (prof?.courseMax > 0) {
            overall_max = prof.courseMax;
          } else {
            overall_max = (parseNumber(ese_max) || 0) + (parseNumber(ce_max) || 0);
          }
        }
        let overall_min = parseNumber(getCell(row, currentHeaderMap, "Overall Minimum", "OverallMinimum", "OverallMin", "Overall Min", "Course Min", "CourseMin", "Total Min"));
        const calculatedOverallMin = Math.ceil(0.35 * (overall_max || 0));
        if (overall_min === null || overall_min === 0 || overall_min < calculatedOverallMin) {
          if (hist && hist.overall_min && hist.overall_min >= calculatedOverallMin) {
            overall_min = hist.overall_min;
          } else if (prof?.courseMin > 0 && prof.courseMin >= calculatedOverallMin) {
            overall_min = prof.courseMin;
          } else {
            overall_min = calculatedOverallMin;
          }
        }
        let course_overall;
        if (is_held) {
          course_overall = "Held";
        } else {
          if (is_improvement || carriedForwardComponents.length > 0 || repeatedComponents.length > 0) {
            course_overall = (parseNumber(ese_obtained) || 0) + (parseNumber(ce_obtained) || 0);
          } else {
            course_overall = parseNumber(getCell(row, currentHeaderMap, "Course Overall Marks ", "Course Overall Marks", "CourseOverallMarks", "Course Overall", "Course Total"));
            if (course_overall === null) {
              course_overall = (parseNumber(ese_obtained) || 0) + (parseNumber(ce_obtained) || 0);
            }
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
          const requiresEse = (ese_max > 0) || (prof?.requiresEseTh ?? true) || (prof?.requiresEsePr ?? false);
          const eseHasNoMarks = (ese_max === 0 || ese_max === null) && numEse === 0;
          if (requiresEse && numEse === 0) {
            ese_deficit = Math.max(ese_min, calculatedEseMin, 1);
            raw_ese_pass = false;
          } else {
            ese_deficit = eseHasNoMarks ? 999 : Math.max(0, ese_min - numEse);
            raw_ese_pass = ese_deficit === 0 && (!requiresEse || numEse >= calculatedEseMin);
          }

          const numOverall = parseNumber(course_overall) || 0;
          const overallHasNoMarks = (overall_max === 0 || overall_max === null) && numOverall === 0;
          overall_deficit = overallHasNoMarks ? 999 : Math.max(0, overall_min - numOverall);
          raw_overall_pass = overall_deficit === 0 && (numOverall >= overall_min);
          raw_course_pass = raw_ese_pass && raw_overall_pass;
        }

        if (gazetteAssurance && gazetteAssurance.isVerifiedReappear) {
          is_improvement = false;
          attempt_type = "SUPPLEMENTARY";
        }

        baseRecords.push({
          identifiers: { faculty, program, seat, prn, code, name, college, collegeCode, collegeName },
          _college: college,
          _collegeCode: collegeCode,
          _collegeName: collegeName,
          attempt_type,
          is_improvement,
          has_historical_record: !!hist,
          historical_source: hist ? hist.sourceFile : null,
          carried_forward_components: Array.from(new Set(carriedForwardComponents)),
          repeated_components: Array.from(new Set(repeatedComponents)),
          expected_components: prof?.expectedComponents || [],
          expected_component_count: prof?.componentCount || 0,
          available_component_count: (prof?.expectedComponents || []).filter(c => !missingComponents.includes(c)).length,
          is_finalized: missingComponents.length === 0 && !is_heldback && !is_malpractice && !is_absent,
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
            "Attempt Type": attempt_type,
            "Historical Baseline": hist ? `Linked (${hist.sourceFile})` : "Fresh / No Prior Record",
            "Carried Forward Components": Array.from(new Set(carriedForwardComponents)).join(", ") || "None",
            "Components Expected": (prof?.expectedComponents || []).join(", ") || "None",
            "Component Count": `${(prof?.expectedComponents || []).filter(c => !missingComponents.includes(c)).length}/${prof?.componentCount || 0}`,
            "Attempt Status": missingComponents.length === 0 ? "Finalized" : `Held (Missing: ${missingComponents.join(", ")})`,
            "Gazette Reappear Status": gazetteAssurance ? (gazetteAssurance.isVerifiedReappear ? `Verified Reappear (${gazetteAssurance.matchingTerm || 'Gazette'})` : (gazetteAssurance.isPriorPass ? `Prior Pass (${gazetteAssurance.matchingTerm || 'Improvement'})` : `Unlisted in Gazette (${gazetteAssurance.latestStatus})`)) : "N/A",
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
          missing_components: missingComponents,
          gazette_assurance: gazetteAssurance
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
      const rawCat = String(getCell(row, currentHeaderMap, "Exam Category", "ExamCategory", "Appearance Type", "AppearanceType", "Exam Type", "ExamType", "Appearance", "Exam_Category", "Exam_Type") || "").trim().toUpperCase();

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

      const studentId = prn || seat;
      const groupKey = faculty + "|||" + program + "|||" + studentId + "|||" + code;

      if (!groups.has(groupKey)) {
        groups.set(groupKey, {
          identifiers: { faculty, program, seat, prn, code, name, rawCollegeCode, rawCollegeName, rawCat },
          components: {},
          tlm: ""
        });
      } else {
        const grp = groups.get(groupKey);
        if (!grp.identifiers.prn && prn) grp.identifiers.prn = prn;
        if (!grp.identifiers.seat && seat) grp.identifiers.seat = seat;
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

      // Historical baseline lookup & Attempt Type determination
      const hist = getHistoricalEntry(identifiers.prn, identifiers.seat, identifiers.code, currentHistoricalMap);
      const gazetteAssurance = verifyCourseAttempt(identifiers.prn, identifiers.seat, identifiers.code, currentGazetteMap, identifiers.name);
      const rawCat = (identifiers.rawCat || "").toUpperCase();

      const hasImp = rawCat.includes("IMP") || rawCat.includes("IMPROVE");
      const hasSupp = rawCat.includes("SUPP") || rawCat.includes("BACKLOG") || rawCat.includes("REPEATER") || rawCat.includes("RE-APPEAR") || rawCat.includes("REAPPEAR");

      let is_improvement = false;
      let attempt_type = "SUPPLEMENTARY";
      if (gazetteAssurance && gazetteAssurance.isVerifiedReappear) {
        is_improvement = false;
        attempt_type = "SUPPLEMENTARY";
      } else if (hasImp && !hasSupp) {
        is_improvement = true;
        attempt_type = "IMPROVEMENT";
      } else if (hasSupp) {
        is_improvement = false;
        attempt_type = "SUPPLEMENTARY";
      } else if (gazetteAssurance && gazetteAssurance.isPriorPass && !hasSupp) {
        is_improvement = true;
        attempt_type = "IMPROVEMENT";
      } else if (hist && (hist.raw_course_pass === true || (hist.course_overall !== null && hist.course_overall >= 35 && hist.raw_ese_pass))) {
        is_improvement = true;
        attempt_type = "IMPROVEMENT";
      } else {
        is_improvement = false;
        attempt_type = "SUPPLEMENTARY";
      }

      let carriedForwardComponents = [];
      let repeatedComponents = [];
      const missingComponents = [];
      const isBlank = (val) => val === undefined || val === null || String(val).trim() === "";

      // Verify Candidate Components against Confirmed Course Structure
      if (is_improvement) {
        repeatedComponents.push("ESE-TH");
        if (prof?.requiresEseTh) {
          const cTh = components["ESE_TH"];
          if (!cTh || !cTh.present || isBlank(cTh.rawMarksVal)) {
            missingComponents.push("ESE-TH");
          } else {
            if (currentImprovementMode === "best" && hist && hist.ese_th_obtained !== "" && hist.ese_th_obtained !== null) {
              const histVal = parseNumber(hist.ese_th_obtained);
              const currVal = parseNumber(cTh.marks) || 0;
              if (histVal !== null && currVal < histVal) {
                cTh.marks = histVal;
                cTh.rawMarksVal = histVal;
              }
            }
          }
        }
        // ESE-PR strictly carried forward from previous baseline along with CE component
        if (prof?.requiresEsePr || (hist && !isBlank(hist.ese_pr_obtained))) {
          if (hist && !isBlank(hist.ese_pr_obtained)) {
            if (!components["ESE_PR"]) components["ESE_PR"] = { present: true };
            components["ESE_PR"].marks = hist.ese_pr_obtained;
            components["ESE_PR"].rawMarksVal = hist.ese_pr_obtained;
            carriedForwardComponents.push("ESE-PR");
          } else if (hist && hist.ese_obtained !== null && !isBlank(hist.ese_th_obtained) && hist.ese_obtained > parseNumber(hist.ese_th_obtained)) {
            const inferredPr = hist.ese_obtained - parseNumber(hist.ese_th_obtained);
            if (!components["ESE_PR"]) components["ESE_PR"] = { present: true };
            components["ESE_PR"].marks = inferredPr;
            components["ESE_PR"].rawMarksVal = inferredPr;
            carriedForwardComponents.push("ESE-PR");
          } else {
            const cPr = components["ESE_PR"];
            if (cPr && cPr.present && !isBlank(cPr.rawMarksVal)) {
              // present in current attempt
            } else {
              missingComponents.push("ESE-PR");
            }
          }
        }
        // CE-TH: check current attempt first, or carry forward from hist
        if (prof?.requiresCeTh) {
          const cTh = components["CE_TH"];
          if (cTh && cTh.present && !isBlank(cTh.rawMarksVal)) {
            // Present in current attempt
          } else if (hist && !isBlank(hist.ce_th_obtained)) {
            if (!components["CE_TH"]) components["CE_TH"] = { present: true };
            components["CE_TH"].marks = hist.ce_th_obtained;
            components["CE_TH"].rawMarksVal = hist.ce_th_obtained;
            carriedForwardComponents.push("CE-TH");
          } else if (hist && !isBlank(hist.ce_obtained) && !prof?.requiresCePr) {
            if (!components["CE_TH"]) components["CE_TH"] = { present: true };
            components["CE_TH"].marks = hist.ce_obtained;
            components["CE_TH"].rawMarksVal = hist.ce_obtained;
            carriedForwardComponents.push("CE-TH");
          } else {
            missingComponents.push("CE-TH");
          }
        }
        // CE-PR: check current attempt first, or carry forward from hist
        if (prof?.requiresCePr) {
          const cPr = components["CE_PR"];
          if (cPr && cPr.present && !isBlank(cPr.rawMarksVal)) {
            // Present in current attempt
          } else if (hist && !isBlank(hist.ce_pr_obtained)) {
            if (!components["CE_PR"]) components["CE_PR"] = { present: true };
            components["CE_PR"].marks = hist.ce_pr_obtained;
            components["CE_PR"].rawMarksVal = hist.ce_pr_obtained;
            carriedForwardComponents.push("CE-PR");
          } else {
            missingComponents.push("CE-PR");
          }
        }
      } else {
        // Supplementary candidate: BOTH ESE-TH and ESE-PR are taken from current attempt!
        if (prof?.requiresEseTh) {
          const cTh = components["ESE_TH"];
          if (!cTh || !cTh.present || isBlank(cTh.rawMarksVal)) {
            missingComponents.push("ESE-TH");
          } else {
            repeatedComponents.push("ESE-TH");
          }
        }
        if (prof?.requiresEsePr) {
          const cPr = components["ESE_PR"];
          if (!cPr || !cPr.present || isBlank(cPr.rawMarksVal)) {
            missingComponents.push("ESE-PR");
          } else {
            repeatedComponents.push("ESE-PR");
          }
        }
        // CE-TH: check current attempt first, or carry forward from hist
        if (prof?.requiresCeTh) {
          const cTh = components["CE_TH"];
          if (cTh && cTh.present && !isBlank(cTh.rawMarksVal)) {
            // Present in current attempt
          } else if (hist && !isBlank(hist.ce_th_obtained)) {
            if (!components["CE_TH"]) components["CE_TH"] = { present: true };
            components["CE_TH"].marks = hist.ce_th_obtained;
            components["CE_TH"].rawMarksVal = hist.ce_th_obtained;
            carriedForwardComponents.push("CE-TH");
          } else if (hist && !isBlank(hist.ce_obtained) && !prof?.requiresCePr) {
            if (!components["CE_TH"]) components["CE_TH"] = { present: true };
            components["CE_TH"].marks = hist.ce_obtained;
            components["CE_TH"].rawMarksVal = hist.ce_obtained;
            carriedForwardComponents.push("CE-TH");
          } else {
            missingComponents.push("CE-TH");
          }
        }
        // CE-PR: check current attempt first, or carry forward from hist
        if (prof?.requiresCePr) {
          const cPr = components["CE_PR"];
          if (cPr && cPr.present && !isBlank(cPr.rawMarksVal)) {
            // Present in current attempt
          } else if (hist && !isBlank(hist.ce_pr_obtained)) {
            if (!components["CE_PR"]) components["CE_PR"] = { present: true };
            components["CE_PR"].marks = hist.ce_pr_obtained;
            components["CE_PR"].rawMarksVal = hist.ce_pr_obtained;
            carriedForwardComponents.push("CE-PR");
          } else {
            missingComponents.push("CE-PR");
          }
        }
      }

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
      // is_held = marks are actually missing (heldback with all marks present is NOT held)
      const is_missing = missingComponents.length > 0;
      const is_held = is_missing;

      const ese_pr = components["ESE_PR"];
      const ese_pr_max = (ese_pr && ese_pr.max !== null && ese_pr.max > 0) ? ese_pr.max : (hist?.ese_pr_max ?? (prof?.requiresEsePr ? (prof?.maxMarks?.ESE_PR || prof?.esePrMax || "") : ""));
      const ese_pr_min = hist?.ese_pr_min ?? (prof?.esePrMin ?? (ese_pr_max !== "" ? 0 : ""));
      let ese_pr_obtained = (ese_pr && ese_pr.marks !== null) ? ese_pr.marks : (missingComponents.includes("ESE-PR") ? "Missing" : (prof?.requiresEsePr ? 0 : ""));

      const ese_th = components["ESE_TH"];
      const ese_th_max = (ese_th && ese_th.max !== null && ese_th.max > 0) ? ese_th.max : (hist?.ese_th_max ?? (prof?.requiresEseTh ? (prof?.maxMarks?.ESE_TH || prof?.eseThMax || "") : ""));
      const ese_th_min = hist?.ese_th_min ?? (prof?.eseThMin ?? (ese_th_max !== "" ? 0 : ""));
      let ese_th_obtained = (ese_th && ese_th.marks !== null) ? ese_th.marks : (missingComponents.includes("ESE-TH") ? "Missing" : (prof?.requiresEseTh ? 0 : ""));

      if (is_malpractice && malpracticeEntry) {
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

      const ese_max = Math.max((parseNumber(ese_pr_max) || 0) + (parseNumber(ese_th_max) || 0), hist?.ese_max || 0, prof?.eseMax || 0);
      const calculatedEseMin = Math.ceil(0.30 * (ese_max || 0));
      let ese_min = hist?.ese_min ?? (prof?.eseMin ?? calculatedEseMin);
      if (ese_min === null || ese_min === 0 || ese_min < calculatedEseMin) {
        if (hist && hist.ese_min && hist.ese_min >= calculatedEseMin) {
          ese_min = hist.ese_min;
        } else if (prof?.eseMin > 0 && prof.eseMin >= calculatedEseMin) {
          ese_min = prof.eseMin;
        } else {
          ese_min = calculatedEseMin;
        }
      }

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
      const ce_pr_max = (ce_pr && ce_pr.max !== null && ce_pr.max > 0) ? ce_pr.max : (hist?.ce_pr_max ?? (prof?.requiresCePr ? (prof?.maxMarks?.CE_PR || prof?.cePrMax || "") : ""));
      const ce_pr_min = hist?.ce_pr_min ?? (prof?.cePrMin ?? (ce_pr_max !== "" ? 0 : ""));
      const ce_pr_obtained = (ce_pr && ce_pr.marks !== null) ? ce_pr.marks : (missingComponents.includes("CE-PR") ? "Missing" : (prof?.requiresCePr ? 0 : ""));

      const ce_th = components["CE_TH"];
      const ce_th_max = (ce_th && ce_th.max !== null && ce_th.max > 0) ? ce_th.max : (hist?.ce_th_max ?? (prof?.requiresCeTh ? (prof?.maxMarks?.CE_TH || prof?.ceThMax || (!prof?.requiresCePr ? (prof?.ceMax || hist?.ce_max) : null) || "") : ""));
      const ce_th_min = hist?.ce_th_min ?? (prof?.ceThMin ?? (ce_th_max !== "" ? 0 : ""));
      const ce_th_obtained = (ce_th && ce_th.marks !== null) ? ce_th.marks : (missingComponents.includes("CE-TH") ? "Missing" : (prof?.requiresCeTh ? 0 : ""));

      const ce_max = Math.max((parseNumber(ce_pr_max) || 0) + (parseNumber(ce_th_max) || 0), hist?.ce_max || 0, prof?.ceMax || 0);
      const ce_min = hist?.ce_min ?? (prof?.ceMin ?? 0);
      let ce_obtained;
      if (missingComponents.includes("CE-PR") || missingComponents.includes("CE-TH")) {
        ce_obtained = "Held";
      } else {
        ce_obtained = (parseNumber(ce_pr_obtained) || 0) + (parseNumber(ce_th_obtained) || 0);
      }

      const overall_max = Math.max(ese_max + ce_max, hist?.overall_max || 0, prof?.courseMax || 0);
      const calculatedOverallMin = Math.ceil(0.35 * (overall_max || 0));
      let overall_min = hist?.overall_min ?? (prof?.courseMin ?? calculatedOverallMin);
      if (overall_min === null || overall_min === 0 || overall_min < calculatedOverallMin) {
        if (hist && hist.overall_min && hist.overall_min >= calculatedOverallMin) {
          overall_min = hist.overall_min;
        } else if (prof?.courseMin > 0 && prof.courseMin >= calculatedOverallMin) {
          overall_min = prof.courseMin;
        } else {
          overall_min = calculatedOverallMin;
        }
      }
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
        const requiresEse = (ese_max > 0) || (prof?.requiresEseTh ?? true) || (prof?.requiresEsePr ?? false);
        const eseHasNoMarks = (ese_max === 0 || ese_max === null) && numEse === 0;
        if (requiresEse && numEse === 0) {
          ese_deficit = Math.max(ese_min, calculatedEseMin, 1);
          raw_ese_pass = false;
        } else {
          ese_deficit = eseHasNoMarks ? 999 : Math.max(0, ese_min - numEse);
          raw_ese_pass = ese_deficit === 0 && (!requiresEse || numEse >= calculatedEseMin);
        }

        const numOverall = parseNumber(course_overall) || 0;
        const overallHasNoMarks = (overall_max === 0 || overall_max === null) && numOverall === 0;
        overall_deficit = overallHasNoMarks ? 999 : Math.max(0, overall_min - numOverall);
        raw_overall_pass = overall_deficit === 0 && (numOverall >= overall_min);
        raw_course_pass = raw_ese_pass && raw_overall_pass;
      }

      const has_ese_pr = (parseNumber(ese_pr_max) || 0) > 0 || (prof?.requiresEsePr ?? false);
      const has_ese_th = (parseNumber(ese_th_max) || 0) > 0 || (prof?.requiresEseTh ?? false) || !has_ese_pr;
      const is_pr_only = has_ese_pr && !has_ese_th;

      const rawC = identifiers.rawCollegeCode || heldbackEntry?.collegeCode || malpracticeEntry?.collegeCode || absentEntry?.collegeCode || "";
      const rawN = identifiers.rawCollegeName || heldbackEntry?.collegeName || malpracticeEntry?.collegeName || absentEntry?.collegeName || "";
      const { collegeCode, collegeName, college } = collegeRegistry.resolve(identifiers.seat, identifiers.prn, rawC, rawN);

      if (gazetteAssurance && gazetteAssurance.isVerifiedReappear) {
        is_improvement = false;
        attempt_type = "SUPPLEMENTARY";
      }

      baseRecords.push({
        identifiers: { ...identifiers, college, collegeCode, collegeName },
        _college: college,
        _collegeCode: collegeCode,
        _collegeName: collegeName,
        attempt_type,
        is_improvement,
        has_historical_record: !!hist,
        historical_source: hist ? hist.sourceFile : null,
        carried_forward_components: Array.from(new Set(carriedForwardComponents)),
        repeated_components: Array.from(new Set(repeatedComponents)),
        expected_components: prof?.expectedComponents || [],
        expected_component_count: prof?.componentCount || 0,
        available_component_count: (prof?.expectedComponents || []).filter(c => !missingComponents.includes(c)).length,
        is_finalized: missingComponents.length === 0 && !is_heldback && !is_malpractice && !is_absent,
        raw: {
          "Faculty": identifiers.faculty,
          "Program Term Name": identifiers.program,
          "Course Code": identifiers.code,
          "Course Name": identifiers.name,
          "Seat Number": identifiers.seat,
          "PRN": identifiers.prn,
          "ESE - PR Max": ese_pr_max !== 0 && ese_pr_max !== "" ? ese_pr_max : "",
          "ESE - PR Min": ese_pr_min !== "" && ese_pr_min !== null ? ese_pr_min : (ese_pr_max !== "" && ese_pr_max > 0 ? 0 : ""),
          "ESE - PR Obtained": (ese_pr && ese_pr_obtained === "Malpractice (MP)") ? "Malpractice (MP)" : ((ese_pr && ese_pr_obtained === "Absent (Ab)") ? "Absent (Ab)" : ese_pr_obtained),
          "ESE - TH Max": ese_th_max !== 0 && ese_th_max !== "" ? ese_th_max : "",
          "ESE - TH Min": ese_th_min !== "" && ese_th_min !== null ? ese_th_min : (ese_th_max !== "" && ese_th_max > 0 ? 0 : ""),
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
          "Attempt Type": attempt_type,
          "Historical Baseline": hist ? `Linked (${hist.sourceFile})` : "Fresh / No Prior Record",
          "Carried Forward Components": Array.from(new Set(carriedForwardComponents)).join(", ") || "None",
          "Components Expected": (prof?.expectedComponents || []).join(", ") || "None",
          "Component Count": `${(prof?.expectedComponents || []).filter(c => !missingComponents.includes(c)).length}/${prof?.componentCount || 0}`,
          "Attempt Status": missingComponents.length === 0 ? "Finalized" : `Held (Missing: ${missingComponents.join(", ")})`,
          "Gazette Reappear Status": gazetteAssurance ? (gazetteAssurance.isVerifiedReappear ? `Verified Reappear (${gazetteAssurance.matchingTerm || 'Gazette'})` : (gazetteAssurance.isPriorPass ? `Prior Pass (${gazetteAssurance.matchingTerm || 'Improvement'})` : `Unlisted in Gazette (${gazetteAssurance.latestStatus})`)) : "N/A",
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
        missing_components: missingComponents,
        gazette_assurance: gazetteAssurance
      });
    });

    return baseRecords;
  };

  // Dynamically compute groupedRecords from active marks and all supporting datasets
  const groupedRecords = useMemo(() => {
    if (!rawRows || rawRows.length === 0 || !headerMap) return [];
    return buildGroupedRecordsFromRows(
      rawRows,
      headerMap,
      absentRecordsMap,
      malpracticeRecordsMap,
      heldbackRecordsMap,
      historicalRecordsMap,
      improvementScoringMode,
      historicalCourseProfilesMap,
      historicalGazetteMap
    );
  }, [
    rawRows,
    headerMap,
    absentRecordsMap,
    malpracticeRecordsMap,
    heldbackRecordsMap,
    historicalRecordsMap,
    improvementScoringMode,
    historicalCourseProfilesMap,
    historicalGazetteMap
  ]);

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
      if (rec.is_heldback && rec.is_held) {
        // Heldback + marks missing → fully excluded (cannot calculate)
        item.heldbackCount = (item.heldbackCount || 0) + 1;
        item.heldCount = (item.heldCount || 0) + 1;
      } else if (rec.is_heldback) {
        // Heldback + marks present → count heldback but also participate in pass/fail stats
        item.heldbackCount = (item.heldbackCount || 0) + 1;
        if (rec.raw_course_pass) {
          item.rawPassed++;
        } else {
          item.rawFailed++;
          const maxDeficit = Math.max(rec.ese_deficit, rec.overall_deficit);
          if (maxDeficit > 0 && maxDeficit <= 5) {
            item.nearPassCount++;
          }
        }
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
      row._prn = rec.identifiers?.prn || row["PRN"] || "";
      row._seat = rec.identifiers?.seat || row["Seat Number"] || "";
      row._studentName = rec.identifiers?.name || row["Student Name"] || "";
      row._college = rec._college || rec.identifiers?.college || "";
      row._collegeCode = rec._collegeCode || rec.identifiers?.collegeCode || "";
      row._collegeName = rec._collegeName || rec.identifiers?.collegeName || "";
      row._attemptType = rec.attempt_type || "SUPPLEMENTARY";
      row._isImprovement = rec.is_improvement || false;
      row._hasHistoricalRecord = rec.has_historical_record || false;
      row._historicalSource = rec.historical_source || "";
      row._carriedForwardComponents = rec.carried_forward_components || [];
      row._repeatedComponents = rec.repeated_components || [];
      row._gazetteAssurance = rec.gazette_assurance || null;
      const normCode = normalizeKey(rec.identifiers.code);
      const modLimit = courseModerationMap[normCode] || 0;

      // Heldback Student Handling:
      // - If heldback AND marks are missing (CE-TH, CE-PR, ESE-TH, ESE-PR) → Result is Fail (Held - Missing, Heldback)
      // - If heldback BUT all marks are present → Calculate pass/fail normally, but suffix result with "(Heldback)"
      if (rec.is_heldback && rec.is_held) {
        const missingComps = rec.missing_components || [];
        const missingCompsStr = missingComps.length > 0 ? missingComps.join(", ") : "Marks Incomplete";
        const hasEseMissing = missingComps.some(c => c.startsWith("ESE"));
        row["ESE Pass"] = hasEseMissing ? "Held (Missing)" : (rec.raw_ese_pass ? "Pass" : "Fail");
        row["Overall pass"] = "Held (Missing)";
        row["Course Pass/Fail"] = `Fail (Held - Missing: ${missingCompsStr}, Heldback)`;
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
        row._isMissingComp = true;
        row._missingComponents = missingComps;
        row._heldReason = `Missing Component: ${missingCompsStr}`;
        row._modLimit = modLimit;
        row._isEligibleForMod = false;
        row._isPrOnly = rec.is_pr_only;
        row._eseDeficit = 999;
        row._overallDeficit = 999;
        return row;
      }
      // If heldback but marks are fully present, fall through to normal calculation below
      // (the result suffix "(Heldback)" is applied at the end)


      // Held / Missing Component Handling (CE-TH, CE-PR, ESE-TH, ESE-PR):
      // Considered Failed because marks are incomplete, but held status and missing components are displayed
      if (rec.is_held) {
        const missingComps = rec.missing_components || [];
        const missingCompsStr = missingComps.length > 0 ? missingComps.join(", ") : "Marks Incomplete";
        const hasEseMissing = missingComps.some(c => c.startsWith("ESE"));
        row["ESE Pass"] = hasEseMissing ? "Held (Missing)" : (rec.raw_ese_pass ? "Pass" : "Fail");
        row["Overall pass"] = "Held (Missing)";
        row["Course Pass/Fail"] = `Fail (Held - Missing: ${missingCompsStr})`;
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
        row._missingComponents = missingComps;
        row._heldReason = `Missing Component: ${missingCompsStr}`;
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
      // Students with 0 marks in ESE are strictly ineligible for moderation
      const numRawEse = parseNumber(rec.raw["ESE Overall"]) ?? ((parseNumber(rec.raw["ESE - TH Obtained"]) || 0) + (parseNumber(rec.raw["ESE - PR Obtained"]) || 0));
      const isEligibleForModeration = (rec.has_ese_th || (rec.is_pr_only && allowPrOnlyModeration)) && numRawEse > 0;

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
      // Suffix (Heldback) to show the administrative hold even though marks are calculable
      row["Course Pass/Fail"] = rec.is_heldback ? `${final_course_pass} (Heldback)` : final_course_pass;
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
      row._isHeldback = rec.is_heldback; // preserve heldback flag for UI colouring
      row._heldbackReason = rec.is_heldback ? (rec.heldback_reason || "Heldback at term-level") : "";
      row._isMissingComp = false;
      row._missingComponents = [];
      row._modLimit = modLimit;
      row._isEligibleForMod = isEligibleForModeration;
      row._isPrOnly = rec.is_pr_only;
      row._eseDeficit = rec.ese_deficit;
      row._overallDeficit = rec.overall_deficit;

      // For heldback-with-marks students: store both raw and moderated results explicitly
      if (rec.is_heldback) {
        row._heldbackRawPass = rec.raw_course_pass;         // Pass/Fail without any moderation
        row._heldbackModPass = (final_course_pass === "Pass"); // Pass/Fail with moderation applied
        row._heldbackModMarks = moderation_awarded;          // Moderation marks used (0 if raw pass)
      }

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

      // Heldback + missing marks → strictly excluded (cannot calculate)
      if (rec.is_heldback && rec.is_held) {
        item.heldbackCount = (item.heldbackCount || 0) + 1;
        item.heldCount = (item.heldCount || 0) + 1;
        return;
      }

      // Heldback + marks present → count as heldback but still participate in pass simulation
      if (rec.is_heldback) {
        item.heldbackCount = (item.heldbackCount || 0) + 1;
        // Fall through to moderation loop below — result shown as Pass/Fail (Heldback) in detail view
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
    // Pre-group processed rows by normalized course code (O(n)) to avoid O(n*m) inner filter
    const rowsByCourse = new Map();
    for (const r of processedRows) {
      const norm = normalizeKey(r["Course Code"]);
      if (!norm) continue;
      let arr = rowsByCourse.get(norm);
      if (!arr) { arr = []; rowsByCourse.set(norm, arr); }
      arr.push(r);
    }

    return distinctCourses.map(c => {
      const currentMod = courseModerationMap[c.normCode] || 0;
      const courseRows = rowsByCourse.get(c.normCode) || [];
      let rescuedInCourse = 0;
      for (const r of courseRows) { if (r._isModeratedPass) rescuedInCourse++; }
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
      const prn = String(row._prn || row["PRN"] || "").trim();
      const seat = String(row._seat || row["Seat Number"] || "").trim();
      const studentName = String(row._studentName || row["Student Name"] || row["Candidate Name"] || row["Name"] || "").trim();
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
          studentName,
          faculty,
          program,
          college,
          collegeCode,
          collegeName,
          totalCourses: 0,
          improvementCourses: 0,
          supplementaryCourses: 0,
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

      const isVerifiedReappear = !!(row._gazetteAssurance && row._gazetteAssurance.isVerifiedReappear);
      const isImprovement = isVerifiedReappear ? false : !!(row._isImprovement || row._attemptType === "IMPROVEMENT");
      if (isImprovement) {
        st.improvementCourses = (st.improvementCourses || 0) + 1;
      } else {
        st.supplementaryCourses = (st.supplementaryCourses || 0) + 1;
      }

      const isRawCoursePass = !!row._rawPass;
      const isFinalCoursePass = row["Course Pass/Fail"] === "Pass" || (row._isHeldback && String(row["Course Pass/Fail"]).startsWith("Pass"));
      const modMarks = row["Moderation Marks"] || 0;

      if (row._isHeldback) {
        st.heldbackCourses = (st.heldbackCourses || 0) + 1;
        if (!st.heldbackReason && row._heldbackReason) {
          st.heldbackReason = row._heldbackReason;
        }
      }
      if (row._isHeld) {
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
        attemptType: isVerifiedReappear ? "SUPPLEMENTARY" : (row._attemptType || (isImprovement ? "IMPROVEMENT" : "SUPPLEMENTARY")),
        isImprovement: isImprovement,
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
        heldbackRawPass: row._heldbackRawPass,
        heldbackModPass: row._heldbackModPass,
        heldbackModMarks: row._heldbackModMarks,
        missingComponents: row._missingComponents || [],
        modMarks,
        eseOverall: row["ESE Overall"],
        eseMin: row["ESE - Min"],
        courseOverall: row["Course Overall Marks "],
        overallMin: row["Overall Minimum"],
        eseThObtained: row["ESE - TH Obtained"],
        eseThMax: row["ESE - TH Max"],
        eseThMin: row["ESE - TH Min"],
        ceThObtained: row["CE - TH Obtained"],
        ceThMax: row["CE - TH Max"],
        ceThMin: row["CE - TH Min"],
        esePrObtained: row["ESE - PR Obtained"],
        esePrMax: row["ESE - PR Max"],
        esePrMin: row["ESE - PR Min"],
        cePrObtained: row["CE - PR Obtained"],
        cePrMax: row["CE - PR Max"],
        cePrMin: row["CE - PR Min"],
        ceOverall: row["CE Overall Marks "],
        ceMax: row["CE - Max"],
        ceMin: row["CE - Min"],
        carriedForwardComponents: row._carriedForwardComponents || [],
        repeatedComponents: row._repeatedComponents || [],
        gazetteAssurance: row._gazetteAssurance || null
      });
    });

    const list = Array.from(map.values()).map(st => {
      const hasHeldback = (st.heldbackCourses || 0) > 0;
      const hasMissingMarks = (st.heldCourses || 0) > 0;
      let semesterResult = "";
      let rawSemesterResult = "";
      let finalSemesterPass = false;
      let rawSemesterPass = false;
      let pendingGazetteBacklogs = [];
      let hasGazetteAssurance = false;

      // Result Gazette Reappear Assurance:
      // Academic Rule: Semester result should ONLY be collected if the Result Gazette (Reappear Assurance) is uploaded; otherwise keep it blank.
      const isGazetteUploaded = (historicalGazetteMap && historicalGazetteMap.size > 0) || (gazetteReports && gazetteReports.length > 0);

      if (isGazetteUploaded) {
        const cleanP = cleanIdKey(st.prn);
        const cleanS = cleanIdKey(st.seatNumber);
        const cleanN = cleanIdKey(st.studentName);

        // PRN is constant across exam sessions; strictly prioritize PRN lookup!
        let gazetteEntries = cleanP ? (historicalGazetteMap.get(cleanP) || []) : [];

        // Fall back to seat number or name if PRN lookup yielded no entries
        if (gazetteEntries.length === 0 && cleanS) {
          gazetteEntries = historicalGazetteMap.get(cleanS) || [];
        }
        if (gazetteEntries.length === 0 && cleanN && cleanN.length > 3) {
          gazetteEntries = historicalGazetteMap.get(cleanN) || [];
        }

        if (gazetteEntries.length > 0) {
          hasGazetteAssurance = true;
          const matchingEntries = gazetteEntries.filter(entry => {
            const hasMatchingCourse = st.courses.some(c => isCourseReappearInGazetteEntry(c.courseCode, c.courseName, entry));
            if (hasMatchingCourse) return true;
            if (entry.term && st.program) {
              const normEntryTerm = entry.term.toUpperCase().replace(/[^A-Z0-9]/g, "");
              const normProgTerm = st.program.toUpperCase().replace(/[^A-Z0-9]/g, "");
              if (normProgTerm.includes(normEntryTerm) || normEntryTerm.includes(normProgTerm)) return true;
            }
            return false;
          });

          // Use matched entries if found; otherwise all student gazette entries apply
          const relevantEntries = matchingEntries.length > 0 ? matchingEntries : gazetteEntries;

          // Extract regular event moderation from relevant gazette entries or student's gazette records
          const regularModMarks = Math.max(
            relevantEntries.reduce((maxOrd, e) => Math.max(maxOrd, parseOrdMarks(e.ordTotal) || 0), 0),
            gazetteEntries.reduce((maxOrd, e) => Math.max(maxOrd, parseOrdMarks(e.ordTotal) || 0), 0)
          );
          st.regularModerationMarks = regularModMarks;

          const declaredReappearItems = [];
          relevantEntries.forEach(entry => {
            const list = (entry.reappearCodes && entry.reappearCodes.length > 0) ? entry.reappearCodes : (entry.reappearTitles || []);
            list.forEach(item => {
              if (item && !declaredReappearItems.some(existing => cleanCourseCode(existing) === cleanCourseCode(item))) {
                declaredReappearItems.push(item);
              }
            });
          });

          // Current courses that passed
          const currentPassedCourses = st.courses.filter(c => c.coursePass === "Pass" || (c.isHeldback && String(c.coursePass).startsWith("Pass")));

          // Any declared reappear not passed in the current session remains an outstanding backlog
          pendingGazetteBacklogs = declaredReappearItems.filter(item => {
            const isCleared = currentPassedCourses.some(c => {
              const codeMatch = cleanCourseCode(c.courseCode) === cleanCourseCode(item);
              const nameMatch = c.courseName && item && (
                cleanCourseName(c.courseName).includes(cleanCourseName(item)) ||
                cleanCourseName(item).includes(cleanCourseName(c.courseName))
              );
              return codeMatch || nameMatch;
            });
            return !isCleared;
          });
        }
      }

      const hasPendingBacklogs = pendingGazetteBacklogs.length > 0;
      const regularModMarks = st.regularModerationMarks || 0;
      const eventModMarks = st.totalModerationMarks || 0;
      const cumulativeModMarks = regularModMarks + eventModMarks;

      if (!isGazetteUploaded) {
        // Result Gazette NOT uploaded: per academic requirement, semester result is NOT evaluated and must remain blank
        rawSemesterPass = false;
        finalSemesterPass = false;
        semesterResult = "";
        rawSemesterResult = "";
      } else if (hasMissingMarks) {
        // Collect missing component details across all held courses (CE-TH, CE-PR, ESE-TH, ESE-PR)
        const heldCourses = st.courses.filter(c => c.isHeld);
        const allMissingComps = Array.from(new Set(heldCourses.flatMap(c => c.missingComponents || [])));
        const missingCompsSummary = allMissingComps.length > 0 ? allMissingComps.join(", ") : "Marks Incomplete";

        // Missing marks: considered Failed because marks are incomplete, but held status, reason and heldback are shown
        rawSemesterPass = false;
        finalSemesterPass = false;
        if (hasHeldback) {
          semesterResult = `Fail (Held - Missing: ${missingCompsSummary}${hasPendingBacklogs ? `, Pending: ${pendingGazetteBacklogs.join(", ")}` : ""}, Heldback)`;
          rawSemesterResult = `Fail (Held - Missing: ${missingCompsSummary}, Heldback)`;
        } else {
          semesterResult = `Fail (Held - Missing: ${missingCompsSummary}${hasPendingBacklogs ? `, Pending: ${pendingGazetteBacklogs.join(", ")}` : ""})`;
          rawSemesterResult = `Fail (Held - Missing: ${missingCompsSummary})`;
        }
      } else if (hasPendingBacklogs) {
        // If student has pending backlogs in the Gazette, semester CANNOT BE PASS!
        rawSemesterPass = false;
        finalSemesterPass = false;
        if (st.finalFailedCourses === 0) {
          semesterResult = `Fail (Pending Backlog${pendingGazetteBacklogs.length > 1 ? "s" : ""}: ${pendingGazetteBacklogs.join(", ")})`;
          rawSemesterResult = `Fail (Pending Backlogs)`;
        } else {
          semesterResult = `Fail (${st.finalFailedCourses} Failed, Pending: ${pendingGazetteBacklogs.join(", ")})`;
          rawSemesterResult = `Fail (${st.rawFailedCourses} Failed, Pending Backlogs)`;
        }
        if (hasHeldback) {
          semesterResult += " (Heldback)";
          rawSemesterResult += " (Heldback)";
        }
      } else {
        // Gazette IS uploaded and 0 pending backlogs!
        rawSemesterPass = st.rawFailedCourses === 0;
        finalSemesterPass = st.finalFailedCourses === 0;

        if (hasHeldback) {
          // Heldback student with all marks available: calculate pass/fail and suffix (Heldback)
          rawSemesterResult = rawSemesterPass ? "Pass (Heldback)" : "Fail (Heldback)";
          if (!rawSemesterPass && finalSemesterPass) {
            semesterResult = "Pass (Rescued, Heldback)";
          } else if (finalSemesterPass) {
            semesterResult = "Pass (Heldback)";
          } else {
            semesterResult = `Fail (${st.finalFailedCourses} Failed, Heldback)`;
          }
        } else {
          // Standard student
          rawSemesterResult = rawSemesterPass ? "Pass" : "Fail";
          semesterResult = finalSemesterPass ? "Pass" : "Fail";
        }
      }

      const isRescuedSemester = isGazetteUploaded && !hasMissingMarks && !hasPendingBacklogs && !rawSemesterPass && finalSemesterPass;

      const heldCourses = st.courses.filter(c => c.isHeld);
      const allMissingComps = Array.from(new Set(heldCourses.flatMap(c => c.missingComponents || [])));
      const missingDetailStr = heldCourses.map(c => {
        const comps = (c.missingComponents && c.missingComponents.length > 0) ? c.missingComponents.join(", ") : "Incomplete";
        return `${cleanCourseCode(c.courseCode)} [Missing: ${comps}]`;
      }).join("; ");

      return {
        ...st,
        eventModerationMarks: eventModMarks,
        regularModerationMarks: regularModMarks,
        cumulativeModerationMarks: cumulativeModMarks,
        totalModerationMarks: cumulativeModMarks,
        ourModTotal: cumulativeModMarks,
        isHeld: hasMissingMarks,
        isHeldback: hasHeldback,
        rawSemesterPass,
        finalSemesterPass,
        isRescuedSemester,
        semesterResult,
        rawSemesterResult,
        missingComponents: allMissingComps,
        missingDetailStr,
        pendingGazetteBacklogs,
        hasGazetteAssurance,
        isGazetteUploaded
      };
    });

    return list.sort((a, b) => {
      if (a.seatNumber && b.seatNumber) {
        return String(a.seatNumber).localeCompare(String(b.seatNumber), undefined, { numeric: true });
      }
      return String(a.prn).localeCompare(String(b.prn), undefined, { numeric: true });
    });
  }, [processedRows, historicalGazetteMap, gazetteReports]);

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
    const isGazetteUploaded = (historicalGazetteMap && historicalGazetteMap.size > 0) || (gazetteReports && gazetteReports.length > 0);
    const total = studentSemesterData.length;
    if (total === 0) {
      return {
        isGazetteUploaded,
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
        pendingBacklogStudents: 0,
        pendingBacklogPct: "0.0",
        rescuedStudents: 0,
        rescuedPct: "0.0",
        absentStudents: 0,
        absentPct: "0.0",
        malpracticeStudents: 0,
        malpracticePct: "0.0",
        improvementStudents: 0,
        supplementaryStudents: 0,
        totalPapersAttempted: 0,
        totalImprovementPapers: 0,
        totalSupplementaryPapers: 0,
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
    let pendingBacklogStudents = 0;
    let improvementStudents = 0;
    let supplementaryStudents = 0;
    let totalPapersAttempted = 0;
    let totalImprovementPapers = 0;
    let totalSupplementaryPapers = 0;

    studentSemesterData.forEach(st => {
      totalPapersAttempted += st.totalCourses;
      totalImprovementPapers += (st.improvementCourses || 0);
      totalSupplementaryPapers += (st.supplementaryCourses || 0);
      if ((st.improvementCourses || 0) > 0) improvementStudents++;
      if ((st.supplementaryCourses || 0) > 0) supplementaryStudents++;

      if ((st.pendingGazetteBacklogs || []).length > 0) {
        pendingBacklogStudents++;
      }
      if (st.isHeldback) {
        heldbackStudents++;
        if (st.isHeld) {
          heldStudents++;
        } else if (isGazetteUploaded) {
          if (st.rawSemesterPass) rawPassedStudents++;
          if (st.finalSemesterPass) finalPassedStudents++;
          if (st.isRescuedSemester) rescuedStudents++;
        }
      } else if (st.isHeld) {
        heldMissingStudents++;
        heldStudents++;
      } else if (isGazetteUploaded) {
        if (st.rawSemesterPass) rawPassedStudents++;
        if (st.finalSemesterPass) finalPassedStudents++;
        if (st.isRescuedSemester) rescuedStudents++;
      }
      if ((st.absentCourses || 0) > 0) absentStudents++;
      if ((st.malpracticeCourses || 0) > 0) malpracticeStudents++;
    });

    const failedStudents = isGazetteUploaded ? (total - finalPassedStudents - heldStudents) : 0;

    return {
      isGazetteUploaded,
      totalStudents: total,
      rawPassedStudents: isGazetteUploaded ? rawPassedStudents : 0,
      rawPassedPct: total > 0 && isGazetteUploaded ? ((rawPassedStudents / total) * 100).toFixed(1) : "0.0",
      finalPassedStudents: isGazetteUploaded ? finalPassedStudents : 0,
      finalPassedPct: total > 0 && isGazetteUploaded ? ((finalPassedStudents / total) * 100).toFixed(1) : "0.0",
      failedStudents,
      failedPct: total > 0 && isGazetteUploaded ? ((failedStudents / total) * 100).toFixed(1) : "0.0",
      heldStudents,
      heldPct: total > 0 ? ((heldStudents / total) * 100).toFixed(1) : "0.0",
      heldbackStudents,
      heldbackPct: total > 0 ? ((heldbackStudents / total) * 100).toFixed(1) : "0.0",
      heldMissingStudents,
      heldMissingPct: total > 0 ? ((heldMissingStudents / total) * 100).toFixed(1) : "0.0",
      pendingBacklogStudents,
      pendingBacklogPct: total > 0 ? ((pendingBacklogStudents / total) * 100).toFixed(1) : "0.0",
      rescuedStudents: isGazetteUploaded ? rescuedStudents : 0,
      rescuedPct: total > 0 && isGazetteUploaded ? ((rescuedStudents / total) * 100).toFixed(1) : "0.0",
      absentStudents,
      absentPct: total > 0 ? ((absentStudents / total) * 100).toFixed(1) : "0.0",
      malpracticeStudents,
      malpracticePct: total > 0 ? ((malpracticeStudents / total) * 100).toFixed(1) : "0.0",
      improvementStudents,
      supplementaryStudents,
      totalPapersAttempted,
      totalImprovementPapers,
      totalSupplementaryPapers,
      avgPapersPerStudent: total > 0 ? (totalPapersAttempted / total).toFixed(1) : "0.0"
    };
  }, [studentSemesterData, historicalGazetteMap, gazetteReports]);

  // Student Metrics calculated directly from the active scoped filters (for main student view)
  const studentMetrics = useMemo(() => {
    const isGazetteUploaded = (historicalGazetteMap && historicalGazetteMap.size > 0) || (gazetteReports && gazetteReports.length > 0);
    const total = scopedStudents.length;
    if (total === 0) {
      return {
        isGazetteUploaded,
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
        pendingBacklogStudents: 0,
        pendingBacklogPct: "0.0",
        rescuedStudents: 0,
        rescuedPct: "0.0",
        absentStudents: 0,
        absentPct: "0.0",
        malpracticeStudents: 0,
        malpracticePct: "0.0",
        improvementStudents: 0,
        supplementaryStudents: 0,
        totalPapersAttempted: 0,
        totalImprovementPapers: 0,
        totalSupplementaryPapers: 0,
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
    let pendingBacklogStudents = 0;
    let improvementStudents = 0;
    let supplementaryStudents = 0;
    let totalPapersAttempted = 0;
    let totalImprovementPapers = 0;
    let totalSupplementaryPapers = 0;

    scopedStudents.forEach(st => {
      totalPapersAttempted += st.totalCourses;
      totalImprovementPapers += (st.improvementCourses || 0);
      totalSupplementaryPapers += (st.supplementaryCourses || 0);
      if ((st.improvementCourses || 0) > 0) improvementStudents++;
      if ((st.supplementaryCourses || 0) > 0) supplementaryStudents++;

      if ((st.pendingGazetteBacklogs || []).length > 0) {
        pendingBacklogStudents++;
      }
      if (st.isHeldback) {
        heldbackStudents++;
        if (st.isHeld) {
          heldStudents++;
        } else if (isGazetteUploaded) {
          if (st.rawSemesterPass) rawPassedStudents++;
          if (st.finalSemesterPass) finalPassedStudents++;
          if (st.isRescuedSemester) rescuedStudents++;
        }
      } else if (st.isHeld) {
        heldMissingStudents++;
        heldStudents++;
      } else if (isGazetteUploaded) {
        if (st.rawSemesterPass) rawPassedStudents++;
        if (st.finalSemesterPass) finalPassedStudents++;
        if (st.isRescuedSemester) rescuedStudents++;
      }
      if ((st.absentCourses || 0) > 0) absentStudents++;
      if ((st.malpracticeCourses || 0) > 0) malpracticeStudents++;
    });

    const failedStudents = isGazetteUploaded ? (total - finalPassedStudents - heldStudents) : 0;

    return {
      isGazetteUploaded,
      totalStudents: total,
      rawPassedStudents: isGazetteUploaded ? rawPassedStudents : 0,
      rawPassedPct: total > 0 && isGazetteUploaded ? ((rawPassedStudents / total) * 100).toFixed(1) : "0.0",
      finalPassedStudents: isGazetteUploaded ? finalPassedStudents : 0,
      finalPassedPct: total > 0 && isGazetteUploaded ? ((finalPassedStudents / total) * 100).toFixed(1) : "0.0",
      failedStudents,
      failedPct: total > 0 && isGazetteUploaded ? ((failedStudents / total) * 100).toFixed(1) : "0.0",
      heldStudents,
      heldPct: total > 0 ? ((heldStudents / total) * 100).toFixed(1) : "0.0",
      heldbackStudents,
      heldbackPct: total > 0 ? ((heldbackStudents / total) * 100).toFixed(1) : "0.0",
      heldMissingStudents,
      heldMissingPct: total > 0 ? ((heldMissingStudents / total) * 100).toFixed(1) : "0.0",
      pendingBacklogStudents,
      pendingBacklogPct: total > 0 ? ((pendingBacklogStudents / total) * 100).toFixed(1) : "0.0",
      rescuedStudents: isGazetteUploaded ? rescuedStudents : 0,
      rescuedPct: total > 0 && isGazetteUploaded ? ((rescuedStudents / total) * 100).toFixed(1) : "0.0",
      absentStudents,
      absentPct: total > 0 ? ((absentStudents / total) * 100).toFixed(1) : "0.0",
      malpracticeStudents,
      malpracticePct: total > 0 ? ((malpracticeStudents / total) * 100).toFixed(1) : "0.0",
      improvementStudents,
      supplementaryStudents,
      totalPapersAttempted,
      totalImprovementPapers,
      totalSupplementaryPapers,
      avgPapersPerStudent: total > 0 ? (totalPapersAttempted / total).toFixed(1) : "0.0"
    };
  }, [scopedStudents, historicalGazetteMap, gazetteReports]);

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
    } else if (studentFilterStatus === "PENDING_BACKLOG") {
      list = list.filter(st => (st.pendingGazetteBacklogs || []).length > 0);
    } else if (studentFilterStatus === "IMPROVEMENT") {
      list = list.filter(st => (st.improvementCourses || 0) > 0);
    } else if (studentFilterStatus === "SUPPLEMENTARY") {
      list = list.filter(st => (st.supplementaryCourses || 0) > 0);
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

  // Result & Ordinance Reconciliation Comparison Dataset
  const comparisonRecords = useMemo(() => {
    if (rawComparisonRows.length === 0) return [];

    // Map studentSemesterData by normalized PRN and Seat
    const prnMap = new Map();
    const seatMap = new Map();
    studentSemesterData.forEach(st => {
      const normPrn = normalizeKey(st.prn);
      const normSeat = normalizeKey(st.seatNumber);
      if (normPrn) prnMap.set(normPrn, st);
      if (normSeat) seatMap.set(normSeat, st);
    });

    return rawComparisonRows.map(pub => {
      const normPrn = normalizeKey(pub.prn);
      const normSeat = normalizeKey(pub.seat);
      const st = (normPrn && prnMap.get(normPrn)) || (normSeat && seatMap.get(normSeat)) || null;

      let isFound = !!st;
      let calcResult = "Not Loaded";
      let calcModMarks = 0;
      let calcEventModMarks = 0;
      let calcRegularModMarks = 0;
      let calcFailedCourses = [];
      let calcIsHeld = false;
      let calcIsHeldback = false;
      let calcHeldbackHasMarks = false; // true = marks present, result calculable without mod
      let calcHeldbackReason = "";
      let calcIsRescued = false;
      let calcRawPass = false;

      if (st) {
        calcIsHeld = st.isHeld;
        calcIsHeldback = st.isHeldback;
        calcHeldbackReason = st.heldbackReason || "";
        calcIsRescued = st.isRescuedSemester;
        calcRawPass = st.rawSemesterPass;

        if (calcIsHeldback) {
          // Heldback students: result is calculated WITHOUT moderation.
          // Their results are NOT published by the university.
          // Check if ALL marks are present (is_held = false on all courses) → calculable
          const hasAnyHeldCourse = st.courses.some(c => c.isHeld);
          calcHeldbackHasMarks = !hasAnyHeldCourse;
          // No moderation is applied to heldback students — always 0
          calcModMarks = 0;
          calcEventModMarks = 0;
          calcRegularModMarks = 0;
          if (calcHeldbackHasMarks) {
            // All marks present → show raw pass/fail result (no moderation)
            calcResult = st.rawSemesterPass ? "Pass (Heldback)" : "Fail (Heldback)";
          } else {
            // Marks missing → considered Failed as marks are incomplete, but held status and heldback are shown
            calcResult = "Fail (Held - Missing, Heldback)";
          }
          const missingCodes = st.courses.filter(c => c.isHeld).map(c => {
            const comps = (c.missingComponents && c.missingComponents.length > 0) ? c.missingComponents.join(", ") : "Incomplete";
            return `${cleanCourseCode(c.courseCode)} (Missing: ${comps})`;
          });
          const regularFailed = st.courses.filter(c => !c.isHeld && c.rawPass === false).map(c => cleanCourseCode(c.courseCode));
          calcFailedCourses = [...regularFailed, ...missingCodes];
        } else if (st.isHeld) {
          calcModMarks = 0;
          calcEventModMarks = 0;
          calcRegularModMarks = 0;
          calcResult = "Fail (Held - Missing)";
          calcFailedCourses = st.courses.filter(c => c.isHeld).map(c => {
            const comps = (c.missingComponents && c.missingComponents.length > 0) ? c.missingComponents.join(", ") : "Incomplete";
            return `${cleanCourseCode(c.courseCode)} (Missing: ${comps})`;
          });
        } else {
          calcEventModMarks = st.eventModerationMarks || 0;
          calcRegularModMarks = st.regularModerationMarks || 0;
          calcModMarks = st.totalModerationMarks !== undefined ? st.totalModerationMarks : (calcRegularModMarks + calcEventModMarks);
          calcResult = st.finalSemesterPass ? "Pass" : "Fail";
          calcFailedCourses = st.courses.filter(c => c.coursePass !== "Pass" && !c.isHeld && !c.isHeldback).map(c => cleanCourseCode(c.courseCode));
        }

        // Direct Gazette Check: if calcRegularModMarks is still 0, check historicalGazetteMap by PRN or Seat directly
        if (calcRegularModMarks === 0) {
          const cleanP = cleanIdKey(pub.prn) || (st ? cleanIdKey(st.prn) : "");
          const cleanS = cleanIdKey(pub.seat) || (st ? cleanIdKey(st.seatNumber) : "");
          let gEntries = cleanP ? (historicalGazetteMap.get(cleanP) || []) : [];
          if (gEntries.length === 0 && cleanS) gEntries = historicalGazetteMap.get(cleanS) || [];
          if (gEntries.length > 0) {
            calcRegularModMarks = gEntries.reduce((maxOrd, e) => Math.max(maxOrd, parseOrdMarks(e.ordTotal) || 0), 0);
            calcModMarks = calcRegularModMarks + calcEventModMarks;
          }
        }
      } else {
        // Fallback: check if regular moderation exists in historical gazette map
        const cleanP = cleanIdKey(pub.prn);
        const cleanS = cleanIdKey(pub.seat);
        let gEntries = cleanP ? (historicalGazetteMap.get(cleanP) || []) : [];
        if (gEntries.length === 0 && cleanS) gEntries = historicalGazetteMap.get(cleanS) || [];
        if (gEntries.length > 0) {
          calcRegularModMarks = gEntries.reduce((maxOrd, e) => Math.max(maxOrd, parseOrdMarks(e.ordTotal) || 0), 0);
          calcModMarks = calcRegularModMarks;
        }
      }

      // Normalize published result and determine exact academic outcome
      const pubResRaw = String(pub.result || "").trim();
      const pubOtherRaw = String(pub.otherStatus || "").trim();
      const isPubHeldback = /held\s*back|hold\s*back/i.test(pubResRaw) || /held\s*back|hold\s*back/i.test(pubOtherRaw);

      // Determine the published academic outcome (Pass / Fail / Held)
      let pubAcademicResult = "Fail";
      if (/pass/i.test(pubResRaw)) {
        pubAcademicResult = "Pass";
      } else if (/fail/i.test(pubResRaw)) {
        pubAcademicResult = "Fail";
      } else if (pub.failCount > 0 || (pub.reappear && pub.reappear.trim() !== "" && !/^(nil|none|-)$/i.test(pub.reappear.trim()))) {
        pubAcademicResult = "Fail";
      } else if (isPubHeldback) {
        // When university published HELDBACK, deduce the underlying academic status:
        if (st && (st.isHeld || !calcHeldbackHasMarks)) {
          pubAcademicResult = "Fail";
        } else if (st) {
          pubAcademicResult = st.rawSemesterPass ? "Pass" : "Fail";
        } else {
          pubAcademicResult = "Heldback";
        }
      } else if (/held|hold/i.test(pubResRaw)) {
        pubAcademicResult = (st && (st.isHeld || !calcHeldbackHasMarks)) ? "Fail" : "Held";
      } else {
        pubAcademicResult = "Fail";
      }

      // Exact published result string with (Heldback) bracket if applicable
      let pubExactResult = pubAcademicResult;
      if (isPubHeldback) {
        if (/pass/i.test(pubResRaw)) {
          pubExactResult = "Pass (Heldback)";
        } else if (/fail/i.test(pubResRaw) || pub.failCount > 0 || (pub.reappear && pub.reappear.trim() !== "" && !/^(nil|none|-)$/i.test(pub.reappear.trim()))) {
          pubExactResult = "Fail (Heldback)";
        } else if (st && (st.isHeld || !calcHeldbackHasMarks)) {
          pubExactResult = "Fail (Held - Missing, Heldback)";
        } else if (st) {
          pubExactResult = st.rawSemesterPass ? "Pass (Heldback)" : "Fail (Heldback)";
        } else {
          pubExactResult = "Heldback";
        }
      } else if (/held|hold/i.test(pubResRaw)) {
        if (st && (st.isHeld || !calcHeldbackHasMarks)) {
          pubExactResult = "Fail (Held - Missing)";
        } else {
          pubExactResult = "Held";
        }
      }

      // Base published result for comparison
      const pubResultStd = (pubAcademicResult === "Pass" || pubAcademicResult === "Fail") ? pubAcademicResult : (isPubHeldback ? "Heldback" : pubAcademicResult);

      const pubModMarks = pub.ordTotal || 0; // Total Published Ordinance
      const pubHasOrd = pub.ord === "YES" || pubModMarks > 0;
      const pubRegularModMarks = calcRegularModMarks;
      // Event Moderation in university published result: difference of total published ordinance and original event moderation
      const pubEventModMarks = Math.max(0, pubModMarks - pubRegularModMarks);

      // Heldback students: result is not published → comparison is N/A, not a real mismatch
      let isResultMatch = null; // null = N/A
      let isModMatch = null;    // null = N/A
      let isEventModMatch = null;
      let isExactMatch = null;  // null = N/A
      let modDiff = 0;
      let eventModDiff = 0;

      let category = "EXACT_MATCH";
      let discrepancyDescription = "Exact Match";

      if (!isFound) {
        category = "NOT_IN_DATA";
        isResultMatch = false;
        isModMatch = false;
        isEventModMatch = false;
        isExactMatch = false;
        discrepancyDescription = "Student record not found in our loaded dataset";
      } else if (calcIsHeldback) {
        // Heldback students: calculated WITHOUT moderation (calcModMarks already = 0 above).
        // Published result CAN be Pass, Fail, or Heldback — do a real comparison.
        const baseRawResult = calcRawPass ? "Pass" : "Fail"; // raw result without moderation

        if (!calcHeldbackHasMarks) {
          // Marks missing: student is considered Failed due to missing marks, with held status and heldback shown
          // If published result is Fail, Fail (Heldback), or Heldback → this is a MATCH!
          if (pubResultStd === "Fail" || pubResultStd === "Heldback" || pubExactResult.startsWith("Fail")) {
            isResultMatch = true;
            isModMatch = Math.round(pubModMarks) === 0;
            isEventModMatch = Math.round(pubEventModMarks) === 0;
            modDiff = -pubModMarks;
            eventModDiff = -pubEventModMarks;
            isExactMatch = isResultMatch && isModMatch;
            category = isExactMatch ? "EXACT_MATCH" : "MOD_DIFF";
            discrepancyDescription = isExactMatch
              ? `Heldback student — held due to missing marks, considered Fail (Heldback). Matches published ${pubExactResult}. Exact match.`
              : `Heldback student — held due to missing marks, considered Fail (Heldback). Mod diff: Our 0 vs Published ${pubModMarks}.`;
          } else {
            isResultMatch = false;
            isModMatch = Math.round(pubModMarks) === 0;
            isEventModMatch = Math.round(pubEventModMarks) === 0;
            modDiff = -pubModMarks;
            eventModDiff = -pubEventModMarks;
            isExactMatch = false;
            category = "FAILED_HERE_PASSED_PUB";
            discrepancyDescription = `Heldback student — marks missing on our side (Fail), but university published: ${pubExactResult}`;
          }
        } else if (pubResultStd === "Heldback") {
          // University sheet only specified "Heldback" with no explicit marks/reappear info
          isResultMatch = true;
          isModMatch = Math.round(pubModMarks) === 0;
          isEventModMatch = Math.round(pubEventModMarks) === 0;
          modDiff = -pubModMarks;
          eventModDiff = -pubEventModMarks;
          isExactMatch = isResultMatch && isModMatch;
          category = isExactMatch ? "EXACT_MATCH" : "MOD_DIFF";
          discrepancyDescription = isExactMatch
            ? `Heldback confirmed — both systems show Heldback. Raw result (no mod): ${baseRawResult} (Heldback). Exact match.`
            : `Heldback confirmed — both systems show Heldback. Raw result (no mod): ${baseRawResult}. Mod diff: Our 0 vs Published ${pubModMarks}.`;
        } else {
          // Published shows concrete academic outcome ("Pass" or "Fail")
          isResultMatch = baseRawResult === pubResultStd;
          isModMatch = Math.round(pubModMarks) === 0;
          isEventModMatch = Math.round(pubEventModMarks) === 0;
          modDiff = -pubModMarks;
          eventModDiff = -pubEventModMarks;
          isExactMatch = isResultMatch && isModMatch;

          if (!isResultMatch) {
            if (baseRawResult === "Pass" && pubResultStd === "Fail") {
              category = "PASSED_HERE_FAILED_PUB";
              discrepancyDescription = `Heldback student — Our raw result: Pass (Heldback), but university published: ${pubExactResult}`;
            } else if (baseRawResult === "Fail" && pubResultStd === "Pass") {
              category = "FAILED_HERE_PASSED_PUB";
              discrepancyDescription = `Heldback student — Our raw result: Fail (Heldback), but university published: ${pubExactResult}`;
            } else {
              category = "HELD_MISMATCH";
              discrepancyDescription = `Heldback — result discrepancy: Our raw=${baseRawResult} (Heldback) vs Published=${pubExactResult}`;
            }
          } else if (!isModMatch) {
            category = "MOD_DIFF";
            discrepancyDescription = `Heldback student — result matches (${baseRawResult}), but mod diff: Our 0 vs Published ${pubModMarks}`;
          } else {
            category = "EXACT_MATCH";
            discrepancyDescription = `Heldback student — exact match: ${baseRawResult} (Heldback) vs Published ${pubExactResult}`;
          }
        }
      } else {
        // Normal comparison for non-heldback students
        const baseCalcStd = calcResult.startsWith("Pass") ? "Pass" : (calcResult === "Held" || calcResult === "Heldback") ? calcResult : "Fail";
        const basePubStd = pubResultStd.startsWith("Pass") ? "Pass" : (pubResultStd === "Held" || pubResultStd === "Heldback") ? pubResultStd : "Fail";
        isResultMatch = baseCalcStd === basePubStd;
        modDiff = calcModMarks - pubModMarks;
        eventModDiff = calcEventModMarks - pubEventModMarks;
        isModMatch = Math.round(calcModMarks) === Math.round(pubModMarks);
        isEventModMatch = Math.round(calcEventModMarks) === Math.round(pubEventModMarks);
        isExactMatch = isResultMatch && isModMatch;

        if (!isResultMatch) {
          if (calcResult === "Pass" && pubResultStd === "Fail") {
            category = "PASSED_HERE_FAILED_PUB";
            discrepancyDescription = "Passed in Our System, but Marked Fail in University Published Data";
          } else if (calcResult === "Fail" && pubResultStd === "Pass") {
            category = "FAILED_HERE_PASSED_PUB";
            discrepancyDescription = "Failed in Our System, but Marked Pass in University Published Data";
          } else {
            category = "HELD_MISMATCH";
            discrepancyDescription = `Result status discrepancy: Our=${calcResult} vs Published=${pubResultStd}`;
          }
        } else if (!isModMatch) {
          category = "MOD_DIFF";
          discrepancyDescription = `Moderation difference: Our Mod Total=${calcModMarks.toFixed(2)} (Event: ${calcEventModMarks.toFixed(2)}, Reg: ${calcRegularModMarks.toFixed(2)}) vs Published Ord Total=${pubModMarks.toFixed(2)} (Pub Event: ${pubEventModMarks.toFixed(2)}, Reg: ${pubRegularModMarks.toFixed(2)}), Diff: ${modDiff > 0 ? "+" : ""}${modDiff.toFixed(2)}`;
        }
      }

      return {
        ...pub,
        matchedStudent: st,
        isFound,
        calcResult,
        calcModMarks,
        calcRegularModMarks,
        calcEventModMarks,
        pubRegularModMarks,
        pubEventModMarks,
        eventModDiff,
        isEventModMatch,
        calcFailedCourses,
        calcIsHeld,
        calcIsHeldback,
        calcHeldbackHasMarks,
        calcHeldbackReason,
        calcIsRescued,
        calcRawPass,
        pubResultStd,
        pubExactResult,
        isPubHeldback,
        pubModMarks,
        pubHasOrd,
        isResultMatch,
        isModMatch,
        isExactMatch,
        modDiff,
        category,
        discrepancyDescription
      };
    });
  }, [rawComparisonRows, studentSemesterData, historicalGazetteMap]);

  // Reconciliation KPIs
  const comparisonKPIs = useMemo(() => {
    const total = comparisonRecords.length;
    if (total === 0) {
      return { total: 0, matchedStudents: 0, matchRate: "0.0", resultMatches: 0, resultMatchRate: "0.0", resultMismatches: 0, modMatches: 0, modMatchRate: "0.0", modDiffs: 0, eventModMatches: 0, eventModMatchRate: "0.0", eventModDiffs: 0, exactMatches: 0, exactMatchRate: "0.0", passedHereFailedPub: 0, failedHerePassedPub: 0, heldMismatches: 0, notInDataset: 0, heldbackCount: 0, totalCalcMod: 0, totalPubMod: 0, totalCalcEventMod: 0, totalPubEventMod: 0, totalRegMod: 0, modTotalDiff: 0, eventModTotalDiff: 0 };
    }

    let matchedStudents = 0;
    let resultMatches = 0;
    let resultMismatches = 0;
    let modMatches = 0;
    let modDiffs = 0;
    let eventModMatches = 0;
    let eventModDiffs = 0;
    let exactMatches = 0;
    let passedHereFailedPub = 0;
    let failedHerePassedPub = 0;
    let heldMismatches = 0;
    let notInDataset = 0;
    let heldbackCount = 0; // informational: how many heldback students are in the dataset
    let totalCalcMod = 0;
    let totalPubMod = 0;
    let totalCalcEventMod = 0;
    let totalPubEventMod = 0;
    let totalRegMod = 0;

    comparisonRecords.forEach(r => {
      if (!r.isFound) { notInDataset++; return; }
      matchedStudents++;
      if (r.calcIsHeldback) heldbackCount++;
      totalCalcMod += r.calcModMarks;
      totalPubMod += r.pubModMarks;
      totalCalcEventMod += (r.calcEventModMarks || 0);
      totalPubEventMod += (r.pubEventModMarks || 0);
      totalRegMod += (r.calcRegularModMarks || 0);

      if (r.isResultMatch) resultMatches++;
      else {
        resultMismatches++;
        if (r.category === "PASSED_HERE_FAILED_PUB") passedHereFailedPub++;
        else if (r.category === "FAILED_HERE_PASSED_PUB") failedHerePassedPub++;
        else if (r.category === "HELD_MISMATCH") heldMismatches++;
      }
      if (r.isModMatch) modMatches++;
      else modDiffs++;

      if (r.isEventModMatch) eventModMatches++;
      else eventModDiffs++;

      if (r.isExactMatch) exactMatches++;
    });

    return {
      total,
      matchedStudents,
      matchRate: total > 0 ? ((matchedStudents / total) * 100).toFixed(1) : "0.0",
      resultMatches,
      resultMatchRate: matchedStudents > 0 ? ((resultMatches / matchedStudents) * 100).toFixed(1) : "0.0",
      resultMismatches,
      modMatches,
      modMatchRate: matchedStudents > 0 ? ((modMatches / matchedStudents) * 100).toFixed(1) : "0.0",
      modDiffs,
      eventModMatches,
      eventModMatchRate: matchedStudents > 0 ? ((eventModMatches / matchedStudents) * 100).toFixed(1) : "0.0",
      eventModDiffs,
      exactMatches,
      exactMatchRate: matchedStudents > 0 ? ((exactMatches / matchedStudents) * 100).toFixed(1) : "0.0",
      passedHereFailedPub,
      failedHerePassedPub,
      heldMismatches,
      notInDataset,
      heldbackCount,
      totalCalcMod: Math.round(totalCalcMod * 100) / 100,
      totalPubMod: Math.round(totalPubMod * 100) / 100,
      totalCalcEventMod: Math.round(totalCalcEventMod * 100) / 100,
      totalPubEventMod: Math.round(totalPubEventMod * 100) / 100,
      totalRegMod: Math.round(totalRegMod * 100) / 100,
      modTotalDiff: Math.round((totalCalcMod - totalPubMod) * 100) / 100,
      eventModTotalDiff: Math.round((totalCalcEventMod - totalPubEventMod) * 100) / 100
    };
  }, [comparisonRecords]);

  const uniqueComparisonColleges = useMemo(() => {
    const set = new Set();
    comparisonRecords.forEach(r => {
      if (r.college) set.add(r.college);
    });
    return Array.from(set).sort();
  }, [comparisonRecords]);

  const filteredComparisonRecords = useMemo(() => {
    let list = comparisonRecords;

    if (comparisonCollegeFilter !== "ALL") {
      list = list.filter(r => r.college === comparisonCollegeFilter);
    }

    if (comparisonFilterStatus === "MISMATCH") {
      list = list.filter(r => r.isResultMatch === false && r.isFound);
    } else if (comparisonFilterStatus === "MOD_DIFF") {
      list = list.filter(r => r.isModMatch === false && r.isFound);
    } else if (comparisonFilterStatus === "EXACT_MATCH") {
      list = list.filter(r => r.isExactMatch === true);
    } else if (comparisonFilterStatus === "PASSED_HERE_FAILED_PUB") {
      list = list.filter(r => r.category === "PASSED_HERE_FAILED_PUB");
    } else if (comparisonFilterStatus === "FAILED_HERE_PASSED_PUB") {
      list = list.filter(r => r.category === "FAILED_HERE_PASSED_PUB");
    } else if (comparisonFilterStatus === "HELD_MISMATCH") {
      list = list.filter(r => r.category === "HELD_MISMATCH" && !r.calcIsHeldback);
    } else if (comparisonFilterStatus === "NOT_IN_DATA") {
      list = list.filter(r => !r.isFound);
    } else if (comparisonFilterStatus === "HELDBACK") {
      list = list.filter(r => !!r.calcIsHeldback);
    }

    if (comparisonSearch.trim()) {
      const q = comparisonSearch.toLowerCase().trim();
      list = list.filter(r => 
        (r.prn && r.prn.toLowerCase().includes(q)) ||
        (r.seat && r.seat.toLowerCase().includes(q)) ||
        (r.name && r.name.toLowerCase().includes(q)) ||
        (r.college && r.college.toLowerCase().includes(q)) ||
        (r.reappear && r.reappear.toLowerCase().includes(q))
      );
    }

    return list;
  }, [comparisonRecords, comparisonCollegeFilter, comparisonFilterStatus, comparisonSearch]);

  const handleExportComparisonExcel = (recordsToExport = filteredComparisonRecords, customFilename = null) => {
    if (!recordsToExport || recordsToExport.length === 0) {
      setStatus("No comparison records to export.", "warning");
      return;
    }

    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Detailed Comparison
      const headers = [
        "Examination Seat Number",
        "PRN",
        "Student Name",
        "College Code",
        "College Name",
        "Our Calculated Result",
        "Published Result (University)",
        "Result Status Agreement",
        "Regular Event Mod (Gazette)",
        "Our Event Mod (Supplementary)",
        "Published Event Mod (Supplementary)",
        "Our Mod Total (Regular + Supp)",
        "Published Ordinance (Ord Total)",
        "Moderation Difference",
        "Moderation Concordance",
        "Our Failed Courses",
        "Published Reappear Paper Codes",
        "Reconciliation Status",
        "Discrepancy Details"
      ];

      const rows = recordsToExport.map(r => {
        return [
          r.seat,
          r.prn,
          r.name,
          r.collegeCode,
          r.college,
          r.calcResult,
          r.pubExactResult || r.pubResultStd,
          r.isResultMatch ? "MATCH" : "MISMATCH",
          r.calcRegularModMarks || 0,
          r.calcEventModMarks || 0,
          r.pubEventModMarks || 0,
          r.calcModMarks,
          r.pubModMarks,
          r.modDiff > 0 ? `+${r.modDiff.toFixed(2)}` : `${r.modDiff.toFixed(2)}`,
          r.isModMatch ? "MATCH" : "DIFF",
          r.calcFailedCourses.join(", ") || "None",
          r.reappear || "None",
          r.category,
          r.discrepancyDescription
        ];
      });

      const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      ws["!cols"] = [
        { wch: 18 }, // Seat
        { wch: 20 }, // PRN
        { wch: 26 }, // Name
        { wch: 14 }, // College Code
        { wch: 36 }, // College Name
        { wch: 22 }, // Calc Result
        { wch: 22 }, // Pub Result
        { wch: 18 }, // Result Agreement
        { wch: 20 }, // Reg Mod (Gazette)
        { wch: 22 }, // Our Event Mod
        { wch: 22 }, // Pub Event Mod
        { wch: 24 }, // Our Mod Total
        { wch: 24 }, // Pub Ord Total
        { wch: 18 }, // Mod Diff
        { wch: 18 }, // Mod Match
        { wch: 30 }, // Calc Failed
        { wch: 30 }, // Pub Reappear
        { wch: 22 }, // Status
        { wch: 40 }  // Details
      ];

      XLSX.utils.book_append_sheet(wb, ws, "Comparison Details");

      // Sheet 2: Summary Reconciliation KPIs
      const summaryRows = [
        ["RESULT RECONCILIATION & ORDINANCE COMPARISON SUMMARY"],
        [],
        ["Metric", "Value"],
        ["Total Published Records Compared", comparisonKPIs.total],
        ["Records Matched in Loaded Dataset", `${comparisonKPIs.matchedStudents} (${comparisonKPIs.matchRate}%)`],
        ["Result Status Agreement Rate", `${comparisonKPIs.resultMatchRate}% (${comparisonKPIs.resultMatches} / ${comparisonKPIs.matchedStudents})`],
        ["Result Status Mismatches", comparisonKPIs.resultMismatches],
        ["  - Passed in Our System, Failed in Published", comparisonKPIs.passedHereFailedPub],
        ["  - Failed in Our System, Passed in Published", comparisonKPIs.failedHerePassedPub],
        ["  - Held Status Discrepancy", comparisonKPIs.heldMismatches],
        ["Ordinance / Moderation Match Rate", `${comparisonKPIs.modMatchRate}% (${comparisonKPIs.modMatches} / ${comparisonKPIs.matchedStudents})`],
        ["Moderation Mark Differences", comparisonKPIs.modDiffs],
        ["Event Moderation Marks (Our System)", comparisonKPIs.totalCalcEventMod],
        ["Event Moderation Marks (University Published)", comparisonKPIs.totalPubEventMod],
        ["Regular Event Moderation Marks (Uploaded Gazette)", comparisonKPIs.totalRegMod],
        ["Total Moderation Marks Awarded (Our Mod Total)", comparisonKPIs.totalCalcMod],
        ["Total Ordinance Marks Awarded (Pub Ord Total)", comparisonKPIs.totalPubMod],
        ["Overall Moderation Net Difference", `${comparisonKPIs.modTotalDiff > 0 ? "+" : ""}${comparisonKPIs.modTotalDiff}`],
        ["Full Exact Concordance (Result + Moderation Match)", `${comparisonKPIs.exactMatches} (${comparisonKPIs.exactMatchRate}%)`],
        [],
        ["Report Generated At", new Date().toLocaleString()]
      ];

      const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
      wsSummary["!cols"] = [{ wch: 48 }, { wch: 28 }];
      XLSX.utils.book_append_sheet(wb, wsSummary, "Reconciliation Summary");

      const fname = customFilename || `result_and_ordinance_comparison_${new Date().toISOString().slice(0, 10)}.xlsx`;
      XLSX.writeFile(wb, fname);
      setStatus(`Exported ${recordsToExport.length} comparison records to ${fname}`, "success");
    } catch (err) {
      console.error(err);
      setStatus(`Failed to export comparison Excel: ${err.message}`, "error");
    }
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

    // Gazette / Tabulation Register indicators
    if (normCols.some(c => c.includes("reappear") || c.includes("failsubject"))) {
      score += 10;
      matchedHeaders.push("Reappear Papers");
    }
    if (normCols.some(c => c.includes("resultstatus") || c === "result" || c.includes("result"))) {
      score += 8;
      matchedHeaders.push("Result Status");
    }
    if (normCols.some(c => c.includes("sgpa") || c.includes("cgpa") || c.includes("grandtotal"))) {
      score += 6;
      matchedHeaders.push("Marks / SGPA");
    }

    return { score, matchedHeaders };
  };

  const parseSheetWithHeaderScan = (ws) => {
    // raw:true returns raw cell values (numbers stay numbers) — faster than formatted output
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
    if (!aoa || aoa.length === 0) return { rows: [], headerMap: {}, headerRowIdx: 0, score: 0, matchedHeaders: [] };

    let bestRowIdx = 0;
    let maxScore = -1;
    let bestMatchedHeaders = [];

    for (let r = 0; r < Math.min(aoa.length, 50); r++) {
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
        validHeaderIndices.push({ idx, name });
      }
    });

    const numCols = validHeaderIndices.length;
    const rows = [];
    for (let r = bestRowIdx + 1; r < aoa.length; r++) {
      const rowArr = aoa[r];
      if (!Array.isArray(rowArr) || rowArr.length === 0) continue;
      // Fast empty-row check: look for any non-empty value in the row
      let hasValue = false;
      for (let c = 0; c < rowArr.length; c++) {
        const v = rowArr[c];
        if (v !== "" && v !== null && v !== undefined) { hasValue = true; break; }
      }
      if (!hasValue) continue;
      const rowObj = {};
      for (let c = 0; c < numCols; c++) {
        const { idx, name } = validHeaderIndices[c];
        rowObj[name] = rowArr[idx] !== undefined ? rowArr[idx] : "";
      }
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
        // Performance: skip cell styles, formatting, HTML and stub cells for faster parsing
        const wb = XLSX.read(data, {
          type: "array",
          cellStyles: false,
          cellNF: false,
          cellHTML: false,
          sheetStubs: false
        });
        setWorkbook(wb);
        setSheetNames(wb.SheetNames);

        const meta = {};
        let bestSheet = wb.SheetNames[0];
        let highestTotalScore = -1;
        const sheetParseCache = new Map(); // Cache parse results to avoid double-parsing

        wb.SheetNames.forEach((sName) => {
          const ws = wb.Sheets[sName];
          const parsed = parseSheetWithHeaderScan(ws);
          sheetParseCache.set(sName, parsed);
          const { score, matchedHeaders, headerRowIdx, rows } = parsed;
          
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

        // Use cached result instead of re-parsing
        const { rows, headerMap: hMap, headerRowIdx, matchedHeaders } = sheetParseCache.get(bestSheet);

        if (!rows || rows.length === 0) {
          setStatus("No valid data rows found in sheet " + bestSheet + ".", "warning");
          setIsProcessing(false);
          return;
        }

        setHeaderMap(hMap);
        setRawRows(rows);

        const baseGrouped = buildGroupedRecordsFromRows(rows, hMap, absentRecordsMap, malpracticeRecordsMap, heldbackRecordsMap, historicalRecordsMap, improvementScoringMode);
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

      const baseGrouped = buildGroupedRecordsFromRows(rows, hMap, absentRecordsMap, malpracticeRecordsMap, heldbackRecordsMap, historicalRecordsMap, improvementScoringMode);
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

  const handleClearSourceFile = () => {
    setSourceFile(null);
    setWorkbook(null);
    setSheetNames([]);
    setSelectedSheet("");
    setSheetMeta({});
    setHeaderMap({});
    setRawRows([]);
    setPage(0);
    if (sourceFileInputRef.current) {
      sourceFileInputRef.current.value = "";
    }
    setStatus("Source file removed. Upload a new ADES Marks Excel file to continue.", "info");
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
        const wb = XLSX.read(data, {
          type: "array",
          cellStyles: false,
          cellNF: false,
          cellHTML: false,
          sheetStubs: false
        });
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        const jsonRows = XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: true });

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
    const allRows = [];
    const parsedFileNames = [];

    for (const file of files) {
      try {
        const buffer = await file.arrayBuffer();
        const data = new Uint8Array(buffer);
        // Performance: skip formatting/style parsing — only need cell values
        const wb = XLSX.read(data, {
          type: "array",
          cellStyles: false,
          cellNF: false,
          cellHTML: false,
          sheetStubs: false
        });
        let fileRowsCount = 0;

        for (const sheetName of wb.SheetNames) {
          const ws = wb.Sheets[sheetName];
          if (!ws || !ws["!ref"]) continue;
          const { rows } = parseSheetWithHeaderScan(ws);
          // Use parsed rows if header detection succeeded; otherwise raw JSON fallback
          const effectiveRows = (rows && rows.length > 0) ? rows : XLSX.utils.sheet_to_json(ws, { defval: "", raw: true });
          if (effectiveRows && effectiveRows.length > 0) {
            // Push individually instead of concat to avoid repeated array copies
            for (let i = 0; i < effectiveRows.length; i++) {
              allRows.push(effectiveRows[i]);
            }
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
        const updatedGrouped = buildGroupedRecordsFromRows(rawRows, headerMap, map, malpracticeRecordsMap, heldbackRecordsMap, historicalRecordsMap, improvementScoringMode);
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

    if (absentFileInputRef.current) {
      absentFileInputRef.current.value = "";
    }

    let currentRows = rawRows;
    let currentHMap = headerMap;
    if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
      const ws = workbook.Sheets[selectedSheet];
      if (ws) {
        const parsed = parseSheetWithHeaderScan(ws);
        currentRows = parsed.rows;
        currentHMap = parsed.headerMap;
        setRawRows(currentRows);
        setHeaderMap(currentHMap);
      }
    }

    if (currentRows && currentRows.length > 0 && currentHMap) {
      const resetGrouped = buildGroupedRecordsFromRows(currentRows, currentHMap, new Map(), malpracticeRecordsMap, heldbackRecordsMap, historicalRecordsMap, improvementScoringMode);
      setGroupedRecords(resetGrouped);
    }
    setPage(0);
    setComparisonPage(0);

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
        "STUDENT ALPHA",
        "2099010000000001",
        "AA99ARTS001",
        "Bachelor of Arts in Sample Studies",
        "Humanities",
        "FYUGP-2099",
        "BA Year II",
        "SEMESTER IV",
        "",
        "AA",
        "Alpha Arts and Science College, Sample City",
        "XT9DSCSAM101",
        "Introduction to Management",
        "Lec-Lab",
        "ESE",
        "TH",
        "Marked Absent During Mark Entry"
      ],
      [
        "STUDENT BETA",
        "2099010000000002",
        "AA99ARTS002",
        "Bachelor of Arts in Sample Studies",
        "Humanities",
        "FYUGP-2099",
        "BA Year II",
        "SEMESTER IV",
        "",
        "AA",
        "Alpha Arts and Science College, Sample City",
        "XT9DSCSAM102",
        "Fundamentals of Economics",
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
        const updatedGrouped = buildGroupedRecordsFromRows(rawRows, headerMap, absentRecordsMap, map, heldbackRecordsMap, historicalRecordsMap, improvementScoringMode);
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

    if (malpracticeFileInputRef.current) {
      malpracticeFileInputRef.current.value = "";
    }

    let currentRows = rawRows;
    let currentHMap = headerMap;
    if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
      const ws = workbook.Sheets[selectedSheet];
      if (ws) {
        const parsed = parseSheetWithHeaderScan(ws);
        currentRows = parsed.rows;
        currentHMap = parsed.headerMap;
        setRawRows(currentRows);
        setHeaderMap(currentHMap);
      }
    }

    if (currentRows && currentRows.length > 0 && currentHMap) {
      const resetGrouped = buildGroupedRecordsFromRows(currentRows, currentHMap, absentRecordsMap, new Map(), heldbackRecordsMap, historicalRecordsMap, improvementScoringMode);
      setGroupedRecords(resetGrouped);
    }
    setPage(0);
    setComparisonPage(0);

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
        "STUDENT ALPHA",
        "2099010000000001",
        "BB99COMM001",
        "Bachelor of Commerce (Sample)",
        "Commerce",
        "FYUGP-2099",
        "BCom Year II",
        "SEMESTER IV",
        "",
        "BB",
        "Beta Commerce College, Sample City",
        "XT9DSCSAM201",
        "1",
        "Beta Commerce College, Sample City",
        "Environmental Studies",
        "Lec-Lab",
        "ESE",
        "TH",
        "EHB",
        "malpractice sample entry",
        "01/01/2099"
      ],
      [
        "STUDENT BETA",
        "2099010000000002",
        "CC99ARTS001",
        "Bachelor of Arts in Sample Studies",
        "Humanities",
        "FYUGP-2099",
        "BA Year II",
        "SEMESTER IV",
        "",
        "CC",
        "Gamma Science College, Sample District",
        "XT9DSCSAM202",
        "1",
        "Gamma Science College, Sample District",
        "Principles of Political Science",
        "Lec-Lab",
        "ESE",
        "TH",
        "EHB",
        "",
        "01/01/2099"
      ],
      [
        "STUDENT GAMMA",
        "2099010000000003",
        "CC99ARTS002",
        "Bachelor of Arts in Sample Studies",
        "Humanities",
        "FYUGP-2099",
        "BA Year II",
        "SEMESTER IV",
        "",
        "CC",
        "Gamma Science College, Sample District",
        "XT9DSCSAM203",
        "1",
        "Gamma Science College, Sample District",
        "Introduction to Sociology",
        "Lec-Lab",
        "ESE",
        "TH",
        "EHB",
        "",
        "01/01/2099"
      ],
      [
        "STUDENT DELTA",
        "2099010000000004",
        "DD99COMM001",
        "Bachelor of Commerce (Sample)",
        "Commerce",
        "FYUGP-2099",
        "BCom Year II",
        "SEMESTER IV",
        "",
        "DD",
        "Delta Academy, Sample Town",
        "XT9DSCSAM204",
        "1",
        "Delta Academy, Sample Town",
        "Financial Accounting",
        "Lec-Lab",
        "ESE",
        "TH",
        "EHB",
        "SMP",
        "01/01/2099"
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

      // Dynamically re-calculate grouped records and all downstream evaluations
      let currentRows = rawRows;
      let currentHMap = headerMap;
      if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
        const ws = workbook.Sheets[selectedSheet];
        if (ws) {
          const parsed = parseSheetWithHeaderScan(ws);
          currentRows = parsed.rows;
          currentHMap = parsed.headerMap;
          setRawRows(currentRows);
          setHeaderMap(currentHMap);
        }
      }

      if (currentRows && currentRows.length > 0 && currentHMap) {
        const updatedGrouped = buildGroupedRecordsFromRows(currentRows, currentHMap, absentRecordsMap, malpracticeRecordsMap, map, historicalRecordsMap, improvementScoringMode);
        setGroupedRecords(updatedGrouped);
      } else if (groupedRecords && groupedRecords.length > 0) {
        // Fallback: direct in-memory update across existing grouped records
        const updatedGrouped = groupedRecords.map(rec => {
          const heldbackEntry = getHeldbackEntry(rec.identifiers.prn, rec.identifiers.seat, rec.identifiers.code, map);
          return {
            ...rec,
            is_heldback: !!heldbackEntry,
            heldback_reason: heldbackEntry ? (heldbackEntry.reason || "Heldback at term-level") : ""
          };
        });
        setGroupedRecords(updatedGrouped);
      }
      setPage(0);
      setComparisonPage(0);

      const fileDetails = parsedFileNames.length === 1 ? `"${parsedFileNames[0]}"` : `${parsedFileNames.length} programme file(s) [${parsedFileNames.join(", ")}]`;
      setStatus(`Successfully imported ${list.length} heldback record(s) from ${fileDetails}. Calculations dynamically updated with "Held (Heldback)" status.`, "success");
    } catch (err) {
      console.error("Error importing heldback sheet(s):", err);
      setStatus("Failed to read heldback report: " + err.message, "error");
    } finally {
      if (heldbackFileInputRef.current) heldbackFileInputRef.current.value = "";
    }
  };

  // Clear / Remove loaded Heldback Data and revert source marks dynamically
  const handleClearHeldbackData = () => {
    setHeldbackRecordsMap(new Map());
    setHeldbackList([]);
    setHeldbackFileName("");

    if (heldbackFileInputRef.current) {
      heldbackFileInputRef.current.value = "";
    }

    let currentRows = rawRows;
    let currentHMap = headerMap;
    if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
      const ws = workbook.Sheets[selectedSheet];
      if (ws) {
        const parsed = parseSheetWithHeaderScan(ws);
        currentRows = parsed.rows;
        currentHMap = parsed.headerMap;
        setRawRows(currentRows);
        setHeaderMap(currentHMap);
      }
    }

    if (currentRows && currentRows.length > 0 && currentHMap) {
      const resetGrouped = buildGroupedRecordsFromRows(currentRows, currentHMap, absentRecordsMap, malpracticeRecordsMap, new Map(), historicalRecordsMap, improvementScoringMode);
      setGroupedRecords(resetGrouped);
    } else if (groupedRecords && groupedRecords.length > 0) {
      const resetGrouped = groupedRecords.map(rec => ({
        ...rec,
        is_heldback: false,
        heldback_reason: ""
      }));
      setGroupedRecords(resetGrouped);
    }
    setPage(0);
    setComparisonPage(0);

    setStatus("Removed heldback records. Student evaluations and pass simulations dynamically recalculated.", "info");
  };

  // Upload Previous Event(s) ADES Reports for Carry Forward
  const handlePreviousReportsUpload = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    setIsProcessing(true);
    setStatus("Analyzing " + files.length + " previous event report(s)...", "info");

    try {
      const newReports = [];

      for (const file of files) {
        const arrayBuffer = await file.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const wb = XLSX.read(data, {
          type: "array",
          cellStyles: false,
          cellNF: false,
          cellHTML: false,
          sheetStubs: false
        });

        let allFileRows = [];
        let chosenHMap = null;
        for (const sName of wb.SheetNames) {
          const ws = wb.Sheets[sName];
          const parsed = parseSheetWithHeaderScan(ws);
          if (parsed.rows.length > 0 && parsed.score >= 5) {
            allFileRows = parsed.rows;
            chosenHMap = parsed.headerMap;
            break;
          }
        }

        if (allFileRows.length > 0 && chosenHMap) {
          const records = extractHistoricalRecordsFromRows(allFileRows, chosenHMap, file.name);
          const courseProfiles = extractHistoricalCourseProfilesFromRows(allFileRows, chosenHMap, file.name);
          newReports.push({
            id: Date.now() + "_" + Math.random().toString(36).substring(2, 7),
            fileName: file.name,
            recordCount: records.length,
            courseCount: courseProfiles.size,
            parsedRecords: records,
            courseProfiles: courseProfiles,
            uploadTime: new Date().toLocaleTimeString()
          });
        }
      }

      if (newReports.length === 0) {
        setStatus("No valid student course records found in uploaded previous event ADES file(s).", "error");
        setIsProcessing(false);
        return;
      }

      const updatedReports = [...previousReports, ...newReports];
      const updatedMap = rebuildHistoricalMap(updatedReports);
      const updatedProfilesMap = rebuildHistoricalCourseProfiles(updatedReports);
      setPreviousReports(updatedReports);
      setHistoricalRecordsMap(updatedMap);
      setHistoricalCourseProfilesMap(updatedProfilesMap);

      let currentRows = rawRows;
      let currentHMap = headerMap;
      if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
        const ws = workbook.Sheets[selectedSheet];
        if (ws) {
          const parsed = parseSheetWithHeaderScan(ws);
          currentRows = parsed.rows;
          currentHMap = parsed.headerMap;
          setRawRows(currentRows);
          setHeaderMap(currentHMap);
        }
      }

      if (currentRows && currentRows.length > 0 && currentHMap) {
        const updatedGrouped = buildGroupedRecordsFromRows(
          currentRows,
          currentHMap,
          absentRecordsMap,
          malpracticeRecordsMap,
          heldbackRecordsMap,
          updatedMap,
          improvementScoringMode,
          updatedProfilesMap,
          historicalGazetteMap
        );
        setGroupedRecords(updatedGrouped);
      }

      setStatus(`Loaded ${newReports.length} earlier event ADES report(s) (${updatedMap.size} student records mapped across ${updatedProfilesMap.size} courses). Carry-forward marks applied.`, "success");
    } catch (err) {
      console.error("Error loading previous reports:", err);
      setStatus("Failed to read earlier event ADES reports: " + err.message, "error");
    } finally {
      if (prevReportsFileInputRef.current) prevReportsFileInputRef.current.value = "";
      setIsProcessing(false);
    }
  };

  const handleRemovePreviousReport = (id) => {
    const updatedReports = previousReports.filter(r => r.id !== id);
    const updatedMap = rebuildHistoricalMap(updatedReports);
    const updatedProfilesMap = rebuildHistoricalCourseProfiles(updatedReports);
    setPreviousReports(updatedReports);
    setHistoricalRecordsMap(updatedMap);
    setHistoricalCourseProfilesMap(updatedProfilesMap);

    let currentRows = rawRows;
    let currentHMap = headerMap;
    if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
      const ws = workbook.Sheets[selectedSheet];
      if (ws) {
        const parsed = parseSheetWithHeaderScan(ws);
        currentRows = parsed.rows;
        currentHMap = parsed.headerMap;
        setRawRows(currentRows);
        setHeaderMap(currentHMap);
      }
    }

    if (currentRows && currentRows.length > 0 && currentHMap) {
      const updatedGrouped = buildGroupedRecordsFromRows(
        currentRows,
        currentHMap,
        absentRecordsMap,
        malpracticeRecordsMap,
        heldbackRecordsMap,
        updatedMap,
        improvementScoringMode,
        updatedProfilesMap,
        historicalGazetteMap
      );
      setGroupedRecords(updatedGrouped);
    }

    setStatus("Removed earlier event ADES report. Calculations dynamically refreshed.", "info");
  };

  const handleClearAllPreviousReports = () => {
    setPreviousReports([]);
    setHistoricalRecordsMap(new Map());
    setHistoricalCourseProfilesMap(new Map());

    let currentRows = rawRows;
    let currentHMap = headerMap;
    if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
      const ws = workbook.Sheets[selectedSheet];
      if (ws) {
        const parsed = parseSheetWithHeaderScan(ws);
        currentRows = parsed.rows;
        currentHMap = parsed.headerMap;
        setRawRows(currentRows);
        setHeaderMap(currentHMap);
      }
    }

    if (currentRows && currentRows.length > 0 && currentHMap) {
      const updatedGrouped = buildGroupedRecordsFromRows(
        currentRows,
        currentHMap,
        absentRecordsMap,
        malpracticeRecordsMap,
        heldbackRecordsMap,
        new Map(),
        improvementScoringMode,
        new Map(),
        historicalGazetteMap
      );
      setGroupedRecords(updatedGrouped);
    }

    setStatus("Cleared all earlier event ADES reports. Calculations reverted to current attempt marks only.", "info");
  };

  // Upload Previous Semester University Result Gazette Reports (Optional Reappear Assurance)
  const handleGazetteReportsUpload = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;

    setIsProcessing(true);
    setStatus("Analyzing " + files.length + " university result gazette(s)...", "info");

    try {
      const newReports = [];

      for (const file of files) {
        const arrayBuffer = await file.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const wb = XLSX.read(data, {
          type: "array",
          cellStyles: false,
          cellNF: false,
          cellHTML: false,
          sheetStubs: false
        });

        let fileGazetteRecords = [];

        // Scan all sheets in the workbook and aggregate student records
        for (const sName of wb.SheetNames) {
          const ws = wb.Sheets[sName];
          const parsed = parseSheetWithHeaderScan(ws);
          if (parsed.rows.length > 0 && parsed.headerMap) {
            const sheetRecs = extractGazetteRecordsFromRows(parsed.rows, parsed.headerMap, `${file.name} [${sName}]`);
            if (sheetRecs.length > 0) {
              fileGazetteRecords.push(...sheetRecs);
            }
          }
        }

        // If no records found with multi-sheet scan, fallback to sheet 0
        if (fileGazetteRecords.length === 0 && wb.SheetNames.length > 0) {
          const ws = wb.Sheets[wb.SheetNames[0]];
          const parsed = parseSheetWithHeaderScan(ws);
          if (parsed.rows.length > 0 && parsed.headerMap) {
            const sheetRecs = extractGazetteRecordsFromRows(parsed.rows, parsed.headerMap, file.name);
            if (sheetRecs.length > 0) {
              fileGazetteRecords.push(...sheetRecs);
            }
          }
        }

        if (fileGazetteRecords.length > 0) {
          const totalReappears = fileGazetteRecords.reduce((acc, r) => acc + (r.reappearCodes ? r.reappearCodes.length : 0), 0);
          newReports.push({
            id: Date.now() + "_" + Math.random().toString(36).substring(2, 7),
            type: "gazette",
            fileName: file.name,
            recordCount: fileGazetteRecords.length,
            reappearCount: totalReappears,
            parsedGazetteRecords: fileGazetteRecords,
            uploadTime: new Date().toLocaleTimeString()
          });
        }
      }

      if (newReports.length === 0) {
        setStatus("No valid student result gazette entries found in uploaded file(s). Ensure sheet includes PRN/Seat and Result Status or Reappear Paper Codes.", "error");
        setIsProcessing(false);
        return;
      }

      const updatedGazetteReports = [...gazetteReports, ...newReports];
      const updatedGazetteMap = rebuildHistoricalGazetteMap(updatedGazetteReports);
      setGazetteReports(updatedGazetteReports);
      setHistoricalGazetteMap(updatedGazetteMap);

      let currentRows = rawRows;
      let currentHMap = headerMap;
      if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
        const ws = workbook.Sheets[selectedSheet];
        if (ws) {
          const parsed = parseSheetWithHeaderScan(ws);
          currentRows = parsed.rows;
          currentHMap = parsed.headerMap;
          setRawRows(currentRows);
          setHeaderMap(currentHMap);
        }
      }

      if (currentRows && currentRows.length > 0 && currentHMap) {
        const updatedGrouped = buildGroupedRecordsFromRows(
          currentRows,
          currentHMap,
          absentRecordsMap,
          malpracticeRecordsMap,
          heldbackRecordsMap,
          historicalRecordsMap,
          improvementScoringMode,
          historicalCourseProfilesMap,
          updatedGazetteMap
        );
        setGroupedRecords(updatedGrouped);
      }

      const totalStudents = updatedGazetteReports.reduce((acc, r) => acc + r.recordCount, 0);
      const totalReappears = updatedGazetteReports.reduce((acc, r) => acc + r.reappearCount, 0);
      setStatus(`Loaded ${newReports.length} university result gazette(s) (${totalStudents} students, ${totalReappears} reappear backlogs indexed for cross-verification).`, "success");
    } catch (err) {
      console.error("Error loading gazette reports:", err);
      setStatus("Failed to read result gazette file(s): " + err.message, "error");
    } finally {
      if (gazetteFileInputRef.current) gazetteFileInputRef.current.value = "";
      setIsProcessing(false);
    }
  };

  const handleRemoveGazetteReport = (id) => {
    const updatedGazetteReports = gazetteReports.filter(r => r.id !== id);
    const updatedGazetteMap = rebuildHistoricalGazetteMap(updatedGazetteReports);
    setGazetteReports(updatedGazetteReports);
    setHistoricalGazetteMap(updatedGazetteMap);

    let currentRows = rawRows;
    let currentHMap = headerMap;
    if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
      const ws = workbook.Sheets[selectedSheet];
      if (ws) {
        const parsed = parseSheetWithHeaderScan(ws);
        currentRows = parsed.rows;
        currentHMap = parsed.headerMap;
        setRawRows(currentRows);
        setHeaderMap(currentHMap);
      }
    }

    if (currentRows && currentRows.length > 0 && currentHMap) {
      const updatedGrouped = buildGroupedRecordsFromRows(
        currentRows,
        currentHMap,
        absentRecordsMap,
        malpracticeRecordsMap,
        heldbackRecordsMap,
        historicalRecordsMap,
        improvementScoringMode,
        historicalCourseProfilesMap,
        updatedGazetteMap
      );
      setGroupedRecords(updatedGrouped);
    }

    setStatus("Removed result gazette report. Cross-verification badges dynamically updated.", "info");
  };

  const handleClearAllGazetteReports = () => {
    setGazetteReports([]);
    setHistoricalGazetteMap(new Map());

    let currentRows = rawRows;
    let currentHMap = headerMap;
    if ((!currentRows || currentRows.length === 0) && workbook && selectedSheet) {
      const ws = workbook.Sheets[selectedSheet];
      if (ws) {
        const parsed = parseSheetWithHeaderScan(ws);
        currentRows = parsed.rows;
        currentHMap = parsed.headerMap;
        setRawRows(currentRows);
        setHeaderMap(currentHMap);
      }
    }

    if (currentRows && currentRows.length > 0 && currentHMap) {
      const updatedGrouped = buildGroupedRecordsFromRows(
        currentRows,
        currentHMap,
        absentRecordsMap,
        malpracticeRecordsMap,
        heldbackRecordsMap,
        historicalRecordsMap,
        improvementScoringMode,
        historicalCourseProfilesMap,
        new Map()
      );
      setGroupedRecords(updatedGrouped);
    }

    setStatus("Cleared all result gazette reports.", "info");
  };

  const handleDownloadGazetteTemplate = () => {
    const headers = [
      "TermUID", "Exam Event", "Program Full Name", "Program Name", "Program Code", "Branch Desc", 
      "Program Part Name", "Program Part Abbrevation", "Program Part Term", "Name of Student", 
      "Vernacular Name", "Mother's name", "PRN-Permanent Registration Number", "RegionalCenterCode", 
      "RegionalCenterName", "College Code", "College Name", "Gender", "Mobile", "Category", 
      "Admitted UNDER Category", "Physically Challenged", "Eligiblity Status", "Examination Seat Number", 
      "Statement Number", "RESULT STATUS", "OTHER STATUS", "Reappear Paper Codes", "No of Fail Subjects", 
      "Grand Total", "Ordinance", "Ord Total", "THTotal", "TPTotal", "PRTotal", "Grade", 
      "Total Credits", "Earned Credits", "EGP", "SGPA", "CGPA", "Result Declaration Date", "Result Processed Date"
    ];

    const sampleRows = [
      [
        "80", "April 2025 (CT)", "Bachelor of Commerce(with Credits)-Regular-FYUGP-2024-Commerce", 
        "Bachelor of Commerce", "COM_BCOM(HONOURS)", "Commerce", "FIRST YEAR BACHELOR OF COMMERCE", 
        "BCom Year I", "SEMESTER II", "MUHAMMED JINAS P", "മുഹമ്മദ് ജിനാസ് പി", "", 
        "2024012600012471", "", "", "WM", "WMO Imam Gazzali Arts and Science College, Wayanad", 
        "Male", "8590333495", "Socially Economic Backward Class", "", "", "", "WM24COMR048", "", 
        "Fail", "", "KU2DSCBBA103", "1", "296", "NO", "", "259", "259", "37", "F", 
        "21.00", "17.00", "105.00", "5.00", "5.33", "30/12/2025", "5/08/2026"
      ],
      [
        "80", "April 2025 (CT)", "Bachelor of Commerce(with Credits)-Regular-FYUGP-2024-Commerce", 
        "Bachelor of Commerce", "COM_BCOM(HONOURS)", "Commerce", "FIRST YEAR BACHELOR OF COMMERCE", 
        "BCom Year I", "SEMESTER II", "JUNAID P", "ജുനൈദ് പി", "", 
        "2024012600012518", "", "", "WM", "WMO Imam Gazzali Arts and Science College, Wayanad", 
        "Male", "9544575118", "Socially Economic Backward Class", "", "", "", "WM24COMR037", "", 
        "Pass", "", "", "0", "321", "NO", "", "280", "280", "41", "B", 
        "21.00", "21.00", "131.00", "6.24", "5.55", "30/12/2025", "5/08/2026"
      ],
      [
        "80", "April 2025 (CT)", "Bachelor of Commerce(with Credits)-Regular-FYUGP-2024-Commerce", 
        "Bachelor of Commerce", "COM_BCOM(HONOURS)", "Commerce", "FIRST YEAR BACHELOR OF COMMERCE", 
        "BCom Year I", "SEMESTER II", "MUHAMMED SWAFWAN NK", "മുഹമ്മദ് സ്വഫ്വാൻ എൻകെ", "", 
        "2024012600012531", "", "", "WM", "WMO Imam Gazzali Arts and Science College, Wayanad", 
        "Male", "9778743795", "Socially Economic Backward Class", "", "", "", "WM24COMR060", "", 
        "Fail", "", "KU2DSCBBA103,KU2DSCCOM105", "2", "303", "NO", "", "239", "239", "64", "F", 
        "21.00", "13.00", "87.00", "4.14", "4.71", "30/12/2025", "5/08/2026"
      ]
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "ResultGazette_Template");
    XLSX.writeFile(wb, "University_Result_Gazette_Template.xlsx");
    setStatus("Downloaded sample University Result Gazette template (.xlsx).", "success");
  };

  const handleDownloadPrevReportTemplate = () => {
    const headers = [
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
      "CE Overall Marks",
      "Overall Maximum",
      "Overall Minimum",
      "Course Overall Marks",
      "Course Pass/Fail",
      "ADEC Code",
      "ADEC Name"
    ];

    const sampleRows = [
      [
        "Faculty of Science", "B.Sc. Semester II", "XT9DSCSAM101", "Introduction to Computing",
        "AA99SCI001", "2099010000000001", 25, 8, 20, 50, 15, 28, 75, 23, 48,
        25, 0, 18, 0, 0, 0, 25, 0, 18, 100, 35, 66, "Pass", "AA", "Alpha Science College"
      ],
      [
        "Faculty of Commerce", "B.Com. Semester II", "XT9DSCSAM102", "Principles of Accounting",
        "AA99COM002", "2099010000000002", 0, 0, 0, 75, 23, 18, 75, 23, 18,
        0, 0, 0, 25, 0, 19, 25, 0, 19, 100, 35, 37, "Fail", "BB", "Beta Commerce College"
      ]
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    XLSX.utils.book_append_sheet(wb, ws, "PreviousEventMarks");
    XLSX.writeFile(wb, "Previous_Event_ADES_Baseline_Template.xlsx");
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
        "AA99ARTS001",
        "2099010000000001",
        "APC WITHIN CONDONABLE LIMIT",
        "Heldback at term-level",
        "AA",
        "STUDENT ALPHA",
        "Heldback at term-level",
        "Alpha Arts and Science College, Sample Town",
        "Heldback at term-level",
        "Heldback at term-level"
      ],
      [
        "BB99COMM001",
        "2099010000000002",
        "APC WCL",
        "Heldback at term-level",
        "BB",
        "STUDENT BETA",
        "Heldback at term-level",
        "Beta Commerce College, Example City",
        "Heldback at term-level",
        "Heldback at term-level"
      ],
      [
        "CC99ARTS002",
        "2099010000000003",
        "APC WCL",
        "Heldback at term-level",
        "CC",
        "STUDENT GAMMA",
        "Heldback at term-level",
        "Gamma Government College, Demo District",
        "Heldback at term-level",
        "Heldback at term-level"
      ],
      [
        "DD99COMM002",
        "2099010000000004",
        "APC WCL",
        "Heldback at term-level",
        "DD",
        "STUDENT DELTA",
        "Heldback at term-level",
        "Delta Science and Arts College, Test Town",
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
        "Improvement Papers",
        "Supplementary Papers",
        "Papers Passed",
        "Papers Failed",
        "Heldback / Missing Papers",
        "Raw Semester Result (0 Mod)",
        "Final Semester Result",
        "Semester Rescued via Moderation",
        "Regular Event Mod (Gazette)",
        "Supplementary Event Mod",
        "Our Mod Total (Regular + Supp)",
        "Failed Courses List",
        "All Attempted Courses Breakdown"
      ];

      // Dynamic summary stats based on the exported subset
      let totalAttempted = 0;
      let totalImpPapers = 0;
      let totalSuppPapers = 0;
      let totalPassed = 0;
      let totalFailed = 0;
      let totalRawPass = 0;
      let totalFinalPass = 0;
      let totalRescued = 0;
      let totalHeld = 0;
      let totalHeldback = 0;

      const isGazetteUploaded = (historicalGazetteMap && historicalGazetteMap.size > 0) || (gazetteReports && gazetteReports.length > 0);
      const rows = studentsToExport.map(st => {
        totalAttempted += st.totalCourses;
        totalImpPapers += (st.improvementCourses || 0);
        totalSuppPapers += (st.supplementaryCourses || 0);
        totalPassed += st.finalPassedCourses;
        totalFailed += st.finalFailedCourses;
        if (st.isHeldback) totalHeldback++;
        if (st.isHeld) totalHeld++;
        if (st.rawSemesterPass) totalRawPass++;
        if (st.finalSemesterPass) totalFinalPass++;
        if (st.isRescuedSemester) totalRescued++;

        const failedCurrent = st.courses.filter(c => c.coursePass === "Fail").map(c => `${c.courseCode} (${c.courseName})`);
        const allPending = [...failedCurrent, ...(st.pendingGazetteBacklogs || []).map(b => `${b} (Pending Gazette Backlog)`)];
        const failedList = allPending.join("; ");
        const coursesSummary = st.courses.map(c => {
          const typeTag = (c.isImprovement || c.attemptType === "IMPROVEMENT") ? "[IMP]" : "[SUPP]";
          return `${c.courseCode} ${typeTag}: ${c.coursePass}${c.modMarks > 0 ? ` (+${c.modMarks} Mod)` : ''}`;
        }).join("; ");
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
          st.improvementCourses || 0,
          st.supplementaryCourses || 0,
          st.finalPassedCourses,
          st.finalFailedCourses,
          heldInfo,
          st.rawSemesterResult || "",
          st.semesterResult || "",
          st.isRescuedSemester ? "Yes (Rescued)" : isGazetteUploaded ? "No" : "",
          st.regularModerationMarks || 0,
          st.eventModerationMarks || 0,
          st.totalModerationMarks || 0,
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
        "-",
        totalAttempted,
        totalImpPapers,
        totalSuppPapers,
        totalPassed,
        totalFailed,
        `${totalHeld} Held (${totalHeldback} Heldback)`,
        isGazetteUploaded ? `${totalRawPass} Passed (${rawPct}%)` : "-",
        isGazetteUploaded ? `${totalFinalPass} Passed (${finalPct}%)` : "-",
        isGazetteUploaded ? `+${totalRescued} Rescued (${resPct}%)` : "-",
        "-",
        "-",
        "-",
        isGazetteUploaded ? `${failCount} Failed (${failPct}%)` : "-",
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
        { wch: 26 }, // College
        { wch: 28 }, // Program
        { wch: 16 }, // Seat Number
        { wch: 18 }, // PRN
        { wch: 16 }, // Papers Attempted
        { wch: 18 }, // Improvement Papers
        { wch: 20 }, // Supplementary Papers
        { wch: 14 }, // Papers Passed
        { wch: 14 }, // Papers Failed
        { wch: 24 }, // Heldback / Missing Papers
        { wch: 24 }, // Raw Semester Result
        { wch: 24 }, // Final Semester Result
        { wch: 22 }, // Semester Rescued
        { wch: 20 }, // Regular Event Mod (Gazette)
        { wch: 20 }, // Supplementary Event Mod
        { wch: 22 }, // Our Mod Total (Regular + Supp)
        { wch: 40 }, // Failed Courses List
        { wch: 65 }  // All Attempted Courses Breakdown
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
    const hasCollegeFilter = selectedCollegeFilter !== "ALL";
    const hasProgFilter = selectedProgramFilter !== "ALL";
    const hasCourseFilter = selectedCourseFilter !== "ALL";
    const hasFacultyFilter = selectedFacultyFilter !== "ALL";
    const hasSearch = !!searchQuery;
    const q = hasSearch ? searchQuery.toLowerCase().trim() : "";
    const cleanedCourseFilter = hasCourseFilter ? cleanCourseCode(selectedCourseFilter) : "";
    const activeFilterEntries = Object.entries(columnFilters);
    const hasColFilters = activeFilterEntries.length > 0;

    // Single-pass filter — avoids 6× full array traversals
    const result = [];
    for (const r of processedRows) {
      if (hasFacultyFilter && r["Faculty"] !== selectedFacultyFilter) continue;
      if (hasCollegeFilter) {
        const col = r._college || r["College Name"] || r["College Code"] || r["College"];
        if (col !== selectedCollegeFilter) continue;
      }
      if (hasProgFilter && r["Program Term Name"] !== selectedProgramFilter) continue;
      if (hasCourseFilter && cleanCourseCode(r["Course Code"]) !== cleanedCourseFilter) continue;
      if (hasSearch) {
        const col = r._college;
        const colMatch = col && col.toLowerCase().includes(q);
        if (!colMatch && !Object.values(r).some(val => String(val || "").toLowerCase().includes(q))) continue;
      }
      if (hasColFilters) {
        let pass = true;
        for (const [colKey, filterVal] of activeFilterEntries) {
          if (!filterVal || filterVal.trim() === "") continue;
          const cellVal = String(r[colKey] !== undefined && r[colKey] !== null ? r[colKey] : "").toLowerCase();
          if (!cellVal.includes(filterVal.toLowerCase().trim())) { pass = false; break; }
        }
        if (!pass) continue;
      }
      result.push(r);
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
    const size = pageSize === "ALL" ? (filteredRows.length || 1) : Number(pageSize);
    const start = page * size;
    return filteredRows.slice(start, start + size);
  }, [filteredRows, page, pageSize]);

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

  // MacBook Air M2 Keyboard Shortcuts & Native macOS Menu Bridge
  useEffect(() => {
    const handleKeyDown = (e) => {
      // Cmd+B / Ctrl+B: Toggle Sidebar
      if ((e.metaKey || e.ctrlKey) && (e.key === "b" || e.key === "B")) {
        e.preventDefault();
        setIsSidebarCollapsed(c => !c);
      }
      // Cmd+Shift+F: Toggle Focus Mode
      if ((e.metaKey || e.ctrlKey) && e.shiftKey && (e.key === "f" || e.key === "F")) {
        e.preventDefault();
        setIsTableMaximized(m => !m);
      }
      // Cmd+O: Open File Picker
      if ((e.metaKey || e.ctrlKey) && (e.key === "o" || e.key === "O")) {
        e.preventDefault();
        sourceFileInputRef.current?.click();
      }
      // Cmd+E: Export Active View
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && (e.key === "e" || e.key === "E")) {
        e.preventDefault();
        if (processedRows.length > 0) {
          handleExportExcel(filteredRows);
        }
      }
    };

    window.addEventListener("keydown", handleKeyDown);

    const unsubscribeMenu = window.electronAPI?.onMenuAction?.((action) => {
      if (action === "toggle-sidebar") setIsSidebarCollapsed(c => !c);
      if (action === "toggle-focus") setIsTableMaximized(m => !m);
      if (action === "open-file") sourceFileInputRef.current?.click();
      if (action === "export-excel" && processedRows.length > 0) handleExportExcel(filteredRows);
      if (action === "export-master" && processedRows.length > 0) handleExportExcel(processedRows);
      if (action === "open-guide") setShowHelpModal(true);
    });

    return () => {
      window.removeEventListener("keydown", handleKeyDown);
      if (unsubscribeMenu) unsubscribeMenu();
    };
  }, [processedRows, filteredRows]);

  const coursePassKey = "Course Pass/Fail";

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", width: "100vw", overflow: "hidden", background: "var(--bg)", color: "var(--ink)" }}>
      
      {/* Top Application Bar */}
      {!isTableMaximized && (
        <header className="app-top-header" style={{ 
          display: "flex", 
          justifyContent: "space-between", 
          alignItems: "center", 
          padding: "10px 24px", 
          paddingLeft: (typeof window !== "undefined" && ((window.electronAPI && window.electronAPI.isDesktop) || /Electron/i.test(navigator.userAgent)) && /Mac/i.test(navigator.platform || navigator.userAgent)) ? "96px" : "24px",
          borderBottom: "1px solid var(--line)", 
          background: "var(--panel)", 
          flexShrink: 0,
          gap: "12px",
          flexWrap: "wrap"
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: "14px", flexWrap: "wrap" }}>
            <Link to="/" style={{ 
              display: "flex", 
              alignItems: "center", 
              gap: "6px", 
              color: "var(--ink)", 
              textDecoration: "none", 
              fontSize: "12.5px", 
              fontWeight: 600,
              padding: "5px 12px",
              borderRadius: "8px",
              background: "var(--bg)",
              border: "1px solid var(--line)",
              transition: "all 0.15s ease"
            }}>
              <ArrowLeft size={15} /> Back
            </Link>
            <div style={{ height: "20px", width: "1px", background: "var(--line)" }} />
            <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
              <div style={{ 
                width: "32px", 
                height: "32px", 
                borderRadius: "8px", 
                background: "rgba(99, 102, 241, 0.12)", 
                display: "flex", 
                alignItems: "center", 
                justifyContent: "center",
                flexShrink: 0 
              }}>
                <Calculator size={18} color="#6366f1" />
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
                <h2 style={{ fontSize: "15px", fontWeight: 700, margin: 0, color: "var(--ink)", letterSpacing: "-0.2px" }}>
                  ADES Supplementary / Improvement Calculator
                </h2>
                <span style={{ fontSize: "10.5px", background: "rgba(99, 102, 241, 0.15)", color: "#6366f1", padding: "2px 8px", borderRadius: "10px", fontWeight: 700 }}>
                  Multi-Event Baseline
                </span>
              </div>
            </div>

            {/* Mode Switcher Pill */}
            <div style={{ display: "flex", gap: "3px", background: "var(--bg)", padding: "3px", borderRadius: "9px", border: "1px solid var(--line)" }}>
              <Link 
                to="/ades-result-calculator"
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "5px",
                  padding: "3px 9px",
                  fontSize: "11.5px",
                  fontWeight: 600,
                  color: "var(--muted)",
                  borderRadius: "6px",
                  textDecoration: "none"
                }}
              >
                🎓 Regular Event
              </Link>
              <span
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: "5px",
                  padding: "3px 9px",
                  fontSize: "11.5px",
                  fontWeight: 700,
                  background: "#6366f1",
                  color: "white",
                  borderRadius: "6px",
                  boxShadow: "0 1px 3px rgba(99, 102, 241, 0.3)"
                }}
              >
                🔄 Supplementary &amp; Improvement
              </span>
            </div>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap" }}>
            <button 
              type="button" 
              className="secondary" 
              onClick={() => setShowHelpModal(true)}
              style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", padding: "6px 12px", borderRadius: "8px" }}
            >
              <HelpCircle size={14} /> Guide
            </button>

            {processedRows.length > 0 && (
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
                  borderRadius: "8px", 
                  fontWeight: 600, 
                  cursor: "pointer",
                  boxShadow: "0 2px 6px rgba(0,0,0,0.08)"
                }}
                title={filteredRows.length !== processedRows.length ? `Export current filtered view (${filteredRows.length} rows)` : "Export all 31-column ADES results"}
              >
                <Download size={14} /> Export 31-Col XLSX ({filteredRows.length !== processedRows.length ? `${filteredRows.length}/${processedRows.length}` : `${filteredRows.length}`})
              </button>
            )}
          </div>
        </header>
      )}

      {/* Sub-Header: Segmented Tab Bar & Secondary Export Actions */}
      <div style={{ 
        display: "flex", 
        justifyContent: "space-between", 
        alignItems: "center", 
        padding: "8px 24px", 
        background: "var(--panel)", 
        borderBottom: "1px solid var(--line)", 
        flexShrink: 0,
        gap: "12px",
        flexWrap: "wrap"
      }}>
        {/* Segmented Tab Pill Navigation */}
        <div style={{ display: "flex", background: "var(--bg)", padding: "3px", borderRadius: "10px", border: "1px solid var(--line)", gap: "3px", overflowX: "auto" }}>
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
              borderRadius: "7px",
              border: "none",
              cursor: "pointer",
              background: activeTab === "results" ? "var(--accent)" : "transparent",
              color: activeTab === "results" ? "white" : "var(--muted)",
              transition: "all 0.15s ease",
              boxShadow: activeTab === "results" ? "0 1px 3px rgba(0,0,0,0.12)" : "none"
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
              borderRadius: "7px",
              border: "none",
              cursor: "pointer",
              background: activeTab === "students" ? "var(--accent)" : "transparent",
              color: activeTab === "students" ? "white" : "var(--muted)",
              transition: "all 0.15s ease",
              boxShadow: activeTab === "students" ? "0 1px 3px rgba(0,0,0,0.12)" : "none"
            }}
          >
            <Users size={13} /> Student Results ({studentMetrics.totalStudents})
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
              borderRadius: "7px",
              border: "none",
              cursor: "pointer",
              background: activeTab === "simulation" ? "var(--accent)" : "transparent",
              color: activeTab === "simulation" ? "white" : "var(--muted)",
              transition: "all 0.15s ease",
              boxShadow: activeTab === "simulation" ? "0 1px 3px rgba(0,0,0,0.12)" : "none"
            }}
          >
            <TrendingUp size={13} /> Pass Simulation (+0..+10)
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
              borderRadius: "7px",
              border: "none",
              cursor: "pointer",
              background: activeTab === "moderation" ? "var(--accent)" : "transparent",
              color: activeTab === "moderation" ? "white" : "var(--muted)",
              transition: "all 0.15s ease",
              boxShadow: activeTab === "moderation" ? "0 1px 3px rgba(0,0,0,0.12)" : "none"
            }}
          >
            <Sliders size={13} /> Moderation Matrix ({Object.values(courseModerationMap).filter(v => v > 0).length})
          </button>
          <button
            type="button"
            onClick={() => setActiveTab("comparison")}
            style={{
              display: "flex",
              alignItems: "center",
              gap: "6px",
              padding: "5px 12px",
              fontSize: "12px",
              fontWeight: 600,
              borderRadius: "7px",
              border: "none",
              cursor: "pointer",
              background: activeTab === "comparison" ? "var(--accent)" : "transparent",
              color: activeTab === "comparison" ? "white" : "var(--muted)",
              transition: "all 0.15s ease",
              boxShadow: activeTab === "comparison" ? "0 1px 3px rgba(0,0,0,0.12)" : "none"
            }}
          >
            <Scale size={13} /> Reconciliation {comparisonRecords.length > 0 ? `(${comparisonRecords.length})` : ""}
          </button>
        </div>

        {/* Action Export Buttons for active tab context */}
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          {activeTab === "comparison" && comparisonRecords.length > 0 && (
            <button 
              type="button" 
              onClick={handleExportComparisonExcel}
              style={{ 
                display: "flex", 
                alignItems: "center", 
                gap: "6px", 
                padding: "5px 12px", 
                fontSize: "12px", 
                background: "#8b5cf6", 
                color: "white", 
                border: "none", 
                borderRadius: "7px", 
                fontWeight: 600, 
                cursor: "pointer",
                boxShadow: "0 1px 4px rgba(139, 92, 246, 0.25)"
              }}
              title="Export Full Result & Ordinance Reconciliation Report to Excel"
            >
              <Download size={13} /> Export Comparison ({filteredComparisonRecords.length})
            </button>
          )}

          {activeTab === "students" && processedRows.length > 0 && (
            <button 
              type="button" 
              onClick={() => handleExportStudentSemesterExcel(filteredStudents)}
              style={{ 
                display: "flex", 
                alignItems: "center", 
                gap: "6px", 
                padding: "5px 12px", 
                fontSize: "12px", 
                background: "#6366f1", 
                color: "white", 
                border: "none", 
                borderRadius: "7px", 
                fontWeight: 600, 
                cursor: "pointer",
                boxShadow: "0 1px 4px rgba(99, 102, 241, 0.25)"
              }}
              title="Export student semester results"
            >
              <Download size={13} /> Export Students ({filteredStudents.length !== studentSemesterData.length ? `${filteredStudents.length}/${studentMetrics.totalStudents}` : studentMetrics.totalStudents})
            </button>
          )}

          {activeTab === "simulation" && processedRows.length > 0 && (
            <button 
              type="button" 
              onClick={handleExportSimulationExcel}
              style={{ 
                display: "flex", 
                alignItems: "center", 
                gap: "6px", 
                padding: "5px 12px", 
                fontSize: "12px", 
                background: "#10b981", 
                color: "white", 
                border: "none", 
                borderRadius: "7px", 
                fontWeight: 600, 
                cursor: "pointer",
                boxShadow: "0 1px 4px rgba(16, 185, 129, 0.25)"
              }}
              title="Export course-wise pass count under 0 to +10 moderation marks"
            >
              <Download size={13} /> Export Simulation (+0..+10)
            </button>
          )}

          {/* Sidebar Collapse/Expand Toggle Button */}
          <button
            type="button"
            onClick={() => setIsSidebarCollapsed(c => !c)}
            style={{
              display: "flex",
              alignItems: "center",
              gap: "6px",
              padding: "5px 10px",
              fontSize: "11.5px",
              fontWeight: 600,
              background: isSidebarCollapsed ? "var(--accent-soft)" : "var(--bg)",
              color: isSidebarCollapsed ? "var(--accent)" : "var(--ink)",
              border: `1px solid ${isSidebarCollapsed ? "var(--accent)" : "var(--line)"}`,
              borderRadius: "7px",
              cursor: "pointer",
              transition: "all 0.15s ease"
            }}
            title={isSidebarCollapsed ? "Expand files and configuration sidebar" : "Collapse sidebar to maximize table width"}
          >
            {isSidebarCollapsed ? <PanelLeftOpen size={14} /> : <PanelLeftClose size={14} />}
            <span>{isSidebarCollapsed ? "Show Sidebar" : "Hide Sidebar"}</span>
          </button>
        </div>
      </div>

      {/* Main Workspace Layout */}
      <div style={{ display: "flex", flex: 1, overflow: "hidden", position: "relative" }}>
        
        {/* Left Sidebar / Config Panel (MacBook Air M2 Optimized Accordions) */}
        <aside style={{ 
          width: isSidebarCollapsed ? "0px" : `${sidebarWidth}px`, 
          borderRight: isSidebarCollapsed ? "none" : "1px solid var(--line)", 
          background: "var(--panel)", 
          display: isSidebarCollapsed ? "none" : "flex", 
          flexDirection: "column", 
          flexShrink: 0, 
          overflowY: "auto", 
          overflowX: "hidden",
          padding: isSidebarCollapsed ? "0" : "10px", 
          gap: "8px",
          transition: isDraggingSidebar.current ? "none" : "width 0.2s ease",
          position: "relative"
        }}>
          
          {/* Sidebar Header with Width Presets, Expand/Collapse All, and Close Button */}
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid var(--line)", paddingBottom: "6px", flexShrink: 0 }}>
            <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
              <Layers size={13} color="var(--accent)" />
              <span style={{ fontSize: "11px", fontWeight: 700, color: "var(--ink)" }}>Controls</span>
              <button
                type="button"
                onClick={() => {
                  const anyClosed = Object.values(sidebarAccordions).some(v => !v);
                  setAllAccordions(anyClosed);
                }}
                title={Object.values(sidebarAccordions).some(v => !v) ? "Expand all sidebar sections" : "Fold all sidebar sections"}
                style={{
                  background: "transparent",
                  border: "none",
                  color: "var(--accent)",
                  cursor: "pointer",
                  fontSize: "10px",
                  padding: "0 3px",
                  fontWeight: 600
                }}
              >
                {Object.values(sidebarAccordions).some(v => !v) ? "Expand All" : "Fold All"}
              </button>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
              <div style={{ display: "flex", background: "var(--bg)", borderRadius: "4px", padding: "1px", border: "1px solid var(--line)" }}>
                <button
                  type="button"
                  onClick={() => setSidebarWidth(220)}
                  title="Slim Sidebar (220px)"
                  style={{ padding: "1px 5px", fontSize: "9px", fontWeight: 600, border: "none", borderRadius: "3px", cursor: "pointer", background: sidebarWidth === 220 ? "var(--accent)" : "transparent", color: sidebarWidth === 220 ? "white" : "var(--muted)" }}
                >
                  Slim
                </button>
                <button
                  type="button"
                  onClick={() => setSidebarWidth(250)}
                  title="Default Sidebar (250px)"
                  style={{ padding: "1px 5px", fontSize: "9px", fontWeight: 600, border: "none", borderRadius: "3px", cursor: "pointer", background: sidebarWidth === 250 ? "var(--accent)" : "transparent", color: sidebarWidth === 250 ? "white" : "var(--muted)" }}
                >
                  250px
                </button>
                <button
                  type="button"
                  onClick={() => setSidebarWidth(290)}
                  title="Wide Sidebar (290px)"
                  style={{ padding: "1px 5px", fontSize: "9px", fontWeight: 600, border: "none", borderRadius: "3px", cursor: "pointer", background: sidebarWidth === 290 ? "var(--accent)" : "transparent", color: sidebarWidth === 290 ? "white" : "var(--muted)" }}
                >
                  290px
                </button>
              </div>
              <button
                type="button"
                onClick={() => setIsSidebarCollapsed(true)}
                title="Collapse sidebar (Cmd+B)"
                style={{ background: "transparent", border: "none", color: "var(--muted)", cursor: "pointer", padding: "2px", borderRadius: "4px", display: "flex", alignItems: "center" }}
              >
                <PanelLeftClose size={13} />
              </button>
            </div>
          </div>

          {/* Accordion 1: Source Marksheet & Sheet Selection */}
          <div style={{ border: "1px solid var(--line)", borderRadius: "7px", overflow: "hidden", background: "var(--panel)" }}>
            <div 
              onClick={() => toggleAccordion("files")}
              style={{ 
                display: "flex", 
                alignItems: "center", 
                justifyContent: "space-between", 
                padding: "7px 9px", 
                background: sidebarAccordions.files ? "var(--bg)" : "var(--panel)", 
                cursor: "pointer", 
                userSelect: "none" 
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                <FileSpreadsheet size={13} color="var(--accent)" />
                <span style={{ fontSize: "11px", fontWeight: 700, color: "var(--ink)" }}>Source Marksheet</span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                <span style={{ 
                  fontSize: "9px", 
                  fontWeight: 600, 
                  padding: "1px 5px", 
                  borderRadius: "8px", 
                  background: sourceFile ? "rgba(16, 185, 129, 0.12)" : "rgba(239, 68, 68, 0.1)", 
                  color: sourceFile ? "#059669" : "#dc2626" 
                }}>
                  {sourceFile ? "Loaded" : "Required"}
                </span>
                {sidebarAccordions.files ? <ChevronDown size={12} color="var(--muted)" /> : <ChevronRight size={12} color="var(--muted)" />}
              </div>
            </div>

            {sidebarAccordions.files && (
              <div style={{ padding: "9px", display: "flex", flexDirection: "column", gap: "8px", borderTop: "1px solid var(--line)" }}>
                {/* File Upload Box */}
                <div style={{ 
                  background: sourceFile ? "rgba(59, 130, 246, 0.05)" : "var(--bg)", 
                  border: sourceFile ? "1.5px solid rgba(59, 130, 246, 0.4)" : "1.5px dashed var(--line)", 
                  borderRadius: "6px", 
                  padding: "10px", 
                  textAlign: "center", 
                  position: "relative" 
                }}>
                  <input 
                    type="file" 
                    ref={sourceFileInputRef}
                    accept=".xlsx,.xls,.csv" 
                    onChange={handleFileUpload}
                    style={{ display: "none" }}
                  />
                  {sourceFile ? (
                    <div style={{ display: "flex", flexDirection: "column", gap: "6px", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: "6px", width: "100%", justifyContent: "space-between" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: "6px", overflow: "hidden" }}>
                          <FileSpreadsheet size={15} color="var(--accent)" style={{ flexShrink: 0 }} />
                          <span style={{ fontSize: "11px", fontWeight: 700, color: "var(--ink)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }} title={sourceFile}>
                            {sourceFile}
                          </span>
                        </div>
                        <button
                          type="button"
                          onClick={handleClearSourceFile}
                          title="Remove source marksheet file"
                          style={{ background: "transparent", border: "none", color: "var(--danger)", cursor: "pointer", padding: "2px", display: "flex", alignItems: "center" }}
                        >
                          <X size={13} />
                        </button>
                      </div>
                      <div style={{ fontSize: "10px", color: "var(--muted)", width: "100%", textAlign: "left" }}>
                        {rawRows.length} raw rows • {groupedRecords.length} courses loaded
                      </div>
                      <div style={{ display: "flex", gap: "6px", width: "100%", marginTop: "2px" }}>
                        <button
                          type="button"
                          onClick={() => sourceFileInputRef.current?.click()}
                          style={{
                            flex: 1,
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            gap: "4px",
                            padding: "4px 8px",
                            fontSize: "10.5px",
                            background: "var(--panel)",
                            color: "var(--ink)",
                            border: "1px solid var(--line)",
                            borderRadius: "4px",
                            fontWeight: 600,
                            cursor: "pointer"
                          }}
                        >
                          <Upload size={11} /> Replace
                        </button>
                        <button
                          type="button"
                          onClick={handleClearSourceFile}
                          style={{
                            padding: "4px 8px",
                            fontSize: "10.5px",
                            background: "rgba(239, 68, 68, 0.1)",
                            color: "#ef4444",
                            border: "1px solid rgba(239, 68, 68, 0.25)",
                            borderRadius: "4px",
                            fontWeight: 600,
                            cursor: "pointer",
                            display: "flex",
                            alignItems: "center",
                            gap: "3px"
                          }}
                        >
                          <Trash2 size={11} /> Clear
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div 
                      onClick={() => sourceFileInputRef.current?.click()}
                      style={{ cursor: "pointer" }}
                    >
                      <FileSpreadsheet size={24} color="var(--accent)" style={{ margin: "0 auto 6px", opacity: 0.8 }} />
                      <div style={{ fontSize: "12px", fontWeight: 600, color: "var(--ink)" }}>
                        Upload ADES Marks Excel
                      </div>
                      <div style={{ fontSize: "10px", color: "var(--muted)", marginTop: "2px" }}>
                        Click to browse or drop .xlsx / .xls
                      </div>
                    </div>
                  )}
                </div>

                {/* Sheet Selector (if multiple sheets exist) */}
                {sheetNames.length > 1 && (
                  <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
                    <label style={{ fontSize: "10.5px", fontWeight: 600, color: "var(--muted)" }}>Source Sheet:</label>
                    <select 
                      value={selectedSheet} 
                      onChange={(e) => handleSheetChange(e.target.value)}
                      style={{ padding: "4px 8px", fontSize: "11px", borderRadius: "5px", border: "1px solid var(--line)", background: "var(--bg)" }}
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
              </div>
            )}
          </div>

          {/* Accordion 2: Baseline & Gazettes */}
          <div style={{ border: "1px solid var(--line)", borderRadius: "7px", overflow: "hidden", background: "var(--panel)" }}>
            <div 
              onClick={() => toggleAccordion("baseline")}
              style={{ 
                display: "flex", 
                alignItems: "center", 
                justifyContent: "space-between", 
                padding: "7px 9px", 
                background: sidebarAccordions.baseline ? "var(--bg)" : "var(--panel)", 
                cursor: "pointer", 
                userSelect: "none" 
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                <Layers size={13} color="#6366f1" />
                <span style={{ fontSize: "11px", fontWeight: 700, color: "var(--ink)" }}>Baseline &amp; Gazettes</span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                <span style={{ 
                  fontSize: "9px", 
                  fontWeight: 600, 
                  padding: "1px 5px", 
                  borderRadius: "8px", 
                  background: (previousReports.length + gazetteReports.length > 0) ? "rgba(99, 102, 241, 0.15)" : "var(--panel)", 
                  color: (previousReports.length + gazetteReports.length > 0) ? "#6366f1" : "var(--muted)", 
                  border: "1px solid var(--line)" 
                }}>
                  {previousReports.length + gazetteReports.length > 0 ? `${previousReports.length + gazetteReports.length} Files` : "Optional"}
                </span>
                {sidebarAccordions.baseline ? <ChevronDown size={12} color="var(--muted)" /> : <ChevronRight size={12} color="var(--muted)" />}
              </div>
            </div>

            {sidebarAccordions.baseline && (
              <div style={{ padding: "9px", display: "flex", flexDirection: "column", gap: "8px", borderTop: "1px solid var(--line)" }}>
                {/* Previous Event(s) Baseline Reports Card */}
                <div style={{ 
                  background: previousReports.length > 0 ? "rgba(99, 102, 241, 0.05)" : "var(--bg)", 
                  border: previousReports.length > 0 ? "1.5px solid rgba(99, 102, 241, 0.35)" : "1px solid var(--line)", 
                  borderRadius: "6px", 
                  padding: "10px", 
                  display: "flex", 
                  flexDirection: "column", 
                  gap: "6px" 
                }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ fontSize: "11px", fontWeight: 700, color: previousReports.length > 0 ? "#6366f1" : "var(--ink)", display: "flex", alignItems: "center", gap: "5px" }}>
                      <Layers size={13} color={previousReports.length > 0 ? "#6366f1" : "var(--muted)"} /> Baseline ADES
                    </div>
                    {previousReports.length > 0 ? (
                      <span style={{ fontSize: "9px", background: "#6366f1", color: "white", padding: "1px 5px", borderRadius: "8px", fontWeight: 700 }}>
                        {previousReports.length} Files ({historicalRecordsMap.size} Recs)
                      </span>
                    ) : (
                      <span style={{ fontSize: "9px", background: "rgba(99, 102, 241, 0.12)", color: "#6366f1", padding: "1px 5px", borderRadius: "8px", fontWeight: 700 }}>
                        CE/PR Baseline
                      </span>
                    )}
                  </div>

                  <div style={{ fontSize: "10px", color: "var(--muted)", lineHeight: "1.3" }}>
                    Carries forward CE &amp; Practical marks and min/max standards.
                  </div>

                  <input 
                    type="file" 
                    ref={prevReportsFileInputRef}
                    accept=".xlsx,.xls,.csv" 
                    multiple
                    onChange={handlePreviousReportsUpload}
                    style={{ display: "none" }}
                  />

                  <div style={{ display: "flex", gap: "5px", marginTop: "2px" }}>
                    <button 
                      type="button"
                      onClick={() => {
                        if (prevReportsFileInputRef.current) prevReportsFileInputRef.current.value = "";
                        prevReportsFileInputRef.current?.click();
                      }}
                      style={{
                        flex: 1,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        gap: "3px",
                        padding: "4px 6px",
                        fontSize: "10.5px",
                        background: previousReports.length > 0 ? "rgba(99, 102, 241, 0.12)" : "var(--panel)",
                        color: previousReports.length > 0 ? "#6366f1" : "var(--ink)",
                        border: previousReports.length > 0 ? "1px solid rgba(99, 102, 241, 0.3)" : "1px solid var(--line)",
                        borderRadius: "4px",
                        fontWeight: 600,
                        cursor: "pointer"
                      }}
                    >
                      <FileUp size={11} /> {previousReports.length > 0 ? "+ Files" : "Upload"}
                    </button>

                    <button 
                      type="button" 
                      onClick={handleDownloadPrevReportTemplate}
                      title="Download baseline template (.xlsx)"
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        gap: "3px",
                        padding: "4px 6px",
                        fontSize: "10.5px",
                        background: "var(--panel)",
                        color: "var(--ink)",
                        border: "1px solid var(--line)",
                        borderRadius: "4px",
                        fontWeight: 600,
                        cursor: "pointer"
                      }}
                    >
                      <FileDown size={11} /> Tpl
                    </button>

                    {previousReports.length > 0 && (
                      <button 
                        type="button" 
                        onClick={handleClearAllPreviousReports}
                        title="Clear all uploaded earlier event reports"
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          padding: "4px 6px",
                          fontSize: "10.5px",
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

                  {/* List of uploaded previous event files */}
                  {previousReports.length > 0 && (
                    <div style={{ display: "flex", flexDirection: "column", gap: "3px", marginTop: "3px", maxHeight: "110px", overflowY: "auto" }}>
                      {previousReports.map((rep) => (
                        <div 
                          key={rep.id} 
                          style={{ 
                            display: "flex", 
                            alignItems: "center", 
                            justifyContent: "space-between", 
                            padding: "3px 6px", 
                            background: "var(--panel)", 
                            border: "1px solid var(--line)", 
                            borderRadius: "4px", 
                            fontSize: "10px" 
                          }}
                        >
                          <div style={{ display: "flex", alignItems: "center", gap: "4px", overflow: "hidden" }}>
                            <FileSpreadsheet size={11} color="#6366f1" style={{ flexShrink: 0 }} />
                            <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "130px" }} title={rep.fileName}>
                              {rep.fileName}
                            </span>
                            <span style={{ fontSize: "8.5px", background: "rgba(99, 102, 241, 0.12)", color: "#6366f1", padding: "1px 3px", borderRadius: "3px", fontWeight: 700 }}>
                              {rep.recordCount}r
                            </span>
                          </div>
                          <button 
                            type="button" 
                            onClick={() => handleRemovePreviousReport(rep.id)}
                            title="Remove file"
                            style={{ background: "transparent", border: "none", color: "#ef4444", cursor: "pointer", display: "flex", alignItems: "center", padding: "1px" }}
                          >
                            <Trash2 size={11} />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}

                  {/* Improvement Scoring Strategy Toggle */}
                  <div style={{ marginTop: "3px", paddingTop: "3px", borderTop: "1px solid var(--line)", display: "flex", flexDirection: "column", gap: "2px" }}>
                    <span style={{ fontSize: "9px", fontWeight: 700, color: "var(--muted)", textTransform: "uppercase" }}>
                      Improvement Policy:
                    </span>
                    <div style={{ display: "flex", gap: "3px" }}>
                      <button
                        type="button"
                        onClick={() => {
                          setImprovementScoringMode("current");
                          if (rawRows && rawRows.length > 0 && headerMap) {
                            const updatedGrouped = buildGroupedRecordsFromRows(
                              rawRows, headerMap, absentRecordsMap, malpracticeRecordsMap, heldbackRecordsMap, historicalRecordsMap, "current", historicalCourseProfilesMap, historicalGazetteMap
                            );
                            setGroupedRecords(updatedGrouped);
                          }
                        }}
                        style={{
                          flex: 1,
                          padding: "2px 4px",
                          fontSize: "9.5px",
                          borderRadius: "3px",
                          border: improvementScoringMode === "current" ? "1px solid #6366f1" : "1px solid var(--line)",
                          background: improvementScoringMode === "current" ? "rgba(99, 102, 241, 0.15)" : "transparent",
                          color: improvementScoringMode === "current" ? "#6366f1" : "var(--muted)",
                          fontWeight: improvementScoringMode === "current" ? 700 : 500,
                          cursor: "pointer"
                        }}
                      >
                        Current
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setImprovementScoringMode("best");
                          if (rawRows && rawRows.length > 0 && headerMap) {
                            const updatedGrouped = buildGroupedRecordsFromRows(
                              rawRows, headerMap, absentRecordsMap, malpracticeRecordsMap, heldbackRecordsMap, historicalRecordsMap, "best", historicalCourseProfilesMap, historicalGazetteMap
                            );
                            setGroupedRecords(updatedGrouped);
                          }
                        }}
                        style={{
                          flex: 1,
                          padding: "2px 4px",
                          fontSize: "9.5px",
                          borderRadius: "3px",
                          border: improvementScoringMode === "best" ? "1px solid #6366f1" : "1px solid var(--line)",
                          background: improvementScoringMode === "best" ? "rgba(99, 102, 241, 0.15)" : "transparent",
                          color: improvementScoringMode === "best" ? "#6366f1" : "var(--muted)",
                          fontWeight: improvementScoringMode === "best" ? 700 : 500,
                          cursor: "pointer"
                        }}
                      >
                        Best of Both
                      </button>
                    </div>
                  </div>
                </div>

                {/* Result Gazette Card */}
                <div style={{ 
                  background: gazetteReports.length > 0 ? "rgba(16, 185, 129, 0.05)" : "var(--bg)", 
                  border: gazetteReports.length > 0 ? "1.5px solid rgba(16, 185, 129, 0.35)" : "1px solid var(--line)", 
                  borderRadius: "6px", 
                  padding: "10px", 
                  display: "flex", 
                  flexDirection: "column", 
                  gap: "6px" 
                }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ fontSize: "11px", fontWeight: 700, color: gazetteReports.length > 0 ? "#10b981" : "var(--ink)", display: "flex", alignItems: "center", gap: "5px" }}>
                      <FileCheck size={13} color={gazetteReports.length > 0 ? "#10b981" : "var(--muted)"} /> Result Gazette
                    </div>
                    {gazetteReports.length > 0 ? (
                      <span style={{ fontSize: "9px", background: "#10b981", color: "white", padding: "1px 5px", borderRadius: "8px", fontWeight: 700 }}>
                        {historicalGazetteMap.size} Students
                      </span>
                    ) : (
                      <span style={{ fontSize: "9px", background: "rgba(16, 185, 129, 0.12)", color: "#059669", padding: "1px 5px", borderRadius: "8px", fontWeight: 700 }}>
                        Optional
                      </span>
                    )}
                  </div>

                  <div style={{ fontSize: "10px", color: "var(--muted)", lineHeight: "1.3" }}>
                    Cross-verifies declared <strong>Reappear Paper Codes</strong> and prior semester pass/fail.
                  </div>

                  <input 
                    type="file" 
                    ref={gazetteFileInputRef}
                    accept=".xlsx,.xls,.csv" 
                    multiple
                    onChange={handleGazetteReportsUpload}
                    style={{ display: "none" }}
                  />

                  <div style={{ display: "flex", gap: "5px", marginTop: "2px" }}>
                    <button 
                      type="button"
                      onClick={() => {
                        if (gazetteFileInputRef.current) gazetteFileInputRef.current.value = "";
                        gazetteFileInputRef.current?.click();
                      }}
                      style={{
                        flex: 1,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        gap: "3px",
                        padding: "4px 6px",
                        fontSize: "10.5px",
                        background: gazetteReports.length > 0 ? "rgba(16, 185, 129, 0.12)" : "var(--panel)",
                        color: gazetteReports.length > 0 ? "#059669" : "var(--ink)",
                        border: gazetteReports.length > 0 ? "1px solid rgba(16, 185, 129, 0.3)" : "1px solid var(--line)",
                        borderRadius: "4px",
                        fontWeight: 600,
                        cursor: "pointer"
                      }}
                    >
                      <FileUp size={11} /> {gazetteReports.length > 0 ? "+ Gazettes" : "Upload"}
                    </button>

                    <button 
                      type="button" 
                      onClick={handleDownloadGazetteTemplate}
                      title="Download gazette template (.xlsx)"
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        gap: "3px",
                        padding: "4px 6px",
                        fontSize: "10.5px",
                        background: "var(--panel)",
                        color: "var(--ink)",
                        border: "1px solid var(--line)",
                        borderRadius: "4px",
                        fontWeight: 600,
                        cursor: "pointer"
                      }}
                    >
                      <FileDown size={11} /> Tpl
                    </button>

                    {gazetteReports.length > 0 && (
                      <button 
                        type="button" 
                        onClick={handleClearAllGazetteReports}
                        title="Clear gazette reports"
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          padding: "4px 6px",
                          fontSize: "10.5px",
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

                  {/* List of uploaded gazette files */}
                  {gazetteReports.length > 0 && (
                    <div style={{ display: "flex", flexDirection: "column", gap: "3px", marginTop: "3px", maxHeight: "100px", overflowY: "auto" }}>
                      {gazetteReports.map((rep) => (
                        <div 
                          key={rep.id} 
                          style={{ 
                            display: "flex", 
                            alignItems: "center", 
                            justifyContent: "space-between", 
                            padding: "3px 6px", 
                            background: "var(--panel)", 
                            border: "1px solid var(--line)", 
                            borderRadius: "4px", 
                            fontSize: "10px" 
                          }}
                        >
                          <div style={{ display: "flex", alignItems: "center", gap: "4px", overflow: "hidden" }}>
                            <FileCheck size={11} color="#10b981" style={{ flexShrink: 0 }} />
                            <span style={{ whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "130px" }} title={rep.fileName}>
                              {rep.fileName}
                            </span>
                          </div>
                          <button 
                            type="button" 
                            onClick={() => handleRemoveGazetteReport(rep.id)}
                            title="Remove file"
                            style={{ background: "transparent", border: "none", color: "#ef4444", cursor: "pointer", display: "flex", alignItems: "center", padding: "1px" }}
                          >
                            <Trash2 size={11} />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>

          {/* Accordion 3: Special Records (Absent, MP, Heldback) */}
          <div style={{ border: "1px solid var(--line)", borderRadius: "7px", overflow: "hidden", background: "var(--panel)" }}>
            <div 
              onClick={() => toggleAccordion("special")}
              style={{ 
                display: "flex", 
                alignItems: "center", 
                justifyContent: "space-between", 
                padding: "7px 9px", 
                background: sidebarAccordions.special ? "var(--bg)" : "var(--panel)", 
                cursor: "pointer", 
                userSelect: "none" 
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                <ShieldAlert size={13} color="#d97706" />
                <span style={{ fontSize: "11px", fontWeight: 700, color: "var(--ink)" }}>Special Records</span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                <span style={{ 
                  fontSize: "9px", 
                  fontWeight: 600, 
                  padding: "1px 5px", 
                  borderRadius: "8px", 
                  background: (absentList.length + malpracticeList.length + heldbackList.length > 0) ? "rgba(245, 158, 11, 0.15)" : "var(--panel)", 
                  color: (absentList.length + malpracticeList.length + heldbackList.length > 0) ? "#b45309" : "var(--muted)", 
                  border: "1px solid var(--line)" 
                }}>
                  {absentList.length + malpracticeList.length + heldbackList.length > 0 ? `${absentList.length + malpracticeList.length + heldbackList.length} Records` : "Optional"}
                </span>
                {sidebarAccordions.special ? <ChevronDown size={12} color="var(--muted)" /> : <ChevronRight size={12} color="var(--muted)" />}
              </div>
            </div>

            {sidebarAccordions.special && (
              <div style={{ padding: "9px", display: "flex", flexDirection: "column", gap: "8px", borderTop: "1px solid var(--line)" }}>
                {/* Absent Mark Entry Card */}
                <div style={{ background: absentList.length > 0 ? "rgba(239, 68, 68, 0.05)" : "var(--bg)", border: absentList.length > 0 ? "1.5px solid rgba(239, 68, 68, 0.35)" : "1px solid var(--line)", borderRadius: "6px", padding: "9px", display: "flex", flexDirection: "column", gap: "6px" }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ fontSize: "11px", fontWeight: 700, color: absentList.length > 0 ? "#ef4444" : "var(--ink)", display: "flex", alignItems: "center", gap: "5px" }}>
                      <UserX size={13} color={absentList.length > 0 ? "#ef4444" : "var(--muted)"} /> Absent Entry
                    </div>
                    {absentList.length > 0 && (
                      <span style={{ fontSize: "9px", background: "#ef4444", color: "white", padding: "1px 5px", borderRadius: "8px", fontWeight: 700 }}>
                        {absentList.length} Active
                      </span>
                    )}
                  </div>
                  <input type="file" ref={absentFileInputRef} accept=".xlsx,.xls,.csv" multiple onChange={handleAbsentExcelUpload} style={{ display: "none" }} />
                  <div style={{ display: "flex", gap: "4px" }}>
                    <button type="button" onClick={() => absentFileInputRef.current?.click()} style={{ flex: 1, padding: "4px 6px", fontSize: "10.5px", background: "var(--panel)", color: "var(--ink)", border: "1px solid var(--line)", borderRadius: "4px", fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: "3px" }}>
                      <FileUp size={11} /> {absentList.length > 0 ? "Replace" : "Upload"}
                    </button>
                    <button type="button" onClick={handleDownloadAbsentTemplate} title="Download template" style={{ padding: "4px 6px", fontSize: "10.5px", background: "var(--panel)", color: "var(--ink)", border: "1px solid var(--line)", borderRadius: "4px", fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: "3px" }}>
                      <FileDown size={11} /> Tpl
                    </button>
                    {absentList.length > 0 && (
                      <button type="button" onClick={handleClearAbsentData} style={{ padding: "4px 6px", fontSize: "10.5px", background: "transparent", color: "#ef4444", border: "1px solid #ef4444", borderRadius: "4px", fontWeight: 600, cursor: "pointer" }}>Clear</button>
                    )}
                  </div>
                </div>

                {/* Malpractice Entry Card */}
                <div style={{ background: malpracticeList.length > 0 ? "rgba(245, 158, 11, 0.05)" : "var(--bg)", border: malpracticeList.length > 0 ? "1.5px solid rgba(245, 158, 11, 0.35)" : "1px solid var(--line)", borderRadius: "6px", padding: "9px", display: "flex", flexDirection: "column", gap: "6px" }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "5px", fontSize: "11px", fontWeight: 700, color: malpracticeList.length > 0 ? "#b45309" : "var(--ink)" }}>
                      <ShieldAlert size={13} color="#d97706" /> Malpractice (MP)
                    </div>
                    {malpracticeList.length > 0 && (
                      <span style={{ fontSize: "9px", background: "rgba(245, 158, 11, 0.15)", color: "#b45309", padding: "1px 5px", borderRadius: "8px", fontWeight: 700 }}>
                        {malpracticeList.length} Active
                      </span>
                    )}
                  </div>
                  <input type="file" ref={malpracticeFileInputRef} accept=".xlsx,.xls,.csv" multiple onChange={handleMalpracticeExcelUpload} style={{ display: "none" }} />
                  <div style={{ display: "flex", gap: "4px" }}>
                    <button type="button" onClick={() => malpracticeFileInputRef.current?.click()} style={{ flex: 1, padding: "4px 6px", fontSize: "10.5px", background: "var(--panel)", color: "var(--ink)", border: "1px solid var(--line)", borderRadius: "4px", fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: "3px" }}>
                      <FileUp size={11} /> {malpracticeList.length > 0 ? "Replace" : "Upload"}
                    </button>
                    <button type="button" onClick={handleDownloadMalpracticeTemplate} title="Download template" style={{ padding: "4px 6px", fontSize: "10.5px", background: "var(--panel)", color: "var(--ink)", border: "1px solid var(--line)", borderRadius: "4px", fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: "3px" }}>
                      <FileDown size={11} /> Tpl
                    </button>
                    {malpracticeList.length > 0 && (
                      <button type="button" onClick={handleClearMalpracticeData} style={{ padding: "4px 6px", fontSize: "10.5px", background: "transparent", color: "#b45309", border: "1px solid #b45309", borderRadius: "4px", fontWeight: 600, cursor: "pointer" }}>Clear</button>
                    )}
                  </div>
                </div>

                {/* Heldback Entry Card */}
                <div style={{ background: heldbackList.length > 0 ? "rgba(192, 38, 211, 0.05)" : "var(--bg)", border: heldbackList.length > 0 ? "1.5px solid rgba(192, 38, 211, 0.35)" : "1px solid var(--line)", borderRadius: "6px", padding: "9px", display: "flex", flexDirection: "column", gap: "6px" }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "5px", fontSize: "11px", fontWeight: 700, color: heldbackList.length > 0 ? "#c026d3" : "var(--ink)" }}>
                      <Lock size={13} color="#c026d3" /> Heldback Entry
                    </div>
                    {heldbackList.length > 0 && (
                      <span style={{ fontSize: "9px", background: "rgba(192, 38, 211, 0.15)", color: "#c026d3", padding: "1px 5px", borderRadius: "8px", fontWeight: 700 }}>
                        {heldbackList.length} Active
                      </span>
                    )}
                  </div>
                  <input type="file" ref={heldbackFileInputRef} accept=".xlsx,.xls,.csv" multiple onChange={handleHeldbackExcelUpload} style={{ display: "none" }} />
                  <div style={{ display: "flex", gap: "4px" }}>
                    <button type="button" onClick={() => heldbackFileInputRef.current?.click()} style={{ flex: 1, padding: "4px 6px", fontSize: "10.5px", background: "var(--panel)", color: "var(--ink)", border: "1px solid var(--line)", borderRadius: "4px", fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: "3px" }}>
                      <FileUp size={11} /> {heldbackList.length > 0 ? "Replace" : "Upload"}
                    </button>
                    <button type="button" onClick={handleDownloadHeldbackTemplate} title="Download template" style={{ padding: "4px 6px", fontSize: "10.5px", background: "var(--panel)", color: "var(--ink)", border: "1px solid var(--line)", borderRadius: "4px", fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: "3px" }}>
                      <FileDown size={11} /> Tpl
                    </button>
                    {heldbackList.length > 0 && (
                      <button type="button" onClick={handleClearHeldbackData} style={{ padding: "4px 6px", fontSize: "10.5px", background: "transparent", color: "#ef4444", border: "1px solid #ef4444", borderRadius: "4px", fontWeight: 600, cursor: "pointer" }}>Clear</button>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Accordion 4: Consolidated KPIs & Pass Simulation */}
          {processedRows.length > 0 && (
            <div style={{ border: "1px solid var(--line)", borderRadius: "7px", overflow: "hidden", background: "var(--panel)" }}>
              <div 
                onClick={() => toggleAccordion("overview")}
                style={{ 
                  display: "flex", 
                  alignItems: "center", 
                  justifyContent: "space-between", 
                  padding: "7px 9px", 
                  background: sidebarAccordions.overview ? "var(--bg)" : "var(--panel)", 
                  cursor: "pointer", 
                  userSelect: "none" 
                }}
              >
                <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                  <BarChart3 size={13} color="#10b981" />
                  <span style={{ fontSize: "11px", fontWeight: 700, color: "var(--ink)" }}>Consolidated KPIs</span>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                  <span style={{ 
                    fontSize: "9px", 
                    fontWeight: 600, 
                    padding: "1px 5px", 
                    borderRadius: "8px", 
                    background: "rgba(16, 185, 129, 0.15)", 
                    color: "#059669", 
                    border: "1px solid var(--line)" 
                  }}>
                    {studentMetrics.isGazetteUploaded ? `${studentMetrics.finalPassedPct}% Pass` : `${consolidatedCourseMetrics.total} Papers`}
                  </span>
                  {sidebarAccordions.overview ? <ChevronDown size={12} color="var(--muted)" /> : <ChevronRight size={12} color="var(--muted)" />}
                </div>
              </div>

              {sidebarAccordions.overview && (
                <div style={{ padding: "9px", display: "flex", flexDirection: "column", gap: "8px", borderTop: "1px solid var(--line)" }}>
                  {/* Student Semester Level Cards */}
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6px", fontSize: "10.5px" }}>
                    <div style={{ background: "var(--panel)", padding: "6px 8px", borderRadius: "5px", border: "1px solid var(--line)" }}>
                      <div style={{ color: "var(--muted)", fontSize: "10px" }}>Students</div>
                      <strong style={{ fontSize: "14px", color: "var(--ink)" }}>{consolidatedStudentMetrics.totalStudents}</strong>
                      <div style={{ fontSize: "9px", color: "var(--muted)" }}>{consolidatedStudentMetrics.totalPapersAttempted} papers</div>
                    </div>

                    <div style={{ background: "rgba(99, 102, 241, 0.08)", padding: "6px 8px", borderRadius: "5px", border: "1px solid rgba(99, 102, 241, 0.25)" }}>
                      <div style={{ color: "#6366f1", fontWeight: 600, fontSize: "10px" }}>Semester Pass %</div>
                      <strong style={{ fontSize: "14px", color: "#6366f1" }}>
                        {consolidatedStudentMetrics.isGazetteUploaded ? `${consolidatedStudentMetrics.finalPassedPct}%` : "—"}
                      </strong>
                      <div style={{ fontSize: "9px", color: "#6366f1" }}>
                        {consolidatedStudentMetrics.isGazetteUploaded ? "All Papers" : "Need Gazette"}
                      </div>
                    </div>

                    <div style={{ background: "rgba(16, 185, 129, 0.1)", padding: "6px 8px", borderRadius: "5px", border: "1px solid rgba(16, 185, 129, 0.3)" }}>
                      <div style={{ color: "#10b981", fontWeight: 600, fontSize: "10px" }}>Passed All</div>
                      <strong style={{ fontSize: "14px", color: "#10b981" }}>
                        {consolidatedStudentMetrics.isGazetteUploaded ? consolidatedStudentMetrics.finalPassedStudents : "—"}
                      </strong>
                    </div>

                    <div style={{ background: "rgba(239, 68, 68, 0.1)", padding: "6px 8px", borderRadius: "5px", border: "1px solid rgba(239, 68, 68, 0.3)" }}>
                      <div style={{ color: "#ef4444", fontWeight: 600, fontSize: "10px" }}>Failed &ge; 1</div>
                      <strong style={{ fontSize: "14px", color: "#ef4444" }}>
                        {consolidatedStudentMetrics.isGazetteUploaded ? consolidatedStudentMetrics.failedStudents : "—"}
                      </strong>
                    </div>
                  </div>

                  {/* Student Moderation Impact */}
                  <div style={{ background: "var(--panel)", padding: "6px 8px", borderRadius: "5px", border: "1px solid var(--line)", fontSize: "10.5px", display: "flex", flexDirection: "column", gap: "3px" }}>
                    <div style={{ fontWeight: 600, color: "var(--ink)", display: "flex", alignItems: "center", gap: "4px" }}>
                      <Zap size={11} color="#f59e0b" /> Mod Impact:
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)" }}>
                      <span>Raw Passed (0 Mod):</span>
                      <strong>{consolidatedStudentMetrics.isGazetteUploaded ? `${consolidatedStudentMetrics.rawPassedStudents} (${consolidatedStudentMetrics.rawPassedPct}%)` : "—"}</strong>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", color: "#10b981", fontWeight: 600 }}>
                      <span>Rescued to Pass:</span>
                      <span>{consolidatedStudentMetrics.isGazetteUploaded ? `+${consolidatedStudentMetrics.rescuedStudents} (${consolidatedStudentMetrics.rescuedPct}%)` : "—"}</span>
                    </div>
                  </div>

                  {/* Course Aggregate */}
                  <div style={{ background: "var(--panel)", padding: "6px 8px", borderRadius: "5px", border: "1px solid var(--line)", fontSize: "10.5px", display: "flex", flexDirection: "column", gap: "3px" }}>
                    <div style={{ fontWeight: 600, color: "var(--ink)", display: "flex", alignItems: "center", gap: "4px" }}>
                      <BookOpen size={11} color="var(--accent)" /> Course Papers:
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)" }}>
                      <span>Total Papers:</span>
                      <strong>{consolidatedCourseMetrics.total}</strong>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", color: "#10b981" }}>
                      <span>Passed Papers:</span>
                      <strong>{consolidatedCourseMetrics.totalPassed} ({consolidatedCourseMetrics.passPct}%)</strong>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", color: "#ef4444" }}>
                      <span>Failed Papers:</span>
                      <strong>{consolidatedCourseMetrics.failed} ({consolidatedCourseMetrics.failedPct}%)</strong>
                    </div>
                  </div>

                  {/* Course Pass Simulation (+0 to +10) Card */}
                  <div style={{ background: "rgba(16, 185, 129, 0.08)", padding: "8px", borderRadius: "5px", border: "1px solid rgba(16, 185, 129, 0.25)", fontSize: "10.5px", display: "flex", flexDirection: "column", gap: "5px" }}>
                    <div style={{ fontWeight: 700, color: "#10b981", display: "flex", alignItems: "center", gap: "4px" }}>
                      <TrendingUp size={12} /> Simulation (+0..+10)
                    </div>
                    <div style={{ display: "flex", gap: "4px", marginTop: "1px" }}>
                      <button
                        type="button"
                        onClick={handleExportSimulationExcel}
                        style={{
                          flex: 1,
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          gap: "3px",
                          padding: "4px 6px",
                          fontSize: "10px",
                          background: "#10b981",
                          color: "white",
                          border: "none",
                          borderRadius: "4px",
                          fontWeight: 600,
                          cursor: "pointer"
                        }}
                      >
                        <Download size={11} /> Excel
                      </button>
                      <button
                        type="button"
                        onClick={() => setActiveTab("matrix")}
                        style={{
                          flex: 1,
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          gap: "3px",
                          padding: "4px 6px",
                          fontSize: "10px",
                          background: "var(--panel)",
                          color: "var(--ink)",
                          border: "1px solid var(--line)",
                          borderRadius: "4px",
                          fontWeight: 600,
                          cursor: "pointer"
                        }}
                      >
                        Matrix
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Accordion 5: Evaluation Rules & Reference */}
          <div style={{ border: "1px solid var(--line)", borderRadius: "7px", overflow: "hidden", background: "var(--panel)" }}>
            <div 
              onClick={() => toggleAccordion("rules")}
              style={{ 
                display: "flex", 
                alignItems: "center", 
                justifyContent: "space-between", 
                padding: "7px 9px", 
                background: sidebarAccordions.rules ? "var(--bg)" : "var(--panel)", 
                cursor: "pointer", 
                userSelect: "none" 
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                <BookOpen size={13} color="var(--accent)" />
                <span style={{ fontSize: "11px", fontWeight: 700, color: "var(--ink)" }}>Evaluation Rules</span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                <span style={{ fontSize: "9px", fontWeight: 600, padding: "1px 5px", borderRadius: "8px", background: "var(--panel)", color: "var(--muted)", border: "1px solid var(--line)" }}>
                  Ref
                </span>
                {sidebarAccordions.rules ? <ChevronDown size={12} color="var(--muted)" /> : <ChevronRight size={12} color="var(--muted)" />}
              </div>
            </div>

            {sidebarAccordions.rules && (
              <div style={{ padding: "9px", fontSize: "10.5px", color: "var(--muted)", lineHeight: "1.35", borderTop: "1px solid var(--line)" }}>
                <ul style={{ margin: 0, paddingLeft: "14px", display: "flex", flexDirection: "column", gap: "3px" }}>
                  <li><strong>ESE Min:</strong> <code>ceil(30% × ESE Max)</code></li>
                  <li><strong>Overall Min:</strong> <code>ceil(35% × Overall Max)</code></li>
                  <li><strong>Pass Condition:</strong> <code>ESE Pass AND Overall Pass</code></li>
                  <li><strong>Moderation Rule:</strong> <code>max(ESE Deficit, Overall Deficit)</code></li>
                  <li><strong>ESE-TH Only:</strong> Moderation applies to courses with ESE-TH.</li>
                </ul>
              </div>
            )}
          </div>

          {/* Status Message */}
          <div style={{ marginTop: "auto", padding: "6px 10px", borderRadius: "5px", fontSize: "10.5px", background: statusType === "error" ? "var(--danger-soft)" : statusType === "success" ? "var(--accent-soft)" : "var(--bg)", color: statusType === "error" ? "var(--danger)" : statusType === "success" ? "var(--accent)" : "var(--muted)", border: "1px solid var(--line)" }}>
            {statusMsg}
          </div>

        </aside>

        {/* Sidebar Drag Resizer Handle (MacBook Air / Trackpad) */}
        {!isSidebarCollapsed && (
          <div
            onMouseDown={handleSidebarMouseDown}
            onDoubleClick={() => setSidebarWidth(250)}
            title="Drag to resize sidebar (double-click to reset to 250px)"
            style={{
              width: "5px",
              cursor: "col-resize",
              flexShrink: 0,
              background: "transparent",
              zIndex: 20,
              transition: "background 0.15s ease",
              userSelect: "none"
            }}
            onMouseEnter={(e) => e.currentTarget.style.background = "rgba(23, 107, 135, 0.4)"}
            onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}
          />
        )}


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

              {/* Gazette Not Uploaded Notice */}
              {!studentMetrics.isGazetteUploaded && (
                <div style={{ background: "rgba(99, 102, 241, 0.08)", border: "1px solid rgba(99, 102, 241, 0.25)", borderRadius: "8px", padding: "10px 14px", fontSize: "12px", color: "var(--ink)", display: "flex", alignItems: "center", gap: "10px" }}>
                  <HelpCircle size={18} color="#6366f1" style={{ flexShrink: 0 }} />
                  <div>
                    <strong style={{ color: "#6366f1" }}>Semester Pass/Fail is Blank:</strong> In supplementary examinations, candidates often attempt only a subset of uncleared courses. Semester-level pass/fail is only evaluated when the <strong>Result Gazette (Reappear Assurance)</strong> is uploaded to cross-check all historical backlogs.
                  </div>
                </div>
              )}

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
                    {(studentMetrics.totalImprovementPapers > 0 || studentMetrics.totalSupplementaryPapers > 0) && (
                      <div style={{ marginTop: "3px", fontSize: "10px", display: "flex", gap: "6px", alignItems: "center" }}>
                        {studentMetrics.totalImprovementPapers > 0 && (
                          <span style={{ color: "#9333ea", fontWeight: 700, background: "rgba(147, 51, 234, 0.1)", padding: "1px 4px", borderRadius: "3px" }}>
                            {studentMetrics.totalImprovementPapers} Imp
                          </span>
                        )}
                        {studentMetrics.totalSupplementaryPapers > 0 && (
                          <span style={{ color: "#2563eb", fontWeight: 700, background: "rgba(59, 130, 246, 0.1)", padding: "1px 4px", borderRadius: "3px" }}>
                            {studentMetrics.totalSupplementaryPapers} Supp
                          </span>
                        )}
                      </div>
                    )}
                  </div>
                </div>

                <div style={{ background: "rgba(16, 185, 129, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(16, 185, 129, 0.3)" }}>
                  <div style={{ fontSize: "11px", color: "#10b981", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                    <CheckCircle2 size={14} /> Passed Semester (All Papers)
                  </div>
                  <div style={{ fontSize: "20px", fontWeight: 700, color: "#10b981", marginTop: "3px" }}>
                    {studentMetrics.isGazetteUploaded ? studentMetrics.finalPassedStudents : "—"}
                  </div>
                  <div style={{ fontSize: "10.5px", color: "#10b981", fontWeight: 600 }}>
                    {studentMetrics.isGazetteUploaded ? `${studentMetrics.finalPassedPct}% of total students` : "Requires Result Gazette"}
                  </div>
                </div>

                <div style={{ background: "rgba(239, 68, 68, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(239, 68, 68, 0.3)" }}>
                  <div style={{ fontSize: "11px", color: "#ef4444", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                    <XCircle size={14} /> Failed Semester (≥1 Paper)
                  </div>
                  <div style={{ fontSize: "20px", fontWeight: 700, color: "#ef4444", marginTop: "3px" }}>
                    {studentMetrics.isGazetteUploaded ? studentMetrics.failedStudents : "—"}
                  </div>
                  <div style={{ fontSize: "10.5px", color: "#ef4444", fontWeight: 600 }}>
                    {studentMetrics.isGazetteUploaded ? `${studentMetrics.failedPct}% of total students` : "Requires Result Gazette"}
                  </div>
                </div>

                <div style={{ background: "rgba(245, 158, 11, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(245, 158, 11, 0.3)" }}>
                  <div style={{ fontSize: "11px", color: "#f59e0b", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                    <Sparkles size={14} /> Rescued to Pass with Mod
                  </div>
                  <div style={{ fontSize: "20px", fontWeight: 700, color: "#f59e0b", marginTop: "3px" }}>
                    {studentMetrics.isGazetteUploaded ? `+${studentMetrics.rescuedStudents}` : "—"}
                  </div>
                  <div style={{ fontSize: "10.5px", color: "#f59e0b", fontWeight: 600 }}>
                    {studentMetrics.isGazetteUploaded ? `${studentMetrics.rescuedPct}% students rescued to pass` : "Requires Result Gazette"}
                  </div>
                </div>

                <div style={{ background: "var(--panel)", padding: "12px 14px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                  <div style={{ fontSize: "11px", color: "var(--muted)", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                    <Calculator size={14} color="var(--muted)" /> Raw Passed (0 Mod)
                  </div>
                  <div style={{ fontSize: "20px", fontWeight: 700, color: "var(--ink)", marginTop: "3px" }}>
                    {studentMetrics.isGazetteUploaded ? studentMetrics.rawPassedStudents : "—"}
                  </div>
                  <div style={{ fontSize: "10.5px", color: "var(--muted)" }}>
                    {studentMetrics.isGazetteUploaded ? `${studentMetrics.rawPassedPct}% without moderation` : "Requires Result Gazette"}
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
                  {studentMetrics.isGazetteUploaded && (
                    <>
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
                    </>
                  )}
                  {studentMetrics.pendingBacklogStudents > 0 && (
                    <button 
                      type="button"
                      onClick={() => setStudentFilterStatus("PENDING_BACKLOG")}
                      style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "PENDING_BACKLOG" ? "#dc2626" : "transparent", color: studentFilterStatus === "PENDING_BACKLOG" ? "white" : "#dc2626" }}
                    >
                      Pending Gazette Backlogs ({studentMetrics.pendingBacklogStudents})
                    </button>
                  )}
                  {studentMetrics.isGazetteUploaded && (
                    <button 
                      type="button"
                      onClick={() => setStudentFilterStatus("RESCUED")}
                      style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "RESCUED" ? "#f59e0b" : "transparent", color: studentFilterStatus === "RESCUED" ? "white" : "var(--muted)" }}
                    >
                      Rescued ({studentMetrics.rescuedStudents})
                    </button>
                  )}
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
                  {studentMetrics.improvementStudents > 0 && (
                    <button 
                      type="button"
                      onClick={() => setStudentFilterStatus("IMPROVEMENT")}
                      style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "IMPROVEMENT" ? "#9333ea" : "transparent", color: studentFilterStatus === "IMPROVEMENT" ? "white" : "#9333ea" }}
                    >
                      Improvement ({studentMetrics.improvementStudents})
                    </button>
                  )}
                  {studentMetrics.supplementaryStudents > 0 && (
                    <button 
                      type="button"
                      onClick={() => setStudentFilterStatus("SUPPLEMENTARY")}
                      style={{ padding: "4px 10px", fontSize: "11.5px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: studentFilterStatus === "SUPPLEMENTARY" ? "#2563eb" : "transparent", color: studentFilterStatus === "SUPPLEMENTARY" ? "white" : "#2563eb" }}
                    >
                      Supplementary ({studentMetrics.supplementaryStudents})
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
                      <th style={{ padding: "10px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center", width: "125px" }} title="Our Mod Total = Regular Mod (from Gazette) + Supplementary Event Mod">Our Mod Total</th>
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
                                <div>{st.totalCourses} Papers</div>
                                {(st.improvementCourses > 0 || st.supplementaryCourses > 0) && (
                                  <div style={{ display: "flex", gap: "3px", justifyContent: "center", marginTop: "3px", flexWrap: "wrap" }}>
                                    {st.improvementCourses > 0 && (
                                      <span style={{ fontSize: "9.5px", background: "rgba(147, 51, 234, 0.12)", color: "#9333ea", padding: "1px 5px", borderRadius: "3px", fontWeight: 700 }} title={`${st.improvementCourses} Improvement paper(s)`}>
                                        {st.improvementCourses} Imp
                                      </span>
                                    )}
                                    {st.supplementaryCourses > 0 && (
                                      <span style={{ fontSize: "9.5px", background: "rgba(59, 130, 246, 0.12)", color: "#2563eb", padding: "1px 5px", borderRadius: "3px", fontWeight: 700 }} title={`${st.supplementaryCourses} Supplementary paper(s)`}>
                                        {st.supplementaryCourses} Supp
                                      </span>
                                    )}
                                  </div>
                                )}
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center", color: "#10b981", fontWeight: 600 }}>
                                {st.finalPassedCourses}
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center", color: st.finalFailedCourses > 0 ? "#ef4444" : "var(--muted)", fontWeight: st.finalFailedCourses > 0 ? 600 : 400 }}>
                                {st.finalFailedCourses}
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center" }}>
                                {st.totalModerationMarks > 0 ? (
                                  <div style={{ display: "inline-flex", flexDirection: "column", alignItems: "center" }}>
                                    <span 
                                      style={{ fontSize: "11px", fontWeight: 700, color: "#f59e0b", background: "rgba(245, 158, 11, 0.12)", padding: "2px 6px", borderRadius: "4px" }}
                                      title={st.regularModerationMarks > 0 
                                        ? `Our Mod Total: +${st.totalModerationMarks} (Supp Event: +${st.eventModerationMarks || 0}, Regular Event: +${st.regularModerationMarks})` 
                                        : `Supplementary Event Moderation: +${st.totalModerationMarks}`}
                                    >
                                      +{st.totalModerationMarks} Marks
                                    </span>
                                    {st.regularModerationMarks > 0 && (
                                      <span style={{ fontSize: "9.5px", color: "var(--muted)", marginTop: "2px" }} title="Breakdown: Supp Event Mod + Regular Event Mod from Gazette">
                                        Event: +{st.eventModerationMarks || 0} | Reg: +{st.regularModerationMarks}
                                      </span>
                                    )}
                                  </div>
                                ) : (
                                  <span style={{ color: "var(--muted)", fontSize: "11px" }}>0</span>
                                )}
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center" }}>
                                <span style={{ 
                                  fontSize: "11px", 
                                  fontWeight: 600, 
                                  color: !st.rawSemesterResult ? "var(--muted)" : st.isHeldback ? "#c026d3" : st.isHeld ? "#7e22ce" : st.rawSemesterPass ? "#10b981" : "#ef4444" 
                                }}>
                                  {st.rawSemesterResult || "—"}
                                </span>
                              </td>
                              <td style={{ padding: "10px 12px", textAlign: "center" }}>
                                {!st.isGazetteUploaded ? (
                                  <span style={{ color: "var(--muted)", fontSize: "13px" }} title="Upload Result Gazette (Reappear Assurance) to evaluate semester pass/fail">
                                    —
                                  </span>
                                ) : st.isHeldback ? (
                                  st.isHeld ? (
                                    <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(239, 68, 68, 0.12)", color: "#ef4444", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }} title={`Held due to missing component(s): ${st.missingDetailStr || "Incomplete"} | Reason: ${st.heldbackReason || "Heldback"}`}>
                                      <XCircle size={12} /> Fail <span style={{ fontSize: "10px", color: "#e11d48", fontWeight: 700 }}>(Held - Missing{st.missingComponents && st.missingComponents.length > 0 ? `: ${st.missingComponents.join(", ")}` : ""})</span> <span style={{ fontSize: "10px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                    </span>
                                  ) : (st.pendingGazetteBacklogs && st.pendingGazetteBacklogs.length > 0) ? (
                                    <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(239, 68, 68, 0.12)", color: "#ef4444", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }} title={`Pending Gazette Backlogs: ${st.pendingGazetteBacklogs.join(", ")} | Reason: ${st.heldbackReason || "Heldback"}`}>
                                      <XCircle size={12} /> Fail <span style={{ fontSize: "10px", color: "#dc2626", fontWeight: 700 }}>({st.pendingGazetteBacklogs.length} Pending Backlog{st.pendingGazetteBacklogs.length > 1 ? "s" : ""})</span> <span style={{ fontSize: "10px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                    </span>
                                  ) : st.isRescuedSemester ? (
                                    <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(245, 158, 11, 0.15)", color: "#d97706", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }} title={`Rescued by moderation | Heldback: ${st.heldbackReason || "Heldback"}`}>
                                      <Sparkles size={12} /> Pass (Rescued) <span style={{ fontSize: "10px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                    </span>
                                  ) : st.finalSemesterPass ? (
                                    <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(16, 185, 129, 0.15)", color: "#10b981", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }} title={`Heldback: ${st.heldbackReason || "Heldback"}`}>
                                      <CheckCircle2 size={12} /> Pass <span style={{ fontSize: "10px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                    </span>
                                  ) : (
                                    <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(239, 68, 68, 0.12)", color: "#ef4444", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }} title={`Heldback: ${st.heldbackReason || "Heldback"}`}>
                                      <XCircle size={12} /> Fail ({st.finalFailedCourses} Failed) <span style={{ fontSize: "10px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                    </span>
                                  )
                                ) : st.isHeld ? (
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(239, 68, 68, 0.12)", color: "#ef4444", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }} title={`Failed due to missing component(s): ${st.missingDetailStr || "Marks Incomplete"}`}>
                                    <XCircle size={12} /> Fail <span style={{ fontSize: "10px", color: "#e11d48", fontWeight: 700 }}>(Held - Missing{st.missingComponents && st.missingComponents.length > 0 ? `: ${st.missingComponents.join(", ")}` : ""})</span>
                                  </span>
                                ) : (st.pendingGazetteBacklogs && st.pendingGazetteBacklogs.length > 0) ? (
                                  <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(239, 68, 68, 0.12)", color: "#ef4444", padding: "3px 8px", borderRadius: "10px", fontWeight: 700, fontSize: "11.5px" }} title={`Semester Incomplete: ${st.pendingGazetteBacklogs.length} pending backlog(s) in Gazette (${st.pendingGazetteBacklogs.join(", ")})${st.finalFailedCourses > 0 ? ` + ${st.finalFailedCourses} failed in current exam` : ""}`}>
                                    <XCircle size={12} /> Fail <span style={{ fontSize: "10px", color: "#dc2626", fontWeight: 700 }}>({st.pendingGazetteBacklogs.length} Pending Backlog{st.pendingGazetteBacklogs.length > 1 ? "s" : ""}{st.finalFailedCourses > 0 ? `, ${st.finalFailedCourses} failed` : ""})</span>
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
                                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "6px", padding: "6px 12px", background: "var(--bg)", fontSize: "11px", fontWeight: 700, color: "var(--muted)", borderBottom: "1px solid var(--line)" }}>
                                      <span>COURSE-BY-COURSE BREAKDOWN FOR {st.seatNumber || st.prn} ({st.courses.length} Attempted Papers{st.improvementCourses > 0 ? `: ${st.improvementCourses} Improvement` : ''}{st.supplementaryCourses > 0 ? `${st.improvementCourses > 0 ? ', ' : ': '}${st.supplementaryCourses} Supplementary` : ''})</span>
                                      {st.pendingGazetteBacklogs && st.pendingGazetteBacklogs.length > 0 && (
                                        <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(239, 68, 68, 0.1)", color: "#dc2626", padding: "2px 8px", borderRadius: "4px", fontSize: "10.5px", fontWeight: 700 }} title="Semester result cannot be Pass until all backlogs declared in the Result Gazette are cleared">
                                          <AlertCircle size={12} /> Outstanding Backlogs in Gazette ({st.pendingGazetteBacklogs.length}): {st.pendingGazetteBacklogs.join(", ")}
                                        </span>
                                      )}
                                      {st.hasGazetteAssurance && (!st.pendingGazetteBacklogs || st.pendingGazetteBacklogs.length === 0) && (
                                        <span style={{ display: "inline-flex", alignItems: "center", gap: "4px", background: "rgba(16, 185, 129, 0.1)", color: "#059669", padding: "2px 8px", borderRadius: "4px", fontSize: "10.5px", fontWeight: 700 }}>
                                          <CheckCircle2 size={12} /> All Gazette Backlog Requirements Satisfied
                                        </span>
                                      )}
                                    </div>
                                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "11.5px" }}>
                                      <thead>
                                        <tr style={{ borderBottom: "1px solid var(--line)", background: "rgba(0,0,0,0.02)" }}>
                                          <th style={{ padding: "6px 10px", textAlign: "left", color: "var(--muted)" }}>Course Code</th>
                                          <th style={{ padding: "6px 10px", textAlign: "left", color: "var(--muted)" }}>Course Name</th>
                                          <th style={{ padding: "6px 8px", textAlign: "center", color: "var(--muted)" }}>Attempt</th>
                                          <th style={{ padding: "6px 8px", textAlign: "center", color: "var(--muted)" }}>ESE-TH</th>
                                          <th style={{ padding: "6px 8px", textAlign: "center", color: "var(--muted)" }}>CE-TH</th>
                                          <th style={{ padding: "6px 8px", textAlign: "center", color: "var(--muted)" }}>ESE-PR</th>
                                          <th style={{ padding: "6px 8px", textAlign: "center", color: "var(--muted)" }}>CE-PR</th>
                                          <th style={{ padding: "6px 8px", textAlign: "center", color: "var(--muted)" }}>CE Total</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>ESE (Marks / Min)</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>Overall (Marks / Min)</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>ESE Pass</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>Overall Pass</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>Mod Awarded</th>
                                          <th style={{ padding: "6px 10px", textAlign: "center", color: "var(--muted)" }}>Course Result</th>
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {st.courses.map((c, cIdx) => {
                                          const renderCompCell = (obt, max, compKey) => {
                                            const isMissing = (c.missingComponents || []).includes(compKey);
                                            if (isMissing) {
                                              return (
                                                <span style={{ color: "#7e22ce", fontWeight: 700, fontSize: "10.5px", background: "rgba(147, 51, 234, 0.12)", padding: "1px 5px", borderRadius: "3px" }}>
                                                  Missing
                                                </span>
                                              );
                                            }
                                            if (c.isAbsent && (obt === "Absent (Ab)" || (c.absentStatus && !max))) {
                                              return <span style={{ color: "#dc2626", fontWeight: 600, fontSize: "10.5px" }}>Ab</span>;
                                            }
                                            if (c.isMalpractice && (obt === "Malpractice (MP)" || c.malpracticeStatus)) {
                                              return <span style={{ color: "#b45309", fontWeight: 600, fontSize: "10.5px" }}>MP</span>;
                                            }
                                            const hasMax = max !== null && max !== undefined && max !== "" && max !== 0;
                                            const hasObt = obt !== null && obt !== undefined && obt !== "";
                                            if (!hasMax && !hasObt) {
                                              return <span style={{ color: "var(--muted)" }}>-</span>;
                                            }
                                            const isCF = (c.carriedForwardComponents || []).includes(compKey);
                                            return (
                                              <span style={{ display: "inline-flex", alignItems: "center", gap: "3px", justifyContent: "center" }}>
                                                <span style={{ color: isCF ? "#059669" : "inherit", fontWeight: isCF ? 600 : 400 }}>
                                                  {hasObt ? String(obt) : "0"}{hasMax ? ` / ${max}` : ""}
                                                </span>
                                                {isCF && (
                                                  <span style={{ fontSize: "8.5px", background: "rgba(16, 185, 129, 0.15)", color: "#059669", padding: "0 3px", borderRadius: "3px", fontWeight: 700 }} title="Carried forward from previous attempt">
                                                    CF
                                                  </span>
                                                )}
                                              </span>
                                            );
                                          };

                                          const renderCeTotalCell = () => {
                                            const hasCeMissing = (c.missingComponents || []).some(m => m.startsWith("CE"));
                                            if (c.isHeld && (c.ceOverall === "Held" || hasCeMissing)) {
                                              return (
                                                <span style={{ color: "#7e22ce", fontWeight: 700, fontSize: "10.5px", background: "rgba(147, 51, 234, 0.12)", padding: "1px 5px", borderRadius: "3px" }}>
                                                  Held
                                                </span>
                                              );
                                            }
                                            const hasMax = c.ceMax !== null && c.ceMax !== undefined && c.ceMax !== "" && c.ceMax !== 0;
                                            const hasObt = c.ceOverall !== null && c.ceOverall !== undefined && c.ceOverall !== "";
                                            if (!hasMax && !hasObt) {
                                              return <span style={{ color: "var(--muted)" }}>-</span>;
                                            }
                                            const allCeCF = (c.carriedForwardComponents || []).some(comp => comp.startsWith("CE"));
                                            return (
                                              <span style={{ display: "inline-flex", alignItems: "center", gap: "3px", justifyContent: "center" }}>
                                                <span style={{ color: allCeCF ? "#059669" : "inherit", fontWeight: allCeCF ? 600 : 500 }}>
                                                  {hasObt ? String(c.ceOverall) : "0"}{hasMax ? ` / ${c.ceMax}` : ""}
                                                </span>
                                                {allCeCF && (
                                                  <span style={{ fontSize: "8.5px", background: "rgba(16, 185, 129, 0.15)", color: "#059669", padding: "0 3px", borderRadius: "3px", fontWeight: 700 }} title="Carried forward from previous attempt">
                                                    CF
                                                  </span>
                                                )}
                                              </span>
                                            );
                                          };

                                          return (
                                            <tr key={cIdx} style={{ borderBottom: "1px solid var(--line)", background: c.coursePass === "Pass" ? "transparent" : c.isHeldback ? "rgba(192, 38, 211, 0.05)" : c.isHeld ? "rgba(147, 51, 234, 0.04)" : "rgba(239, 68, 68, 0.03)" }}>
                                              <td style={{ padding: "6px 10px" }}>
                                                <div style={{ fontWeight: 600 }}>{c.courseCode}</div>
                                                {c.gazetteAssurance && c.gazetteAssurance.isVerifiedReappear && (
                                                  <div style={{ marginTop: "2px" }}>
                                                    <span style={{ fontSize: "9.5px", background: "rgba(16, 185, 129, 0.12)", color: "#059669", padding: "1px 5px", borderRadius: "3px", fontWeight: 700 }} title={`Officially declared reappear paper in ${c.gazetteAssurance.matchingTerm || 'previous semester'} (${c.gazetteAssurance.matchingEvent || 'Gazette'})`}>
                                                      ✓ Verified Reappear ({c.gazetteAssurance.matchingTerm || 'Gazette'})
                                                    </span>
                                                  </div>
                                                )}
                                                {c.gazetteAssurance && c.gazetteAssurance.isPriorPass && (
                                                  <div style={{ marginTop: "2px" }}>
                                                    <span style={{ fontSize: "9.5px", background: "rgba(37, 99, 235, 0.12)", color: "#2563eb", padding: "1px 5px", borderRadius: "3px", fontWeight: 700 }} title="Passed in previous semester (Improvement Attempt)">
                                                      ✓ Improvement ({c.gazetteAssurance.matchingTerm || 'Prior Pass'})
                                                    </span>
                                                  </div>
                                                )}
                                                {c.gazetteAssurance && !c.gazetteAssurance.isVerifiedReappear && !c.gazetteAssurance.isPriorPass && (
                                                  <div style={{ marginTop: "2px" }}>
                                                    <span style={{ fontSize: "9.5px", background: "rgba(245, 158, 11, 0.12)", color: "#b45309", padding: "1px 5px", borderRadius: "3px", fontWeight: 600 }} title={`Not listed in previous fail list (${c.gazetteAssurance.matchingTerm || 'Gazette'})`}>
                                                      ⚠ Unlisted in Gazette Reappears
                                                    </span>
                                                  </div>
                                                )}
                                                {c.gazetteAssurance && c.gazetteAssurance.allPendingReappears?.length > 0 && (
                                                  <div style={{ fontSize: "8.5px", color: "#6b7280", marginTop: "2px" }} title={`Other pending failed papers: ${c.gazetteAssurance.allPendingReappears.join(", ")}`}>
                                                    Other pending: {c.gazetteAssurance.allPendingReappears.join(", ")}
                                                  </div>
                                                )}
                                              </td>
                                              <td style={{ padding: "6px 10px" }}>{c.courseName}</td>
                                              <td style={{ padding: "6px 8px", textAlign: "center" }}>
                                                {(c.isImprovement || c.attemptType === "IMPROVEMENT") ? (
                                                  <span style={{ 
                                                    display: "inline-flex", 
                                                    alignItems: "center", 
                                                    gap: "3px", 
                                                    background: "rgba(147, 51, 234, 0.12)", 
                                                    color: "#9333ea", 
                                                    padding: "2px 7px", 
                                                    borderRadius: "4px", 
                                                    fontSize: "10.5px", 
                                                    fontWeight: 700 
                                                  }} title="Improvement Attempt (Student previously passed and is retaking to improve score)">
                                                    Improvement
                                                  </span>
                                                ) : (
                                                  <span style={{ 
                                                    display: "inline-flex", 
                                                    alignItems: "center", 
                                                    gap: "3px", 
                                                    background: "rgba(59, 130, 246, 0.12)", 
                                                    color: "#2563eb", 
                                                    padding: "2px 7px", 
                                                    borderRadius: "4px", 
                                                    fontSize: "10.5px", 
                                                    fontWeight: 700 
                                                  }} title="Supplementary Attempt (Student previously failed/absent and is retaking to clear backlog)">
                                                    Supplementary
                                                  </span>
                                                )}
                                              </td>
                                              <td style={{ padding: "6px 8px", textAlign: "center" }}>{renderCompCell(c.eseThObtained, c.eseThMax, "ESE-TH")}</td>
                                              <td style={{ padding: "6px 8px", textAlign: "center" }}>{renderCompCell(c.ceThObtained, c.ceThMax, "CE-TH")}</td>
                                              <td style={{ padding: "6px 8px", textAlign: "center" }}>{renderCompCell(c.esePrObtained, c.esePrMax, "ESE-PR")}</td>
                                              <td style={{ padding: "6px 8px", textAlign: "center" }}>{renderCompCell(c.cePrObtained, c.cePrMax, "CE-PR")}</td>
                                              <td style={{ padding: "6px 8px", textAlign: "center" }}>{renderCeTotalCell()}</td>
                                              <td style={{ padding: "6px 10px", textAlign: "center" }}>
                                              {c.isHeld ? (
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
                                              {c.isHeld ? (
                                                <span style={{ color: "#e11d48", fontWeight: 700, background: "rgba(239, 68, 68, 0.12)", padding: "1px 6px", borderRadius: "4px", fontSize: "10.5px" }} title={`Missing component(s): ${c.missingComponents?.join(", ") || "Component"}`}>
                                                  Held (Missing{c.missingComponents?.length ? `: ${c.missingComponents.join(", ")}` : ""})
                                                </span>
                                              ) : c.courseOverall !== null ? (
                                                `${c.courseOverall} / ${c.overallMin ?? "-"}`
                                              ) : "-"}
                                            </td>
                                            <td style={{ padding: "6px 10px", textAlign: "center", fontWeight: 600, color: c.esePass === "Pass" ? "#10b981" : c.isHeld ? "#e11d48" : "#ef4444" }}>
                                              {c.esePass || "-"}
                                            </td>
                                            <td style={{ padding: "6px 10px", textAlign: "center", fontWeight: 600, color: c.overallPass === "Pass" ? "#10b981" : c.isHeld ? "#e11d48" : "#ef4444" }}>
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
                                              {c.isHeldback && c.isHeld ? (
                                                <span style={{ display: "inline-flex", flexDirection: "column", alignItems: "center", gap: "2px" }}>
                                                  <span style={{ color: "#ef4444", fontWeight: 700, fontSize: "11px", background: "rgba(239, 68, 68, 0.12)", padding: "1px 6px", borderRadius: "4px" }} title={`Course Failed: Missing component(s) ${c.missingComponents?.join(", ") || "Component"} | Reason: ${c.heldbackReason || "Heldback"}`}>
                                                    Fail <span style={{ color: "#e11d48", fontWeight: 700 }}>(Held - Missing{c.missingComponents?.length ? `: ${c.missingComponents.join(", ")}` : ""})</span>
                                                  </span>
                                                  <span style={{ color: "#a855f7", fontWeight: 700, fontSize: "10px" }} title={c.heldbackReason || "Heldback record"}>
                                                    (Heldback)
                                                  </span>
                                                </span>
                                              ) : c.isHeldback ? (
                                                <span style={{ display: "inline-flex", flexDirection: "column", alignItems: "center", gap: "2px" }}>
                                                  {c.rawPass ? (
                                                    <span style={{ color: "#10b981", fontWeight: 600, fontSize: "11px", background: "rgba(16, 185, 129, 0.12)", padding: "1px 6px", borderRadius: "4px" }} title="Pass without moderation | Heldback">
                                                      Pass <span style={{ color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                                    </span>
                                                  ) : c.isModeratedPass ? (
                                                    <>
                                                      <span style={{ color: "#ef4444", fontSize: "10px", fontWeight: 500 }} title="Failed without moderation">
                                                        Raw: Fail (No Mod)
                                                      </span>
                                                      <span style={{ color: "#d97706", fontWeight: 700, fontSize: "11px", background: "rgba(245, 158, 11, 0.15)", padding: "1px 6px", borderRadius: "4px" }} title={`Pass with +${c.modMarks} moderation marks | Heldback`}>
                                                        Pass (Mod +{c.modMarks}) <span style={{ color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                                      </span>
                                                    </>
                                                  ) : (
                                                    <span style={{ color: "#ef4444", fontWeight: 600, fontSize: "11px", background: "rgba(239, 68, 68, 0.12)", padding: "1px 6px", borderRadius: "4px" }} title="Failed both raw and with moderation | Heldback">
                                                      Fail <span style={{ color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                                    </span>
                                                  )}
                                                </span>
                                              ) : c.isHeld ? (
                                                <span style={{ color: "#ef4444", fontWeight: 700, background: "rgba(239, 68, 68, 0.12)", padding: "2px 6px", borderRadius: "4px", display: "inline-flex", alignItems: "center", gap: "3px", fontSize: "11px" }} title={`Course Failed: Missing component(s) ${c.missingComponents?.join(", ") || "Component"}`}>
                                                  <XCircle size={11} /> Fail <span style={{ color: "#e11d48", fontWeight: 700 }}>(Held - Missing{c.missingComponents?.length ? `: ${c.missingComponents.join(", ")}` : ""})</span>
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
                                         );
                                       })}
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
          ) : activeTab === "comparison" ? (
            /* Result & Ordinance Comparison Panel */
            <div style={{ display: "flex", flexDirection: "column", height: "100%", overflowY: "auto", padding: "16px", gap: "16px", background: "var(--bg)" }}>
              {/* File Upload / Status Card */}
              {rawComparisonRows.length === 0 ? (
                <div 
                  style={{ 
                    border: "2px dashed var(--line)", 
                    borderRadius: "12px", 
                    padding: "40px 24px", 
                    textAlign: "center", 
                    background: "var(--panel)",
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    justifyContent: "center",
                    gap: "12px",
                    maxWidth: "700px",
                    margin: "40px auto",
                    width: "100%"
                  }}
                >
                  <div style={{ width: "56px", height: "56px", borderRadius: "50%", background: "rgba(139, 92, 246, 0.1)", display: "flex", alignItems: "center", justifyContent: "center", color: "#8b5cf6" }}>
                    <Scale size={28} />
                  </div>
                  <div>
                    <h3 style={{ margin: "0 0 6px", fontSize: "16px", fontWeight: 700, color: "var(--ink)" }}>Upload University Published Result / Ordinance Summary</h3>
                    <p style={{ margin: 0, fontSize: "12.5px", color: "var(--muted)", maxWidth: "480px", lineHeight: "1.5" }}>
                      Upload the official university published Excel file (e.g. <code>Student DataIV Sem(R)...xlsx</code>) containing <strong>RESULT STATUS</strong>, <strong>Ordinance</strong>, and <strong>Ord Total</strong> to automatically reconcile results and moderation marks against our calculations.
                    </p>
                  </div>
                  <div style={{ marginTop: "8px" }}>
                    <label 
                      htmlFor="comparison-file-upload-hero"
                      style={{ 
                        display: "inline-flex", 
                        alignItems: "center", 
                        gap: "8px", 
                        padding: "8px 18px", 
                        background: "#8b5cf6", 
                        color: "white", 
                        borderRadius: "8px", 
                        fontWeight: 600, 
                        fontSize: "13px", 
                        cursor: "pointer",
                        boxShadow: "0 2px 4px rgba(139, 92, 246, 0.2)"
                      }}
                    >
                      <Upload size={16} /> Choose Result Summary Excel File
                    </label>
                    <input 
                      id="comparison-file-upload-hero"
                      type="file" 
                      accept=".xlsx,.xls,.csv" 
                      onChange={handleComparisonFileUpload} 
                      style={{ display: "none" }} 
                    />
                  </div>
                </div>
              ) : (
                /* Comparison Controls & KPIs */
                <>
                  {/* File Metadata Bar */}
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "var(--panel)", padding: "10px 16px", borderRadius: "8px", border: "1px solid var(--line)", flexWrap: "wrap", gap: "10px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "10px", flexWrap: "wrap" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: "6px", fontWeight: 600, fontSize: "13px", color: "var(--ink)" }}>
                        <FileText size={16} color="#8b5cf6" />
                        <span>{comparisonFileName}</span>
                      </div>
                      <span style={{ fontSize: "11.5px", background: "rgba(139, 92, 246, 0.1)", color: "#8b5cf6", padding: "2px 8px", borderRadius: "10px", fontWeight: 600 }}>
                        {comparisonRecords.length} Student Records
                      </span>

                      {/* Sheet Selector if multiple sheets */}
                      {comparisonSheetNames.length > 1 && (
                        <div style={{ display: "flex", alignItems: "center", gap: "6px", fontSize: "12px", color: "var(--muted)" }}>
                          <span>Sheet:</span>
                          <select 
                            value={comparisonSelectedSheet} 
                            onChange={(e) => parseComparisonSheet(comparisonWorkbook, e.target.value)}
                            style={{ padding: "3px 8px", fontSize: "11.5px", borderRadius: "4px", border: "1px solid var(--line)", background: "var(--bg)" }}
                          >
                            {comparisonSheetNames.map(s => <option key={s} value={s}>{s}</option>)}
                          </select>
                        </div>
                      )}
                    </div>

                    <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                      <label 
                        htmlFor="comparison-file-upload-btn"
                        style={{ 
                          display: "inline-flex", 
                          alignItems: "center", 
                          gap: "5px", 
                          padding: "5px 10px", 
                          background: "var(--bg)", 
                          color: "var(--ink)", 
                          border: "1px solid var(--line)", 
                          borderRadius: "6px", 
                          fontSize: "11.5px", 
                          fontWeight: 600, 
                          cursor: "pointer" 
                        }}
                      >
                        <RefreshCw size={13} /> Change File
                      </label>
                      <input 
                        id="comparison-file-upload-btn"
                        type="file" 
                        accept=".xlsx,.xls,.csv" 
                        onChange={handleComparisonFileUpload} 
                        style={{ display: "none" }} 
                      />
                      <button 
                        type="button" 
                        onClick={handleClearComparisonFile}
                        style={{ 
                          display: "flex", 
                          alignItems: "center", 
                          gap: "4px", 
                          padding: "5px 10px", 
                          background: "rgba(239, 68, 68, 0.1)", 
                          color: "#ef4444", 
                          border: "1px solid rgba(239, 68, 68, 0.2)", 
                          borderRadius: "6px", 
                          fontSize: "11.5px", 
                          fontWeight: 600, 
                          cursor: "pointer" 
                        }}
                      >
                        <X size={13} /> Clear
                      </button>
                    </div>
                  </div>

                  {/* 5 KPI Metric Cards */}
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: "10px" }}>
                    {/* Card 1: Total Published */}
                    <div style={{ background: "var(--panel)", padding: "12px 14px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                      <div style={{ fontSize: "11px", color: "var(--muted)", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                        <Users size={14} /> Total Compared
                      </div>
                      <div style={{ fontSize: "20px", fontWeight: 700, color: "var(--ink)", marginTop: "3px" }}>
                        {comparisonKPIs.total}
                      </div>
                      <div style={{ fontSize: "10.5px", color: "var(--muted)" }}>
                        {comparisonKPIs.matchedStudents} matched to active dataset ({comparisonKPIs.matchRate}%)
                      </div>
                    </div>

                    {/* Card 2: Result Concordance */}
                    <div style={{ background: "rgba(16, 185, 129, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(16, 185, 129, 0.3)" }}>
                      <div style={{ fontSize: "11px", color: "#10b981", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                        <CheckCircle2 size={14} /> Result Concordance
                      </div>
                      <div style={{ fontSize: "20px", fontWeight: 700, color: "#10b981", marginTop: "3px" }}>
                        {comparisonKPIs.resultMatchRate}%
                      </div>
                      <div style={{ fontSize: "10.5px", color: "#10b981", fontWeight: 600 }}>
                        {comparisonKPIs.resultMatches} identical pass/fail/held decisions
                      </div>
                    </div>

                    {/* Card 3: Result Mismatches */}
                    <div style={{ background: comparisonKPIs.resultMismatches > 0 ? "rgba(239, 68, 68, 0.08)" : "var(--panel)", padding: "12px 14px", borderRadius: "8px", border: comparisonKPIs.resultMismatches > 0 ? "1px solid rgba(239, 68, 68, 0.3)" : "1px solid var(--line)" }}>
                      <div style={{ fontSize: "11px", color: comparisonKPIs.resultMismatches > 0 ? "#ef4444" : "var(--muted)", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                        <AlertTriangle size={14} /> Result Mismatches
                      </div>
                      <div style={{ fontSize: "20px", fontWeight: 700, color: comparisonKPIs.resultMismatches > 0 ? "#ef4444" : "var(--ink)", marginTop: "3px" }}>
                        {comparisonKPIs.resultMismatches}
                      </div>
                      <div style={{ fontSize: "10.5px", color: comparisonKPIs.resultMismatches > 0 ? "#ef4444" : "var(--muted)", fontWeight: 600 }}>
                        {comparisonKPIs.passedHereFailedPub > 0 ? `${comparisonKPIs.passedHereFailedPub} Pass/Fail · ` : ""}{comparisonKPIs.failedHerePassedPub > 0 ? `${comparisonKPIs.failedHerePassedPub} Fail/Pass · ` : ""}{comparisonKPIs.heldMismatches > 0 ? `${comparisonKPIs.heldMismatches} Held` : comparisonKPIs.resultMismatches === 0 ? "Zero discrepancies" : ""}
                      </div>
                    </div>

                    {/* Card 4: Moderation Concordance */}
                    <div style={{ background: "rgba(99, 102, 241, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(99, 102, 241, 0.3)" }}>
                      <div style={{ fontSize: "11px", color: "#6366f1", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                        <Scale size={14} /> Moderation Concordance
                      </div>
                      <div style={{ fontSize: "20px", fontWeight: 700, color: "#6366f1", marginTop: "3px" }}>
                        {comparisonKPIs.modMatchRate}%
                      </div>
                      <div style={{ fontSize: "10.5px", color: "#6366f1", fontWeight: 600 }}>
                        {comparisonKPIs.modMatches} exact marks ({comparisonKPIs.modDiffs} diffs)
                      </div>
                    </div>

                    {/* Card 5: Moderation Total Comparison */}
                    <div style={{ background: "rgba(139, 92, 246, 0.08)", padding: "12px 14px", borderRadius: "8px", border: "1px solid rgba(139, 92, 246, 0.3)" }}>
                      <div style={{ fontSize: "11px", color: "#8b5cf6", fontWeight: 600, display: "flex", alignItems: "center", gap: "5px" }}>
                        <ArrowRightLeft size={14} /> Moderation Marks Total
                      </div>
                      <div style={{ fontSize: "16px", fontWeight: 700, color: "#8b5cf6", marginTop: "5px", display: "flex", alignItems: "baseline", gap: "6px" }}>
                        <span>Our: {comparisonKPIs.totalCalcMod}</span>
                        <span style={{ fontSize: "12px", color: "var(--muted)" }}>vs</span>
                        <span>Pub: {comparisonKPIs.totalPubMod}</span>
                      </div>
                      <div style={{ fontSize: "10px", color: "var(--muted)", marginTop: "2px" }}>
                        Event Mod: Our {comparisonKPIs.totalCalcEventMod} vs Pub {comparisonKPIs.totalPubEventMod} | Reg: {comparisonKPIs.totalRegMod}
                      </div>
                      <div style={{ fontSize: "10.5px", color: comparisonKPIs.modTotalDiff === 0 ? "#10b981" : "#f59e0b", fontWeight: 600, marginTop: "2px" }}>
                        Net Diff: {comparisonKPIs.modTotalDiff >= 0 ? `+${comparisonKPIs.modTotalDiff}` : comparisonKPIs.modTotalDiff} marks
                      </div>
                    </div>
                  </div>

                  {/* Toolbar & Filter Bar */}
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "10px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap", flex: 1 }}>
                      {/* Search */}
                      <div style={{ position: "relative", minWidth: "220px", flex: 1, maxWidth: "320px" }}>
                        <Search size={14} style={{ position: "absolute", left: "10px", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                        <input 
                          type="text"
                          placeholder="Search PRN, Seat, Student, College, Reappear..."
                          value={comparisonSearch}
                          onChange={(e) => { setComparisonSearch(e.target.value); setComparisonPage(0); }}
                          style={{ width: "100%", padding: "6px 10px 6px 30px", fontSize: "12px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--panel)" }}
                        />
                        {comparisonSearch && (
                          <X 
                            size={13} 
                            onClick={() => { setComparisonSearch(""); setComparisonPage(0); }}
                            style={{ position: "absolute", right: "10px", top: "50%", transform: "translateY(-50%)", cursor: "pointer", color: "var(--muted)" }}
                          />
                        )}
                      </div>

                      {/* College Filter */}
                      {uniqueComparisonColleges.length > 0 && (
                        <select 
                          value={comparisonCollegeFilter} 
                          onChange={(e) => { setComparisonCollegeFilter(e.target.value); setComparisonPage(0); }}
                          style={{ padding: "6px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--panel)", maxWidth: "200px" }}
                          title="Filter comparison by College"
                        >
                          <option value="ALL">All Colleges ({uniqueComparisonColleges.length})</option>
                          {uniqueComparisonColleges.map(c => (
                            <option key={c} value={c}>{c}</option>
                          ))}
                        </select>
                      )}
                    </div>

                    {/* Filter Pills */}
                    <div style={{ display: "flex", background: "var(--panel)", padding: "2px", borderRadius: "6px", border: "1px solid var(--line)", gap: "2px", flexWrap: "wrap" }}>
                      <button 
                        type="button"
                        onClick={() => { setComparisonFilterStatus("ALL"); setComparisonPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: comparisonFilterStatus === "ALL" ? "var(--accent)" : "transparent", color: comparisonFilterStatus === "ALL" ? "white" : "var(--muted)" }}
                      >
                        All ({comparisonRecords.length})
                      </button>
                      <button 
                        type="button"
                        onClick={() => { setComparisonFilterStatus("MISMATCH"); setComparisonPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: comparisonFilterStatus === "MISMATCH" ? "#ef4444" : "transparent", color: comparisonFilterStatus === "MISMATCH" ? "white" : "var(--muted)" }}
                      >
                        ⚠️ Result Mismatches ({comparisonKPIs.resultMismatches})
                      </button>
                      <button 
                        type="button"
                        onClick={() => { setComparisonFilterStatus("MOD_DIFF"); setComparisonPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: comparisonFilterStatus === "MOD_DIFF" ? "#f59e0b" : "transparent", color: comparisonFilterStatus === "MOD_DIFF" ? "white" : "var(--muted)" }}
                      >
                        ⚖️ Mod Diffs ({comparisonKPIs.modDiffs})
                      </button>
                      <button 
                        type="button"
                        onClick={() => { setComparisonFilterStatus("EXACT_MATCH"); setComparisonPage(0); }}
                        style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: comparisonFilterStatus === "EXACT_MATCH" ? "#10b981" : "transparent", color: comparisonFilterStatus === "EXACT_MATCH" ? "white" : "var(--muted)" }}
                      >
                        ✅ Exact Match ({comparisonKPIs.exactMatches})
                      </button>
                      {comparisonKPIs.passedHereFailedPub > 0 && (
                        <button 
                          type="button"
                          onClick={() => { setComparisonFilterStatus("PASSED_HERE_FAILED_PUB"); setComparisonPage(0); }}
                          style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: comparisonFilterStatus === "PASSED_HERE_FAILED_PUB" ? "#3b82f6" : "transparent", color: comparisonFilterStatus === "PASSED_HERE_FAILED_PUB" ? "white" : "var(--muted)" }}
                        >
                          Pass Here / Fail Pub ({comparisonKPIs.passedHereFailedPub})
                        </button>
                      )}
                      {comparisonKPIs.failedHerePassedPub > 0 && (
                        <button 
                          type="button"
                          onClick={() => { setComparisonFilterStatus("FAILED_HERE_PASSED_PUB"); setComparisonPage(0); }}
                          style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: comparisonFilterStatus === "FAILED_HERE_PASSED_PUB" ? "#dc2626" : "transparent", color: comparisonFilterStatus === "FAILED_HERE_PASSED_PUB" ? "white" : "var(--muted)" }}
                        >
                          Fail Here / Pass Pub ({comparisonKPIs.failedHerePassedPub})
                        </button>
                      )}
                      {comparisonKPIs.notInDataset > 0 && (
                        <button 
                          type="button"
                          onClick={() => { setComparisonFilterStatus("NOT_IN_DATA"); setComparisonPage(0); }}
                          style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: comparisonFilterStatus === "NOT_IN_DATA" ? "#64748b" : "transparent", color: comparisonFilterStatus === "NOT_IN_DATA" ? "white" : "var(--muted)" }}
                        >
                          Not In Dataset ({comparisonKPIs.notInDataset})
                        </button>
                      )}
                      {comparisonKPIs.heldbackCount > 0 && (
                        <button 
                          type="button"
                          onClick={() => { setComparisonFilterStatus("HELDBACK"); setComparisonPage(0); }}
                          style={{ padding: "3px 8px", fontSize: "11px", fontWeight: 600, border: "none", borderRadius: "4px", cursor: "pointer", background: comparisonFilterStatus === "HELDBACK" ? "#7c3aed" : "transparent", color: comparisonFilterStatus === "HELDBACK" ? "white" : "var(--muted)" }}
                        >
                          🔒 Heldback ({comparisonKPIs.heldbackCount})
                        </button>
                      )}
                    </div>
                  </div>

                  {/* Reconciliation Table */}
                  <div style={{ background: "var(--panel)", borderRadius: "8px", border: "1px solid var(--line)", overflow: "hidden", display: "flex", flexDirection: "column", flex: 1 }}>
                    <div style={{ overflowX: "auto", maxHeight: "calc(100vh - 380px)", overflowY: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "11.5px", textAlign: "left" }}>
                        <thead>
                          <tr style={{ background: "var(--bg)", borderBottom: "1px solid var(--line)", position: "sticky", top: 0, zIndex: 10 }}>
                            <th style={{ padding: "8px 10px", width: "40px", color: "var(--muted)" }}>#</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600 }}>PRN</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600 }}>Seat No</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600 }}>Student Name</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600 }}>College</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600, textAlign: "center" }}>Calculated Result</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600, textAlign: "center" }}>Published Result</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600, textAlign: "center" }}>Result Concordance</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600, textAlign: "center" }} title="Moderation applied in Original / Regular Event from uploaded Gazette">Reg Mod</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600, textAlign: "center" }} title="Supplementary Event Moderation: Our System vs University Published (Pub Ord Total - Reg Mod)">Event Mod (Our / Pub)</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600, textAlign: "center" }} title="Our Mod Total = Regular Mod + Our Event Mod">Our Mod Total</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600, textAlign: "center" }} title="Published Ordinance Total from University Result">Pub Ord Total</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600, textAlign: "center" }}>Mod Match</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600 }}>Our Failed Courses</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600 }}>Pub Reappear Courses</th>
                            <th style={{ padding: "8px 10px", fontWeight: 600 }}>Reconciliation Notes</th>
                          </tr>
                        </thead>
                        <tbody>
                          {filteredComparisonRecords.length === 0 ? (
                            <tr>
                              <td colSpan={16} style={{ padding: "40px", textAlign: "center", color: "var(--muted)" }}>
                                No student comparison records found matching the current filters.
                              </td>
                            </tr>
                          ) : (
                            filteredComparisonRecords
                              .slice(comparisonPage * rowsPerPage, (comparisonPage + 1) * rowsPerPage)
                              .map((rec, idx) => {
                                const rowIdx = comparisonPage * rowsPerPage + idx + 1;
                                const isMismatch = !rec.isResultMatch && rec.isFound;
                                const isModDiff = !rec.isModMatch && rec.isFound;

                                return (
                                  <tr 
                                    key={`${rec.prn}-${rec.seat}-${idx}`}
                                    style={{ 
                                      borderBottom: "1px solid var(--line)", 
                                      background: isMismatch ? "rgba(239, 68, 68, 0.04)" : isModDiff ? "rgba(245, 158, 11, 0.04)" : idx % 2 === 1 ? "rgba(255, 255, 255, 0.02)" : "transparent"
                                    }}
                                  >
                                    <td style={{ padding: "7px 10px", color: "var(--muted)" }}>{rowIdx}</td>
                                    <td style={{ padding: "7px 10px", fontFamily: "monospace", fontWeight: 600 }}>{rec.prn || "-"}</td>
                                    <td style={{ padding: "7px 10px", fontFamily: "monospace", fontWeight: 600 }}>{rec.seat || "-"}</td>
                                    <td style={{ padding: "7px 10px", fontWeight: 500 }}>{rec.name || "-"}</td>
                                    <td style={{ padding: "7px 10px", maxWidth: "180px", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={rec.college}>
                                      {rec.college || "-"}
                                    </td>
                                    
                                    {/* Calculated Result */}
                                    <td style={{ padding: "7px 10px", textAlign: "center" }}>
                                      {!rec.isFound ? (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(148, 163, 184, 0.15)", color: "#64748b", fontWeight: 600 }}>
                                          Not Loaded
                                        </span>
                                      ) : rec.calcIsHeldback ? (
                                        rec.calcResult.startsWith("Pass") ? (
                                          <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(16, 185, 129, 0.15)", color: "#10b981", fontWeight: 600 }}>
                                            PASS <span style={{ fontSize: "9.5px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                          </span>
                                        ) : (rec.calcIsHeld || !rec.calcHeldbackHasMarks || rec.calcResult.includes("Held - Missing")) ? (
                                          <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(239, 68, 68, 0.15)", color: "#ef4444", fontWeight: 600 }}>
                                            FAIL <span style={{ fontSize: "9.5px", color: "#e11d48", fontWeight: 600 }}>(Held - Missing)</span> <span style={{ fontSize: "9.5px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                          </span>
                                        ) : (
                                          <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(239, 68, 68, 0.15)", color: "#ef4444", fontWeight: 600 }}>
                                            FAIL <span style={{ fontSize: "9.5px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>
                                          </span>
                                        )
                                      ) : rec.calcResult.startsWith("Pass") || rec.calcResult === "Pass" ? (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(16, 185, 129, 0.15)", color: "#10b981", fontWeight: 600 }}>
                                          PASS {rec.calcIsRescued ? "(Mod)" : ""}
                                        </span>
                                      ) : (rec.calcIsHeld || rec.calcResult.includes("Held - Missing") || rec.calcResult === "Held (Missing Component)") ? (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(239, 68, 68, 0.15)", color: "#ef4444", fontWeight: 600 }}>
                                          FAIL <span style={{ fontSize: "9.5px", color: "#e11d48", fontWeight: 600 }}>(Held - Missing)</span>
                                        </span>
                                      ) : rec.calcResult === "Held" || rec.calcResult === "Heldback" ? (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(192, 38, 211, 0.15)", color: "#c026d3", fontWeight: 600 }}>
                                          {rec.calcResult.toUpperCase()}
                                        </span>
                                      ) : (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(239, 68, 68, 0.15)", color: "#ef4444", fontWeight: 600 }}>
                                          FAIL
                                        </span>
                                      )}
                                    </td>

                                    {/* Published Result */}
                                    <td style={{ padding: "7px 10px", textAlign: "center" }}>
                                      {rec.pubExactResult?.startsWith("Pass") ? (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(16, 185, 129, 0.15)", color: "#10b981", fontWeight: 600 }}>
                                          PASS {rec.isPubHeldback && <span style={{ fontSize: "9.5px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>}
                                        </span>
                                      ) : rec.pubExactResult?.includes("Held - Missing") ? (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(239, 68, 68, 0.15)", color: "#ef4444", fontWeight: 600 }}>
                                          FAIL <span style={{ fontSize: "9.5px", color: "#e11d48", fontWeight: 600 }}>(Held - Missing)</span> {rec.isPubHeldback && <span style={{ fontSize: "9.5px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>}
                                        </span>
                                      ) : rec.pubExactResult?.startsWith("Fail") || (!rec.isPubHeldback && rec.pubResultStd === "Fail") ? (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(239, 68, 68, 0.15)", color: "#ef4444", fontWeight: 600 }}>
                                          FAIL {rec.isPubHeldback && <span style={{ fontSize: "9.5px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>}
                                        </span>
                                      ) : rec.pubResultStd === "Held" || rec.pubResultStd === "Heldback" ? (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(192, 38, 211, 0.15)", color: "#c026d3", fontWeight: 600 }}>
                                          {rec.pubResultStd.toUpperCase()} {rec.isPubHeldback && rec.pubResultStd !== "Heldback" && <span style={{ fontSize: "9.5px", color: "#a855f7", fontWeight: 700 }}>(Heldback)</span>}
                                        </span>
                                      ) : (
                                        <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(239, 68, 68, 0.15)", color: "#ef4444", fontWeight: 600 }}>
                                          {String(rec.pubExactResult || rec.pubResultStd || "FAIL").toUpperCase()}
                                        </span>
                                      )}
                                    </td>

                                    {/* Result Match */}
                                    <td style={{ padding: "7px 10px", textAlign: "center" }}>
                                      {!rec.isFound ? (
                                        <span style={{ fontSize: "10.5px", color: "#64748b" }}>—</span>
                                      ) : rec.isResultMatch ? (
                                        <span style={{ fontSize: "10.5px", color: "#10b981", fontWeight: 600, display: "inline-flex", alignItems: "center", gap: "3px" }}>
                                          <CheckCircle2 size={12} /> Match
                                        </span>
                                      ) : (
                                        <span style={{ fontSize: "10.5px", color: "#ef4444", fontWeight: 700, display: "inline-flex", alignItems: "center", gap: "3px" }}>
                                          <AlertTriangle size={12} /> MISMATCH
                                        </span>
                                      )}
                                    </td>

                                    {/* Reg Mod (from Gazette) */}
                                    <td style={{ padding: "7px 10px", textAlign: "center", fontWeight: rec.calcRegularModMarks > 0 ? 600 : 400, color: rec.calcRegularModMarks > 0 ? "#0284c7" : "var(--muted)" }}>
                                      {rec.calcRegularModMarks > 0 ? (
                                        <span style={{ fontSize: "10.5px", background: "rgba(2, 132, 199, 0.1)", color: "#0284c7", padding: "1px 5px", borderRadius: "3px", fontWeight: 600 }} title={`Regular Event Moderation: ${rec.calcRegularModMarks} marks`}>
                                          {rec.calcRegularModMarks.toFixed(2)}
                                        </span>
                                      ) : (
                                        <span style={{ color: "var(--muted)", fontSize: "10.5px" }}>0</span>
                                      )}
                                    </td>

                                    {/* Event Mod (Our / Pub) */}
                                    <td style={{ padding: "7px 10px", textAlign: "center", whiteSpace: "nowrap" }}>
                                      {!rec.isFound ? (
                                        <span style={{ fontSize: "10.5px", color: "#64748b" }}>—</span>
                                      ) : (
                                        <div style={{ display: "inline-flex", alignItems: "center", gap: "3px", fontSize: "11px" }}>
                                          <span style={{ fontWeight: rec.calcEventModMarks > 0 ? 700 : 500, color: rec.calcEventModMarks > 0 ? "#8b5cf6" : "var(--muted)" }} title="Our Calculated Supplementary Event Moderation">
                                            {rec.calcEventModMarks.toFixed(2)}
                                          </span>
                                          <span style={{ color: "var(--muted)", fontSize: "10px" }}>/</span>
                                          <span style={{ fontWeight: rec.pubEventModMarks > 0 ? 700 : 500, color: rec.pubEventModMarks > 0 ? "#8b5cf6" : "var(--muted)" }} title={`Published Event Moderation = Pub Ord Total (${rec.pubModMarks}) - Reg Mod (${rec.calcRegularModMarks}) = ${rec.pubEventModMarks.toFixed(2)}`}>
                                            {rec.pubEventModMarks.toFixed(2)}
                                          </span>
                                          {!rec.isEventModMatch && (
                                            <span style={{ fontSize: "9px", color: "#f59e0b", fontWeight: 700, marginLeft: "2px" }} title={`Event Mod Diff: ${rec.eventModDiff > 0 ? '+' : ''}${rec.eventModDiff.toFixed(2)}`}>
                                              ({rec.eventModDiff > 0 ? `+${rec.eventModDiff.toFixed(1)}` : rec.eventModDiff.toFixed(1)})
                                            </span>
                                          )}
                                        </div>
                                      )}
                                    </td>

                                    {/* Our Mod Total */}
                                    <td style={{ padding: "7px 10px", textAlign: "center", fontWeight: rec.calcModMarks > 0 ? 700 : 400, color: rec.calcModMarks > 0 ? "#8b5cf6" : "var(--muted)" }} title={`Our Mod Total = Reg Mod (${rec.calcRegularModMarks}) + Event Mod (${rec.calcEventModMarks})`}>
                                      {rec.isFound ? rec.calcModMarks.toFixed(2) : "—"}
                                    </td>

                                    {/* Pub Ord Total */}
                                    <td style={{ padding: "7px 10px", textAlign: "center", fontWeight: rec.pubModMarks > 0 ? 700 : 400, color: rec.pubModMarks > 0 ? "#8b5cf6" : "var(--muted)" }} title="Published Ordinance Total">
                                      {rec.pubModMarks.toFixed(2)}
                                      {rec.pubHasOrd && (
                                        <span style={{ marginLeft: "4px", fontSize: "9.5px", background: "rgba(139, 92, 246, 0.15)", color: "#8b5cf6", padding: "1px 4px", borderRadius: "3px" }}>
                                          ORD
                                        </span>
                                      )}
                                    </td>

                                    {/* Mod Match */}
                                    <td style={{ padding: "7px 10px", textAlign: "center" }}>
                                      {!rec.isFound ? (
                                        <span style={{ fontSize: "10.5px", color: "#64748b" }}>—</span>
                                      ) : rec.isModMatch ? (
                                        <span style={{ fontSize: "10.5px", color: "#10b981", fontWeight: 600 }} title="Both Total Moderation and Event Moderation Match">
                                          <CheckCircle2 size={12} />
                                        </span>
                                      ) : (
                                        <span style={{ fontSize: "10.5px", color: "#f59e0b", fontWeight: 700 }} title={`Moderation Diff: ${rec.modDiff > 0 ? '+' : ''}${rec.modDiff.toFixed(2)}`}>
                                          {rec.modDiff > 0 ? `+${rec.modDiff.toFixed(2)}` : rec.modDiff.toFixed(2)}
                                        </span>
                                      )}
                                    </td>

                                    {/* Our Failed Courses */}
                                    <td style={{ padding: "7px 10px", maxWidth: "160px" }}>
                                      {rec.calcFailedCourses && rec.calcFailedCourses.length > 0 ? (
                                        <div style={{ display: "flex", flexWrap: "wrap", gap: "2px" }}>
                                          {rec.calcFailedCourses.map(c => (
                                            <span key={c} style={{ fontSize: "9.5px", background: "rgba(239, 68, 68, 0.1)", color: "#ef4444", padding: "1px 4px", borderRadius: "3px", fontFamily: "monospace" }}>
                                              {c}
                                            </span>
                                          ))}
                                        </div>
                                      ) : (
                                        <span style={{ color: "var(--muted)", fontSize: "10.5px" }}>None</span>
                                      )}
                                    </td>

                                    {/* Pub Reappear Courses */}
                                    <td style={{ padding: "7px 10px", maxWidth: "160px" }}>
                                      {rec.reappear ? (
                                        <div style={{ display: "flex", flexWrap: "wrap", gap: "2px" }}>
                                          {rec.reappear.split(",").map(c => c.trim()).filter(Boolean).map(c => (
                                            <span key={c} style={{ fontSize: "9.5px", background: "rgba(245, 158, 11, 0.1)", color: "#d97706", padding: "1px 4px", borderRadius: "3px", fontFamily: "monospace" }}>
                                              {c}
                                            </span>
                                          ))}
                                        </div>
                                      ) : (
                                        <span style={{ color: "var(--muted)", fontSize: "10.5px" }}>None</span>
                                      )}
                                    </td>

                                    {/* Reconciliation Notes */}
                                    <td style={{ padding: "7px 10px", fontSize: "11px", color: isMismatch ? "#ef4444" : isModDiff ? "#d97706" : "var(--muted)" }}>
                                      {rec.discrepancyDescription}
                                    </td>
                                  </tr>
                                );
                              })
                          )}
                        </tbody>
                      </table>
                    </div>

                    {/* Pagination Footer */}
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "8px 16px", borderTop: "1px solid var(--line)", background: "var(--bg)", fontSize: "11.5px" }}>
                      <div style={{ color: "var(--muted)" }}>
                        Showing {filteredComparisonRecords.length === 0 ? 0 : comparisonPage * rowsPerPage + 1} to {Math.min((comparisonPage + 1) * rowsPerPage, filteredComparisonRecords.length)} of {filteredComparisonRecords.length} records
                      </div>

                      <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                        <button 
                          type="button" 
                          disabled={comparisonPage === 0} 
                          onClick={() => setComparisonPage(p => Math.max(0, p - 1))}
                          style={{ padding: "4px 8px", fontSize: "11px", borderRadius: "4px", border: "1px solid var(--line)", background: "var(--panel)", cursor: comparisonPage === 0 ? "not-allowed" : "pointer", opacity: comparisonPage === 0 ? 0.5 : 1 }}
                        >
                          Previous
                        </button>
                        <span>Page {comparisonPage + 1} of {Math.max(1, Math.ceil(filteredComparisonRecords.length / rowsPerPage))}</span>
                        <button 
                          type="button" 
                          disabled={(comparisonPage + 1) * rowsPerPage >= filteredComparisonRecords.length} 
                          onClick={() => setComparisonPage(p => p + 1)}
                          style={{ padding: "4px 8px", fontSize: "11px", borderRadius: "4px", border: "1px solid var(--line)", background: "var(--panel)", cursor: (comparisonPage + 1) * rowsPerPage >= filteredComparisonRecords.length ? "not-allowed" : "pointer", opacity: (comparisonPage + 1) * rowsPerPage >= filteredComparisonRecords.length ? 0.5 : 1 }}
                        >
                          Next
                        </button>
                      </div>
                    </div>
                  </div>
                </>
              )}
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

                  <div style={{ display: "flex", alignItems: "center", gap: "8px", flexWrap: "wrap" }}>
                    
                    {/* View Controls Group */}
                    <div style={{ display: "flex", alignItems: "center", gap: "6px", flexWrap: "wrap" }}>
                      
                      {/* Column Preset Pills */}
                      <div style={{ display: "flex", alignItems: "center", background: "var(--bg)", borderRadius: "6px", border: "1px solid var(--line)", padding: "2px", gap: "2px" }}>
                        <span style={{ fontSize: "10.5px", fontWeight: 700, color: "var(--muted)", padding: "0 4px" }}>Cols:</span>
                        <button
                          type="button"
                          onClick={() => setColumnPreset("ALL")}
                          style={{
                            padding: "3px 7px",
                            fontSize: "10.5px",
                            fontWeight: 600,
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                            background: columnPreset === "ALL" ? "var(--accent)" : "transparent",
                            color: columnPreset === "ALL" ? "white" : "var(--muted)"
                          }}
                          title="Show all 31 university output columns"
                        >
                          All (31)
                        </button>
                        <button
                          type="button"
                          onClick={() => setColumnPreset("ESSENTIAL")}
                          style={{
                            padding: "3px 7px",
                            fontSize: "10.5px",
                            fontWeight: 600,
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                            background: columnPreset === "ESSENTIAL" ? "var(--accent)" : "transparent",
                            color: columnPreset === "ESSENTIAL" ? "white" : "var(--muted)"
                          }}
                          title="Show only key columns: Code, Name, Seat, PRN, ESE Obt, CE Obt, Total, Status, Mod"
                        >
                          ⭐ Essential (10)
                        </button>
                        <button
                          type="button"
                          onClick={() => setColumnPreset("SCORES")}
                          style={{
                            padding: "3px 7px",
                            fontSize: "10.5px",
                            fontWeight: 600,
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                            background: columnPreset === "SCORES" ? "var(--accent)" : "transparent",
                            color: columnPreset === "SCORES" ? "white" : "var(--muted)"
                          }}
                          title="Show student identifiers and all obtained marks breakdown"
                        >
                          📊 Scores (14)
                        </button>
                        <button
                          type="button"
                          onClick={() => setShowColumnPicker(v => !v)}
                          style={{
                            padding: "3px 7px",
                            fontSize: "10.5px",
                            fontWeight: 600,
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                            display: "flex",
                            alignItems: "center",
                            gap: "3px",
                            background: (columnPreset === "CUSTOM" || showColumnPicker) ? "var(--accent)" : "transparent",
                            color: (columnPreset === "CUSTOM" || showColumnPicker) ? "white" : "var(--muted)"
                          }}
                          title="Customize column visibility with checkboxes"
                        >
                          <Columns size={11} /> {columnPreset === "CUSTOM" ? `Custom (${visibleColumns.size})` : "Pick..."}
                        </button>
                      </div>

                      {/* Density Selector */}
                      <div style={{ display: "flex", alignItems: "center", background: "var(--bg)", borderRadius: "6px", border: "1px solid var(--line)", padding: "2px", gap: "2px" }}>
                        <span style={{ fontSize: "10.5px", fontWeight: 700, color: "var(--muted)", padding: "0 4px" }}>Density:</span>
                        <button
                          type="button"
                          onClick={() => setTableDensity("compact")}
                          style={{
                            padding: "3px 6px",
                            fontSize: "10.5px",
                            fontWeight: 600,
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                            background: tableDensity === "compact" ? "var(--accent)" : "transparent",
                            color: tableDensity === "compact" ? "white" : "var(--muted)"
                          }}
                          title="Compact row spacing to view more data on screen"
                        >
                          Compact
                        </button>
                        <button
                          type="button"
                          onClick={() => setTableDensity("normal")}
                          style={{
                            padding: "3px 6px",
                            fontSize: "10.5px",
                            fontWeight: 600,
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                            background: tableDensity === "normal" ? "var(--accent)" : "transparent",
                            color: tableDensity === "normal" ? "white" : "var(--muted)"
                          }}
                          title="Standard row spacing"
                        >
                          Normal
                        </button>
                        <button
                          type="button"
                          onClick={() => setTableDensity("comfortable")}
                          style={{
                            padding: "3px 6px",
                            fontSize: "10.5px",
                            fontWeight: 600,
                            border: "none",
                            borderRadius: "4px",
                            cursor: "pointer",
                            background: tableDensity === "comfortable" ? "var(--accent)" : "transparent",
                            color: tableDensity === "comfortable" ? "white" : "var(--muted)"
                          }}
                          title="Roomy spacing"
                        >
                          Roomy
                        </button>
                      </div>

                      {/* Page Size Dropdown */}
                      <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
                        <span style={{ fontSize: "11px", color: "var(--muted)", fontWeight: 600 }}>Rows:</span>
                        <select
                          value={pageSize}
                          onChange={(e) => {
                            const val = e.target.value === "ALL" ? "ALL" : Number(e.target.value);
                            setPageSize(val);
                            setPage(0);
                          }}
                          style={{
                            padding: "3px 6px",
                            fontSize: "11px",
                            borderRadius: "5px",
                            border: "1px solid var(--line)",
                            background: "var(--bg)",
                            cursor: "pointer"
                          }}
                        >
                          <option value={25}>25</option>
                          <option value={50}>50</option>
                          <option value={100}>100</option>
                          <option value={250}>250</option>
                          <option value="ALL">All ({filteredRows.length})</option>
                        </select>
                      </div>

                      {/* Focus / Maximize Button */}
                      <button
                        type="button"
                        onClick={() => {
                          setIsTableMaximized(v => !v);
                          if (!isTableMaximized) setIsSidebarCollapsed(true);
                        }}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: "4px",
                          padding: "4px 8px",
                          fontSize: "11px",
                          fontWeight: 600,
                          borderRadius: "5px",
                          border: `1px solid ${isTableMaximized ? "var(--accent)" : "var(--line)"}`,
                          background: isTableMaximized ? "var(--accent-soft)" : "var(--bg)",
                          color: isTableMaximized ? "var(--accent)" : "var(--muted)",
                          cursor: "pointer"
                        }}
                        title={isTableMaximized ? "Restore standard view" : "Focus Mode: maximize table to full workspace"}
                      >
                        {isTableMaximized ? <Minimize2 size={12} /> : <Maximize2 size={12} />}
                        <span>{isTableMaximized ? "Restore" : "Focus"}</span>
                      </button>

                    </div>

                    <button 
                      type="button" 
                      onClick={() => handleExportExcel(filteredRows)}
                      style={{ display: "flex", alignItems: "center", gap: "4px", padding: "4px 9px", fontSize: "11px", borderRadius: "5px", border: "1px solid var(--accent)", background: "var(--accent-soft)", color: "var(--accent)", cursor: "pointer", fontWeight: 600 }}
                      title={filteredRows.length !== processedRows.length ? `Export current filtered view (${filteredRows.length} rows) to 31-Col ADES XLSX` : "Export all rows to 31-Col ADES XLSX"}
                    >
                      <Download size={12} /> Export Current View ({filteredRows.length})
                    </button>

                    <span style={{ fontSize: "11.5px", color: "var(--muted)", whiteSpace: "nowrap" }}>
                      Showing {filteredRows.length > 0 ? page * (pageSize === "ALL" ? filteredRows.length : pageSize) + 1 : 0} - {pageSize === "ALL" ? filteredRows.length : Math.min((page + 1) * pageSize, filteredRows.length)} of {filteredRows.length} rows
                    </span>

                    {/* Pagination Controls */}
                    {pageSize !== "ALL" && (
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
                    )}
                  </div>

                </div>

                {/* Column Picker Modal / Popover */}
                {showColumnPicker && (
                  <div 
                    style={{
                      position: "absolute",
                      right: "20px",
                      top: "85px",
                      zIndex: 60,
                      background: "var(--panel)",
                      border: "1px solid var(--line)",
                      borderRadius: "10px",
                      boxShadow: "0 10px 30px rgba(0,0,0,0.18)",
                      width: "320px",
                      maxHeight: "440px",
                      display: "flex",
                      flexDirection: "column",
                      padding: "12px"
                    }}
                  >
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "8px", borderBottom: "1px solid var(--line)", paddingBottom: "8px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                        <Columns size={14} color="var(--accent)" />
                        <span style={{ fontSize: "12px", fontWeight: 700 }}>Choose Columns ({visibleColumns.size}/{ADES_OUTPUT_HEADERS.length})</span>
                      </div>
                      <button 
                        type="button" 
                        onClick={() => setShowColumnPicker(false)}
                        style={{ background: "transparent", border: "none", color: "var(--muted)", cursor: "pointer", padding: "2px" }}
                      >
                        <X size={14} />
                      </button>
                    </div>

                    <div style={{ display: "flex", gap: "6px", marginBottom: "8px" }}>
                      <button
                        type="button"
                        onClick={() => {
                          setVisibleColumns(new Set(ADES_OUTPUT_HEADERS));
                          setColumnPreset("ALL");
                        }}
                        style={{ flex: 1, padding: "3px 6px", fontSize: "10.5px", background: "var(--bg)", border: "1px solid var(--line)", borderRadius: "4px", cursor: "pointer", fontWeight: 600 }}
                      >
                        Select All
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setColumnPreset("ESSENTIAL");
                          setVisibleColumns(new Set([
                            "Course Code", "Course Name", "Seat Number", "PRN", 
                            "ESE - TH Obtained", "CE - TH Obtained", "Course Overall Marks ", 
                            "ESE Pass", "Course Pass/Fail", "Moderation Marks"
                          ]));
                        }}
                        style={{ flex: 1, padding: "3px 6px", fontSize: "10.5px", background: "var(--bg)", border: "1px solid var(--line)", borderRadius: "4px", cursor: "pointer", fontWeight: 600 }}
                      >
                        Essential
                      </button>
                      <button
                        type="button"
                        onClick={() => {
                          setColumnPreset("SCORES");
                          setVisibleColumns(new Set([
                            "Seat Number", "PRN", "Course Code", "Course Name",
                            "ESE - PR Obtained", "ESE - TH Obtained", "ESE Overall",
                            "CE - PR Obtained", "CE - TH Obtained", "CE Overall Marks ",
                            "Course Overall Marks ", "ESE Pass", "Course Pass/Fail", "Moderation Marks"
                          ]));
                        }}
                        style={{ flex: 1, padding: "3px 6px", fontSize: "10.5px", background: "var(--bg)", border: "1px solid var(--line)", borderRadius: "4px", cursor: "pointer", fontWeight: 600 }}
                      >
                        Scores
                      </button>
                    </div>

                    <div style={{ flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: "4px", paddingRight: "4px" }}>
                      {ADES_OUTPUT_HEADERS.map((col) => {
                        const isChecked = visibleColumns.has(col);
                        return (
                          <label 
                            key={col} 
                            style={{ 
                              display: "flex", 
                              alignItems: "center", 
                              gap: "8px", 
                              fontSize: "11px", 
                              padding: "3px 6px", 
                              borderRadius: "4px", 
                              cursor: "pointer",
                              background: isChecked ? "var(--bg)" : "transparent"
                            }}
                          >
                            <input 
                              type="checkbox"
                              checked={isChecked}
                              onChange={(e) => {
                                const next = new Set(visibleColumns);
                                if (e.target.checked) next.add(col);
                                else next.delete(col);
                                setVisibleColumns(next);
                                setColumnPreset("CUSTOM");
                              }}
                              style={{ cursor: "pointer" }}
                            />
                            <span style={{ color: isChecked ? "var(--ink)" : "var(--muted)", fontWeight: isChecked ? 600 : 400 }}>{col}</span>
                          </label>
                        );
                      })}
                    </div>

                    <div style={{ display: "flex", justifyContent: "flex-end", marginTop: "8px", borderTop: "1px solid var(--line)", paddingTop: "8px" }}>
                      <button
                        type="button"
                        onClick={() => setShowColumnPicker(false)}
                        style={{ padding: "4px 12px", fontSize: "11px", fontWeight: 600, background: "var(--accent)", color: "white", border: "none", borderRadius: "5px", cursor: "pointer" }}
                      >
                        Done
                      </button>
                    </div>
                  </div>
                )}
              </div>

              {/* Data Table Grid */}
              <div style={{ flex: 1, overflow: "auto", position: "relative" }}>
                <table style={{ 
                  width: "100%", 
                  borderCollapse: "collapse", 
                  fontSize: tableDensity === "compact" ? "10.5px" : tableDensity === "comfortable" ? "12px" : "11.5px", 
                  whiteSpace: "nowrap" 
                }}>
                  <thead style={{ position: "sticky", top: 0, background: "var(--panel)", zIndex: 10, borderBottom: "1px solid var(--line)", boxShadow: "0 1px 3px rgba(0,0,0,0.05)" }}>
                    <tr>
                      <th style={{ padding: tableDensity === "compact" ? "3px 6px" : tableDensity === "comfortable" ? "8px 10px" : "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center", width: "40px", color: "var(--muted)" }}>#</th>
                      {displayedHeaders.map((header) => {
                        const isSorted = sortConfig.column === header;
                        const isPassCol = header.includes("Pass") || header.includes("Pass/Fail");
                        const isModCol = header === "Moderation Marks";
                        
                        return (
                          <th 
                            key={header} 
                            style={{ 
                              padding: tableDensity === "compact" ? "3px 6px" : tableDensity === "comfortable" ? "8px 10px" : "5px 8px", 
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
                        <td colSpan={displayedHeaders.length + 1} style={{ padding: "40px", textAlign: "center", color: "var(--muted)" }}>
                          {processedRows.length === 0 
                            ? "Upload an ADES Marks Excel sheet (.xlsx, .xls) to view calculated student course results and apply moderation." 
                            : "No student records match the active search or column filters."}
                        </td>
                      </tr>
                    ) : (
                      pagedRows.map((row, idx) => {
                        const globalIdx = page * (pageSize === "ALL" ? filteredRows.length : pageSize) + idx + 1;
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
                            <td style={{ padding: tableDensity === "compact" ? "3px 6px" : tableDensity === "comfortable" ? "8px 10px" : "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center", color: "var(--muted)" }}>
                              {globalIdx}
                            </td>

                            {displayedHeaders.map((col) => {
                              const val = row[col];
                              const isCoursePass = col === coursePassKey;

                              if (col === "Course Code") {
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)" }}>
                                    <div style={{ display: "flex", alignItems: "center", gap: "5px" }}>
                                      <span style={{ fontWeight: 600 }}>{val}</span>
                                      {row._isImprovement ? (
                                        <span style={{ fontSize: "9px", background: "rgba(147, 51, 234, 0.15)", color: "#9333ea", padding: "1px 5px", borderRadius: "4px", fontWeight: 700 }} title="Improvement candidate (Only ESE-TH repeated; other components carried forward from previous event)">
                                          IMP
                                        </span>
                                      ) : (
                                        <span style={{ fontSize: "9px", background: "rgba(217, 119, 6, 0.15)", color: "#b45309", padding: "1px 5px", borderRadius: "4px", fontWeight: 700 }} title="Supplementary candidate (ESE-TH/PR repeated; CE marks carried forward from previous event)">
                                          SUPP
                                        </span>
                                      )}
                                      {row._hasHistoricalRecord ? (
                                        <span style={{ fontSize: "9.5px", color: "#10b981", fontWeight: 700 }} title={`Baseline linked from: ${row._historicalSource} | Carried forward: ${(row._carriedForwardComponents || []).join(", ") || "None"}`}>
                                          ✓
                                        </span>
                                      ) : (
                                        <span style={{ fontSize: "9px", color: "#ef4444", fontWeight: 700 }} title="No previous event record matched">
                                          ⚠
                                        </span>
                                      )}
                                    </div>
                                  </td>
                                );
                              }
                              const isOtherPass = col === "ESE Pass" || col === "Overall pass";
                              const isModMarks = col === "Moderation Marks";
                              const isAbsentMark = val === "Absent (Ab)";
                              const isMalpracticeMark = val === "Malpractice (MP)";
                              const isMissingMark = val === "Missing";

                              if (isCoursePass) {
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center" }}>
                                     {isHeldback && !isHeld ? (
                                      // Heldback but marks are present — show both raw & moderated result
                                      <span style={{ display: "inline-flex", flexDirection: "column", alignItems: "center", gap: "3px" }}>
                                        {/* Raw result (no moderation) */}
                                        <span
                                          style={{
                                            display: "inline-flex", alignItems: "center", gap: "3px",
                                            background: row._heldbackRawPass ? "rgba(16,185,129,0.12)" : "rgba(239,68,68,0.12)",
                                            color: row._heldbackRawPass ? "#10b981" : "#ef4444",
                                            padding: "2px 6px", borderRadius: "4px", fontWeight: 600, fontSize: "11px",
                                            border: "1px solid rgba(192,38,211,0.3)"
                                          }}
                                          title={`Without moderation | Heldback: ${row._heldbackReason || "Heldback at term-level"}`}
                                        >
                                          {row._heldbackRawPass
                                            ? <><CheckCircle2 size={10} /> Pass (No Mod)</>
                                            : <><XCircle size={10} /> Fail (No Mod)</>
                                          }
                                        </span>
                                        {/* Moderated result */}
                                        <span
                                          style={{
                                            display: "inline-flex", alignItems: "center", gap: "3px",
                                            background: row._heldbackModPass ? "rgba(16,185,129,0.2)" : "rgba(239,68,68,0.18)",
                                            color: row._heldbackModPass ? "#059669" : "#dc2626",
                                            padding: "2px 6px", borderRadius: "4px", fontWeight: 700, fontSize: "11px",
                                            border: "1px solid rgba(192,38,211,0.4)"
                                          }}
                                          title={`With moderation (+${row._heldbackModMarks || 0}) | Heldback: ${row._heldbackReason || "Heldback at term-level"}`}
                                        >
                                          <Lock size={10} />
                                          {row._heldbackModPass
                                            ? `Pass +${row._heldbackModMarks || 0} (Heldback)`
                                            : "Fail (Heldback)"
                                          }
                                        </span>
                                      </span>
                                    ) : isHeldback ? (
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

                              const isCarriedForwardCol = 
                                (col === "ESE - PR Obtained" && (row._carriedForwardComponents || []).includes("ESE-PR")) ||
                                (col === "CE - TH Obtained" && (row._carriedForwardComponents || []).includes("CE-TH")) ||
                                (col === "CE - PR Obtained" && (row._carriedForwardComponents || []).includes("CE-PR"));

                              return (
                                <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", color: "var(--ink)" }}>
                                  {val !== undefined && val !== null ? (
                                    isCarriedForwardCol ? (
                                      <span style={{ display: "inline-flex", alignItems: "center", gap: "3px" }} title="Carried forward from previous attempt baseline">
                                        <span style={{ color: "#059669", fontWeight: 600 }}>{String(val)}</span>
                                        <span style={{ fontSize: "8.5px", background: "rgba(16, 185, 129, 0.15)", color: "#059669", padding: "0 3px", borderRadius: "3px", fontWeight: 700 }}>CF</span>
                                      </span>
                                    ) : (
                                      String(val)
                                    )
                                  ) : ""}
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
              
              <div style={{ background: "rgba(16, 185, 129, 0.08)", border: "1px solid rgba(16, 185, 129, 0.25)", borderRadius: "8px", padding: "12px 14px" }}>
                <h4 style={{ margin: "0 0 8px", fontSize: "14px", color: "#10b981", display: "flex", alignItems: "center", gap: "6px" }}>
                  <TrendingUp size={16} /> ADES Result Calculation Lifecycle &amp; Sequence of Events
                </h4>
                <ol style={{ margin: 0, paddingLeft: "18px", color: "var(--ink)", display: "flex", flexDirection: "column", gap: "6px" }}>
                  <li>
                    <strong>Step 1: Course &amp; Student Semester Result Calculation (Raw):</strong> Results are evaluated based on available assessment conditions (30% ESE threshold, 35% overall threshold, absentees, malpractices, held/missing components, and heldback entries) <em>without moderation</em>.
                  </li>
                  <li>
                    <strong>Step 2: Pass Simulation Generation (+0 to +10):</strong> Based on raw conditions from Step 1, what-if pass simulations across 0 to +10 moderation marks are modeled course-wise and institution-wide to project pass percentage improvements.
                  </li>
                  <li>
                    <strong>Step 3: University Decision:</strong> Examination authorities review the Pass Simulation and decide what moderation marks policy/limit to approve.
                  </li>
                  <li>
                    <strong>Step 4: Approved Moderation Application:</strong> The approved moderation rule is entered into the Moderation Matrix and applied to eligible students to finalize moderated semester results, SGPA, and rescued status.
                  </li>
                  <li>
                    <strong>Step 5: Reconciliation &amp; Audit:</strong> Results are reconciled against university published gazettes to confirm exact agreement and ordinance concordance.
                  </li>
                </ol>
              </div>

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
