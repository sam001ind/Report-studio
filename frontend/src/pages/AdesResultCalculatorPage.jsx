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
  BarChart3
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
  
  // Moderation & Simulation State
  const [courseModerationMap, setCourseModerationMap] = useState({});
  const [allowPrOnlyModeration, setAllowPrOnlyModeration] = useState(false); // Moderation on solely ESE-PR courses toggle (default: false)
  const [activeTab, setActiveTab] = useState("results"); // "results" | "moderation" | "simulation"
  const [moderationSearch, setModerationSearch] = useState("");
  const [simSearchQuery, setSimSearchQuery] = useState("");
  const [bulkModValue, setBulkModValue] = useState(4);
  const modFileInputRef = useRef(null);

  // Table Filters & Pagination
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedFacultyFilter, setSelectedFacultyFilter] = useState("ALL");
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

  const getCell = (row, hMap, ...aliases) => {
    for (const alias of aliases) {
      const norm = normalizeKey(alias);
      const actualKey = hMap[norm];
      if (actualKey && row[actualKey] !== undefined && row[actualKey] !== null) {
        return row[actualKey];
      }
    }
    return "";
  };

  const parseNumber = (val) => {
    if (val === null || val === undefined) return null;
    const s = String(val).trim();
    if (s === "") return null;
    const num = Number(s);
    return isNaN(num) ? null : num;
  };

  const isAlreadyAggregatedSheet = (hMap) => {
    const keys = Object.keys(hMap);
    const hasEseOverall = keys.some(k => k.includes("eseoverall") || k.includes("esemax") || k.includes("esethobtained") || k.includes("esethmax"));
    const hasCourseOverall = keys.some(k => k.includes("courseoverall") || k.includes("overallmax") || k.includes("overallmin"));
    const hasRawAssessment = keys.some(k => k === "assessmentmethod" || k === "am" || k.includes("assessmenttype") || k === "at");
    return (hasEseOverall || hasCourseOverall) && !hasRawAssessment;
  };

  // Group raw assessment rows OR parse pre-aggregated rows into Student-Course Base Aggregates
  const buildGroupedRecordsFromRows = (rows, currentHeaderMap) => {
    const isAgg = isAlreadyAggregatedSheet(currentHeaderMap);

    if (isAgg) {
      const baseRecords = [];

      rows.forEach((row) => {
        const faculty = String(getCell(row, currentHeaderMap, "Faculty", "Fac", "FacultyName", "Department") || "").trim();
        const program = String(getCell(row, currentHeaderMap, "Program Term Name", "ProgramTermName", "ProgramTerm", "Program Term", "Degree", "Term", "Semester") || "").trim();
        const seat = String(getCell(row, currentHeaderMap, "Seat Number", "SeatNumber", "SeatNo", "Seat_Number", "RollNo", "Roll Number") || "").trim();
        const prn = String(getCell(row, currentHeaderMap, "PRN", "PRN Number", "PRNNo", "RegisterNo", "RegNo", "StudentID") || "").trim();
        const code = String(getCell(row, currentHeaderMap, "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || "").trim();
        const name = String(getCell(row, currentHeaderMap, "Course Name", "CourseName", "PaperName", "SubjectName", "CourseTitle") || "").trim();

        const ese_pr_max = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Max", "ESEPRMax")) ?? "";
        const ese_pr_min = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Min", "ESEPRMin")) ?? "";
        const ese_pr_obtained = parseNumber(getCell(row, currentHeaderMap, "ESE - PR Obtained", "ESEPRObtained")) ?? "";

        const ese_th_max = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Max", "ESETHMax")) ?? "";
        const ese_th_min = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Min", "ESETHMin")) ?? (ese_th_max !== "" ? 0 : "");
        const ese_th_obtained = parseNumber(getCell(row, currentHeaderMap, "ESE - TH Obtained", "ESETHObtained")) ?? "";

        let ese_max = parseNumber(getCell(row, currentHeaderMap, "ESE - Max", "ESEMax"));
        if (ese_max === null) {
          ese_max = (parseNumber(ese_pr_max) || 0) + (parseNumber(ese_th_max) || 0);
        }
        let ese_min = parseNumber(getCell(row, currentHeaderMap, "ESE - Min", "ESEMin"));
        if (ese_min === null) {
          ese_min = Math.ceil(0.30 * ese_max);
        }
        let ese_obtained = parseNumber(getCell(row, currentHeaderMap, "ESE Overall", "ESEOverall"));
        if (ese_obtained === null) {
          ese_obtained = (parseNumber(ese_pr_obtained) || 0) + (parseNumber(ese_th_obtained) || 0);
        }

        const ce_pr_max = parseNumber(getCell(row, currentHeaderMap, "CE - PR Max", "CEPRMax")) ?? "";
        const ce_pr_min = parseNumber(getCell(row, currentHeaderMap, "CE - PR Min", "CEPRMin")) ?? (ce_pr_max !== "" ? 0 : "");
        const ce_pr_obtained = parseNumber(getCell(row, currentHeaderMap, "CE - PR Obtained", "CEPRObtained")) ?? "";

        const ce_th_max = parseNumber(getCell(row, currentHeaderMap, "CE - TH Max", "CETHMax")) ?? "";
        const ce_th_min = parseNumber(getCell(row, currentHeaderMap, "CE - TH Min", "CETHMin")) ?? (ce_th_max !== "" ? 0 : "");
        const ce_th_obtained = parseNumber(getCell(row, currentHeaderMap, "CE - TH Obtained", "CETHObtained")) ?? "";

        let ce_max = parseNumber(getCell(row, currentHeaderMap, "CE - Max", "CEMax"));
        if (ce_max === null) {
          ce_max = (parseNumber(ce_pr_max) || 0) + (parseNumber(ce_th_max) || 0);
        }
        let ce_min = 0;
        let ce_obtained = parseNumber(getCell(row, currentHeaderMap, "CE Overall Marks ", "CE Overall Marks", "CEOverallMarks", "CEOverall"));
        if (ce_obtained === null) {
          ce_obtained = (parseNumber(ce_pr_obtained) || 0) + (parseNumber(ce_th_obtained) || 0);
        }

        let overall_max = parseNumber(getCell(row, currentHeaderMap, "Overall Maximum", "OverallMaximum", "OverallMax"));
        if (overall_max === null) {
          overall_max = ese_max + ce_max;
        }
        let overall_min = parseNumber(getCell(row, currentHeaderMap, "Overall Minimum", "OverallMinimum", "OverallMin"));
        if (overall_min === null) {
          overall_min = Math.ceil(0.35 * overall_max);
        }
        let course_overall = parseNumber(getCell(row, currentHeaderMap, "Course Overall Marks ", "Course Overall Marks", "CourseOverallMarks"));
        if (course_overall === null) {
          course_overall = ese_obtained + ce_obtained;
        }

        const ese_deficit = Math.max(0, ese_min - ese_obtained);
        const overall_deficit = Math.max(0, overall_min - course_overall);
        const raw_ese_pass = ese_deficit === 0;
        const raw_overall_pass = overall_deficit === 0;
        const raw_course_pass = raw_ese_pass && raw_overall_pass;

        const has_ese_th = (parseNumber(ese_th_max) || 0) > 0;
        const has_ese_pr = (parseNumber(ese_pr_max) || 0) > 0;
        const is_pr_only = has_ese_pr && !has_ese_th;

        baseRecords.push({
          identifiers: { faculty, program, seat, prn, code, name },
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
          raw_course_pass
        });
      });

      return baseRecords;
    }

    // Standard Raw Assessment Grouping
    const groups = new Map();

    rows.forEach((row) => {
      const faculty = String(getCell(row, currentHeaderMap, "Faculty", "Fac", "FacultyName", "Department") || "").trim();
      const program = String(getCell(row, currentHeaderMap, "Program Term Name", "ProgramTermName", "ProgramTerm", "Program Term", "Degree", "Term", "Semester") || "").trim();
      const seat = String(getCell(row, currentHeaderMap, "Seat Number", "SeatNumber", "SeatNo", "Seat_Number", "RollNo", "Roll Number") || "").trim();
      const prn = String(getCell(row, currentHeaderMap, "PRN", "PRN Number", "PRNNo", "RegisterNo", "RegNo", "StudentID") || "").trim();
      const code = String(getCell(row, currentHeaderMap, "Course Code", "CourseCode", "PaperCode", "SubjectCode", "Course") || "").trim();
      const name = String(getCell(row, currentHeaderMap, "Course Name", "CourseName", "PaperName", "SubjectName", "CourseTitle") || "").trim();

      const methodRaw = String(getCell(row, currentHeaderMap, "Assessment Method", "AssessmentMethod", "AM", "Method") || "").trim().toUpperCase();
      const typeRaw = String(getCell(row, currentHeaderMap, "Assessment Type", "AssessmentType", "AT", "Type") || "").trim().toUpperCase();
      const marksRaw = parseNumber(getCell(row, currentHeaderMap, "Marks", "ObtainedMarks", "Mark", "Obtained"));
      const atMaxRaw = parseNumber(getCell(row, currentHeaderMap, "AT Max Marks", "ATMaxMarks", "MaxMarks", "Max Marks", "Max"));

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

      const groupKey = faculty + "|||" + program + "|||" + seat + "|||" + prn + "|||" + code + "|||" + name;

      if (!groups.has(groupKey)) {
        groups.set(groupKey, {
          identifiers: { faculty, program, seat, prn, code, name },
          components: {}
        });
      }

      const group = groups.get(groupKey);
      group.components[method + "_" + type] = {
        marks: marksRaw !== null ? marksRaw : null,
        max: atMaxRaw !== null ? atMaxRaw : null,
        present: true
      };
    });

    const baseRecords = [];

    groups.forEach(({ identifiers, components }) => {
      const ese_pr = components["ESE_PR"];
      const ese_pr_max = (ese_pr && ese_pr.max !== null) ? ese_pr.max : 0;
      const ese_pr_obtained = (ese_pr && ese_pr.marks !== null) ? ese_pr.marks : 0;

      const ese_th = components["ESE_TH"];
      const ese_th_max = (ese_th && ese_th.max !== null) ? ese_th.max : 0;
      const ese_th_obtained = (ese_th && ese_th.marks !== null) ? ese_th.marks : 0;

      const ese_max = ese_pr_max + ese_th_max;
      const ese_min = Math.ceil(0.30 * ese_max);
      const ese_obtained = ese_pr_obtained + ese_th_obtained;

      const ce_pr = components["CE_PR"];
      const ce_pr_max = (ce_pr && ce_pr.max !== null) ? ce_pr.max : 0;
      const ce_pr_obtained = (ce_pr && ce_pr.marks !== null) ? ce_pr.marks : 0;

      const ce_th = components["CE_TH"];
      const ce_th_max = (ce_th && ce_th.max !== null) ? ce_th.max : 0;
      const ce_th_obtained = (ce_th && ce_th.marks !== null) ? ce_th.marks : 0;

      const ce_max = ce_pr_max + ce_th_max;
      const ce_min = 0;
      const ce_obtained = ce_pr_obtained + ce_th_obtained;

      const overall_max = ese_max + ce_max;
      const overall_min = Math.ceil(0.35 * overall_max);
      const course_overall = ese_obtained + ce_obtained;

      const ese_deficit = Math.max(0, ese_min - ese_obtained);
      const overall_deficit = Math.max(0, overall_min - course_overall);
      const raw_ese_pass = ese_deficit === 0;
      const raw_overall_pass = overall_deficit === 0;
      const raw_course_pass = raw_ese_pass && raw_overall_pass;

      const has_ese_th = ese_th_max > 0;
      const has_ese_pr = ese_pr_max > 0;
      const is_pr_only = has_ese_pr && !has_ese_th;

      baseRecords.push({
        identifiers,
        raw: {
          "Faculty": identifiers.faculty,
          "Program Term Name": identifiers.program,
          "Course Code": identifiers.code,
          "Course Name": identifiers.name,
          "Seat Number": identifiers.seat,
          "PRN": identifiers.prn,
          "ESE - PR Max": (ese_pr && ese_pr.max !== null) ? ese_pr.max : "",
          "ESE - PR Min": "",
          "ESE - PR Obtained": (ese_pr && ese_pr.marks !== null) ? ese_pr.marks : "",
          "ESE - TH Max": (ese_th && ese_th.max !== null) ? ese_th.max : "",
          "ESE - TH Min": ese_th ? 0 : "",
          "ESE - TH Obtained": (ese_th && ese_th.marks !== null) ? ese_th.marks : "",
          "ESE - Max": Math.round(ese_max),
          "ESE - Min": ese_min,
          "ESE Overall": Math.round(ese_obtained),
          "CE - PR Max": (ce_pr && ce_pr.max !== null) ? ce_pr.max : "",
          "CE - PR Min": ce_pr ? 0 : "",
          "CE - PR Obtained": (ce_pr && ce_pr.marks !== null) ? ce_pr.marks : "",
          "CE - TH Max": (ce_th && ce_th.max !== null) ? ce_th.max : "",
          "CE - TH Min": ce_th ? 0 : "",
          "CE - TH Obtained": (ce_th && ce_th.marks !== null) ? ce_th.marks : "",
          "CE - Max": Math.round(ce_max),
          "CE - Min": ce_min,
          "CE Overall Marks ": Math.round(ce_obtained),
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
        raw_course_pass
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
          hasEseTh: rec.has_ese_th,
          hasEsePr: rec.has_ese_pr,
          isPrOnly: rec.is_pr_only,
          totalStudents: 0,
          rawPassed: 0,
          rawFailed: 0,
          nearPassCount: 0
        });
      }

      const item = map.get(norm);
      item.totalStudents++;
      if (rec.raw_course_pass) {
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
      const normCode = normalizeKey(rec.identifiers.code);
      const modLimit = courseModerationMap[normCode] || 0;

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
      row._modLimit = modLimit;
      row._isEligibleForMod = isEligibleForModeration;
      row._isPrOnly = rec.is_pr_only;
      row._eseDeficit = rec.ese_deficit;
      row._overallDeficit = rec.overall_deficit;

      return row;
    });
  }, [groupedRecords, courseModerationMap, allowPrOnlyModeration]);

  // Course-Wise Pass Simulation (0 to +10 Moderation Marks)
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
          hasEseTh: rec.has_ese_th,
          hasEsePr: rec.has_ese_pr,
          isPrOnly: rec.is_pr_only,
          isEligible: rec.has_ese_th || (rec.is_pr_only && allowPrOnlyModeration),
          totalStudents: 0,
          rawPassCount: 0,
          passCountAtMod: Array(11).fill(0), // indices 0 to 10
        });
      }

      const item = map.get(norm);
      item.totalStudents++;

      const ese_deficit = rec.ese_deficit;
      const overall_deficit = rec.overall_deficit;
      const marks_needed = Math.max(ese_deficit, overall_deficit);
      const is_raw_pass = rec.raw_course_pass;

      if (is_raw_pass) {
        item.rawPassCount++;
      }

      const isEligible = rec.has_ese_th || (rec.is_pr_only && allowPrOnlyModeration);

      // Evaluate simulated pass for each moderation mark level from 0 to 10
      for (let m = 0; m <= 10; m++) {
        if (is_raw_pass) {
          item.passCountAtMod[m]++;
        } else if (isEligible && marks_needed <= m) {
          item.passCountAtMod[m]++;
        }
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

  // Overall Simulation Totals across filtered courses
  const simTotals = useMemo(() => {
    let totalStudents = 0;
    let rawPass = 0;
    const modPass = Array(11).fill(0);

    filteredSimulationCourses.forEach(c => {
      totalStudents += c.totalStudents;
      rawPass += c.rawPassCount;
      for (let m = 0; m <= 10; m++) {
        modPass[m] += c.passCountAtMod[m];
      }
    });

    return {
      totalCourses: filteredSimulationCourses.length,
      totalStudents,
      rawPass,
      rawPassPct: totalStudents > 0 ? ((rawPass / totalStudents) * 100).toFixed(1) : "0.0",
      modPass,
      modPassPct: (m) => totalStudents > 0 ? ((modPass[m] / totalStudents) * 100).toFixed(1) : "0.0",
      rescuedAtMod: (m) => modPass[m] - rawPass
    };
  }, [filteredSimulationCourses]);

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

        const baseGrouped = buildGroupedRecordsFromRows(rows, hMap);
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

      const baseGrouped = buildGroupedRecordsFromRows(rows, hMap);
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
              courseCode = String(val).trim();
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
      "Normal Pass (0 Mod)",
      "Normal Pass %",
      "+1 Mod Pass",
      "+2 Mod Pass",
      "+3 Mod Pass",
      "+4 Mod Pass",
      "+5 Mod Pass",
      "+6 Mod Pass",
      "+7 Mod Pass",
      "+8 Mod Pass",
      "+9 Mod Pass",
      "+10 Mod Pass",
      "+10 Mod Pass %",
      "Max Rescued (+10)"
    ];

    let totalAllStudents = 0;
    let totalAllRawPass = 0;
    const totalAllModPass = Array(11).fill(0);

    const rows = courseSimulationData.map(c => {
      totalAllStudents += c.totalStudents;
      totalAllRawPass += c.rawPassCount;
      for (let m = 1; m <= 10; m++) {
        totalAllModPass[m] += c.passCountAtMod[m];
      }

      const rawPct = c.totalStudents > 0 ? ((c.rawPassCount / c.totalStudents) * 100).toFixed(2) + "%" : "0.00%";
      const plus10Pct = c.totalStudents > 0 ? ((c.passCountAtMod[10] / c.totalStudents) * 100).toFixed(2) + "%" : "0.00%";
      const maxRescued = c.passCountAtMod[10] - c.rawPassCount;

      const compType = c.hasEseTh ? "ESE-TH (+PR)" : (c.isPrOnly ? "ESE-PR Only" : "Other");

      return [
        c.faculty,
        c.program,
        c.courseCode,
        c.courseName,
        compType,
        c.totalStudents,
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
      { wch: 18 }, // Normal Pass (0 Mod)
      { wch: 14 }, // Normal Pass %
      { wch: 13 }, // +1 Mod Pass
      { wch: 13 }, // +2 Mod Pass
      { wch: 13 }, // +3 Mod Pass
      { wch: 13 }, // +4 Mod Pass
      { wch: 13 }, // +5 Mod Pass
      { wch: 13 }, // +6 Mod Pass
      { wch: 13 }, // +7 Mod Pass
      { wch: 13 }, // +8 Mod Pass
      { wch: 13 }, // +9 Mod Pass
      { wch: 14 }, // +10 Mod Pass
      { wch: 16 }, // +10 Mod Pass %
      { wch: 16 }  // Max Rescued (+10)
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Course_Pass_Simulation");
    XLSX.writeFile(wb, "course_wise_pass_simulation.xlsx");
    setStatus("Generated & downloaded Course-Wise Pass Simulation Report (+0 to +10 Moderation).", "success");
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
    setSelectedProgramFilter("ALL");
    setSelectedCourseFilter("ALL");
    setSelectedResultFilter("ALL");
    setSearchQuery("");
    setPage(0);
  };

  // Unique Lists for Dropdown Filters
  const uniqueFaculties = useMemo(() => {
    const set = new Set();
    processedRows.forEach(r => { if (r["Faculty"]) set.add(r["Faculty"]); });
    return Array.from(set).sort();
  }, [processedRows]);

  const uniquePrograms = useMemo(() => {
    const set = new Set();
    processedRows.forEach(r => { if (r["Program Term Name"]) set.add(r["Program Term Name"]); });
    return Array.from(set).sort();
  }, [processedRows]);

  const uniqueCourses = useMemo(() => {
    const set = new Set();
    processedRows.forEach(r => { if (r["Course Code"]) set.add(r["Course Code"]); });
    return Array.from(set).sort();
  }, [processedRows]);

  // Statistics Metrics
  const metrics = useMemo(() => {
    const total = processedRows.length;
    if (total === 0) return { total: 0, uniqueStudents: 0, rawPassed: 0, moderatedPassed: 0, totalPassed: 0, failed: 0, passPct: 0, rawPassPct: 0, eseFailed: 0, overallFailed: 0 };

    const prnSet = new Set();
    let rawPassed = 0;
    let moderatedPassed = 0;
    let totalPassed = 0;
    let failed = 0;
    let eseFailed = 0;
    let overallFailed = 0;

    const coursePassKey = "Course Pass/Fail";
    const esePassKey = "ESE Pass";
    const overallPassKey = "Overall pass";

    processedRows.forEach(r => {
      if (r["PRN"]) prnSet.add(r["PRN"]);
      
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

    return {
      total,
      uniqueStudents: prnSet.size,
      rawPassed,
      moderatedPassed,
      totalPassed,
      failed,
      passPct: ((totalPassed / total) * 100).toFixed(1),
      rawPassPct: ((rawPassed / total) * 100).toFixed(1),
      eseFailed,
      overallFailed
    };
  }, [processedRows]);

  // Filtered & Sorted Rows
  const filteredRows = useMemo(() => {
    let result = [...processedRows];

    const coursePassKey = "Course Pass/Fail";
    const esePassKey = "ESE Pass";
    const overallPassKey = "Overall pass";

    // 1. Result Status Filter
    if (selectedResultFilter === "PASS") {
      result = result.filter(r => r[coursePassKey] === "Pass");
    } else if (selectedResultFilter === "PASS_MOD") {
      result = result.filter(r => r._isModeratedPass);
    } else if (selectedResultFilter === "FAIL") {
      result = result.filter(r => r[coursePassKey] === "Fail");
    } else if (selectedResultFilter === "ESE_FAIL") {
      result = result.filter(r => r[esePassKey] === "Fail");
    } else if (selectedResultFilter === "OVERALL_FAIL") {
      result = result.filter(r => r[overallPassKey] === "Fail");
    }

    // 2. Dropdown Filters
    if (selectedFacultyFilter !== "ALL") {
      result = result.filter(r => r["Faculty"] === selectedFacultyFilter);
    }
    if (selectedProgramFilter !== "ALL") {
      result = result.filter(r => r["Program Term Name"] === selectedProgramFilter);
    }
    if (selectedCourseFilter !== "ALL") {
      result = result.filter(r => r["Course Code"] === selectedCourseFilter);
    }

    // 3. Search Query
    if (searchQuery) {
      const q = searchQuery.toLowerCase().trim();
      result = result.filter(r =>
        Object.values(r).some(val => String(val || "").toLowerCase().includes(q))
      );
    }

    // 4. Column Filters
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

    // 5. Sorting
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
  }, [processedRows, selectedResultFilter, selectedFacultyFilter, selectedProgramFilter, selectedCourseFilter, searchQuery, columnFilters, sortConfig]);

  const pagedRows = useMemo(() => {
    const start = page * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, page]);

  // Export to Excel Matching exact 31 columns structure with "Output file " sheet name
  const handleExportExcel = (rowsToExport = filteredRows, filename = "converted_output_card.xlsx") => {
    if (!rowsToExport || rowsToExport.length === 0) {
      alert("No rows to export.");
      return;
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
              <Table size={13} /> Results & Grid
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
                <Download size={14} /> Export Simulation (+0 to +10 XLSX)
              </button>
              <button 
                type="button" 
                onClick={() => handleExportExcel(filteredRows, "converted_output_card.xlsx")}
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
              >
                <Download size={14} /> Export 31-Col XLSX ({filteredRows.length} Rows)
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

          {/* Result Overview Stat Card */}
          {processedRows.length > 0 && (
            <div style={{ background: "var(--bg)", border: "1px solid var(--line)", borderRadius: "8px", padding: "12px", display: "flex", flexDirection: "column", gap: "10px" }}>
              <div style={{ fontSize: "12px", fontWeight: 700, color: "var(--ink)", display: "flex", alignItems: "center", gap: "6px" }}>
                <Sparkles size={14} color="var(--accent)" /> Evaluation Metrics
              </div>
              
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px", fontSize: "11px" }}>
                <div style={{ background: "var(--panel)", padding: "8px", borderRadius: "6px", border: "1px solid var(--line)" }}>
                  <div style={{ color: "var(--muted)" }}>Evaluated Courses</div>
                  <strong style={{ fontSize: "15px", color: "var(--ink)" }}>{metrics.total}</strong>
                </div>

                <div style={{ background: "var(--panel)", padding: "8px", borderRadius: "6px", border: "1px solid var(--line)" }}>
                  <div style={{ color: "var(--muted)" }}>Unique Students</div>
                  <strong style={{ fontSize: "15px", color: "var(--ink)" }}>{metrics.uniqueStudents}</strong>
                </div>

                <div style={{ background: "rgba(16, 185, 129, 0.1)", padding: "8px", borderRadius: "6px", border: "1px solid rgba(16, 185, 129, 0.3)" }}>
                  <div style={{ color: "#10b981", fontWeight: 600 }}>Total Passed</div>
                  <strong style={{ fontSize: "15px", color: "#10b981" }}>{metrics.totalPassed}</strong>
                  <div style={{ fontSize: "10px", color: "#10b981" }}>({metrics.passPct}%)</div>
                </div>

                <div style={{ background: "rgba(239, 68, 68, 0.1)", padding: "8px", borderRadius: "6px", border: "1px solid rgba(239, 68, 68, 0.3)" }}>
                  <div style={{ color: "#ef4444", fontWeight: 600 }}>Failed Courses</div>
                  <strong style={{ fontSize: "15px", color: "#ef4444" }}>{metrics.failed}</strong>
                  <div style={{ fontSize: "10px", color: "#ef4444" }}>({(100 - metrics.passPct).toFixed(1)}%)</div>
                </div>
              </div>

              {/* Moderation Impact Breakdown */}
              <div style={{ background: "var(--panel)", padding: "8px", borderRadius: "6px", border: "1px solid var(--line)", fontSize: "11px", display: "flex", flexDirection: "column", gap: "4px" }}>
                <div style={{ fontWeight: 600, color: "var(--ink)", display: "flex", alignItems: "center", gap: "4px" }}>
                  <Zap size={12} color="#f59e0b" /> Moderation Impact:
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)" }}>
                  <span>Raw Passed:</span>
                  <strong>{metrics.rawPassed} ({metrics.rawPassPct}%)</strong>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", color: "#10b981", fontWeight: 600 }}>
                  <span>Rescued via Moderation:</span>
                  <span>+{metrics.moderatedPassed}</span>
                </div>
                <div style={{ fontSize: "10px", color: "var(--muted)", marginTop: "2px" }}>
                  Rule: {allowPrOnlyModeration ? "Applied to TH & PR-Only" : "Restricted to ESE-TH courses"}
                </div>
              </div>

              {/* Failure Breakdown */}
              {metrics.failed > 0 && (
                <div style={{ background: "var(--panel)", padding: "8px", borderRadius: "6px", border: "1px solid var(--line)", fontSize: "11px", display: "flex", flexDirection: "column", gap: "4px" }}>
                  <div style={{ fontWeight: 600, color: "var(--ink)" }}>Failure Breakdown:</div>
                  <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)" }}>
                    <span>ESE Failed (&lt; 30%):</span>
                    <strong style={{ color: "#ef4444" }}>{metrics.eseFailed}</strong>
                  </div>
                  <div style={{ display: "flex", justifyContent: "space-between", color: "var(--muted)" }}>
                    <span>Aggregate Failed (&lt; 35%):</span>
                    <strong style={{ color: "#ef4444" }}>{metrics.overallFailed}</strong>
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
                    onClick={() => setActiveTab("simulation")}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      gap: "4px",
                      padding: "5px 8px",
                      fontSize: "11px",
                      background: "var(--bg)",
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
          
          {activeTab === "moderation" ? (
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

              {/* Course Search in Moderation Tab */}
              <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                <div style={{ position: "relative", flex: 1, maxWidth: "400px" }}>
                  <Search size={14} style={{ position: "absolute", left: "10px", top: "50%", transform: "translateY(-50%)", color: "var(--muted)" }} />
                  <input 
                    type="text"
                    placeholder="Search courses by code or title..."
                    value={moderationSearch}
                    onChange={(e) => setModerationSearch(e.target.value)}
                    style={{ width: "100%", padding: "6px 10px 6px 30px", fontSize: "12px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--panel)" }}
                  />
                </div>
                <span style={{ fontSize: "12px", color: "var(--muted)" }}>
                  Showing {distinctCourses.filter(c => !moderationSearch || c.courseCode.toLowerCase().includes(moderationSearch.toLowerCase()) || c.courseName.toLowerCase().includes(moderationSearch.toLowerCase())).length} of {distinctCourses.length} unique courses
                </span>
              </div>

              {/* Moderation Course List Table */}
              <div style={{ flex: 1, overflow: "auto", background: "var(--panel)", border: "1px solid var(--line)", borderRadius: "8px" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px", textAlign: "left" }}>
                  <thead style={{ position: "sticky", top: 0, background: "var(--panel)", borderBottom: "1px solid var(--line)", zIndex: 10 }}>
                    <tr>
                      <th style={{ padding: "8px 12px", color: "var(--muted)", fontWeight: 600 }}>Course Code</th>
                      <th style={{ padding: "8px 12px", color: "var(--muted)", fontWeight: 600 }}>Course Name</th>
                      <th style={{ padding: "8px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center" }}>Component Type</th>
                      <th style={{ padding: "8px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center" }}>Total Enrolled</th>
                      <th style={{ padding: "8px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center" }}>Raw Failed</th>
                      <th style={{ padding: "8px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center" }}>Deficit &le; 5 Marks</th>
                      <th style={{ padding: "8px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center", width: "150px" }}>Current Moderation Marks</th>
                      <th style={{ padding: "8px 12px", color: "var(--muted)", fontWeight: 600, textAlign: "center" }}>Rescued with Mod</th>
                    </tr>
                  </thead>
                  <tbody>
                    {distinctCourses.length === 0 ? (
                      <tr>
                        <td colSpan={8} style={{ padding: "30px", textAlign: "center", color: "var(--muted)" }}>
                          No courses loaded. Upload an ADES Marks Excel sheet to begin.
                        </td>
                      </tr>
                    ) : (
                      distinctCourses
                        .filter(c => !moderationSearch || c.courseCode.toLowerCase().includes(moderationSearch.toLowerCase()) || c.courseName.toLowerCase().includes(moderationSearch.toLowerCase()))
                        .map(c => {
                          const currentMod = courseModerationMap[c.normCode] || 0;
                          const courseRows = processedRows.filter(r => normalizeKey(r["Course Code"]) === c.normCode);
                          const rescuedInCourse = courseRows.filter(r => r._isModeratedPass).length;
                          const isPrOnly = c.isPrOnly;

                          return (
                            <tr key={c.normCode} style={{ borderBottom: "1px solid var(--line)", background: currentMod > 0 ? "rgba(245, 158, 11, 0.04)" : "transparent" }}>
                              <td style={{ padding: "8px 12px", fontWeight: 600, color: "var(--ink)" }}>{c.courseCode}</td>
                              <td style={{ padding: "8px 12px", color: "var(--ink)" }}>{c.courseName}</td>
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
                              <td style={{ padding: "8px 12px", textAlign: "center" }}>{c.totalStudents}</td>
                              <td style={{ padding: "8px 12px", textAlign: "center", color: c.rawFailed > 0 ? "#ef4444" : "var(--muted)", fontWeight: c.rawFailed > 0 ? 600 : 400 }}>
                                {c.rawFailed}
                              </td>
                              <td style={{ padding: "8px 12px", textAlign: "center", color: "#f59e0b", fontWeight: 600 }}>
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
                                      width: "60px", 
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

              {/* KPI Summary Cards */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(170px, 1fr))", gap: "12px" }}>
                <div style={{ background: "var(--panel)", padding: "12px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                  <div style={{ fontSize: "11px", color: "var(--muted)", fontWeight: 600 }}>Total Courses</div>
                  <div style={{ fontSize: "18px", fontWeight: 700, color: "var(--ink)", marginTop: "2px" }}>{simTotals.totalCourses}</div>
                  <div style={{ fontSize: "11px", color: "var(--muted)", marginTop: "2px" }}>{simTotals.totalStudents} total student entries</div>
                </div>

                <div style={{ background: "var(--panel)", padding: "12px", borderRadius: "8px", border: "1px solid var(--line)" }}>
                  <div style={{ fontSize: "11px", color: "var(--muted)", fontWeight: 600 }}>Normal Pass (0 Mod)</div>
                  <div style={{ fontSize: "18px", fontWeight: 700, color: "#3b82f6", marginTop: "2px" }}>{simTotals.rawPass}</div>
                  <div style={{ fontSize: "11px", color: "#3b82f6", fontWeight: 600, marginTop: "2px" }}>{simTotals.rawPassPct}% raw pass rate</div>
                </div>

                <div style={{ background: "rgba(245, 158, 11, 0.08)", padding: "12px", borderRadius: "8px", border: "1px solid rgba(245, 158, 11, 0.25)" }}>
                  <div style={{ fontSize: "11px", color: "#f59e0b", fontWeight: 600 }}>Pass at +3 Mod</div>
                  <div style={{ fontSize: "18px", fontWeight: 700, color: "#f59e0b", marginTop: "2px" }}>{simTotals.modPass[3]}</div>
                  <div style={{ fontSize: "11px", color: "#f59e0b", fontWeight: 600, marginTop: "2px" }}>{simTotals.modPassPct(3)}% (+{simTotals.rescuedAtMod(3)} rescued)</div>
                </div>

                <div style={{ background: "rgba(245, 158, 11, 0.12)", padding: "12px", borderRadius: "8px", border: "1px solid rgba(245, 158, 11, 0.35)" }}>
                  <div style={{ fontSize: "11px", color: "#d97706", fontWeight: 600 }}>Pass at +5 Mod</div>
                  <div style={{ fontSize: "18px", fontWeight: 700, color: "#d97706", marginTop: "2px" }}>{simTotals.modPass[5]}</div>
                  <div style={{ fontSize: "11px", color: "#d97706", fontWeight: 600, marginTop: "2px" }}>{simTotals.modPassPct(5)}% (+{simTotals.rescuedAtMod(5)} rescued)</div>
                </div>

                <div style={{ background: "rgba(16, 185, 129, 0.12)", padding: "12px", borderRadius: "8px", border: "1px solid rgba(16, 185, 129, 0.35)" }}>
                  <div style={{ fontSize: "11px", color: "#10b981", fontWeight: 600 }}>Pass at +10 Mod</div>
                  <div style={{ fontSize: "18px", fontWeight: 700, color: "#10b981", marginTop: "2px" }}>{simTotals.modPass[10]}</div>
                  <div style={{ fontSize: "11px", color: "#10b981", fontWeight: 600, marginTop: "2px" }}>{simTotals.modPassPct(10)}% (+{simTotals.rescuedAtMod(10)} rescued)</div>
                </div>
              </div>

              {/* Simulation Table Container */}
              <div style={{ flex: 1, overflow: "auto", border: "1px solid var(--line)", borderRadius: "8px", background: "var(--panel)" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
                  <thead style={{ position: "sticky", top: 0, background: "var(--bg)", zIndex: 10, borderBottom: "2px solid var(--line)" }}>
                    <tr>
                      <th style={{ padding: "10px 12px", textAlign: "left", color: "var(--muted)", fontWeight: 600, width: "120px" }}>Course Code</th>
                      <th style={{ padding: "10px 12px", textAlign: "left", color: "var(--muted)", fontWeight: 600, minWidth: "180px" }}>Course Name</th>
                      <th style={{ padding: "10px 12px", textAlign: "center", color: "var(--muted)", fontWeight: 600, width: "100px" }}>Type</th>
                      <th style={{ padding: "10px 12px", textAlign: "center", color: "var(--muted)", fontWeight: 600, width: "70px" }}>Total</th>
                      <th style={{ padding: "10px 12px", textAlign: "center", color: "var(--ink)", fontWeight: 700, width: "85px", background: "rgba(59, 130, 246, 0.08)" }}>Normal (0)</th>
                      {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(m => (
                        <th 
                          key={m} 
                          style={{ 
                            padding: "10px 8px", 
                            textAlign: "center", 
                            color: m === 5 ? "#d97706" : m === 10 ? "#10b981" : "var(--muted)", 
                            fontWeight: m === 5 || m === 10 ? 700 : 600, 
                            width: "55px",
                            background: m === 5 ? "rgba(245, 158, 11, 0.1)" : m === 10 ? "rgba(16, 185, 129, 0.1)" : "transparent"
                          }}
                        >
                          +{m}
                        </th>
                      ))}
                      <th style={{ padding: "10px 12px", textAlign: "center", color: "#10b981", fontWeight: 700, width: "100px", background: "rgba(16, 185, 129, 0.08)" }}>Max Gain (+10)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredSimulationCourses.length === 0 ? (
                      <tr>
                        <td colSpan={16} style={{ padding: "32px", textAlign: "center", color: "var(--muted)" }}>
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
                              {c.hasEseTh ? (
                                <span style={{ fontSize: "10.5px", padding: "2px 6px", borderRadius: "4px", background: "rgba(59, 130, 246, 0.12)", color: "#3b82f6", fontWeight: 600 }}>
                                  ESE-TH
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
                                <span style={{ fontSize: "10.5px", color: "var(--muted)" }}>-</span>
                              )}
                            </td>
                            <td style={{ padding: "8px 12px", textAlign: "center", fontWeight: 600 }}>{c.totalStudents}</td>
                            <td style={{ padding: "8px 12px", textAlign: "center", background: "rgba(59, 130, 246, 0.04)" }}>
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
                                    padding: "8px 6px", 
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
                        <td style={{ padding: "10px 12px", textAlign: "center", color: "#3b82f6" }}>
                          {simTotals.rawPass} ({simTotals.rawPassPct}%)
                        </td>
                        {[1, 2, 3, 4, 5, 6, 7, 8, 9, 10].map(m => (
                          <td 
                            key={m} 
                            style={{ 
                              padding: "10px 6px", 
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
                    </div>

                    {/* Faculty Filter */}
                    {uniqueFaculties.length > 1 && (
                      <select 
                        value={selectedFacultyFilter} 
                        onChange={(e) => { setSelectedFacultyFilter(e.target.value); setPage(0); }}
                        style={{ padding: "4px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)" }}
                      >
                        <option value="ALL">All Faculties</option>
                        {uniqueFaculties.map(f => <option key={f} value={f}>{f}</option>)}
                      </select>
                    )}

                    {/* Program Filter */}
                    {uniquePrograms.length > 1 && (
                      <select 
                        value={selectedProgramFilter} 
                        onChange={(e) => { setSelectedProgramFilter(e.target.value); setPage(0); }}
                        style={{ padding: "4px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)" }}
                      >
                        <option value="ALL">All Programs</option>
                        {uniquePrograms.map(p => <option key={p} value={p}>{p}</option>)}
                      </select>
                    )}

                    {/* Course Filter */}
                    {uniqueCourses.length > 1 && (
                      <select 
                        value={selectedCourseFilter} 
                        onChange={(e) => { setSelectedCourseFilter(e.target.value); setPage(0); }}
                        style={{ padding: "4px 8px", fontSize: "11.5px", borderRadius: "6px", border: "1px solid var(--line)", background: "var(--bg)" }}
                      >
                        <option value="ALL">All Courses</option>
                        {uniqueCourses.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    )}

                    {(Object.keys(columnFilters).length > 0 || sortConfig.column || searchQuery || selectedResultFilter !== "ALL" || selectedFacultyFilter !== "ALL" || selectedProgramFilter !== "ALL" || selectedCourseFilter !== "ALL") && (
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
                        const isPass = row[coursePassKey] === "Pass";
                        const isModPass = row._isModeratedPass;

                        return (
                          <tr 
                            key={globalIdx} 
                            style={{ 
                              borderBottom: "1px solid var(--line)", 
                              background: isModPass 
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

                              if (isCoursePass) {
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center" }}>
                                    {isModPass ? (
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

                              if (isOtherPass) {
                                const passVal = String(val);
                                return (
                                  <td key={col} style={{ padding: "5px 8px", borderRight: "1px solid var(--line)", textAlign: "center" }}>
                                    <span style={{ 
                                      color: passVal === "Pass" ? "#10b981" : "#ef4444",
                                      fontWeight: 600,
                                      background: passVal === "Pass" ? "rgba(16, 185, 129, 0.08)" : "rgba(239, 68, 68, 0.08)",
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
