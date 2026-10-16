import { useState, useMemo, useCallback } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { 
  ArrowLeft, 
  Upload, 
  Download, 
  FileSpreadsheet, 
  Sparkles, 
  Search, 
  HelpCircle, 
  X, 
  RotateCcw, 
  CheckCircle2, 
  FileText,
  Sliders,
  ChevronLeft,
  ChevronRight,
  Database,
  GraduationCap
} from 'lucide-react';
import { parseWorkbookFromBuffer } from '../utils/excelParser';

export const SAMPLE_FYIMP_WIDE_DATA = [
  {
    "ApplicationNo": "2024CAP001",
    "CandidateName": "STUDENT ALPHA",
    "Gender": "Female",
    "Mobile": "9876543210",
    "CollegeCode": "AA",
    "CollegeName": "Alpha Arts & Science College",
    "Program": "B.A. History (FYIMP)",
    "Part": "1",
    "Term": "1",
    "AllottedCategory": "General",
    "ParentGroupName1": "Major Core",
    "ImmediateParentGroupName1": "Discipline Specific Core",
    "CourseCode1": "KU4DSCHIS101",
    "CourseName1": "Early Indian History and Culture",
    "ParentGroupName2": "Minor Group",
    "ImmediateParentGroupName2": "Minor Stream A",
    "CourseCode2": "KU4MINPOL101",
    "CourseName2": "Principles of Political Theory",
    "ParentGroupName3": "Multi-Disciplinary",
    "ImmediateParentGroupName3": "Open Elective",
    "CourseCode3": "KU4MDCECO101",
    "CourseName3": "Introductory Microeconomics",
    "ParentGroupName4": "AEC Language",
    "ImmediateParentGroupName4": "English Communication",
    "CourseCode4": "KU4AECENG101",
    "CourseName4": "English Fluency & Academic Writing",
    "ParentGroupName5": "",
    "ImmediateParentGroupName5": "",
    "CourseCode5": "",
    "CourseName5": ""
  },
  {
    "ApplicationNo": "2024CAP002",
    "CandidateName": "STUDENT BETA",
    "Gender": "Male",
    "Mobile": "9876543211",
    "CollegeCode": "AA",
    "CollegeName": "Alpha Arts & Science College",
    "Program": "B.A. English (FYIMP)",
    "Part": "1",
    "Term": "1",
    "AllottedCategory": "OBC",
    "ParentGroupName1": "Major Core",
    "ImmediateParentGroupName1": "Discipline Specific Core",
    "CourseCode1": "KU4DSCENG101",
    "CourseName1": "Introduction to Literary Studies",
    "ParentGroupName2": "Minor Group",
    "ImmediateParentGroupName2": "Minor Stream B",
    "CourseCode2": "KU4MINSOC101",
    "CourseName2": "Foundations of Sociology",
    "ParentGroupName3": "",
    "ImmediateParentGroupName3": "General Psychology Minor",
    "CourseCode3": "KU4MDCPSY101",
    "CourseName3": "General Psychology",
    "ParentGroupName4": "",
    "ImmediateParentGroupName4": "",
    "CourseCode4": "",
    "CourseName4": "",
    "ParentGroupName5": "",
    "ImmediateParentGroupName5": "",
    "CourseCode5": "",
    "CourseName5": ""
  },
  {
    "ApplicationNo": "2024CAP003",
    "CandidateName": "STUDENT GAMMA",
    "Gender": "Female",
    "Mobile": "9876543212",
    "CollegeCode": "BB",
    "CollegeName": "Beta Commerce College",
    "Program": "B.Com. Finance (FYIMP)",
    "Part": "1",
    "Term": "1",
    "AllottedCategory": "SC",
    "ParentGroupName1": "Major Core",
    "ImmediateParentGroupName1": "Discipline Specific Core",
    "CourseCode1": "KU4DSCCOM101",
    "CourseName1": "Financial Accounting Fundamentals",
    "ParentGroupName2": "Minor Group",
    "ImmediateParentGroupName2": "Commerce Minor",
    "CourseCode2": "KU4MINMGT101",
    "CourseName2": "Management Principles & Practice",
    "ParentGroupName3": "Skill Enhancement",
    "ImmediateParentGroupName3": "SEC Core",
    "CourseCode3": "KU4SECIT101",
    "CourseName3": "Computerised Financial Accounting",
    "ParentGroupName4": "AEC Language",
    "ImmediateParentGroupName4": "Modern Indian Language",
    "CourseCode4": "KU4AECMAL101",
    "CourseName4": "Malayalam Prose & Poetry",
    "ParentGroupName5": "Value Added",
    "ImmediateParentGroupName5": "VAC Environmental",
    "CourseCode5": "KU4VACENV101",
    "CourseName5": "Environmental Studies & Disaster Management"
  },
  {
    "ApplicationNo": "2024CAP004",
    "CandidateName": "STUDENT DELTA",
    "Gender": "Male",
    "Mobile": "9876543213",
    "CollegeCode": "CC",
    "CollegeName": "Gamma Government College",
    "Program": "B.Sc. Physics (FYIMP)",
    "Part": "1",
    "Term": "1",
    "AllottedCategory": "General",
    "ParentGroupName1": "Major Core",
    "ImmediateParentGroupName1": "Discipline Specific Core",
    "CourseCode1": "KU4DSCPHY101",
    "CourseName1": "Mechanics and Properties of Matter",
    "ParentGroupName2": "Minor Group",
    "ImmediateParentGroupName2": "Mathematics Minor",
    "CourseCode2": "KU4MINMAT101",
    "CourseName2": "Differential Calculus & Analytical Geometry",
    "ParentGroupName3": "",
    "ImmediateParentGroupName3": "",
    "CourseCode3": "",
    "CourseName3": "",
    "ParentGroupName4": "",
    "ImmediateParentGroupName4": "",
    "CourseCode4": "",
    "CourseName4": "",
    "ParentGroupName5": "",
    "ImmediateParentGroupName5": "",
    "CourseCode5": "",
    "CourseName5": ""
  },
  {
    "ApplicationNo": "2024CAP005",
    "CandidateName": "STUDENT EPSILON",
    "Gender": "Male",
    "Mobile": "9876543214",
    "CollegeCode": "AA",
    "CollegeName": "Alpha Arts & Science College",
    "Program": "B.A. History (FYIMP)",
    "Part": "1",
    "Term": "1",
    "AllottedCategory": "General",
    "ParentGroupName1": "Major Core",
    "ImmediateParentGroupName1": "Discipline Specific Core",
    "CourseCode1": "KU4DSCHIS101",
    "CourseName1": "Early Indian History and Culture",
    "ParentGroupName2": "Minor Group",
    "ImmediateParentGroupName2": "Minor Stream A",
    "CourseCode2": "KU4MINPOL101",
    "CourseName2": "Principles of Political Theory",
    "ParentGroupName3": "Multi-Disciplinary",
    "ImmediateParentGroupName3": "Open Elective",
    "CourseCode3": "KU4MDCECO101",
    "CourseName3": "Introductory Microeconomics",
    "ParentGroupName4": "",
    "ImmediateParentGroupName4": "",
    "CourseCode4": "",
    "CourseName4": "",
    "ParentGroupName5": "",
    "ImmediateParentGroupName5": "",
    "CourseCode5": "",
    "CourseName5": ""
  }
];

/**
 * Intelligent helper to detect column names based on candidate aliases
 */
export const detectColumnMatch = (columns, candidateNames) => {
  if (!columns || columns.length === 0) return '';
  const lowerCols = columns.map(c => ({
    original: c,
    clean: String(c).trim().toLowerCase().replace(/[\s_]+/g, '')
  }));

  // Exact clean match first
  for (const cand of candidateNames) {
    const cleanCand = cand.toLowerCase().replace(/[\s_]+/g, '');
    const found = lowerCols.find(c => c.clean === cleanCand);
    if (found) return found.original;
  }

  // Substring match fallback
  for (const cand of candidateNames) {
    const cleanCand = cand.toLowerCase().replace(/[\s_]+/g, '');
    const found = lowerCols.find(c => c.clean.includes(cleanCand) || cleanCand.includes(c.clean));
    if (found) return found.original;
  }
  return '';
};

/**
 * Generates a unique course master report containing:
 * Program, Part, Term, ParentGroupName, ImmediateParentGroupName, CourseCode, CourseName
 */
export const generateUniqueCourseReport = (transformedRows, options = {}) => {
  const {
    programCol = '',
    partCol = '',
    termCol = '',
    fallbackProgram = 'Default Program',
    fallbackPart = '1',
    fallbackTerm = '1'
  } = options;

  if (!transformedRows || !Array.isArray(transformedRows)) return [];

  const uniqueMap = new Map();

  for (const row of transformedRows) {
    const courseCode = (row['CourseCode'] || '').trim();
    const courseName = (row['CourseName'] || '').trim();

    // Skip empty course slots
    if (!courseCode && !courseName) continue;

    let program = programCol && row[programCol] !== undefined ? String(row[programCol]).trim() : '';
    if (!program) program = fallbackProgram || 'Default Program';

    let part = partCol && row[partCol] !== undefined ? String(row[partCol]).trim() : '';
    if (!part) part = fallbackPart || '1';

    let term = termCol && row[termCol] !== undefined ? String(row[termCol]).trim() : '';
    if (!term) term = fallbackTerm || '1';

    const parentGroupName = (row['ParentGroupName'] || '').trim();
    const immediateParentGroupName = (row['ImmediateParentGroupName'] || '').trim();

    const dedupKey = [
      program.toLowerCase(),
      part.toLowerCase(),
      term.toLowerCase(),
      parentGroupName.toLowerCase(),
      immediateParentGroupName.toLowerCase(),
      courseCode.toLowerCase(),
      courseName.toLowerCase()
    ].join('|||');

    if (!uniqueMap.has(dedupKey)) {
      uniqueMap.set(dedupKey, {
        Program: program,
        Part: part,
        Term: term,
        ParentGroupName: parentGroupName,
        ImmediateParentGroupName: immediateParentGroupName,
        CourseCode: courseCode,
        CourseName: courseName
      });
    }
  }

  const list = Array.from(uniqueMap.values());
  list.sort((a, b) => {
    const pCmp = a.Program.localeCompare(b.Program);
    if (pCmp !== 0) return pCmp;
    const ptCmp = a.Part.localeCompare(b.Part, undefined, { numeric: true });
    if (ptCmp !== 0) return ptCmp;
    const tCmp = a.Term.localeCompare(b.Term, undefined, { numeric: true });
    if (tCmp !== 0) return tCmp;
    const pgCmp = a.ParentGroupName.localeCompare(b.ParentGroupName);
    if (pgCmp !== 0) return pgCmp;
    const ipgCmp = a.ImmediateParentGroupName.localeCompare(b.ImmediateParentGroupName);
    if (ipgCmp !== 0) return ipgCmp;
    return a.CourseCode.localeCompare(b.CourseCode);
  });

  return list;
};

export const transformWideToRowWise = (columns, rows, options = {}) => {
  const {
    includeSlotNumber = false,
    trimStrings = true,
    fallbackBlankParentGroup = true
  } = options;

  const courseFields = [
    'ParentGroupName',
    'ImmediateParentGroupName',
    'CourseCode',
    'CourseName'
  ];

  // Build normalized column lookup
  const colMap = new Map();
  columns.forEach(col => {
    colMap.set(col.trim().toLowerCase(), col);
  });

  const findColumn = (field, index) => {
    // Generate list of field variations to try (including ParentGroup / ParentGroupName)
    const fieldAliases = [];
    if (field === 'ParentGroupName') {
      fieldAliases.push('ParentGroupName', 'ParentGroup');
    } else if (field === 'ImmediateParentGroupName') {
      fieldAliases.push('ImmediateParentGroupName', 'ImmediateParentGroup');
    } else {
      fieldAliases.push(field);
    }

    for (const alias of fieldAliases) {
      const exactName = `${alias}${index}`;
      if (columns.includes(exactName)) return exactName;

      const exactNorm = exactName.toLowerCase();
      if (colMap.has(exactNorm)) return colMap.get(exactNorm);

      // Support flexible spacing/underscores: e.g. "Parent Group 1", "Parent Group Name 1", "Parent_Group_1"
      const regex = new RegExp(`^${alias.replace(/([A-Z])/g, '[\\s_]*$1').trim()}[\\s_]*${index}$`, 'i');
      for (const col of columns) {
        if (regex.test(col.trim())) {
          return col;
        }
      }
    }
    return null;
  };

  // Detect all numeric suffixes dynamically
  const indicesSet = new Set();
  columns.forEach(col => {
    const raw = col.trim();
    const match = raw.match(/^course[\s_]*code[\s_]*(\d+)$/i);
    if (match) {
      indicesSet.add(parseInt(match[1], 10));
    }
  });

  // Fallback: check CourseName1..N if CourseCode not explicitly numbered
  if (indicesSet.size === 0) {
    columns.forEach(col => {
      const match = col.trim().match(/^course[\s_]*name[\s_]*(\d+)$/i);
      if (match) {
        indicesSet.add(parseInt(match[1], 10));
      }
    });
  }

  const indices = Array.from(indicesSet).sort((a, b) => a - b);

  // Identify all repeated columns present
  const repeatedColsSet = new Set();
  for (const field of courseFields) {
    for (const i of indices) {
      const colName = findColumn(field, i);
      if (colName) {
        repeatedColsSet.add(colName);
      }
    }
  }

  // Identify base (non-course) columns
  const baseCols = columns.filter(col => !repeatedColsSet.has(col));

  const isValidCourse = (val) => {
    if (val === null || val === undefined) return false;
    const s = String(val).trim();
    if (!s) return false;
    const lower = s.toLowerCase();
    if (lower === 'nan' || lower === 'null' || lower === 'undefined' || lower === '\\n' || lower === '#n/a' || lower === '-') return false;
    return true;
  };

  const cleanVal = (val) => {
    if (val === null || val === undefined) return '';
    const s = String(val);
    if (trimStrings) {
      const trimmed = s.trim();
      const lower = trimmed.toLowerCase();
      if (lower === 'nan' || lower === 'null' || lower === 'undefined' || lower === '\\n') return '';
      return trimmed;
    }
    return s;
  };

  const indexColMapping = indices.map(i => ({
    index: i,
    parentCol: findColumn('ParentGroupName', i),
    immParentCol: findColumn('ImmediateParentGroupName', i),
    codeCol: findColumn('CourseCode', i),
    nameCol: findColumn('CourseName', i)
  }));

  const transformedRows = [];
  let omittedSlotCount = 0;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const baseData = {};
    for (let c = 0; c < baseCols.length; c++) {
      const col = baseCols[c];
      baseData[col] = cleanVal(row[col]);
    }

    for (let k = 0; k < indexColMapping.length; k++) {
      const { index, parentCol, immParentCol, codeCol, nameCol } = indexColMapping[k];
      const courseCode = codeCol ? row[codeCol] : undefined;
      const courseName = nameCol ? row[nameCol] : undefined;

      if (isValidCourse(courseCode) || isValidCourse(courseName)) {
        const record = { ...baseData };
        if (includeSlotNumber) {
          record['CourseSlot'] = index;
        }

        let parentGroupName = parentCol ? cleanVal(row[parentCol]) : '';
        const immParentGroupName = immParentCol ? cleanVal(row[immParentCol]) : '';

        // Copy ImmediateParentGroupName to ParentGroupName if ParentGroupName is blank
        if (fallbackBlankParentGroup && !parentGroupName && immParentGroupName) {
          parentGroupName = immParentGroupName;
        }

        record['ParentGroupName'] = parentGroupName;
        record['ImmediateParentGroupName'] = immParentGroupName;
        record['CourseCode'] = cleanVal(courseCode);
        record['CourseName'] = cleanVal(courseName);
        transformedRows.push(record);
      } else {
        omittedSlotCount++;
      }
    }
  }

  const outputHeaders = [...baseCols];
  if (includeSlotNumber) {
    outputHeaders.push('CourseSlot');
  }
  outputHeaders.push(
    'ParentGroupName',
    'ImmediateParentGroupName',
    'CourseCode',
    'CourseName'
  );

  return {
    transformedRows,
    outputHeaders,
    indices,
    baseCols,
    omittedSlotCount,
    inputRowCount: rows.length,
    outputRowCount: transformedRows.length
  };
};

export default function AdmissionRowWisePage() {
  const [fileName, setFileName] = useState('');
  const [fileSize, setFileSize] = useState(0);
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  
  const [rawColumns, setRawColumns] = useState([]);
  const [rawRows, setRawRows] = useState([]);

  // Options
  const [includeSlotNumber, setIncludeSlotNumber] = useState(false);
  const [trimStrings, setTrimStrings] = useState(true);
  const [fallbackBlankParentGroup, setFallbackBlankParentGroup] = useState(true);

  // Unique Report Mapping & Settings
  const [selectedProgramCol, setSelectedProgramCol] = useState('');
  const [selectedPartCol, setSelectedPartCol] = useState('');
  const [selectedTermCol, setSelectedTermCol] = useState('');
  const [fallbackProgram, setFallbackProgram] = useState('FYIMP Undergraduate');
  const [fallbackPart, setFallbackPart] = useState('1');
  const [fallbackTerm, setFallbackTerm] = useState('1');

  // Active View Tab: 'rowwise' | 'uniqueCourses'
  const [activeViewTab, setActiveViewTab] = useState('rowwise');
  const [uniqueSearchQuery, setUniqueSearchQuery] = useState('');
  const [uniquePage, setUniquePage] = useState(1);
  const [uniquePageSize, setUniquePageSize] = useState(50);

  // Results
  const [transformResult, setTransformResult] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Upload a wide admission file to reshape.');
  const [statusType, setStatusType] = useState('info');

  // Interactive View
  const [searchQuery, setSearchQuery] = useState('');
  const [slotFilter, setSlotFilter] = useState('ALL');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);
  const [showHelpModal, setShowHelpModal] = useState(false);

  const setStatus = (msg, type = 'info') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  const updateColumnMappings = (cols) => {
    const detectedProg = detectColumnMatch(cols, ['Program', 'ProgramName', 'Programme', 'ProgrammeName', 'Degree', 'CourseStream', 'Stream']);
    const detectedPart = detectColumnMatch(cols, ['Part', 'PartName', 'PartNo', 'Year']);
    const detectedTerm = detectColumnMatch(cols, ['Term', 'TermName', 'TermNo', 'Semester', 'Sem']);

    setSelectedProgramCol(detectedProg || '');
    setSelectedPartCol(detectedPart || '');
    setSelectedTermCol(detectedTerm || '');
  };

  // Process rows whenever raw data or options change
  const executeReshape = useCallback((cols, dataRows, opts) => {
    if (!cols || cols.length === 0 || !dataRows || dataRows.length === 0) {
      setTransformResult(null);
      return;
    }

    setIsProcessing(true);
    setStatus('Reshaping course data into row-wise records...', 'info');

    setTimeout(() => {
      try {
        const result = transformWideToRowWise(cols, dataRows, opts);
        setTransformResult(result);
        setPage(1);

        if (result.indices.length === 0) {
          setStatus(`No numeric CourseCode columns found (e.g. CourseCode1, CourseCode2). Reshaped 0 rows.`, 'warning');
        } else {
          setStatus(`Successfully transformed ${result.inputRowCount.toLocaleString()} wide rows into ${result.outputRowCount.toLocaleString()} row-wise course allocations across ${result.indices.length} course slots!`, 'success');
        }
      } catch (err) {
        console.error('Reshape error:', err);
        setStatus(`Transformation failed: ${err.message}`, 'error');
      } finally {
        setIsProcessing(false);
      }
    }, 20);
  }, []);

  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setFileSize(file.size);
    setIsProcessing(true);
    setStatus(`Reading ${file.name}...`, 'info');

    try {
      const buffer = await file.arrayBuffer();
      const wb = parseWorkbookFromBuffer(buffer);

      if (!wb || !wb.SheetNames || wb.SheetNames.length === 0) {
        throw new Error('Could not parse valid workbook or sheets from file.');
      }

      setWorkbook(wb);
      setSheetNames(wb.SheetNames);
      
      const firstSheet = wb.SheetNames[0];
      setSelectedSheet(firstSheet);
      loadSheetData(wb, firstSheet);
    } catch (err) {
      console.error(err);
      setStatus(`Failed to read file: ${err.message}`, 'error');
      setIsProcessing(false);
    }
  };

  const loadSheetData = (wb, sheetName) => {
    try {
      const ws = wb.Sheets[sheetName];
      if (!ws) throw new Error(`Sheet ${sheetName} not found.`);

      // Read raw array of objects
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
      if (!json || json.length === 0) {
        setRawColumns([]);
        setRawRows([]);
        setTransformResult(null);
        setStatus(`Sheet "${sheetName}" contains no data rows.`, 'warning');
        setIsProcessing(false);
        return;
      }

      const cols = Object.keys(json[0] || {});
      setRawColumns(cols);
      setRawRows(json);
      updateColumnMappings(cols);

      executeReshape(cols, json, { includeSlotNumber, trimStrings, fallbackBlankParentGroup });
    } catch (err) {
      console.error(err);
      setStatus(`Error reading sheet "${sheetName}": ${err.message}`, 'error');
      setIsProcessing(false);
    }
  };

  const handleSheetChange = (sheetName) => {
    setSelectedSheet(sheetName);
    if (workbook) {
      loadSheetData(workbook, sheetName);
    }
  };

  const handleLoadSample = () => {
    setFileName('Sample_FYIMP_Wide_Admission_Report.xlsx');
    setFileSize(42800);
    setWorkbook(null);
    setSheetNames(['Admission_Wide_Data']);
    setSelectedSheet('Admission_Wide_Data');

    const cols = Object.keys(SAMPLE_FYIMP_WIDE_DATA[0]);
    setRawColumns(cols);
    setRawRows(SAMPLE_FYIMP_WIDE_DATA);
    updateColumnMappings(cols);

    executeReshape(cols, SAMPLE_FYIMP_WIDE_DATA, { includeSlotNumber, trimStrings, fallbackBlankParentGroup });
  };

  const handleReset = () => {
    setFileName('');
    setFileSize(0);
    setWorkbook(null);
    setSheetNames([]);
    setSelectedSheet('');
    setRawColumns([]);
    setRawRows([]);
    setTransformResult(null);
    setSearchQuery('');
    setSlotFilter('ALL');
    setPage(1);
    setStatus('Upload a wide admission file to reshape.');
  };

  const handleOptionChange = (newOpts) => {
    if (newOpts.includeSlotNumber !== undefined) {
      setIncludeSlotNumber(newOpts.includeSlotNumber);
    }
    if (newOpts.trimStrings !== undefined) {
      setTrimStrings(newOpts.trimStrings);
    }
    if (newOpts.fallbackBlankParentGroup !== undefined) {
      setFallbackBlankParentGroup(newOpts.fallbackBlankParentGroup);
    }

    if (rawColumns.length > 0 && rawRows.length > 0) {
      executeReshape(rawColumns, rawRows, {
        includeSlotNumber: newOpts.includeSlotNumber ?? includeSlotNumber,
        trimStrings: newOpts.trimStrings ?? trimStrings,
        fallbackBlankParentGroup: newOpts.fallbackBlankParentGroup ?? fallbackBlankParentGroup
      });
    }
  };

  // Filtered & Paginated items
  const filteredItems = useMemo(() => {
    if (!transformResult || !transformResult.transformedRows) return [];

    let list = transformResult.transformedRows;

    if (slotFilter !== 'ALL') {
      const sNum = parseInt(slotFilter, 10);
      list = list.filter(r => r.CourseSlot === sNum);
    }

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase().trim();
      list = list.filter(row => {
        return Object.values(row).some(val => 
          val !== null && val !== undefined && String(val).toLowerCase().includes(q)
        );
      });
    }

    return list;
  }, [transformResult, slotFilter, searchQuery]);

  const totalPages = Math.max(1, Math.ceil(filteredItems.length / pageSize));
  const safePage = Math.min(Math.max(1, page), totalPages);

  const pagedItems = useMemo(() => {
    const start = (safePage - 1) * pageSize;
    return filteredItems.slice(start, start + pageSize);
  }, [filteredItems, safePage, pageSize]);

  // Download Reshaped Excel
  const handleExportExcel = () => {
    if (!transformResult || transformResult.transformedRows.length === 0) return;

    try {
      const ws = XLSX.utils.json_to_sheet(transformResult.transformedRows, {
        header: transformResult.outputHeaders
      });

      // Auto-fit column widths
      const colWidths = transformResult.outputHeaders.map(h => {
        const headerLen = String(h).length;
        let maxDataLen = 0;
        for (let i = 0; i < Math.min(transformResult.transformedRows.length, 50); i++) {
          const val = transformResult.transformedRows[i][h];
          if (val) maxDataLen = Math.max(maxDataLen, String(val).length);
        }
        return { wch: Math.max(headerLen + 2, Math.min(maxDataLen + 2, 45)) };
      });
      ws['!cols'] = colWidths;

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'RowWise_Courses');

      const baseName = fileName ? fileName.replace(/\.[^/.]+$/, '') : 'Admission_RowWise_Report';
      XLSX.writeFile(wb, `${baseName}_RowWise.xlsx`);
      setStatus(`Exported ${transformResult.transformedRows.length.toLocaleString()} rows to Excel successfully!`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to export Excel: ${err.message}`, 'error');
    }
  };

  // Download Reshaped CSV
  const handleExportCsv = () => {
    if (!transformResult || transformResult.transformedRows.length === 0) return;

    try {
      const ws = XLSX.utils.json_to_sheet(transformResult.transformedRows, {
        header: transformResult.outputHeaders
      });
      const csv = XLSX.utils.sheet_to_csv(ws);

      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const baseName = fileName ? fileName.replace(/\.[^/.]+$/, '') : 'Admission_RowWise_Report';
      link.href = url;
      link.setAttribute('download', `${baseName}_RowWise.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      setStatus(`Exported ${transformResult.transformedRows.length.toLocaleString()} rows to CSV successfully!`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to export CSV: ${err.message}`, 'error');
    }
  };

  // Computed Unique Courses List (Program, Part, Term, ParentGroupName, ImmediateParentGroupName, CourseCode, CourseName)
  const uniqueCoursesList = useMemo(() => {
    if (!transformResult || !transformResult.transformedRows || transformResult.transformedRows.length === 0) {
      return [];
    }
    return generateUniqueCourseReport(transformResult.transformedRows, {
      programCol: selectedProgramCol,
      partCol: selectedPartCol,
      termCol: selectedTermCol,
      fallbackProgram,
      fallbackPart,
      fallbackTerm
    });
  }, [transformResult, selectedProgramCol, selectedPartCol, selectedTermCol, fallbackProgram, fallbackPart, fallbackTerm]);

  const uniqueProgramsCount = useMemo(() => {
    return new Set(uniqueCoursesList.map(c => c.Program)).size;
  }, [uniqueCoursesList]);

  const filteredUniqueItems = useMemo(() => {
    if (!uniqueSearchQuery.trim()) return uniqueCoursesList;
    const q = uniqueSearchQuery.toLowerCase().trim();
    return uniqueCoursesList.filter(row => {
      return Object.values(row).some(val => 
        val !== null && val !== undefined && String(val).toLowerCase().includes(q)
      );
    });
  }, [uniqueCoursesList, uniqueSearchQuery]);

  const totalUniquePages = Math.max(1, Math.ceil(filteredUniqueItems.length / uniquePageSize));
  const safeUniquePage = Math.min(Math.max(1, uniquePage), totalUniquePages);

  const pagedUniqueItems = useMemo(() => {
    const start = (safeUniquePage - 1) * uniquePageSize;
    return filteredUniqueItems.slice(start, start + uniquePageSize);
  }, [filteredUniqueItems, safeUniquePage, uniquePageSize]);

  // Download Unique Courses Master Excel
  const handleExportUniqueCoursesExcel = () => {
    if (!uniqueCoursesList || uniqueCoursesList.length === 0) {
      setStatus('No unique course records available to export.', 'warning');
      return;
    }

    try {
      const headers = [
        'Program',
        'Part',
        'Term',
        'ParentGroupName',
        'ImmediateParentGroupName',
        'CourseCode',
        'CourseName'
      ];

      const ws = XLSX.utils.json_to_sheet(uniqueCoursesList, { header: headers });

      const colWidths = headers.map(h => {
        const headerLen = String(h).length;
        let maxDataLen = 0;
        for (let i = 0; i < Math.min(uniqueCoursesList.length, 100); i++) {
          const val = uniqueCoursesList[i][h];
          if (val) maxDataLen = Math.max(maxDataLen, String(val).length);
        }
        return { wch: Math.max(headerLen + 2, Math.min(maxDataLen + 2, 45)) };
      });
      ws['!cols'] = colWidths;

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Unique_Courses');

      const baseName = fileName ? fileName.replace(/\.[^/.]+$/, '') : 'Admission_Report';
      XLSX.writeFile(wb, `${baseName}_Unique_Program_Courses.xlsx`);
      setStatus(`Exported ${uniqueCoursesList.length.toLocaleString()} unique course records to Excel successfully!`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to export Unique Courses Excel: ${err.message}`, 'error');
    }
  };

  // Download Unique Courses Master CSV
  const handleExportUniqueCoursesCsv = () => {
    if (!uniqueCoursesList || uniqueCoursesList.length === 0) {
      setStatus('No unique course records available to export.', 'warning');
      return;
    }

    try {
      const headers = [
        'Program',
        'Part',
        'Term',
        'ParentGroupName',
        'ImmediateParentGroupName',
        'CourseCode',
        'CourseName'
      ];

      const ws = XLSX.utils.json_to_sheet(uniqueCoursesList, { header: headers });
      const csv = XLSX.utils.sheet_to_csv(ws);

      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const baseName = fileName ? fileName.replace(/\.[^/.]+$/, '') : 'Admission_Report';
      link.href = url;
      link.setAttribute('download', `${baseName}_Unique_Program_Courses.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      setStatus(`Exported ${uniqueCoursesList.length.toLocaleString()} unique course records to CSV successfully!`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to export Unique Courses CSV: ${err.message}`, 'error');
    }
  };

  return (
    <div style={{ minHeight: '100vh', background: 'var(--bg)', color: 'var(--ink)', padding: '24px 32px', fontFamily: 'var(--font-family)' }}>
      <div style={{ maxWidth: '1440px', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '20px' }}>
        
        {/* Top Breadcrumb & Actions Bar */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <Link 
              to="/" 
              style={{ 
                display: 'inline-flex', 
                alignItems: 'center', 
                gap: '6px', 
                fontSize: '13px', 
                fontWeight: 600, 
                color: 'var(--muted)', 
                textDecoration: 'none',
                padding: '6px 12px',
                borderRadius: '6px',
                background: 'var(--panel)',
                border: '1px solid var(--line)'
              }}
            >
              <ArrowLeft size={14} /> Back to Dashboard
            </Link>
            <div style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '11px', fontWeight: 700, padding: '3px 8px', borderRadius: '12px', background: 'rgba(23,107,135,0.1)', color: 'var(--accent)' }}>
              <Database size={12} /> Row-Wise Reshaper
            </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            {transformResult && transformResult.transformedRows.length > 0 && (
              <>
                <button 
                  type="button" 
                  onClick={handleExportExcel}
                  style={{ 
                    display: 'inline-flex', 
                    alignItems: 'center', 
                    gap: '6px', 
                    fontSize: '12.5px', 
                    padding: '7px 14px', 
                    background: 'var(--panel)', 
                    color: 'var(--ink)', 
                    border: '1px solid var(--line)', 
                    borderRadius: '6px', 
                    fontWeight: 700, 
                    cursor: 'pointer' 
                  }}
                >
                  <Download size={14} /> Row-Wise Excel ({transformResult.transformedRows.length.toLocaleString()})
                </button>

                <button 
                  type="button" 
                  onClick={handleExportUniqueCoursesExcel}
                  style={{ 
                    display: 'inline-flex', 
                    alignItems: 'center', 
                    gap: '6px', 
                    fontSize: '12.5px', 
                    padding: '7px 14px', 
                    background: '#10b981', 
                    color: 'white', 
                    border: 'none', 
                    borderRadius: '6px', 
                    fontWeight: 700, 
                    cursor: 'pointer' 
                  }}
                >
                  <GraduationCap size={14} /> Unique Report Excel ({uniqueCoursesList.length.toLocaleString()})
                </button>
              </>
            )}

            <button 
              type="button" 
              className="secondary" 
              onClick={() => setShowHelpModal(true)}
              style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '6px 12px' }}
            >
              <HelpCircle size={14} /> Extraction Logic & Guide
            </button>

            {rawRows.length > 0 && (
              <button 
                type="button" 
                className="secondary" 
                onClick={handleReset}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '6px 12px' }}
              >
                <RotateCcw size={14} /> Reset Data
              </button>
            )}
          </div>
        </div>

        {/* Hero Section */}
        <div className="card" style={{ padding: '24px', background: 'linear-gradient(135deg, rgba(23,107,135,0.06) 0%, rgba(92,187,212,0.03) 100%)', border: '1px solid var(--line)' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px' }}>
            <div style={{ flex: 1, minWidth: '300px' }}>
              <h1 style={{ fontSize: '24px', fontWeight: 800, margin: '0 0 6px 0', color: 'var(--ink)', letterSpacing: '-0.02em' }}>
                Admission report maker - Row Wise
              </h1>
              <p style={{ margin: 0, fontSize: '13.5px', color: 'var(--muted)', lineHeight: '1.5' }}>
                Unpivot wide-format admission spreadsheets with multiple course attributes (e.g. <code>CourseCode1..N</code>, <code>CourseName1..N</code>, <code>ParentGroupName1..N</code>) into normalized, row-by-row student-course records, plus generate unique <strong>Program - Part - Term - Courses</strong> Master Reports.
              </p>
            </div>

            <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap', alignItems: 'center' }}>
              {transformResult && transformResult.transformedRows.length > 0 ? (
                <>
                  <button 
                    type="button"
                    onClick={handleExportUniqueCoursesExcel}
                    style={{ 
                      display: 'inline-flex', 
                      alignItems: 'center', 
                      gap: '8px', 
                      padding: '11px 18px', 
                      fontSize: '13.5px', 
                      fontWeight: 700,
                      background: '#10b981',
                      color: 'white',
                      border: 'none',
                      borderRadius: '6px',
                      cursor: 'pointer',
                      boxShadow: '0 2px 8px rgba(16, 185, 129, 0.25)'
                    }}
                  >
                    <GraduationCap size={16} /> Download Unique Courses Report (.xlsx)
                  </button>

                  <button 
                    type="button"
                    onClick={handleExportExcel}
                    style={{ 
                      display: 'inline-flex', 
                      alignItems: 'center', 
                      gap: '8px', 
                      padding: '11px 18px', 
                      fontSize: '13.5px', 
                      fontWeight: 700,
                      background: 'var(--panel)',
                      color: 'var(--ink)',
                      border: '1px solid var(--line)',
                      borderRadius: '6px',
                      cursor: 'pointer'
                    }}
                  >
                    <Download size={16} /> Row-Wise Excel (.xlsx)
                  </button>

                  <button 
                    type="button"
                    className="secondary"
                    onClick={activeViewTab === 'uniqueCourses' ? handleExportUniqueCoursesCsv : handleExportCsv}
                    style={{ 
                      display: 'inline-flex', 
                      alignItems: 'center', 
                      gap: '6px', 
                      padding: '11px 16px', 
                      fontSize: '13.5px', 
                      fontWeight: 600 
                    }}
                  >
                    <FileText size={16} /> Export CSV
                  </button>
                </>
              ) : (
                <button 
                  type="button"
                  onClick={handleLoadSample}
                  style={{ 
                    display: 'inline-flex', 
                    alignItems: 'center', 
                    gap: '6px', 
                    padding: '9px 16px', 
                    fontSize: '13px', 
                    fontWeight: 600,
                    background: 'var(--panel)',
                    color: 'var(--accent)',
                    border: '1px solid var(--accent)',
                    borderRadius: '6px',
                    cursor: 'pointer'
                  }}
                >
                  <Sparkles size={15} /> Load Sample FYIMP Data
                </button>
              )}
            </div>
          </div>
        </div>

        {/* Upload & Configuration Deck */}
        <div style={{ display: 'grid', gridTemplateColumns: rawRows.length > 0 ? '1fr 1.2fr' : '1fr', gap: '16px' }}>
          
          {/* File Upload Dropzone */}
          <div className="card" style={{ padding: '20px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
              <h3 style={{ margin: 0, fontSize: '14px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Upload size={16} color="var(--accent)" /> 1. Upload Wide Admission File
              </h3>
              {fileName && (
                <span style={{ fontSize: '11px', color: 'var(--muted)' }}>
                  {(fileSize / 1024).toFixed(1)} KB
                </span>
              )}
            </div>

            <label style={{
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              padding: '28px 20px',
              border: '2px dashed var(--line)',
              borderRadius: '8px',
              cursor: 'pointer',
              background: 'var(--bg)',
              transition: 'all 0.2s ease',
              textAlign: 'center',
              gap: '8px'
            }}>
              <FileSpreadsheet size={32} color="var(--accent)" />
              <div style={{ fontSize: '13px', fontWeight: 600, color: 'var(--ink)' }}>
                {fileName ? fileName : 'Choose .xlsx, .xls, or .csv admission spreadsheet'}
              </div>
              <div style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                Detects all CourseCode1..N, CourseName1..N, and ParentGroupName1..N columns dynamically
              </div>
              <input 
                type="file" 
                accept=".xlsx,.xls,.csv" 
                onChange={handleFileUpload} 
                style={{ display: 'none' }} 
              />
            </label>

            {sheetNames.length > 1 && (
              <div style={{ marginTop: '14px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                <label style={{ fontSize: '12.5px', fontWeight: 600, color: 'var(--muted)' }}>Select Sheet:</label>
                <select 
                  value={selectedSheet} 
                  onChange={(e) => handleSheetChange(e.target.value)}
                  style={{ flex: 1, padding: '6px 10px', fontSize: '12.5px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--panel)', color: 'var(--ink)' }}
                >
                  {sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
            )}

            {rawRows.length === 0 && (
              <div style={{ marginTop: '14px', padding: '12px 14px', borderRadius: '6px', background: 'var(--panel)', border: '1px solid var(--line)', fontSize: '12px', color: 'var(--muted)', display: 'flex', alignItems: 'center', gap: '10px' }}>
                <FileSpreadsheet size={18} color="var(--accent)" style={{ flexShrink: 0 }} />
                <div>
                  <strong>How to download:</strong> Upload your wide Excel/CSV file above, or click <strong>"Load Sample FYIMP Data"</strong> in the top header. The <strong>Download Reshaped Excel (.xlsx)</strong> and <strong>Unique Courses Report (.xlsx)</strong> buttons will activate immediately.
                </div>
              </div>
            )}
          </div>

          {/* Configuration & Detected Columns Panel */}
          {rawRows.length > 0 && transformResult && (
            <div className="card" style={{ padding: '20px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', gap: '14px' }}>
              <div>
                <h3 style={{ margin: '0 0 12px 0', fontSize: '14px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <Sliders size={16} color="var(--accent)" /> 2. Reshape Settings & Schema Detection
                </h3>

                {/* Detected Slots Tag Cloud */}
                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', marginBottom: '6px' }}>
                    Detected Course Slots ({transformResult.indices.length}):
                  </div>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                    {transformResult.indices.length > 0 ? (
                      transformResult.indices.map(idx => (
                        <span key={idx} style={{ fontSize: '11px', fontWeight: 700, padding: '3px 8px', borderRadius: '4px', background: 'rgba(23,107,135,0.1)', color: 'var(--accent)', border: '1px solid rgba(23,107,135,0.2)' }}>
                          Slot {idx}
                        </span>
                      ))
                    ) : (
                      <span style={{ fontSize: '12px', color: '#ef4444' }}>None detected</span>
                    )}
                  </div>
                </div>

                {/* Base Columns Count */}
                <div style={{ marginBottom: '14px' }}>
                  <div style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', marginBottom: '6px' }}>
                    Preserved Base Columns ({transformResult.baseCols.length}):
                  </div>
                  <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px', maxHeight: '60px', overflowY: 'auto' }}>
                    {transformResult.baseCols.map(col => (
                      <span key={col} style={{ fontSize: '10.5px', padding: '2px 6px', borderRadius: '3px', background: 'var(--bg)', color: 'var(--muted)', border: '1px solid var(--line)' }}>
                        {col}
                      </span>
                    ))}
                  </div>
                </div>

                {/* Toggles */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', cursor: 'pointer' }}>
                    <input 
                      type="checkbox" 
                      checked={includeSlotNumber} 
                      onChange={(e) => handleOptionChange({ includeSlotNumber: e.target.checked })} 
                    />
                    <span>Include <strong>CourseSlot</strong> (Slot #) in transformed output</span>
                  </label>

                  <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', cursor: 'pointer' }}>
                    <input 
                      type="checkbox" 
                      checked={trimStrings} 
                      onChange={(e) => handleOptionChange({ trimStrings: e.target.checked })} 
                    />
                    <span>Sanitize escape artifacts (convert <code>\N</code>, <code>null</code>, <code>NaN</code> to clean blanks)</span>
                  </label>

                  <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', cursor: 'pointer' }}>
                    <input 
                      type="checkbox" 
                      checked={fallbackBlankParentGroup} 
                      onChange={(e) => handleOptionChange({ fallbackBlankParentGroup: e.target.checked })} 
                    />
                    <span>Copy <strong>ImmediateParentGroupName</strong> to <strong>ParentGroupName</strong> when ParentGroupName is blank</span>
                  </label>
                </div>

                {/* Unique Course Report Column Mapping */}
                <div style={{ marginTop: '14px', paddingTop: '12px', borderTop: '1px dashed var(--line)' }}>
                  <div style={{ fontSize: '12.5px', fontWeight: 700, color: 'var(--ink)', marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <GraduationCap size={15} color="#059669" />
                    <span>Unique Course Master Mapping:</span>
                    <span style={{ fontSize: '11px', fontWeight: 600, color: '#059669', background: 'rgba(16,185,129,0.1)', padding: '2px 7px', borderRadius: '10px' }}>
                      {uniqueCoursesList.length} Unique Courses
                    </span>
                  </div>

                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(130px, 1fr))', gap: '8px', marginBottom: '8px' }}>
                    <div>
                      <label style={{ fontSize: '11px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '3px' }}>
                        Program Column:
                      </label>
                      <select 
                        value={selectedProgramCol} 
                        onChange={(e) => setSelectedProgramCol(e.target.value)}
                        style={{ width: '100%', fontSize: '12px', padding: '5px 8px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', color: 'var(--ink)' }}
                      >
                        <option value="">-- Fixed/Default Value --</option>
                        {transformResult.baseCols.map(col => <option key={col} value={col}>{col}</option>)}
                      </select>
                    </div>

                    <div>
                      <label style={{ fontSize: '11px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '3px' }}>
                        Part Column:
                      </label>
                      <select 
                        value={selectedPartCol} 
                        onChange={(e) => setSelectedPartCol(e.target.value)}
                        style={{ width: '100%', fontSize: '12px', padding: '5px 8px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', color: 'var(--ink)' }}
                      >
                        <option value="">-- Fixed: 1 --</option>
                        {transformResult.baseCols.map(col => <option key={col} value={col}>{col}</option>)}
                      </select>
                    </div>

                    <div>
                      <label style={{ fontSize: '11px', fontWeight: 600, color: 'var(--muted)', display: 'block', marginBottom: '3px' }}>
                        Term Column:
                      </label>
                      <select 
                        value={selectedTermCol} 
                        onChange={(e) => setSelectedTermCol(e.target.value)}
                        style={{ width: '100%', fontSize: '12px', padding: '5px 8px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', color: 'var(--ink)' }}
                      >
                        <option value="">-- Fixed: 1 --</option>
                        {transformResult.baseCols.map(col => <option key={col} value={col}>{col}</option>)}
                      </select>
                    </div>
                  </div>

                  {!selectedProgramCol && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px', marginTop: '6px' }}>
                      <span style={{ fontSize: '11px', color: 'var(--muted)', whiteSpace: 'nowrap' }}>Default Program:</span>
                      <input 
                        type="text" 
                        value={fallbackProgram} 
                        onChange={(e) => setFallbackProgram(e.target.value)}
                        placeholder="e.g. FYIMP Undergraduate"
                        style={{ flex: 1, fontSize: '11.5px', padding: '4px 8px', borderRadius: '4px', border: '1px solid var(--line)' }}
                      />
                    </div>
                  )}
                </div>

                {/* Direct Download Actions Box */}
                <div style={{ marginTop: '16px', background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px', padding: '14px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--ink)' }}>📥 Download Available Reports:</span>
                    <span style={{ fontSize: '11.5px', fontWeight: 700, color: '#10b981' }}>{transformResult.outputRowCount.toLocaleString()} Rows | {uniqueCoursesList.length.toLocaleString()} Unique</span>
                  </div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: '8px' }}>
                    <button 
                      type="button"
                      onClick={handleExportExcel}
                      style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '6px', padding: '9px 12px', fontSize: '12.5px', fontWeight: 700, background: 'var(--panel)', color: 'var(--ink)', border: '1px solid var(--line)', borderRadius: '6px', cursor: 'pointer' }}
                    >
                      <Download size={14} /> Full Row-Wise (.xlsx)
                    </button>
                    <button 
                      type="button"
                      onClick={handleExportUniqueCoursesExcel}
                      style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '6px', padding: '9px 12px', fontSize: '12.5px', fontWeight: 700, background: '#10b981', color: 'white', border: 'none', borderRadius: '6px', cursor: 'pointer' }}
                    >
                      <GraduationCap size={14} /> Unique Report (.xlsx)
                    </button>
                  </div>
                </div>
              </div>

              {/* Status Banner */}
              <div style={{ 
                padding: '8px 12px', 
                borderRadius: '6px', 
                fontSize: '12px', 
                fontWeight: 600,
                background: statusType === 'success' ? 'rgba(16, 185, 129, 0.1)' : statusType === 'error' ? 'rgba(239, 68, 68, 0.1)' : 'rgba(23, 107, 135, 0.08)',
                color: statusType === 'success' ? '#059669' : statusType === 'error' ? '#dc2626' : 'var(--accent)',
                border: `1px solid ${statusType === 'success' ? 'rgba(16, 185, 129, 0.2)' : statusType === 'error' ? 'rgba(239, 68, 68, 0.2)' : 'rgba(23, 107, 135, 0.2)'}`,
                display: 'flex',
                alignItems: 'center',
                gap: '8px'
              }}>
                {isProcessing && (
                  <div style={{
                    width: '12px',
                    height: '12px',
                    border: '2px solid currentColor',
                    borderTopColor: 'transparent',
                    borderRadius: '50%',
                    animation: 'spin 0.7s linear infinite'
                  }} />
                )}
                {statusType === 'success' && !isProcessing && <CheckCircle2 size={15} />}
                <span>{statusMsg}</span>
              </div>

            </div>
          )}
        </div>

        {/* KPI Summary Metrics Cards */}
        {transformResult && (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: '14px' }}>
            <div className="card" style={{ padding: '16px 20px', borderLeft: '4px solid var(--accent)' }}>
              <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', marginBottom: '4px' }}>Input Students</div>
              <div style={{ fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>{transformResult.inputRowCount.toLocaleString()}</div>
              <div style={{ fontSize: '11.5px', color: 'var(--muted)' }}>Wide applicant rows</div>
            </div>

            <div className="card" style={{ padding: '16px 20px', borderLeft: '4px solid #10b981' }}>
              <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', marginBottom: '4px' }}>Reshaped Records</div>
              <div style={{ fontSize: '24px', fontWeight: 800, color: '#059669' }}>{transformResult.outputRowCount.toLocaleString()}</div>
              <div style={{ fontSize: '11.5px', color: 'var(--muted)' }}>Normalized student-course rows</div>
            </div>

            <div className="card" style={{ padding: '16px 20px', borderLeft: '4px solid #059669', background: 'rgba(16,185,129,0.02)' }}>
              <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', marginBottom: '4px' }}>Unique Courses</div>
              <div style={{ fontSize: '24px', fontWeight: 800, color: '#059669' }}>{uniqueCoursesList.length.toLocaleString()}</div>
              <div style={{ fontSize: '11.5px', color: 'var(--muted)' }}>Program-Term Master Combinations</div>
            </div>

            <div className="card" style={{ padding: '16px 20px', borderLeft: '4px solid #8b5cf6' }}>
              <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', marginBottom: '4px' }}>Unique Programs</div>
              <div style={{ fontSize: '24px', fontWeight: 800, color: '#7c3aed' }}>{uniqueProgramsCount}</div>
              <div style={{ fontSize: '11.5px', color: 'var(--muted)' }}>Distinct academic streams</div>
            </div>

            <div className="card" style={{ padding: '16px 20px', borderLeft: '4px solid #f59e0b' }}>
              <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', marginBottom: '4px' }}>Empty Slots Omitted</div>
              <div style={{ fontSize: '24px', fontWeight: 800, color: '#d97706' }}>{transformResult.omittedSlotCount.toLocaleString()}</div>
              <div style={{ fontSize: '11.5px', color: 'var(--muted)' }}>Blanks skipped automatically</div>
            </div>
          </div>
        )}

        {/* Results Deck with Interactive Search, Export & Paginated Grid */}
        {transformResult && transformResult.transformedRows.length > 0 && (
          <div className="card" style={{ padding: '0', overflow: 'hidden' }}>
            
            {/* View Switcher Tabs */}
            <div style={{ display: 'flex', borderBottom: '1px solid var(--line)', background: 'var(--bg)' }}>
              <button
                type="button"
                onClick={() => setActiveViewTab('rowwise')}
                style={{
                  padding: '12px 20px',
                  fontSize: '13px',
                  fontWeight: 700,
                  background: activeViewTab === 'rowwise' ? 'var(--panel)' : 'transparent',
                  border: 'none',
                  borderBottom: activeViewTab === 'rowwise' ? '2px solid var(--accent)' : '2px solid transparent',
                  color: activeViewTab === 'rowwise' ? 'var(--accent)' : 'var(--muted)',
                  cursor: 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '8px'
                }}
              >
                <FileSpreadsheet size={15} />
                <span>Row-Wise Course Allocations</span>
                <span style={{ fontSize: '11px', padding: '2px 7px', borderRadius: '10px', background: 'rgba(23,107,135,0.1)', color: 'var(--accent)' }}>
                  {transformResult.outputRowCount.toLocaleString()}
                </span>
              </button>

              <button
                type="button"
                onClick={() => setActiveViewTab('uniqueCourses')}
                style={{
                  padding: '12px 20px',
                  fontSize: '13px',
                  fontWeight: 700,
                  background: activeViewTab === 'uniqueCourses' ? 'var(--panel)' : 'transparent',
                  border: 'none',
                  borderBottom: activeViewTab === 'uniqueCourses' ? '2px solid #10b981' : '2px solid transparent',
                  color: activeViewTab === 'uniqueCourses' ? '#059669' : 'var(--muted)',
                  cursor: 'pointer',
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: '8px'
                }}
              >
                <GraduationCap size={16} />
                <span>Unique Course Master Report</span>
                <span style={{ fontSize: '11px', padding: '2px 7px', borderRadius: '10px', background: 'rgba(16,185,129,0.1)', color: '#059669' }}>
                  {uniqueCoursesList.length.toLocaleString()} Unique Courses
                </span>
              </button>
            </div>

            {/* TAB 1: ROW-WISE ALLOCATIONS VIEW */}
            {activeViewTab === 'rowwise' && (
              <>
                {/* Toolbar */}
                <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                  
                  {/* Search & Slot Filter */}
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flex: 1, minWidth: '320px', flexWrap: 'wrap' }}>
                    <div style={{ position: 'relative', flex: 1, minWidth: '220px' }}>
                      <Search size={15} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
                      <input 
                        type="text" 
                        placeholder="Search across student name, application no, course code..."
                        value={searchQuery}
                        onChange={(e) => { setSearchQuery(e.target.value); setPage(1); }}
                        style={{ width: '100%', paddingLeft: '32px', fontSize: '13px' }}
                      />
                    </div>

                    {includeSlotNumber && (
                      <select 
                        value={slotFilter} 
                        onChange={(e) => { setSlotFilter(e.target.value); setPage(1); }}
                        style={{ padding: '8px 12px', fontSize: '13px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--panel)' }}
                      >
                        <option value="ALL">All Slots</option>
                        {transformResult.indices.map(i => <option key={i} value={i}>Slot {i}</option>)}
                      </select>
                    )}
                  </div>

                  {/* Export Buttons */}
                  <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                    <button 
                      type="button"
                      onClick={handleExportExcel}
                      style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '8px 16px', background: 'var(--accent)', color: 'white' }}
                    >
                      <Download size={15} /> Download Reshaped Excel (.xlsx)
                    </button>

                    <button 
                      type="button"
                      className="secondary"
                      onClick={handleExportCsv}
                      style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '8px 16px' }}
                    >
                      <FileText size={15} /> Export CSV (.csv)
                    </button>
                  </div>

                </div>

                {/* Table Container */}
                <div style={{ overflowX: 'auto', maxHeight: '560px' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                    <thead>
                      <tr style={{ background: 'var(--bg)', borderBottom: '2px solid var(--line)', position: 'sticky', top: 0, zIndex: 2 }}>
                        <th style={{ padding: '10px 14px', width: '50px', textAlign: 'center' }}>#</th>
                        {transformResult.outputHeaders.map(h => {
                          const isCourseCol = ['CourseCode', 'CourseName', 'ParentGroupName', 'ImmediateParentGroupName', 'CourseSlot'].includes(h);
                          return (
                            <th 
                              key={h} 
                              style={{ 
                                padding: '10px 14px', 
                                textAlign: 'left', 
                                whiteSpace: 'nowrap',
                                color: isCourseCol ? 'var(--accent)' : 'var(--ink)',
                                background: isCourseCol ? 'rgba(23,107,135,0.04)' : 'transparent'
                              }}
                            >
                              {h}
                            </th>
                          );
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      {pagedItems.length > 0 ? (
                        pagedItems.map((row, idx) => {
                          const globalIdx = (safePage - 1) * pageSize + idx + 1;
                          return (
                            <tr 
                              key={idx} 
                              style={{ 
                                borderBottom: '1px solid var(--line)', 
                                background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.015)' 
                              }}
                            >
                              <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 600, color: 'var(--muted)' }}>
                                {globalIdx}
                              </td>
                              {transformResult.outputHeaders.map(h => {
                                const val = row[h];
                                if (h === 'CourseCode') {
                                  return (
                                    <td key={h} style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                                      <span style={{ 
                                        fontFamily: 'monospace', 
                                        fontWeight: 700, 
                                        padding: '2px 7px', 
                                        borderRadius: '4px', 
                                        background: 'rgba(23,107,135,0.1)', 
                                        color: 'var(--accent)',
                                        fontSize: '11.5px'
                                      }}>
                                        {val || '—'}
                                      </span>
                                    </td>
                                  );
                                }
                                if (h === 'CourseSlot') {
                                  return (
                                    <td key={h} style={{ padding: '10px 14px', textAlign: 'center' }}>
                                      <span style={{ fontSize: '11px', fontWeight: 700, padding: '2px 6px', borderRadius: '10px', background: 'var(--panel)', border: '1px solid var(--line)' }}>
                                        #{val}
                                      </span>
                                    </td>
                                  );
                                }
                                return (
                                  <td key={h} style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                                    {val !== undefined && val !== null && String(val).trim() !== '' ? String(val) : <span style={{ color: 'var(--muted)' }}>—</span>}
                                  </td>
                                );
                              })}
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={transformResult.outputHeaders.length + 1} style={{ textAlign: 'center', padding: '40px', color: 'var(--muted)' }}>
                            No records match query "{searchQuery}".
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Pagination Footer */}
                <div style={{ 
                  padding: '12px 20px', 
                  background: 'var(--bg)', 
                  borderTop: '1px solid var(--line)', 
                  display: 'flex', 
                  justifyContent: 'space-between', 
                  alignItems: 'center', 
                  flexWrap: 'wrap',
                  gap: '12px',
                  fontSize: '13px', 
                  color: 'var(--muted)' 
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
                    <div>
                      Showing <strong>{filteredItems.length === 0 ? 0 : (safePage - 1) * pageSize + 1}</strong> to <strong>{Math.min(safePage * pageSize, filteredItems.length)}</strong> of <strong>{filteredItems.length.toLocaleString()}</strong> record(s)
                    </div>

                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <span>Rows per page:</span>
                      <select 
                        value={pageSize} 
                        onChange={(e) => { 
                          setPageSize(Number(e.target.value)); 
                          setPage(1); 
                        }}
                        style={{ padding: '4px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)' }}
                      >
                        <option value={25}>25</option>
                        <option value={50}>50</option>
                        <option value={100}>100</option>
                        <option value={250}>250</option>
                        <option value={500}>500</option>
                      </select>
                    </div>
                  </div>

                  {totalPages > 1 && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <button 
                        type="button"
                        className="secondary" 
                        onClick={() => setPage(1)} 
                        disabled={safePage <= 1}
                        style={{ padding: '4px 10px', fontSize: '12px' }}
                      >
                        « First
                      </button>

                      <button 
                        type="button"
                        className="secondary" 
                        onClick={() => setPage(p => Math.max(1, p - 1))} 
                        disabled={safePage <= 1}
                        style={{ padding: '4px 10px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '2px' }}
                      >
                        <ChevronLeft size={14} /> Prev
                      </button>

                      <span style={{ margin: '0 8px', fontSize: '12.5px', fontWeight: 600, color: 'var(--ink)' }}>
                        Page {safePage} of {totalPages}
                      </span>

                      <button 
                        type="button"
                        className="secondary" 
                        onClick={() => setPage(p => Math.min(totalPages, p + 1))} 
                        disabled={safePage >= totalPages}
                        style={{ padding: '4px 10px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '2px' }}
                      >
                        Next <ChevronRight size={14} />
                      </button>

                      <button 
                        type="button"
                        className="secondary" 
                        onClick={() => setPage(totalPages)} 
                        disabled={safePage >= totalPages}
                        style={{ padding: '4px 10px', fontSize: '12px' }}
                      >
                        Last »
                      </button>
                    </div>
                  )}
                </div>
              </>
            )}

            {/* TAB 2: UNIQUE COURSE MASTER REPORT VIEW */}
            {activeViewTab === 'uniqueCourses' && (
              <>
                {/* Unique Toolbar */}
                <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
                  
                  {/* Search Unique Courses */}
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flex: 1, minWidth: '320px', flexWrap: 'wrap' }}>
                    <div style={{ position: 'relative', flex: 1, minWidth: '220px' }}>
                      <Search size={15} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
                      <input 
                        type="text" 
                        placeholder="Search across CourseCode, CourseName, Program, Groups..."
                        value={uniqueSearchQuery}
                        onChange={(e) => { setUniqueSearchQuery(e.target.value); setUniquePage(1); }}
                        style={{ width: '100%', paddingLeft: '32px', fontSize: '13px' }}
                      />
                    </div>
                    <span style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)' }}>
                      {filteredUniqueItems.length.toLocaleString()} matching courses
                    </span>
                  </div>

                  {/* Export Buttons */}
                  <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                    <button 
                      type="button"
                      onClick={handleExportUniqueCoursesExcel}
                      style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '8px 16px', background: '#10b981', color: 'white', border: 'none', borderRadius: '6px', fontWeight: 700, cursor: 'pointer' }}
                    >
                      <Download size={15} /> Download Unique Courses (.xlsx)
                    </button>

                    <button 
                      type="button"
                      className="secondary"
                      onClick={handleExportUniqueCoursesCsv}
                      style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '8px 16px' }}
                    >
                      <FileText size={15} /> Export CSV (.csv)
                    </button>
                  </div>

                </div>

                {/* Sub-header Banner */}
                <div style={{ padding: '8px 20px', background: 'rgba(16,185,129,0.04)', borderBottom: '1px solid var(--line)', fontSize: '12px', color: '#059669', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <GraduationCap size={15} />
                  <span>Unique combinations of <strong>Program</strong>, <strong>Part</strong>, <strong>Term</strong>, <strong>ParentGroupName</strong>, <strong>ImmediateParentGroupName</strong>, <strong>CourseCode</strong>, and <strong>CourseName</strong>.</span>
                </div>

                {/* Unique Table Container */}
                <div style={{ overflowX: 'auto', maxHeight: '560px' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                    <thead>
                      <tr style={{ background: 'var(--bg)', borderBottom: '2px solid var(--line)', position: 'sticky', top: 0, zIndex: 2 }}>
                        <th style={{ padding: '10px 14px', width: '50px', textAlign: 'center' }}>#</th>
                        <th style={{ padding: '10px 14px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--accent)' }}>Program</th>
                        <th style={{ padding: '10px 14px', textAlign: 'center', width: '70px', whiteSpace: 'nowrap' }}>Part</th>
                        <th style={{ padding: '10px 14px', textAlign: 'center', width: '70px', whiteSpace: 'nowrap' }}>Term</th>
                        <th style={{ padding: '10px 14px', textAlign: 'left', whiteSpace: 'nowrap' }}>ParentGroupName</th>
                        <th style={{ padding: '10px 14px', textAlign: 'left', whiteSpace: 'nowrap' }}>ImmediateParentGroupName</th>
                        <th style={{ padding: '10px 14px', textAlign: 'left', whiteSpace: 'nowrap', color: 'var(--accent)' }}>CourseCode</th>
                        <th style={{ padding: '10px 14px', textAlign: 'left', whiteSpace: 'nowrap' }}>CourseName</th>
                      </tr>
                    </thead>
                    <tbody>
                      {pagedUniqueItems.length > 0 ? (
                        pagedUniqueItems.map((item, idx) => {
                          const globalIdx = (safeUniquePage - 1) * uniquePageSize + idx + 1;
                          return (
                            <tr 
                              key={idx} 
                              style={{ 
                                borderBottom: '1px solid var(--line)', 
                                background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.015)' 
                              }}
                            >
                              <td style={{ padding: '10px 14px', textAlign: 'center', fontWeight: 600, color: 'var(--muted)' }}>
                                {globalIdx}
                              </td>
                              <td style={{ padding: '10px 14px', fontWeight: 600, color: 'var(--ink)' }}>
                                {item.Program || '—'}
                              </td>
                              <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                                <span style={{ fontSize: '11.5px', padding: '2px 7px', borderRadius: '4px', background: 'var(--bg)', border: '1px solid var(--line)' }}>
                                  {item.Part || '—'}
                                </span>
                              </td>
                              <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                                <span style={{ fontSize: '11.5px', padding: '2px 7px', borderRadius: '4px', background: 'var(--bg)', border: '1px solid var(--line)' }}>
                                  {item.Term || '—'}
                                </span>
                              </td>
                              <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                                {item.ParentGroupName || <span style={{ color: 'var(--muted)' }}>—</span>}
                              </td>
                              <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                                {item.ImmediateParentGroupName || <span style={{ color: 'var(--muted)' }}>—</span>}
                              </td>
                              <td style={{ padding: '10px 14px', whiteSpace: 'nowrap' }}>
                                <span style={{ 
                                  fontFamily: 'monospace', 
                                  fontWeight: 700, 
                                  padding: '2px 7px', 
                                  borderRadius: '4px', 
                                  background: 'rgba(16,185,129,0.1)', 
                                  color: '#059669',
                                  fontSize: '11.5px'
                                }}>
                                  {item.CourseCode || '—'}
                                </span>
                              </td>
                              <td style={{ padding: '10px 14px', fontWeight: 500 }}>
                                {item.CourseName || <span style={{ color: 'var(--muted)' }}>—</span>}
                              </td>
                            </tr>
                          );
                        })
                      ) : (
                        <tr>
                          <td colSpan={8} style={{ textAlign: 'center', padding: '40px', color: 'var(--muted)' }}>
                            No unique courses match query "{uniqueSearchQuery}".
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>

                {/* Unique Pagination Footer */}
                <div style={{ 
                  padding: '12px 20px', 
                  background: 'var(--bg)', 
                  borderTop: '1px solid var(--line)', 
                  display: 'flex', 
                  justifyContent: 'space-between', 
                  alignItems: 'center', 
                  flexWrap: 'wrap',
                  gap: '12px',
                  fontSize: '13px', 
                  color: 'var(--muted)' 
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
                    <div>
                      Showing <strong>{filteredUniqueItems.length === 0 ? 0 : (safeUniquePage - 1) * uniquePageSize + 1}</strong> to <strong>{Math.min(safeUniquePage * uniquePageSize, filteredUniqueItems.length)}</strong> of <strong>{filteredUniqueItems.length.toLocaleString()}</strong> unique course(s)
                    </div>

                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <span>Rows per page:</span>
                      <select 
                        value={uniquePageSize} 
                        onChange={(e) => { 
                          setUniquePageSize(Number(e.target.value)); 
                          setUniquePage(1); 
                        }}
                        style={{ padding: '4px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)' }}
                      >
                        <option value={25}>25</option>
                        <option value={50}>50</option>
                        <option value={100}>100</option>
                        <option value={250}>250</option>
                      </select>
                    </div>
                  </div>

                  {totalUniquePages > 1 && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                      <button 
                        type="button"
                        className="secondary" 
                        onClick={() => setUniquePage(1)} 
                        disabled={safeUniquePage <= 1}
                        style={{ padding: '4px 10px', fontSize: '12px' }}
                      >
                        « First
                      </button>

                      <button 
                        type="button"
                        className="secondary" 
                        onClick={() => setUniquePage(p => Math.max(1, p - 1))} 
                        disabled={safeUniquePage <= 1}
                        style={{ padding: '4px 10px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '2px' }}
                      >
                        <ChevronLeft size={14} /> Prev
                      </button>

                      <span style={{ margin: '0 8px', fontSize: '12.5px', fontWeight: 600, color: 'var(--ink)' }}>
                        Page {safeUniquePage} of {totalUniquePages}
                      </span>

                      <button 
                        type="button"
                        className="secondary" 
                        onClick={() => setUniquePage(p => Math.min(totalUniquePages, p + 1))} 
                        disabled={safeUniquePage >= totalUniquePages}
                        style={{ padding: '4px 10px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '2px' }}
                      >
                        Next <ChevronRight size={14} />
                      </button>

                      <button 
                        type="button"
                        className="secondary" 
                        onClick={() => setUniquePage(totalUniquePages)} 
                        disabled={safeUniquePage >= totalUniquePages}
                        style={{ padding: '4px 10px', fontSize: '12px' }}
                      >
                        Last »
                      </button>
                    </div>
                  )}
                </div>
              </>
            )}

          </div>
        )}

      </div>

      {/* Extraction Logic & Guide Modal */}
      {showHelpModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          backgroundColor: 'rgba(0, 0, 0, 0.6)',
          backdropFilter: 'blur(3px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          padding: '20px'
        }}>
          <div style={{
            backgroundColor: 'var(--panel)',
            border: '1px solid var(--line)',
            borderRadius: '12px',
            maxWidth: '720px',
            width: '100%',
            maxHeight: '85vh',
            overflowY: 'auto',
            padding: '28px',
            position: 'relative',
            boxShadow: '0 20px 40px rgba(0,0,0,0.25)'
          }}>
            <button 
              type="button"
              onClick={() => setShowHelpModal(false)}
              style={{
                position: 'absolute',
                top: '18px',
                right: '18px',
                background: 'transparent',
                border: 'none',
                cursor: 'pointer',
                color: 'var(--muted)'
              }}
            >
              <X size={20} />
            </button>

            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '16px' }}>
              <div style={{ width: '36px', height: '36px', borderRadius: '8px', background: 'rgba(23,107,135,0.1)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                <Database size={20} color="var(--accent)" />
              </div>
              <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 800 }}>Admission Report Maker - Row Wise Guide</h2>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', fontSize: '13.5px', lineHeight: '1.6', color: 'var(--ink)' }}>
              <div>
                <h4 style={{ margin: '0 0 6px 0', color: 'var(--accent)', fontSize: '14px' }}>1. What this tool does</h4>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  In university admission registers (such as FYUGP and FYIMP 4-year undergraduate programmes), student course selections are exported as multiple wide columns on a single row (e.g. <code>CourseCode1</code>, <code>CourseCode2</code>, up to <code>CourseCode10</code>). This tool unpivots the wide structure into standardized row-by-row records where each student-course pair occupies its own row.
                </p>
              </div>

              <div>
                <h4 style={{ margin: '0 0 6px 0', color: 'var(--accent)', fontSize: '14px' }}>2. Supported Course Attribute Fields</h4>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  For each numeric index <code>i</code> (1, 2, ..., N), the tool detects and unpivots:
                </p>
                <ul style={{ margin: '6px 0 0 20px', padding: 0, color: 'var(--muted)' }}>
                  <li><code>ParentGroupName&#123;i&#125;</code> (e.g. Major Core, Minor Group)</li>
                  <li><code>ImmediateParentGroupName&#123;i&#125;</code> (e.g. Discipline Specific Core)</li>
                  <li><code>CourseCode&#123;i&#125;</code> (e.g. KU4DSCHIS101)</li>
                  <li><code>CourseName&#123;i&#125;</code> (e.g. Early Indian History)</li>
                </ul>
              </div>

              <div>
                <h4 style={{ margin: '0 0 6px 0', color: 'var(--accent)', fontSize: '14px' }}>3. Filtering Empty Slots</h4>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  If a student did not register for slot <code>i</code> (i.e. both <code>CourseCode&#123;i&#125;</code> and <code>CourseName&#123;i&#125;</code> are empty, <code>\N</code>, or <code>NaN</code>), that slot is automatically omitted without generating blank rows.
                </p>
              </div>

              <div>
                <h4 style={{ margin: '0 0 6px 0', color: 'var(--accent)', fontSize: '14px' }}>4. Full Base Data Preservation</h4>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  All candidate base columns (e.g. <code>ApplicationNo</code>, <code>CandidateName</code>, <code>Mobile</code>, <code>CollegeCode</code>, <code>AllottedCategory</code>, etc.) are replicated faithfully onto every unpivoted course row.
                </p>
              </div>

              <div>
                <h4 style={{ margin: '0 0 6px 0', color: '#059669', fontSize: '14px' }}>5. Unique Course Master Report (Excel)</h4>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  Generates an exportable curriculum report containing deduplicated records across 7 core columns: <code>Program</code>, <code>Part</code>, <code>Term</code>, <code>ParentGroupName</code>, <code>ImmediateParentGroupName</code>, <code>CourseCode</code>, and <code>CourseName</code>. Ideal for populating university syllabus master databases and verifying course codes.
                </p>
              </div>
            </div>

            <div style={{ marginTop: '24px', textAlign: 'right' }}>
              <button 
                type="button" 
                onClick={() => setShowHelpModal(false)}
                style={{ padding: '8px 20px', background: 'var(--accent)', color: 'white', borderRadius: '6px', border: 'none', fontWeight: 600, fontSize: '13px', cursor: 'pointer' }}
              >
                Got it
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
