import React, { useState, useMemo } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
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
  UserX
} from 'lucide-react';

export const ADES_OUTPUT_HEADERS = [
  'Faculty',
  'Program Term Name',
  'Course Code',
  'Course Name',
  'Seat Number',
  'PRN',
  'ESE - PR Max',
  'ESE - PR Min',
  'ESE - PR Obtained',
  'ESE - TH Max',
  'ESE - TH Min',
  'ESE - TH Obtained',
  'ESE - Max',
  'ESE - Min',
  'ESE Overall',
  'CE - PR Max',
  'CE - PR Min',
  'CE - PR Obtained',
  'CE - TH Max',
  'CE - TH Min',
  'CE - TH Obtained',
  'CE - Max',
  'CE - Min',
  'CE Overall Marks ',
  'Overall Maximum',
  'Overall Minimum',
  'Course Overall Marks ',
  'ESE Pass (ESE obtained  >= ESE Min is pass if less then fail)',
  'Overall pass (Course overall Marks >= Overall Minimum is pass if less then fail)',
  'Course Pass/Fail ("Pass"only if both ESE Pass and Overall pass are pass if any condition is fail then course pass/fail status is fail )'
];

const normalizeKey = (key) => String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');

export default function AdesResultCalculatorPage() {
  const [sourceFile, setSourceFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [workbook, setWorkbook] = useState(null);
  const [rawRows, setRawRows] = useState([]);
  const [processedRows, setProcessedRows] = useState([]);
  const [headerMap, setHeaderMap] = useState({});
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('info');
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedFacultyFilter, setSelectedFacultyFilter] = useState('ALL');
  const [selectedProgramFilter, setSelectedProgramFilter] = useState('ALL');
  const [selectedCourseFilter, setSelectedCourseFilter] = useState('ALL');
  const [selectedResultFilter, setSelectedResultFilter] = useState('ALL'); // 'ALL' | 'PASS' | 'FAIL' | 'ESE_FAIL' | 'OVERALL_FAIL'
  const [columnFilters, setColumnFilters] = useState({});
  const [sortConfig, setSortConfig] = useState({ column: null, direction: null });
  const [page, setPage] = useState(0);
  const pageSize = 50;
  const [showHelpModal, setShowHelpModal] = useState(false);
  const [sheetMetadata, setSheetMetadata] = useState({});

  const setStatus = (msg, type = 'info') => {
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
    return '';
  };

  const parseNumber = (val) => {
    if (val === null || val === undefined || val === '') return null;
    const num = Number(val);
    return isNaN(num) ? null : num;
  };

  // Process rows into ADES Aggregated Student-Course Result Format
  const processDataFromRows = (rows, currentHeaderMap) => {
    const groups = new Map();

    rows.forEach((row) => {
      const faculty = String(getCell(row, currentHeaderMap, 'Faculty', 'Fac', 'FacultyName', 'Department') || '').trim();
      const program = String(getCell(row, currentHeaderMap, 'Program Term Name', 'ProgramTermName', 'ProgramTerm', 'Program Term', 'Degree', 'Term', 'Semester') || '').trim();
      const seat = String(getCell(row, currentHeaderMap, 'Seat Number', 'SeatNumber', 'SeatNo', 'Seat_Number', 'RollNo', 'Roll Number') || '').trim();
      const prn = String(getCell(row, currentHeaderMap, 'PRN', 'PRN Number', 'PRNNo', 'RegisterNo', 'RegNo', 'StudentID') || '').trim();
      const code = String(getCell(row, currentHeaderMap, 'Course Code', 'CourseCode', 'PaperCode', 'SubjectCode', 'Course') || '').trim();
      const name = String(getCell(row, currentHeaderMap, 'Course Name', 'CourseName', 'PaperName', 'SubjectName', 'CourseTitle') || '').trim();

      const methodRaw = String(getCell(row, currentHeaderMap, 'Assessment Method', 'AssessmentMethod', 'AM', 'Method') || '').trim().toUpperCase();
      const typeRaw = String(getCell(row, currentHeaderMap, 'Assessment Type', 'AssessmentType', 'AT', 'Type') || '').trim().toUpperCase();
      const marksRaw = parseNumber(getCell(row, currentHeaderMap, 'Marks', 'ObtainedMarks', 'Mark', 'Obtained'));
      const atMaxRaw = parseNumber(getCell(row, currentHeaderMap, 'AT Max Marks', 'ATMaxMarks', 'MaxMarks', 'Max Marks', 'Max'));

      // Normalize Assessment Method (ESE vs CE/IA/CCA)
      let method = 'ESE';
      if (methodRaw.includes('CE') || methodRaw.includes('CA') || methodRaw.includes('IA') || methodRaw.includes('CCA') || methodRaw.includes('INTERNAL')) {
        method = 'CE';
      } else if (methodRaw.includes('ESE') || methodRaw.includes('EXT') || methodRaw.includes('EXTERNAL') || methodRaw.includes('THEORY')) {
        method = 'ESE';
      }

      // Normalize Assessment Type (PR vs TH)
      let type = 'TH';
      if (typeRaw.includes('PR') || typeRaw.includes('PRACTICAL') || typeRaw.includes('VIVA') || typeRaw.includes('LAB')) {
        type = 'PR';
      } else if (typeRaw.includes('TH') || typeRaw.includes('THEORY')) {
        type = 'TH';
      }

      const groupKey = `${faculty}|||${program}|||${seat}|||${prn}|||${code}|||${name}`;

      if (!groups.has(groupKey)) {
        groups.set(groupKey, {
          identifiers: { faculty, program, seat, prn, code, name },
          components: {}
        });
      }

      const group = groups.get(groupKey);
      group.components[`${method}_${type}`] = {
        marks: marksRaw !== null ? marksRaw : null,
        max: atMaxRaw !== null ? atMaxRaw : null,
        present: true
      };
    });

    const output = [];

    groups.forEach(({ identifiers, components }) => {
      const row = {
        'Faculty': identifiers.faculty,
        'Program Term Name': identifiers.program,
        'Course Code': identifiers.code,
        'Course Name': identifiers.name,
        'Seat Number': identifiers.seat,
        'PRN': identifiers.prn
      };

      // 1. ESE Components (Practical & Theory)
      const ese_pr = components['ESE_PR'];
      row['ESE - PR Max'] = (ese_pr && ese_pr.max !== null) ? ese_pr.max : '';
      row['ESE - PR Min'] = '';
      row['ESE - PR Obtained'] = (ese_pr && ese_pr.marks !== null) ? ese_pr.marks : '';

      const ese_th = components['ESE_TH'];
      row['ESE - TH Max'] = (ese_th && ese_th.max !== null) ? ese_th.max : '';
      row['ESE - TH Min'] = ese_th ? 0 : '';
      row['ESE - TH Obtained'] = (ese_th && ese_th.marks !== null) ? ese_th.marks : '';

      // ESE Totals & 30% Pass Minimum (Ceiling)
      const ese_pr_max = (ese_pr && ese_pr.max !== null) ? ese_pr.max : 0;
      const ese_th_max = (ese_th && ese_th.max !== null) ? ese_th.max : 0;
      const ese_max = ese_pr_max + ese_th_max;

      const ese_pr_obtained = (ese_pr && ese_pr.marks !== null) ? ese_pr.marks : 0;
      const ese_th_obtained = (ese_th && ese_th.marks !== null) ? ese_th.marks : 0;
      const ese_obtained = ese_pr_obtained + ese_th_obtained;

      row['ESE - Max'] = Math.round(ese_max);
      row['ESE - Min'] = Math.ceil(0.30 * ese_max);
      row['ESE Overall'] = Math.round(ese_obtained);

      // 2. CE Components (Practical & Theory)
      const ce_pr = components['CE_PR'];
      row['CE - PR Max'] = (ce_pr && ce_pr.max !== null) ? ce_pr.max : '';
      row['CE - PR Min'] = ce_pr ? 0 : '';
      row['CE - PR Obtained'] = (ce_pr && ce_pr.marks !== null) ? ce_pr.marks : '';

      const ce_th = components['CE_TH'];
      row['CE - TH Max'] = (ce_th && ce_th.max !== null) ? ce_th.max : '';
      row['CE - TH Min'] = ce_th ? 0 : '';
      row['CE - TH Obtained'] = (ce_th && ce_th.marks !== null) ? ce_th.marks : '';

      // CE Totals (Sum of CE - TH Max + CE - PR Max)
      const ce_pr_max = (ce_pr && ce_pr.max !== null) ? ce_pr.max : 0;
      const ce_th_max = (ce_th && ce_th.max !== null) ? ce_th.max : 0;
      const ce_max = ce_pr_max + ce_th_max;

      const ce_pr_obtained = (ce_pr && ce_pr.marks !== null) ? ce_pr.marks : 0;
      const ce_th_obtained = (ce_th && ce_th.marks !== null) ? ce_th.marks : 0;
      const ce_obtained = ce_pr_obtained + ce_th_obtained;

      row['CE - Max'] = Math.round(ce_max);
      row['CE - Min'] = 0;
      row['CE Overall Marks '] = Math.round(ce_obtained);

      // 3. Aggregate Course Calculations (Overall Maximum, 35% Ceiling, Course Overall)
      const overall_max = row['ESE - Max'] + row['CE - Max'];
      const overall_min = Math.ceil(0.35 * overall_max);
      const course_overall = row['ESE Overall'] + row['CE Overall Marks '];

      row['Overall Maximum'] = overall_max;
      row['Overall Minimum'] = overall_min;
      row['Course Overall Marks '] = course_overall;

      // 4. Pass / Fail Evaluations
      const ese_pass = row['ESE Overall'] >= row['ESE - Min'] ? 'Pass' : 'Fail';
      const overall_pass = course_overall >= overall_min ? 'Pass' : 'Fail';
      const course_pass = (ese_pass === 'Pass' && overall_pass === 'Pass') ? 'Pass' : 'Fail';

      row['ESE Pass (ESE obtained  >= ESE Min is pass if less then fail)'] = ese_pass;
      row['Overall pass (Course overall Marks >= Overall Minimum is pass if less then fail)'] = overall_pass;
      row['Course Pass/Fail ("Pass"only if both ESE Pass and Overall pass are pass if any condition is fail then course pass/fail status is fail )'] = course_pass;

      output.push(row);
    });

    return output;
  };

  const scoreHeaderRow = (rowArray) => {
    if (!Array.isArray(rowArray)) return { score: 0, matchedHeaders: [] };
    let score = 0;
    const matchedHeaders = [];
    const normCols = rowArray.map(c => normalizeKey(c));

    if (normCols.some(c => c.includes('assessmentmethod') || c === 'am' || c === 'method')) {
      score += 10;
      matchedHeaders.push('Assessment Method');
    }
    if (normCols.some(c => c.includes('assessmenttype') || c === 'at' || c === 'type')) {
      score += 10;
      matchedHeaders.push('Assessment Type');
    }
    if (normCols.some(c => c.includes('atmaxmarks') || c.includes('maxmarks') || c.includes('maxmark'))) {
      score += 8;
      matchedHeaders.push('AT Max Marks');
    }
    if (normCols.some(c => c === 'marks' || c.includes('obtained') || c === 'mark')) {
      score += 8;
      matchedHeaders.push('Marks');
    }
    if (normCols.some(c => c.includes('coursecode') || c.includes('papercode') || c.includes('subjectcode'))) {
      score += 6;
      matchedHeaders.push('Course Code');
    }
    if (normCols.some(c => c.includes('coursename') || c.includes('papername') || c.includes('subjectname'))) {
      score += 6;
      matchedHeaders.push('Course Name');
    }
    if (normCols.some(c => c.includes('prn') || c.includes('registerno') || c.includes('studentid'))) {
      score += 6;
      matchedHeaders.push('PRN');
    }
    if (normCols.some(c => c.includes('seat') || c.includes('rollno'))) {
      score += 4;
      matchedHeaders.push('Seat Number');
    }
    if (normCols.some(c => c.includes('program') || c.includes('semester') || c.includes('term'))) {
      score += 4;
      matchedHeaders.push('Program Term');
    }
    if (normCols.some(c => c.includes('faculty') || c.includes('department'))) {
      score += 4;
      matchedHeaders.push('Faculty');
    }

    return { score, matchedHeaders };
  };

  const parseSheetWithHeaderScan = (ws) => {
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
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
      const name = String(colName || '').trim();
      if (name) {
        const norm = normalizeKey(name);
        headerMap[norm] = name;
        validHeaderIndices.push({ idx, name, norm });
      }
    });

    const rows = [];
    for (let r = bestRowIdx + 1; r < aoa.length; r++) {
      const rowArr = aoa[r];
      if (!Array.isArray(rowArr) || rowArr.every(c => String(c || '').trim() === '')) continue;
      const rowObj = {};
      validHeaderIndices.forEach(({ idx, name }) => {
        rowObj[name] = rowArr[idx] !== undefined ? rowArr[idx] : '';
      });
      rows.push(rowObj);
    }

    return { rows, headerMap, headerRowIdx: bestRowIdx, score: maxScore, matchedHeaders: bestMatchedHeaders };
  };

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setStatus(`Analyzing workbook ${file.name}...`, 'info');
    setSourceFile(file.name);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const wb = XLSX.read(data, { type: 'array' });
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
          if (sNorm.includes('sourcefile') || sNorm.includes('source') || sNorm.includes('raw')) nameBonus += 10;
          else if (sNorm.includes('result') || sNorm.includes('marks') || sNorm.includes('exam')) nameBonus += 6;

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
          setStatus(`No valid data rows found in sheet "${bestSheet}".`, 'warning');
          setIsProcessing(false);
          return;
        }

        setHeaderMap(hMap);
        setRawRows(rows);

        const calculated = processDataFromRows(rows, hMap);
        setProcessedRows(calculated);
        setPage(0);

        const headerInfo = matchedHeaders.length > 0 ? ` (Detected Headers: ${matchedHeaders.join(', ')} on Row ${headerRowIdx + 1})` : '';
        setStatus(`Successfully calculated results for ${calculated.length} courses from "${bestSheet}"!${headerInfo}`, 'success');
      } catch (err) {
        console.error('Error parsing sheet:', err);
        setStatus(`Failed to read file: ${err.message}`, 'error');
      } finally {
        setIsProcessing(false);
      }
    };

    reader.onerror = () => {
      setStatus('File reading error.', 'error');
      setIsProcessing(false);
    };

    reader.readAsArrayBuffer(file);
  };

  const handleSheetChange = (sheetName) => {
    if (!workbook) return;
    setSelectedSheet(sheetName);
    setIsProcessing(true);
    setStatus(`Loading sheet "${sheetName}"...`, 'info');

    try {
      const ws = workbook.Sheets[sheetName];
      const { rows, headerMap: hMap, headerRowIdx, matchedHeaders } = parseSheetWithHeaderScan(ws);

      if (!rows || rows.length === 0) {
        setStatus(`Sheet "${sheetName}" is empty or has no valid rows.`, 'warning');
        setRawRows([]);
        setProcessedRows([]);
        setIsProcessing(false);
        return;
      }

      setHeaderMap(hMap);
      setRawRows(rows);

      const calculated = processDataFromRows(rows, hMap);
      setProcessedRows(calculated);
      setPage(0);

      const headerInfo = matchedHeaders.length > 0 ? ` (Headers: ${matchedHeaders.join(', ')} on Row ${headerRowIdx + 1})` : '';
      setStatus(`Loaded "${sheetName}": ${calculated.length} calculated course result rows.${headerInfo}`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to load sheet: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleSort = (column) => {
    setSortConfig(prev => {
      if (prev.column !== column) {
        return { column, direction: 'asc' };
      }
      if (prev.direction === 'asc') {
        return { column, direction: 'desc' };
      }
      return { column: null, direction: null };
    });
    setPage(0);
  };

  const updateColumnFilter = (colKey, val) => {
    setColumnFilters(prev => {
      const next = { ...prev };
      if (val === undefined || val === null || String(val).trim() === '') {
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
    setSelectedFacultyFilter('ALL');
    setSelectedProgramFilter('ALL');
    setSelectedCourseFilter('ALL');
    setSelectedResultFilter('ALL');
    setSearchQuery('');
    setPage(0);
  };

  // Unique Lists for Dropdown Filters
  const uniqueFaculties = useMemo(() => {
    const set = new Set();
    processedRows.forEach(r => { if (r['Faculty']) set.add(r['Faculty']); });
    return Array.from(set).sort();
  }, [processedRows]);

  const uniquePrograms = useMemo(() => {
    const set = new Set();
    processedRows.forEach(r => { if (r['Program Term Name']) set.add(r['Program Term Name']); });
    return Array.from(set).sort();
  }, [processedRows]);

  const uniqueCourses = useMemo(() => {
    const set = new Set();
    processedRows.forEach(r => { if (r['Course Code']) set.add(r['Course Code']); });
    return Array.from(set).sort();
  }, [processedRows]);

  // Statistics Metrics
  const metrics = useMemo(() => {
    const total = processedRows.length;
    if (total === 0) return { total: 0, uniqueStudents: 0, passed: 0, failed: 0, passPct: 0, eseFailed: 0, overallFailed: 0 };

    const prnSet = new Set();
    let passed = 0;
    let failed = 0;
    let eseFailed = 0;
    let overallFailed = 0;

    const coursePassKey = 'Course Pass/Fail ("Pass"only if both ESE Pass and Overall pass are pass if any condition is fail then course pass/fail status is fail )';
    const esePassKey = 'ESE Pass (ESE obtained  >= ESE Min is pass if less then fail)';
    const overallPassKey = 'Overall pass (Course overall Marks >= Overall Minimum is pass if less then fail)';

    processedRows.forEach(r => {
      if (r['PRN']) prnSet.add(r['PRN']);
      if (r[coursePassKey] === 'Pass') passed++;
      else failed++;

      if (r[esePassKey] === 'Fail') eseFailed++;
      if (r[overallPassKey] === 'Fail') overallFailed++;
    });

    return {
      total,
      uniqueStudents: prnSet.size,
      passed,
      failed,
      passPct: ((passed / total) * 100).toFixed(1),
      eseFailed,
      overallFailed
    };
  }, [processedRows]);

  // Filtered & Sorted Rows
  const filteredRows = useMemo(() => {
    let result = [...processedRows];

    const coursePassKey = 'Course Pass/Fail ("Pass"only if both ESE Pass and Overall pass are pass if any condition is fail then course pass/fail status is fail )';
    const esePassKey = 'ESE Pass (ESE obtained  >= ESE Min is pass if less then fail)';
    const overallPassKey = 'Overall pass (Course overall Marks >= Overall Minimum is pass if less then fail)';

    // 1. Result Status Filter
    if (selectedResultFilter === 'PASS') {
      result = result.filter(r => r[coursePassKey] === 'Pass');
    } else if (selectedResultFilter === 'FAIL') {
      result = result.filter(r => r[coursePassKey] === 'Fail');
    } else if (selectedResultFilter === 'ESE_FAIL') {
      result = result.filter(r => r[esePassKey] === 'Fail');
    } else if (selectedResultFilter === 'OVERALL_FAIL') {
      result = result.filter(r => r[overallPassKey] === 'Fail');
    }

    // 2. Dropdown Filters
    if (selectedFacultyFilter !== 'ALL') {
      result = result.filter(r => r['Faculty'] === selectedFacultyFilter);
    }
    if (selectedProgramFilter !== 'ALL') {
      result = result.filter(r => r['Program Term Name'] === selectedProgramFilter);
    }
    if (selectedCourseFilter !== 'ALL') {
      result = result.filter(r => r['Course Code'] === selectedCourseFilter);
    }

    // 3. Search Query
    if (searchQuery) {
      const q = searchQuery.toLowerCase().trim();
      result = result.filter(r =>
        Object.values(r).some(val => String(val || '').toLowerCase().includes(q))
      );
    }

    // 4. Column Filters
    const activeFilterEntries = Object.entries(columnFilters);
    if (activeFilterEntries.length > 0) {
      result = result.filter(r => {
        return activeFilterEntries.every(([colKey, filterVal]) => {
          if (!filterVal || String(filterVal).trim() === '') return true;
          const cellVal = String(r[colKey] !== undefined && r[colKey] !== null ? r[colKey] : '').toLowerCase();
          return cellVal.includes(String(filterVal).toLowerCase().trim());
        });
      });
    }

    // 5. Sorting
    if (sortConfig.column && sortConfig.direction) {
      const col = sortConfig.column;
      const dir = sortConfig.direction === 'asc' ? 1 : -1;

      result.sort((a, b) => {
        const valA = a[col] !== undefined && a[col] !== null ? a[col] : '';
        const valB = b[col] !== undefined && b[col] !== null ? b[col] : '';

        const numA = Number(valA);
        const numB = Number(valB);
        const isNumA = typeof valA === 'number' || (String(valA).trim() !== '' && !isNaN(numA));
        const isNumB = typeof valB === 'number' || (String(valB).trim() !== '' && !isNaN(numB));

        if (isNumA && isNumB) {
          return (numA - numB) * dir;
        }

        return String(valA).localeCompare(String(valB), undefined, { numeric: true, sensitivity: 'base' }) * dir;
      });
    }

    return result;
  }, [processedRows, selectedResultFilter, selectedFacultyFilter, selectedProgramFilter, selectedCourseFilter, searchQuery, columnFilters, sortConfig]);

  const pagedRows = useMemo(() => {
    const start = page * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, page]);

  // Export to Excel Matching exact python structure
  const handleExportExcel = (rowsToExport = filteredRows, filename = 'converted_output_card.xlsx') => {
    if (!rowsToExport || rowsToExport.length === 0) {
      alert('No rows to export.');
      return;
    }

    setIsProcessing(true);
    setStatus('Generating ADES Result Excel export...', 'info');

    try {
      const aoa = [
        ADES_OUTPUT_HEADERS,
        ...rowsToExport.map(r => ADES_OUTPUT_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))
      ];

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });

      ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: aoa.length - 1, c: ADES_OUTPUT_HEADERS.length - 1 }
        })
      };

      ws['!cols'] = ADES_OUTPUT_HEADERS.map(h => ({
        wch: Math.min(Math.max(h.length + 3, 14), 45)
      }));

      XLSX.utils.book_append_sheet(wb, ws, 'Output file ');
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);

      setStatus(`Successfully exported ${rowsToExport.length} calculated rows to ${filename}!`, 'success');
    } catch (err) {
      console.error('Export Error:', err);
      setStatus(`Export failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const coursePassKey = 'Course Pass/Fail ("Pass"only if both ESE Pass and Overall pass are pass if any condition is fail then course pass/fail status is fail )';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', overflow: 'hidden', background: 'var(--bg)', color: 'var(--ink)' }}>
      
      {/* Top Navigation Bar */}
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 20px', borderBottom: '1px solid var(--line)', background: 'var(--panel)', flexShrink: 0 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
          <Link to="/" style={{ display: 'flex', alignItems: 'center', gap: '6px', color: 'var(--muted)', textDecoration: 'none', fontSize: '13px', fontWeight: 600 }}>
            <ArrowLeft size={16} /> Back to Portal
          </Link>
          <div style={{ height: '18px', width: '1px', background: 'var(--line)' }} />
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Calculator size={20} color="var(--accent)" />
            <h2 style={{ fontSize: '16px', fontWeight: 700, margin: 0 }}>ADES Result Calculator</h2>
            <span style={{ fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '12px', fontWeight: 600 }}>
              ESE & CE Evaluator
            </span>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <button 
            type="button" 
            className="secondary" 
            onClick={() => setShowHelpModal(true)}
            style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', padding: '6px 12px' }}
          >
            <HelpCircle size={14} /> Extraction Logic & Guide
          </button>

          {processedRows.length > 0 && (
            <button 
              type="button" 
              onClick={() => handleExportExcel(filteredRows, 'converted_output_card.xlsx')}
              style={{ 
                display: 'flex', 
                alignItems: 'center', 
                gap: '6px', 
                padding: '6px 14px', 
                fontSize: '12.5px', 
                background: 'var(--accent)', 
                color: 'white', 
                border: 'none', 
                borderRadius: '6px', 
                fontWeight: 600, 
                cursor: 'pointer' 
              }}
            >
              <Download size={14} /> Export Output Excel ({filteredRows.length} Rows)
            </button>
          )}
        </div>
      </header>

      {/* Main Workspace Layout */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        
        {/* Left Sidebar / Config Panel */}
        <aside style={{ width: '320px', borderRight: '1px solid var(--line)', background: 'var(--panel)', display: 'flex', flexDirection: 'column', flexShrink: 0, overflowY: 'auto', padding: '16px', gap: '16px' }}>
          
          {/* File Upload Box */}
          <div style={{ background: 'var(--bg)', border: '1.5px dashed var(--line)', borderRadius: '8px', padding: '16px', textAlign: 'center', position: 'relative' }}>
            <input 
              type="file" 
              accept=".xlsx,.xls,.csv" 
              onChange={handleFileUpload}
              style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', opacity: 0, cursor: 'pointer' }}
            />
            <FileSpreadsheet size={32} color="var(--accent)" style={{ margin: '0 auto 8px', opacity: 0.8 }} />
            <div style={{ fontSize: '13px', fontWeight: 600, color: 'var(--ink)' }}>
              {sourceFile ? sourceFile : 'Upload ADES Marks Excel'}
            </div>
            <div style={{ fontSize: '11px', color: 'var(--muted)', marginTop: '4px' }}>
              Drop .xlsx / .xls file here (e.g. new tool card_updated.xlsx)
            </div>
          </div>

          {/* Sheet Selector (if multiple sheets exist) */}
          {sheetNames.length > 1 && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <label style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)' }}>Select Source Sheet:</label>
              <select 
                value={selectedSheet} 
                onChange={(e) => handleSheetChange(e.target.value)}
                style={{ padding: '6px 10px', fontSize: '12px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)' }}
              >
                {sheetNames.map(s => {
                  const meta = sheetMetadata[s];
                  const hasHeaders = meta && meta.matchedHeaders && meta.matchedHeaders.length > 0;
                  return (
                    <option key={s} value={s}>
                      {s} {hasHeaders ? `(✓ ${meta.matchedHeaders.length} headers)` : ''}
                    </option>
                  );
                })}
              </select>
            </div>
          )}

          {/* Result Overview Stat Card */}
          {processedRows.length > 0 && (
            <div style={{ background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px', padding: '12px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <div style={{ fontSize: '12px', fontWeight: 700, color: 'var(--ink)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <Sparkles size={14} color="var(--accent)" /> Result Evaluation Metrics
              </div>
              
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px', fontSize: '11px' }}>
                <div style={{ background: 'var(--panel)', padding: '8px', borderRadius: '6px', border: '1px solid var(--line)' }}>
                  <div style={{ color: 'var(--muted)' }}>Evaluated Courses</div>
                  <strong style={{ fontSize: '15px', color: 'var(--ink)' }}>{metrics.total}</strong>
                </div>

                <div style={{ background: 'var(--panel)', padding: '8px', borderRadius: '6px', border: '1px solid var(--line)' }}>
                  <div style={{ color: 'var(--muted)' }}>Unique Students</div>
                  <strong style={{ fontSize: '15px', color: 'var(--ink)' }}>{metrics.uniqueStudents}</strong>
                </div>

                <div style={{ background: 'rgba(16, 185, 129, 0.1)', padding: '8px', borderRadius: '6px', border: '1px solid rgba(16, 185, 129, 0.3)' }}>
                  <div style={{ color: '#10b981', fontWeight: 600 }}>Passed Courses</div>
                  <strong style={{ fontSize: '15px', color: '#10b981' }}>{metrics.passed}</strong>
                  <div style={{ fontSize: '10px', color: '#10b981' }}>({metrics.passPct}%)</div>
                </div>

                <div style={{ background: 'rgba(239, 68, 68, 0.1)', padding: '8px', borderRadius: '6px', border: '1px solid rgba(239, 68, 68, 0.3)' }}>
                  <div style={{ color: '#ef4444', fontWeight: 600 }}>Failed Courses</div>
                  <strong style={{ fontSize: '15px', color: '#ef4444' }}>{metrics.failed}</strong>
                  <div style={{ fontSize: '10px', color: '#ef4444' }}>({(100 - metrics.passPct).toFixed(1)}%)</div>
                </div>
              </div>

              {/* Failure Breakdown */}
              {metrics.failed > 0 && (
                <div style={{ background: 'var(--panel)', padding: '8px', borderRadius: '6px', border: '1px solid var(--line)', fontSize: '11px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <div style={{ fontWeight: 600, color: 'var(--ink)' }}>Failure Breakdown:</div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', color: 'var(--muted)' }}>
                    <span>ESE Failed (&lt; 30%):</span>
                    <strong style={{ color: '#ef4444' }}>{metrics.eseFailed}</strong>
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', color: 'var(--muted)' }}>
                    <span>Aggregate Failed (&lt; 35%):</span>
                    <strong style={{ color: '#ef4444' }}>{metrics.overallFailed}</strong>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* Quick Rules Summary */}
          <div style={{ background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px', padding: '12px', fontSize: '11.5px', color: 'var(--muted)', lineHeight: '1.4' }}>
            <strong style={{ color: 'var(--ink)', display: 'block', marginBottom: '6px' }}>Evaluation Rules:</strong>
            <ul style={{ margin: 0, paddingLeft: '16px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
              <li><strong>ESE Min:</strong> <code>ceil(30% × ESE Max)</code></li>
              <li><strong>ESE Pass:</strong> <code>ESE Overall &ge; ESE Min</code></li>
              <li><strong>Overall Min:</strong> <code>ceil(35% × Overall Max)</code></li>
              <li><strong>Overall Pass:</strong> <code>Course Overall &ge; Overall Min</code></li>
              <li><strong>Final Course Status:</strong> <code>"Pass" only if both ESE Pass and Overall Pass are met</code></li>
            </ul>
          </div>

          {/* Status Message */}
          <div style={{ marginTop: 'auto', padding: '8px 12px', borderRadius: '6px', fontSize: '11.5px', background: statusType === 'error' ? 'var(--danger-soft)' : statusType === 'success' ? 'var(--accent-soft)' : 'var(--bg)', color: statusType === 'error' ? 'var(--danger)' : statusType === 'success' ? 'var(--accent)' : 'var(--muted)', border: '1px solid var(--line)' }}>
            {statusMsg}
          </div>

        </aside>

        {/* Right Content / Data Table View */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', background: 'var(--bg)' }}>
          
          {/* Interactive Toolbar */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '10px 16px', borderBottom: '1px solid var(--line)', background: 'var(--panel)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
              
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                
                {/* Search Box */}
                <div style={{ position: 'relative', width: '180px' }}>
                  <input 
                    type="text" 
                    placeholder="Search all columns..." 
                    value={searchQuery} 
                    onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }} 
                    style={{ width: '100%', padding: '5px 8px 5px 26px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--bg)' }} 
                  />
                  <Search size={13} color="var(--muted)" style={{ position: 'absolute', left: '8px', top: '7px' }} />
                </div>

                {/* Result Filter Tabs */}
                <div style={{ display: 'flex', alignItems: 'center', background: 'var(--bg)', padding: '2px', borderRadius: '6px', border: '1px solid var(--line)' }}>
                  <button 
                    type="button" 
                    onClick={() => { setSelectedResultFilter('ALL'); setPage(0); }}
                    style={{ padding: '3px 8px', fontSize: '11px', fontWeight: 600, border: 'none', borderRadius: '4px', cursor: 'pointer', background: selectedResultFilter === 'ALL' ? 'var(--panel)' : 'transparent', color: selectedResultFilter === 'ALL' ? 'var(--ink)' : 'var(--muted)' }}
                  >
                    All ({processedRows.length})
                  </button>
                  <button 
                    type="button" 
                    onClick={() => { setSelectedResultFilter('PASS'); setPage(0); }}
                    style={{ padding: '3px 8px', fontSize: '11px', fontWeight: 600, border: 'none', borderRadius: '4px', cursor: 'pointer', background: selectedResultFilter === 'PASS' ? 'rgba(16,185,129,0.15)' : 'transparent', color: selectedResultFilter === 'PASS' ? '#10b981' : 'var(--muted)' }}
                  >
                    Passed ({metrics.passed})
                  </button>
                  <button 
                    type="button" 
                    onClick={() => { setSelectedResultFilter('FAIL'); setPage(0); }}
                    style={{ padding: '3px 8px', fontSize: '11px', fontWeight: 600, border: 'none', borderRadius: '4px', cursor: 'pointer', background: selectedResultFilter === 'FAIL' ? 'rgba(239,68,68,0.15)' : 'transparent', color: selectedResultFilter === 'FAIL' ? '#ef4444' : 'var(--muted)' }}
                  >
                    Failed ({metrics.failed})
                  </button>
                </div>

                {/* Program Filter */}
                {uniquePrograms.length > 0 && (
                  <select 
                    value={selectedProgramFilter} 
                    onChange={(e) => { setSelectedProgramFilter(e.target.value); setPage(0); }}
                    style={{ padding: '5px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--bg)', maxWidth: '140px' }}
                  >
                    <option value="ALL">All Programs ({uniquePrograms.length})</option>
                    {uniquePrograms.map(p => (<option key={p} value={p}>{p}</option>))}
                  </select>
                )}

                {/* Course Code Filter */}
                {uniqueCourses.length > 0 && (
                  <select 
                    value={selectedCourseFilter} 
                    onChange={(e) => { setSelectedCourseFilter(e.target.value); setPage(0); }}
                    style={{ padding: '5px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--bg)', maxWidth: '130px' }}
                  >
                    <option value="ALL">All Courses ({uniqueCourses.length})</option>
                    {uniqueCourses.map(c => (<option key={c} value={c}>{c}</option>))}
                  </select>
                )}

                {/* Faculty Filter */}
                {uniqueFaculties.length > 0 && (
                  <select 
                    value={selectedFacultyFilter} 
                    onChange={(e) => { setSelectedFacultyFilter(e.target.value); setPage(0); }}
                    style={{ padding: '5px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--bg)', maxWidth: '130px' }}
                  >
                    <option value="ALL">All Faculties</option>
                    {uniqueFaculties.map(f => (<option key={f} value={f}>{f}</option>))}
                  </select>
                )}

                {/* Reset Filters */}
                {(selectedFacultyFilter !== 'ALL' || selectedProgramFilter !== 'ALL' || selectedCourseFilter !== 'ALL' || selectedResultFilter !== 'ALL' || searchQuery || Object.keys(columnFilters).length > 0 || sortConfig.column) && (
                  <button 
                    type="button" 
                    className="secondary" 
                    onClick={clearAllFilters}
                    style={{ padding: '3px 8px', fontSize: '11px', color: 'var(--danger)' }}
                  >
                    Reset
                  </button>
                )}
              </div>

              {/* Pagination Controls */}
              {filteredRows.length > pageSize && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', color: 'var(--muted)' }}>
                  <span>Page {page + 1} of {Math.ceil(filteredRows.length / pageSize)}</span>
                  <button 
                    type="button" 
                    className="secondary" 
                    disabled={page === 0} 
                    onClick={() => setPage(p => p - 1)} 
                    style={{ padding: '2px 6px', fontSize: '11px' }}
                  >
                    Prev
                  </button>
                  <button 
                    type="button" 
                    className="secondary" 
                    disabled={(page + 1) * pageSize >= filteredRows.length} 
                    onClick={() => setPage(p => p + 1)} 
                    style={{ padding: '2px 6px', fontSize: '11px' }}
                  >
                    Next
                  </button>
                </div>
              )}
            </div>

            {/* Active Sort / Column Filter Badges */}
            {(sortConfig.column || Object.keys(columnFilters).length > 0) && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap', paddingTop: '2px' }}>
                <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 600 }}>Active:</span>
                
                {sortConfig.column && (
                  <span style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '10.5px', background: 'var(--accent-soft)', color: 'var(--accent)', border: '1px solid var(--accent)', padding: '2px 6px', borderRadius: '4px', fontWeight: 600 }}>
                    Sorted: {sortConfig.column} ({sortConfig.direction === 'asc' ? '▲ ASC' : '▼ DESC'})
                    <button type="button" onClick={() => setSortConfig({ column: null, direction: null })} style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--accent)', display: 'flex' }}>
                      <X size={11} />
                    </button>
                  </span>
                )}

                {Object.entries(columnFilters).map(([colKey, filterVal]) => (
                  <span key={colKey} style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '10.5px', background: 'var(--accent-soft)', color: 'var(--accent)', border: '1px solid var(--accent)', padding: '2px 6px', borderRadius: '4px', fontWeight: 600 }}>
                    {colKey}: "{filterVal}"
                    <button type="button" onClick={() => updateColumnFilter(colKey, '')} style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--accent)', display: 'flex' }}>
                      <X size={11} />
                    </button>
                  </span>
                ))}
              </div>
            )}
          </div>

          {/* Table Container */}
          <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
            {processedRows.length === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'var(--muted)', gap: '10px' }}>
                <Calculator size={44} style={{ opacity: 0.3 }} />
                <strong style={{ fontSize: '15px', color: 'var(--ink)' }}>No Course Marks Loaded</strong>
                <span style={{ fontSize: '12px', maxWidth: '420px', textAlign: 'center' }}>
                  Upload your <code>new tool card_updated.xlsx</code> on the left to aggregate ESE & CE marks and calculate course pass/fail results.
                </span>
              </div>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px', whiteSpace: 'nowrap' }}>
                <thead>
                  <tr style={{ background: 'var(--panel)', position: 'sticky', top: 0, zIndex: 10 }}>
                    <th style={{ padding: '8px 10px', textAlign: 'center', borderBottom: '1.5px solid var(--line)', color: 'var(--muted)', width: '40px' }}>#</th>
                    {ADES_OUTPUT_HEADERS.map((col, idx) => {
                      const isSorted = sortConfig.column === col;
                      const isStatusCol = col.includes('Pass') || col.includes('pass') || col.includes('Fail');
                      const isOverallCol = col.includes('Overall');
                      const isEseCol = col.startsWith('ESE');
                      const isCeCol = col.startsWith('CE');

                      let thBg = 'var(--panel)';
                      if (isSorted) thBg = 'var(--accent-soft)';
                      else if (isStatusCol) thBg = 'rgba(23, 107, 135, 0.08)';

                      return (
                        <th 
                          key={col} 
                          onClick={() => handleSort(col)}
                          title={`Click to sort by Col ${idx + 1} (${col})`}
                          style={{ 
                            padding: '8px 10px', 
                            textAlign: 'left', 
                            borderBottom: '1.5px solid var(--line)', 
                            color: isSorted ? 'var(--accent)' : 'var(--ink)', 
                            fontWeight: 700,
                            borderRight: '1px solid var(--line)',
                            background: thBg,
                            cursor: 'pointer',
                            userSelect: 'none'
                          }}
                        >
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '6px' }}>
                            <div>
                              <span style={{ fontSize: '9px', color: isSorted ? 'var(--accent)' : 'var(--muted)', display: 'block', fontWeight: 600 }}>Col {idx + 1}</span>
                              {col}
                            </div>
                            <div style={{ display: 'flex', alignItems: 'center' }}>
                              {isSorted && sortConfig.direction === 'asc' && <ArrowUp size={12} color="var(--accent)" />}
                              {isSorted && sortConfig.direction === 'desc' && <ArrowDown size={12} color="var(--accent)" />}
                              {!isSorted && <ArrowUpDown size={11} style={{ opacity: 0.25 }} />}
                            </div>
                          </div>
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {pagedRows.map((row, rowIdx) => {
                    const actualIdx = (page * pageSize) + rowIdx;
                    const finalStatus = row[coursePassKey];
                    const isRowPass = finalStatus === 'Pass';

                    return (
                      <tr 
                        key={actualIdx} 
                        style={{ 
                          borderBottom: '1px solid var(--line)',
                          background: isRowPass ? 'transparent' : 'rgba(239, 68, 68, 0.02)'
                        }}
                      >
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: 'var(--muted)', borderRight: '1px solid var(--line)', fontWeight: 600 }}>
                          {actualIdx + 1}
                        </td>
                        {ADES_OUTPUT_HEADERS.map(col => {
                          const val = row[col];
                          const isPassFail = val === 'Pass' || val === 'Fail';

                          return (
                            <td 
                              key={col} 
                              style={{ 
                                padding: '6px 10px', 
                                borderRight: '1px solid var(--line)',
                                color: (val === 'Pass') ? '#10b981' : (val === 'Fail') ? '#ef4444' : 'var(--ink)',
                                fontWeight: (isPassFail || col === 'PRN' || col === 'Seat Number' || col === 'Course Code') ? 700 : 400
                              }}
                            >
                              {isPassFail ? (
                                <span style={{
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '4px',
                                  padding: '2px 8px',
                                  borderRadius: '4px',
                                  fontSize: '11px',
                                  background: val === 'Pass' ? 'rgba(16, 185, 129, 0.15)' : 'rgba(239, 68, 68, 0.15)',
                                  color: val === 'Pass' ? '#10b981' : '#ef4444'
                                }}>
                                  {val === 'Pass' ? <CheckCircle2 size={12} /> : <XCircle size={12} />}
                                  {val}
                                </span>
                              ) : (
                                val !== undefined ? String(val) : ''
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>

          {/* Footer Info */}
          {processedRows.length > 0 && (
            <div style={{ padding: '8px 16px', background: 'var(--panel)', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
              <span>Showing {pagedRows.length} of {filteredRows.length} filtered rows (Total evaluated courses: <strong>{processedRows.length}</strong>)</span>
              <span>30 Standard ADES Result Calculation Columns</span>
            </div>
          )}

        </main>
      </div>

      {/* Extraction Logic & Help Modal */}
      {showHelpModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0,0,0,0.6)',
          zIndex: 1000,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          padding: '20px'
        }}>
          <div style={{
            background: 'var(--panel)',
            border: '1px solid var(--line)',
            borderRadius: '12px',
            width: '640px',
            maxWidth: '90vw',
            maxHeight: '85vh',
            display: 'flex',
            flexDirection: 'column',
            boxShadow: '0 20px 40px rgba(0,0,0,0.3)',
            overflow: 'hidden'
          }}>
            <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Calculator size={20} color="var(--accent)" />
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>ADES Result Calculator - Logic & Rules</h3>
              </div>
              <button 
                type="button" 
                onClick={() => setShowHelpModal(false)}
                style={{ background: 'none', border: 'none', color: 'var(--muted)', cursor: 'pointer', padding: '4px' }}
              >
                <X size={18} />
              </button>
            </div>

            <div style={{ padding: '20px', overflowY: 'auto', fontSize: '12.5px', lineHeight: '1.6', display: 'flex', flexDirection: 'column', gap: '14px' }}>
              <div>
                <strong>1. Student-Course Grouping:</strong>
                <p style={{ margin: '4px 0', color: 'var(--muted)' }}>
                  Groups records uniquely by <code>Faculty</code>, <code>Program Term Name</code>, <code>Seat Number</code>, <code>PRN</code>, <code>Course Code</code>, and <code>Course Name</code>.
                </p>
              </div>

              <div>
                <strong>2. Component Mapping (Assessment Method & Type):</strong>
                <p style={{ margin: '4px 0', color: 'var(--muted)' }}>
                  Extracts <code>(Method, Type)</code> metrics:
                </p>
                <ul style={{ margin: '2px 0 0', paddingLeft: '16px', color: 'var(--ink)' }}>
                  <li><code>(ESE, PR)</code>: Practical End Semester Exam Marks & Max</li>
                  <li><code>(ESE, TH)</code>: Theory End Semester Exam Marks & Max</li>
                  <li><code>(CE, PR)</code>: Practical Continuous Evaluation Marks & Max</li>
                  <li><code>(CE, TH)</code>: Theory Continuous Evaluation Marks & Max</li>
                </ul>
              </div>

              <div>
                <strong>3. ESE Minimum & 30% Pass Rule:</strong>
                <pre style={{ background: 'var(--bg)', padding: '8px 12px', borderRadius: '6px', border: '1px solid var(--line)', marginTop: '4px', fontSize: '11.5px' }}>
                  ESE Max = ESE - PR Max + ESE - TH Max
                  ESE Min = ceil(0.30 * ESE Max)   // 30% Pass Threshold
                  ESE Overall = ESE - PR Obtained + ESE - TH Obtained
                  ESE Pass = (ESE Overall &ge; ESE Min) ? "Pass" : "Fail"
                </pre>
              </div>

              <div>
                <strong>4. CE Totals & Aggregate 35% Overall Pass Rule:</strong>
                <pre style={{ background: 'var(--bg)', padding: '8px 12px', borderRadius: '6px', border: '1px solid var(--line)', marginTop: '4px', fontSize: '11.5px' }}>
                  CE Max = CE - PR Max + CE - TH Max
                  CE Overall = CE - PR Obtained + CE - TH Obtained
                  Overall Maximum = ESE Max + CE Max
                  Overall Minimum = ceil(0.35 * Overall Maximum)  // 35% Pass Threshold
                  Course Overall = ESE Overall + CE Overall
                  Overall Pass = (Course Overall &ge; Overall Minimum) ? "Pass" : "Fail"
                </pre>
              </div>

              <div>
                <strong>5. Final Course Pass/Fail Condition:</strong>
                <p style={{ margin: '4px 0', color: 'var(--muted)' }}>
                  A course is marked <strong>"Pass" only if both ESE Pass and Overall Pass conditions are satisfied</strong>. If any condition fails, the final status is <strong>"Fail"</strong>.
                </p>
              </div>
            </div>

            <div style={{ padding: '12px 20px', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'flex-end', background: 'var(--bg)' }}>
              <button 
                type="button" 
                onClick={() => setShowHelpModal(false)}
                style={{ padding: '6px 16px', background: 'var(--accent)', color: 'white', borderRadius: '6px', border: 'none', fontWeight: 600, cursor: 'pointer' }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
