import React, { useState, useMemo } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { 
  ArrowLeft, 
  Upload, 
  Download, 
  FileSpreadsheet, 
  CheckCircle2, 
  AlertTriangle, 
  Search, 
  Filter, 
  RefreshCw,
  Table,
  Sparkles,
  HelpCircle,
  Layers,
  GraduationCap,
  Building2,
  BookOpen,
  ArrowUp,
  ArrowDown,
  ArrowUpDown,
  X,
  ListFilter,
  Layers2,
  FileCheck2
} from 'lucide-react';

// Format 1: 7-Column Deduplicated Report
export const DEDUPLICATED_HEADERS = [
  'College Code', 
  'College Name', 
  'Programme Year', 
  'Program Term Name', 
  'Course Details',
  'Course Code', 
  'Course Name'
];

// Format 2: 9-Column All Rows Exploded Report
export const ALL_ROWS_HEADERS = [
  'College Code', 
  'College Name', 
  'Program Code', 
  'Program Term',
  'Programme Year', 
  'Program Term Name', 
  'Course Details',
  'Course Code', 
  'Course Name'
];

const normalizeKey = (key) => String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');

export default function AffiliatedProgrammePage() {
  const [sourceFile, setSourceFile] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const [workbook, setWorkbook] = useState(null);
  const [rawRows, setRawRows] = useState([]);
  const [allExplodedRows, setAllExplodedRows] = useState([]);
  const [deduplicatedRows, setDeduplicatedRows] = useState([]);
  const [duplicatesCount, setDuplicatesCount] = useState(0);
  const [activeTab, setActiveTab] = useState('deduplicated'); // 'deduplicated' (7 cols) or 'all_rows' (9 cols)

  const [headerMap, setHeaderMap] = useState({});
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('info');
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedCollegeFilter, setSelectedCollegeFilter] = useState('ALL');
  const [selectedProgramFilter, setSelectedProgramFilter] = useState('ALL');
  const [selectedYearFilter, setSelectedYearFilter] = useState('ALL');
  const [selectedSemesterFilter, setSelectedSemesterFilter] = useState('ALL');
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
        return String(row[actualKey]).trim();
      }
    }
    return '';
  };

  // Process raw rows into both (1) All Rows Exploded (9 cols) and (2) Deduplicated Unique (7 cols)
  const processDataFromRows = (rows, currentHeaderMap) => {
    const allExploded = [];
    const dedupeList = [];
    const seen = new Set();
    let dupCount = 0;

    rows.forEach((row) => {
      const collegeCode = getCell(row, currentHeaderMap, 'ADEC Code', 'ADECCode', 'ADEC_Code', 'ADEC', 'College Code', 'CollegeCode', 'College_Code', 'InstCode', 'CenterCode', 'Code');
      const collegeName = getCell(row, currentHeaderMap, 'ADEC Name', 'ADECName', 'ADEC_Name', 'ADEC', 'College Name', 'CollegeName', 'College_Name', 'InstituteName', 'CenterName', 'College');
      const programCode = getCell(row, currentHeaderMap, 'Program Code', 'ProgramCode', 'Program_Code', 'ProgCode', 'DegreeCode', 'Program');
      const programTerm = getCell(row, currentHeaderMap, 'Program Term', 'ProgramTerm', 'Program_Term', 'Term', 'SemesterYear', 'Sem');
      const rawCourseDetails = getCell(row, currentHeaderMap, 'Course Details', 'CourseDetails', 'Course_Details', 'Courses', 'Subjects', 'SubjectDetails');

      // Extract Year and Semester from 'Program Term'
      const yearMatch = programTerm.match(/(Year\s+[IVXLCDM\d]+)/i);
      const semesterMatch = programTerm.match(/(SEMESTER\s+[IVXLCDM\d]+)/i);
      const programmeYear = yearMatch ? yearMatch[1] : '';
      const programTermName = semesterMatch ? semesterMatch[1] : '';

      // Split and explode the 'Course Details' column by commas separating courses: r',\s*(?=\()'
      let courseItems = [];
      if (rawCourseDetails) {
        courseItems = rawCourseDetails.split(/,\s*(?=\()/).map(s => s.trim()).filter(Boolean);
      }

      if (courseItems.length === 0) {
        const itemCourseDetails = rawCourseDetails || '';
        
        // Format 2: All Rows (9 columns)
        allExploded.push({
          'College Code': collegeCode,
          'College Name': collegeName,
          'Program Code': programCode,
          'Program Term': programTerm,
          'Programme Year': programmeYear,
          'Program Term Name': programTermName,
          'Course Details': itemCourseDetails,
          'Course Code': '',
          'Course Name': itemCourseDetails
        });

        // Format 1: Deduplicated (7 columns)
        const dedupeKey = `${collegeCode.trim().toLowerCase()}|${programCode.trim().toLowerCase()}|${programmeYear.trim().toLowerCase()}|${programTermName.trim().toLowerCase()}|${itemCourseDetails.trim().toLowerCase()}`;
        if (dedupeKey.replace(/\|/g, '') && seen.has(dedupeKey)) {
          dupCount++;
        } else {
          if (dedupeKey.replace(/\|/g, '')) seen.add(dedupeKey);
          dedupeList.push({
            'College Code': collegeCode,
            'College Name': collegeName,
            'Programme Year': programmeYear,
            'Program Term Name': programTermName,
            'Course Details': itemCourseDetails,
            'Course Code': '',
            'Course Name': itemCourseDetails
          });
        }
      } else {
        courseItems.forEach((courseStr) => {
          // Extract 'Course Code' and 'Course Name' from 'Course Details': r'^\((?P<Course_Code>[^)]+)\)\s*(?P<Course_Name>.*)$'
          const coursePatternMatch = courseStr.match(/^\(([^)]+)\)\s*(.*)$/);
          let courseCode = '';
          let courseName = '';

          if (coursePatternMatch) {
            courseCode = (coursePatternMatch[1] || '').trim();
            courseName = (coursePatternMatch[2] || '').trim();
          } else {
            courseName = courseStr;
          }

          // Format 2: All Rows (9 columns)
          allExploded.push({
            'College Code': collegeCode,
            'College Name': collegeName,
            'Program Code': programCode,
            'Program Term': programTerm,
            'Programme Year': programmeYear,
            'Program Term Name': programTermName,
            'Course Details': courseStr,
            'Course Code': courseCode,
            'Course Name': courseName
          });

          // Format 1: Deduplicated (7 columns)
          // Key: College + Program + Year + Term + (Course Code / Course Name)
          const distinctCourseIdentifier = (courseCode || courseName || courseStr).trim().toLowerCase();
          const dedupeKey = `${collegeCode.trim().toLowerCase()}|${programCode.trim().toLowerCase()}|${programmeYear.trim().toLowerCase()}|${programTermName.trim().toLowerCase()}|${distinctCourseIdentifier}`;

          if (seen.has(dedupeKey)) {
            dupCount++;
          } else {
            seen.add(dedupeKey);
            dedupeList.push({
              'College Code': collegeCode,
              'College Name': collegeName,
              'Programme Year': programmeYear,
              'Program Term Name': programTermName,
              'Course Details': courseStr,
              'Course Code': courseCode,
              'Course Name': courseName
            });
          }
        });
      }
    });

    setDuplicatesCount(dupCount);
    return { allExploded, dedupeList, dupCount };
  };

  const scoreHeaderRow = (rowArray) => {
    if (!Array.isArray(rowArray)) return { score: 0, matchedHeaders: [] };
    let score = 0;
    const matchedHeaders = [];
    const normCols = rowArray.map(c => normalizeKey(c));

    if (normCols.some(c => c.includes('coursedetail') || c.includes('course') || c.includes('subject'))) {
      score += 10;
      matchedHeaders.push('Course Details');
    }
    if (normCols.some(c => c.includes('programterm') || c.includes('term') || c.includes('semester'))) {
      score += 6;
      matchedHeaders.push('Program Term');
    }
    if (normCols.some(c => c.includes('collegecode') || c.includes('adeccode') || c.includes('instcode') || c.includes('centercode') || c === 'code')) {
      score += 4;
      matchedHeaders.push('College Code');
    }
    if (normCols.some(c => c.includes('collegename') || c.includes('adecname') || c.includes('adec') || c.includes('institutename') || c === 'college' || c === 'centername')) {
      score += 4;
      matchedHeaders.push('College Name');
    }
    if (normCols.some(c => c.includes('programcode') || c.includes('progcode') || c.includes('degreecode') || c === 'program')) {
      score += 4;
      matchedHeaders.push('Program Code');
    }

    return { score, matchedHeaders };
  };

  const parseSheetWithHeaderScan = (ws) => {
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (!aoa || aoa.length === 0) return { rows: [], headerMap: {}, headerRowIdx: 0, score: 0, matchedHeaders: [] };

    // Scan top 15 rows to find the best header row
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
        rowObj[name] = rowArr[idx] !== undefined ? String(rowArr[idx]).trim() : '';
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

        // Analyze all sheets to calculate header match scores
        const meta = {};
        let bestSheet = wb.SheetNames[0];
        let highestTotalScore = -1;

        wb.SheetNames.forEach((sName) => {
          const ws = wb.Sheets[sName];
          const { score, matchedHeaders, headerRowIdx, rows } = parseSheetWithHeaderScan(ws);
          
          let nameBonus = 0;
          const sNorm = normalizeKey(sName);
          if (sNorm === 'affiliatedprograms' || sNorm.includes('affiliatedprogram')) nameBonus += 12;
          else if (sNorm.includes('affiliated')) nameBonus += 6;
          else if (sNorm.includes('program')) nameBonus += 4;
          else if (sNorm === 'source') nameBonus += 5;

          const totalScore = score + nameBonus + (rows.length > 0 ? 3 : 0);
          meta[sName] = { score: totalScore, matchedHeaders, headerRowIdx, rowCount: rows.length };

          if (totalScore > highestTotalScore) {
            highestTotalScore = totalScore;
            bestSheet = sName;
          }
        });

        setSheetMetadata(meta);
        setSelectedSheet(bestSheet);

        // Load best detected sheet
        const ws = wb.Sheets[bestSheet];
        const { rows, headerMap: hMap, headerRowIdx, matchedHeaders } = parseSheetWithHeaderScan(ws);

        if (!rows || rows.length === 0) {
          setStatus(`No valid data rows found in sheet "${bestSheet}".`, 'warning');
          setIsProcessing(false);
          return;
        }

        setHeaderMap(hMap);
        setRawRows(rows);

        const { allExploded, dedupeList, dupCount } = processDataFromRows(rows, hMap);
        setAllExplodedRows(allExploded);
        setDeduplicatedRows(dedupeList);
        setPage(0);
        
        const headerInfo = matchedHeaders.length > 0 ? ` (Detected Headers: ${matchedHeaders.join(', ')} at Row ${headerRowIdx + 1})` : '';
        setStatus(`Loaded "${bestSheet}"${headerInfo}: ${rows.length} source rows -> ${dedupeList.length} unique records (${dupCount} repeats filtered), ${allExploded.length} total exploded rows.`, 'success');
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
        setAllExplodedRows([]);
        setDeduplicatedRows([]);
        setIsProcessing(false);
        return;
      }

      setHeaderMap(hMap);
      setRawRows(rows);

      const { allExploded, dedupeList, dupCount } = processDataFromRows(rows, hMap);
      setAllExplodedRows(allExploded);
      setDeduplicatedRows(dedupeList);
      setPage(0);

      const headerInfo = matchedHeaders.length > 0 ? ` (Headers: ${matchedHeaders.join(', ')} on Row ${headerRowIdx + 1})` : '';
      setStatus(`Loaded "${sheetName}"${headerInfo}: ${rows.length} source rows -> ${dedupeList.length} unique (${dupCount} filtered), ${allExploded.length} all rows.`, 'success');
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
    setSelectedCollegeFilter('ALL');
    setSelectedProgramFilter('ALL');
    setSelectedYearFilter('ALL');
    setSelectedSemesterFilter('ALL');
    setSearchQuery('');
    setPage(0);
  };

  // Determine active dataset and columns based on selected tab
  const activeHeaders = activeTab === 'deduplicated' ? DEDUPLICATED_HEADERS : ALL_ROWS_HEADERS;
  const currentDataset = activeTab === 'deduplicated' ? deduplicatedRows : allExplodedRows;

  // Unique Lists for Dropdown Filters
  const uniqueColleges = useMemo(() => {
    const set = new Set();
    currentDataset.forEach(r => {
      if (r['College Name']) set.add(r['College Name']);
      else if (r['College Code']) set.add(r['College Code']);
    });
    return Array.from(set).sort();
  }, [currentDataset]);

  const uniquePrograms = useMemo(() => {
    const set = new Set();
    allExplodedRows.forEach(r => {
      if (r['Program Code']) set.add(r['Program Code']);
    });
    return Array.from(set).sort();
  }, [allExplodedRows]);

  const uniqueYears = useMemo(() => {
    const set = new Set();
    currentDataset.forEach(r => {
      if (r['Programme Year']) set.add(r['Programme Year']);
    });
    return Array.from(set).sort();
  }, [currentDataset]);

  const uniqueSemesters = useMemo(() => {
    const set = new Set();
    currentDataset.forEach(r => {
      if (r['Program Term Name']) set.add(r['Program Term Name']);
    });
    return Array.from(set).sort();
  }, [currentDataset]);

  // Filtered & Sorted Rows for the active tab
  const filteredRows = useMemo(() => {
    let result = [...currentDataset];

    // 1. Dropdown Filters
    if (selectedCollegeFilter !== 'ALL') {
      result = result.filter(r => r['College Name'] === selectedCollegeFilter || r['College Code'] === selectedCollegeFilter);
    }
    if (activeTab === 'all_rows' && selectedProgramFilter !== 'ALL') {
      result = result.filter(r => r['Program Code'] === selectedProgramFilter);
    }
    if (selectedYearFilter !== 'ALL') {
      result = result.filter(r => r['Programme Year'] === selectedYearFilter);
    }
    if (selectedSemesterFilter !== 'ALL') {
      result = result.filter(r => r['Program Term Name'] === selectedSemesterFilter);
    }

    // 2. Global Search Query
    if (searchQuery) {
      const q = searchQuery.toLowerCase().trim();
      result = result.filter(r => 
        Object.values(r).some(val => String(val || '').toLowerCase().includes(q))
      );
    }

    // 3. Column-specific filters
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

    // 4. Sorting
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
  }, [currentDataset, activeTab, selectedCollegeFilter, selectedProgramFilter, selectedYearFilter, selectedSemesterFilter, searchQuery, columnFilters, sortConfig]);

  const pagedRows = useMemo(() => {
    const start = page * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, page]);

  // Export a specific dataset to single-sheet Excel
  const handleExportSingleReport = (rowsToExport, headers, sheetTitle, filename) => {
    if (!rowsToExport || rowsToExport.length === 0) {
      alert('No rows to export.');
      return;
    }

    setIsProcessing(true);
    setStatus(`Generating ${sheetTitle} export...`, 'info');

    try {
      const aoa = [
        headers,
        ...rowsToExport.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''))
      ];

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });

      ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: aoa.length - 1, c: headers.length - 1 }
        })
      };

      ws['!cols'] = headers.map(h => ({
        wch: Math.max(h.length + 3, 15)
      }));

      XLSX.utils.book_append_sheet(wb, ws, sheetTitle);
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

      setStatus(`Successfully exported ${rowsToExport.length} rows to ${filename}!`, 'success');
    } catch (err) {
      console.error('Export Error:', err);
      setStatus(`Export failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Export Combined Multi-Tab Workbook (Sheet 1: Deduplicated 7-Col, Sheet 2: All Rows 9-Col)
  const handleExportCombinedExcel = () => {
    if (deduplicatedRows.length === 0 && allExplodedRows.length === 0) {
      alert('No data to export.');
      return;
    }

    setIsProcessing(true);
    setStatus('Generating 2-Sheet Combined Excel Workbook...', 'info');

    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Deduplicated Unique Report (7 Columns)
      const aoaDedupe = [
        DEDUPLICATED_HEADERS,
        ...deduplicatedRows.map(r => DEDUPLICATED_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))
      ];
      const wsDedupe = XLSX.utils.aoa_to_sheet(aoaDedupe, { dense: true });
      wsDedupe['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: aoaDedupe.length - 1, c: DEDUPLICATED_HEADERS.length - 1 }
        })
      };
      wsDedupe['!cols'] = DEDUPLICATED_HEADERS.map(h => ({ wch: Math.max(h.length + 3, 15) }));
      XLSX.utils.book_append_sheet(wb, wsDedupe, 'Unique_Deduplicated_Courses');

      // Sheet 2: All Rows Exploded Report (9 Columns)
      const aoaAll = [
        ALL_ROWS_HEADERS,
        ...allExplodedRows.map(r => ALL_ROWS_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))
      ];
      const wsAll = XLSX.utils.aoa_to_sheet(aoaAll, { dense: true });
      wsAll['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: aoaAll.length - 1, c: ALL_ROWS_HEADERS.length - 1 }
        })
      };
      wsAll['!cols'] = ALL_ROWS_HEADERS.map(h => ({ wch: Math.max(h.length + 3, 15) }));
      XLSX.utils.book_append_sheet(wb, wsAll, 'All_Exploded_Rows');

      const filename = 'Affiliated_Programs_Combined_Report.xlsx';
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

      setStatus(`Successfully exported combined 2-sheet workbook with ${deduplicatedRows.length} unique and ${allExplodedRows.length} all rows!`, 'success');
    } catch (err) {
      console.error('Combined Export Error:', err);
      setStatus(`Export failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

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
            <Building2 size={20} color="var(--accent)" />
            <h2 style={{ fontSize: '16px', fontWeight: 700, margin: 0 }}>Affiliated Programme Details</h2>
            <span style={{ fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '12px', fontWeight: 600 }}>
              Dual Report Formatter
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

          {(deduplicatedRows.length > 0 || allExplodedRows.length > 0) && (
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              {/* Export Active Tab Single Sheet */}
              <button 
                type="button" 
                onClick={() => {
                  if (activeTab === 'deduplicated') {
                    handleExportSingleReport(filteredRows, DEDUPLICATED_HEADERS, 'Unique_Deduplicated_Courses', 'Affiliated_Programs_Deduplicated_Report.xlsx');
                  } else {
                    handleExportSingleReport(filteredRows, ALL_ROWS_HEADERS, 'All_Exploded_Rows', 'Affiliated_Programs_All_Rows_Report.xlsx');
                  }
                }}
                style={{ 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '6px', 
                  padding: '6px 14px', 
                  fontSize: '12px', 
                  background: 'var(--accent)', 
                  color: 'white', 
                  border: 'none', 
                  borderRadius: '6px', 
                  fontWeight: 600,
                  cursor: 'pointer' 
                }}
                title={`Export ${activeTab === 'deduplicated' ? '7-Column Deduplicated' : '9-Column All Rows'} report to Excel`}
              >
                <Download size={14} /> Export {activeTab === 'deduplicated' ? 'Deduplicated (7 Cols)' : 'All Rows (9 Cols)'} ({filteredRows.length})
              </button>

              {/* Export Combined 2-Sheet Excel */}
              <button 
                type="button" 
                className="secondary"
                onClick={handleExportCombinedExcel}
                style={{ 
                  display: 'flex', 
                  alignItems: 'center', 
                  gap: '6px', 
                  padding: '6px 12px', 
                  fontSize: '12px', 
                  fontWeight: 600,
                  cursor: 'pointer' 
                }}
                title="Download single Excel workbook containing both Deduplicated & All Rows sheets"
              >
                <FileSpreadsheet size={14} color="var(--accent)" /> Export Combined (2-Sheet)
              </button>
            </div>
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
              {sourceFile ? sourceFile : 'Upload Affiliated Programs Excel'}
            </div>
            <div style={{ fontSize: '11px', color: 'var(--muted)', marginTop: '4px' }}>
              Drop .xlsx / .xls file here (e.g. Affiliated_Programs_Sheet.xlsx)
            </div>
          </div>

          {/* Sheet Selector (if multiple sheets exist) */}
          {sheetNames.length > 1 && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <label style={{ fontSize: '11.5px', fontWeight: 600, color: 'var(--muted)' }}>Select Sheet:</label>
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
                      {s} {hasHeaders ? `(✓ ${meta.matchedHeaders.length} headers)` : ''} {s.toLowerCase().includes('affiliated') ? '★' : ''}
                    </option>
                  );
                })}
              </select>
            </div>
          )}

          {/* Summary Stats Card */}
          {rawRows.length > 0 && (
            <div style={{ background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px', padding: '12px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <div style={{ fontSize: '12px', fontWeight: 700, color: 'var(--ink)', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <Sparkles size={14} color="var(--accent)" /> Extraction Overview
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px', fontSize: '11px' }}>
                <div style={{ background: 'var(--panel)', padding: '8px', borderRadius: '6px', border: '1px solid var(--line)' }}>
                  <div style={{ color: 'var(--muted)' }}>Raw Source Rows</div>
                  <strong style={{ fontSize: '15px', color: 'var(--ink)' }}>{rawRows.length}</strong>
                </div>
                <div style={{ background: activeTab === 'deduplicated' ? 'var(--accent-soft)' : 'var(--panel)', padding: '8px', borderRadius: '6px', border: activeTab === 'deduplicated' ? '1px solid var(--accent)' : '1px solid var(--line)' }}>
                  <div style={{ color: activeTab === 'deduplicated' ? 'var(--accent)' : 'var(--muted)', fontWeight: 600 }}>Deduplicated (7 Col)</div>
                  <strong style={{ fontSize: '15px', color: activeTab === 'deduplicated' ? 'var(--accent)' : 'var(--ink)' }}>{deduplicatedRows.length}</strong>
                </div>
                <div style={{ background: activeTab === 'all_rows' ? 'var(--accent-soft)' : 'var(--panel)', padding: '8px', borderRadius: '6px', border: activeTab === 'all_rows' ? '1px solid var(--accent)' : '1px solid var(--line)' }}>
                  <div style={{ color: activeTab === 'all_rows' ? 'var(--accent)' : 'var(--muted)', fontWeight: 600 }}>All Exploded (9 Col)</div>
                  <strong style={{ fontSize: '15px', color: activeTab === 'all_rows' ? 'var(--accent)' : 'var(--ink)' }}>{allExplodedRows.length}</strong>
                </div>
                <div style={{ background: 'var(--panel)', padding: '8px', borderRadius: '6px', border: '1px solid var(--line)' }}>
                  <div style={{ color: 'var(--muted)' }}>Duplicates Filtered</div>
                  <strong style={{ fontSize: '15px', color: 'var(--danger, #e11d48)' }}>{duplicatesCount}</strong>
                </div>
                <div style={{ background: 'var(--panel)', padding: '8px', borderRadius: '6px', border: '1px solid var(--line)' }}>
                  <div style={{ color: 'var(--muted)' }}>Colleges</div>
                  <strong style={{ fontSize: '13px', color: 'var(--ink)' }}>{uniqueColleges.length}</strong>
                </div>
                <div style={{ background: 'var(--panel)', padding: '8px', borderRadius: '6px', border: '1px solid var(--line)' }}>
                  <div style={{ color: 'var(--muted)' }}>Programs</div>
                  <strong style={{ fontSize: '13px', color: 'var(--ink)' }}>{uniquePrograms.length}</strong>
                </div>
              </div>
            </div>
          )}

          {/* Formats Definition Card */}
          <div style={{ background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px', padding: '12px', fontSize: '11.5px', color: 'var(--muted)', lineHeight: '1.4' }}>
            <strong style={{ color: 'var(--ink)', display: 'block', marginBottom: '6px' }}>2 Output Report Formats:</strong>
            
            <div style={{ marginBottom: '8px', padding: '6px 8px', background: 'var(--panel)', borderRadius: '6px', border: '1px solid var(--line)' }}>
              <div style={{ fontWeight: 700, color: 'var(--accent)', marginBottom: '2px' }}>🛡️ Format 1: Deduplicated Report (7 Cols)</div>
              <div style={{ fontSize: '10.5px' }}>
                <code>College Code</code>, <code>College Name</code>, <code>Programme Year</code>, <code>Program Term Name</code>, <code>Course Details</code>, <code>Course Code</code>, <code>Course Name</code>
                <div style={{ marginTop: '2px', color: 'var(--ink)' }}>• Unique courses per college, program, year & term.</div>
              </div>
            </div>

            <div style={{ padding: '6px 8px', background: 'var(--panel)', borderRadius: '6px', border: '1px solid var(--line)' }}>
              <div style={{ fontWeight: 700, color: 'var(--ink)', marginBottom: '2px' }}>📄 Format 2: All Rows Exploded (9 Cols)</div>
              <div style={{ fontSize: '10.5px' }}>
                <code>College Code</code>, <code>College Name</code>, <code>Program Code</code>, <code>Program Term</code>, <code>Programme Year</code>, <code>Program Term Name</code>, <code>Course Details</code>, <code>Course Code</code>, <code>Course Name</code>
                <div style={{ marginTop: '2px', color: 'var(--ink)' }}>• Retains all exploded rows without deduplication.</div>
              </div>
            </div>
          </div>

          {/* Status Message */}
          <div style={{ marginTop: 'auto', padding: '8px 12px', borderRadius: '6px', fontSize: '11.5px', background: statusType === 'error' ? 'var(--danger-soft)' : statusType === 'success' ? 'var(--accent-soft)' : 'var(--bg)', color: statusType === 'error' ? 'var(--danger)' : statusType === 'success' ? 'var(--accent)' : 'var(--muted)', border: '1px solid var(--line)' }}>
            {statusMsg}
          </div>

        </aside>

        {/* Right Content / Data Table View */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', background: 'var(--bg)' }}>
          
          {/* Top Report Type Tabs Switcher */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '8px 16px', background: 'var(--panel)', borderBottom: '1px solid var(--line)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <button
                type="button"
                onClick={() => { setActiveTab('deduplicated'); setPage(0); }}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px',
                  padding: '7px 14px',
                  borderRadius: '6px',
                  fontSize: '12.5px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  border: activeTab === 'deduplicated' ? '1.5px solid var(--accent)' : '1px solid var(--line)',
                  background: activeTab === 'deduplicated' ? 'var(--accent-soft)' : 'var(--bg)',
                  color: activeTab === 'deduplicated' ? 'var(--accent)' : 'var(--muted)',
                  transition: 'all 0.15s ease'
                }}
              >
                <ListFilter size={15} />
                <span>Format 1: Deduplicated Report (7 Columns)</span>
                {deduplicatedRows.length > 0 && (
                  <span style={{ fontSize: '10.5px', padding: '1px 6px', borderRadius: '10px', background: activeTab === 'deduplicated' ? 'var(--accent)' : 'var(--line)', color: activeTab === 'deduplicated' ? 'white' : 'var(--ink)' }}>
                    {deduplicatedRows.length}
                  </span>
                )}
              </button>

              <button
                type="button"
                onClick={() => { setActiveTab('all_rows'); setPage(0); }}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '8px',
                  padding: '7px 14px',
                  borderRadius: '6px',
                  fontSize: '12.5px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  border: activeTab === 'all_rows' ? '1.5px solid var(--accent)' : '1px solid var(--line)',
                  background: activeTab === 'all_rows' ? 'var(--accent-soft)' : 'var(--bg)',
                  color: activeTab === 'all_rows' ? 'var(--accent)' : 'var(--muted)',
                  transition: 'all 0.15s ease'
                }}
              >
                <Layers2 size={15} />
                <span>Format 2: All Rows Exploded (9 Columns)</span>
                {allExplodedRows.length > 0 && (
                  <span style={{ fontSize: '10.5px', padding: '1px 6px', borderRadius: '10px', background: activeTab === 'all_rows' ? 'var(--accent)' : 'var(--line)', color: activeTab === 'all_rows' ? 'white' : 'var(--ink)' }}>
                    {allExplodedRows.length}
                  </span>
                )}
              </button>
            </div>

            {duplicatesCount > 0 && activeTab === 'deduplicated' && (
              <span style={{ fontSize: '11.5px', color: 'var(--muted)', display: 'flex', alignItems: 'center', gap: '4px' }}>
                <CheckCircle2 size={13} color="var(--accent)" />
                <strong>{duplicatesCount}</strong> duplicate course instances filtered
              </span>
            )}
          </div>

          {/* Interactive Toolbar */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '10px 16px', borderBottom: '1px solid var(--line)', background: 'var(--bg)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '8px' }}>
              
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                {/* Search Box */}
                <div style={{ position: 'relative', width: '200px' }}>
                  <input 
                    type="text" 
                    placeholder="Search all columns..." 
                    value={searchQuery} 
                    onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }} 
                    style={{ width: '100%', padding: '5px 8px 5px 26px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)' }} 
                  />
                  <Search size={13} color="var(--muted)" style={{ position: 'absolute', left: '8px', top: '7px' }} />
                </div>

                {/* College Filter */}
                {uniqueColleges.length > 0 && (
                  <select 
                    value={selectedCollegeFilter} 
                    onChange={(e) => { setSelectedCollegeFilter(e.target.value); setPage(0); }}
                    style={{ padding: '5px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', maxWidth: '160px' }}
                  >
                    <option value="ALL">All Colleges ({uniqueColleges.length})</option>
                    {uniqueColleges.map(c => (
                      <option key={c} value={c}>{c}</option>
                    ))}
                  </select>
                )}

                {/* Program Filter (For All Rows Mode) */}
                {activeTab === 'all_rows' && uniquePrograms.length > 0 && (
                  <select 
                    value={selectedProgramFilter} 
                    onChange={(e) => { setSelectedProgramFilter(e.target.value); setPage(0); }}
                    style={{ padding: '5px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', maxWidth: '140px' }}
                  >
                    <option value="ALL">All Programs ({uniquePrograms.length})</option>
                    {uniquePrograms.map(p => (
                      <option key={p} value={p}>{p}</option>
                    ))}
                  </select>
                )}

                {/* Year Filter */}
                {uniqueYears.length > 0 && (
                  <select 
                    value={selectedYearFilter} 
                    onChange={(e) => { setSelectedYearFilter(e.target.value); setPage(0); }}
                    style={{ padding: '5px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', maxWidth: '120px' }}
                  >
                    <option value="ALL">All Years</option>
                    {uniqueYears.map(y => (
                      <option key={y} value={y}>{y}</option>
                    ))}
                  </select>
                )}

                {/* Semester Filter */}
                {uniqueSemesters.length > 0 && (
                  <select 
                    value={selectedSemesterFilter} 
                    onChange={(e) => { setSelectedSemesterFilter(e.target.value); setPage(0); }}
                    style={{ padding: '5px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)', maxWidth: '140px' }}
                  >
                    <option value="ALL">All Semesters</option>
                    {uniqueSemesters.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                )}

                {/* Reset Filters */}
                {(selectedCollegeFilter !== 'ALL' || selectedProgramFilter !== 'ALL' || selectedYearFilter !== 'ALL' || selectedSemesterFilter !== 'ALL' || searchQuery || Object.keys(columnFilters).length > 0 || sortConfig.column) && (
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
            {currentDataset.length === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'var(--muted)', gap: '10px' }}>
                <Building2 size={44} style={{ opacity: 0.3 }} />
                <strong style={{ fontSize: '15px', color: 'var(--ink)' }}>No Records Loaded</strong>
                <span style={{ fontSize: '12px', maxWidth: '420px', textAlign: 'center' }}>
                  Upload your <code>Affiliated_Programs_Sheet.xlsx</code> on the left to process both <strong>Format 1 (Deduplicated 7 Columns)</strong> and <strong>Format 2 (All Rows 9 Columns)</strong> reports.
                </span>
              </div>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px', whiteSpace: 'nowrap' }}>
                <thead>
                  <tr style={{ background: 'var(--panel)', position: 'sticky', top: 0, zIndex: 10 }}>
                    <th style={{ padding: '8px 10px', textAlign: 'center', borderBottom: '1.5px solid var(--line)', color: 'var(--muted)', width: '40px' }}>#</th>
                    {activeHeaders.map((col, idx) => {
                      const isSorted = sortConfig.column === col;
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
                            background: isSorted ? 'var(--accent-soft)' : 'var(--panel)',
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
                    return (
                      <tr key={actualIdx} style={{ borderBottom: '1px solid var(--line)' }}>
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: 'var(--muted)', borderRight: '1px solid var(--line)', fontWeight: 600 }}>
                          {actualIdx + 1}
                        </td>
                        {activeHeaders.map(col => (
                          <td 
                            key={col} 
                            style={{ 
                              padding: '6px 10px', 
                              borderRight: '1px solid var(--line)',
                              color: (col === 'Course Code' || col === 'Programme Year' || col === 'Program Term Name') ? 'var(--accent)' : 'var(--ink)',
                              fontWeight: (col === 'Course Code' || col === 'College Code' || col === 'Program Code') ? 600 : 400
                            }}
                          >
                            {row[col] !== undefined ? String(row[col]) : ''}
                          </td>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>

          {/* Footer Info */}
          {currentDataset.length > 0 && (
            <div style={{ padding: '8px 16px', background: 'var(--panel)', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
              <span>
                Showing {pagedRows.length} of {filteredRows.length} filtered rows | Current View: <strong>{activeTab === 'deduplicated' ? 'Format 1: Deduplicated (7 Columns)' : 'Format 2: All Exploded (9 Columns)'}</strong> (Total: {currentDataset.length})
              </span>
              <span>
                {activeTab === 'deduplicated' 
                  ? 'Format: College Code, College Name, Programme Year, Program Term Name, Course Details, Course Code, Course Name' 
                  : 'Format: College Code, College Name, Program Code, Program Term, Programme Year, Program Term Name, Course Details, Course Code, Course Name'}
              </span>
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
                <Building2 size={20} color="var(--accent)" />
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>Affiliated Programme Dual Extraction Logic</h3>
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
                <strong>1. Format 1: Deduplicated Report (7 Columns):</strong>
                <p style={{ margin: '4px 0 6px', color: 'var(--muted)' }}>
                  Filters duplicate course instances so courses are not repeated for a programme for a specific college, year, and semester term.
                </p>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                  {DEDUPLICATED_HEADERS.map((h, i) => (
                    <span key={h} style={{ background: 'var(--accent-soft)', color: 'var(--accent)', border: '1px solid var(--accent)', padding: '3px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 600 }}>
                      {i + 1}. {h}
                    </span>
                  ))}
                </div>
              </div>

              <div>
                <strong>2. Format 2: All Rows Exploded Report (9 Columns):</strong>
                <p style={{ margin: '4px 0 6px', color: 'var(--muted)' }}>
                  Complete exploded list retaining Program Code and raw Program Term columns across all exploded items.
                </p>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px' }}>
                  {ALL_ROWS_HEADERS.map((h, i) => (
                    <span key={h} style={{ background: 'var(--bg)', border: '1px solid var(--line)', padding: '3px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 600 }}>
                      {i + 1}. {h}
                    </span>
                  ))}
                </div>
              </div>

              <div>
                <strong>3. Regular Expression Transformations:</strong>
                <ul style={{ margin: '4px 0 0', paddingLeft: '16px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <li><strong>Course Exploder:</strong> <code>r',\s*(?=\()'</code> (Splits multi-course cells by comma before opening bracket).</li>
                  <li><strong>Programme Year:</strong> <code>r'(Year\s+[IVXLCDM]+)'</code></li>
                  <li><strong>Program Term Name:</strong> <code>r'(SEMESTER\s+[IVXLCDM]+)'</code></li>
                  <li><strong>Course Code & Name:</strong> <code>r'^\(([^)]+)\)\s*(.*)$'</code></li>
                </ul>
              </div>

              <div>
                <strong>4. Combined Export:</strong>
                <p style={{ margin: '4px 0', color: 'var(--muted)' }}>
                  Click <strong>"Export Combined (2-Sheet)"</strong> to download an Excel workbook containing both <code>Unique_Deduplicated_Courses</code> and <code>All_Exploded_Rows</code> sheets in one file.
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
