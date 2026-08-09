import { useState, useMemo, useCallback } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import JSZip from 'jszip';
import { 
  GitCompare, 
  Upload, 
  Settings2, 
  BarChart3, 
  Download, 
  Search, 
  CheckCircle2, 
  AlertTriangle, 
  XCircle, 
  Layers, 
  ArrowRightLeft, 
  Eye, 
  Sparkles, 
  RotateCcw, 
  FileSpreadsheet, 
  FileArchive, 
  Plus, 
  Trash2, 
  Percent
} from 'lucide-react';
import { normalizeText, stringSimilarity, calculateCompositeSimilarity } from '../utils/fuzzyMatch';

const SAMPLE_DATASET_A = [
  { "PRN": "KU2025001", "StudentName": "Muhammed Rashid K", "CourseCode": "ENG101", "CourseTitle": "English Literature", "College": "Government College Kasaragod", "Marks": "85" },
  { "PRN": "KU2025002", "StudentName": "Ananya S Nair", "CourseCode": "MAL102", "CourseTitle": "Malayalam Poetry", "College": "Sree Narayana Guru College", "Marks": "78" },
  { "PRN": "KU2025003", "StudentName": "Fathima Hameed", "CourseCode": "ARB103", "CourseTitle": "Arabic Grammar", "College": "Sir Syed College Taliparamba", "Marks": "92" },
  { "PRN": "KU2025004", "StudentName": "Abhijith T", "CourseCode": "CS104", "CourseTitle": "Data Structures", "College": "Government Brennen College", "Marks": "95" },
  { "PRN": "KU2025005", "StudentName": "Devika Menon", "CourseCode": "PHY105", "CourseTitle": "Modern Physics", "College": "Payyanur College", "Marks": "64" },
  { "PRN": "KU2025006", "StudentName": "Rahul K V", "CourseCode": "CHE106", "CourseTitle": "Organic Chemistry", "College": "Nehru Arts & Science College", "Marks": "88" },
  { "PRN": "KU2025007", "StudentName": "Sruthi Radhakrishnan", "CourseCode": "MAT107", "CourseTitle": "Calculus & Matrices", "College": "Krishna Menon Memorial College", "Marks": "72" },
  { "PRN": "KU2025008", "StudentName": "Gokul Prasad", "CourseCode": "HIS108", "CourseTitle": "World History", "College": "SES College Sreekandapuram", "Marks": "60" }
];

const SAMPLE_DATASET_B = [
  { "RegisterNo": "KU2025001", "CandidateName": "Mohammed Rashid K.", "SubjectCode": "ENG101", "SubjectTitle": "English Literature", "ExamCentre": "Govt College Kasaragod", "Score": "85" },
  { "RegisterNo": "KU2025002", "CandidateName": "Ananya S Nair", "SubjectCode": "MAL102", "SubjectTitle": "Malayalam Poetry", "ExamCentre": "Sree Narayana Guru College", "Score": "78" },
  { "RegisterNo": "KU2025003", "CandidateName": "Fathima Hameed M", "SubjectCode": "ARB103", "SubjectTitle": "Arabic Grammar", "ExamCentre": "Sir Syed College", "Score": "90" }, // Score discrepancy
  { "RegisterNo": "KU2025004", "CandidateName": "Abhijith T", "SubjectCode": "CS104", "SubjectTitle": "Data Structures & Algorithms", "ExamCentre": "Govt Brennen College Dharmadam", "Score": "95" },
  { "RegisterNo": "KU2025005", "CandidateName": "Devika M.", "SubjectCode": "PHY105", "SubjectTitle": "Modern Physics", "ExamCentre": "Payyanur College", "Score": "64" },
  { "RegisterNo": "KU2025006", "CandidateName": "Rahul K V", "SubjectCode": "CHE106", "SubjectTitle": "Organic Chemistry", "ExamCentre": "Nehru Arts and Science College", "Score": "88" },
  { "RegisterNo": "KU2025009", "CandidateName": "Vishnu Mohan", "SubjectCode": "COM109", "SubjectTitle": "Financial Accounting", "ExamCentre": "PRNSS College Mattannur", "Score": "81" }, // Unmatched Right
  { "RegisterNo": "KU2025010", "CandidateName": "Aiswarya K", "SubjectCode": "ZOO110", "SubjectTitle": "Animal Diversity", "ExamCentre": "St. Pius X College Rajapuram", "Score": "89" } // Unmatched Right
];

const DataComparisonPage = () => {
  // Navigation Stepper: 1: Upload, 2: Rules & Mapping, 3: Analytics & Results, 4: Export
  const [currentStep, setCurrentStep] = useState(1);

  // Dataset A (Left / Reference)
  const [datasetA, setDatasetA] = useState({
    name: 'Dataset_A',
    columns: [],
    rows: [],
    fileName: '',
    sheets: [],
    selectedSheet: '',
    rawWorkbook: null
  });

  // Dataset B (Right / Comparison)
  const [datasetB, setDatasetB] = useState({
    name: 'Dataset_B',
    columns: [],
    rows: [],
    fileName: '',
    sheets: [],
    selectedSheet: '',
    rawWorkbook: null
  });

  // Key Column Mappings (Composite Keys support)
  const [keyMappings, setKeyMappings] = useState([
    { id: 1, leftCol: '', rightCol: '' }
  ]);

  // Non-key columns to compare for attribute drift/discrepancies
  const [valueCompareMappings, setValueCompareMappings] = useState([]);

  // Matching Rules & Parameters
  const [matchMode, setMatchMode] = useState('fuzzy'); // 'exact' | 'fuzzy'
  const [fuzzyThreshold, setFuzzyThreshold] = useState(80); // 50 - 100 (%)
  const [comparisonType, setComparisonType] = useState('one_to_one'); // 'one_to_one' | 'one_to_many' | 'many_to_many'
  
  // Normalization Options
  const [normOptions, setNormOptions] = useState({
    ignoreCase: true,
    stripSpaces: true,
    stripPunctuation: true,
    stripLeadingZeros: true
  });

  // Processing & Status
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready to upload datasets.');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Comparison Output Results
  const [comparisonResults, setComparisonResults] = useState(null);

  // UI Filter & Inspection States
  const [activeResultTab, setActiveResultTab] = useState('all'); // 'all' | 'exact' | 'partial' | 'discrepancy' | 'unmatched_a' | 'unmatched_b' | 'duplicates'
  const [searchQuery, setSearchQuery] = useState('');
  const [inspectModalItem, setInspectModalItem] = useState(null);

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  // Helper to load Sample Datasets for 1-click test
  const loadSampleData = () => {
    const colsA = Object.keys(SAMPLE_DATASET_A[0]);
    const colsB = Object.keys(SAMPLE_DATASET_B[0]);

    setDatasetA({
      name: 'Admission_Master_List.xlsx',
      columns: colsA,
      rows: SAMPLE_DATASET_A,
      fileName: 'Admission_Master_List.xlsx',
      sheets: ['Master_2025'],
      selectedSheet: 'Master_2025',
      rawWorkbook: null
    });

    setDatasetB({
      name: 'Exam_Registration_Report.xlsx',
      columns: colsB,
      rows: SAMPLE_DATASET_B,
      fileName: 'Exam_Registration_Report.xlsx',
      sheets: ['Registrations_Nov2025'],
      selectedSheet: 'Registrations_Nov2025',
      rawWorkbook: null
    });

    setKeyMappings([
      { id: 1, leftCol: 'PRN', rightCol: 'RegisterNo' }
    ]);

    setValueCompareMappings([
      { id: 101, leftCol: 'Marks', rightCol: 'Score' },
      { id: 102, leftCol: 'StudentName', rightCol: 'CandidateName' }
    ]);

    setStatus('Sample datasets loaded successfully. Proceed to Configure Rules.', 'success');
    setCurrentStep(2);
  };

  // Universal File Parser (handles .xlsx, .xls, .csv, .zip)
  const parseUploadedFile = async (file) => {
    const ext = file.name.split('.').pop().toLowerCase();
    
    if (ext === 'zip') {
      const zip = new JSZip();
      const unzipped = await zip.loadAsync(file);
      const excelFiles = Object.keys(unzipped.files).filter(fn => 
        !unzipped.files[fn].dir && (fn.endsWith('.xlsx') || fn.endsWith('.xls') || fn.endsWith('.csv'))
      );

      if (excelFiles.length === 0) {
        throw new Error('No valid Excel or CSV files found inside the ZIP archive.');
      }

      // Pick the first spreadsheet or extract all
      const targetFileName = excelFiles[0];
      const blob = await unzipped.files[targetFileName].async('blob');
      return await parseSpreadsheetBlob(blob, targetFileName);
    } else {
      return await parseSpreadsheetBlob(file, file.name);
    }
  };

  const parseSpreadsheetBlob = (fileOrBlob, originalFileName) => {
    return new Promise((resolve, reject) => {
      const ext = originalFileName.split('.').pop().toLowerCase();

      if (ext === 'csv') {
        Papa.parse(fileOrBlob, {
          header: true,
          skipEmptyLines: true,
          complete: (results) => {
            if (!results.data || results.data.length === 0) {
              reject(new Error(`CSV file ${originalFileName} is empty.`));
              return;
            }
            const columns = Object.keys(results.data[0] || {}).map(c => c.trim()).filter(Boolean);
            resolve({
              columns,
              rows: results.data,
              fileName: originalFileName,
              sheets: ['Sheet1'],
              selectedSheet: 'Sheet1',
              rawWorkbook: null
            });
          },
          error: reject
        });
      } else {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            
            if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
              reject(new Error(`Workbook ${originalFileName} has no sheets.`));
              return;
            }

            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            if (rows.length === 0) {
              reject(new Error(`Sheet ${firstSheetName} in ${originalFileName} is empty.`));
              return;
            }

            const columns = Object.keys(rows[0] || {}).map(c => c.trim()).filter(Boolean);
            resolve({
              columns,
              rows,
              fileName: originalFileName,
              sheets: workbook.SheetNames,
              selectedSheet: firstSheetName,
              rawWorkbook: workbook
            });
          } catch (err) {
            reject(err);
          }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(fileOrBlob);
      }
    });
  };

  // Handle Sheet Change for Dataset A
  const handleSheetChangeA = (sheetName) => {
    if (!datasetA.rawWorkbook || !datasetA.rawWorkbook.Sheets[sheetName]) return;
    const worksheet = datasetA.rawWorkbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    const columns = rows.length > 0 ? Object.keys(rows[0]).map(c => c.trim()).filter(Boolean) : [];
    setDatasetA(prev => ({
      ...prev,
      selectedSheet: sheetName,
      columns,
      rows
    }));
  };

  // Handle Sheet Change for Dataset B
  const handleSheetChangeB = (sheetName) => {
    if (!datasetB.rawWorkbook || !datasetB.rawWorkbook.Sheets[sheetName]) return;
    const worksheet = datasetB.rawWorkbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
    const columns = rows.length > 0 ? Object.keys(rows[0]).map(c => c.trim()).filter(Boolean) : [];
    setDatasetB(prev => ({
      ...prev,
      selectedSheet: sheetName,
      columns,
      rows
    }));
  };

  // Upload Dataset A Handler
  const handleUploadA = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setStatus(`Processing Dataset A: ${file.name}...`);

    try {
      const parsed = await parseUploadedFile(file);
      setDatasetA({
        name: file.name,
        ...parsed
      });
      setStatus(`Dataset A loaded (${parsed.rows.length} rows, ${parsed.columns.length} cols).`, 'success');
      
      // Auto-suggest initial key if Dataset B already loaded
      if (datasetB.columns.length > 0) {
        autoSuggestKeyMappings(parsed.columns, datasetB.columns);
      }
    } catch (err) {
      console.error(err);
      setStatus(`Error loading Dataset A: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Upload Dataset B Handler
  const handleUploadB = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setStatus(`Processing Dataset B: ${file.name}...`);

    try {
      const parsed = await parseUploadedFile(file);
      setDatasetB({
        name: file.name,
        ...parsed
      });
      setStatus(`Dataset B loaded (${parsed.rows.length} rows, ${parsed.columns.length} cols).`, 'success');

      // Auto-suggest initial key if Dataset A already loaded
      if (datasetA.columns.length > 0) {
        autoSuggestKeyMappings(datasetA.columns, parsed.columns);
      }
    } catch (err) {
      console.error(err);
      setStatus(`Error loading Dataset B: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Auto-Suggest Key Mappings based on name similarity
  const autoSuggestKeyMappings = (colsA, colsB) => {
    const suggestions = [];
    const usedB = new Set();

    colsA.forEach((colA) => {
      const normA = colA.toLowerCase().replace(/[^a-z0-9]/g, '');
      let bestMatch = null;
      let bestScore = 0;

      colsB.forEach((colB) => {
        if (usedB.has(colB)) return;
        const normB = colB.toLowerCase().replace(/[^a-z0-9]/g, '');

        if (normA === normB) {
          bestMatch = colB;
          bestScore = 1.0;
        } else {
          const score = stringSimilarity(normA, normB);
          if (score > 0.7 && score > bestScore) {
            bestMatch = colB;
            bestScore = score;
          }
        }
      });

      if (bestMatch && bestScore >= 0.75) {
        suggestions.push({
          id: Date.now() + Math.random(),
          leftCol: colA,
          rightCol: bestMatch
        });
        usedB.add(bestMatch);
      }
    });

    if (suggestions.length > 0) {
      setKeyMappings(suggestions.slice(0, 2)); // Use top 1-2 matches as key columns
    } else {
      setKeyMappings([{ id: 1, leftCol: colsA[0] || '', rightCol: colsB[0] || '' }]);
    }
  };

  // Key Mapping Manipulators
  const addKeyMapping = () => {
    setKeyMappings(prev => [
      ...prev,
      { id: Date.now(), leftCol: datasetA.columns[0] || '', rightCol: datasetB.columns[0] || '' }
    ]);
  };

  const updateKeyMapping = (id, side, value) => {
    setKeyMappings(prev => prev.map(m => m.id === id ? { ...m, [side]: value } : m));
  };

  const removeKeyMapping = (id) => {
    if (keyMappings.length <= 1) return;
    setKeyMappings(prev => prev.filter(m => m.id !== id));
  };

  // Value Compare Mapping Manipulators
  const addValueCompare = () => {
    setValueCompareMappings(prev => [
      ...prev,
      { id: Date.now(), leftCol: datasetA.columns[0] || '', rightCol: datasetB.columns[0] || '' }
    ]);
  };

  const updateValueCompare = (id, side, value) => {
    setValueCompareMappings(prev => prev.map(m => m.id === id ? { ...m, [side]: value } : m));
  };

  const removeValueCompare = (id) => {
    setValueCompareMappings(prev => prev.filter(m => m.id !== id));
  };

  // CORE COMPARISON & RECONCILIATION ENGINE
  const runComparisonEngine = useCallback(() => {
    if (!datasetA.rows.length || !datasetB.rows.length) {
      alert("Please ensure both Dataset A and Dataset B have data loaded.");
      return;
    }

    const activeKeys = keyMappings.filter(k => k.leftCol && k.rightCol);
    if (activeKeys.length === 0) {
      alert("Please select at least one Key Column pair for comparison.");
      return;
    }

    setIsProcessing(true);
    setStatus("Executing Data Comparison Engine...");

    setTimeout(() => {
      try {
        const rowsA = [...datasetA.rows];
        const rowsB = [...datasetB.rows];

        // Deduplication Detection
        const duplicateA = [];
        const duplicateB = [];
        const seenKeysA = new Set();
        const seenKeysB = new Set();

        const getCompositeKey = (row, cols, isLeft) => {
          return cols.map(c => normalizeText(row[isLeft ? c.leftCol : c.rightCol], normOptions)).join('||');
        };

        rowsA.forEach((r, idx) => {
          const k = getCompositeKey(r, activeKeys, true);
          if (seenKeysA.has(k)) {
            duplicateA.push({ ...r, _source: 'Dataset A', _rowIndex: idx + 1, _dupKey: k });
          } else {
            seenKeysA.add(k);
          }
        });

        rowsB.forEach((r, idx) => {
          const k = getCompositeKey(r, activeKeys, false);
          if (seenKeysB.has(k)) {
            duplicateB.push({ ...r, _source: 'Dataset B', _rowIndex: idx + 1, _dupKey: k });
          } else {
            seenKeysB.add(k);
          }
        });

        // Comparison Buckets
        const exactMatches = [];
        const partialMatches = [];
        const valueDiscrepancies = [];
        const matchedIndexB = new Set();
        const matchedIndexA = new Set();

        const thresholdRatio = fuzzyThreshold / 100.0;
        const activeValueComps = valueCompareMappings.filter(v => v.leftCol && v.rightCol);

        // PASS 1: Exact Key Matching
        const indexMapB = new Map();
        rowsB.forEach((rowB, idxB) => {
          const keyB = getCompositeKey(rowB, activeKeys, false);
          if (!indexMapB.has(keyB)) indexMapB.set(keyB, []);
          indexMapB.get(keyB).push({ row: rowB, index: idxB });
        });

        rowsA.forEach((rowA, idxA) => {
          const keyA = getCompositeKey(rowA, activeKeys, true);

          if (indexMapB.has(keyA) && indexMapB.get(keyA).length > 0) {
            const matchEntry = indexMapB.get(keyA)[0];
            const rowB = matchEntry.row;
            const idxB = matchEntry.index;

            matchedIndexA.add(idxA);
            matchedIndexB.add(idxB);

            // Check non-key value discrepancies
            const discrepancies = [];
            activeValueComps.forEach(comp => {
              const valA = normalizeText(rowA[comp.leftCol], normOptions);
              const valB = normalizeText(rowB[comp.rightCol], normOptions);
              if (valA !== valB) {
                discrepancies.push({
                  fieldA: comp.leftCol,
                  fieldB: comp.rightCol,
                  valA: rowA[comp.leftCol],
                  valB: rowB[comp.rightCol]
                });
              }
            });

            if (discrepancies.length > 0) {
              valueDiscrepancies.push({
                id: `disc_${idxA}_${idxB}`,
                status: 'Value Discrepancy',
                confidence: 100,
                key: keyA,
                rowA,
                rowB,
                discrepancies,
                matchType: 'Exact Key Match with Attribute Differences'
              });
            } else {
              exactMatches.push({
                id: `exact_${idxA}_${idxB}`,
                status: 'Exact Match',
                confidence: 100,
                key: keyA,
                rowA,
                rowB,
                discrepancies: [],
                matchType: '100% Exact Match'
              });
            }
          }
        });

        // PASS 2: Fuzzy Matching for Unmatched rows
        if (matchMode === 'fuzzy') {
          rowsA.forEach((rowA, idxA) => {
            if (matchedIndexA.has(idxA)) return;

            let bestMatchB = null;
            let bestScore = 0;
            let bestIdxB = -1;

            rowsB.forEach((rowB, idxB) => {
              if (matchedIndexB.has(idxB)) return;

              const score = calculateCompositeSimilarity(rowA, rowB, activeKeys, normOptions);
              if (score >= thresholdRatio && score > bestScore) {
                bestScore = score;
                bestMatchB = rowB;
                bestIdxB = idxB;
              }
            });

            if (bestMatchB && bestIdxB !== -1) {
              matchedIndexA.add(idxA);
              matchedIndexB.add(bestIdxB);

              const discrepancies = [];
              activeValueComps.forEach(comp => {
                const valA = normalizeText(rowA[comp.leftCol], normOptions);
                const valB = normalizeText(bestMatchB[comp.rightCol], normOptions);
                if (valA !== valB) {
                  discrepancies.push({
                    fieldA: comp.leftCol,
                    fieldB: comp.rightCol,
                    valA: rowA[comp.leftCol],
                    valB: bestMatchB[comp.rightCol]
                  });
                }
              });

              partialMatches.push({
                id: `fuzzy_${idxA}_${bestIdxB}`,
                status: 'Partial Match',
                confidence: Math.round(bestScore * 100),
                key: getCompositeKey(rowA, activeKeys, true),
                rowA,
                rowB: bestMatchB,
                discrepancies,
                matchType: `Fuzzy Similarity (${Math.round(bestScore * 100)}%)`
              });
            }
          });
        }

        // PASS 3: Unmatched Left and Unmatched Right
        const unmatchedA = [];
        rowsA.forEach((rowA, idxA) => {
          if (!matchedIndexA.has(idxA)) {
            unmatchedA.push({
              id: `un_a_${idxA}`,
              status: 'Unmatched (Dataset A Only)',
              confidence: 0,
              rowA,
              rowB: null,
              key: getCompositeKey(rowA, activeKeys, true)
            });
          }
        });

        const unmatchedB = [];
        rowsB.forEach((rowB, idxB) => {
          if (!matchedIndexB.has(idxB)) {
            unmatchedB.push({
              id: `un_b_${idxB}`,
              status: 'Unmatched (Dataset B Only)',
              confidence: 0,
              rowA: null,
              rowB,
              key: getCompositeKey(rowB, activeKeys, false)
            });
          }
        });

        // Column Discrepancy Statistics
        const columnDiscrepancyCounts = {};
        activeValueComps.forEach(comp => {
          const label = `${comp.leftCol} ↔ ${comp.rightCol}`;
          columnDiscrepancyCounts[label] = 0;
        });

        [...valueDiscrepancies, ...partialMatches].forEach(item => {
          item.discrepancies.forEach(d => {
            const label = `${d.fieldA} ↔ ${d.fieldB}`;
            if (columnDiscrepancyCounts[label] !== undefined) {
              columnDiscrepancyCounts[label] += 1;
            }
          });
        });

        const totalMatchedRecords = exactMatches.length + partialMatches.length + valueDiscrepancies.length;
        const totalBaseRecords = Math.max(rowsA.length, rowsB.length);
        const matchPercentage = totalBaseRecords > 0 ? ((totalMatchedRecords / totalBaseRecords) * 100).toFixed(1) : '0';

        const results = {
          totalRowsA: rowsA.length,
          totalRowsB: rowsB.length,
          totalProcessed: rowsA.length + rowsB.length,
          exactMatches,
          partialMatches,
          valueDiscrepancies,
          unmatchedA,
          unmatchedB,
          duplicatesA: duplicateA,
          duplicatesB: duplicateB,
          totalDuplicates: duplicateA.length + duplicateB.length,
          totalMatchedRecords,
          matchPercentage,
          columnDiscrepancyCounts,
          comparedKeys: activeKeys,
          comparedValues: activeValueComps,
          timestamp: new Date().toLocaleString()
        };

        setComparisonResults(results);
        setCurrentStep(3);
        setStatus(`Reconciliation Complete! ${totalMatchedRecords} records matched (${matchPercentage}% Match Rate).`, 'success');
      } catch (err) {
        console.error("Comparison Engine error:", err);
        setStatus(`Comparison error: ${err.message}`, 'error');
      } finally {
        setIsProcessing(false);
      }
    }, 250);
  }, [datasetA, datasetB, keyMappings, valueCompareMappings, matchMode, fuzzyThreshold, normOptions]);

  // Filtered rows for results table
  const filteredResultItems = useMemo(() => {
    if (!comparisonResults) return [];

    let list = [];
    if (activeResultTab === 'all') {
      list = [
        ...comparisonResults.exactMatches,
        ...comparisonResults.partialMatches,
        ...comparisonResults.valueDiscrepancies,
        ...comparisonResults.unmatchedA,
        ...comparisonResults.unmatchedB
      ];
    } else if (activeResultTab === 'exact') {
      list = comparisonResults.exactMatches;
    } else if (activeResultTab === 'partial') {
      list = comparisonResults.partialMatches;
    } else if (activeResultTab === 'discrepancy') {
      list = comparisonResults.valueDiscrepancies;
    } else if (activeResultTab === 'unmatched_a') {
      list = comparisonResults.unmatchedA;
    } else if (activeResultTab === 'unmatched_b') {
      list = comparisonResults.unmatchedB;
    } else if (activeResultTab === 'duplicates') {
      list = [
        ...comparisonResults.duplicatesA.map(d => ({ ...d, status: 'Duplicate in Dataset A' })),
        ...comparisonResults.duplicatesB.map(d => ({ ...d, status: 'Duplicate in Dataset B' }))
      ];
    }

    if (!searchQuery.trim()) return list;

    const q = searchQuery.toLowerCase();
    return list.filter(item => {
      const strA = item.rowA ? JSON.stringify(item.rowA).toLowerCase() : '';
      const strB = item.rowB ? JSON.stringify(item.rowB).toLowerCase() : '';
      const keyStr = item.key ? String(item.key).toLowerCase() : '';
      const statusStr = item.status ? String(item.status).toLowerCase() : '';
      return strA.includes(q) || strB.includes(q) || keyStr.includes(q) || statusStr.includes(q);
    });
  }, [comparisonResults, activeResultTab, searchQuery]);

  // Helper for tab names
  const getTabLabel = (tabId) => {
    switch(tabId) {
      case 'exact': return 'Exact Matches';
      case 'partial': return 'Partial Matches';
      case 'discrepancy': return 'Value Discrepancies';
      case 'unmatched_a': return 'Unmatched Left (A)';
      case 'unmatched_b': return 'Unmatched Right (B)';
      case 'duplicates': return 'Duplicates';
      default: return 'All Records';
    }
  };

  // EXPORT CURRENT ACTIVE TAB DIRECTLY AS EXCEL
  const exportCurrentTabExcel = () => {
    if (!comparisonResults) return;

    const wb = XLSX.utils.book_new();
    let sheetName = 'Records';
    let dataToExport = [];

    if (activeResultTab === 'all') {
      sheetName = 'All_Compared_Records';
      dataToExport = [
        ...comparisonResults.exactMatches.map(m => ({ "Match_Status": m.status, "Confidence_%": m.confidence, ...m.rowA })),
        ...comparisonResults.partialMatches.map(m => {
          const row = { "Match_Status": m.status, "Confidence_%": m.confidence, "Match_Type": m.matchType };
          Object.keys(m.rowA || {}).forEach(k => { row[`A_${k}`] = m.rowA[k]; });
          Object.keys(m.rowB || {}).forEach(k => { row[`B_${k}`] = m.rowB[k]; });
          return row;
        }),
        ...comparisonResults.valueDiscrepancies.map(m => {
          const row = { "Match_Status": m.status, "Confidence_%": m.confidence, "Discrepancies": m.discrepancies.map(d => `${d.fieldA} vs ${d.fieldB}`).join('; ') };
          Object.keys(m.rowA || {}).forEach(k => { row[`A_${k}`] = m.rowA[k]; });
          Object.keys(m.rowB || {}).forEach(k => { row[`B_${k}`] = m.rowB[k]; });
          return row;
        }),
        ...comparisonResults.unmatchedA.map(u => ({ "Match_Status": u.status, ...u.rowA })),
        ...comparisonResults.unmatchedB.map(u => ({ "Match_Status": u.status, ...u.rowB }))
      ];
    } else if (activeResultTab === 'exact') {
      sheetName = 'Exact_Matches';
      dataToExport = comparisonResults.exactMatches.map((m, idx) => ({
        "Match_ID": idx + 1,
        "Status": m.status,
        "Confidence_%": m.confidence,
        ...m.rowA
      }));
    } else if (activeResultTab === 'partial') {
      sheetName = 'Partial_Matches';
      dataToExport = comparisonResults.partialMatches.map((m, idx) => {
        const row = {
          "Match_ID": idx + 1,
          "Status": m.status,
          "Confidence_%": m.confidence,
          "Match_Explanation": m.matchType
        };
        Object.keys(m.rowA || {}).forEach(k => { row[`A_${k}`] = m.rowA[k]; });
        Object.keys(m.rowB || {}).forEach(k => { row[`B_${k}`] = m.rowB[k]; });
        return row;
      });
    } else if (activeResultTab === 'discrepancy') {
      sheetName = 'Value_Discrepancies';
      dataToExport = comparisonResults.valueDiscrepancies.map((m, idx) => {
        const row = {
          "Match_ID": idx + 1,
          "Status": m.status,
          "Confidence_%": m.confidence,
          "Discrepancies": m.discrepancies.map(d => `${d.fieldA}("${d.valA}" vs "${d.valB}")`).join('; ')
        };
        Object.keys(m.rowA || {}).forEach(k => { row[`A_${k}`] = m.rowA[k]; });
        Object.keys(m.rowB || {}).forEach(k => { row[`B_${k}`] = m.rowB[k]; });
        return row;
      });
    } else if (activeResultTab === 'unmatched_a') {
      sheetName = 'Unmatched_Left_A';
      dataToExport = comparisonResults.unmatchedA.map(u => u.rowA);
    } else if (activeResultTab === 'unmatched_b') {
      sheetName = 'Unmatched_Right_B';
      dataToExport = comparisonResults.unmatchedB.map(u => u.rowB);
    } else if (activeResultTab === 'duplicates') {
      sheetName = 'Duplicates';
      dataToExport = [
        ...comparisonResults.duplicatesA.map(d => ({ "Origin": "Dataset A", ...d })),
        ...comparisonResults.duplicatesB.map(d => ({ "Origin": "Dataset B", ...d }))
      ];
    }

    if (dataToExport.length === 0) {
      alert(`No records available to export for "${getTabLabel(activeResultTab)}".`);
      return;
    }

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    XLSX.utils.book_append_sheet(wb, ws, sheetName.slice(0, 31));
    XLSX.writeFile(wb, `${sheetName}_Report.xlsx`);
    setStatus(`Downloaded ${sheetName}_Report.xlsx successfully!`, 'success');
  };

  // EXPORT 1: Master Multi-Sheet Excel Workbook
  const exportMasterExcelWorkbook = () => {
    if (!comparisonResults) return;

    const wb = XLSX.utils.book_new();

    // 1. Summary Sheet
    const summaryData = [
      ["DATA RECONCILIATION & COMPARISON REPORT"],
      ["Generated At", comparisonResults.timestamp],
      ["Dataset A (Reference)", datasetA.name],
      ["Dataset B (Comparison)", datasetB.name],
      [""],
      ["EXECUTIVE METRICS", "COUNT", "PERCENTAGE"],
      ["Total Dataset A Records", comparisonResults.totalRowsA, "100%"],
      ["Total Dataset B Records", comparisonResults.totalRowsB, "100%"],
      ["Total Matched Records", comparisonResults.totalMatchedRecords, `${comparisonResults.matchPercentage}%`],
      ["  - Exact Matches (100%)", comparisonResults.exactMatches.length, `${((comparisonResults.exactMatches.length / comparisonResults.totalRowsA) * 100 || 0).toFixed(1)}%`],
      ["  - Partial Matches (Fuzzy)", comparisonResults.partialMatches.length, `${((comparisonResults.partialMatches.length / comparisonResults.totalRowsA) * 100 || 0).toFixed(1)}%`],
      ["  - Value Discrepancies", comparisonResults.valueDiscrepancies.length, `${((comparisonResults.valueDiscrepancies.length / comparisonResults.totalRowsA) * 100 || 0).toFixed(1)}%`],
      ["Unmatched in Dataset A Only", comparisonResults.unmatchedA.length, `${((comparisonResults.unmatchedA.length / comparisonResults.totalRowsA) * 100 || 0).toFixed(1)}%`],
      ["Unmatched in Dataset B Only", comparisonResults.unmatchedB.length, `${((comparisonResults.unmatchedB.length / comparisonResults.totalRowsB) * 100 || 0).toFixed(1)}%`],
      ["Duplicates Detected (A+B)", comparisonResults.totalDuplicates, "—"],
      [""],
      ["COLUMN-LEVEL DISCREPANCIES BREAKDOWN"],
      ["Column Mapping", "Mismatched Rows Count"]
    ];

    Object.entries(comparisonResults.columnDiscrepancyCounts).forEach(([k, v]) => {
      summaryData.push([k, v]);
    });

    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, wsSummary, "Summary_Dashboard");

    // 2. Exact Matches Sheet
    if (comparisonResults.exactMatches.length > 0) {
      const flatExact = comparisonResults.exactMatches.map((it, idx) => ({
        "Match_ID": idx + 1,
        "Status": it.status,
        "Confidence_%": it.confidence,
        ...it.rowA
      }));
      const wsExact = XLSX.utils.json_to_sheet(flatExact);
      XLSX.utils.book_append_sheet(wb, wsExact, "Exact_Matches");
    }

    // 3. Partial Matches Sheet
    if (comparisonResults.partialMatches.length > 0) {
      const flatPartial = comparisonResults.partialMatches.map((it, idx) => {
        const flat = {
          "Match_ID": idx + 1,
          "Status": it.status,
          "Confidence_%": it.confidence,
          "Match_Explanation": it.matchType
        };
        Object.keys(it.rowA || {}).forEach(k => { flat[`A_${k}`] = it.rowA[k]; });
        Object.keys(it.rowB || {}).forEach(k => { flat[`B_${k}`] = it.rowB[k]; });
        return flat;
      });
      const wsPartial = XLSX.utils.json_to_sheet(flatPartial);
      XLSX.utils.book_append_sheet(wb, wsPartial, "Partial_Matches");
    }

    // 4. Value Discrepancies Sheet
    if (comparisonResults.valueDiscrepancies.length > 0) {
      const flatDisc = comparisonResults.valueDiscrepancies.map((it, idx) => {
        const discSummary = it.discrepancies.map(d => `${d.fieldA}("${d.valA}" vs "${d.valB}")`).join('; ');
        const flat = {
          "Match_ID": idx + 1,
          "Status": it.status,
          "Confidence_%": it.confidence,
          "Discrepancies_Found": discSummary
        };
        Object.keys(it.rowA || {}).forEach(k => { flat[`A_${k}`] = it.rowA[k]; });
        Object.keys(it.rowB || {}).forEach(k => { flat[`B_${k}`] = it.rowB[k]; });
        return flat;
      });
      const wsDisc = XLSX.utils.json_to_sheet(flatDisc);
      XLSX.utils.book_append_sheet(wb, wsDisc, "Value_Discrepancies");
    }

    // 5. Unmatched A Sheet
    if (comparisonResults.unmatchedA.length > 0) {
      const flatUnA = comparisonResults.unmatchedA.map(it => it.rowA);
      const wsUnA = XLSX.utils.json_to_sheet(flatUnA);
      XLSX.utils.book_append_sheet(wb, wsUnA, "Unmatched_Left_A");
    }

    // 6. Unmatched B Sheet
    if (comparisonResults.unmatchedB.length > 0) {
      const flatUnB = comparisonResults.unmatchedB.map(it => it.rowB);
      const wsUnB = XLSX.utils.json_to_sheet(flatUnB);
      XLSX.utils.book_append_sheet(wb, wsUnB, "Unmatched_Right_B");
    }

    // 7. Duplicates Sheet
    if (comparisonResults.totalDuplicates > 0) {
      const dups = [
        ...comparisonResults.duplicatesA.map(d => ({ "Duplicate_Origin": "Dataset A", ...d })),
        ...comparisonResults.duplicatesB.map(d => ({ "Duplicate_Origin": "Dataset B", ...d }))
      ];
      const wsDups = XLSX.utils.json_to_sheet(dups);
      XLSX.utils.book_append_sheet(wb, wsDups, "Duplicates");
    }

    const safeA = datasetA.name.replace(/[^a-zA-Z0-9_-]/g, '_');
    const safeB = datasetB.name.replace(/[^a-zA-Z0-9_-]/g, '_');
    XLSX.writeFile(wb, `Reconciliation_Report_${safeA}_vs_${safeB}.xlsx`);
    setStatus("Master Excel Workbook downloaded successfully!", "success");
  };

  // EXPORT 2: Export All as ZIP Bundle
  const exportAllZipBundle = async () => {
    if (!comparisonResults) return;

    setStatus("Building ZIP archive package...", "normal");
    const zip = new JSZip();

    // Add individual CSVs
    const addCsvToZip = (fileName, jsonArray) => {
      if (!jsonArray || jsonArray.length === 0) return;
      const csvContent = Papa.unparse(jsonArray);
      zip.file(fileName, csvContent);
    };

    if (comparisonResults.exactMatches.length > 0) {
      addCsvToZip("01_Exact_Matches.csv", comparisonResults.exactMatches.map(m => ({ ...m.rowA, Match_Status: m.status, Confidence: m.confidence })));
    }
    if (comparisonResults.partialMatches.length > 0) {
      addCsvToZip("02_Partial_Matches.csv", comparisonResults.partialMatches.map(m => {
        const item = { Match_Status: m.status, Confidence: m.confidence, Match_Type: m.matchType };
        Object.keys(m.rowA || {}).forEach(k => { item[`A_${k}`] = m.rowA[k]; });
        Object.keys(m.rowB || {}).forEach(k => { item[`B_${k}`] = m.rowB[k]; });
        return item;
      }));
    }
    if (comparisonResults.valueDiscrepancies.length > 0) {
      addCsvToZip("03_Value_Discrepancies.csv", comparisonResults.valueDiscrepancies.map(m => {
        const item = { Match_Status: m.status, Discrepancies: m.discrepancies.map(d => `${d.fieldA} vs ${d.fieldB}`).join('; ') };
        Object.keys(m.rowA || {}).forEach(k => { item[`A_${k}`] = m.rowA[k]; });
        Object.keys(m.rowB || {}).forEach(k => { item[`B_${k}`] = m.rowB[k]; });
        return item;
      }));
    }
    if (comparisonResults.unmatchedA.length > 0) {
      addCsvToZip("04_Unmatched_Dataset_A.csv", comparisonResults.unmatchedA.map(u => u.rowA));
    }
    if (comparisonResults.unmatchedB.length > 0) {
      addCsvToZip("05_Unmatched_Dataset_B.csv", comparisonResults.unmatchedB.map(u => u.rowB));
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const downloadUrl = URL.createObjectURL(zipBlob);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = `Reconciliation_Reports_Bundle.zip`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(downloadUrl);

    setStatus("ZIP bundle downloaded successfully!", "success");
  };

  return (
    <div style={{ maxWidth: '1440px', margin: '0 auto', padding: '24px 24px 80px', fontFamily: 'var(--font-family)' }}>
      
      {/* Top Header & Navigation Bar */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px', flexWrap: 'wrap', gap: '16px' }}>
        <div>
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '8px', fontSize: '14px', marginBottom: '6px' }}>
            ← Back to Portal
          </Link>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div style={{ background: 'var(--accent)', color: 'white', padding: '8px', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <GitCompare size={24} />
            </div>
            <div>
              <h1 style={{ margin: 0, fontSize: '24px', fontWeight: 800, color: 'var(--ink)' }}>Data Comparison & Reconciliation Studio</h1>
              <p style={{ margin: 0, color: 'var(--muted)', fontSize: '14px' }}>
                Compare, match, merge, and identify discrepancies across Excel spreadsheets, CSVs, and ZIP archives.
              </p>
            </div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <button 
            className="secondary" 
            onClick={loadSampleData}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
          >
            <Sparkles size={16} color="var(--accent)" /> Load Sample Datasets
          </button>

          <div style={{
            padding: '6px 14px',
            borderRadius: '20px',
            fontSize: '13px',
            fontWeight: 600,
            background: statusType === 'error' ? 'var(--danger-soft, #fee2e2)' : statusType === 'success' ? 'var(--accent-soft)' : 'var(--panel)',
            color: statusType === 'error' ? 'var(--danger, #dc2626)' : statusType === 'success' ? 'var(--accent)' : 'var(--muted)',
            border: '1px solid var(--line)'
          }}>
            {statusMsg}
          </div>
        </div>
      </div>

      {/* 4-Step Interactive Stepper Header */}
      <div className="card" style={{ padding: '16px 20px', marginBottom: '24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
        {[
          { step: 1, title: '1. Upload Datasets', icon: Upload },
          { step: 2, title: '2. Rules & Mapping', icon: Settings2 },
          { step: 3, title: '3. Analytics & Results', icon: BarChart3 },
          { step: 4, title: '4. Export Studio', icon: Download }
        ].map((item) => {
          const Icon = item.icon;
          const isActive = currentStep === item.step;
          const isDone = currentStep > item.step;

          return (
            <button
              key={item.step}
              onClick={() => {
                if (item.step === 2 && (!datasetA.rows.length || !datasetB.rows.length)) {
                  alert("Upload both datasets in Step 1 first.");
                  return;
                }
                if (item.step >= 3 && !comparisonResults) {
                  alert("Please run comparison in Step 2 first.");
                  return;
                }
                setCurrentStep(item.step);
              }}
              style={{
                flex: 1,
                minWidth: '200px',
                padding: '10px 16px',
                borderRadius: '8px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '8px',
                fontWeight: 700,
                fontSize: '13.5px',
                background: isActive ? 'var(--accent)' : isDone ? 'var(--accent-soft)' : 'transparent',
                color: isActive ? 'white' : isDone ? 'var(--accent)' : 'var(--muted)',
                border: isActive ? 'none' : '1px solid var(--line)',
                cursor: 'pointer',
                transition: 'all 0.2s ease'
              }}
            >
              <Icon size={18} />
              {item.title}
              {isDone && <CheckCircle2 size={15} color="var(--accent)" />}
            </button>
          );
        })}
      </div>

      {/* ========================================================================= */}
      {/* STEP 1: UPLOAD DATASETS                                                   */}
      {/* ========================================================================= */}
      {currentStep === 1 && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(460px, 1fr))', gap: '24px' }}>
          
          {/* Dataset A Card */}
          <div className="card" style={{ padding: '24px', display: 'flex', flexDirection: 'column' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <div style={{ background: 'var(--accent)', color: 'white', width: '28px', height: '28px', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800 }}>A</div>
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700 }}>Dataset A (Reference / Master)</h3>
              </div>
              {datasetA.rows.length > 0 && (
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--accent)', background: 'var(--accent-soft)', padding: '4px 10px', borderRadius: '12px' }}>
                  {datasetA.rows.length} Rows • {datasetA.columns.length} Cols
                </span>
              )}
            </div>

            <p style={{ color: 'var(--muted)', fontSize: '13px', margin: '0 0 16px 0', lineHeight: '1.5' }}>
              Select the primary reference dataset (e.g. Master Admission Roster, Previous Term Data, Official Roll).
            </p>

            <div style={{ border: '2px dashed var(--line)', borderRadius: '12px', padding: '24px', textAlign: 'center', background: 'var(--bg)', marginBottom: '16px' }}>
              <Upload size={32} color="var(--accent)" style={{ margin: '0 auto 10px', opacity: 0.8 }} />
              <div style={{ fontWeight: 600, fontSize: '14px', marginBottom: '6px' }}>
                {datasetA.fileName ? datasetA.fileName : 'Upload Excel (.xlsx, .xls), CSV, or ZIP'}
              </div>
              <p style={{ color: 'var(--muted)', fontSize: '12px', margin: '0 0 14px 0' }}>Supports multi-sheet workbooks and zipped files</p>
              <label className="button" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', cursor: 'pointer', padding: '8px 18px', fontSize: '13px' }}>
                Browse File A
                <input type="file" accept=".xlsx, .xls, .csv, .zip" onChange={handleUploadA} style={{ display: 'none' }} />
              </label>
            </div>

            {datasetA.sheets.length > 1 && (
              <div style={{ marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--muted)' }}>Select Sheet:</span>
                <select 
                  value={datasetA.selectedSheet} 
                  onChange={(e) => handleSheetChangeA(e.target.value)}
                  style={{ flex: 1, padding: '6px 12px', borderRadius: '6px', fontSize: '13px' }}
                >
                  {datasetA.sheets.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
            )}

            {datasetA.columns.length > 0 && (
              <div>
                <div style={{ fontSize: '12px', fontWeight: 700, textTransform: 'uppercase', color: 'var(--muted)', marginBottom: '8px' }}>
                  Detected Columns ({datasetA.columns.length})
                </div>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', maxHeight: '100px', overflowY: 'auto', marginBottom: '16px' }}>
                  {datasetA.columns.map(c => (
                    <span key={c} style={{ fontSize: '11.5px', background: 'var(--panel)', border: '1px solid var(--line)', padding: '3px 8px', borderRadius: '6px', fontWeight: 500 }}>
                      {c}
                    </span>
                  ))}
                </div>

                <div style={{ fontSize: '12px', fontWeight: 700, textTransform: 'uppercase', color: 'var(--muted)', marginBottom: '8px' }}>
                  Data Preview (First 4 Rows)
                </div>
                <div style={{ overflowX: 'auto', border: '1px solid var(--line)', borderRadius: '6px', maxHeight: '160px' }}>
                  <table style={{ width: '100%', fontSize: '11.5px', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)' }}>
                        {datasetA.columns.slice(0, 5).map(c => <th key={c} style={{ padding: '6px 8px', textAlign: 'left' }}>{c}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {datasetA.rows.slice(0, 4).map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--line)' }}>
                          {datasetA.columns.slice(0, 5).map(c => <td key={c} style={{ padding: '6px 8px', whiteSpace: 'nowrap' }}>{String(r[c] || '')}</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>

          {/* Dataset B Card */}
          <div className="card" style={{ padding: '24px', display: 'flex', flexDirection: 'column' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <div style={{ background: 'var(--accent)', color: 'white', width: '28px', height: '28px', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800 }}>B</div>
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700 }}>Dataset B (Comparison / Incoming)</h3>
              </div>
              {datasetB.rows.length > 0 && (
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--accent)', background: 'var(--accent-soft)', padding: '4px 10px', borderRadius: '12px' }}>
                  {datasetB.rows.length} Rows • {datasetB.columns.length} Cols
                </span>
              )}
            </div>

            <p style={{ color: 'var(--muted)', fontSize: '13px', margin: '0 0 16px 0', lineHeight: '1.5' }}>
              Select the comparison dataset to reconcile against Dataset A (e.g. Exam Registrations, Fee Receipts, Attendance).
            </p>

            <div style={{ border: '2px dashed var(--line)', borderRadius: '12px', padding: '24px', textAlign: 'center', background: 'var(--bg)', marginBottom: '16px' }}>
              <Upload size={32} color="var(--accent)" style={{ margin: '0 auto 10px', opacity: 0.8 }} />
              <div style={{ fontWeight: 600, fontSize: '14px', marginBottom: '6px' }}>
                {datasetB.fileName ? datasetB.fileName : 'Upload Excel (.xlsx, .xls), CSV, or ZIP'}
              </div>
              <p style={{ color: 'var(--muted)', fontSize: '12px', margin: '0 0 14px 0' }}>Supports multi-sheet workbooks and zipped files</p>
              <label className="button" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', cursor: 'pointer', padding: '8px 18px', fontSize: '13px' }}>
                Browse File B
                <input type="file" accept=".xlsx, .xls, .csv, .zip" onChange={handleUploadB} style={{ display: 'none' }} />
              </label>
            </div>

            {datasetB.sheets.length > 1 && (
              <div style={{ marginBottom: '16px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--muted)' }}>Select Sheet:</span>
                <select 
                  value={datasetB.selectedSheet} 
                  onChange={(e) => handleSheetChangeB(e.target.value)}
                  style={{ flex: 1, padding: '6px 12px', borderRadius: '6px', fontSize: '13px' }}
                >
                  {datasetB.sheets.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
            )}

            {datasetB.columns.length > 0 && (
              <div>
                <div style={{ fontSize: '12px', fontWeight: 700, textTransform: 'uppercase', color: 'var(--muted)', marginBottom: '8px' }}>
                  Detected Columns ({datasetB.columns.length})
                </div>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', maxHeight: '100px', overflowY: 'auto', marginBottom: '16px' }}>
                  {datasetB.columns.map(c => (
                    <span key={c} style={{ fontSize: '11.5px', background: 'var(--panel)', border: '1px solid var(--line)', padding: '3px 8px', borderRadius: '6px', fontWeight: 500 }}>
                      {c}
                    </span>
                  ))}
                </div>

                <div style={{ fontSize: '12px', fontWeight: 700, textTransform: 'uppercase', color: 'var(--muted)', marginBottom: '8px' }}>
                  Data Preview (First 4 Rows)
                </div>
                <div style={{ overflowX: 'auto', border: '1px solid var(--line)', borderRadius: '6px', maxHeight: '160px' }}>
                  <table style={{ width: '100%', fontSize: '11.5px', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)' }}>
                        {datasetB.columns.slice(0, 5).map(c => <th key={c} style={{ padding: '6px 8px', textAlign: 'left' }}>{c}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {datasetB.rows.slice(0, 4).map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--line)' }}>
                          {datasetB.columns.slice(0, 5).map(c => <td key={c} style={{ padding: '6px 8px', whiteSpace: 'nowrap' }}>{String(r[c] || '')}</td>)}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>

          {/* Bottom Next Action */}
          <div style={{ gridColumn: '1 / -1', display: 'flex', justifyContent: 'flex-end', marginTop: '8px' }}>
            <button 
              disabled={datasetA.rows.length === 0 || datasetB.rows.length === 0}
              onClick={() => setCurrentStep(2)}
              style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '12px 28px', fontSize: '14.5px', fontWeight: 700 }}
            >
              Continue to Rules & Mapping →
            </button>
          </div>

        </div>
      )}

      {/* ========================================================================= */}
      {/* STEP 2: CONFIGURE RULES & MAPPING                                         */}
      {/* ========================================================================= */}
      {currentStep === 2 && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          {/* 1. Key Matching Columns (Composite Keys) */}
          <div className="card" style={{ padding: '24px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '10px' }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <ArrowRightLeft size={20} color="var(--accent)" /> Primary Matching Keys (Composite Key Pairs)
                </h3>
                <p style={{ color: 'var(--muted)', fontSize: '13px', margin: '4px 0 0 0' }}>
                  Define the unique identifier columns used to align records between Dataset A and Dataset B.
                </p>
              </div>
              <button 
                className="secondary" 
                onClick={addKeyMapping}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
              >
                <Plus size={16} /> Add Key Column Pair
              </button>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              {keyMappings.map((km, idx) => (
                <div key={km.id} style={{ display: 'flex', alignItems: 'center', gap: '16px', background: 'var(--bg)', padding: '14px 18px', borderRadius: '8px', border: '1px solid var(--line)', flexWrap: 'wrap' }}>
                  <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--accent)', minWidth: '80px' }}>
                    Key #{idx + 1}:
                  </span>

                  <div style={{ flex: 1, minWidth: '200px' }}>
                    <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, color: 'var(--muted)', marginBottom: '4px' }}>
                      DATASET A COLUMN (LEFT)
                    </label>
                    <select 
                      value={km.leftCol} 
                      onChange={(e) => updateKeyMapping(km.id, 'leftCol', e.target.value)}
                      style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', fontSize: '13.5px', fontWeight: 600 }}
                    >
                      <option value="">-- Select Column from A --</option>
                      {datasetA.columns.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>

                  <ArrowRightLeft size={18} color="var(--muted)" style={{ marginTop: '16px' }} />

                  <div style={{ flex: 1, minWidth: '200px' }}>
                    <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, color: 'var(--muted)', marginBottom: '4px' }}>
                      DATASET B COLUMN (RIGHT)
                    </label>
                    <select 
                      value={km.rightCol} 
                      onChange={(e) => updateKeyMapping(km.id, 'rightCol', e.target.value)}
                      style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', fontSize: '13.5px', fontWeight: 600 }}
                    >
                      <option value="">-- Select Column from B --</option>
                      {datasetB.columns.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>

                  {keyMappings.length > 1 && (
                    <button 
                      onClick={() => removeKeyMapping(km.id)}
                      style={{ background: 'transparent', border: 'none', color: 'var(--danger, #dc2626)', cursor: 'pointer', marginTop: '16px', padding: '6px' }}
                      title="Remove Key Pair"
                    >
                      <Trash2 size={18} />
                    </button>
                  )}
                </div>
              ))}
            </div>
          </div>

          {/* 2. Value Comparison Columns (Attribute Discrepancy Checks) */}
          <div className="card" style={{ padding: '24px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '10px' }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <Layers size={20} color="var(--accent)" /> Value Comparison Columns (Attribute Discrepancy Checks)
                </h3>
                <p style={{ color: 'var(--muted)', fontSize: '13px', margin: '4px 0 0 0' }}>
                  Compare secondary attributes (e.g. Marks, Candidate Name, Status, College) across matched records to flag differences.
                </p>
              </div>
              <button 
                className="secondary" 
                onClick={addValueCompare}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}
              >
                <Plus size={16} /> Add Comparison Field Pair
              </button>
            </div>

            {valueCompareMappings.length === 0 ? (
              <div style={{ textAlign: 'center', padding: '20px', background: 'var(--bg)', borderRadius: '8px', border: '1px dashed var(--line)', color: 'var(--muted)', fontSize: '13px' }}>
                No secondary value comparison fields added. Click <strong>"+ Add Comparison Field Pair"</strong> above if you wish to verify attributes like Marks or Names.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                {valueCompareMappings.map((vc, idx) => (
                  <div key={vc.id} style={{ display: 'flex', alignItems: 'center', gap: '16px', background: 'var(--bg)', padding: '12px 18px', borderRadius: '8px', border: '1px solid var(--line)', flexWrap: 'wrap' }}>
                    <span style={{ fontSize: '13px', fontWeight: 600, color: 'var(--muted)', minWidth: '80px' }}>
                      Field #{idx + 1}:
                    </span>

                    <div style={{ flex: 1, minWidth: '200px' }}>
                      <select 
                        value={vc.leftCol} 
                        onChange={(e) => updateValueCompare(vc.id, 'leftCol', e.target.value)}
                        style={{ width: '100%', padding: '7px 12px', borderRadius: '6px', fontSize: '13px' }}
                      >
                        <option value="">-- Column from A --</option>
                        {datasetA.columns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <ArrowRightLeft size={16} color="var(--muted)" />

                    <div style={{ flex: 1, minWidth: '200px' }}>
                      <select 
                        value={vc.rightCol} 
                        onChange={(e) => updateValueCompare(vc.id, 'rightCol', e.target.value)}
                        style={{ width: '100%', padding: '7px 12px', borderRadius: '6px', fontSize: '13px' }}
                      >
                        <option value="">-- Column from B --</option>
                        {datasetB.columns.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>

                    <button 
                      onClick={() => removeValueCompare(vc.id)}
                      style={{ background: 'transparent', border: 'none', color: 'var(--danger, #dc2626)', cursor: 'pointer', padding: '4px' }}
                      title="Remove Field Pair"
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          {/* 3. Matching Logic & Normalization Options */}
          <div className="card" style={{ padding: '24px' }}>
            <h3 style={{ margin: '0 0 16px 0', fontSize: '17px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
              <Settings2 size={20} color="var(--accent)" /> Matching Engine & Normalization Parameters
            </h3>

            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '20px', marginBottom: '24px' }}>
              
              {/* Match Mode Toggle */}
              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '8px' }}>Matching Mode:</label>
                <div style={{ display: 'flex', gap: '10px' }}>
                  <button 
                    onClick={() => setMatchMode('exact')}
                    style={{
                      flex: 1,
                      padding: '10px',
                      fontSize: '13px',
                      fontWeight: 600,
                      background: matchMode === 'exact' ? 'var(--accent)' : 'var(--bg)',
                      color: matchMode === 'exact' ? 'white' : 'var(--ink)',
                      border: '1px solid var(--line)'
                    }}
                  >
                    Exact Match Only
                  </button>
                  <button 
                    onClick={() => setMatchMode('fuzzy')}
                    style={{
                      flex: 1,
                      padding: '10px',
                      fontSize: '13px',
                      fontWeight: 600,
                      background: matchMode === 'fuzzy' ? 'var(--accent)' : 'var(--bg)',
                      color: matchMode === 'fuzzy' ? 'white' : 'var(--ink)',
                      border: '1px solid var(--line)'
                    }}
                  >
                    Fuzzy (Partial) Match
                  </button>
                </div>
              </div>

              {/* Fuzzy Similarity Threshold Slider */}
              {matchMode === 'fuzzy' && (
                <div>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                    <label style={{ fontSize: '13px', fontWeight: 700 }}>Similarity Threshold:</label>
                    <span style={{ fontSize: '14px', fontWeight: 800, color: 'var(--accent)' }}>{fuzzyThreshold}%</span>
                  </div>
                  <input 
                    type="range" 
                    min="50" 
                    max="100" 
                    value={fuzzyThreshold} 
                    onChange={(e) => setFuzzyThreshold(parseInt(e.target.value, 10))}
                    style={{ width: '100%', accentColor: 'var(--accent)', cursor: 'pointer' }}
                  />
                  <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11px', color: 'var(--muted)', marginTop: '4px' }}>
                    <span>50% (Loose)</span>
                    <span>80% (Recommended)</span>
                    <span>100% (Strict)</span>
                  </div>
                </div>
              )}

              {/* Comparison Type */}
              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '8px' }}>Comparison Cardinality:</label>
                <select 
                  value={comparisonType} 
                  onChange={(e) => setComparisonType(e.target.value)}
                  style={{ width: '100%', padding: '10px 12px', borderRadius: '6px', fontSize: '13.5px' }}
                >
                  <option value="one_to_one">One-to-One (1:1 Matching)</option>
                  <option value="one_to_many">One-to-Many (1:N Matching)</option>
                  <option value="many_to_many">Many-to-Many (N:M Matching)</option>
                </select>
              </div>

            </div>

            {/* Normalization Checkboxes */}
            <div style={{ background: 'var(--bg)', padding: '16px 20px', borderRadius: '8px', border: '1px solid var(--line)' }}>
              <div style={{ fontSize: '13px', fontWeight: 700, color: 'var(--ink)', marginBottom: '12px' }}>
                Text Cleansing & String Normalization:
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '12px' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', cursor: 'pointer' }}>
                  <input 
                    type="checkbox" 
                    checked={normOptions.ignoreCase} 
                    onChange={(e) => setNormOptions(p => ({ ...p, ignoreCase: e.target.checked }))} 
                  />
                  <span>Case-Insensitive Match</span>
                </label>

                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', cursor: 'pointer' }}>
                  <input 
                    type="checkbox" 
                    checked={normOptions.stripSpaces} 
                    onChange={(e) => setNormOptions(p => ({ ...p, stripSpaces: e.target.checked }))} 
                  />
                  <span>Trim & Collapse Extra Spaces</span>
                </label>

                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', cursor: 'pointer' }}>
                  <input 
                    type="checkbox" 
                    checked={normOptions.stripPunctuation} 
                    onChange={(e) => setNormOptions(p => ({ ...p, stripPunctuation: e.target.checked }))} 
                  />
                  <span>Ignore Punctuation & Special Chars (.,-_#)</span>
                </label>

                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '13px', cursor: 'pointer' }}>
                  <input 
                    type="checkbox" 
                    checked={normOptions.stripLeadingZeros} 
                    onChange={(e) => setNormOptions(p => ({ ...p, stripLeadingZeros: e.target.checked }))} 
                  />
                  <span>Strip Leading Zeros (e.g. 00123 → 123)</span>
                </label>
              </div>
            </div>

          </div>

          {/* Action Row */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '8px' }}>
            <button className="secondary" onClick={() => setCurrentStep(1)}>
              ← Back to Upload
            </button>
            <button 
              disabled={isProcessing}
              onClick={runComparisonEngine}
              style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '12px 32px', fontSize: '15px', fontWeight: 800 }}
            >
              {isProcessing ? 'Reconciling Records...' : '⚡ Execute Comparison Engine →'}
            </button>
          </div>

        </div>
      )}

      {/* ========================================================================= */}
      {/* STEP 3: ANALYTICS & RECONCILIATION RESULTS                                */}
      {/* ========================================================================= */}
      {currentStep === 3 && comparisonResults && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          {/* Executive Summary Metrics Deck */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px' }}>
            
            {/* Overall Match Rate */}
            <div className="card" style={{ padding: '20px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', borderLeft: '4px solid var(--accent)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>Match Rate</span>
                <Percent size={18} color="var(--accent)" />
              </div>
              <div style={{ fontSize: '28px', fontWeight: 800, color: 'var(--ink)', marginBottom: '4px' }}>
                {comparisonResults.matchPercentage}%
              </div>
              <div style={{ fontSize: '12px', color: 'var(--muted)' }}>
                {comparisonResults.totalMatchedRecords} of {Math.max(comparisonResults.totalRowsA, comparisonResults.totalRowsB)} records
              </div>
            </div>

            {/* Exact Matches */}
            <div className="card" style={{ padding: '20px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', borderLeft: '4px solid #10b981' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>Exact (100%)</span>
                <CheckCircle2 size={18} color="#10b981" />
              </div>
              <div style={{ fontSize: '28px', fontWeight: 800, color: '#10b981', marginBottom: '4px' }}>
                {comparisonResults.exactMatches.length}
              </div>
              <div style={{ fontSize: '12px', color: 'var(--muted)' }}>Identical key & attributes</div>
            </div>

            {/* Partial / Fuzzy Matches */}
            <div className="card" style={{ padding: '20px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', borderLeft: '4px solid #f59e0b' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>Fuzzy Matches</span>
                <Sparkles size={18} color="#f59e0b" />
              </div>
              <div style={{ fontSize: '28px', fontWeight: 800, color: '#f59e0b', marginBottom: '4px' }}>
                {comparisonResults.partialMatches.length}
              </div>
              <div style={{ fontSize: '12px', color: 'var(--muted)' }}>Above {fuzzyThreshold}% threshold</div>
            </div>

            {/* Value Discrepancies */}
            <div className="card" style={{ padding: '20px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', borderLeft: '4px solid #8b5cf6' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>Discrepancies</span>
                <AlertTriangle size={18} color="#8b5cf6" />
              </div>
              <div style={{ fontSize: '28px', fontWeight: 800, color: '#8b5cf6', marginBottom: '4px' }}>
                {comparisonResults.valueDiscrepancies.length}
              </div>
              <div style={{ fontSize: '12px', color: 'var(--muted)' }}>Key match, differing fields</div>
            </div>

            {/* Unmatched Left */}
            <div className="card" style={{ padding: '20px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', borderLeft: '4px solid #ef4444' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>Unmatched A</span>
                <XCircle size={18} color="#ef4444" />
              </div>
              <div style={{ fontSize: '28px', fontWeight: 800, color: '#ef4444', marginBottom: '4px' }}>
                {comparisonResults.unmatchedA.length}
              </div>
              <div style={{ fontSize: '12px', color: 'var(--muted)' }}>Only in {datasetA.name}</div>
            </div>

            {/* Unmatched Right */}
            <div className="card" style={{ padding: '20px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', borderLeft: '4px solid #ec4899' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>Unmatched B</span>
                <XCircle size={18} color="#ec4899" />
              </div>
              <div style={{ fontSize: '28px', fontWeight: 800, color: '#ec4899', marginBottom: '4px' }}>
                {comparisonResults.unmatchedB.length}
              </div>
              <div style={{ fontSize: '12px', color: 'var(--muted)' }}>Only in {datasetB.name}</div>
            </div>

          </div>

          {/* Column-Level Discrepancy Breakdown */}
          {Object.keys(comparisonResults.columnDiscrepancyCounts).length > 0 && (
            <div className="card" style={{ padding: '20px 24px', background: 'var(--panel)', border: '1px solid var(--line)' }}>
              <h4 style={{ margin: '0 0 12px 0', fontSize: '15px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px' }}>
                <BarChart3 size={18} color="var(--accent)" /> Column-Level Discrepancy Summary
              </h4>
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: '12px' }}>
                {Object.entries(comparisonResults.columnDiscrepancyCounts).map(([colPair, count]) => (
                  <div key={colPair} style={{ background: 'var(--bg)', border: '1px solid var(--line)', padding: '8px 14px', borderRadius: '8px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <span style={{ fontSize: '13px', fontWeight: 600 }}>{colPair}:</span>
                    <span style={{ 
                      padding: '2px 8px', 
                      borderRadius: '10px', 
                      fontSize: '12px', 
                      fontWeight: 800, 
                      background: count > 0 ? 'rgba(239, 68, 68, 0.15)' : 'rgba(16, 185, 129, 0.15)',
                      color: count > 0 ? '#dc2626' : '#059669'
                    }}>
                      {count} {count === 1 ? 'mismatch' : 'mismatches'}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Toolbar & Filter Tabs */}
          <div className="card" style={{ padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
            
            {/* Filter Tabs */}
            <div style={{ display: 'flex', gap: '8px', overflowX: 'auto', paddingBottom: '4px' }}>
              {[
                { id: 'all', label: 'All Records', count: comparisonResults.totalProcessed - comparisonResults.totalDuplicates },
                { id: 'exact', label: 'Exact Matches (100%)', count: comparisonResults.exactMatches.length, color: '#10b981' },
                { id: 'partial', label: 'Partial Matches (Fuzzy)', count: comparisonResults.partialMatches.length, color: '#f59e0b' },
                { id: 'discrepancy', label: 'Value Discrepancies', count: comparisonResults.valueDiscrepancies.length, color: '#8b5cf6' },
                { id: 'unmatched_a', label: 'Unmatched Left (A)', count: comparisonResults.unmatchedA.length, color: '#ef4444' },
                { id: 'unmatched_b', label: 'Unmatched Right (B)', count: comparisonResults.unmatchedB.length, color: '#ec4899' },
                { id: 'duplicates', label: 'Duplicates', count: comparisonResults.totalDuplicates, color: '#6b7280' }
              ].map(tab => (
                <button
                  key={tab.id}
                  onClick={() => setActiveResultTab(tab.id)}
                  style={{
                    padding: '8px 14px',
                    borderRadius: '20px',
                    fontSize: '13px',
                    fontWeight: 600,
                    whiteSpace: 'nowrap',
                    background: activeResultTab === tab.id ? (tab.color || 'var(--accent)') : 'var(--bg)',
                    color: activeResultTab === tab.id ? 'white' : 'var(--ink)',
                    border: '1px solid var(--line)',
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '8px',
                    cursor: 'pointer'
                  }}
                >
                  {tab.label}
                  <span style={{ 
                    background: activeResultTab === tab.id ? 'rgba(255,255,255,0.25)' : 'rgba(0,0,0,0.06)', 
                    padding: '2px 7px', 
                    borderRadius: '10px', 
                    fontSize: '11px',
                    fontWeight: 800
                  }}>
                    {tab.count}
                  </span>
                </button>
              ))}
            </div>

            {/* Search Bar & Quick Export */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
              <div style={{ position: 'relative', flex: 1, minWidth: '280px' }}>
                <Search size={16} style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
                <input 
                  type="text" 
                  placeholder="Search in results (names, keys, values)..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  style={{ width: '100%', paddingLeft: '36px', fontSize: '13px' }}
                />
              </div>

              <div style={{ display: 'flex', gap: '10px', flexWrap: 'wrap' }}>
                <button 
                  onClick={exportCurrentTabExcel}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '9px 16px', background: 'var(--accent)', color: 'white' }}
                  title="Download records in currently selected tab as Excel spreadsheet"
                >
                  <Download size={15} /> Download {getTabLabel(activeResultTab)} (.xlsx)
                </button>

                <button 
                  className="secondary"
                  onClick={exportMasterExcelWorkbook}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '9px 16px' }}
                  title="Download full multi-sheet workbook"
                >
                  <FileSpreadsheet size={15} /> Master Excel (.xlsx)
                </button>
              </div>
            </div>

          </div>

          {/* Interactive Results Table Grid */}
          <div className="card" style={{ padding: '0', overflow: 'hidden' }}>
            <div style={{ overflowX: 'auto', maxHeight: '550px' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', borderBottom: '2px solid var(--line)', position: 'sticky', top: 0, zIndex: 2 }}>
                    <th style={{ padding: '12px 14px', width: '50px', textAlign: 'center' }}>#</th>
                    <th style={{ padding: '12px 14px', width: '160px', textAlign: 'left' }}>Status & Confidence</th>
                    <th style={{ padding: '12px 14px', textAlign: 'left' }}>Dataset A (Left Record)</th>
                    <th style={{ padding: '12px 14px', textAlign: 'left' }}>Dataset B (Right Record)</th>
                    <th style={{ padding: '12px 14px', width: '100px', textAlign: 'center' }}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredResultItems.length > 0 ? (
                    filteredResultItems.map((item, idx) => {
                      const isExact = item.status === 'Exact Match';
                      const isPartial = item.status === 'Partial Match';
                      const isDisc = item.status === 'Value Discrepancy';
                      const isUnA = item.status?.includes('Dataset A Only');
                      const isUnB = item.status?.includes('Dataset B Only');

                      return (
                        <tr key={item.id || idx} style={{ borderBottom: '1px solid var(--line)', background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.015)' }}>
                          <td style={{ padding: '12px 14px', textAlign: 'center', fontWeight: 600, color: 'var(--muted)' }}>
                            {idx + 1}
                          </td>

                          {/* Status Badge */}
                          <td style={{ padding: '12px 14px', verticalAlign: 'middle' }}>
                            <div style={{
                              display: 'inline-flex',
                              alignItems: 'center',
                              gap: '6px',
                              padding: '4px 10px',
                              borderRadius: '12px',
                              fontSize: '12px',
                              fontWeight: 700,
                              background: isExact ? 'rgba(16, 185, 129, 0.15)' : isPartial ? 'rgba(245, 158, 11, 0.15)' : isDisc ? 'rgba(139, 92, 246, 0.15)' : 'rgba(239, 68, 68, 0.15)',
                              color: isExact ? '#059669' : isPartial ? '#d97706' : isDisc ? '#7c3aed' : '#dc2626'
                            }}>
                              {isExact && <CheckCircle2 size={13} />}
                              {isPartial && <Sparkles size={13} />}
                              {isDisc && <AlertTriangle size={13} />}
                              {(isUnA || isUnB) && <XCircle size={13} />}
                              {item.status}
                            </div>
                            {item.confidence > 0 && (
                              <div style={{ fontSize: '11px', color: 'var(--muted)', marginTop: '4px', fontWeight: 600 }}>
                                Score: {item.confidence}%
                              </div>
                            )}
                          </td>

                          {/* Row A Content */}
                          <td style={{ padding: '12px 14px', verticalAlign: 'top' }}>
                            {item.rowA ? (
                              <div>
                                {Object.entries(item.rowA).slice(0, 4).map(([k, v]) => (
                                  <div key={k} style={{ fontSize: '12px', marginBottom: '2px' }}>
                                    <strong style={{ color: 'var(--muted)' }}>{k}:</strong> {String(v)}
                                  </div>
                                ))}
                              </div>
                            ) : (
                              <span style={{ color: 'var(--muted)', fontStyle: 'italic' }}>— Not Present in Dataset A —</span>
                            )}
                          </td>

                          {/* Row B Content */}
                          <td style={{ padding: '12px 14px', verticalAlign: 'top' }}>
                            {item.rowB ? (
                              <div>
                                {Object.entries(item.rowB).slice(0, 4).map(([k, v]) => (
                                  <div key={k} style={{ fontSize: '12px', marginBottom: '2px' }}>
                                    <strong style={{ color: 'var(--muted)' }}>{k}:</strong> {String(v)}
                                  </div>
                                ))}
                              </div>
                            ) : (
                              <span style={{ color: 'var(--muted)', fontStyle: 'italic' }}>— Not Present in Dataset B —</span>
                            )}
                          </td>

                          {/* Inspect Action */}
                          <td style={{ padding: '12px 14px', textAlign: 'center', verticalAlign: 'middle' }}>
                            {(item.rowA && item.rowB) ? (
                              <button 
                                className="secondary"
                                onClick={() => setInspectModalItem(item)}
                                style={{ padding: '6px 10px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '4px' }}
                              >
                                <Eye size={14} /> Diff
                              </button>
                            ) : (
                              <span style={{ color: 'var(--muted)', fontSize: '12px' }}>Solo</span>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={5} style={{ textAlign: 'center', padding: '40px', color: 'var(--muted)' }}>
                        No records matching filter "{activeResultTab}" and search "{searchQuery}".
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Table Footer */}
            <div style={{ padding: '14px 20px', background: 'var(--bg)', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '13px', color: 'var(--muted)' }}>
              <div>Showing <strong>{filteredResultItems.length}</strong> record(s)</div>
              <div>Comparison generated at {comparisonResults.timestamp}</div>
            </div>
          </div>

          {/* Bottom Step Actions */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '8px' }}>
            <button className="secondary" onClick={() => setCurrentStep(2)}>
              ← Back to Mapping
            </button>
            <button 
              onClick={() => setCurrentStep(4)}
              style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '12px 28px', fontSize: '14.5px', fontWeight: 700 }}
            >
              Proceed to Export Studio →
            </button>
          </div>

        </div>
      )}

      {/* ========================================================================= */}
      {/* STEP 4: EXPORT STUDIO                                                     */}
      {/* ========================================================================= */}
      {currentStep === 4 && comparisonResults && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          {/* Export Options Deck */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '24px' }}>
            
            {/* Master Excel Report */}
            <div className="card" style={{ padding: '28px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', border: '1.5px solid var(--accent)' }}>
              <div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '14px' }}>
                  <div style={{ background: 'var(--accent)', color: 'white', padding: '10px', borderRadius: '10px' }}>
                    <FileSpreadsheet size={24} />
                  </div>
                  <div>
                    <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800 }}>Master Multi-Sheet Excel</h3>
                    <span style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600 }}>.XLSX Workbook Format</span>
                  </div>
                </div>
                <p style={{ color: 'var(--muted)', fontSize: '13.5px', lineHeight: '1.5', margin: '0 0 16px 0' }}>
                  Generates an executive-ready multi-tab Excel workbook containing Summary Dashboard, Exact Matches, Partial Matches, Discrepancies, and Unmatched sheets.
                </p>
                <ul style={{ fontSize: '12.5px', color: 'var(--ink)', paddingLeft: '18px', margin: '0 0 20px 0', lineHeight: '1.6' }}>
                  <li>Executive Summary Sheet with Match KPI Cards</li>
                  <li>Side-by-side key and confidence scoring</li>
                  <li>Highlighted attribute drift columns</li>
                </ul>
              </div>

              <button 
                onClick={exportMasterExcelWorkbook}
                style={{ width: '100%', padding: '12px', fontSize: '14px', fontWeight: 700, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}
              >
                <Download size={18} /> Download Master Excel (.xlsx)
              </button>
            </div>

            {/* ZIP Archive Bundle */}
            <div className="card" style={{ padding: '28px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between' }}>
              <div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '14px' }}>
                  <div style={{ background: 'var(--panel)', color: 'var(--accent)', border: '1px solid var(--line)', padding: '10px', borderRadius: '10px' }}>
                    <FileArchive size={24} />
                  </div>
                  <div>
                    <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800 }}>Complete ZIP Archive</h3>
                    <span style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600 }}>.ZIP Bundle with CSVs</span>
                  </div>
                </div>
                <p style={{ color: 'var(--muted)', fontSize: '13.5px', lineHeight: '1.5', margin: '0 0 16px 0' }}>
                  Download all matched, partial, discrepancy, and unmatched sets partitioned into separate, clean CSV files packaged in a single ZIP file.
                </p>
                <ul style={{ fontSize: '12.5px', color: 'var(--ink)', paddingLeft: '18px', margin: '0 0 20px 0', lineHeight: '1.6' }}>
                  <li>Individual CSV for each match category</li>
                  <li>Ready for import into databases, Power BI, or ERPs</li>
                </ul>
              </div>

              <button 
                className="secondary"
                onClick={exportAllZipBundle}
                style={{ width: '100%', padding: '12px', fontSize: '14px', fontWeight: 700, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}
              >
                <FileArchive size={18} /> Download All as ZIP (.zip)
              </button>
            </div>

          </div>

          {/* Quick Summary of Output */}
          <div className="card" style={{ padding: '24px', background: 'var(--panel)' }}>
            <h4 style={{ margin: '0 0 16px 0', fontSize: '16px', fontWeight: 700 }}>Reconciliation Output Breakdown</h4>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px' }}>
              <div style={{ background: 'var(--bg)', padding: '14px', borderRadius: '8px', border: '1px solid var(--line)' }}>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600 }}>Exact Matched Records</div>
                <div style={{ fontSize: '20px', fontWeight: 800, color: '#10b981', marginTop: '4px' }}>{comparisonResults.exactMatches.length}</div>
              </div>
              <div style={{ background: 'var(--bg)', padding: '14px', borderRadius: '8px', border: '1px solid var(--line)' }}>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600 }}>Fuzzy Partial Matches</div>
                <div style={{ fontSize: '20px', fontWeight: 800, color: '#f59e0b', marginTop: '4px' }}>{comparisonResults.partialMatches.length}</div>
              </div>
              <div style={{ background: 'var(--bg)', padding: '14px', borderRadius: '8px', border: '1px solid var(--line)' }}>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600 }}>Attribute Discrepancies</div>
                <div style={{ fontSize: '20px', fontWeight: 800, color: '#8b5cf6', marginTop: '4px' }}>{comparisonResults.valueDiscrepancies.length}</div>
              </div>
              <div style={{ background: 'var(--bg)', padding: '14px', borderRadius: '8px', border: '1px solid var(--line)' }}>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600 }}>Unmatched Left / Right</div>
                <div style={{ fontSize: '20px', fontWeight: 800, color: '#ef4444', marginTop: '4px' }}>
                  {comparisonResults.unmatchedA.length + comparisonResults.unmatchedB.length}
                </div>
              </div>
            </div>
          </div>

          {/* Navigation */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '8px' }}>
            <button className="secondary" onClick={() => setCurrentStep(3)}>
              ← Back to Results & Grid
            </button>
            <button 
              className="secondary"
              onClick={() => {
                setCurrentStep(1);
                setComparisonResults(null);
              }}
              style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
            >
              <RotateCcw size={16} /> Start New Comparison
            </button>
          </div>

        </div>
      )}

      {/* ========================================================================= */}
      {/* SIDE-BY-SIDE DIFF INSPECTOR MODAL                                         */}
      {/* ========================================================================= */}
      {inspectModalItem && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0,0,0,0.65)',
          backdropFilter: 'blur(4px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999,
          padding: '20px'
        }}>
          <div className="card" style={{ width: '100%', maxWidth: '850px', maxHeight: '90vh', overflowY: 'auto', padding: '28px', background: 'var(--panel)' }}>
            
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', borderBottom: '1px solid var(--line)', paddingBottom: '14px' }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '19px', fontWeight: 800, color: 'var(--ink)' }}>Side-by-Side Record Diff Inspector</h3>
                <div style={{ fontSize: '13px', color: 'var(--muted)', marginTop: '4px' }}>
                  Status: <strong>{inspectModalItem.status}</strong> • Confidence: <strong>{inspectModalItem.confidence}%</strong>
                </div>
              </div>
              <button 
                onClick={() => setInspectModalItem(null)}
                style={{ background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '50%', width: '32px', height: '32px', cursor: 'pointer', fontSize: '16px', fontWeight: 700 }}
              >
                ✕
              </button>
            </div>

            {/* Diff Comparison Grid */}
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px', marginBottom: '24px' }}>
              
              {/* Dataset A Card */}
              <div style={{ background: 'var(--bg)', padding: '16px', borderRadius: '8px', border: '1px solid var(--line)' }}>
                <div style={{ fontSize: '13px', fontWeight: 700, color: 'var(--accent)', marginBottom: '12px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <span>Dataset A (Left):</span> {datasetA.name}
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', fontSize: '12.5px' }}>
                  {Object.entries(inspectModalItem.rowA || {}).map(([k, v]) => (
                    <div key={k} style={{ padding: '6px 8px', background: 'var(--panel)', borderRadius: '4px', border: '1px solid var(--line)' }}>
                      <span style={{ color: 'var(--muted)', fontWeight: 600 }}>{k}:</span>
                      <div style={{ fontWeight: 700, color: 'var(--ink)', wordBreak: 'break-word', marginTop: '2px' }}>
                        {String(v)}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Dataset B Card */}
              <div style={{ background: 'var(--bg)', padding: '16px', borderRadius: '8px', border: '1px solid var(--line)' }}>
                <div style={{ fontSize: '13px', fontWeight: 700, color: 'var(--accent)', marginBottom: '12px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <span>Dataset B (Right):</span> {datasetB.name}
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', fontSize: '12.5px' }}>
                  {Object.entries(inspectModalItem.rowB || {}).map(([k, v]) => (
                    <div key={k} style={{ padding: '6px 8px', background: 'var(--panel)', borderRadius: '4px', border: '1px solid var(--line)' }}>
                      <span style={{ color: 'var(--muted)', fontWeight: 600 }}>{k}:</span>
                      <div style={{ fontWeight: 700, color: 'var(--ink)', wordBreak: 'break-word', marginTop: '2px' }}>
                        {String(v)}
                      </div>
                    </div>
                  ))}
                </div>
              </div>

            </div>

            {/* Discrepancies Alert */}
            {inspectModalItem.discrepancies?.length > 0 && (
              <div style={{ background: 'rgba(245, 158, 11, 0.1)', border: '1px solid #f59e0b', padding: '14px 18px', borderRadius: '8px', marginBottom: '20px' }}>
                <div style={{ fontWeight: 700, color: '#d97706', fontSize: '13px', marginBottom: '6px' }}>
                  ⚠️ Attribute Discrepancies Detected:
                </div>
                <ul style={{ margin: 0, paddingLeft: '20px', fontSize: '12.5px', color: 'var(--ink)' }}>
                  {inspectModalItem.discrepancies.map((d, i) => (
                    <li key={i} style={{ marginBottom: '4px' }}>
                      <strong>{d.fieldA}</strong> ("{d.valA}") vs <strong>{d.fieldB}</strong> ("{d.valB}")
                    </li>
                  ))}
                </ul>
              </div>
            )}

            <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
              <button onClick={() => setInspectModalItem(null)}>
                Close Inspector
              </button>
            </div>

          </div>
        </div>
      )}

    </div>
  );
};

export default DataComparisonPage;
