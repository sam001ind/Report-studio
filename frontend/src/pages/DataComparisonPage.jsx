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
  Percent,
  Sliders,
  FileText,
  ChevronLeft,
  ChevronRight
} from 'lucide-react';
import { parseWorkbookFromBuffer } from '../utils/excelParser';
import { stringSimilarity, calculateCompositeSimilarity } from '../utils/fuzzyMatch';

const SAMPLE_DATASET_A = [
  { "PRN": "XT2099001", "StudentName": "Student Alpha", "CourseCode": "ENG101", "CourseTitle": "English Literature", "College": "Alpha Arts and Science College", "Marks": "85" },
  { "PRN": "XT2099002", "StudentName": "Student Beta", "CourseCode": "MAL102", "CourseTitle": "Language Studies", "College": "Beta Commerce and Science College", "Marks": "78" },
  { "PRN": "XT2099003", "StudentName": "Student Gamma", "CourseCode": "ARB103", "CourseTitle": "Classical Studies", "College": "Gamma Government College", "Marks": "92" },
  { "PRN": "XT2099004", "StudentName": "Student Delta", "CourseCode": "CS104", "CourseTitle": "Data Structures", "College": "Delta Arts College", "Marks": "95" },
  { "PRN": "XT2099005", "StudentName": "Student Epsilon", "CourseCode": "PHY105", "CourseTitle": "Modern Physics", "College": "Epsilon Science College", "Marks": "64" },
  { "PRN": "XT2099006", "StudentName": "Student Zeta", "CourseCode": "CHE106", "CourseTitle": "Organic Chemistry", "College": "Zeta Government College", "Marks": "88" },
  { "PRN": "XT2099007", "StudentName": "Student Eta", "CourseCode": "MAT107", "CourseTitle": "Calculus & Matrices", "College": "Alpha Arts and Science College", "Marks": "72" },
  { "PRN": "XT2099008", "StudentName": "Student Theta", "CourseCode": "HIS108", "CourseTitle": "World History", "College": "Beta Commerce and Science College", "Marks": "60" }
];

const SAMPLE_DATASET_B = [
  { "RegisterNo": "XT2099001", "CandidateName": "Student Alpha", "SubjectCode": "ENG101", "SubjectTitle": "English Literature", "ExamCentre": "Alpha Arts College", "Score": "85" },
  { "RegisterNo": "XT2099002", "CandidateName": "Student Beta", "SubjectCode": "MAL102", "SubjectTitle": "Language Studies", "ExamCentre": "Beta Science College", "Score": "78" },
  { "RegisterNo": "XT2099003", "CandidateName": "Student Gamma X", "SubjectCode": "ARB103", "SubjectTitle": "Classical Studies", "ExamCentre": "Gamma College", "Score": "90" }, // Score discrepancy
  { "RegisterNo": "XT2099004", "CandidateName": "Student Delta", "SubjectCode": "CS104", "SubjectTitle": "Data Structures & Algorithms", "ExamCentre": "Delta Arts College", "Score": "95" },
  { "RegisterNo": "XT2099005", "CandidateName": "Student Epsilon M.", "SubjectCode": "PHY105", "SubjectTitle": "Modern Physics", "ExamCentre": "Epsilon Science College", "Score": "64" },
  { "RegisterNo": "XT2099006", "CandidateName": "Student Zeta", "SubjectCode": "CHE106", "SubjectTitle": "Organic Chemistry", "ExamCentre": "Zeta Government College", "Score": "88" },
  { "RegisterNo": "XT2099009", "CandidateName": "Student Iota", "SubjectCode": "COM109", "SubjectTitle": "Financial Accounting", "ExamCentre": "Iota Commerce College", "Score": "81" }, // Unmatched Right
  { "RegisterNo": "XT2099010", "CandidateName": "Student Kappa", "SubjectCode": "ZOO110", "SubjectTitle": "Animal Diversity", "ExamCentre": "Kappa Science College", "Score": "89" } // Unmatched Right
];

const formatFileSize = (bytes) => {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
};

const fileListStyles = {
  filesListContainer: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    background: 'var(--bg)',
    borderRadius: '8px',
    border: '1px solid var(--line)',
    padding: '10px 12px',
    marginBottom: '16px'
  },
  filesListHeader: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    paddingBottom: '6px',
    borderBottom: '1px dashed var(--line)'
  },
  filesListTitle: {
    fontSize: '11px',
    fontWeight: 700,
    textTransform: 'uppercase',
    letterSpacing: '0.5px',
    color: 'var(--muted)'
  },
  clearAllBtn: {
    background: 'transparent',
    border: 'none',
    color: '#ef4444',
    fontSize: '11px',
    fontWeight: 600,
    cursor: 'pointer',
    padding: '2px 6px',
    borderRadius: '4px'
  },
  filesScrollList: {
    display: 'flex',
    flexDirection: 'column',
    gap: '6px',
    maxHeight: '160px',
    overflowY: 'auto'
  },
  fileChip: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: '8px',
    background: 'var(--panel)',
    border: '1px solid var(--line)',
    borderRadius: '6px',
    padding: '6px 10px'
  },
  fileChipLeft: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
    minWidth: 0,
    flex: 1
  },
  fileChipName: {
    fontSize: '12px',
    fontWeight: 600,
    color: 'var(--ink)',
    whiteSpace: 'nowrap',
    overflow: 'hidden',
    textOverflow: 'ellipsis',
    maxWidth: '200px'
  },
  fileChipMeta: {
    fontSize: '10px',
    color: 'var(--muted)'
  },
  fileChipRight: {
    display: 'flex',
    alignItems: 'center',
    gap: '6px',
    flexShrink: 0
  },
  sheetMiniSelect: {
    fontSize: '10.5px',
    background: 'var(--bg)',
    color: 'var(--ink)',
    border: '1px solid var(--line)',
    borderRadius: '4px',
    padding: '2px 4px',
    maxWidth: '110px'
  },
  fileRemoveBtn: {
    background: 'transparent',
    border: 'none',
    color: 'var(--muted)',
    cursor: 'pointer',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    padding: '4px',
    borderRadius: '4px',
    transition: 'color 0.15s ease'
  }
};

const DataComparisonPage = () => {
  // Navigation Stepper: 1: Upload, 2: Rules & Mapping, 3: Analytics & Results, 4: Export
  const [currentStep, setCurrentStep] = useState(1);

  // Dataset A (Left / Reference) Files List
  const [datasetAFiles, setDatasetAFiles] = useState([]);
  const [isDragOverA, setIsDragOverA] = useState(false);

  // Dataset B (Right / Comparison) Files List
  const [datasetBFiles, setDatasetBFiles] = useState([]);
  const [isDragOverB, setIsDragOverB] = useState(false);

  // Derived Dataset A across all uploaded files (Memory-Optimized for 200k+ rows)
  const datasetA = useMemo(() => {
    if (datasetAFiles.length === 0) {
      return {
        name: 'Dataset_A',
        columns: [],
        rows: [],
        fileName: '',
        sheets: [],
        selectedSheet: '',
        rawWorkbook: null
      };
    }

    const primaryFile = datasetAFiles[0];

    // Single file fast-path: Avoid re-allocating 200,000+ objects
    if (datasetAFiles.length === 1) {
      return {
        name: primaryFile.name,
        columns: primaryFile.columns || [],
        rows: primaryFile.rows || [],
        fileName: primaryFile.name,
        sheets: primaryFile.sheetNames || [],
        selectedSheet: primaryFile.selectedSheet || '',
        rawWorkbook: primaryFile.workbook || null
      };
    }

    const allRows = [];
    const colSet = new Set();
    const orderedCols = [];

    datasetAFiles.forEach((file) => {
      (file.columns || []).forEach((c) => {
        if (!c.startsWith('_') && !colSet.has(c)) {
          colSet.add(c);
          orderedCols.push(c);
        }
      });
      (file.rows || []).forEach((r) => {
        allRows.push({
          ...r,
          _sourceFileName: file.name,
          _sourceFileId: file.id
        });
      });
    });

    const name = `${primaryFile.name} (+${datasetAFiles.length - 1} more)`;
    const fileName = datasetAFiles.map(f => f.name).join(', ');

    return {
      name,
      columns: orderedCols,
      rows: allRows,
      fileName,
      sheets: primaryFile.sheetNames || [],
      selectedSheet: primaryFile.selectedSheet || '',
      rawWorkbook: primaryFile.workbook || null
    };
  }, [datasetAFiles]);

  // Derived Dataset B across all uploaded files (Memory-Optimized for 200k+ rows)
  const datasetB = useMemo(() => {
    if (datasetBFiles.length === 0) {
      return {
        name: 'Dataset_B',
        columns: [],
        rows: [],
        fileName: '',
        sheets: [],
        selectedSheet: '',
        rawWorkbook: null
      };
    }

    const primaryFile = datasetBFiles[0];

    // Single file fast-path: Avoid re-allocating 200,000+ objects
    if (datasetBFiles.length === 1) {
      return {
        name: primaryFile.name,
        columns: primaryFile.columns || [],
        rows: primaryFile.rows || [],
        fileName: primaryFile.name,
        sheets: primaryFile.sheetNames || [],
        selectedSheet: primaryFile.selectedSheet || '',
        rawWorkbook: primaryFile.workbook || null
      };
    }

    const allRows = [];
    const colSet = new Set();
    const orderedCols = [];

    datasetBFiles.forEach((file) => {
      (file.columns || []).forEach((c) => {
        if (!c.startsWith('_') && !colSet.has(c)) {
          colSet.add(c);
          orderedCols.push(c);
        }
      });
      (file.rows || []).forEach((r) => {
        allRows.push({
          ...r,
          _sourceFileName: file.name,
          _sourceFileId: file.id
        });
      });
    });

    const name = `${primaryFile.name} (+${datasetBFiles.length - 1} more)`;
    const fileName = datasetBFiles.map(f => f.name).join(', ');

    return {
      name,
      columns: orderedCols,
      rows: allRows,
      fileName,
      sheets: primaryFile.sheetNames || [],
      selectedSheet: primaryFile.selectedSheet || '',
      rawWorkbook: primaryFile.workbook || null
    };
  }, [datasetBFiles]);

  // Key Column Mappings (Composite Keys support)
  const [keyMappings, setKeyMappings] = useState([
    { id: 1, leftCol: '', rightCol: '' }
  ]);

  // Non-key columns to compare for attribute drift/discrepancies
  const [valueCompareMappings, setValueCompareMappings] = useState([]);

  // Matching Rules & Parameters: Default to 'exact' for blazing-fast 200k+ performance
  const [matchMode, setMatchMode] = useState('exact'); // 'exact' | 'fuzzy'
  const [fuzzyThreshold, setFuzzyThreshold] = useState(80); // 50 - 100 (%)
  const [comparisonType, setComparisonType] = useState('one_to_one'); // 'one_to_one' | 'one_to_many' | 'many_to_many'
  
  // Normalization Options
  const [normOptions, setNormOptions] = useState({
    ignoreCase: true,
    stripSpaces: true,
    stripPunctuation: true,
    stripLeadingZeros: true
  });

  // Processing, Progress & Status
  const [isProcessing, setIsProcessing] = useState(false);
  const [calculationProgress, setCalculationProgress] = useState(0);
  const [statusMsg, setStatusMsg] = useState('Ready to upload datasets.');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Comparison Output Results
  const [comparisonResults, setComparisonResults] = useState(null);

  // UI Filter, Search & Pagination States
  const [activeResultTab, setActiveResultTab] = useState('all'); // 'all' | 'exact' | 'partial' | 'discrepancy' | 'unmatched_a' | 'unmatched_b' | 'duplicates'
  const [searchQuery, setSearchQuery] = useState('');
  const [inspectModalItem, setInspectModalItem] = useState(null);
  const [resultPage, setResultPage] = useState(1);
  const [resultPageSize, setResultPageSize] = useState(50);

  // Custom Report Builder States (Step 4)
  const [customSelectedCols, setCustomSelectedCols] = useState(null);
  const [customRecordScope, setCustomRecordScope] = useState('all'); // 'all' | 'matched' | 'discrepancy' | 'partial' | 'unmatched_a' | 'unmatched_b'
  const [customReportName, setCustomReportName] = useState('Custom_Reconciled_Report');

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  // Helper to load Sample Datasets for 1-click test
  const loadSampleData = () => {
    const colsA = Object.keys(SAMPLE_DATASET_A[0]);
    const colsB = Object.keys(SAMPLE_DATASET_B[0]);

    setDatasetAFiles([
      {
        id: 'sample_file_a',
        name: 'Admission_Master_List.xlsx',
        size: 10240,
        sizeFormatted: '10.0 KB',
        workbook: null,
        sheetNames: ['Master_2025'],
        selectedSheet: 'Master_2025',
        rows: SAMPLE_DATASET_A,
        columns: colsA
      }
    ]);

    setDatasetBFiles([
      {
        id: 'sample_file_b',
        name: 'Exam_Registration_Report.xlsx',
        size: 12288,
        sizeFormatted: '12.0 KB',
        workbook: null,
        sheetNames: ['Registrations_Nov2025'],
        selectedSheet: 'Registrations_Nov2025',
        rows: SAMPLE_DATASET_B,
        columns: colsB
      }
    ]);

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
        !unzipped.files[fn].dir && !fn.startsWith('__MACOSX') && (fn.endsWith('.xlsx') || fn.endsWith('.xls') || fn.endsWith('.csv'))
      );

      if (excelFiles.length === 0) {
        throw new Error(`No valid Excel or CSV files found inside the ZIP archive "${file.name}".`);
      }

      const parsedList = await Promise.all(
        excelFiles.map(async (targetFileName) => {
          const blob = await unzipped.files[targetFileName].async('blob');
          const cleanName = targetFileName.split('/').pop();
          return await parseSpreadsheetBlob(blob, cleanName, blob.size);
        })
      );
      return parsedList;
    } else {
      const parsed = await parseSpreadsheetBlob(file, file.name, file.size);
      return [parsed];
    }
  };

  const parseSpreadsheetBlob = (fileOrBlob, originalFileName, fileSize = 0) => {
    return new Promise((resolve, reject) => {
      const ext = originalFileName.split('.').pop().toLowerCase();
      const fileId = `${originalFileName}_${fileSize || 0}_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
      const sizeFormatted = formatFileSize(fileSize || (fileOrBlob.size || 0));

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
              id: fileId,
              name: originalFileName,
              size: fileSize || fileOrBlob.size || 0,
              sizeFormatted,
              columns,
              rows: results.data,
              fileName: originalFileName,
              sheetNames: ['Sheet1'],
              selectedSheet: 'Sheet1',
              workbook: null
            });
          },
          error: reject
        });
      } else {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const buffer = e.target.result;
            const workbook = parseWorkbookFromBuffer(buffer);
            
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
              id: fileId,
              name: originalFileName,
              size: fileSize || fileOrBlob.size || 0,
              sizeFormatted,
              columns,
              rows,
              fileName: originalFileName,
              sheetNames: workbook.SheetNames,
              selectedSheet: firstSheetName,
              workbook
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

  // Upload Batch Handlers for Dataset A
  const handleUploadFilesA = async (fileList) => {
    const files = Array.from(fileList || []);
    if (!files.length) return;

    setIsProcessing(true);
    setStatus(`Processing ${files.length} file(s) for Dataset A...`);

    try {
      const parsedBatches = await Promise.all(files.map(f => parseUploadedFile(f)));
      const newFiles = parsedBatches.flat();

      setDatasetAFiles(prev => {
        const existingKeys = new Set(prev.map(f => `${f.name}_${f.size}`));
        const uniqueNew = newFiles.filter(f => !existingKeys.has(`${f.name}_${f.size}`));
        return [...prev, ...uniqueNew];
      });

      const totalLoadedRows = newFiles.reduce((acc, f) => acc + f.rows.length, 0);
      setStatus(`Dataset A loaded ${newFiles.length} file(s) with ${totalLoadedRows} rows.`, 'success');

      // Auto-suggest initial key if Dataset B already loaded
      const allColsA = Array.from(new Set(newFiles.flatMap(f => f.columns)));
      if (datasetB.columns.length > 0 && (!keyMappings[0]?.leftCol || !keyMappings[0]?.rightCol)) {
        autoSuggestKeyMappings(allColsA, datasetB.columns);
      }
    } catch (err) {
      console.error(err);
      setStatus(`Error loading Dataset A: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleUploadA = async (e) => {
    await handleUploadFilesA(e.target.files);
    if (e.target) e.target.value = '';
  };

  // Upload Batch Handlers for Dataset B
  const handleUploadFilesB = async (fileList) => {
    const files = Array.from(fileList || []);
    if (!files.length) return;

    setIsProcessing(true);
    setStatus(`Processing ${files.length} file(s) for Dataset B...`);

    try {
      const parsedBatches = await Promise.all(files.map(f => parseUploadedFile(f)));
      const newFiles = parsedBatches.flat();

      setDatasetBFiles(prev => {
        const existingKeys = new Set(prev.map(f => `${f.name}_${f.size}`));
        const uniqueNew = newFiles.filter(f => !existingKeys.has(`${f.name}_${f.size}`));
        return [...prev, ...uniqueNew];
      });

      const totalLoadedRows = newFiles.reduce((acc, f) => acc + f.rows.length, 0);
      setStatus(`Dataset B loaded ${newFiles.length} file(s) with ${totalLoadedRows} rows.`, 'success');

      // Auto-suggest initial key if Dataset A already loaded
      const allColsB = Array.from(new Set(newFiles.flatMap(f => f.columns)));
      if (datasetA.columns.length > 0 && (!keyMappings[0]?.leftCol || !keyMappings[0]?.rightCol)) {
        autoSuggestKeyMappings(datasetA.columns, allColsB);
      }
    } catch (err) {
      console.error(err);
      setStatus(`Error loading Dataset B: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleUploadB = async (e) => {
    await handleUploadFilesB(e.target.files);
    if (e.target) e.target.value = '';
  };

  // Handle Sheet Change per File in Dataset A
  const handleSheetChangeFileA = (fileId, newSheet) => {
    setDatasetAFiles(prev =>
      prev.map(f => {
        if (f.id !== fileId || !f.workbook || !f.workbook.Sheets[newSheet]) return f;
        const worksheet = f.workbook.Sheets[newSheet];
        const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        const columns = rows.length > 0 ? Object.keys(rows[0]).map(c => c.trim()).filter(Boolean) : [];
        return {
          ...f,
          selectedSheet: newSheet,
          rows,
          columns
        };
      })
    );
  };

  // Handle Sheet Change per File in Dataset B
  const handleSheetChangeFileB = (fileId, newSheet) => {
    setDatasetBFiles(prev =>
      prev.map(f => {
        if (f.id !== fileId || !f.workbook || !f.workbook.Sheets[newSheet]) return f;
        const worksheet = f.workbook.Sheets[newSheet];
        const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        const columns = rows.length > 0 ? Object.keys(rows[0]).map(c => c.trim()).filter(Boolean) : [];
        return {
          ...f,
          selectedSheet: newSheet,
          rows,
          columns
        };
      })
    );
  };

  const handleRemoveFileA = (fileId) => {
    setDatasetAFiles(prev => prev.filter(f => f.id !== fileId));
  };

  const handleRemoveFileB = (fileId) => {
    setDatasetBFiles(prev => prev.filter(f => f.id !== fileId));
  };

  const handleClearFilesA = () => {
    setDatasetAFiles([]);
    setStatus('Cleared all Dataset A files.', 'normal');
  };

  const handleClearFilesB = () => {
    setDatasetBFiles([]);
    setStatus('Cleared all Dataset B files.', 'normal');
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

  // CORE HIGH-SPEED COMPARISON & RECONCILIATION ENGINE (Optimized for 200k+ rows)
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
    setCalculationProgress(5);
    setStatus("Preparing records and composite keys...");

    // Yield to UI thread so progress indicator renders cleanly
    setTimeout(() => {
      try {
        const rowsA = datasetA.rows;
        const rowsB = datasetB.rows;
        const lenA = rowsA.length;
        const lenB = rowsB.length;
        const isLargeDataset = lenA > 5000 || lenB > 5000;

        const fastNorm = (text) => {
          if (text === null || text === undefined) return '';
          let str = String(text);
          if (normOptions.stripSpaces) str = str.trim();
          if (normOptions.ignoreCase) str = str.toLowerCase();
          if (normOptions.stripPunctuation) str = str.replace(/[\s.,\-_#;:/\\]+/g, normOptions.stripSpaces ? '' : ' ');
          else if (normOptions.stripSpaces) str = str.replace(/\s+/g, ' ');
          if (normOptions.stripLeadingZeros && str.length > 1 && str.charCodeAt(0) === 48) {
            str = str.replace(/^0+(?=\d)/, '');
          }
          return str;
        };

        const getCompositeKey = (row, isLeft) => {
          let k = '';
          for (let i = 0; i < activeKeys.length; i++) {
            const col = isLeft ? activeKeys[i].leftCol : activeKeys[i].rightCol;
            const p = fastNorm(row[col]);
            k += (i === 0 ? p : '||' + p);
          }
          return k;
        };

        // 1. Precompute keys and detect duplicates
        const keysA = new Array(lenA);
        const seenKeysA = new Set();
        const duplicateA = [];
        for (let i = 0; i < lenA; i++) {
          const k = getCompositeKey(rowsA[i], true);
          keysA[i] = k;
          if (k && seenKeysA.has(k)) {
            duplicateA.push({ ...rowsA[i], _source: 'Dataset A', _rowIndex: i + 1, _dupKey: k });
          } else if (k) {
            seenKeysA.add(k);
          }
        }

        const keysB = new Array(lenB);
        const seenKeysB = new Set();
        const duplicateB = [];
        const indexMapB = new Map();
        for (let i = 0; i < lenB; i++) {
          const k = getCompositeKey(rowsB[i], false);
          keysB[i] = k;
          if (k && seenKeysB.has(k)) {
            duplicateB.push({ ...rowsB[i], _source: 'Dataset B', _rowIndex: i + 1, _dupKey: k });
          } else if (k) {
            seenKeysB.add(k);
          }

          if (k) {
            const existing = indexMapB.get(k);
            if (existing === undefined) {
              indexMapB.set(k, i);
            } else if (typeof existing === 'number') {
              indexMapB.set(k, [existing, i]);
            } else {
              existing.push(i);
            }
          }
        }

        setCalculationProgress(35);
        setStatus(`Indexed ${lenB.toLocaleString()} reference keys. Comparing records...`);

        setTimeout(() => {
          try {
            // 2. High-speed Exact Matching & Value Discrepancies
            const matchedIndexA = new Uint8Array(lenA);
            const matchedIndexB = new Uint8Array(lenB);
            const exactMatches = [];
            const valueDiscrepancies = [];
            const activeValueComps = valueCompareMappings.filter(v => v.leftCol && v.rightCol);

            for (let idxA = 0; idxA < lenA; idxA++) {
              const keyA = keysA[idxA];
              if (!keyA) continue;

              const targetB = indexMapB.get(keyA);
              if (targetB !== undefined) {
                let idxB = -1;
                if (typeof targetB === 'number') {
                  if (matchedIndexB[targetB] === 0 || comparisonType !== 'one_to_one') {
                    idxB = targetB;
                  }
                } else {
                  if (comparisonType === 'one_to_one') {
                    for (let j = 0; j < targetB.length; j++) {
                      const candidate = targetB[j];
                      if (matchedIndexB[candidate] === 0) {
                        idxB = candidate;
                        break;
                      }
                    }
                  } else {
                    idxB = targetB[0];
                  }
                }

                if (idxB !== -1) {
                  matchedIndexA[idxA] = 1;
                  matchedIndexB[idxB] = 1;
                  const rowA = rowsA[idxA];
                  const rowB = rowsB[idxB];

                  let hasDisc = false;
                  const discrepancies = [];
                  for (let c = 0; c < activeValueComps.length; c++) {
                    const comp = activeValueComps[c];
                    const valA = fastNorm(rowA[comp.leftCol]);
                    const valB = fastNorm(rowB[comp.rightCol]);
                    if (valA !== valB) {
                      hasDisc = true;
                      discrepancies.push({
                        fieldA: comp.leftCol,
                        fieldB: comp.rightCol,
                        valA: rowA[comp.leftCol],
                        valB: rowB[comp.rightCol]
                      });
                    }
                  }

                  if (hasDisc) {
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
              }
            }

            setCalculationProgress(70);

            // 3. Fuzzy Matching (Safe Blocked Candidate Pass)
            const partialMatches = [];
            const thresholdRatio = fuzzyThreshold / 100.0;

            if (matchMode === 'fuzzy') {
              setStatus("Evaluating candidate fuzzy similarities...");

              const unmatchedIndicesA = [];
              for (let i = 0; i < lenA; i++) if (matchedIndexA[i] === 0 && keysA[i]) unmatchedIndicesA.push(i);

              const unmatchedIndicesB = [];
              for (let i = 0; i < lenB; i++) if (matchedIndexB[i] === 0 && keysB[i]) unmatchedIndicesB.push(i);

              if (unmatchedIndicesA.length > 0 && unmatchedIndicesB.length > 0) {
                if (isLargeDataset) {
                  // Prefix blocking for large datasets to prevent browser hang
                  const blockMapB = new Map();
                  for (let b = 0; b < unmatchedIndicesB.length; b++) {
                    const idxB = unmatchedIndicesB[b];
                    const prefix = keysB[idxB].slice(0, 3);
                    let list = blockMapB.get(prefix);
                    if (!list) {
                      list = [];
                      blockMapB.set(prefix, list);
                    }
                    if (list.length < 50) list.push(idxB);
                  }

                  for (let a = 0; a < unmatchedIndicesA.length; a++) {
                    const idxA = unmatchedIndicesA[a];
                    if (matchedIndexA[idxA] === 1) continue;
                    const prefix = keysA[idxA].slice(0, 3);
                    const candidateList = blockMapB.get(prefix);
                    if (!candidateList || candidateList.length === 0) continue;

                    let bestIdxB = -1;
                    let bestScore = 0;

                    for (let c = 0; c < candidateList.length; c++) {
                      const idxB = candidateList[c];
                      if (matchedIndexB[idxB] === 1) continue;
                      const score = calculateCompositeSimilarity(rowsA[idxA], rowsB[idxB], activeKeys, normOptions);
                      if (score >= thresholdRatio && score > bestScore) {
                        bestScore = score;
                        bestIdxB = idxB;
                      }
                    }

                    if (bestIdxB !== -1) {
                      matchedIndexA[idxA] = 1;
                      matchedIndexB[bestIdxB] = 1;
                      const rowA = rowsA[idxA];
                      const rowB = rowsB[bestIdxB];

                      const discrepancies = [];
                      for (let vc = 0; vc < activeValueComps.length; vc++) {
                        const comp = activeValueComps[vc];
                        const valA = fastNorm(rowA[comp.leftCol]);
                        const valB = fastNorm(rowB[comp.rightCol]);
                        if (valA !== valB) {
                          discrepancies.push({
                            fieldA: comp.leftCol,
                            fieldB: comp.rightCol,
                            valA: rowA[comp.leftCol],
                            valB: rowB[comp.rightCol]
                          });
                        }
                      }

                      partialMatches.push({
                        id: `fuzzy_${idxA}_${bestIdxB}`,
                        status: 'Partial Match',
                        confidence: Math.round(bestScore * 100),
                        key: keysA[idxA],
                        rowA,
                        rowB,
                        discrepancies,
                        matchType: `Fuzzy Similarity (${Math.round(bestScore * 100)}%)`
                      });
                    }
                  }
                } else {
                  // Small dataset (< 5000 rows): pairwise comparison
                  for (let a = 0; a < unmatchedIndicesA.length; a++) {
                    const idxA = unmatchedIndicesA[a];
                    if (matchedIndexA[idxA] === 1) continue;

                    let bestIdxB = -1;
                    let bestScore = 0;

                    for (let b = 0; b < unmatchedIndicesB.length; b++) {
                      const idxB = unmatchedIndicesB[b];
                      if (matchedIndexB[idxB] === 1) continue;

                      const score = calculateCompositeSimilarity(rowsA[idxA], rowsB[idxB], activeKeys, normOptions);
                      if (score >= thresholdRatio && score > bestScore) {
                        bestScore = score;
                        bestIdxB = idxB;
                      }
                    }

                    if (bestIdxB !== -1) {
                      matchedIndexA[idxA] = 1;
                      matchedIndexB[bestIdxB] = 1;
                      const rowA = rowsA[idxA];
                      const rowB = rowsB[bestIdxB];

                      const discrepancies = [];
                      for (let vc = 0; vc < activeValueComps.length; vc++) {
                        const comp = activeValueComps[vc];
                        const valA = fastNorm(rowA[comp.leftCol]);
                        const valB = fastNorm(rowB[comp.rightCol]);
                        if (valA !== valB) {
                          discrepancies.push({
                            fieldA: comp.leftCol,
                            fieldB: comp.rightCol,
                            valA: rowA[comp.leftCol],
                            valB: rowB[comp.rightCol]
                          });
                        }
                      }

                      partialMatches.push({
                        id: `fuzzy_${idxA}_${bestIdxB}`,
                        status: 'Partial Match',
                        confidence: Math.round(bestScore * 100),
                        key: keysA[idxA],
                        rowA,
                        rowB,
                        discrepancies,
                        matchType: `Fuzzy Similarity (${Math.round(bestScore * 100)}%)`
                      });
                    }
                  }
                }
              }
            }

            setCalculationProgress(90);
            setStatus("Compiling reconciliation statistics...");

            // 4. Unmatched left & right
            const unmatchedA = [];
            for (let i = 0; i < lenA; i++) {
              if (matchedIndexA[i] === 0) {
                unmatchedA.push({
                  id: `un_a_${i}`,
                  status: 'Unmatched (Dataset A Only)',
                  confidence: 0,
                  rowA: rowsA[i],
                  rowB: null,
                  key: keysA[i]
                });
              }
            }

            const unmatchedB = [];
            for (let i = 0; i < lenB; i++) {
              if (matchedIndexB[i] === 0) {
                unmatchedB.push({
                  id: `un_b_${i}`,
                  status: 'Unmatched (Dataset B Only)',
                  confidence: 0,
                  rowA: null,
                  rowB: rowsB[i],
                  key: keysB[i]
                });
              }
            }

            // Column Discrepancy Statistics
            const columnDiscrepancyCounts = {};
            activeValueComps.forEach(comp => {
              const label = `${comp.leftCol} ↔ ${comp.rightCol}`;
              columnDiscrepancyCounts[label] = 0;
            });

            for (let d = 0; d < valueDiscrepancies.length; d++) {
              const item = valueDiscrepancies[d];
              for (let c = 0; c < item.discrepancies.length; c++) {
                const label = `${item.discrepancies[c].fieldA} ↔ ${item.discrepancies[c].fieldB}`;
                if (columnDiscrepancyCounts[label] !== undefined) {
                  columnDiscrepancyCounts[label] += 1;
                }
              }
            }
            for (let p = 0; p < partialMatches.length; p++) {
              const item = partialMatches[p];
              for (let c = 0; c < item.discrepancies.length; c++) {
                const label = `${item.discrepancies[c].fieldA} ↔ ${item.discrepancies[c].fieldB}`;
                if (columnDiscrepancyCounts[label] !== undefined) {
                  columnDiscrepancyCounts[label] += 1;
                }
              }
            }

            const totalMatchedRecords = exactMatches.length + partialMatches.length + valueDiscrepancies.length;
            const totalBaseRecords = Math.max(lenA, lenB);
            const matchPercentage = totalBaseRecords > 0 ? ((totalMatchedRecords / totalBaseRecords) * 100).toFixed(1) : '0';

            const results = {
              totalRowsA: lenA,
              totalRowsB: lenB,
              totalProcessed: lenA + lenB,
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
            setResultPage(1);
            setCalculationProgress(100);
            setCurrentStep(3);
            setStatus(`Reconciliation Complete! ${totalMatchedRecords.toLocaleString()} records matched (${matchPercentage}% Match Rate).`, 'success');
          } catch (innerErr) {
            console.error("Comparison Engine execution error:", innerErr);
            setStatus(`Comparison error: ${innerErr.message}`, 'error');
          } finally {
            setIsProcessing(false);
          }
        }, 30);
      } catch (err) {
        console.error("Comparison Engine initialization error:", err);
        setStatus(`Comparison error: ${err.message}`, 'error');
        setIsProcessing(false);
      }
    }, 30);
  }, [datasetA, datasetB, keyMappings, valueCompareMappings, matchMode, fuzzyThreshold, normOptions, comparisonType]);

  // Filtered rows for results table (Fast search without JSON.stringify memory bloat)
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

    const q = searchQuery.toLowerCase().trim();
    return list.filter(item => {
      if (item.key && String(item.key).toLowerCase().includes(q)) return true;
      if (item.status && String(item.status).toLowerCase().includes(q)) return true;
      if (item.rowA) {
        for (const k in item.rowA) {
          if (!k.startsWith('_') && String(item.rowA[k] || '').toLowerCase().includes(q)) return true;
        }
      }
      if (item.rowB) {
        for (const k in item.rowB) {
          if (!k.startsWith('_') && String(item.rowB[k] || '').toLowerCase().includes(q)) return true;
        }
      }
      return false;
    });
  }, [comparisonResults, activeResultTab, searchQuery]);

  // Server/Client-side Table Pagination (Critical for 200k+ rows)
  const totalResultCount = filteredResultItems.length;
  const totalResultPages = Math.max(1, Math.ceil(totalResultCount / resultPageSize));
  const safePage = Math.min(Math.max(1, resultPage), totalResultPages);

  const pagedResultItems = useMemo(() => {
    const startIdx = (safePage - 1) * resultPageSize;
    return filteredResultItems.slice(startIdx, startIdx + resultPageSize);
  }, [filteredResultItems, safePage, resultPageSize]);

  // Tab switch handler: auto-reset page to 1
  const handleTabSwitch = (tabId) => {
    setActiveResultTab(tabId);
    setResultPage(1);
  };

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
        ...comparisonResults.exactMatches.map(m => ({
          "Match_Status": m.status,
          "Confidence_%": m.confidence,
          ...(datasetAFiles.length > 1 ? { "Source_File_A": m.rowA?._sourceFileName || '' } : {}),
          ...(datasetBFiles.length > 1 ? { "Source_File_B": m.rowB?._sourceFileName || '' } : {}),
          ...Object.fromEntries(Object.entries(m.rowA || {}).filter(([k]) => !k.startsWith('_')))
        })),
        ...comparisonResults.partialMatches.map(m => {
          const row = {
            "Match_Status": m.status,
            "Confidence_%": m.confidence,
            "Match_Type": m.matchType,
            ...(datasetAFiles.length > 1 ? { "Source_File_A": m.rowA?._sourceFileName || '' } : {}),
            ...(datasetBFiles.length > 1 ? { "Source_File_B": m.rowB?._sourceFileName || '' } : {})
          };
          Object.keys(m.rowA || {}).filter(k => !k.startsWith('_')).forEach(k => { row[`A_${k}`] = m.rowA[k]; });
          Object.keys(m.rowB || {}).filter(k => !k.startsWith('_')).forEach(k => { row[`B_${k}`] = m.rowB[k]; });
          return row;
        }),
        ...comparisonResults.valueDiscrepancies.map(m => {
          const row = {
            "Match_Status": m.status,
            "Confidence_%": m.confidence,
            "Discrepancies": m.discrepancies.map(d => `${d.fieldA} vs ${d.fieldB}`).join('; '),
            ...(datasetAFiles.length > 1 ? { "Source_File_A": m.rowA?._sourceFileName || '' } : {}),
            ...(datasetBFiles.length > 1 ? { "Source_File_B": m.rowB?._sourceFileName || '' } : {})
          };
          Object.keys(m.rowA || {}).filter(k => !k.startsWith('_')).forEach(k => { row[`A_${k}`] = m.rowA[k]; });
          Object.keys(m.rowB || {}).filter(k => !k.startsWith('_')).forEach(k => { row[`B_${k}`] = m.rowB[k]; });
          return row;
        }),
        ...comparisonResults.unmatchedA.map(u => ({
          "Match_Status": u.status,
          ...(datasetAFiles.length > 1 ? { "Source_File_A": u.rowA?._sourceFileName || '' } : {}),
          ...Object.fromEntries(Object.entries(u.rowA || {}).filter(([k]) => !k.startsWith('_')))
        })),
        ...comparisonResults.unmatchedB.map(u => ({
          "Match_Status": u.status,
          ...(datasetBFiles.length > 1 ? { "Source_File_B": u.rowB?._sourceFileName || '' } : {}),
          ...Object.fromEntries(Object.entries(u.rowB || {}).filter(([k]) => !k.startsWith('_')))
        }))
      ];
    } else if (activeResultTab === 'exact') {
      sheetName = 'Exact_Matches';
      dataToExport = comparisonResults.exactMatches.map((m, idx) => ({
        "Match_ID": idx + 1,
        "Status": m.status,
        "Confidence_%": m.confidence,
        ...(datasetAFiles.length > 1 ? { "Source_File_A": m.rowA?._sourceFileName || '' } : {}),
        ...(datasetBFiles.length > 1 ? { "Source_File_B": m.rowB?._sourceFileName || '' } : {}),
        ...Object.fromEntries(Object.entries(m.rowA || {}).filter(([k]) => !k.startsWith('_')))
      }));
    } else if (activeResultTab === 'partial') {
      sheetName = 'Partial_Matches';
      dataToExport = comparisonResults.partialMatches.map((m, idx) => {
        const row = {
          "Match_ID": idx + 1,
          "Status": m.status,
          "Confidence_%": m.confidence,
          "Match_Explanation": m.matchType,
          ...(datasetAFiles.length > 1 ? { "Source_File_A": m.rowA?._sourceFileName || '' } : {}),
          ...(datasetBFiles.length > 1 ? { "Source_File_B": m.rowB?._sourceFileName || '' } : {})
        };
        Object.keys(m.rowA || {}).filter(k => !k.startsWith('_')).forEach(k => { row[`A_${k}`] = m.rowA[k]; });
        Object.keys(m.rowB || {}).filter(k => !k.startsWith('_')).forEach(k => { row[`B_${k}`] = m.rowB[k]; });
        return row;
      });
    } else if (activeResultTab === 'discrepancy') {
      sheetName = 'Value_Discrepancies';
      dataToExport = comparisonResults.valueDiscrepancies.map((m, idx) => {
        const row = {
          "Match_ID": idx + 1,
          "Status": m.status,
          "Confidence_%": m.confidence,
          "Discrepancies": m.discrepancies.map(d => `${d.fieldA}("${d.valA}" vs "${d.valB}")`).join('; '),
          ...(datasetAFiles.length > 1 ? { "Source_File_A": m.rowA?._sourceFileName || '' } : {}),
          ...(datasetBFiles.length > 1 ? { "Source_File_B": m.rowB?._sourceFileName || '' } : {})
        };
        Object.keys(m.rowA || {}).filter(k => !k.startsWith('_')).forEach(k => { row[`A_${k}`] = m.rowA[k]; });
        Object.keys(m.rowB || {}).filter(k => !k.startsWith('_')).forEach(k => { row[`B_${k}`] = m.rowB[k]; });
        return row;
      });
    } else if (activeResultTab === 'unmatched_a') {
      sheetName = 'Unmatched_Left_A';
      dataToExport = comparisonResults.unmatchedA.map(u => ({
        ...(datasetAFiles.length > 1 ? { "Source_File_A": u.rowA?._sourceFileName || '' } : {}),
        ...Object.fromEntries(Object.entries(u.rowA || {}).filter(([k]) => !k.startsWith('_')))
      }));
    } else if (activeResultTab === 'unmatched_b') {
      sheetName = 'Unmatched_Right_B';
      dataToExport = comparisonResults.unmatchedB.map(u => ({
        ...(datasetBFiles.length > 1 ? { "Source_File_B": u.rowB?._sourceFileName || '' } : {}),
        ...Object.fromEntries(Object.entries(u.rowB || {}).filter(([k]) => !k.startsWith('_')))
      }));
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
      ["Dataset B (Comparison)", datasetB.name]
    ];

    if (datasetAFiles.length > 0) {
      summaryData.push(["Dataset A Source Files", datasetAFiles.map(f => `${f.name} (${f.rows.length} rows)`).join('; ')]);
    }
    if (datasetBFiles.length > 0) {
      summaryData.push(["Dataset B Source Files", datasetBFiles.map(f => `${f.name} (${f.rows.length} rows)`).join('; ')]);
    }

    summaryData.push(
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
    );

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
        ...(datasetAFiles.length > 1 ? { "Source_File_A": it.rowA?._sourceFileName || '' } : {}),
        ...(datasetBFiles.length > 1 ? { "Source_File_B": it.rowB?._sourceFileName || '' } : {}),
        ...Object.fromEntries(Object.entries(it.rowA || {}).filter(([k]) => !k.startsWith('_')))
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
          "Match_Explanation": it.matchType,
          ...(datasetAFiles.length > 1 ? { "Source_File_A": it.rowA?._sourceFileName || '' } : {}),
          ...(datasetBFiles.length > 1 ? { "Source_File_B": it.rowB?._sourceFileName || '' } : {})
        };
        Object.keys(it.rowA || {}).filter(k => !k.startsWith('_')).forEach(k => { flat[`A_${k}`] = it.rowA[k]; });
        Object.keys(it.rowB || {}).filter(k => !k.startsWith('_')).forEach(k => { flat[`B_${k}`] = it.rowB[k]; });
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
          "Discrepancies_Found": discSummary,
          ...(datasetAFiles.length > 1 ? { "Source_File_A": it.rowA?._sourceFileName || '' } : {}),
          ...(datasetBFiles.length > 1 ? { "Source_File_B": it.rowB?._sourceFileName || '' } : {})
        };
        Object.keys(it.rowA || {}).filter(k => !k.startsWith('_')).forEach(k => { flat[`A_${k}`] = it.rowA[k]; });
        Object.keys(it.rowB || {}).filter(k => !k.startsWith('_')).forEach(k => { flat[`B_${k}`] = it.rowB[k]; });
        return flat;
      });
      const wsDisc = XLSX.utils.json_to_sheet(flatDisc);
      XLSX.utils.book_append_sheet(wb, wsDisc, "Value_Discrepancies");
    }

    // 5. Unmatched A Sheet
    if (comparisonResults.unmatchedA.length > 0) {
      const flatUnA = comparisonResults.unmatchedA.map(it => ({
        ...(datasetAFiles.length > 1 ? { "Source_File_A": it.rowA?._sourceFileName || '' } : {}),
        ...Object.fromEntries(Object.entries(it.rowA || {}).filter(([k]) => !k.startsWith('_')))
      }));
      const wsUnA = XLSX.utils.json_to_sheet(flatUnA);
      XLSX.utils.book_append_sheet(wb, wsUnA, "Unmatched_Left_A");
    }

    // 6. Unmatched B Sheet
    if (comparisonResults.unmatchedB.length > 0) {
      const flatUnB = comparisonResults.unmatchedB.map(it => ({
        ...(datasetBFiles.length > 1 ? { "Source_File_B": it.rowB?._sourceFileName || '' } : {}),
        ...Object.fromEntries(Object.entries(it.rowB || {}).filter(([k]) => !k.startsWith('_')))
      }));
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

  // Available custom column options from both uploaded files + audit metadata
  const allCustomColumnsList = useMemo(() => {
    const list = [
      { key: 'meta:status', label: 'Match Status', origin: 'meta', badge: 'Audit Status' },
      { key: 'meta:confidence', label: 'Confidence Score (%)', origin: 'meta', badge: 'Audit Score' },
      { key: 'meta:discrepancies', label: 'Discrepancy Details', origin: 'meta', badge: 'Discrepancy' }
    ];

    if (datasetAFiles.length > 1) {
      list.push({ key: 'meta:sourceA', label: 'Source File A', origin: 'meta', badge: 'File Origin' });
    }
    if (datasetBFiles.length > 1) {
      list.push({ key: 'meta:sourceB', label: 'Source File B', origin: 'meta', badge: 'File Origin' });
    }

    (datasetA.columns || []).forEach(c => {
      list.push({ key: `A:${c}`, label: c, originalName: c, origin: 'A', badge: `Dataset A: ${datasetA.name || 'File A'}` });
    });

    (datasetB.columns || []).forEach(c => {
      list.push({ key: `B:${c}`, label: c, originalName: c, origin: 'B', badge: `Dataset B: ${datasetB.name || 'File B'}` });
    });

    return list;
  }, [datasetA.columns, datasetA.name, datasetB.columns, datasetB.name, datasetAFiles.length, datasetBFiles.length]);

  // Current active selected column keys
  const effectiveSelectedCols = useMemo(() => {
    if (customSelectedCols !== null) return customSelectedCols;
    return allCustomColumnsList.map(c => c.key);
  }, [customSelectedCols, allCustomColumnsList]);

  const toggleCustomCol = (colKey) => {
    if (effectiveSelectedCols.includes(colKey)) {
      setCustomSelectedCols(effectiveSelectedCols.filter(k => k !== colKey));
    } else {
      setCustomSelectedCols([...effectiveSelectedCols, colKey]);
    }
  };

  const selectAllCustomCols = () => {
    setCustomSelectedCols(allCustomColumnsList.map(c => c.key));
  };

  const selectGroupCols = (groupKey) => {
    const groupColKeys = allCustomColumnsList.filter(c => c.origin === groupKey).map(c => c.key);
    const combined = Array.from(new Set([...effectiveSelectedCols, ...groupColKeys]));
    setCustomSelectedCols(combined);
  };

  const deselectGroupCols = (groupKey) => {
    setCustomSelectedCols(effectiveSelectedCols.filter(k => {
      const col = allCustomColumnsList.find(c => c.key === k);
      return col && col.origin !== groupKey;
    }));
  };

  const clearAllCustomCols = () => {
    setCustomSelectedCols([]);
  };

  // Generate Custom Output Rows based on selected columns & record scope
  const getCustomExportRows = () => {
    if (!comparisonResults) return [];

    let targetItems = [];
    if (customRecordScope === 'all') {
      targetItems = [
        ...comparisonResults.exactMatches,
        ...comparisonResults.partialMatches,
        ...comparisonResults.valueDiscrepancies,
        ...comparisonResults.unmatchedA,
        ...comparisonResults.unmatchedB
      ];
    } else if (customRecordScope === 'matched') {
      targetItems = [
        ...comparisonResults.exactMatches,
        ...comparisonResults.partialMatches
      ];
    } else if (customRecordScope === 'discrepancy') {
      targetItems = comparisonResults.valueDiscrepancies;
    } else if (customRecordScope === 'partial') {
      targetItems = comparisonResults.partialMatches;
    } else if (customRecordScope === 'unmatched_a') {
      targetItems = comparisonResults.unmatchedA;
    } else if (customRecordScope === 'unmatched_b') {
      targetItems = comparisonResults.unmatchedB;
    }

    return targetItems.map((item, idx) => {
      const row = { "Row_Index": idx + 1 };
      
      effectiveSelectedCols.forEach(k => {
        if (k === 'meta:status') {
          row['Match_Status'] = item.status || 'Unmatched';
        } else if (k === 'meta:confidence') {
          row['Confidence_%'] = item.confidence ?? '';
        } else if (k === 'meta:discrepancies') {
          row['Discrepancy_Details'] = (item.discrepancies && item.discrepancies.length > 0)
            ? item.discrepancies.map(d => `${d.fieldA}("${d.valA}" vs "${d.valB}")`).join('; ')
            : (item.status === 'Exact Match' ? 'None (100% Match)' : '');
        } else if (k === 'meta:sourceA') {
          row['Source_File_A'] = item.rowA ? (item.rowA._sourceFileName || '') : '';
        } else if (k === 'meta:sourceB') {
          row['Source_File_B'] = item.rowB ? (item.rowB._sourceFileName || '') : '';
        } else if (k.startsWith('A:')) {
          const col = k.slice(2);
          row[`A_${col}`] = item.rowA ? (item.rowA[col] ?? '') : '';
        } else if (k.startsWith('B:')) {
          const col = k.slice(2);
          row[`B_${col}`] = item.rowB ? (item.rowB[col] ?? '') : '';
        }
      });

      return row;
    });
  };

  const exportCustomExcel = () => {
    if (effectiveSelectedCols.length === 0) {
      alert("Please select at least one column from Dataset A, Dataset B, or Audit Metadata to include in your customized report.");
      return;
    }
    const rows = getCustomExportRows();
    if (rows.length === 0) {
      alert("No records found matching the selected record scope.");
      return;
    }

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    const fileName = `${customReportName.trim() || 'Custom_Reconciled_Report'}.xlsx`;
    XLSX.utils.book_append_sheet(wb, ws, "Custom_Report");
    XLSX.writeFile(wb, fileName);
    setStatus(`Downloaded custom report ${fileName} successfully!`, 'success');
  };

  const exportCustomCsv = () => {
    if (effectiveSelectedCols.length === 0) {
      alert("Please select at least one column to include in your customized report.");
      return;
    }
    const rows = getCustomExportRows();
    if (rows.length === 0) {
      alert("No records found matching the selected record scope.");
      return;
    }

    const ws = XLSX.utils.json_to_sheet(rows);
    const csvData = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
    const fileName = `${customReportName.trim() || 'Custom_Reconciled_Report'}.csv`;
    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(downloadUrl);
    setStatus(`Downloaded custom CSV report ${fileName} successfully!`, 'success');
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
                <div>
                  <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700 }}>Dataset A (Reference / Master)</h3>
                  {datasetAFiles.length > 0 && (
                    <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 500 }}>
                      {datasetAFiles.length} {datasetAFiles.length === 1 ? 'file' : 'files'} uploaded
                    </span>
                  )}
                </div>
              </div>
              {datasetA.rows.length > 0 && (
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--accent)', background: 'var(--accent-soft)', padding: '4px 10px', borderRadius: '12px' }}>
                  {datasetA.rows.length} Rows • {datasetA.columns.length} Cols • {datasetAFiles.length} {datasetAFiles.length === 1 ? 'File' : 'Files'}
                </span>
              )}
            </div>

            <p style={{ color: 'var(--muted)', fontSize: '13px', margin: '0 0 16px 0', lineHeight: '1.5' }}>
              Select the primary reference dataset (e.g. Master Admission Roster, Previous Term Data, Official Roll). Upload one or multiple Excel / CSV files.
            </p>

            <div 
              onDragOver={(e) => { e.preventDefault(); setIsDragOverA(true); }}
              onDragLeave={() => setIsDragOverA(false)}
              onDrop={(e) => {
                e.preventDefault();
                setIsDragOverA(false);
                handleUploadFilesA(e.dataTransfer.files);
              }}
              style={{ 
                border: `2px dashed ${isDragOverA ? 'var(--accent)' : 'var(--line)'}`, 
                borderRadius: '12px', 
                padding: '22px 16px', 
                textAlign: 'center', 
                background: isDragOverA ? 'var(--accent-soft)' : 'var(--bg)', 
                marginBottom: '16px',
                transition: 'all 0.15s ease'
              }}
            >
              <Upload size={30} color="var(--accent)" style={{ margin: '0 auto 8px', opacity: 0.8 }} />
              <div style={{ fontWeight: 600, fontSize: '13.5px', marginBottom: '4px' }}>
                {datasetAFiles.length === 0 
                  ? 'Upload Excel (.xlsx, .xls), CSV, or ZIP' 
                  : `Add More Files (${datasetAFiles.length} currently uploaded)`}
              </div>
              <p style={{ color: 'var(--muted)', fontSize: '11.5px', margin: '0 0 12px 0' }}>
                Select multiple files at once or drop a ZIP archive of spreadsheets
              </p>
              <label className="button" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', cursor: 'pointer', padding: '7px 16px', fontSize: '12.5px' }}>
                <Plus size={14} /> Browse File(s) A
                <input type="file" accept=".xlsx, .xls, .csv, .zip" multiple onChange={handleUploadA} style={{ display: 'none' }} />
              </label>
            </div>

            {/* Uploaded Files Chips for Dataset A */}
            {datasetAFiles.length > 0 && (
              <div style={fileListStyles.filesListContainer}>
                <div style={fileListStyles.filesListHeader}>
                  <span style={fileListStyles.filesListTitle}>
                    Uploaded Reference Files ({datasetAFiles.length})
                  </span>
                  <button
                    type="button"
                    style={fileListStyles.clearAllBtn}
                    onClick={handleClearFilesA}
                    title="Clear all uploaded Dataset A files"
                  >
                    Clear All
                  </button>
                </div>
                <div style={fileListStyles.filesScrollList}>
                  {datasetAFiles.map((f) => (
                    <div key={f.id} style={fileListStyles.fileChip}>
                      <div style={fileListStyles.fileChipLeft}>
                        <FileSpreadsheet size={15} color="var(--accent)" style={{ flexShrink: 0 }} />
                        <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
                          <span style={fileListStyles.fileChipName} title={f.name}>
                            {f.name}
                          </span>
                          <span style={fileListStyles.fileChipMeta}>
                            {f.sizeFormatted} • {f.rows.length} rows • {f.columns.length} cols
                          </span>
                        </div>
                      </div>
                      <div style={fileListStyles.fileChipRight}>
                        {f.sheetNames.length > 1 && (
                          <select
                            value={f.selectedSheet}
                            onChange={(e) => handleSheetChangeFileA(f.id, e.target.value)}
                            style={fileListStyles.sheetMiniSelect}
                            title="Select active worksheet for this file"
                          >
                            {f.sheetNames.map((s) => (
                              <option key={s} value={s}>{s}</option>
                            ))}
                          </select>
                        )}
                        <button
                          type="button"
                          style={fileListStyles.fileRemoveBtn}
                          onClick={() => handleRemoveFileA(f.id)}
                          title={`Remove ${f.name}`}
                        >
                          <Trash2 size={13} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {datasetA.columns.length > 0 && (
              <div>
                <div style={{ fontSize: '12px', fontWeight: 700, textTransform: 'uppercase', color: 'var(--muted)', marginBottom: '8px' }}>
                  Unified Columns ({datasetA.columns.length})
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
                        {datasetAFiles.length > 1 && (
                          <th style={{ padding: '6px 8px', textAlign: 'left', color: 'var(--accent)' }}>Source File</th>
                        )}
                        {datasetA.columns.slice(0, 5).map(c => <th key={c} style={{ padding: '6px 8px', textAlign: 'left' }}>{c}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {datasetA.rows.slice(0, 4).map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--line)' }}>
                          {datasetAFiles.length > 1 && (
                            <td style={{ padding: '6px 8px', whiteSpace: 'nowrap', color: 'var(--accent)', fontWeight: 600 }}>
                              {r._sourceFileName || '—'}
                            </td>
                          )}
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
                <div>
                  <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700 }}>Dataset B (Comparison / Incoming)</h3>
                  {datasetBFiles.length > 0 && (
                    <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 500 }}>
                      {datasetBFiles.length} {datasetBFiles.length === 1 ? 'file' : 'files'} uploaded
                    </span>
                  )}
                </div>
              </div>
              {datasetB.rows.length > 0 && (
                <span style={{ fontSize: '12px', fontWeight: 700, color: 'var(--accent)', background: 'var(--accent-soft)', padding: '4px 10px', borderRadius: '12px' }}>
                  {datasetB.rows.length} Rows • {datasetB.columns.length} Cols • {datasetBFiles.length} {datasetBFiles.length === 1 ? 'File' : 'Files'}
                </span>
              )}
            </div>

            <p style={{ color: 'var(--muted)', fontSize: '13px', margin: '0 0 16px 0', lineHeight: '1.5' }}>
              Select the comparison dataset to reconcile against Dataset A (e.g. Exam Registrations, Fee Receipts, Attendance). Upload one or multiple Excel / CSV files.
            </p>

            <div 
              onDragOver={(e) => { e.preventDefault(); setIsDragOverB(true); }}
              onDragLeave={() => setIsDragOverB(false)}
              onDrop={(e) => {
                e.preventDefault();
                setIsDragOverB(false);
                handleUploadFilesB(e.dataTransfer.files);
              }}
              style={{ 
                border: `2px dashed ${isDragOverB ? 'var(--accent)' : 'var(--line)'}`, 
                borderRadius: '12px', 
                padding: '22px 16px', 
                textAlign: 'center', 
                background: isDragOverB ? 'var(--accent-soft)' : 'var(--bg)', 
                marginBottom: '16px',
                transition: 'all 0.15s ease'
              }}
            >
              <Upload size={30} color="var(--accent)" style={{ margin: '0 auto 8px', opacity: 0.8 }} />
              <div style={{ fontWeight: 600, fontSize: '13.5px', marginBottom: '4px' }}>
                {datasetBFiles.length === 0 
                  ? 'Upload Excel (.xlsx, .xls), CSV, or ZIP' 
                  : `Add More Files (${datasetBFiles.length} currently uploaded)`}
              </div>
              <p style={{ color: 'var(--muted)', fontSize: '11.5px', margin: '0 0 12px 0' }}>
                Select multiple files at once or drop a ZIP archive of spreadsheets
              </p>
              <label className="button" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', cursor: 'pointer', padding: '7px 16px', fontSize: '12.5px' }}>
                <Plus size={14} /> Browse File(s) B
                <input type="file" accept=".xlsx, .xls, .csv, .zip" multiple onChange={handleUploadB} style={{ display: 'none' }} />
              </label>
            </div>

            {/* Uploaded Files Chips for Dataset B */}
            {datasetBFiles.length > 0 && (
              <div style={fileListStyles.filesListContainer}>
                <div style={fileListStyles.filesListHeader}>
                  <span style={fileListStyles.filesListTitle}>
                    Uploaded Comparison Files ({datasetBFiles.length})
                  </span>
                  <button
                    type="button"
                    style={fileListStyles.clearAllBtn}
                    onClick={handleClearFilesB}
                    title="Clear all uploaded Dataset B files"
                  >
                    Clear All
                  </button>
                </div>
                <div style={fileListStyles.filesScrollList}>
                  {datasetBFiles.map((f) => (
                    <div key={f.id} style={fileListStyles.fileChip}>
                      <div style={fileListStyles.fileChipLeft}>
                        <FileSpreadsheet size={15} color="var(--accent)" style={{ flexShrink: 0 }} />
                        <div style={{ display: 'flex', flexDirection: 'column', minWidth: 0 }}>
                          <span style={fileListStyles.fileChipName} title={f.name}>
                            {f.name}
                          </span>
                          <span style={fileListStyles.fileChipMeta}>
                            {f.sizeFormatted} • {f.rows.length} rows • {f.columns.length} cols
                          </span>
                        </div>
                      </div>
                      <div style={fileListStyles.fileChipRight}>
                        {f.sheetNames.length > 1 && (
                          <select
                            value={f.selectedSheet}
                            onChange={(e) => handleSheetChangeFileB(f.id, e.target.value)}
                            style={fileListStyles.sheetMiniSelect}
                            title="Select active worksheet for this file"
                          >
                            {f.sheetNames.map((s) => (
                              <option key={s} value={s}>{s}</option>
                            ))}
                          </select>
                        )}
                        <button
                          type="button"
                          style={fileListStyles.fileRemoveBtn}
                          onClick={() => handleRemoveFileB(f.id)}
                          title={`Remove ${f.name}`}
                        >
                          <Trash2 size={13} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {datasetB.columns.length > 0 && (
              <div>
                <div style={{ fontSize: '12px', fontWeight: 700, textTransform: 'uppercase', color: 'var(--muted)', marginBottom: '8px' }}>
                  Unified Columns ({datasetB.columns.length})
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
                        {datasetBFiles.length > 1 && (
                          <th style={{ padding: '6px 8px', textAlign: 'left', color: 'var(--accent)' }}>Source File</th>
                        )}
                        {datasetB.columns.slice(0, 5).map(c => <th key={c} style={{ padding: '6px 8px', textAlign: 'left' }}>{c}</th>)}
                      </tr>
                    </thead>
                    <tbody>
                      {datasetB.rows.slice(0, 4).map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--line)' }}>
                          {datasetBFiles.length > 1 && (
                            <td style={{ padding: '6px 8px', whiteSpace: 'nowrap', color: 'var(--accent)', fontWeight: 600 }}>
                              {r._sourceFileName || '—'}
                            </td>
                          )}
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
          <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginTop: '8px' }}>
            {isProcessing && (
              <div style={{ background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '8px', padding: '14px 18px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '13px', fontWeight: 700, marginBottom: '8px', color: 'var(--ink)' }}>
                  <span>Reconciling datasets ({((datasetA?.rows?.length) || 0).toLocaleString()} vs {((datasetB?.rows?.length) || 0).toLocaleString()} rows)...</span>
                  <span style={{ color: 'var(--accent)', fontWeight: 800 }}>{calculationProgress}%</span>
                </div>
                <div style={{ width: '100%', height: '8px', background: 'var(--line)', borderRadius: '4px', overflow: 'hidden' }}>
                  <div style={{ width: `${calculationProgress}%`, height: '100%', background: 'linear-gradient(90deg, var(--accent) 0%, #10b981 100%)', transition: 'width 0.2s ease' }} />
                </div>
                <div style={{ fontSize: '11.5px', color: 'var(--muted)', marginTop: '6px' }}>
                  Running ultra-fast composite-index comparison with sub-second memory optimization.
                </div>
              </div>
            )}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <button className="secondary" onClick={() => setCurrentStep(1)} disabled={isProcessing}>
                ← Back to Upload
              </button>
              <button 
                disabled={isProcessing}
                onClick={runComparisonEngine}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '12px 32px', fontSize: '15px', fontWeight: 800 }}
              >
                {isProcessing ? `Reconciling Records (${calculationProgress}%)...` : '⚡ Execute Comparison Engine →'}
              </button>
            </div>
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
                  onClick={() => handleTabSwitch(tab.id)}
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
                    pagedResultItems.map((item, idx) => {
                      const globalIdx = (safePage - 1) * resultPageSize + idx + 1;
                      const isExact = item.status === 'Exact Match';
                      const isPartial = item.status === 'Partial Match';
                      const isDisc = item.status === 'Value Discrepancy';
                      const isUnA = item.status?.includes('Dataset A Only');
                      const isUnB = item.status?.includes('Dataset B Only');

                      return (
                        <tr key={item.id || globalIdx} style={{ borderBottom: '1px solid var(--line)', background: idx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.015)' }}>
                          <td style={{ padding: '12px 14px', textAlign: 'center', fontWeight: 600, color: 'var(--muted)' }}>
                            {globalIdx}
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
                                {datasetAFiles.length > 1 && item.rowA._sourceFileName && (
                                  <div style={{ fontSize: '10.5px', background: 'var(--panel)', border: '1px solid var(--line)', padding: '2px 6px', borderRadius: '4px', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '4px', marginBottom: '6px' }}>
                                    <FileSpreadsheet size={11} /> {item.rowA._sourceFileName}
                                  </div>
                                )}
                                {Object.entries(item.rowA).filter(([k]) => !k.startsWith('_')).slice(0, 4).map(([k, v]) => (
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
                                {datasetBFiles.length > 1 && item.rowB._sourceFileName && (
                                  <div style={{ fontSize: '10.5px', background: 'var(--panel)', border: '1px solid var(--line)', padding: '2px 6px', borderRadius: '4px', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '4px', marginBottom: '6px' }}>
                                    <FileSpreadsheet size={11} /> {item.rowB._sourceFileName}
                                  </div>
                                )}
                                {Object.entries(item.rowB).filter(([k]) => !k.startsWith('_')).slice(0, 4).map(([k, v]) => (
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

            {/* Table Footer with Pagination Controls */}
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
                  Showing <strong>{totalResultCount === 0 ? 0 : (safePage - 1) * resultPageSize + 1}</strong> to <strong>{Math.min(safePage * resultPageSize, totalResultCount)}</strong> of <strong>{(totalResultCount || 0).toLocaleString()}</strong> record(s)
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <span>Rows per page:</span>
                  <select 
                    value={resultPageSize} 
                    onChange={(e) => { 
                      setResultPageSize(Number(e.target.value)); 
                      setResultPage(1); 
                    }}
                    style={{ padding: '4px 8px', fontSize: '12px', borderRadius: '4px' }}
                  >
                    <option value={25}>25</option>
                    <option value={50}>50</option>
                    <option value={100}>100</option>
                    <option value={250}>250</option>
                    <option value={500}>500</option>
                  </select>
                </div>
              </div>

              {totalResultPages > 1 && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <button 
                    className="secondary" 
                    onClick={() => setResultPage(1)} 
                    disabled={safePage <= 1}
                    style={{ padding: '4px 10px', fontSize: '12px' }}
                    title="First Page"
                  >
                    « First
                  </button>

                  <button 
                    className="secondary" 
                    onClick={() => setResultPage(p => Math.max(1, p - 1))} 
                    disabled={safePage <= 1}
                    style={{ padding: '4px 10px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '2px' }}
                  >
                    <ChevronLeft size={14} /> Prev
                  </button>

                  <span style={{ margin: '0 8px', fontSize: '12.5px', fontWeight: 600, color: 'var(--ink)' }}>
                    Page {safePage} of {totalResultPages}
                  </span>

                  <button 
                    className="secondary" 
                    onClick={() => setResultPage(p => Math.min(totalResultPages, p + 1))} 
                    disabled={safePage >= totalResultPages}
                    style={{ padding: '4px 10px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '2px' }}
                  >
                    Next <ChevronRight size={14} />
                  </button>

                  <button 
                    className="secondary" 
                    onClick={() => setResultPage(totalResultPages)} 
                    disabled={safePage >= totalResultPages}
                    style={{ padding: '4px 10px', fontSize: '12px' }}
                    title="Last Page"
                  >
                    Last »
                  </button>
                </div>
              )}
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

          {/* Custom Column & Master Merged Report Studio */}
          <div className="card" style={{ padding: '32px', border: '1.5px solid var(--accent)', background: 'var(--panel)' }}>
            
            {/* Header */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', flexWrap: 'wrap', gap: '16px', marginBottom: '24px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
                <div style={{ background: 'linear-gradient(135deg, var(--accent), #5cbbd4)', color: 'white', padding: '12px', borderRadius: '12px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                  <Sliders size={26} />
                </div>
                <div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <h3 style={{ margin: 0, fontSize: '20px', fontWeight: 800 }}>Custom Column & Master Merged Report Builder</h3>
                    <span style={{ fontSize: '11px', fontWeight: 800, padding: '3px 8px', borderRadius: '12px', background: 'rgba(23, 107, 135, 0.12)', color: 'var(--accent)' }}>
                      CUSTOMIZABLE OUTPUT
                    </span>
                  </div>
                  <p style={{ color: 'var(--muted)', fontSize: '13.5px', margin: '4px 0 0 0' }}>
                    Select and customize any headers from <strong>Dataset A</strong> ({datasetA.name}), <strong>Dataset B</strong> ({datasetB.name}), and <strong>Audit Status Flags</strong> to build your bespoke spreadsheet.
                  </p>
                </div>
              </div>
            </div>

            {/* Config Controls Row: Report Name & Record Scope */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '18px', background: 'var(--bg)', padding: '20px', borderRadius: '10px', border: '1px solid var(--line)', marginBottom: '24px' }}>
              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                  Custom Report File Name:
                </label>
                <input 
                  type="text" 
                  value={customReportName} 
                  onChange={(e) => setCustomReportName(e.target.value)}
                  placeholder="Custom_Reconciled_Report"
                  style={{ width: '100%', padding: '9px 12px', fontSize: '13.5px' }}
                />
              </div>

              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                  Record Scope Filter:
                </label>
                <select 
                  value={customRecordScope} 
                  onChange={(e) => setCustomRecordScope(e.target.value)}
                  style={{ width: '100%', padding: '9px 12px', fontSize: '13.5px' }}
                >
                  <option value="all">All Reconciled Records (Full Merged Outer View)</option>
                  <option value="matched">Matched Records Only (Exact + Fuzzy)</option>
                  <option value="discrepancy">Value Discrepancies Only</option>
                  <option value="partial">Partial / Fuzzy Matches Only</option>
                  <option value="unmatched_a">Unmatched Dataset A Only</option>
                  <option value="unmatched_b">Unmatched Dataset B Only</option>
                </select>
              </div>
            </div>

            {/* Column Presets / Quick Selection Bar */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px', marginBottom: '18px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Quick Selectors:
                </span>
                <button className="secondary" onClick={selectAllCustomCols} style={{ padding: '5px 10px', fontSize: '12px' }}>
                  Select All Columns
                </button>
                <button className="secondary" onClick={() => selectGroupCols('A')} style={{ padding: '5px 10px', fontSize: '12px' }}>
                  All Dataset A ({datasetA.columns.length})
                </button>
                <button className="secondary" onClick={() => selectGroupCols('B')} style={{ padding: '5px 10px', fontSize: '12px' }}>
                  All Dataset B ({datasetB.columns.length})
                </button>
                <button className="secondary" onClick={() => selectGroupCols('meta')} style={{ padding: '5px 10px', fontSize: '12px' }}>
                  Audit Metadata (3)
                </button>
                <button className="secondary" onClick={clearAllCustomCols} style={{ padding: '5px 10px', fontSize: '12px', color: 'var(--danger)' }}>
                  Deselect All
                </button>
              </div>

              <div style={{ fontSize: '13px', fontWeight: 700, color: 'var(--accent)' }}>
                {effectiveSelectedCols.length} of {allCustomColumnsList.length} Headers Selected
              </div>
            </div>

            {/* Categorized Column Selection Deck */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '20px', marginBottom: '24px' }}>
              
              {/* Group 1: Dataset A Columns */}
              <div style={{ background: 'var(--bg)', borderRadius: '10px', padding: '16px', border: '1.5px solid rgba(23, 107, 135, 0.25)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', paddingBottom: '8px', borderBottom: '1px solid var(--line)' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <div style={{ background: 'var(--accent)', color: 'white', width: '20px', height: '20px', borderRadius: '50%', fontSize: '11px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800 }}>A</div>
                    <strong style={{ fontSize: '13px' }}>Dataset A Headers ({datasetA.columns.length})</strong>
                  </div>
                  <div style={{ display: 'flex', gap: '4px' }}>
                    <button className="secondary" onClick={() => selectGroupCols('A')} style={{ padding: '2px 6px', fontSize: '11px' }}>All</button>
                    <button className="secondary" onClick={() => deselectGroupCols('A')} style={{ padding: '2px 6px', fontSize: '11px' }}>None</button>
                  </div>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '220px', overflowY: 'auto' }}>
                  {datasetA.columns.map(c => {
                    const colKey = `A:${c}`;
                    const isChecked = effectiveSelectedCols.includes(colKey);
                    return (
                      <label 
                        key={colKey}
                        onClick={() => toggleCustomCol(colKey)}
                        style={{ 
                          display: 'flex', 
                          alignItems: 'center', 
                          gap: '10px', 
                          padding: '7px 10px', 
                          borderRadius: '6px', 
                          fontSize: '12.5px', 
                          cursor: 'pointer',
                          background: isChecked ? 'rgba(23, 107, 135, 0.1)' : 'var(--panel)',
                          border: isChecked ? '1px solid var(--accent)' : '1px solid var(--line)',
                          fontWeight: isChecked ? 600 : 400
                        }}
                      >
                        <input 
                          type="checkbox" 
                          checked={isChecked} 
                          onChange={() => {}} 
                          style={{ cursor: 'pointer' }}
                        />
                        <span style={{ flex: 1, wordBreak: 'break-word' }}>{c}</span>
                        <span style={{ fontSize: '10px', color: 'var(--muted)', background: 'var(--bg)', padding: '2px 6px', borderRadius: '4px' }}>A</span>
                      </label>
                    );
                  })}
                </div>
              </div>

              {/* Group 2: Dataset B Columns */}
              <div style={{ background: 'var(--bg)', borderRadius: '10px', padding: '16px', border: '1.5px solid rgba(139, 92, 246, 0.25)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', paddingBottom: '8px', borderBottom: '1px solid var(--line)' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <div style={{ background: '#8b5cf6', color: 'white', width: '20px', height: '20px', borderRadius: '50%', fontSize: '11px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800 }}>B</div>
                    <strong style={{ fontSize: '13px' }}>Dataset B Headers ({datasetB.columns.length})</strong>
                  </div>
                  <div style={{ display: 'flex', gap: '4px' }}>
                    <button className="secondary" onClick={() => selectGroupCols('B')} style={{ padding: '2px 6px', fontSize: '11px' }}>All</button>
                    <button className="secondary" onClick={() => deselectGroupCols('B')} style={{ padding: '2px 6px', fontSize: '11px' }}>None</button>
                  </div>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '220px', overflowY: 'auto' }}>
                  {datasetB.columns.map(c => {
                    const colKey = `B:${c}`;
                    const isChecked = effectiveSelectedCols.includes(colKey);
                    return (
                      <label 
                        key={colKey}
                        onClick={() => toggleCustomCol(colKey)}
                        style={{ 
                          display: 'flex', 
                          alignItems: 'center', 
                          gap: '10px', 
                          padding: '7px 10px', 
                          borderRadius: '6px', 
                          fontSize: '12.5px', 
                          cursor: 'pointer',
                          background: isChecked ? 'rgba(139, 92, 246, 0.1)' : 'var(--panel)',
                          border: isChecked ? '1px solid #8b5cf6' : '1px solid var(--line)',
                          fontWeight: isChecked ? 600 : 400
                        }}
                      >
                        <input 
                          type="checkbox" 
                          checked={isChecked} 
                          onChange={() => {}} 
                          style={{ cursor: 'pointer' }}
                        />
                        <span style={{ flex: 1, wordBreak: 'break-word' }}>{c}</span>
                        <span style={{ fontSize: '10px', color: 'var(--muted)', background: 'var(--bg)', padding: '2px 6px', borderRadius: '4px' }}>B</span>
                      </label>
                    );
                  })}
                </div>
              </div>

              {/* Group 3: Audit Metadata Columns */}
              <div style={{ background: 'var(--bg)', borderRadius: '10px', padding: '16px', border: '1.5px solid rgba(16, 185, 129, 0.25)' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', paddingBottom: '8px', borderBottom: '1px solid var(--line)' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <div style={{ background: '#10b981', color: 'white', width: '20px', height: '20px', borderRadius: '50%', fontSize: '11px', display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 800 }}>✓</div>
                    <strong style={{ fontSize: '13px' }}>Audit & Reconciled Metadata (3)</strong>
                  </div>
                  <div style={{ display: 'flex', gap: '4px' }}>
                    <button className="secondary" onClick={() => selectGroupCols('meta')} style={{ padding: '2px 6px', fontSize: '11px' }}>All</button>
                    <button className="secondary" onClick={() => deselectGroupCols('meta')} style={{ padding: '2px 6px', fontSize: '11px' }}>None</button>
                  </div>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '220px', overflowY: 'auto' }}>
                  {[
                    { key: 'meta:status', label: 'Match Status (Exact / Partial / Discrepancy / Unmatched)', sub: 'Reconciled classification badge' },
                    { key: 'meta:confidence', label: 'Confidence Score (%)', sub: 'Calculated fuzzy similarity percentage' },
                    { key: 'meta:discrepancies', label: 'Discrepancy Details', sub: 'Summary of mismatched attributes' }
                  ].map(m => {
                    const isChecked = effectiveSelectedCols.includes(m.key);
                    return (
                      <label 
                        key={m.key}
                        onClick={() => toggleCustomCol(m.key)}
                        style={{ 
                          display: 'flex', 
                          alignItems: 'flex-start', 
                          gap: '10px', 
                          padding: '7px 10px', 
                          borderRadius: '6px', 
                          fontSize: '12.5px', 
                          cursor: 'pointer',
                          background: isChecked ? 'rgba(16, 185, 129, 0.1)' : 'var(--panel)',
                          border: isChecked ? '1px solid #10b981' : '1px solid var(--line)',
                          fontWeight: isChecked ? 600 : 400
                        }}
                      >
                        <input 
                          type="checkbox" 
                          checked={isChecked} 
                          onChange={() => {}} 
                          style={{ cursor: 'pointer', marginTop: '3px' }}
                        />
                        <div style={{ flex: 1 }}>
                          <div>{m.label}</div>
                          <div style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 400 }}>{m.sub}</div>
                        </div>
                      </label>
                    );
                  })}
                </div>
              </div>

            </div>

            {/* Custom Output Action Buttons */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '16px', paddingTop: '20px', borderTop: '1px solid var(--line)' }}>
              <div style={{ fontSize: '13.5px', color: 'var(--muted)' }}>
                Ready to generate custom report with <strong>{effectiveSelectedCols.length} column(s)</strong>.
              </div>

              <div style={{ display: 'flex', gap: '12px', flexWrap: 'wrap' }}>
                <button 
                  onClick={exportCustomExcel}
                  disabled={effectiveSelectedCols.length === 0}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '12px 24px', fontSize: '14px', fontWeight: 700 }}
                >
                  <Download size={18} /> Download Custom Excel (.xlsx)
                </button>

                <button 
                  className="secondary"
                  onClick={exportCustomCsv}
                  disabled={effectiveSelectedCols.length === 0}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '12px 20px', fontSize: '14px', fontWeight: 700 }}
                >
                  <FileText size={18} /> Download Custom CSV (.csv)
                </button>
              </div>
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
                  {inspectModalItem.rowA?._sourceFileName && (
                    <div style={{ padding: '6px 8px', background: 'var(--panel)', borderRadius: '4px', border: '1px solid var(--accent)', color: 'var(--accent)', fontWeight: 600 }}>
                      📄 Source File: {inspectModalItem.rowA._sourceFileName}
                    </div>
                  )}
                  {Object.entries(inspectModalItem.rowA || {}).filter(([k]) => !k.startsWith('_')).map(([k, v]) => (
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
                  {inspectModalItem.rowB?._sourceFileName && (
                    <div style={{ padding: '6px 8px', background: 'var(--panel)', borderRadius: '4px', border: '1px solid var(--accent)', color: 'var(--accent)', fontWeight: 600 }}>
                      📄 Source File: {inspectModalItem.rowB._sourceFileName}
                    </div>
                  )}
                  {Object.entries(inspectModalItem.rowB || {}).filter(([k]) => !k.startsWith('_')).map(([k, v]) => (
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
