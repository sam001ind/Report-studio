import { useState, useMemo } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { 
  ArrowLeft, 
  Upload, 
  Download, 
  FileSpreadsheet, 
  AlertTriangle, 
  Search, 
  Sparkles,
  Database,
  SlidersHorizontal
} from 'lucide-react';
import { 
  TARGET_COLUMNS, 
  CATEGORY_LOOKUP, 
  CASTE_LOOKUP, 
  RELIGION_LOOKUP, 
  PROGRAM_LOOKUP, 
  KANNUR_PROGRAMMES,
  COLLEGE_LOOKUP,
  STREAM_LOOKUP,
  QUALIFICATION_LOOKUP,
  DEFAULT_LOOKUP_DATA,
  loadSavedLookups,
  buildLookupMaps,
  normalizeLookupKey
} from '../data/admissionLookups';
import LookupIdManagerModal from '../components/LookupIdManagerModal';

const normalizeKey = (key) => String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');

const cleanEscapeArtifacts = (val) => {
  if (val === null || val === undefined) return '';
  const str = String(val).trim();
  if (str === '\\N' || str.toLowerCase() === 'null' || str.toLowerCase() === 'undefined' || str.toLowerCase() === 'nan') {
    return '';
  }
  return str;
};

// Formats date to YYYY-MM-DD
export const parseDobYMD = (val) => {
  const clean = cleanEscapeArtifacts(val);
  if (!clean) return '';
  if (/^\d{4,5}$/.test(clean)) {
    const serial = parseInt(clean, 10);
    if (serial > 10000 && serial < 60000) {
      const d = new Date((serial - 25569) * 86400 * 1000);
      if (!isNaN(d.getTime())) {
        const y = d.getUTCFullYear();
        const m = String(d.getUTCMonth() + 1).padStart(2, '0');
        const day = String(d.getUTCDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
      }
    }
  }
  // Check YY-MM-DD (e.g. "08-12-22" from dump)
  const yyMatch = clean.match(/^(\d{2})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (yyMatch) {
    let year = parseInt(yyMatch[1], 10);
    year = year > 50 ? 1900 + year : 2000 + year;
    const m = String(parseInt(yyMatch[2], 10)).padStart(2, '0');
    const d = String(parseInt(yyMatch[3], 10)).padStart(2, '0');
    return `${year}-${m}-${d}`;
  }
  // Check DD-MM-YYYY or DD/MM/YYYY
  const dmyMatch = clean.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (dmyMatch) {
    const d = String(parseInt(dmyMatch[1], 10)).padStart(2, '0');
    const m = String(parseInt(dmyMatch[2], 10)).padStart(2, '0');
    const y = dmyMatch[3];
    return `${y}-${m}-${d}`;
  }
  // Standard JS Date
  const d = new Date(clean);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  return clean;
};

// Formats date to DD/MM/YYYY
export const parseDobDMY = (val) => {
  const ymd = parseDobYMD(val);
  if (!ymd) return '';
  const parts = ymd.split('-');
  if (parts.length === 3) {
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  return ymd;
};

// Admission allotment header recognition weights
export const ADMISSION_HEADER_WEIGHTS = [
  { keys: ['name', 'candidatename', 'studentname', 'fullname', 'applicantname', 'nameofthecandidate'], weight: 12, label: 'Name' },
  { keys: ['appno', 'applicationno', 'applno', 'applicationnumber', 'regno', 'registerno', 'registrationno', 'prn', 'candidateid', 'username', 'slno', 'sno', 'rollno', 'seatno', 'admnum'], weight: 10, label: 'Identifier / SlNo' },
  { keys: ['gender', 'sex'], weight: 5, label: 'Gender' },
  { keys: ['dob', 'dateofbirth', 'birthdate'], weight: 5, label: 'DOB' },
  { keys: ['category', 'community', 'allotmentquota', 'quota', 'allottedcategory', 'caste', 'religion', 'actual_category'], weight: 8, label: 'Category' },
  { keys: ['mobile', 'mobilenumber', 'phone', 'contact', 'contactnumber'], weight: 8, label: 'Mobile' },
  { keys: ['email', 'emailid', 'mail'], weight: 6, label: 'Email' },
  { keys: ['marks', 'marksobtained', 'indexmark', 'totalmarks', 'percentage', 'percent', 'cgpa', 'rank', 'allotmentrank', 'securedmarks', 'marksoutof'], weight: 8, label: 'Marks' },
  { keys: ['course', 'coursename', 'program', 'programname', 'programme', 'degree', 'subject', 'stream'], weight: 8, label: 'Course/Program' },
  { keys: ['college', 'collegename', 'institution', 'center', 'centre', 'collegecode', 'centercode', 'colcode'], weight: 6, label: 'College' },
  { keys: ['aadhaar', 'aadhaarnumber', 'adhar', 'abcid', 'certificatenumber', 'qcno', 'qcboard', 'qcyear', 'admit_date'], weight: 8, label: 'Student Master Dump Identifiers' }
];

export const scoreAdmissionHeaderRow = (rowArray) => {
  if (!Array.isArray(rowArray) || rowArray.length === 0) {
    return { score: -100, matchedHeaders: [], nonEmptyCount: 0 };
  }

  const nonEmpty = rowArray.filter(c => c !== undefined && c !== null && String(c).trim() !== '');
  if (nonEmpty.length < 2) {
    return { score: -100, matchedHeaders: [], nonEmptyCount: nonEmpty.length };
  }

  const normCols = rowArray.map(c => normalizeKey(c));
  let score = 0;
  const matchedHeaders = [];
  const matchedCategories = new Set();

  for (const item of ADMISSION_HEADER_WEIGHTS) {
    const hasMatch = normCols.some(col => item.keys.some(k => col === k || (col.length >= 4 && col.includes(k))));
    if (hasMatch) {
      score += item.weight;
      matchedHeaders.push(item.label);
      matchedCategories.add(item.label);
    }
  }

  // Bonus for candidate data density
  if (matchedCategories.size >= 3) {
    score += matchedCategories.size * 4;
  }

  // Penalty if row looks like an institutional banner / title line rather than table headers
  const joinedText = rowArray.map(c => String(c || '').toLowerCase()).join(' ');
  if (/university|centralized admission|allotment process|government of|provisional allotment|notification dated/i.test(joinedText)) {
    score -= 15;
  }

  return { score, matchedHeaders, nonEmptyCount: nonEmpty.length };
};

export const extractAdmissionRowsFromSheet = (sheet) => {
  if (!sheet) return { rows: [], headers: [], score: 0, headerRowIdx: 0, matchedHeaders: [] };

  const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
  if (!aoa || aoa.length === 0) {
    return { rows: [], headers: [], score: 0, headerRowIdx: 0, matchedHeaders: [] };
  }

  // Scan top 30 rows
  let bestRowIdx = -1;
  let maxScore = -99;
  let bestMatchedHeaders = [];

  const maxScan = Math.min(aoa.length, 30);
  for (let r = 0; r < maxScan; r++) {
    const rowArr = aoa[r];
    const { score, matchedHeaders } = scoreAdmissionHeaderRow(rowArr);
    if (score > maxScore) {
      maxScore = score;
      bestRowIdx = r;
      bestMatchedHeaders = matchedHeaders;
    }
  }

  // If no good header match was found (> 0), fallback to row with the most columns
  if (maxScore <= 0 || bestRowIdx === -1) {
    let maxCols = 0;
    bestRowIdx = 0;
    for (let r = 0; r < Math.min(aoa.length, 15); r++) {
      const nonEmpty = (aoa[r] || []).filter(c => String(c ?? '').trim() !== '').length;
      if (nonEmpty > maxCols) {
        maxCols = nonEmpty;
        bestRowIdx = r;
      }
    }
  }

  const rawHeaders = aoa[bestRowIdx] || [];
  const headers = [];
  const usedKeys = new Set();

  rawHeaders.forEach((colName, idx) => {
    let name = String(colName || '').trim();
    if (!name) name = `Column_${idx + 1}`;
    let unique = name;
    let counter = 2;
    while (usedKeys.has(unique.toLowerCase())) {
      unique = `${name}_${counter}`;
      counter++;
    }
    usedKeys.add(unique.toLowerCase());
    headers.push({ idx, name: unique, norm: normalizeKey(unique) });
  });

  const rows = [];
  for (let r = bestRowIdx + 1; r < aoa.length; r++) {
    const rowArr = aoa[r];
    if (!Array.isArray(rowArr)) continue;
    
    // Check if row has at least one non-empty value
    const nonEmptyCells = rowArr.filter(c => c !== undefined && c !== null && String(c).trim() !== '');
    if (nonEmptyCells.length === 0) continue;

    // Filter out footer rows, page numbers, signatures, or print timestamps
    const rowStr = nonEmptyCells.map(c => String(c).trim()).join(' ').toLowerCase();
    const isSummaryOrFooter = (nonEmptyCells.length <= 3 && /^(total|page \d|report generated|printed on|signature|disclaimer|coordinator|principal|verified by|date:|note:)/i.test(rowStr));
    if (isSummaryOrFooter) continue;

    const rowObj = {};
    headers.forEach(({ idx, name }) => {
      rowObj[name] = rowArr[idx] !== undefined && rowArr[idx] !== null ? String(rowArr[idx]).trim() : '';
    });
    rows.push(rowObj);
  }

  return {
    rows,
    headers: headers.map(h => h.name),
    score: Math.max(maxScore, 0),
    headerRowIdx: bestRowIdx,
    matchedHeaders: bestMatchedHeaders
  };
};

export const parseWorkbookFromBytes = (buffer) => {
  const uint8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  // 1. Try binary Excel
  try {
    return XLSX.read(uint8, { type: 'array', cellDates: true, raw: false });
  } catch {
    // 2. Try UTF-8 string (HTML tables, CSV, TSV)
    try {
      const text = new TextDecoder('utf-8').decode(uint8);
      return XLSX.read(text, { type: 'string', cellDates: true, raw: false });
    } catch {
      // 3. Try Windows-1252 string
      try {
        const text = new TextDecoder('windows-1252').decode(uint8);
        return XLSX.read(text, { type: 'string', cellDates: true, raw: false });
      } catch (err) {
        throw new Error(`Failed to read spreadsheet buffer: ${err.message}`, { cause: err });
      }
    }
  }
};

export const parseWorkbookBuffer = async (buffer, fileName = '') => {
  const isZip = String(fileName).toLowerCase().endsWith('.zip');
  if (isZip) {
    const zip = await JSZip.loadAsync(buffer);
    const validEntries = Object.keys(zip.files).filter(name => {
      if (!name || zip.files[name].dir) return false;
      if (name.includes('__MACOSX') || name.includes('.DS_Store')) return false;
      const base = name.split('/').pop();
      return !base.startsWith('._') && !base.startsWith('.');
    });

    // Find spreadsheet files first
    const spreadsheetFile = validEntries.find(name => /\.(xlsx|xls|csv|tsv|xlsm|ods)$/i.test(name)) || validEntries[0];
    if (!spreadsheetFile) {
      throw new Error('No spreadsheet or data file found in the uploaded ZIP archive.');
    }
    const unzippedData = await zip.files[spreadsheetFile].async('uint8array');
    return parseWorkbookFromBytes(unzippedData);
  }

  return parseWorkbookFromBytes(buffer);
};

export const scanWorkbookSheets = (workbook) => {
  const sheetSummaries = [];
  let bestSheetName = '';
  let bestExtraction = null;
  let maxSheetScore = -1;

  for (const name of workbook.SheetNames) {
    const sheet = workbook.Sheets[name];
    if (!sheet) continue;
    const extraction = extractAdmissionRowsFromSheet(sheet);
    // Score based on candidate count and header quality
    const sheetScore = (extraction.rows.length * 10) + extraction.score;
    sheetSummaries.push({
      name,
      rowCount: extraction.rows.length,
      headerRowIdx: extraction.headerRowIdx,
      score: extraction.score,
      totalScore: sheetScore,
      matchedHeaders: extraction.matchedHeaders
    });

    if (sheetScore > maxSheetScore) {
      maxSheetScore = sheetScore;
      bestSheetName = name;
      bestExtraction = extraction;
    }
  }

  // Fallback: If best extraction has 0 rows, check naive sheet_to_json across sheets
  if (!bestExtraction || bestExtraction.rows.length === 0) {
    for (const name of workbook.SheetNames) {
      const sheet = workbook.Sheets[name];
      if (!sheet) continue;
      const naiveRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      if (naiveRows && naiveRows.length > 0) {
        const headers = Object.keys(naiveRows[0]);
        bestSheetName = name;
        bestExtraction = {
          rows: naiveRows,
          headers,
          score: 1,
          headerRowIdx: 0,
          matchedHeaders: []
        };
        break;
      }
    }
  }

  return {
    bestSheetName: bestSheetName || workbook.SheetNames[0],
    bestExtraction: bestExtraction || { rows: [], headers: [], score: 0, headerRowIdx: 0, matchedHeaders: [] },
    sheetSummaries
  };
};

export const generateSampleAllotmentRows = () => [
  {
    'Sl No': 1,
    'Application Number': '2099010001',
    'Candidate Name': 'STUDENT ALPHA',
    'Gender': 'Male',
    'Date of Birth': '2005-04-12',
    'Category': 'General',
    'Caste': 'Nair',
    'Religion': 'Hindu',
    'Mobile Number': '9876543210',
    'Email ID': 'alpha.student@sample-univ.edu',
    'Program Name': 'B.A. History (Honours)',
    'College Name': 'Alpha Arts and Science College',
    'College Code': '101',
    'Total Marks': '1200',
    'Marks Obtained': '1120',
    'Register Number': '2099HSE0101',
    'Stream': 'Humanities',
    'Allotment Quota': 'Open Merit'
  },
  {
    'Sl No': 2,
    'Application Number': '2099010002',
    'Candidate Name': 'STUDENT BETA',
    'Gender': 'Female',
    'Date of Birth': '2005-08-20',
    'Category': 'OBC',
    'Caste': 'Ezhava',
    'Religion': 'Hindu',
    'Mobile Number': '9876543211',
    'Email ID': 'beta.student@sample-univ.edu',
    'Program Name': 'B.Com Finance (Honours)',
    'College Name': 'Beta Commerce College',
    'College Code': '304',
    'Total Marks': '1200',
    'Marks Obtained': '1085',
    'Register Number': '2099HSE0102',
    'Stream': 'Commerce',
    'Allotment Quota': 'SEBC / OBC'
  },
  {
    'Sl No': 3,
    'Application Number': '2099010003',
    'Candidate Name': 'STUDENT GAMMA',
    'Gender': 'Male',
    'Date of Birth': '2005-01-15',
    'Category': 'SC',
    'Caste': 'Pulaya',
    'Religion': 'Hindu',
    'Mobile Number': '9876543212',
    'Email ID': 'gamma.student@sample-univ.edu',
    'Program Name': 'B.Sc Physics (Honours)',
    'College Name': 'Gamma Science College',
    'College Code': '347',
    'Total Marks': '1200',
    'Marks Obtained': '1150',
    'Register Number': '2099HSE0103',
    'Stream': 'Science',
    'Allotment Quota': 'Scheduled Caste'
  },
  {
    'Sl No': 4,
    'Application Number': '2099010004',
    'Candidate Name': 'STUDENT DELTA',
    'Gender': 'Female',
    'Date of Birth': '2004-11-03',
    'Category': 'General',
    'Caste': 'Christian',
    'Religion': 'Christian',
    'Mobile Number': '9876543213',
    'Email ID': 'delta.student@sample-univ.edu',
    'Program Name': 'B.A. English (Honours)',
    'College Name': 'Delta Academy',
    'College Code': '358',
    'Total Marks': '1200',
    'Marks Obtained': '1040',
    'Register Number': '2099CBSE0104',
    'Stream': 'Humanities',
    'Allotment Quota': 'Open Merit'
  },
  {
    'Sl No': 5,
    'Application Number': '2099010005',
    'Candidate Name': 'STUDENT EPSILON',
    'Gender': 'Female',
    'Date of Birth': '2005-05-25',
    'Category': 'EWS',
    'Caste': 'General-EWS',
    'Religion': 'Hindu',
    'Mobile Number': '9876543214',
    'Email ID': 'epsilon.student@sample-univ.edu',
    'Program Name': 'B.Sc Mathematics (Honours)',
    'College Name': 'Alpha Arts and Science College',
    'College Code': '101',
    'Total Marks': '1200',
    'Marks Obtained': '1190',
    'Register Number': '2099HSE0105',
    'Stream': 'Science',
    'Allotment Quota': 'Economically Weaker Section'
  },
  {
    'Sl No': 6,
    'Application Number': '2099010006',
    'Candidate Name': 'STUDENT ZETA',
    'Gender': 'Male',
    'Date of Birth': '2005-09-18',
    'Category': 'Muslim',
    'Caste': 'Muslim',
    'Religion': 'Islam',
    'Mobile Number': '9876543215',
    'Email ID': 'zeta.student@sample-univ.edu',
    'Program Name': 'Integrated M.Sc (IMPES)',
    'College Name': 'Beta Commerce College',
    'College Code': '304',
    'Total Marks': '1200',
    'Marks Obtained': '1135',
    'Register Number': '2099HSE0106',
    'Stream': 'Science',
    'Allotment Quota': 'Muslim Merit'
  }
];

export default function AdmissionImportPage() {
  const [sourceFile, setSourceFile] = useState(null);
  const [currentWorkbook, setCurrentWorkbook] = useState(null);
  const [sheetList, setSheetList] = useState([]);
  const [selectedSheetName, setSelectedSheetName] = useState('');
  const [normalizedRows, setNormalizedRows] = useState([]);
  const [sourceHeaders, setSourceHeaders] = useState([]);
  const [stats, setStats] = useState({ total: 0, warnings: 0, ready: 0 });
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('info');
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [filterMode, setFilterMode] = useState('all'); // 'all' | 'warnings' | 'ready'
  const [page, setPage] = useState(0);
  const pageSize = 50;

  // Custom User-Defined Lookups
  const [lookups, setLookups] = useState(() => loadSavedLookups());
  const [isLookupModalOpen, setIsLookupModalOpen] = useState(false);
  const [lastExtractedData, setLastExtractedData] = useState(null);

  const lookupMaps = useMemo(() => buildLookupMaps(lookups), [lookups]);

  const setStatus = (msg, type = 'info') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  const transformRow = (rawRow, headerMap, activeLookupMaps = lookupMaps) => {
    const getVal = (...keys) => {
      for (const k of keys) {
        const norm = normalizeKey(k);
        const actualKey = headerMap[norm];
        if (actualKey && rawRow[actualKey] !== undefined && rawRow[actualKey] !== null) {
          const val = cleanEscapeArtifacts(rawRow[actualKey]);
          // Strip literal header artifacts (e.g. literal 'CollegeId' in data cell)
          if (val && normalizeKey(val) !== norm) {
            return val;
          }
        }
      }
      return '';
    };

    // 1. Name
    const name = getVal('name', 'fullname', 'candidatename', 'studentname', 'cname');

    // 2. Gender
    const rawGender = getVal('gender', 'sex');
    let gender = 'Male';
    if (rawGender.toLowerCase().startsWith('f')) gender = 'Female';
    else if (rawGender.toLowerCase().startsWith('m')) gender = 'Male';
    else if (rawGender) gender = rawGender;

    // 3. Mobile (clean 10 digits)
    const rawMobile = getVal('mobile', 'mobilenumber', 'contact', 'contactnumber', 'phone');
    const mobile = rawMobile.replace(/\D/g, '').slice(-10);

    // 4. Email (clean lowercase)
    const rawEmail = getVal('email', 'emailid', 'mail');
    const email = rawEmail.toLowerCase().trim();

    // 5. Username (Application Number / Reg No / Admn No)
    const username = getVal('username', 'applno', 'applicationnumber', 'appno', 'regno', 'registerno', 'admnum', 'prn', 'candidateid');

    // 6 & 7. Date of Birth: YYYY-MM-DD for `dob` and DD/MM/YYYY for `Date of Birth (DD/MM/YY)`
    const rawDob = getVal('dob', 'dateofbirth', 'birthdate');
    const dob = parseDobYMD(rawDob);
    const dobDMY = parseDobDMY(rawDob);

    // 8. AadhaarNumber (Preserve exact 12 digits as string)
    const rawAadhaar = getVal('aadhaarnumber', 'aadhar', 'adharnumber', 'adhar');
    const aadhaarNumber = rawAadhaar.replace(/\D/g, '').slice(0, 12);

    // 9. ABCId (Preserve exact 12 digits as string)
    const rawABC = getVal('abcid', 'abc', 'academicbankofcredits');
    const abcId = rawABC.replace(/\D/g, '').slice(0, 12);

    // 10 & 11. CountryId & CountryName
    const rawCountry = getVal('countryname', 'country', 'nationality');
    const normCountry = normalizeLookupKey(rawCountry);
    const customCountry = activeLookupMaps?.countryByName?.[normCountry] || activeLookupMaps?.countryByName?.[rawCountry.toLowerCase()];
    const countryId = getVal('countryid') || (customCountry ? customCountry.id : '') || '107';
    let countryName = rawCountry || (activeLookupMaps?.countryById?.[countryId]?.name) || (customCountry ? customCountry.name : '') || 'India';
    if (countryId === '107' || normCountry === 'indian' || normCountry === 'india') {
      countryName = 'India';
    }

    // 12 & 13. State (corrStateId & corrStateName)
    const rawState = getVal('corrstatename', 'statename', 'state', 'corrstate');
    const normState = normalizeLookupKey(rawState);
    const customState = activeLookupMaps?.stateByName?.[normState] || activeLookupMaps?.stateByName?.[rawState.toLowerCase()];
    const corrStateId = getVal('corrstateid', 'stateid') || (customState ? customState.id : '') || (normState === 'kerala' ? '18' : '');
    const corrStateName = rawState || (activeLookupMaps?.stateById?.[corrStateId]?.name) || (customState ? customState.name : '') || 'Kerala';

    // 14, 15, 16. Address, City, Pin
    const corrAddress = getVal('corraddress', 'address', 'communicationaddress', 'permaddress');
    let corrCity = getVal('corrcity', 'city', 'town', 'place', 'district');
    const rawPin = getVal('corrpin', 'pincode', 'pin', 'postalcode');
    const corrPin = rawPin.replace(/\D/g, '').slice(0, 6);

    // 17 & 18. Category (categoryId & categoryName)
    const rawCategory = getVal('category', 'categoryname', 'castecategory', 'community', 'actual_category');
    const normCategory = normalizeLookupKey(rawCategory);
    const customCategory = activeLookupMaps?.categoryByName?.[normCategory] || activeLookupMaps?.categoryByName?.[rawCategory.toLowerCase()];
    const fallbackCategory = CATEGORY_LOOKUP[normCategory] || { id: '10', name: rawCategory || 'General' };
    const categoryId = getVal('categoryid') || (customCategory ? customCategory.id : '') || fallbackCategory.id;
    const categoryName = rawCategory || (customCategory ? customCategory.name : '') || fallbackCategory.name;

    // 19 & 20. Caste (CasteId & Caste)
    const rawCaste = getVal('caste', 'subcaste');
    const normCaste = normalizeLookupKey(rawCaste);
    const customCaste = activeLookupMaps?.casteByName?.[normCaste] || activeLookupMaps?.casteByName?.[rawCaste.toLowerCase()];
    const fallbackCasteId = CASTE_LOOKUP[normCaste] || CASTE_LOOKUP[rawCaste] || '1';
    const casteId = getVal('casteid') || (customCaste ? customCaste.id : '') || fallbackCasteId;
    const caste = rawCaste || (customCaste ? customCaste.name : '') || (activeLookupMaps?.casteById?.[casteId]?.name) || 'General';

    // 21 & 22. Religion (religionId & religion)
    const rawReligion = getVal('religion', 'religionname');
    const normReligion = normalizeLookupKey(rawReligion);
    const customReligion = activeLookupMaps?.religionByName?.[normReligion] || activeLookupMaps?.religionByName?.[rawReligion.toLowerCase()];
    const fallbackReligion = RELIGION_LOOKUP[normReligion] || { id: '1', name: rawReligion || 'Hindu' };
    const religionId = getVal('religionid') || (customReligion ? customReligion.id : '') || fallbackReligion.id;
    const religion = rawReligion || (customReligion ? customReligion.name : '') || fallbackReligion.name;

    // 23, 24, 25, 26. Program Details
    // Sourced from Program Name, Programme, Course, etc.
    const rawProgram = getVal('programname', 'program_name', 'program', 'programme', 'programmename', 'programmname', 'coursename', 'course_name', 'course', 'pgmname', 'pgm_name', 'degree', 'subject');
    const normProg = normalizeKey(rawProgram);
    let programMatch = activeLookupMaps?.programByName?.[rawProgram.toLowerCase()] || 
                       activeLookupMaps?.programByName?.[normProg] || 
                       PROGRAM_LOOKUP[rawProgram.toLowerCase()] || 
                       PROGRAM_LOOKUP[normProg];
    if (!programMatch && activeLookupMaps?.programByName) {
      for (const [pk, pv] of Object.entries(activeLookupMaps.programByName)) {
        const npk = normalizeKey(pk);
        if (normProg && npk && (normProg.includes(npk) || npk.includes(normProg))) {
          programMatch = pv;
          break;
        }
      }
    }
    if (!programMatch) {
      for (const [pk, pv] of Object.entries(PROGRAM_LOOKUP)) {
        const npk = normalizeKey(pk);
        if (normProg && npk && (normProg.includes(npk) || npk.includes(normProg))) {
          programMatch = pv;
          break;
        }
      }
    }

    // Keyword & discipline fallback matching for Kannur programmes
    if (!programMatch && normProg) {
      const disciplines = [
        { terms: ['clinicalpsychology', 'psychology'], code: 'PFPSYGS25' },
        { terms: ['computational', 'computationalscience'], code: 'PFCLSGS25' },
        { terms: ['computerapplication', 'mca'], code: 'PFMCAGS25' },
        { terms: ['businessadministration', 'management', 'bba', 'mba'], code: 'PFMBAGA25' },
        { terms: ['geography'], code: 'PFGEOGS25' },
        { terms: ['music'], code: 'PFMUSGA25' },
        { terms: ['physicaleducation', 'sports', 'impes'], code: 'PFPESGS25' },
        { terms: ['anthropolog'], code: 'PFANTGS25' },
        { terms: ['english'], code: 'PFENGGA25' },
        { terms: ['envirn', 'environ'], code: 'PFEVSGS25' },
        { terms: ['historical', 'history'], code: 'PFHISGS25' },
        { terms: ['mathemat'], code: 'PFMATGS25' },
        { terms: ['physicalscience'], code: 'PFPLSGS25' },
        { terms: ['socialscience', 'economics'], code: 'PFSSEGS25' },
        { terms: ['commerce', 'bcom', 'mcom'], code: 'PFCOMGS25' }
      ];
      for (const d of disciplines) {
        if (d.terms.some(t => normProg.includes(t))) {
          programMatch = activeLookupMaps?.programByCode?.[d.code] || KANNUR_PROGRAMMES.find(p => p.code === d.code) || null;
          if (programMatch) break;
        }
      }
    }

    // Fallback to code lookup if a valid program code is supplied in source file
    const rawProgCodeCandidate = getVal('programcode', 'progcode', 'degreecode');
    if (!programMatch && rawProgCodeCandidate) {
      programMatch = activeLookupMaps?.programByCode?.[rawProgCodeCandidate] || 
                     activeLookupMaps?.programByCode?.[rawProgCodeCandidate.toLowerCase()] ||
                     KANNUR_PROGRAMMES.find(p => p.code === rawProgCodeCandidate) || null;
    }

    const pgmName = (programMatch ? programMatch.pgmName : '') || getVal('pgmname', 'pgm_name') || rawProgram || 'Five Year Integrated Masters Programme in  Business Administration';
    const programName = (programMatch ? programMatch.name : '') || rawProgram || 'Five Year Integrated Masters Programme in Business Administration(with Credits)-Regular-FYIMP-2025-Business Administration-FYIMP Year I SEMESTER I';

    // Col 24 workFlowId: strictly use the WORKFLOWID looked up against the Program Name
    const lookedUpWorkflowId = (programMatch && programMatch.workflowId && programMatch.workflowId !== 'NA') ? String(programMatch.workflowId).trim() : '';
    const rawWorkflowId = getVal('workflowid', 'workflow_id', 'wfid');
    const workFlowId = lookedUpWorkflowId || (rawWorkflowId && rawWorkflowId !== 'NA' ? rawWorkflowId : '') || '162';

    // Col 26 ProgramCode: strictly use the PROGRAMCODE looked up against the Program Name
    const lookedUpProgramCode = (programMatch && programMatch.code) ? String(programMatch.code).trim() : '';
    const rawProgramCode = getVal('programcode', 'progcode', 'degreecode');
    const programCode = lookedUpProgramCode || (rawProgramCode && isNaN(Number(rawProgramCode)) ? rawProgramCode : '') || rawProgramCode || 'PFMBAGA25';

    // 27, 28, 29, 30. College Details
    const rawCollegeCode = getVal('collegecode', 'colcode', 'centre', 'center', 'centercode');
    const rawCollegeName = getVal('collegename', 'college', 'centername');
    const customCollege = activeLookupMaps?.collegeByName?.[normalizeKey(rawCollegeName)] ||
                          activeLookupMaps?.collegeByName?.[rawCollegeName.toLowerCase()] ||
                          activeLookupMaps?.collegeByCode?.[rawCollegeCode] ||
                          activeLookupMaps?.collegeById?.[rawCollegeCode];
    const collegeLookup = customCollege || COLLEGE_LOOKUP[normalizeKey(rawCollegeName)] || COLLEGE_LOOKUP[rawCollegeCode] || COLLEGE_LOOKUP[normalizeKey(rawCollegeCode)] || {
      id: (getVal('collegeid') !== 'CollegeId' ? getVal('collegeid') : '') || rawCollegeCode || '133',
      code: rawCollegeCode || '133',
      name: rawCollegeName || 'Department of Management Studies',
      clgRef: rawCollegeName || 'Department of Management Studies,Thavakkara Campus,Thavakkara,pin:670002'
    };
    const clgName = getVal('clgname') || collegeLookup.clgRef || collegeLookup.name;
    const collegeId = (collegeLookup ? collegeLookup.id : '') || (getVal('collegeid') !== 'CollegeId' ? getVal('collegeid') : '') || '133';
    const collegeName = collegeLookup.name || rawCollegeName;
    const collegeCode = collegeLookup.code || rawCollegeCode;

    // 31 & 32. Qualification (QualificationFacultyId & QualificationName)
    const rawQual = getVal('qualificationname', 'qcboard', 'board', 'qualifyingexam', 'boardname');
    const normQual = normalizeLookupKey(rawQual);
    const customQual = activeLookupMaps?.qualificationByName?.[normQual] || activeLookupMaps?.qualificationByName?.[rawQual.toLowerCase()];
    const fallbackQual = QUALIFICATION_LOOKUP[rawQual.toLowerCase()] || QUALIFICATION_LOOKUP[normQual] || null;
    let qualificationFacultyId = getVal('qualificationfacultyid') || (customQual ? customQual.id : '') || (fallbackQual ? fallbackQual.facultyId : '2');
    if (!customQual && !fallbackQual && normQual && !normQual.includes('hse') && !normQual.includes('cbse') && !normQual.includes('vhse') && !normQual.includes('isc')) {
      qualificationFacultyId = '7';
    }
    const qualificationName = rawQual || (customQual ? customQual.name : '') || 'HSE - Kerala';

    // 33 & 34. Stream (Stream & QualificationSpecializationId)
    const rawStream = getVal('stream', 'group', 'specialization', 'branch');
    const normStream = normalizeLookupKey(rawStream);
    const customStream = activeLookupMaps?.streamByName?.[normStream] || activeLookupMaps?.streamByName?.[rawStream.toLowerCase()];
    let qualificationSpecializationId = '2'; // Default Commerce (2)
    if (customStream) {
      qualificationSpecializationId = customStream.id;
    } else if (STREAM_LOOKUP[normStream]) {
      qualificationSpecializationId = STREAM_LOOKUP[normStream].id;
    } else if (normStream.includes('sci')) {
      qualificationSpecializationId = '1';
    } else if (normStream.includes('comm')) {
      qualificationSpecializationId = '2';
    } else if (normStream.includes('hum') || normStream.includes('art')) {
      qualificationSpecializationId = '3';
    }
    const stream = rawStream || (customStream ? customStream.name : '') || (qualificationSpecializationId === '1' ? 'Science' : qualificationSpecializationId === '3' ? 'Humanities' : 'Commerce');

    // 35. Certificate Number
    const certificateNumber = getVal('certificatenumber', 'qcno', 'regno', 'rollno', 'register_number', 'registerno');

    // 36. Attempts
    const rawAttempts = getVal('noofattempts', 'attempts');
    const noOfAttempts = rawAttempts ? parseInt(rawAttempts, 10) || 1 : 1;

    // 37, 38, 39, 40. Marks, Percentage, CGPA
    const rawMarks = getVal('marksobtained', 'totalmarks', 'securedmarks', 'marks');
    const marksObtained = rawMarks ? (parseFloat(rawMarks) || '') : '';
    const rawMax = getVal('marksoutof', 'maxmarks', 'maximummarks', 'total');
    const marksOutOf = rawMax ? (parseFloat(rawMax) || '1200') : (marksObtained ? '1200' : '');

    let percentage = getVal('percentage', 'percent');
    if (!percentage && marksOutOf > 0 && marksObtained > 0) {
      const calcPct = (parseFloat(marksObtained) / parseFloat(marksOutOf)) * 100;
      percentage = Number(calcPct.toFixed(4)).toString();
    } else if (percentage) {
      percentage = String(percentage).trim();
    } else {
      percentage = '';
    }

    const rawCgpa = getVal('cgpa', 'gpa');
    let cgpa = '';
    if (rawCgpa) {
      cgpa = String(rawCgpa).trim();
    } else if (percentage) {
      cgpa = percentage;
    }

    // 41. SeatNumber
    const rawSeat = getVal('seatnumber', 'seatno', 'qcno');
    const seatNumber = rawSeat || certificateNumber;

    // 42. Passing Year (Column 42 with blank header '')
    const rawResultDate = getVal('resultdate', 'passeddate', 'admit_date');
    let passingYear = getVal('qcyear', 'passingyear', 'admyear');
    let resultDate = '';

    // Result Date in format YYYY-MM-DD
    if (rawResultDate && /^\d{4}$/.test(rawResultDate.trim())) {
      passingYear = rawResultDate.trim();
      resultDate = `${passingYear}-05-31`;
    } else if (rawResultDate) {
      resultDate = parseDobYMD(rawResultDate);
    }

    if (!passingYear) {
      if (resultDate && /^\d{4}/.test(resultDate)) {
        passingYear = resultDate.substring(0, 4);
      } else {
        passingYear = '2026';
      }
    }
    if (!resultDate) resultDate = `${passingYear}-05-31`;

    // 44. ResultStatus
    const resultStatus = getVal('resultstatus', 'status', 'result') || 'Passed';

    // Warnings flag
    let warnings = [];
    if (!name) warnings.push('Missing Name');
    if (!mobile || mobile.length < 10) warnings.push('Invalid Mobile');
    if (!username && !certificateNumber) warnings.push('Missing Identifier/RegNo');

    return {
      Name: name,
      Gender: gender,
      Mobile: mobile,
      Email: email,
      username: username || certificateNumber,
      dob: dob,
      'Date of Birth (DD/MM/YY)': dobDMY,
      AadhaarNumber: aadhaarNumber,
      ABCId: abcId,
      CountryId: countryId,
      CountryName: countryName,
      corrStateId: corrStateId,
      corrStateName: corrStateName,
      corrAddress: corrAddress,
      corrCity: corrCity,
      corrPin: corrPin,
      categoryId: categoryId,
      categoryName: categoryName,
      CasteId: casteId,
      Caste: caste,
      religionId: religionId,
      religion: religion,
      'PGM Name': pgmName,
      WORKFLOWID: workFlowId,
      workFlowId: workFlowId,
      'Program Name': programName,
      PROGRAMCODE: programCode,
      ProgramCode: programCode,
      'CLG NAME': clgName,
      CollegeId: collegeId,
      'College Name': collegeName,
      CollegeCode: collegeCode,
      QualificationFacultyId: qualificationFacultyId,
      QualificationName: qualificationName,
      QualificationSpecializationId: qualificationSpecializationId,
      Stream: stream,
      CertificateNumber: certificateNumber,
      NoOfAttempts: noOfAttempts,
      MarksObtained: marksObtained,
      Percentage: percentage,
      MarksOutOf: marksOutOf,
      CGPA: cgpa,
      SeatNumber: seatNumber,
      '': passingYear,
      'Result Date ': resultDate,
      'Result Date': resultDate,
      ResultDate: resultDate,
      ResultStatus: resultStatus,
      _warnings: warnings
    };
  };

  const applyExtractedRows = (extractedData, originName = 'file', customLookupMaps = lookupMaps) => {
    const { rows, headers } = extractedData;
    if (!rows || rows.length === 0) {
      throw new Error('No candidate data rows found in the selected worksheet.');
    }

    setLastExtractedData({ rows, headers, originName });
    setSourceHeaders(headers);
    const headerMap = {};
    headers.forEach(h => {
      headerMap[normalizeKey(h)] = h;
    });

    const transformed = rows.map(row => transformRow(row, headerMap, customLookupMaps));
    setNormalizedRows(transformed);

    const warningCount = transformed.filter(r => r._warnings && r._warnings.length > 0).length;
    setStats({
      total: transformed.length,
      warnings: warningCount,
      ready: transformed.length - warningCount
    });

    setStatus(`Successfully normalized ${transformed.length} candidate record(s) from "${originName}" into ${TARGET_COLUMNS.length} master columns!`, 'success');
  };

  const handleSaveLookups = (updatedLookups) => {
    setLookups(updatedLookups);
    const updatedMaps = buildLookupMaps(updatedLookups);
    if (lastExtractedData && lastExtractedData.rows && lastExtractedData.rows.length > 0) {
      applyExtractedRows(lastExtractedData, lastExtractedData.originName, updatedMaps);
    }
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setSourceFile(file);
    setIsProcessing(true);
    setStatus('Reading and analyzing spreadsheet structure...', 'info');

    try {
      const buffer = await file.arrayBuffer();
      const workbook = await parseWorkbookBuffer(buffer, file.name);
      if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
        throw new Error('Unable to read workbook sheets from the selected file.');
      }

      setCurrentWorkbook(workbook);

      const { bestSheetName, bestExtraction, sheetSummaries } = scanWorkbookSheets(workbook);
      setSheetList(sheetSummaries);
      setSelectedSheetName(bestSheetName);

      if (!bestExtraction || bestExtraction.rows.length === 0) {
        const sheetNamesStr = workbook.SheetNames.join(', ');
        throw new Error(`The uploaded spreadsheet contains no data rows across examined sheet(s): [${sheetNamesStr}]. Please verify the file contains candidate columns.`);
      }

      applyExtractedRows(bestExtraction, bestSheetName);
    } catch (err) {
      console.error('Admission ingestion error:', err);
      setStatus(`Ingestion failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
      if (e.target) e.target.value = '';
    }
  };

  const handleSheetChange = (newSheetName) => {
    if (!currentWorkbook || !currentWorkbook.Sheets[newSheetName]) return;
    setSelectedSheetName(newSheetName);
    setIsProcessing(true);
    setStatus(`Extracting candidates from sheet "${newSheetName}"...`, 'info');

    try {
      const sheet = currentWorkbook.Sheets[newSheetName];
      const extraction = extractAdmissionRowsFromSheet(sheet);
      if (!extraction.rows.length) {
        const naive = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        if (naive && naive.length > 0) {
          applyExtractedRows({ rows: naive, headers: Object.keys(naive[0]) }, newSheetName);
          return;
        }
        throw new Error(`Worksheet "${newSheetName}" contains no candidate data rows.`);
      }
      applyExtractedRows(extraction, newSheetName);
    } catch (err) {
      console.error('Sheet switch error:', err);
      setStatus(`Worksheet extraction failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleLoadSampleData = () => {
    setIsProcessing(true);
    setStatus('Loading sample FYUGP & FYIMP candidate allotment roster...', 'info');

    try {
      const sampleRows = generateSampleAllotmentRows();
      // Construct a realistic in-memory workbook with institutional title banner to test header detection
      const aoa = [
        ['STATE CENTRALIZED ALLOTMENT PROCESS (CAP 2026-27)'],
        ['COLLEGE ALLOTMENT MERIT ROSTER - ADMISSION IMPORT ENGINE'],
        Object.keys(sampleRows[0]),
        ...sampleRows.map(r => Object.values(r))
      ];

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sample_Allotment_Roster');

      setCurrentWorkbook(wb);
      setSheetList([{
        name: 'Sample_Allotment_Roster',
        rowCount: sampleRows.length,
        headerRowIdx: 2,
        score: 55,
        totalScore: (sampleRows.length * 10) + 55,
        matchedHeaders: ['Name', 'Identifier / SlNo', 'Mobile', 'Email', 'Marks']
      }]);
      setSelectedSheetName('Sample_Allotment_Roster');
      setSourceFile({ name: 'sample_allotment_roster_2026.xlsx' });

      const extraction = extractAdmissionRowsFromSheet(ws);
      applyExtractedRows(extraction, 'Sample_Allotment_Roster');
      setStatus(`Loaded hypothetical sample allotment roster (${sampleRows.length} candidate records across FYUGP & FYIMP)!`, 'success');
    } catch (err) {
      console.error('Sample data error:', err);
      setStatus(`Failed to load sample data: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const updateCell = (rowIndex, colKey, newValue) => {
    setNormalizedRows(prev => {
      const copy = [...prev];
      copy[rowIndex] = { ...copy[rowIndex], [colKey]: newValue };
      return copy;
    });
  };

  const exportStandardXlsx = () => {
    if (!normalizedRows.length) return alert('No data available to export.');
    setIsProcessing(true);
    setStatus(`Generating standardized ${TARGET_COLUMNS.length}-column master Excel workbook...`, 'info');

    try {
      // Export array of arrays for dense mode memory efficiency
      const exportAoa = [TARGET_COLUMNS];
      normalizedRows.forEach(row => {
        exportAoa.push(TARGET_COLUMNS.map(col => {
          const val = row[col] !== undefined && row[col] !== null ? row[col] : '';
          if (typeof val === 'number' && (col === 'AadhaarNumber' || col === 'ABCId' || col === 'Mobile' || col === 'corrPin' || col === 'CertificateNumber' || col === 'SeatNumber')) {
            return String(val);
          }
          return val;
        }));
      });

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(exportAoa, { dense: true });

      // Auto-filter on all columns
      ws['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: Math.max(exportAoa.length - 1, 0), c: TARGET_COLUMNS.length - 1 }
        })
      };

      // Set column widths
      ws['!cols'] = TARGET_COLUMNS.map(c => ({ wch: Math.max((c || 'Year').length + 3, 14) }));

      XLSX.utils.book_append_sheet(wb, ws, 'Master_Admission');

      const outBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
      const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const baseName = (sourceFile?.name || 'allotment').replace(/\.[^/.]+$/, '');
      a.download = `${baseName}_Master_Admission.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);

      setStatus(`Exported ${TARGET_COLUMNS.length}-column Master XLSX with ${normalizedRows.length} records!`, 'success');
    } catch (err) {
      console.error('Export error:', err);
      setStatus(`Export failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const filteredRows = useMemo(() => {
    return normalizedRows.filter((row) => {
      if (filterMode === 'warnings' && (!row._warnings || !row._warnings.length)) return false;
      if (filterMode === 'ready' && row._warnings && row._warnings.length) return false;
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        return (
          String(row.Name || '').toLowerCase().includes(q) ||
          String(row.username || '').toLowerCase().includes(q) ||
          String(row.Mobile || '').includes(q) ||
          String(row.Email || '').toLowerCase().includes(q) ||
          String(row['Program Name'] || '').toLowerCase().includes(q)
        );
      }
      return true;
    });
  }, [normalizedRows, filterMode, searchQuery]);

  const pagedRows = useMemo(() => {
    const start = page * pageSize;
    return filteredRows.slice(start, start + pageSize);
  }, [filteredRows, page]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', overflow: 'hidden', background: 'var(--bg)', color: 'var(--ink)' }}>
      {/* Top Header Bar */}
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 24px', background: 'var(--panel)', borderBottom: '1px solid var(--line)', zIndex: 20 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}>
            <ArrowLeft size={16} /> Back to Portal
          </Link>
          <div style={{ height: '18px', width: '1px', background: 'var(--line)' }} />
          <h2 style={{ fontSize: '16px', margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Database size={18} color="var(--accent)" /> Admission Import & Transformation Engine
            <span style={{ fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '12px', fontWeight: 700 }}>FYUGP & FYIMP v1.1</span>
          </h2>
        </div>

        {/* Global Status Pill & Export Button */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
          <div style={{
            padding: '4px 12px',
            borderRadius: '16px',
            fontSize: '12px',
            fontWeight: 600,
            background: statusType === 'success' ? 'var(--accent-soft)' : statusType === 'error' ? 'var(--danger)' : 'var(--bg)',
            color: statusType === 'success' ? 'var(--accent)' : statusType === 'error' ? 'white' : 'var(--muted)',
            border: '1px solid var(--line)'
          }}>
            {statusMsg}
          </div>

          <button
            type="button"
            onClick={() => setIsLookupModalOpen(true)}
            title="Define and manage ID mappings for Category, Caste, Religion, Country, State, Qualification, and Stream"
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              padding: '6px 14px',
              fontSize: '12px',
              fontWeight: 600,
              background: 'var(--panel)',
              color: 'var(--ink)',
              border: '1px solid var(--line)',
              borderRadius: '6px',
              cursor: 'pointer'
            }}
          >
            <SlidersHorizontal size={14} color="var(--accent)" />
            <span>Define / Manage IDs</span>
            <span style={{
              fontSize: '10px',
              background: 'var(--accent-soft)',
              color: 'var(--accent)',
              padding: '1px 6px',
              borderRadius: '10px',
              marginLeft: '2px',
              fontWeight: 700
            }}>
              9 Master Entities
            </span>
          </button>

          <button 
            type="button" 
            disabled={!normalizedRows.length || isProcessing}
            onClick={exportStandardXlsx}
            style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 16px', fontSize: '12px' }}
          >
            <Download size={14} /> Export {TARGET_COLUMNS.length}-Col Master XLSX
          </button>
        </div>
      </header>

      {/* Main Container */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden', padding: '16px 24px', gap: '16px' }}>
        
        {/* Left Control & Metrics Panel */}
        <aside style={{ width: '320px', display: 'flex', flexDirection: 'column', gap: '14px', overflowY: 'auto' }}>
          
          {/* File Upload Card */}
          <div className="card" style={{ padding: '20px', margin: 0, textAlign: 'center', border: '1.5px dashed var(--accent)', background: 'var(--accent-soft)' }}>
            <input 
              type="file" 
              id="admissionFileInput" 
              accept=".xlsx,.xls,.csv,.tsv,.zip" 
              onChange={handleFileUpload} 
              style={{ display: 'none' }} 
            />
            <label htmlFor="admissionFileInput" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '6px' }}>
              <Upload size={28} color="var(--accent)" />
              <strong style={{ fontSize: '13.5px', color: 'var(--ink)' }}>Upload Allotment Spreadsheet</strong>
              <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Supports .xlsx, .xls (HTML/ERP), .csv, .zip</span>
            </label>
          </div>

          {/* Quick Demo Sample Data Button */}
          <button
            type="button"
            className="secondary"
            onClick={handleLoadSampleData}
            disabled={isProcessing}
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '6px',
              padding: '8px 12px',
              fontSize: '12px',
              width: '100%',
              fontWeight: 600,
              background: 'var(--panel)',
              border: '1px dashed var(--accent)',
              color: 'var(--accent)',
              borderRadius: '6px',
              cursor: 'pointer'
            }}
          >
            <Sparkles size={14} /> Load Sample Allotment Data (Instant Demo)
          </button>

          {/* ID Lookups & Master Codes Card */}
          <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '8px', background: 'var(--panel)', border: '1px solid var(--line)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', display: 'flex', alignItems: 'center', gap: '5px' }}>
                <SlidersHorizontal size={13} color="var(--accent)" /> ID Lookups & Master Codes
              </span>
              <span style={{ fontSize: '10.5px', color: 'var(--accent)', fontWeight: 600 }}>Active</span>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px', fontSize: '11px' }}>
              <div style={{ padding: '6px 8px', background: 'var(--bg)', borderRadius: '4px', border: '1px solid var(--line)' }}>
                <span style={{ color: 'var(--muted)' }}>Castes:</span> <strong>{lookups.caste?.length || 0}</strong>
              </div>
              <div style={{ padding: '6px 8px', background: 'var(--bg)', borderRadius: '4px', border: '1px solid var(--line)' }}>
                <span style={{ color: 'var(--muted)' }}>Categories:</span> <strong>{lookups.category?.length || 0}</strong>
              </div>
              <div style={{ padding: '6px 8px', background: 'var(--bg)', borderRadius: '4px', border: '1px solid var(--line)' }}>
                <span style={{ color: 'var(--muted)' }}>Religions:</span> <strong>{lookups.religion?.length || 0}</strong>
              </div>
              <div style={{ padding: '6px 8px', background: 'var(--bg)', borderRadius: '4px', border: '1px solid var(--line)' }}>
                <span style={{ color: 'var(--muted)' }}>Countries:</span> <strong>{lookups.country?.length || 0}</strong>
              </div>
              <div style={{ padding: '6px 8px', background: 'var(--bg)', borderRadius: '4px', border: '1px solid var(--line)' }}>
                <span style={{ color: 'var(--muted)' }}>States:</span> <strong>{lookups.state?.length || 0}</strong>
              </div>
              <div style={{ padding: '6px 8px', background: 'var(--bg)', borderRadius: '4px', border: '1px solid var(--line)' }}>
                <span style={{ color: 'var(--muted)' }}>Boards:</span> <strong>{lookups.qualification?.length || 0}</strong>
              </div>
            </div>
            <button
              type="button"
              onClick={() => setIsLookupModalOpen(true)}
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '6px',
                padding: '7px 10px',
                fontSize: '11.5px',
                fontWeight: 600,
                background: 'var(--bg)',
                border: '1px solid var(--accent)',
                color: 'var(--accent)',
                borderRadius: '5px',
                cursor: 'pointer',
                marginTop: '2px'
              }}
            >
              <SlidersHorizontal size={13} /> Manage / Upload Lookup IDs
            </button>
          </div>

          {/* Worksheet Selector (Multi-Sheet Workbooks) */}
          {sheetList.length > 1 && (
            <div className="card" style={{ padding: '12px 14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '6px', background: 'var(--panel)' }}>
              <label htmlFor="worksheetSelect" style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', display: 'flex', alignItems: 'center', gap: '5px' }}>
                <FileSpreadsheet size={14} color="var(--accent)" /> Active Worksheet ({sheetList.length} sheets found)
              </label>
              <select
                id="worksheetSelect"
                value={selectedSheetName}
                onChange={(e) => handleSheetChange(e.target.value)}
                style={{
                  width: '100%',
                  padding: '6px 8px',
                  fontSize: '12px',
                  borderRadius: '4px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  fontWeight: 600
                }}
              >
                {sheetList.map(s => (
                  <option key={s.name} value={s.name}>
                    {s.name} ({s.rowCount} rows{s.name === selectedSheetName ? ' • Selected' : ''})
                  </option>
                ))}
              </select>
              <span style={{ fontSize: '10.5px', color: 'var(--muted)', lineHeight: 1.3 }}>
                Automatically picked worksheet with highest candidate density. Switch if needed.
              </span>
            </div>
          )}

          {/* Ingestion Metrics Card */}
          {normalizedRows.length > 0 && (
            <div className="card" style={{ padding: '16px', margin: 0, display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <h3 style={{ margin: 0, fontSize: '13px', textTransform: 'uppercase', color: 'var(--muted)', fontWeight: 700 }}>Transformation Summary</h3>
                <span style={{ fontSize: '11px', color: 'var(--accent)', fontWeight: 600 }}>{selectedSheetName || 'Active'}</span>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                <div style={{ padding: '10px', background: 'var(--bg)', borderRadius: '6px', border: '1px solid var(--line)', textAlign: 'center' }}>
                  <div style={{ fontSize: '20px', fontWeight: 800, color: 'var(--accent)' }}>{stats.total}</div>
                  <div style={{ fontSize: '11px', color: 'var(--muted)' }}>Total Records</div>
                </div>
                <div style={{ padding: '10px', background: 'var(--bg)', borderRadius: '6px', border: '1px solid var(--line)', textAlign: 'center' }}>
                  <div style={{ fontSize: '20px', fontWeight: 800, color: '#16a34a' }}>40 / 40</div>
                  <div style={{ fontSize: '11px', color: 'var(--muted)' }}>Standard Columns</div>
                </div>
              </div>

              {sourceHeaders.length > 0 && (
                <div style={{ fontSize: '11px', color: 'var(--muted)', display: 'flex', justifyContent: 'space-between' }}>
                  <span>Source Headers Mapped:</span>
                  <strong style={{ color: 'var(--ink)' }}>{sourceHeaders.length} columns</strong>
                </div>
              )}

              {stats.warnings > 0 && (
                <div style={{ padding: '8px 12px', background: 'rgba(234, 179, 8, 0.1)', border: '1px solid rgba(234, 179, 8, 0.3)', borderRadius: '6px', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '11.5px', color: '#854d0e' }}>
                  <AlertTriangle size={16} />
                  <span><strong>{stats.warnings} record(s)</strong> have partial fields requiring review.</span>
                </div>
              )}
            </div>
          )}

          {/* Sanitization Rules v1.1 Reference Card */}
          <div className="card" style={{ padding: '16px', margin: 0, display: 'flex', flexDirection: 'column', gap: '8px', fontSize: '12px' }}>
            <h3 style={{ margin: 0, fontSize: '13px', fontWeight: 700 }}>Auto-Sanitization Rules (v1.1)</h3>
            <ul style={{ margin: 0, paddingLeft: '16px', color: 'var(--muted)', display: 'flex', flexDirection: 'column', gap: '4px', lineHeight: 1.4 }}>
              <li><strong>Pre-Header Title Scanner:</strong> Skips university banner rows & locates candidate headers.</li>
              <li><strong>Multi-Sheet Auto-Detect:</strong> Selects sheet with actual candidates over instruction sheets.</li>
              <li><strong>Excel Serial Dates:</strong> Converts serial integers (e.g. 38302) to ISO YYYY-MM-DD.</li>
              <li><strong>Mobile Normalization:</strong> Strips non-digits & prefixes.</li>
              <li><strong>\N & Null Trapping:</strong> Cleans MySQL export artifacts.</li>
              <li><strong>ZIP Archive Ingestion:</strong> Unpacks and scans contained spreadsheets.</li>
              <li><strong>Multi-Program:</strong> FYUGP & FYIMP / IMPES compatible.</li>
            </ul>
          </div>
        </aside>

        {/* Center/Right Live Validation & Editable Grid */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '10px', overflow: 'hidden' }}>
          
          {/* Table Toolbar */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 16px', borderBottom: '1px solid var(--line)', background: 'var(--bg)', flexWrap: 'wrap', gap: '10px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <div style={{ position: 'relative', width: '240px' }}>
                <input 
                  type="text" 
                  placeholder="Filter candidates, mobile, email..." 
                  value={searchQuery} 
                  onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }} 
                  style={{ width: '100%', padding: '5px 8px 5px 26px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)' }} 
                />
                <Search size={13} color="var(--muted)" style={{ position: 'absolute', left: '8px', top: '7px' }} />
              </div>

              <div style={{ display: 'flex', gap: '4px' }}>
                {['all', 'warnings', 'ready'].map(mode => (
                  <button
                    key={mode}
                    type="button"
                    onClick={() => { setFilterMode(mode); setPage(0); }}
                    style={{
                      padding: '3px 8px',
                      fontSize: '11px',
                      borderRadius: '12px',
                      border: '1px solid',
                      borderColor: filterMode === mode ? 'var(--accent)' : 'var(--line)',
                      background: filterMode === mode ? 'var(--accent)' : 'transparent',
                      color: filterMode === mode ? 'white' : 'var(--muted)',
                      fontWeight: 600,
                      textTransform: 'capitalize'
                    }}
                  >
                    {mode === 'all' ? `All (${normalizedRows.length})` : mode === 'warnings' ? `Warnings (${stats.warnings})` : `Clean (${stats.ready})`}
                  </button>
                ))}
              </div>
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

          {/* Master 40-Column Table Container */}
          <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
            {normalizedRows.length === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'var(--muted)', gap: '8px' }}>
                <FileSpreadsheet size={40} style={{ opacity: 0.3 }} />
                <strong style={{ fontSize: '14px', color: 'var(--ink)' }}>No Admission Allotment File Uploaded</strong>
                <span style={{ fontSize: '12px' }}>Upload your raw candidate spreadsheet from the left panel to begin.</span>
              </div>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px', whiteSpace: 'nowrap' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', position: 'sticky', top: 0, zIndex: 10 }}>
                    <th style={{ padding: '8px 10px', textAlign: 'center', borderBottom: '1.5px solid var(--line)', color: 'var(--muted)', width: '40px' }}>#</th>
                    {TARGET_COLUMNS.map((col, idx) => (
                      <th 
                        key={col || `col_${idx}`} 
                        style={{ 
                          padding: '8px 10px', 
                          textAlign: 'left', 
                          borderBottom: '1.5px solid var(--line)', 
                          color: 'var(--ink)', 
                          fontWeight: 700,
                          borderRight: '1px solid var(--line)',
                          background: 'var(--bg)'
                        }}
                      >
                        <span style={{ fontSize: '9px', color: 'var(--muted)', display: 'block', fontWeight: 600 }}>Col {idx + 1}</span>
                        {col || <span style={{ fontStyle: 'italic', color: 'var(--muted)' }}>(Passing Year)</span>}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {pagedRows.map((row, rowIdx) => {
                    const actualIdx = (page * pageSize) + rowIdx;
                    const hasWarn = row._warnings && row._warnings.length > 0;
                    return (
                      <tr 
                        key={actualIdx} 
                        style={{ 
                          borderBottom: '1px solid var(--line)',
                          background: hasWarn ? 'rgba(234, 179, 8, 0.03)' : 'transparent'
                        }}
                      >
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: 'var(--muted)', borderRight: '1px solid var(--line)', fontWeight: 600 }}>
                          {actualIdx + 1}
                        </td>
                        {TARGET_COLUMNS.map((col, cIdx) => (
                          <td 
                            key={col || `col_${cIdx}`} 
                            style={{ 
                              padding: '5px 8px', 
                              borderRight: '1px solid var(--line)',
                              color: 'var(--ink)'
                            }}
                          >
                            <input 
                              type="text" 
                              value={row[col] !== undefined ? row[col] : ''} 
                              onChange={(e) => updateCell(actualIdx, col, e.target.value)} 
                              style={{ 
                                width: '100%', 
                                border: 'none', 
                                background: 'transparent', 
                                fontSize: '11.5px', 
                                color: 'inherit',
                                padding: 0,
                                fontFamily: 'inherit'
                              }} 
                            />
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
          {normalizedRows.length > 0 && (
            <div style={{ padding: '8px 16px', background: 'var(--bg)', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
              <span>Showing {pagedRows.length} of {filteredRows.length} filtered candidate records • All {TARGET_COLUMNS.length} columns verified</span>
              <span>Click any table cell to perform inline edits before exporting</span>
            </div>
          )}

        </main>
      </div>

      {/* ID Lookups Definition & Management Modal */}
      <LookupIdManagerModal
        isOpen={isLookupModalOpen}
        onClose={() => setIsLookupModalOpen(false)}
        lookups={lookups}
        onSaveLookups={handleSaveLookups}
      />
    </div>
  );
}
