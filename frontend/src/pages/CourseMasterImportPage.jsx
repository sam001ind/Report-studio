import { useState, useMemo, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { 
  ArrowLeft, 
  Upload, 
  Download, 
  FileSpreadsheet, 
  Search, 
  BookOpen, 
  Copy, 
  Sparkles, 
  Plus, 
  Trash2, 
  GitFork,
  TableProperties,
  Edit3,
  X,
  ListFilter,
  ArrowUp, 
  ArrowDown, 
  ArrowUpDown,
  Sliders,
  Trash,
  Settings2
} from 'lucide-react';

const OUTPUT_HEADERS = [
  'UniqueProgramTermCode', 
  'ImmidiateParentGroup', 
  'GroupMaxCoursesforAdmission',
  'GroupMinCoursesforAdmission', 
  'GroupMaxCreditsforAdmission',
  'GroupMinCreditsforAdmission', 
  'GroupMaxMarks', 
  'GroupMinMarks',
  'GroupMaxCredits', 
  'GroupMinCredits', 
  'CourseCode', 
  'CourseName',
  'CourseShortName', 
  'CourseType', 
  'CourseLevel', 
  'Faculty',
  'Subject', 
  'FollowCreditSystem', 
  'Credits', 
  'CourseEvaluationSystem',
  'CourseMaxMarks', 
  'CourseMinMarks', 
  'CourseEvaluationTemplate',
  'TeachingLearningMethod', 
  'TeachingHours', 
  'AssessmentMethod',
  'AMEvaluationSystem', 
  'AMCredits', 
  'AMMaxMarks', 
  'AMMinMarks',
  'AMEvaluationTemplate', 
  'AssessmentType', 
  'ATEvaluationSystem',
  'ATCredits', 
  'ATMaxMarks', 
  'ATMinMarks', 
  'ATEvaluationTemplate'
];

const GROUP_MASTER_HEADERS = [
  'UniqueProgramTermCode',
  'ParentGroupName',
  'SubGroupName',
  'ParentGroupMinSubGroups',
  'ParentGroupMaxSubGroups',
  'ParentGroupEvaluationSystem',
  'ParentGroupMaxMarks',
  'ParentGroupMinMarks',
  'ParentGroupMaxCredits',
  'ParentGroupMinCredits'
];

const normalizeKey = (key) => String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');

const parseSheetWithSmartHeaders = (sheet) => {
  const rawGrid = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (!rawGrid || rawGrid.length === 0) return { parsed: [], headerMap: {} };
  if (rawGrid.length === 1) {
    const p = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const map = {};
    if (p.length) {
      Object.keys(p[0]).forEach(h => { map[normalizeKey(h)] = h; });
    }
    return { parsed: p, headerMap: map };
  }

  const row0 = rawGrid[0] || [];
  const row1 = rawGrid[1] || [];

  let courseCodeIdx = -1;
  row0.forEach((h, idx) => {
    const norm = normalizeKey(h);
    if (norm === 'coursecode' || norm === 'code') courseCodeIdx = idx;
  });

  const hasCourseCodeInData = courseCodeIdx !== -1 && row1[courseCodeIdx] && String(row1[courseCodeIdx]).trim().length > 1;

  const exactSubheaders = ['theory', 'practical', 'th', 'pr', 'totalmarks', 'minimumpassingmarks', 'minpassingmarks', 'passingmarks', 'maxmarks', 'minmarks'];
  const matchingSubheaders = row1.filter(cell => {
    const norm = normalizeKey(cell);
    return exactSubheaders.includes(norm);
  });

  const hasSubheaders = !hasCourseCodeInData && matchingSubheaders.length >= 2;

  if (!hasSubheaders) {
    const p = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    const map = {};
    if (p.length) {
      Object.keys(p[0]).forEach(h => { map[normalizeKey(h)] = h; });
    }
    return { parsed: p, headerMap: map };
  }

  const mergedHeaders = [];
  let lastParentHeader = '';

  const maxCols = Math.max(row0.length, row1.length);
  for (let c = 0; c < maxCols; c++) {
    const parent = String(row0[c] || '').trim();
    const child = String(row1[c] || '').trim();

    if (parent) {
      lastParentHeader = parent;
    }

    let colName;
    const normChild = normalizeKey(child);
    const normParent = normalizeKey(parent);

    if (normChild === 'theory' || normChild === 'th') {
      colName = lastParentHeader ? `${lastParentHeader} - TH` : 'Theory';
    } else if (normChild === 'practical' || normChild === 'pr') {
      colName = lastParentHeader ? `${lastParentHeader} - PR` : 'Practical';
    } else if (child && (!parent || normParent === normChild)) {
      colName = child;
    } else if (parent && !child) {
      colName = parent;
    } else if (parent && child) {
      colName = `${parent} - ${child}`;
    } else if (lastParentHeader && child) {
      colName = `${lastParentHeader} - ${child}`;
    } else {
      colName = `__col_${c}`;
    }

    mergedHeaders.push(colName);
  }

  const rows = [];
  for (let r = 2; r < rawGrid.length; r++) {
    const dataRow = rawGrid[r];
    if (!dataRow || dataRow.every(v => v === '' || v === undefined || v === null)) continue;
    const obj = {};
    mergedHeaders.forEach((h, idx) => {
      obj[h] = dataRow[idx] !== undefined ? dataRow[idx] : '';
    });
    rows.push(obj);
  }

  const map = {};
  mergedHeaders.forEach(h => {
    map[normalizeKey(h)] = h;
  });

  return { parsed: rows, headerMap: map };
};

const createDefaultConfigForGroup = (groupKey, defaultMarks = 100, defaultCredits = null) => {
  const upper = String(groupKey || '').toUpperCase().trim();
  const isMajor = upper.includes('DSC') || upper.includes('MAJOR') || upper.includes('CORE') || upper.includes('DSE');
  return {
    pattern: isMajor ? 'group_subject' : 'group_only',
    copies: isMajor ? 2 : 1,
    suffixStr: isMajor ? ', .' : '',
    maxCredits: defaultCredits !== null ? defaultCredits : (isMajor ? 4 : 3),
    maxMarks: defaultMarks || 100,
    maxCourses: 1,
    minCourses: 1
  };
};

export default function CourseMasterImportPage() {
  const [sourceFile, setSourceFile] = useState(null);
  const [rawRows, setRawRows] = useState([]);
  const [headerMap, setHeaderMap] = useState({});
  const [activeView, setActiveView] = useState('course_master'); // 'course_master' | 'group_master'
  const [useDuplication, setUseDuplication] = useState(true); // MASTER DUPLICATION TOGGLE
  const [groupConfigs, setGroupConfigs] = useState({});
  const [customGroupInput, setCustomGroupInput] = useState('');
  const [showAddGroupInput, setShowAddGroupInput] = useState(false);
  const [exportMode, setExportMode] = useState('zip_bundle'); // 'zip_bundle' | 'zip_courses' | 'zip_groups' | 'combined_courses' | 'combined_groups'
  const [selectedSubjectFilter, setSelectedSubjectFilter] = useState('ALL');
  const [selectedGroupFilter, setSelectedGroupFilter] = useState('ALL');
  const [searchQuery, setSearchQuery] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('info');
  const [page, setPage] = useState(0);
  const pageSize = 50;

  // Fully Customizable User-Controlled Group & Subgroup Hierarchy
  // Format: [ { id, groupName, subGroups: [ { id, name, pattern, suffix, subjects: [] } ] } ]
  const [groupHierarchy, setGroupHierarchy] = useState([]);
  const [newGroupNameInput, setNewGroupNameInput] = useState('');
  const [showNewGroupModal, setShowNewGroupModal] = useState(false);
  const [targetGroupIdForSubgroup, setTargetGroupIdForSubgroup] = useState(null);
  const [newSubgroupNameInput, setNewSubgroupNameInput] = useState('');
  const [newSubgroupPattern, setNewSubgroupPattern] = useState('group_subject');
  const [newSubgroupSuffix, setNewSubgroupSuffix] = useState('');

  // Subject Assignment Modal State
  const [editingSubgroup, setEditingSubgroup] = useState(null); // { groupId, subGroupId, subGroupName, groupName }
  const [subjectModalSearch, setSubjectModalSearch] = useState('');

  // 37-Column Deep Filtering States
  const [columnFilters, setColumnFilters] = useState({}); // { [colName]: string }
  const [showColumnFilterRow, setShowColumnFilterRow] = useState(false);

  // Column Sorting State: { column: string | null, direction: 'asc' | 'desc' | null }
  const [sortConfig, setSortConfig] = useState({ column: null, direction: null });

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

  const clearAllColumnFilters = () => {
    setColumnFilters({});
    setSortConfig({ column: null, direction: null });
    setSelectedSubjectFilter('ALL');
    setSelectedGroupFilter('ALL');
    setSearchQuery('');
    setPage(0);
  };

  const setStatus = (msg, type = 'info') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  const getCell = (row, ...aliases) => {
    for (const alias of aliases) {
      const norm = normalizeKey(alias);
      const actualKey = headerMap[norm];
      if (actualKey && row[actualKey] !== undefined && row[actualKey] !== null) {
        return String(row[actualKey]).trim();
      }
    }
    return '';
  };

  const getNumber = (row, ...aliases) => {
    const val = getCell(row, ...aliases);
    const num = Number(val);
    return isNaN(num) ? 0 : num;
  };

  // Extract all unique group names, subjects, and group-to-subjects map dynamically from uploaded file
  // Extract all unique group names, subjects, group-to-subjects map, and marks/credits map dynamically from uploaded file
  const { detectedGroups, uniqueSubjects, groupToSubjectsMap, groupMarksMap, groupCreditsMap } = useMemo(() => {
    if (!rawRows.length) return { detectedGroups: [], uniqueSubjects: [], groupToSubjectsMap: {}, groupMarksMap: {}, groupCreditsMap: {} };
    const groups = new Set();
    const subjects = new Set();
    const map = {};
    const marksMap = {};
    const creditsMap = {};

    rawRows.forEach(row => {
      const g = getCell(row, 'Group Name', 'groupname', 'group', 'parentgroup').trim().toUpperCase();
      const s = getCell(row, 'Subject', 'subject').trim();
      const tm = getNumber(row, 'Total Marks', 'totalmarks', 'marks', 'coursemarks', 'coursemaxmarks', 'maxmarks');
      const tc = getNumber(row, 'Total Credits', 'totalcredits', 'credits', 'credit');

      if (g) {
        groups.add(g);
        if (!map[g]) map[g] = new Set();
        if (s) map[g].add(s);
        if (tm > 0) marksMap[g] = Math.max(marksMap[g] || 0, tm);
        if (tc > 0) creditsMap[g] = Math.max(creditsMap[g] || 0, tc);
      }
      if (s) subjects.add(s);
    });

    const finalMap = {};
    Object.keys(map).forEach(k => {
      finalMap[k] = Array.from(map[k]).sort();
    });

    return {
      detectedGroups: Array.from(groups).sort(),
      uniqueSubjects: Array.from(subjects).sort(),
      groupToSubjectsMap: finalMap,
      groupMarksMap: marksMap,
      groupCreditsMap: creditsMap
    };
  }, [rawRows, headerMap]);

  // Synchronize dynamic group configurations whenever a new file is uploaded
  useEffect(() => {
    if (!detectedGroups.length) return;
    setGroupConfigs(prev => {
      const next = { ...prev };
      detectedGroups.forEach(g => {
        if (!next[g]) {
          next[g] = createDefaultConfigForGroup(g, groupMarksMap[g] || 100, groupCreditsMap[g] || 4);
        }
      });
      return next;
    });
  }, [detectedGroups, groupMarksMap, groupCreditsMap]);

  // 1. MASTER COURSE FILE TRANSFORMATION PIPELINE (37 COLUMNS)
  const generateProcessedRows = (forSubject = null, forGroup = null) => {
    if (!rawRows.length) return [];
    const rowsForSubject = [];

    const targetRows = rawRows.filter(r => {
      const s = getCell(r, 'Subject', 'subject').trim().toLowerCase();
      const g = getCell(r, 'Group Name', 'groupname', 'group', 'parentgroup').trim().toUpperCase();
      if (forSubject && s !== forSubject.trim().toLowerCase()) return false;
      if (forGroup && g !== forGroup.trim().toUpperCase()) return false;
      return true;
    });

    targetRows.forEach(row => {
      const subject = getCell(row, 'Subject', 'subject');
      const rawGroup = getCell(row, 'Group Name', 'groupname', 'group', 'parentgroup');
      const groupKey = rawGroup.trim().toUpperCase();

      if (!subject) return;

      let eseMaxTh = getNumber(row, 'ESE Max - TH', 'esemaxth', 'eseth', 'esemax_th', 'esetheory', 'theoryesemax', 'esemax');
      let ccaMaxTh = getNumber(row, 'CCA Max - TH', 'ccamaxth', 'ccath', 'ccamax_th', 'cemaxth', 'ccatheory', 'ceth', 'ccamax', 'cemax');
      let eseMaxPr = getNumber(row, 'ESE Max - PR', 'esemaxpr', 'esepr', 'esemax_pr', 'esepractical', 'practicalesemax');
      let ccaMaxPr = getNumber(row, 'CCA Max - PR', 'ccamaxpr', 'ccapr', 'ccamax_pr', 'cemaxpr', 'ccapractical', 'cepr');

      let eseMinTh = getNumber(row, 'ESE Min - TH', 'eseminth', 'esethmin', 'esemin_th', 'esetheorymin', 'esemin');
      let ccaMinTh = getNumber(row, 'CCA Min - TH', 'ccaminth', 'ccathmin', 'ccamin_th', 'ceminth', 'cethmin', 'ccatheorymin', 'ccamin', 'cemin');
      let eseMinPr = getNumber(row, 'ESE Min - PR', 'eseminpr', 'eseprmin', 'esemin_pr', 'esepracticalmin');
      let ccaMinPr = getNumber(row, 'CCA Min - PR', 'ccaminpr', 'ccaprmin', 'ccamin_pr', 'ceminpr', 'ceprmin', 'ccapracticalmin');

      const thCredits = getNumber(row, 'Theory Credits', 'theorycredits', 'thcredits', 'theory credit', 'theorycredit');
      const prCredits = getNumber(row, 'Practical Credits', 'practicalcredits', 'prcredits', 'practical credit', 'practicalcredit');
      const totalCredits = getNumber(row, 'Total Credits', 'totalcredits', 'credits', 'credit');

      // Auto-correct any column-order shift from older misaligned template downloads
      if (prCredits === 0 && thCredits > 0) {
        if (eseMinTh === 0 && eseMaxPr > 0) eseMinTh = eseMaxPr;
        if (ccaMinTh === 0 && ccaMaxPr > 0) ccaMinTh = ccaMaxPr;
        eseMaxPr = 0;
        ccaMaxPr = 0;
        eseMinPr = 0;
        ccaMinPr = 0;
      } else if (thCredits === 0 && prCredits > 0) {
        if (eseMaxPr === 0 && eseMinTh > 0) eseMaxPr = eseMinTh;
        if (ccaMaxPr === 0 && ccaMinTh > 0) ccaMaxPr = ccaMinTh;
        eseMaxTh = 0;
        ccaMaxTh = 0;
        eseMinTh = 0;
        ccaMinTh = 0;
      }

      // Determine strictly whether TH or PR components exist
      let hasThAssessments = false;
      let hasPrAssessments = false;

      if (thCredits > 0 && prCredits === 0) {
        hasThAssessments = true;
        hasPrAssessments = false;
      } else if (prCredits > 0 && thCredits === 0) {
        hasThAssessments = false;
        hasPrAssessments = true;
      } else if (thCredits > 0 && prCredits > 0) {
        hasThAssessments = true;
        hasPrAssessments = true;
      } else {
        const hasThMarks = eseMaxTh > 0 || ccaMaxTh > 0 || eseMinTh > 0 || ccaMinTh > 0;
        const hasPrMarks = eseMaxPr > 0 || ccaMaxPr > 0 || eseMinPr > 0 || ccaMinPr > 0;
        if (hasThMarks && !hasPrMarks) {
          hasThAssessments = true;
          hasPrAssessments = false;
        } else if (hasPrMarks && !hasThMarks) {
          hasThAssessments = false;
          hasPrAssessments = true;
        } else if (hasThMarks && hasPrMarks) {
          hasThAssessments = true;
          hasPrAssessments = true;
        } else {
          const cName = getCell(row, 'Course Name', 'coursename', 'title').toLowerCase();
          if (cName.includes('(practical)') || cName.includes('practical')) {
            hasPrAssessments = true;
            hasThAssessments = false;
          } else {
            hasThAssessments = true;
            hasPrAssessments = false;
          }
        }
      }

      let combinations = [];
      if (hasThAssessments) {
        combinations.push(['ESE', 'TH'], ['CE', 'TH']);
      }
      if (hasPrAssessments) {
        combinations.push(['ESE', 'PR'], ['CE', 'PR']);
      }
      if (combinations.length === 0) {
        combinations.push(['ESE', 'TH'], ['CE', 'TH']);
      }

      const cfg = groupConfigs[groupKey] || createDefaultConfigForGroup(groupKey);
      const copies = useDuplication ? Math.max(1, Number(cfg.copies) || 1) : 1;
      const rawSuffixes = (cfg.suffixStr || '').split(',').map(s => s.trim());
      const suffixList = [];
      for (let i = 0; i < copies; i++) {
        suffixList.push(rawSuffixes[i] !== undefined ? rawSuffixes[i] : (i === 0 ? '' : ` ${i + 1}`));
      }

      combinations.forEach(([amMethod, atType]) => {
        let totalMarks = getNumber(row, 'Total Marks', 'totalmarks', 'marks', 'coursemarks', 'coursemaxmarks', 'maxmarks');
        let minMarks = getNumber(row, 'Minimum Passing Marks', 'minimumpassingmarks', 'minpassingmarks', 'minmarks', 'courseminmarks');

        if (totalMarks === 0) {
          const sumAssessments = eseMaxTh + eseMaxPr + ccaMaxTh + ccaMaxPr;
          totalMarks = sumAssessments > 0 ? sumAssessments : (totalCredits > 0 ? totalCredits * 25 : 0);
        }
        if (minMarks === 0 && totalMarks > 0) {
          const sumMinAssessments = eseMinTh + eseMinPr + ccaMinTh + ccaMinPr;
          minMarks = sumMinAssessments > 0 ? sumMinAssessments : Math.round(totalMarks * 0.4);
        }

        const amCreditsFromSheet = getNumber(row, 'AM Credits', 'amcredits', 'amcredit', 'am credit');
        const atCredits = atType === 'TH' ? (thCredits || totalCredits) : (prCredits || totalCredits);
        const amCreditsVal = amCreditsFromSheet > 0 ? amCreditsFromSheet : atCredits;

        for (let copyIdx = 0; copyIdx < copies; copyIdx++) {
          const newRow = {};
          newRow['UniqueProgramTermCode'] = '';

          let baseGroup = rawGroup || 'General';
          if (cfg.pattern === 'group_subject') {
            baseGroup = `${rawGroup.trim()} - ${subject.trim()}`;
          }

          const currentSuffix = suffixList[copyIdx] || '';
          const immidiateParentGroup = `${baseGroup}${currentSuffix}`;
          newRow['ImmidiateParentGroup'] = immidiateParentGroup;

          newRow['GroupMaxCoursesforAdmission'] = Number(cfg.maxCourses) || 1;
          newRow['GroupMinCoursesforAdmission'] = Number(cfg.minCourses) || 1;
          newRow['GroupMaxCreditsforAdmission'] = cfg.maxCredits > 0 ? Number(cfg.maxCredits) : (totalCredits || 0);
          newRow['GroupMaxMarks'] = totalMarks; // GroupMaxMarks is same as CourseMaxMarks
          newRow['GroupMinMarks'] = 0;
          newRow['GroupMinCreditsforAdmission'] = 0;
          newRow['GroupMaxCredits'] = newRow['GroupMaxCreditsforAdmission'];
          newRow['GroupMinCredits'] = 0;

          newRow['CourseCode'] = getCell(row, 'Course Code', 'coursecode', 'code');
          newRow['CourseName'] = getCell(row, 'Course Name', 'coursename', 'title');
          newRow['CourseShortName'] = newRow['CourseCode'];
          newRow['CourseType'] = 'General';
          newRow['CourseLevel'] = 'General';
          newRow['Faculty'] = getCell(row, 'Faculty', 'faculty');
          newRow['Subject'] = subject;
          newRow['FollowCreditSystem'] = 'Yes';
          newRow['Credits'] = totalCredits;
          newRow['CourseEvaluationSystem'] = 'Indirect Grade System';
          newRow['CourseMaxMarks'] = totalMarks;
          newRow['CourseMinMarks'] = minMarks;
          newRow['CourseEvaluationTemplate'] = 'Eight Level';
          newRow['TeachingLearningMethod'] = 'Lec-Lab';
          newRow['TeachingHours'] = totalCredits === 3 ? 60 : (totalCredits === 4 ? 75 : 0);

          newRow['AssessmentMethod'] = amMethod;
          newRow['AssessmentType'] = atType;
          newRow['AMEvaluationSystem'] = 'Marks System';
          newRow['AMCredits'] = amCreditsVal;
          newRow['AMEvaluationTemplate'] = 'Eight Level';
          newRow['ATEvaluationSystem'] = 'Marks System';
          newRow['ATCredits'] = atCredits;
          newRow['ATEvaluationTemplate'] = 'Eight Level';

          let amMaxMarksVal = 0;
          let amMinMarksVal = 0;
          let atMaxMarksVal = 0;

          if (amMethod === 'ESE') {
            amMaxMarksVal = (hasThAssessments ? eseMaxTh : 0) + (hasPrAssessments ? eseMaxPr : 0);
            amMinMarksVal = (hasThAssessments ? eseMinTh : 0) + (hasPrAssessments ? eseMinPr : 0);
            atMaxMarksVal = atType === 'TH' ? eseMaxTh : eseMaxPr;
          } else if (amMethod === 'CE') {
            amMaxMarksVal = (hasThAssessments ? ccaMaxTh : 0) + (hasPrAssessments ? ccaMaxPr : 0);
            amMinMarksVal = (hasThAssessments ? ccaMinTh : 0) + (hasPrAssessments ? ccaMinPr : 0);
            atMaxMarksVal = atType === 'TH' ? ccaMaxTh : ccaMaxPr;
          }

          newRow['ATMaxMarks'] = atMaxMarksVal;
          newRow['ATMinMarks'] = 0; // No pass condition for AT, pass condition is only at AM level
          newRow['AMMaxMarks'] = amMaxMarksVal;
          newRow['AMMinMarks'] = amMinMarksVal;

          rowsForSubject.push(newRow);
        }
      });
    });

    rowsForSubject.sort((a, b) => {
      const order = (group) => {
        if (group.startsWith('DSC') && !group.endsWith('.')) return 1;
        if (group.startsWith('DSC') && group.endsWith('.')) return 2;
        if (group.startsWith('MDC')) return 3;
        if (group.startsWith('VAC')) return 4;
        if (group.startsWith('AEC')) return 5;
        if (group.startsWith('SEC')) return 6;
        return 7;
      };
      return order(a.ImmidiateParentGroup) - order(b.ImmidiateParentGroup);
    });

    return rowsForSubject;
  };

  // Extract all unique ImmidiateParentGroup values produced by Course Master rules
  const allImmediateParentGroups = useMemo(() => {
    if (!rawRows.length) return [];
    const rows = generateProcessedRows(null, null);
    const set = new Set();
    rows.forEach(r => {
      if (r.ImmidiateParentGroup) set.add(r.ImmidiateParentGroup);
    });
    return Array.from(set).sort();
  }, [rawRows, headerMap, groupConfigs, useDuplication]);

  const getImmediateParentGroupsForGroup = (groupName) => {
    if (!groupName) return allImmediateParentGroups;
    const upper = groupName.trim().toUpperCase();
    const matched = allImmediateParentGroups.filter(ipg => 
      ipg.toUpperCase() === upper ||
      ipg.toUpperCase().startsWith(`${upper} -`) ||
      ipg.toUpperCase().startsWith(`${upper} `) ||
      ipg.toUpperCase().startsWith(`${upper}.`)
    );
    return matched.length > 0 ? matched : allImmediateParentGroups;
  };

  // Helper to build default standard hierarchy automatically using ImmidiateParentGroup
  const autoPopulateHierarchy = (targetSubj = null) => {
    if (!detectedGroups.length) return;
    const activeSubj = targetSubj || (selectedSubjectFilter !== 'ALL' ? selectedSubjectFilter : uniqueSubjects[0]);
    const nextHierarchy = [];

    detectedGroups.forEach(gKey => {
      const cfg = groupConfigs[gKey] || createDefaultConfigForGroup(gKey);
      const groupIPGs = getImmediateParentGroupsForGroup(gKey);
      const copies = useDuplication ? Math.max(1, Number(cfg.copies) || 1) : 1;

      const subGroups = [];

      if (copies === 1 && cfg.pattern === 'group_only') {
        subGroups.push({
          id: `${gKey}_single`,
          name: gKey,
          pattern: 'group_only',
          suffix: '',
          subjects: groupIPGs.length ? groupIPGs : [gKey]
        });
      } else {
        // Find matching primary IPG for active subject
        const primaryMatch = groupIPGs.find(ipg => 
          ipg.toLowerCase().includes(activeSubj.toLowerCase()) && 
          !ipg.endsWith('.') && 
          !ipg.endsWith(' 2')
        ) || groupIPGs[0] || `${gKey} - ${activeSubj}`;

        // Bucket 1 (Primary / Target Subject IPG)
        subGroups.push({
          id: `${gKey}_1`,
          name: `${gKey} - 1`,
          pattern: cfg.pattern,
          suffix: '',
          subjects: [primaryMatch]
        });

        // Repetition buckets (DSC - 2, DSC - 3, etc.)
        for (let i = 1; i < (copies + 1); i++) {
          let bucketIPGs;
          if (i === 1) {
            // DSC - 2: Minor subjects (no dot / 1st repetition)
            bucketIPGs = groupIPGs.filter(ipg => 
              ipg !== primaryMatch && 
              !ipg.endsWith('.') && 
              !ipg.endsWith(' 2') && 
              !ipg.endsWith(' M2')
            );
          } else if (i === 2) {
            // DSC - 3: Minor subjects with dot / 2nd repetition
            bucketIPGs = groupIPGs.filter(ipg => 
              ipg.endsWith('.') || 
              ipg.endsWith(' 2') || 
              ipg.endsWith(' M2')
            );
          } else {
            bucketIPGs = groupIPGs.filter(ipg => ipg !== primaryMatch);
          }

          subGroups.push({
            id: `${gKey}_${i + 1}`,
            name: `${gKey} - ${i + 1}`,
            pattern: cfg.pattern,
            suffix: '',
            subjects: bucketIPGs
          });
        }
      }

      nextHierarchy.push({
        id: `g_${gKey}`,
        groupName: gKey,
        subGroups
      });
    });

    setGroupHierarchy(nextHierarchy);
    setStatus(`Auto-populated ${nextHierarchy.length} group(s) with ImmidiateParentGroups!`, 'success');
  };

  // Apply Quick Presets to all detected groups and rebuild groupHierarchy
  const applyPreset = (presetKey) => {
    if (!detectedGroups.length) return;
    const nextConfigs = { ...groupConfigs };
    let dupFlag = true;

    if (presetKey === 'fyugp_dot') {
      dupFlag = true;
      setUseDuplication(true);
      detectedGroups.forEach(g => {
        const marks = groupMarksMap[g] || 100;
        const creds = groupCreditsMap[g] || ((g.includes('DSC') || g.includes('MAJOR') || g.includes('CORE') || g.includes('DSE')) ? 4 : 3);
        if (g.includes('DSC') || g.includes('MAJOR') || g.includes('CORE') || g.includes('DSE')) {
          nextConfigs[g] = { pattern: 'group_subject', copies: 2, suffixStr: ', .', maxCredits: creds, maxMarks: marks, maxCourses: 1, minCourses: 1 };
        } else {
          nextConfigs[g] = { pattern: 'group_only', copies: 1, suffixStr: '', maxCredits: creds, maxMarks: marks, maxCourses: 1, minCourses: 1 };
        }
      });
      setGroupConfigs(nextConfigs);
      setStatus('Applied: Standard Pattern (Subj & .)', 'success');
    } else if (presetKey === 'numbered_series') {
      dupFlag = true;
      setUseDuplication(true);
      detectedGroups.forEach(g => {
        const marks = groupMarksMap[g] || 100;
        const isMaj = g.includes('DSC') || g.includes('MAJOR');
        const creds = groupCreditsMap[g] || (isMaj ? 4 : 3);
        nextConfigs[g] = {
          pattern: isMaj ? 'group_subject' : 'group_only',
          copies: 2,
          suffixStr: ' 1,  2',
          maxCredits: creds,
          maxMarks: marks,
          maxCourses: 1,
          minCourses: 1
        };
      });
      setGroupConfigs(nextConfigs);
      setStatus('Applied: Numbers (1, 2) Series', 'success');
    } else if (presetKey === 'major_minor') {
      dupFlag = true;
      setUseDuplication(true);
      detectedGroups.forEach(g => {
        const marks = groupMarksMap[g] || 100;
        const isMaj = g.includes('DSC') || g.includes('MAJOR');
        const creds = groupCreditsMap[g] || (isMaj ? 4 : 3);
        if (isMaj) {
          nextConfigs[g] = { pattern: 'group_subject', copies: 2, suffixStr: ' M1,  M2', maxCredits: creds, maxMarks: marks, maxCourses: 1, minCourses: 1 };
        } else {
          nextConfigs[g] = { pattern: 'group_only', copies: 2, suffixStr: ' 1,  2', maxCredits: creds, maxMarks: marks, maxCourses: 1, minCourses: 1 };
        }
      });
      setGroupConfigs(nextConfigs);
      setStatus('Applied: Multipliers (M1, M2)', 'success');
    } else if (presetKey === 'clean_single') {
      dupFlag = false;
      setUseDuplication(false);
      detectedGroups.forEach(g => {
        const marks = groupMarksMap[g] || 100;
        const isMaj = g.includes('DSC') || g.includes('MAJOR');
        const creds = groupCreditsMap[g] || (isMaj ? 4 : 3);
        nextConfigs[g] = {
          pattern: isMaj ? 'group_subject' : 'group_only',
          copies: 1,
          suffixStr: '',
          maxCredits: creds,
          maxMarks: marks,
          maxCourses: 1,
          minCourses: 1
        };
      });
      setGroupConfigs(nextConfigs);
      setStatus('Applied: Clean Single (No Duplications)', 'success');
    }

    // Populate groupHierarchy accordingly
    const activeSubj = selectedSubjectFilter !== 'ALL' ? selectedSubjectFilter : uniqueSubjects[0];
    const nextHierarchy = [];

    detectedGroups.forEach(gKey => {
      const cfg = nextConfigs[gKey] || createDefaultConfigForGroup(gKey);
      const subjsForThisGroup = groupToSubjectsMap[gKey] || uniqueSubjects;
      const copies = dupFlag ? Math.max(1, Number(cfg.copies) || 1) : 1;
      const suffixes = dupFlag ? (cfg.suffixStr || '').split(',').map(s => s.trim()) : [''];
      const subGroups = [];

      if (copies === 1 && cfg.pattern === 'group_only') {
        subGroups.push({
          id: `${gKey}_single`,
          name: gKey,
          pattern: 'group_only',
          suffix: '',
          subjects: []
        });
      } else {
        // Bucket 1 (Primary / Target Subject)
        subGroups.push({
          id: `${gKey}_1`,
          name: `${gKey} - 1`,
          pattern: cfg.pattern,
          suffix: suffixes[0] || '',
          subjects: activeSubj ? [activeSubj] : (subjsForThisGroup.slice(0, 1))
        });

        // Repetition buckets (DSC - 2, DSC - 3, etc.)
        for (let i = 1; i < (copies + 1); i++) {
          const sfx = suffixes[i - 1] || '';
          subGroups.push({
            id: `${gKey}_${i + 1}`,
            name: `${gKey} - ${i + 1}`,
            pattern: cfg.pattern,
            suffix: sfx,
            subjects: subjsForThisGroup.filter(s => s.toLowerCase() !== (activeSubj || '').toLowerCase())
          });
        }
      }

      nextHierarchy.push({
        id: `g_${gKey}`,
        groupName: gKey,
        subGroups
      });
    });

    setGroupHierarchy(nextHierarchy);
  };

  const updateGroupConfig = (group, field, value) => {
    setGroupConfigs(prev => ({
      ...prev,
      [group]: {
        ...(prev[group] || createDefaultConfigForGroup(group)),
        [field]: value
      }
    }));
  };

  const applyPresetToGroup = (groupKey, type) => {
    let updatedCfg;
    if (type === 'subj_dot') {
      updatedCfg = { pattern: 'group_subject', copies: 2, suffixStr: ', .', maxCredits: 4, maxMarks: 100, maxCourses: 1, minCourses: 1 };
    } else if (type === 'numbers') {
      updatedCfg = { pattern: 'group_subject', copies: 2, suffixStr: ' 1,  2', maxCredits: 4, maxMarks: 100, maxCourses: 1, minCourses: 1 };
    } else if (type === 'multipliers') {
      updatedCfg = { pattern: 'group_subject', copies: 2, suffixStr: ' M1,  M2', maxCredits: 4, maxMarks: 100, maxCourses: 1, minCourses: 1 };
    } else if (type === 'subj_single') {
      updatedCfg = { pattern: 'group_subject', copies: 1, suffixStr: '', maxCredits: 4, maxMarks: 100, maxCourses: 1, minCourses: 1 };
    } else if (type === 'group_single') {
      updatedCfg = { pattern: 'group_only', copies: 1, suffixStr: '', maxCredits: 3, maxMarks: 75, maxCourses: 1, minCourses: 1 };
    }

    if (updatedCfg) {
      setGroupConfigs(prev => ({
        ...prev,
        [groupKey]: {
          ...(prev[groupKey] || createDefaultConfigForGroup(groupKey)),
          ...updatedCfg
        }
      }));
      setStatus(`Applied ${type} preset to Group "${groupKey}"!`, 'success');
    }
  };

  const handleAddCustomGroup = () => {
    const trimmed = customGroupInput.trim().toUpperCase();
    if (!trimmed) return;
    if (!groupConfigs[trimmed]) {
      setGroupConfigs(prev => ({
        ...prev,
        [trimmed]: createDefaultConfigForGroup(trimmed)
      }));
      setStatus(`Added custom group rule for "${trimmed}"`, 'success');
    }
    setCustomGroupInput('');
    setShowAddGroupInput(false);
  };

  // Group Hierarchy Operations
  const handleAddGroup = () => {
    const trimmed = newGroupNameInput.trim().toUpperCase();
    if (!trimmed) return;
    if (groupHierarchy.some(g => g.groupName === trimmed)) {
      alert(`Group "${trimmed}" already exists.`);
      return;
    }
    setGroupHierarchy(prev => [
      ...prev,
      {
        id: `g_${Date.now()}`,
        groupName: trimmed,
        subGroups: []
      }
    ]);
    setNewGroupNameInput('');
    setShowNewGroupModal(false);
    setStatus(`Added group "${trimmed}"`, 'success');
  };

  const handleRemoveGroup = (groupId) => {
    setGroupHierarchy(prev => prev.filter(g => g.id !== groupId));
  };

  const handleAddSubgroup = () => {
    const trimmed = newSubgroupNameInput.trim();
    if (!trimmed || !targetGroupIdForSubgroup) return;

    setGroupHierarchy(prev => prev.map(g => {
      if (g.id !== targetGroupIdForSubgroup) return g;
      return {
        ...g,
        subGroups: [
          ...g.subGroups,
          {
            id: `sg_${Date.now()}`,
            name: trimmed,
            pattern: newSubgroupPattern,
            suffix: newSubgroupSuffix,
            subjects: []
          }
        ]
      };
    }));

    setNewSubgroupNameInput('');
    setNewSubgroupSuffix('');
    setTargetGroupIdForSubgroup(null);
    setStatus(`Added sub-group "${trimmed}"`, 'success');
  };

  const handleRemoveSubgroup = (groupId, subGroupId) => {
    setGroupHierarchy(prev => prev.map(g => {
      if (g.id !== groupId) return g;
      return {
        ...g,
        subGroups: g.subGroups.filter(sg => sg.id !== subGroupId)
      };
    }));
  };

  const toggleSubjectInSubgroup = (groupId, subGroupId, ipgName) => {
    setGroupHierarchy(prev => prev.map(g => {
      if (g.id !== groupId) return g;
      return {
        ...g,
        subGroups: g.subGroups.map(sg => {
          if (sg.id !== subGroupId) return sg;
          const exists = (sg.subjects || []).includes(ipgName);
          return {
            ...sg,
            subjects: exists ? sg.subjects.filter(s => s !== ipgName) : [...(sg.subjects || []), ipgName]
          };
        })
      };
    }));
  };

  const selectAllForSubgroup = (groupId, subGroupId, groupName) => {
    const allIPGs = getImmediateParentGroupsForGroup(groupName);
    setGroupHierarchy(prev => prev.map(g => {
      if (g.id !== groupId) return g;
      return {
        ...g,
        subGroups: g.subGroups.map(sg => {
          if (sg.id !== subGroupId) return sg;
          return { ...sg, subjects: [...allIPGs] };
        })
      };
    }));
  };

  const clearAllForSubgroup = (groupId, subGroupId) => {
    setGroupHierarchy(prev => prev.map(g => {
      if (g.id !== groupId) return g;
      return {
        ...g,
        subGroups: g.subGroups.map(sg => {
          if (sg.id !== subGroupId) return sg;
          return { ...sg, subjects: [] };
        })
      };
    }));
  };

  // 2. GROUP MASTER HIERARCHY TRANSFORMATION PIPELINE (10 COLUMNS)
  // Completely driven by user-defined groupHierarchy state or fallback to dynamic generation
  const generateGroupMasterRows = (forSubject = null) => {
    if (!rawRows.length) return [];
    const result = [];

    const getGroupMarks = (groupName) => {
      const upper = String(groupName || '').toUpperCase().trim();
      for (const [k, marks] of Object.entries(groupMarksMap)) {
        if (upper === k || upper.startsWith(`${k} `) || upper.startsWith(`${k}-`) || upper.startsWith(`${k} -`)) {
          return marks;
        }
      }
      const cfg = groupConfigs[upper];
      if (cfg && cfg.maxMarks > 0) return Number(cfg.maxMarks);
      return 100;
    };

    const getGroupCredits = (groupName) => {
      const upper = String(groupName || '').toUpperCase().trim();
      for (const [k, creds] of Object.entries(groupCreditsMap)) {
        if (upper === k || upper.startsWith(`${k} `) || upper.startsWith(`${k}-`) || upper.startsWith(`${k} -`)) {
          return creds;
        }
      }
      const cfg = groupConfigs[upper];
      if (cfg && cfg.maxCredits > 0) return Number(cfg.maxCredits);
      return 4;
    };

    const addRow = (parent, sub, maxM, minM = 0, maxC, minC = 0, maxSub = 1, minSub = 1) => {
      const finalMaxM = (maxM !== undefined && maxM !== null) ? maxM : (getGroupMarks(sub) || getGroupMarks(parent) || 100);
      const finalMaxC = (maxC !== undefined && maxC !== null) ? maxC : (getGroupCredits(sub) || getGroupCredits(parent) || 4);
      result.push({
        UniqueProgramTermCode: '',
        ParentGroupName: parent,
        SubGroupName: sub,
        ParentGroupMinSubGroups: minSub,
        ParentGroupMaxSubGroups: maxSub,
        ParentGroupEvaluationSystem: 'Marks System',
        ParentGroupMaxMarks: finalMaxM,
        ParentGroupMinMarks: minM,
        ParentGroupMaxCredits: finalMaxC,
        ParentGroupMinCredits: minC
      });
    };

    if (groupHierarchy.length > 0) {
      // Use user-curated groupHierarchy
      groupHierarchy.forEach(g => {
        const grpMarks = getGroupMarks(g.groupName);
        const grpCredits = getGroupCredits(g.groupName);
        if (!g.subGroups || g.subGroups.length === 0) {
          addRow(g.groupName, g.groupName, grpMarks, 0, grpCredits, 0);
        } else {
          g.subGroups.forEach(sg => {
            const sgMarks = getGroupMarks(sg.name) || grpMarks;
            const sgCredits = getGroupCredits(sg.name) || grpCredits;
            addRow(g.groupName, sg.name, grpMarks, 0, grpCredits, 0);
            let itemsToRender = sg.subjects || [];
            if (forSubject) {
              if (sg.name.endsWith('- 1')) {
                // Find matching primary IPG for this subject
                const primaryMatch = allImmediateParentGroups.find(ipg => 
                  ipg.toLowerCase().includes(forSubject.toLowerCase()) && 
                  !ipg.endsWith('.') && 
                  !ipg.endsWith(' 2') &&
                  ipg.toUpperCase().startsWith(g.groupName.toUpperCase())
                );
                itemsToRender = primaryMatch ? [primaryMatch] : itemsToRender.slice(0, 1);
              } else {
                itemsToRender = itemsToRender.filter(item => 
                  !item.toLowerCase().includes(` ${forSubject.toLowerCase()}`) && 
                  !item.toLowerCase().includes(`-${forSubject.toLowerCase()}`)
                );
              }
            }
            itemsToRender.forEach(item => {
              addRow(sg.name, item, sgMarks, 0, sgCredits, 0);
            });
          });
        }
      });
    } else {
      // Fallback: build standard default structure dynamically
      detectedGroups.forEach(gKey => {
        const cfg = groupConfigs[gKey] || createDefaultConfigForGroup(gKey);
        const grpMarks = groupMarksMap[gKey] || (cfg.maxMarks > 0 ? Number(cfg.maxMarks) : 100);
        const grpCredits = groupCreditsMap[gKey] || (cfg.maxCredits > 0 ? Number(cfg.maxCredits) : 4);
        const copies = useDuplication ? Math.max(1, Number(cfg.copies) || 1) : 1;
        const rawSuffixes = (cfg.suffixStr || '').split(',').map(s => s.trim());
        const subjsForThisGroup = groupToSubjectsMap[gKey] || uniqueSubjects;
        const activeSubj = forSubject || (selectedSubjectFilter !== 'ALL' ? selectedSubjectFilter : subjsForThisGroup[0]);

        if (copies === 1 && cfg.pattern === 'group_only') {
          addRow(gKey, gKey, grpMarks, 0, grpCredits, 0);
        } else {
          for (let copyIdx = 0; copyIdx < copies; copyIdx++) {
            const bucketName = `${gKey} - ${copyIdx + 1}`;
            const sfx = rawSuffixes[copyIdx] || '';
            addRow(gKey, bucketName, grpMarks, 0, grpCredits, 0);

            if (cfg.pattern === 'group_subject') {
              if (copyIdx === 0) {
                if (activeSubj) addRow(bucketName, `${gKey} - ${activeSubj}${sfx}`, grpMarks, 0, grpCredits, 0);
              } else {
                subjsForThisGroup
                  .filter(s => s.toLowerCase() !== (activeSubj || '').toLowerCase())
                  .forEach(s => {
                    addRow(bucketName, `${gKey} - ${s}${sfx}`, grpMarks, 0, grpCredits, 0);
                  });
              }
            } else {
              addRow(bucketName, `${gKey}${sfx}`, grpMarks, 0, grpCredits, 0);
            }
          }
        }
      });
    }

    return result;
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setSourceFile(file);
    setIsProcessing(true);
    setStatus('Reading and extracting Group Names from spreadsheet...', 'info');

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const { parsed, headerMap: map } = parseSheetWithSmartHeaders(firstSheet);

      if (!parsed.length) throw new Error('No data found in uploaded sheet.');

      setHeaderMap(map);
      setRawRows(parsed);

      const foundGroups = new Set();
      parsed.forEach(row => {
        for (const alias of ['Group Name', 'groupname', 'group', 'parentgroup']) {
          const actualKey = map[normalizeKey(alias)];
          if (actualKey && row[actualKey]) {
            foundGroups.add(String(row[actualKey]).trim().toUpperCase());
            break;
          }
        }
      });

      const groupArray = Array.from(foundGroups).sort();
      const initialConfigs = {};
      groupArray.forEach(g => {
        initialConfigs[g] = createDefaultConfigForGroup(g);
      });
      setGroupConfigs(initialConfigs);
      setGroupHierarchy([]); // Keep it clean / blank as requested!

      setStatus(`Detected ${groupArray.length} unique Group(s): [${groupArray.join(', ')}] across ${parsed.length} rows!`, 'success');
    } catch (err) {
      console.error('Course Master Ingestion Error:', err);
      setStatus(`Upload failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Download standardized Syllabus / Course Master Input Template with pre-filled examples & instructions
  const handleDownloadTemplate = () => {
    try {
      const templateHeaders = [
        'Semester',
        'Faculty',
        'Subject',
        'Group Name',
        'Course Code',
        'Course Name',
        'Total Credits',
        'Theory Credits',
        'Practical Credits',
        'ESE Max - TH',
        'ESE Max - PR',
        'ESE Min - TH',
        'ESE Min - PR',
        'CCA Max - TH',
        'CCA Max - PR',
        'CCA Min - TH',
        'CCA Min - PR',
        'Total Marks',
        'Minimum Passing Marks'
      ];

      const sampleData = [
        [
          'Semester V',
          'Faculty of Social Sciences',
          'Anthropological Sciences',
          'DSC',
          'KU05DSCANT301',
          'Anthropological Theories (Theory Only Course)',
          4,
          4,
          0,
          50,
          0,
          20,
          0,
          50,
          0,
          20,
          0,
          100,
          40
        ],
        [
          'Semester V',
          'Faculty of Social Sciences',
          'Anthropological Sciences',
          'SEC',
          'KU05SECANT302',
          'Report Writing using Office tools (Practical Only Course)',
          3,
          0,
          3,
          0,
          50,
          0,
          20,
          0,
          50,
          0,
          20,
          100,
          40
        ],
        [
          'Semester V',
          'Faculty of Science',
          'Chemistry',
          'DSC',
          'KU05DSCCHE301',
          'Inorganic and Physical Chemistry (Theory + Practical Integrated)',
          4,
          3,
          1,
          50,
          15,
          20,
          0,
          25,
          10,
          10,
          0,
          100,
          40
        ],
        [
          'Semester V',
          'Faculty of Language and Literature',
          'English',
          'VAC',
          'KU05VACENG301',
          'Reading Literature in English (Elective / 3 Credits)',
          3,
          3,
          0,
          50,
          0,
          15,
          0,
          25,
          0,
          0,
          0,
          75,
          27
        ]
      ];

      const instructions = [
        ['Field Name', 'Required?', 'Sample Value', 'Purpose & Instructions'],
        ['Semester', 'Yes', 'Semester V', 'Semester or term identifier (e.g. Semester V, Semester 5).'],
        ['Faculty', 'Yes', 'Faculty of Social Sciences', 'Faculty or department offering the syllabus.'],
        ['Subject', 'Yes', 'Anthropological Sciences', 'The subject name. Courses are grouped by subject for individual exports.'],
        ['Group Name', 'Yes', 'DSC', 'Group code: DSC (Major/Core), DSE (Elective), SEC (Skill), VAC (Value Added), MDC (Multidisciplinary), AEC (Ability).'],
        ['Course Code', 'Yes', 'KU05DSCANT301', 'Official unique University course code.'],
        ['Course Name', 'Yes', 'Anthropological Theories', 'Full title of the paper or course.'],
        ['Total Credits', 'Yes', 4, 'Total credits for this course (usually 2, 3, or 4). Sum of Theory and Practical credits.'],
        ['Theory Credits', 'Optional', 4, 'Credits assigned for Theory lecture hours (leave 0 or blank if practical only).'],
        ['Practical Credits', 'Optional', 0, 'Credits assigned for Practical laboratory hours (leave 0 or blank if theory only).'],
        ['ESE Max - TH', 'Optional', 50, 'End Semester Exam maximum marks for Theory component.'],
        ['ESE Max - PR', 'Optional', 0, 'End Semester Exam maximum marks for Practical component (leave 0 or blank if theory only).'],
        ['ESE Min - TH', 'Optional', 20, 'Minimum marks required to pass ESE Theory (usually 40% of ESE Max).'],
        ['ESE Min - PR', 'Optional', 0, 'Minimum marks required to pass ESE Practical (leave 0 or blank if theory only).'],
        ['CCA Max - TH', 'Optional', 50, 'Continuous Assessment / Internal maximum marks for Theory component.'],
        ['CCA Max - PR', 'Optional', 0, 'Continuous Assessment / Internal maximum marks for Practical component (leave 0 or blank if theory only).'],
        ['CCA Min - TH', 'Optional', 20, 'Minimum passing marks for CCA Theory (if applicable, else 0).'],
        ['CCA Min - PR', 'Optional', 0, 'Minimum passing marks for CCA Practical (if applicable, else 0).'],
        ['Total Marks', 'Yes', 100, 'Overall maximum marks for this course (sum of all ESE + CCA components, e.g. 100 or 75).'],
        ['Minimum Passing Marks', 'Yes', 40, 'Minimum total marks required to pass the course (usually 40% of Total Marks).']
      ];

      const wb = XLSX.utils.book_new();

      // Sheet 1: Main Template (pre-filled with sample courses that user can edit or replace)
      const wsTemplate = XLSX.utils.aoa_to_sheet([templateHeaders, ...sampleData]);
      wsTemplate['!cols'] = templateHeaders.map(h => ({ wch: Math.max(h.length + 3, 14) }));
      wsTemplate['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: sampleData.length, c: templateHeaders.length - 1 } }) };
      XLSX.utils.book_append_sheet(wb, wsTemplate, 'Course Template');

      // Sheet 2: Field Instructions & Guide
      const wsGuide = XLSX.utils.aoa_to_sheet(instructions);
      wsGuide['!cols'] = [
        { wch: 24 },
        { wch: 12 },
        { wch: 28 },
        { wch: 75 }
      ];
      XLSX.utils.book_append_sheet(wb, wsGuide, 'Field Guide & Instructions');

      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Course_Master_Input_Template.xlsx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);

      setStatus('Downloaded Course Master Input Template (.xlsx)!', 'success');
    } catch (err) {
      console.error('Template Download Error:', err);
      alert('Failed to generate template: ' + err.message);
    }
  };

  // Master Comprehensive Export Dispatcher
  const executeExport = async () => {
    if (!rawRows.length) return alert('Please upload a course master sheet first.');
    setIsProcessing(true);
    setStatus('Generating customized workbooks...', 'info');

    try {
      if (exportMode === 'zip_bundle') {
        const zip = new JSZip();
        for (const subj of uniqueSubjects) {
          // 1. Subject Course Master
          const courseRows = generateProcessedRows(subj, null);
          const courseAoa = [OUTPUT_HEADERS, ...courseRows.map(r => OUTPUT_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))];
          const wbCourse = XLSX.utils.book_new();
          const wsCourse = XLSX.utils.aoa_to_sheet(courseAoa, { dense: true });
          wsCourse['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: courseAoa.length - 1, c: OUTPUT_HEADERS.length - 1 } }) };
          wsCourse['!cols'] = OUTPUT_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }));
          XLSX.utils.book_append_sheet(wbCourse, wsCourse, subj.slice(0, 30));
          const outCourse = XLSX.write(wbCourse, { bookType: 'xlsx', type: 'array' });
          zip.file(`${subj} Course Master.xlsx`, outCourse);

          // 2. Subject Group Master
          const groupRows = generateGroupMasterRows(subj);
          const groupAoa = [GROUP_MASTER_HEADERS, ...groupRows.map(r => GROUP_MASTER_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))];
          const wbGroup = XLSX.utils.book_new();
          const wsGroup = XLSX.utils.aoa_to_sheet(groupAoa, { dense: true });
          wsGroup['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: groupAoa.length - 1, c: GROUP_MASTER_HEADERS.length - 1 } }) };
          wsGroup['!cols'] = GROUP_MASTER_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }));
          XLSX.utils.book_append_sheet(wbGroup, wsGroup, `${subj} Groups`.slice(0, 30));
          const outGroup = XLSX.write(wbGroup, { bookType: 'xlsx', type: 'array' });
          zip.file(`${subj} Group Master.xlsx`, outGroup);
        }

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Complete_Master_Package_${uniqueSubjects.length}_Subjects.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported complete Course & Group Master bundle into ZIP!`, 'success');

      } else if (exportMode === 'zip_courses') {
        const zip = new JSZip();
        for (const subj of uniqueSubjects) {
          const rows = generateProcessedRows(subj, selectedGroupFilter === 'ALL' ? null : selectedGroupFilter);
          const aoa = [OUTPUT_HEADERS, ...rows.map(r => OUTPUT_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))];
          const wb = XLSX.utils.book_new();
          const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
          ws['!cols'] = OUTPUT_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }));
          XLSX.utils.book_append_sheet(wb, ws, subj.slice(0, 30));
          const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          zip.file(`${subj} Courses.xlsx`, out);
        }
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Course_Master_Subjects.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported Course Master ZIP!`, 'success');

      } else if (exportMode === 'zip_groups') {
        const zip = new JSZip();
        for (const subj of uniqueSubjects) {
          const rows = generateGroupMasterRows(subj);
          const aoa = [GROUP_MASTER_HEADERS, ...rows.map(r => GROUP_MASTER_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))];
          const wb = XLSX.utils.book_new();
          const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
          ws['!cols'] = GROUP_MASTER_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }));
          XLSX.utils.book_append_sheet(wb, ws, subj.slice(0, 30));
          const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          zip.file(`${subj} Group Master.xlsx`, out);
        }
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Group_Master_Subjects.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported Group Master ZIP!`, 'success');

      } else if (exportMode === 'combined_groups') {
        const targetSubj = selectedSubjectFilter === 'ALL' ? uniqueSubjects[0] : selectedSubjectFilter;
        const rows = generateGroupMasterRows(targetSubj);
        const aoa = [GROUP_MASTER_HEADERS, ...rows.map(r => GROUP_MASTER_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))];
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
        ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: GROUP_MASTER_HEADERS.length - 1 } }) };
        ws['!cols'] = GROUP_MASTER_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }));
        XLSX.utils.book_append_sheet(wb, ws, 'Group_Master');
        const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${targetSubj || 'All'}_Group_Master.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported Group Master XLSX (${rows.length} rows)!`, 'success');

      } else {
        const allRows = generateProcessedRows(
          selectedSubjectFilter === 'ALL' ? null : selectedSubjectFilter,
          selectedGroupFilter === 'ALL' ? null : selectedGroupFilter
        );
        const aoa = [OUTPUT_HEADERS, ...allRows.map(r => OUTPUT_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))];
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
        ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: OUTPUT_HEADERS.length - 1 } }) };
        ws['!cols'] = OUTPUT_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }));
        XLSX.utils.book_append_sheet(wb, ws, 'Course_Master');
        const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Course_Master_Combined.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported combined course master XLSX!`, 'success');
      }
    } catch (err) {
      console.error('Export Error:', err);
      setStatus(`Export failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Download ONLY the filtered subset currently shown
  const exportFilteredSubset = () => {
    if (!previewRows.length) return alert('No filtered rows to download.');
    setIsProcessing(true);
    setStatus('Exporting filtered subset...', 'info');

    try {
      const headers = activeView === 'group_master' ? GROUP_MASTER_HEADERS : OUTPUT_HEADERS;
      const aoa = [headers, ...previewRows.map(r => headers.map(h => r[h] !== undefined ? r[h] : ''))];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: headers.length - 1 } }) };
      ws['!cols'] = headers.map(h => ({ wch: Math.max(h.length + 2, 12) }));

      const sheetTitle = (activeView === 'group_master' ? 'Group_Master' : (selectedSubjectFilter !== 'ALL' ? selectedSubjectFilter : 'Filtered')).slice(0, 31);
      XLSX.utils.book_append_sheet(wb, ws, sheetTitle);

      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      
      const sName = selectedSubjectFilter !== 'ALL' ? selectedSubjectFilter.replace(/\s+/g, '_') : 'AllSubj';
      const fileType = activeView === 'group_master' ? 'Group_Master' : 'Course_Master';
      a.download = `${sName}_${fileType}_${useDuplication ? 'WithDup' : 'NoDup'}.xlsx`;
      
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setStatus(`Exported ${previewRows.length} filtered rows (${fileType})!`, 'success');
    } catch (err) {
      console.error('Filtered Export Error:', err);
      setStatus(`Export failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const previewRows = useMemo(() => {
    let generated;
    if (activeView === 'group_master') {
      const targetSubj = selectedSubjectFilter === 'ALL' ? null : selectedSubjectFilter;
      generated = generateGroupMasterRows(targetSubj);
    } else {
      const targetSubj = selectedSubjectFilter === 'ALL' ? null : selectedSubjectFilter;
      const targetGroup = selectedGroupFilter === 'ALL' ? null : selectedGroupFilter;
      generated = generateProcessedRows(targetSubj, targetGroup);
    }

    // 1. Apply global search query
    if (searchQuery) {
      const q = searchQuery.toLowerCase();
      generated = generated.filter(r => 
        Object.values(r).some(val => String(val || '').toLowerCase().includes(q))
      );
    }

    // 2. Apply Column-specific filters across all 37 columns
    const activeFilters = Object.entries(columnFilters);
    if (activeFilters.length > 0) {
      generated = generated.filter(r => {
        return activeFilters.every(([colKey, filterVal]) => {
          if (!filterVal || String(filterVal).trim() === '') return true;
          const cellVal = String(r[colKey] !== undefined && r[colKey] !== null ? r[colKey] : '').toLowerCase();
          return cellVal.includes(String(filterVal).toLowerCase().trim());
        });
      });
    }

    // 3. Apply Column Sorting across all headers (1 to 37)
    if (sortConfig.column && sortConfig.direction) {
      const col = sortConfig.column;
      const dir = sortConfig.direction === 'asc' ? 1 : -1;

      generated.sort((a, b) => {
        const valA = a[col] !== undefined && a[col] !== null ? a[col] : '';
        const valB = b[col] !== undefined && b[col] !== null ? b[col] : '';

        // Check if numeric
        const numA = Number(valA);
        const numB = Number(valB);
        const isNumA = typeof valA === 'number' || (String(valA).trim() !== '' && !isNaN(numA));
        const isNumB = typeof valB === 'number' || (String(valB).trim() !== '' && !isNaN(numB));

        if (isNumA && isNumB) {
          return (numA - numB) * dir;
        }

        // Natural alphanumeric comparison
        return String(valA).localeCompare(String(valB), undefined, { numeric: true, sensitivity: 'base' }) * dir;
      });
    }

    return generated;
  }, [rawRows, headerMap, groupConfigs, useDuplication, activeView, selectedSubjectFilter, selectedGroupFilter, searchQuery, groupHierarchy, columnFilters, sortConfig]);

  const pagedRows = useMemo(() => {
    const start = page * pageSize;
    return previewRows.slice(start, start + pageSize);
  }, [previewRows, page]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', overflow: 'hidden', background: 'var(--bg)', color: 'var(--ink)' }}>
      
      {/* Subject Assignment Modal for Sub-Groups */}
      {editingSubgroup && (
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
            width: '650px',
            maxHeight: '85vh',
            display: 'flex',
            flexDirection: 'column',
            overflow: 'hidden',
            boxShadow: '0 20px 40px rgba(0,0,0,0.25)'
          }}>
            {/* Modal Header */}
            <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--bg)' }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '15px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <ListFilter size={18} color="var(--accent)" /> Assign ImmidiateParentGroups to: <span style={{ color: 'var(--accent)' }}>{editingSubgroup.subGroupName}</span>
                </h3>
                <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                  Group: <strong>{editingSubgroup.groupName}</strong> • Pick from {getImmediateParentGroupsForGroup(editingSubgroup.groupName).length} generated ImmidiateParentGroup values
                </span>
              </div>
              <button 
                type="button" 
                className="secondary" 
                onClick={() => setEditingSubgroup(null)} 
                style={{ padding: '4px 8px', borderRadius: '50%' }}
              >
                <X size={16} />
              </button>
            </div>

            {/* Modal Filter & Action Toolbar */}
            <div style={{ padding: '10px 20px', borderBottom: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '10px', background: 'var(--panel)' }}>
              <div style={{ position: 'relative', flex: 1 }}>
                <input 
                  type="text" 
                  placeholder="Search ImmidiateParentGroup..." 
                  value={subjectModalSearch} 
                  onChange={(e) => setSubjectModalSearch(e.target.value)} 
                  style={{ width: '100%', padding: '5px 8px 5px 26px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)' }} 
                />
                <Search size={13} color="var(--muted)" style={{ position: 'absolute', left: '8px', top: '7px' }} />
              </div>

              <div style={{ display: 'flex', gap: '6px' }}>
                <button type="button" className="secondary" onClick={() => selectAllForSubgroup(editingSubgroup.groupId, editingSubgroup.subGroupId, editingSubgroup.groupName)} style={{ fontSize: '11px', padding: '4px 8px' }}>
                  Select All
                </button>
                <button type="button" className="secondary" onClick={() => clearAllForSubgroup(editingSubgroup.groupId, editingSubgroup.subGroupId)} style={{ fontSize: '11px', padding: '4px 8px' }}>
                  Clear All
                </button>
              </div>
            </div>

            {/* Modal ImmidiateParentGroup Checkboxes Grid */}
            <div style={{ flex: 1, overflowY: 'auto', padding: '16px 20px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
              {getImmediateParentGroupsForGroup(editingSubgroup.groupName)
                .filter(ipg => !subjectModalSearch || ipg.toLowerCase().includes(subjectModalSearch.toLowerCase()))
                .map(ipg => {
                  const currGroup = groupHierarchy.find(g => g.id === editingSubgroup.groupId);
                  const currSubgroup = currGroup?.subGroups?.find(sg => sg.id === editingSubgroup.subGroupId);
                  const isAssigned = (currSubgroup?.subjects || []).includes(ipg);
                  return (
                    <label 
                      key={ipg} 
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                        padding: '6px 10px',
                        borderRadius: '6px',
                        border: '1px solid',
                        borderColor: isAssigned ? 'var(--accent)' : 'var(--line)',
                        background: isAssigned ? 'var(--accent-soft)' : 'var(--bg)',
                        cursor: 'pointer',
                        fontSize: '11.5px'
                      }}
                    >
                      <input 
                        type="checkbox" 
                        checked={isAssigned} 
                        onChange={() => toggleSubjectInSubgroup(editingSubgroup.groupId, editingSubgroup.subGroupId, ipg)} 
                      />
                      <span style={{ flex: 1, color: 'var(--ink)', fontWeight: isAssigned ? 600 : 400 }}>
                        {ipg}
                      </span>
                    </label>
                  );
                })}
            </div>

            {/* Modal Footer */}
            <div style={{ padding: '12px 20px', borderTop: '1px solid var(--line)', background: 'var(--bg)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                Live updates reflected in Group Master table
              </span>
              <button 
                type="button" 
                onClick={() => setEditingSubgroup(null)}
                style={{ padding: '6px 18px', fontSize: '12px' }}
              >
                Done
              </button>
            </div>

          </div>
        </div>
      )}

      {/* Top Header Bar */}
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 24px', background: 'var(--panel)', borderBottom: '1px solid var(--line)', zIndex: 20 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}>
            <ArrowLeft size={16} /> Back to Portal
          </Link>
          <div style={{ height: '18px', width: '1px', background: 'var(--line)' }} />
          <h2 style={{ fontSize: '16px', margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
            <BookOpen size={18} color="var(--accent)" /> Course & Group Master Engine
            <span style={{ fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '12px', fontWeight: 700 }}>
              {useDuplication ? '✨ Duplication Active' : '📄 Single Mode'}
            </span>
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
            disabled={!rawRows.length || isProcessing}
            onClick={executeExport}
            style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 16px', fontSize: '12px' }}
          >
            <Download size={14} /> 
            {exportMode === 'zip_bundle' ? '📦 Download Master Bundle ZIP' : 
             exportMode === 'zip_courses' ? 'Download Course ZIP' : 
             exportMode === 'zip_groups' ? 'Download Group ZIP' : 
             exportMode === 'combined_groups' ? 'Download Group Master XLSX' : 'Export Combined XLSX'}
          </button>
        </div>
      </header>

      {/* Main Workspace Layout */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden', padding: '16px 24px', gap: '16px' }}>
        
        {/* Left Settings Sidebar */}
        <aside style={{ width: '380px', display: 'flex', flexDirection: 'column', gap: '14px', overflowY: 'auto' }}>
          
          {/* File Upload Dropzone */}
          <div className="card" style={{ padding: '16px', margin: 0, textAlign: 'center', border: '1.5px dashed var(--accent)', background: 'var(--accent-soft)' }}>
            <input 
              type="file" 
              id="courseMasterFileInput" 
              accept=".xlsx,.xls,.csv" 
              onChange={handleFileUpload} 
              style={{ display: 'none' }} 
            />
            <label htmlFor="courseMasterFileInput" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '6px' }}>
              <Upload size={26} color="var(--accent)" />
              <strong style={{ fontSize: '13px', color: 'var(--ink)' }}>
                {sourceFile ? sourceFile.name : 'Upload Course Master Sheet'}
              </strong>
              <span style={{ fontSize: '11px', color: 'var(--muted)' }}>
                {sourceFile ? `Loaded (${rawRows.length} courses)` : 'Extracts Group Names and lets you manage groups & buckets'}
              </span>
            </label>

            {/* DOWNLOAD INPUT TEMPLATE ACTION */}
            <div style={{ marginTop: '14px', paddingTop: '12px', borderTop: '1px dashed rgba(23,107,135,0.35)', display: 'flex', flexDirection: 'column', gap: '5px' }}>
              <button 
                type="button"
                onClick={handleDownloadTemplate}
                className="secondary"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: '6px',
                  padding: '7px 12px',
                  fontSize: '11.5px',
                  fontWeight: 600,
                  background: 'var(--panel)',
                  border: '1px solid var(--accent)',
                  color: 'var(--accent)',
                  borderRadius: '6px',
                  cursor: 'pointer',
                  boxShadow: '0 1px 3px rgba(0,0,0,0.05)'
                }}
                title="Download standard Excel template with pre-filled sample courses and field guide"
              >
                <Download size={13} /> Download Input Template (.xlsx)
              </button>
              <span style={{ fontSize: '10px', color: 'var(--muted)' }}>
                Includes sample rows (Theory, Practical, Integrated) + Field Guide
              </span>
            </div>
          </div>

          {/* MASTER DUPLICATION MODE SWITCH CARD */}
          <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '10px', border: '1.5px solid var(--accent)', background: 'linear-gradient(135deg, rgba(23,107,135,0.05), transparent)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3 style={{ margin: 0, fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700 }}>
                <Copy size={15} color="var(--accent)" /> Duplication Engine
              </h3>
              <span style={{ 
                fontSize: '10.5px', 
                padding: '2px 8px', 
                borderRadius: '10px', 
                fontWeight: 700, 
                background: useDuplication ? 'var(--accent)' : 'var(--muted)',
                color: 'white'
              }}>
                {useDuplication ? 'DUPLICATION ON' : 'DUPLICATION OFF'}
              </span>
            </div>

            <p style={{ margin: 0, fontSize: '11.5px', color: 'var(--muted)' }}>
              Controls repetition copies and group buckets (e.g. <code>DSC - 1, DSC - 2, VAC - 1, VAC - 2, MDC - 1, MDC - 2</code>):
            </p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
              <label style={{ 
                display: 'flex', 
                alignItems: 'flex-start', 
                gap: '8px', 
                cursor: 'pointer', 
                fontSize: '11.5px', 
                padding: '8px 10px', 
                borderRadius: '6px', 
                border: '1.5px solid', 
                borderColor: useDuplication ? 'var(--accent)' : 'var(--line)', 
                background: useDuplication ? 'var(--accent-soft)' : 'var(--bg)' 
              }}>
                <input 
                  type="radio" 
                  name="masterDuplicationOption" 
                  checked={useDuplication} 
                  onChange={() => {
                    setUseDuplication(true);
                    setStatus('Duplication enabled: Generating multiple group buckets and repetition copies.', 'info');
                  }} 
                  style={{ marginTop: '2px' }}
                />
                <div>
                  <strong style={{ display: 'block', color: 'var(--ink)' }}>✨ With Duplication (Default)</strong>
                  <span style={{ fontSize: '10.5px', color: 'var(--muted)' }}>
                    Generates multiple repetition copies (e.g. DSC - 1, DSC - 2, DSC - 3, VAC - 1, VAC - 2).
                  </span>
                </div>
              </label>

              <label style={{ 
                display: 'flex', 
                alignItems: 'flex-start', 
                gap: '8px', 
                cursor: 'pointer', 
                fontSize: '11.5px', 
                padding: '8px 10px', 
                borderRadius: '6px', 
                border: '1.5px solid', 
                borderColor: !useDuplication ? 'var(--accent)' : 'var(--line)', 
                background: !useDuplication ? 'var(--accent-soft)' : 'var(--bg)' 
              }}>
                <input 
                  type="radio" 
                  name="masterDuplicationOption" 
                  checked={!useDuplication} 
                  onChange={() => {
                    setUseDuplication(false);
                    setStatus('Duplication disabled: Generating single master rows and single group hierarchy.', 'info');
                  }} 
                  style={{ marginTop: '2px' }}
                />
                <div>
                  <strong style={{ display: 'block', color: 'var(--ink)' }}>📄 Without Duplication (Single Mode)</strong>
                  <span style={{ fontSize: '10.5px', color: 'var(--muted)' }}>
                    Single copy per course and clean single group entries without numbered repetition buckets.
                  </span>
                </div>
              </label>
            </div>
          </div>

          {/* COURSE MASTER: GROUP RULES & IMMIDIATEPARENTGROUP GENERATOR */}
          <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '12px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3 style={{ margin: 0, fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700 }}>
                <Sliders size={15} color="var(--accent)" /> Course Master: ImmidiateParentGroup Rules
              </h3>
              <button 
                type="button" 
                className="secondary" 
                onClick={() => setShowAddGroupInput(!showAddGroupInput)}
                style={{ padding: '2px 6px', fontSize: '10.5px', display: 'flex', alignItems: 'center', gap: '4px' }}
              >
                <Plus size={11} /> Add Group
              </button>
            </div>

            <p style={{ margin: 0, fontSize: '11px', color: 'var(--muted)' }}>
              Controls the naming pattern of <code>ImmidiateParentGroup</code> (Col 2 in Course Master):
            </p>

            {/* Optional Manual Add Group Input */}
            {showAddGroupInput && (
              <div style={{ display: 'flex', gap: '6px', padding: '6px', background: 'var(--bg)', borderRadius: '6px', border: '1px solid var(--line)' }}>
                <input 
                  type="text" 
                  placeholder="e.g. DSE or AEC" 
                  value={customGroupInput} 
                  onChange={(e) => setCustomGroupInput(e.target.value)} 
                  style={{ flex: 1, padding: '4px 6px', fontSize: '11.5px' }} 
                />
                <button type="button" onClick={handleAddCustomGroup} style={{ padding: '4px 8px', fontSize: '11px' }}>
                  Add
                </button>
              </div>
            )}

            {/* Detected Group Rules List */}
            {detectedGroups.length === 0 && Object.keys(groupConfigs).length === 0 ? (
              <div style={{ padding: '12px', background: 'var(--bg)', borderRadius: '6px', border: '1px dashed var(--line)', textAlign: 'center', color: 'var(--muted)', fontSize: '11.5px' }}>
                Upload an Excel sheet to extract Group Names from the source file.
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                {Object.keys(groupConfigs).map(groupKey => {
                  const cfg = groupConfigs[groupKey] || createDefaultConfigForGroup(groupKey);
                  return (
                    <div 
                      key={groupKey} 
                      style={{ 
                        padding: '8px 10px', 
                        background: 'var(--bg)', 
                        border: '1px solid var(--line)', 
                        borderRadius: '6px',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: '6px'
                      }}
                    >
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                        <strong style={{ fontSize: '12px', color: 'var(--accent)' }}>Group: {groupKey}</strong>
                        {useDuplication && (
                          <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                            <span style={{ fontSize: '10.5px', color: 'var(--muted)' }}>Copies:</span>
                            <input 
                              type="number" 
                              min={1} 
                              max={10} 
                              value={cfg.copies || 1} 
                              onChange={(e) => updateGroupConfig(groupKey, 'copies', Math.max(1, parseInt(e.target.value, 10) || 1))}
                              style={{ width: '42px', padding: '2px 4px', fontSize: '11px', textAlign: 'center' }} 
                            />
                          </div>
                        )}
                      </div>

                      {/* Smart Per-Group Quick Presets */}
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px', alignItems: 'center' }}>
                        <span style={{ fontSize: '9.5px', color: 'var(--muted)' }}>Smart Presets:</span>
                        <button 
                          type="button" 
                          className="secondary" 
                          onClick={() => applyPresetToGroup(groupKey, 'subj_single')}
                          style={{ padding: '1px 5px', fontSize: '9.5px', borderRadius: '4px', background: cfg.pattern === 'group_subject' && (!cfg.suffixStr || cfg.copies === 1) ? 'var(--accent-soft)' : 'var(--panel)', color: 'var(--ink)' }}
                          title={`Set ${groupKey} - Subject (Single copy)`}
                        >
                          {groupKey} - Subj
                        </button>
                        <button 
                          type="button" 
                          className="secondary" 
                          onClick={() => applyPresetToGroup(groupKey, 'group_single')}
                          style={{ padding: '1px 5px', fontSize: '9.5px', borderRadius: '4px', background: cfg.pattern === 'group_only' && (!cfg.suffixStr || cfg.copies === 1) ? 'var(--accent-soft)' : 'var(--panel)', color: 'var(--ink)' }}
                          title={`Set ${groupKey} Only (Single copy)`}
                        >
                          {groupKey} Only
                        </button>
                        <button 
                          type="button" 
                          className="secondary" 
                          onClick={() => applyPresetToGroup(groupKey, 'subj_dot')}
                          style={{ padding: '1px 5px', fontSize: '9.5px', borderRadius: '4px', background: cfg.pattern === 'group_subject' && cfg.suffixStr === ', .' ? 'var(--accent-soft)' : 'var(--panel)', color: 'var(--ink)' }}
                          title={`Set ${groupKey} - Subj & . (2 copies)`}
                        >
                          Subj & .
                        </button>
                        <button 
                          type="button" 
                          className="secondary" 
                          onClick={() => applyPresetToGroup(groupKey, 'numbers')}
                          style={{ padding: '1px 5px', fontSize: '9.5px', borderRadius: '4px', background: cfg.suffixStr?.includes('1') ? 'var(--accent-soft)' : 'var(--panel)', color: 'var(--ink)' }}
                          title={`Set ${groupKey} - Subj 1, 2 (2 copies)`}
                        >
                          1, 2
                        </button>
                        <button 
                          type="button" 
                          className="secondary" 
                          onClick={() => applyPresetToGroup(groupKey, 'multipliers')}
                          style={{ padding: '1px 5px', fontSize: '9.5px', borderRadius: '4px', background: cfg.suffixStr?.includes('M1') ? 'var(--accent-soft)' : 'var(--panel)', color: 'var(--ink)' }}
                          title={`Set ${groupKey} - Subj M1, M2 (2 copies)`}
                        >
                          M1, M2
                        </button>
                      </div>

                      <div style={{ display: 'grid', gridTemplateColumns: useDuplication ? '1.2fr 1.8fr' : '1fr', gap: '6px' }}>
                        <div className="form-group" style={{ margin: 0 }}>
                          <label style={{ fontSize: '10px' }}>Pattern</label>
                          <select 
                            value={cfg.pattern || 'group_only'} 
                            onChange={(e) => updateGroupConfig(groupKey, 'pattern', e.target.value)}
                            style={{ fontSize: '11px', padding: '2px 4px' }}
                          >
                            <option value="group_subject">{groupKey} - Subject</option>
                            <option value="group_only">{groupKey} Only</option>
                          </select>
                        </div>

                        {useDuplication && (
                          <div className="form-group" style={{ margin: 0 }}>
                            <label style={{ fontSize: '10px' }}>Suffixes (comma-separated)</label>
                            <input 
                              type="text" 
                              placeholder="e.g. , ." 
                              value={cfg.suffixStr !== undefined ? cfg.suffixStr : ''} 
                              onChange={(e) => updateGroupConfig(groupKey, 'suffixStr', e.target.value)}
                              style={{ fontSize: '11px', padding: '2px 4px' }} 
                            />
                          </div>
                        )}
                      </div>

                      {/* Live Preview of ImmidiateParentGroup */}
                      <div style={{ fontSize: '10px', color: 'var(--muted)', background: 'var(--panel)', padding: '3px 6px', borderRadius: '4px' }}>
                        ImmidiateParentGroup: {useDuplication ? (
                          (cfg.suffixStr || '').split(',').map((s, idx) => (
                            <span key={idx} style={{ fontWeight: 600, color: 'var(--ink)' }}>
                              {idx > 0 && ' | '}
                              {cfg.pattern === 'group_subject' ? `${groupKey} - Subj${s.trim()}` : `${groupKey}${s.trim()}`}
                            </span>
                          ))
                        ) : (
                          <span style={{ fontWeight: 600, color: 'var(--ink)' }}>
                            {cfg.pattern === 'group_subject' ? `${groupKey} - Subj` : `${groupKey}`}
                          </span>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>

          {/* GROUP MASTER (10 COLUMNS): FULLY CUSTOMIZABLE HIERARCHY MANAGER */}
          <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '12px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3 style={{ margin: 0, fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700 }}>
                <GitFork size={15} color="var(--accent)" /> Group Master: Hierarchy Manager
              </h3>
              <div style={{ display: 'flex', gap: '4px' }}>
                <button 
                  type="button" 
                  className="secondary" 
                  onClick={() => autoPopulateHierarchy()}
                  style={{ padding: '2px 6px', fontSize: '10.5px', color: 'var(--accent)' }}
                  title="Auto fill standard groups and buckets from uploaded file"
                >
                  ⚡ Auto-Fill
                </button>
                <button 
                  type="button" 
                  onClick={() => setShowNewGroupModal(!showNewGroupModal)}
                  style={{ padding: '2px 8px', fontSize: '10.5px', display: 'flex', alignItems: 'center', gap: '4px' }}
                >
                  <Plus size={11} /> Add Group
                </button>
              </div>
            </div>

            <p style={{ margin: 0, fontSize: '11px', color: 'var(--muted)' }}>
              Controls <code>ParentGroupName</code> & <code>SubGroupName</code> (in 10-Col Group Master):
            </p>

            {/* Quick Add Group Input Bar */}
            {showNewGroupModal && (
              <div style={{ display: 'flex', gap: '6px', padding: '8px', background: 'var(--bg)', borderRadius: '6px', border: '1px solid var(--line)' }}>
                {detectedGroups.length > 0 ? (
                  <select 
                    value={newGroupNameInput} 
                    onChange={(e) => setNewGroupNameInput(e.target.value)}
                    style={{ flex: 1, padding: '4px 6px', fontSize: '11.5px' }}
                  >
                    <option value="">-- Pick from Detected or Type --</option>
                    {detectedGroups.map(g => (
                      <option key={g} value={g}>{g}</option>
                    ))}
                  </select>
                ) : null}
                <input 
                  type="text" 
                  placeholder="e.g. DSC, AEC, MDC" 
                  value={newGroupNameInput} 
                  onChange={(e) => setNewGroupNameInput(e.target.value)} 
                  style={{ flex: 1, padding: '4px 6px', fontSize: '11.5px' }} 
                />
                <button type="button" onClick={handleAddGroup} style={{ padding: '4px 8px', fontSize: '11px' }}>
                  Save
                </button>
              </div>
            )}

            {/* Group Hierarchy Tree */}
            {groupHierarchy.length === 0 ? (
              <div style={{ padding: '14px', background: 'var(--bg)', borderRadius: '6px', border: '1px dashed var(--line)', textAlign: 'center', color: 'var(--muted)', fontSize: '11.5px' }}>
                <p style={{ margin: '0 0 8px 0' }}>No custom hierarchy defined yet (Starts blank for custom selection).</p>
                <div style={{ display: 'flex', justifyContent: 'center', gap: '6px' }}>
                  <button type="button" className="secondary" onClick={() => autoPopulateHierarchy()} style={{ fontSize: '11px', padding: '4px 10px' }}>
                    ⚡ Quick Auto-Fill Standard Groups
                  </button>
                  <button type="button" onClick={() => setShowNewGroupModal(true)} style={{ fontSize: '11px', padding: '4px 10px' }}>
                    <Plus size={12} style={{ display: 'inline', marginRight: '4px' }} /> Add Group
                  </button>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <span style={{ fontSize: '11px', color: 'var(--muted)' }}>
                    Configured {groupHierarchy.length} Group(s):
                  </span>
                  <button 
                    type="button" 
                    className="secondary" 
                    onClick={() => setGroupHierarchy([])} 
                    style={{ fontSize: '10px', padding: '2px 6px', color: 'var(--danger)' }}
                  >
                    Clear All
                  </button>
                </div>

                {groupHierarchy.map(group => (
                  <div 
                    key={group.id} 
                    style={{ 
                      background: 'var(--bg)', 
                      border: '1px solid var(--line)', 
                      borderRadius: '8px', 
                      padding: '10px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '8px'
                    }}
                  >
                    {/* Group Header */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <strong style={{ fontSize: '12.5px', color: 'var(--accent)' }}>
                        Group: {group.groupName}
                      </strong>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <button 
                          type="button" 
                          className="secondary" 
                          onClick={() => {
                            setTargetGroupIdForSubgroup(group.id);
                            setNewSubgroupNameInput(`${group.groupName} - ${(group.subGroups?.length || 0) + 1}`);
                          }}
                          style={{ padding: '2px 6px', fontSize: '10px', display: 'flex', alignItems: 'center', gap: '2px' }}
                        >
                          <Plus size={10} /> Add Subgroup
                        </button>
                        <button 
                          type="button" 
                          className="secondary" 
                          onClick={() => handleRemoveGroup(group.id)}
                          style={{ padding: '2px 4px', fontSize: '10px', color: 'var(--danger)' }}
                          title="Delete Group"
                        >
                          <Trash2 size={11} />
                        </button>
                      </div>
                    </div>

                    {/* Subgroup Add Inline Form */}
                    {targetGroupIdForSubgroup === group.id && (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', padding: '8px', background: 'var(--panel)', borderRadius: '6px', border: '1px solid var(--line)' }}>
                        <div style={{ display: 'flex', gap: '6px' }}>
                          <input 
                            type="text" 
                            placeholder="Subgroup name (e.g. DSC - 1, VAC - 1)" 
                            value={newSubgroupNameInput} 
                            onChange={(e) => setNewSubgroupNameInput(e.target.value)} 
                            style={{ flex: 1, padding: '3px 6px', fontSize: '11px' }} 
                          />
                          <select 
                            value={newSubgroupPattern} 
                            onChange={(e) => setNewSubgroupPattern(e.target.value)}
                            style={{ padding: '3px 6px', fontSize: '11px' }}
                          >
                            <option value="group_subject">Group - Subject</option>
                            <option value="group_only">Group Only</option>
                          </select>
                        </div>
                        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                          <input 
                            type="text" 
                            placeholder="Optional suffix (e.g. . or 1)" 
                            value={newSubgroupSuffix} 
                            onChange={(e) => setNewSubgroupSuffix(e.target.value)} 
                            style={{ width: '140px', padding: '3px 6px', fontSize: '11px' }} 
                          />
                          <div style={{ display: 'flex', gap: '4px' }}>
                            <button type="button" className="secondary" onClick={() => setTargetGroupIdForSubgroup(null)} style={{ padding: '2px 6px', fontSize: '10.5px' }}>
                              Cancel
                            </button>
                            <button type="button" onClick={handleAddSubgroup} style={{ padding: '2px 8px', fontSize: '10.5px' }}>
                              Add
                            </button>
                          </div>
                        </div>
                      </div>
                    )}

                    {/* Subgroups List */}
                    {(!group.subGroups || group.subGroups.length === 0) ? (
                      <div style={{ fontSize: '10.5px', color: 'var(--muted)', fontStyle: 'italic', padding: '4px' }}>
                        No subgroups. Maps directly as <code>{group.groupName}</code> $\rightarrow$ <code>{group.groupName}</code>.
                      </div>
                    ) : (
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                        {group.subGroups.map(sg => {
                          const subjCount = (sg.subjects || []).length;
                          return (
                            <div 
                              key={sg.id} 
                              style={{ 
                                display: 'flex', 
                                justifyContent: 'space-between', 
                                alignItems: 'center', 
                                padding: '5px 8px', 
                                background: 'var(--panel)', 
                                borderRadius: '4px', 
                                border: '1px solid var(--line)',
                                fontSize: '11px' 
                              }}
                            >
                              <div>
                                <strong style={{ color: 'var(--ink)' }}>{sg.name}</strong>
                                <span style={{ fontSize: '10px', color: 'var(--muted)', marginLeft: '4px' }}>
                                  ({subjCount} items)
                                </span>
                              </div>

                              <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                                <button 
                                  type="button" 
                                  className="secondary" 
                                  onClick={() => {
                                    setEditingSubgroup({
                                      groupId: group.id,
                                      subGroupId: sg.id,
                                      subGroupName: sg.name,
                                      groupName: group.groupName
                                    });
                                    setSubjectModalSearch('');
                                  }}
                                  style={{ padding: '2px 6px', fontSize: '10px', display: 'flex', alignItems: 'center', gap: '3px' }}
                                >
                                  <Edit3 size={10} /> Assign ({subjCount})
                                </button>
                                <button 
                                  type="button" 
                                  className="secondary" 
                                  onClick={() => handleRemoveSubgroup(group.id, sg.id)}
                                  style={{ padding: '2px 4px', fontSize: '10px', color: 'var(--danger)' }}
                                  title="Delete Subgroup"
                                >
                                  <Trash size={10} />
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}

                  </div>
                ))}
              </div>
            )}
          </div>

          {/* Quick Presets Bar (Active only when groups detected) */}
          {detectedGroups.length > 0 && (
            <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <div style={{ fontSize: '11.5px', textTransform: 'uppercase', color: 'var(--muted)', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '6px' }}>
                <Sparkles size={14} color="var(--accent)" /> Quick Group Presets
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' }}>
                <button type="button" className="secondary" onClick={() => applyPreset('fyugp_dot')} style={{ fontSize: '11px', padding: '5px 8px', textAlign: 'left' }}>
                  ⚡ Standard (Subj & .)
                </button>
                <button type="button" className="secondary" onClick={() => applyPreset('numbered_series')} style={{ fontSize: '11px', padding: '5px 8px', textAlign: 'left' }}>
                  ⚡ Numbers (1, 2)
                </button>
                <button type="button" className="secondary" onClick={() => applyPreset('major_minor')} style={{ fontSize: '11px', padding: '5px 8px', textAlign: 'left' }}>
                  ⚡ Multipliers (M1, M2)
                </button>
                <button type="button" className="secondary" onClick={() => applyPreset('clean_single')} style={{ fontSize: '11px', padding: '5px 8px', textAlign: 'left' }}>
                  ⚡ Single (No Dup)
                </button>
              </div>
            </div>
          )}

          {/* Export Output Package Format */}
          <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <h3 style={{ margin: 0, fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700 }}>
              <Settings2 size={15} color="var(--accent)" /> Output Package Mode
            </h3>
            
            <select value={exportMode} onChange={(e) => setExportMode(e.target.value)} style={{ fontSize: '12px' }}>
              <option value="zip_bundle">📦 Complete Bundle ZIP (Both Courses & Groups per Subject)</option>
              <option value="zip_courses">📑 Course Master ZIP Only (Separate .xlsx per Subject)</option>
              <option value="zip_groups">🗂️ Group Master ZIP Only (Separate .xlsx per Subject)</option>
              <option value="combined_groups">📑 Single Group Master File (.xlsx)</option>
              <option value="combined_courses">📚 Combined Course Master (.xlsx)</option>
            </select>
          </div>

        </aside>

        {/* Center Live Validation & Preview Grid */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '10px', overflow: 'hidden' }}>
          
          {/* Top View Selector Tab Bar */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 16px', background: 'var(--panel)', borderBottom: '1px solid var(--line)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <button 
                type="button" 
                onClick={() => { setActiveView('course_master'); setPage(0); }}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '6px 14px',
                  fontSize: '12.5px',
                  borderRadius: '6px',
                  border: '1px solid',
                  borderColor: activeView === 'course_master' ? 'var(--accent)' : 'var(--line)',
                  background: activeView === 'course_master' ? 'var(--accent)' : 'var(--bg)',
                  color: activeView === 'course_master' ? 'white' : 'var(--ink)',
                  fontWeight: activeView === 'course_master' ? 700 : 500
                }}
              >
                <TableProperties size={15} /> Course Master (37 Columns)
              </button>

              <button 
                type="button" 
                onClick={() => { setActiveView('group_master'); setPage(0); }}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '6px 14px',
                  fontSize: '12.5px',
                  borderRadius: '6px',
                  border: '1px solid',
                  borderColor: activeView === 'group_master' ? 'var(--accent)' : 'var(--line)',
                  background: activeView === 'group_master' ? 'var(--accent)' : 'var(--bg)',
                  color: activeView === 'group_master' ? 'white' : 'var(--ink)',
                  fontWeight: activeView === 'group_master' ? 700 : 500
                }}
              >
                <GitFork size={15} /> Group Master Hierarchy (10 Columns)
              </button>
            </div>

            <div style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
              Active View: <strong>{activeView === 'group_master' ? 'Group Master File' : 'Course Master Sheet'}</strong>
            </div>
          </div>

          {/* Table Header Controls Bar */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', padding: '10px 16px', borderBottom: '1px solid var(--line)', background: 'var(--bg)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                
                {/* Global Search Box */}
                <div style={{ position: 'relative', width: '160px' }}>
                  <input 
                    type="text" 
                    placeholder={activeView === 'group_master' ? "Global search..." : "Global search..."} 
                    value={searchQuery} 
                    onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }} 
                    style={{ width: '100%', padding: '5px 8px 5px 26px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)' }} 
                  />
                  <Search size={13} color="var(--muted)" style={{ position: 'absolute', left: '8px', top: '7px' }} />
                </div>

                {/* Sort Selector Dropdown (All 37 Headers) */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                  <select 
                    value={sortConfig.column || ''} 
                    onChange={(e) => {
                      const col = e.target.value;
                      if (!col) {
                        setSortConfig({ column: null, direction: null });
                      } else {
                        setSortConfig({ column: col, direction: 'asc' });
                      }
                      setPage(0);
                    }}
                    style={{ 
                      padding: '5px 8px', 
                      fontSize: '12px', 
                      borderRadius: '4px', 
                      border: sortConfig.column ? '1.5px solid var(--accent)' : '1px solid var(--line)', 
                      maxWidth: '220px', 
                      background: sortConfig.column ? 'var(--accent-soft)' : 'var(--bg)',
                      fontWeight: sortConfig.column ? 600 : 400
                    }}
                  >
                    <option value="">↕️ Sort by Header (Cols 1-37)...</option>
                    {(activeView === 'group_master' ? GROUP_MASTER_HEADERS : OUTPUT_HEADERS).map((col, idx) => (
                      <option key={col} value={col}>
                        Col {idx + 1}: {col} {sortConfig.column === col ? `(${sortConfig.direction.toUpperCase()})` : ''}
                      </option>
                    ))}
                  </select>

                  {sortConfig.column && (
                    <button 
                      type="button" 
                      className="secondary"
                      onClick={() => {
                        setSortConfig(prev => ({
                          ...prev,
                          direction: prev.direction === 'asc' ? 'desc' : 'asc'
                        }));
                        setPage(0);
                      }}
                      style={{ padding: '4px 8px', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px', color: 'var(--accent)', fontWeight: 600 }}
                      title="Toggle Ascending / Descending"
                    >
                      {sortConfig.direction === 'asc' ? <ArrowUp size={12} /> : <ArrowDown size={12} />}
                      {sortConfig.direction === 'asc' ? 'ASC' : 'DESC'}
                    </button>
                  )}
                </div>

                {/* Subject Filter Dropdown */}
                {uniqueSubjects.length > 0 && (
                  <select 
                    value={selectedSubjectFilter} 
                    onChange={(e) => { setSelectedSubjectFilter(e.target.value); setPage(0); }}
                    style={{ padding: '5px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', maxWidth: '140px' }}
                  >
                    <option value="ALL">All Subjects ({uniqueSubjects.length})</option>
                    {uniqueSubjects.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                )}

                {/* Group Filter Dropdown */}
                {activeView === 'course_master' && detectedGroups.length > 0 && (
                  <select 
                    value={selectedGroupFilter} 
                    onChange={(e) => { setSelectedGroupFilter(e.target.value); setPage(0); }}
                    style={{ padding: '5px 8px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', maxWidth: '130px' }}
                  >
                    <option value="ALL">All Groups ({detectedGroups.length})</option>
                    {detectedGroups.map(g => (
                      <option key={g} value={g}>{g}</option>
                    ))}
                  </select>
                )}

                {/* Reset Filters & Sort */}
                {(selectedSubjectFilter !== 'ALL' || selectedGroupFilter !== 'ALL' || searchQuery || Object.keys(columnFilters).length > 0 || sortConfig.column) && (
                  <button 
                    type="button" 
                    className="secondary" 
                    onClick={clearAllColumnFilters}
                    style={{ padding: '3px 8px', fontSize: '11px', color: 'var(--danger)' }}
                  >
                    Reset
                  </button>
                )}

                {/* Toggle Inline Column Filters Button */}
                <button
                  type="button"
                  className="secondary"
                  onClick={() => setShowColumnFilterRow(prev => !prev)}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '4px',
                    padding: '4px 8px',
                    fontSize: '11.5px',
                    borderRadius: '4px',
                    background: showColumnFilterRow ? 'var(--accent-soft)' : 'var(--bg)',
                    color: showColumnFilterRow ? 'var(--accent)' : 'var(--ink)',
                    border: showColumnFilterRow ? '1.5px solid var(--accent)' : '1px solid var(--line)'
                  }}
                  title={showColumnFilterRow ? "Hide 37-column search row" : "Show 37-column search row"}
                >
                  <ListFilter size={13} /> Filters {Object.keys(columnFilters).length > 0 ? `(${Object.keys(columnFilters).length})` : ''}
                </button>

                {/* Direct Download Filtered Subset Button */}
                {previewRows.length > 0 && (
                  <button 
                    type="button" 
                    onClick={exportFilteredSubset}
                    style={{ 
                      display: 'flex', 
                      alignItems: 'center', 
                      gap: '5px', 
                      padding: '4px 10px', 
                      fontSize: '11.5px', 
                      background: 'var(--accent)', 
                      color: 'white', 
                      borderRadius: '4px',
                      border: 'none'
                    }}
                    title="Download only the currently filtered/sorted rows"
                  >
                    <Download size={13} /> Export ({previewRows.length})
                  </button>
                )}

                {/* Quick Mode Toggle on Header Bar */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '4px', background: 'var(--panel)', padding: '2px 6px', borderRadius: '14px', border: '1px solid var(--line)' }}>
                  <button 
                    type="button" 
                    onClick={() => setUseDuplication(true)}
                    style={{
                      padding: '2px 8px',
                      fontSize: '11px',
                      borderRadius: '12px',
                      border: 'none',
                      background: useDuplication ? 'var(--accent)' : 'transparent',
                      color: useDuplication ? 'white' : 'var(--muted)',
                      fontWeight: 600
                    }}
                  >
                    ✨ With Duplication
                  </button>
                  <button 
                    type="button" 
                    onClick={() => setUseDuplication(false)}
                    style={{
                      padding: '2px 8px',
                      fontSize: '11px',
                      borderRadius: '12px',
                      border: 'none',
                      background: !useDuplication ? 'var(--accent)' : 'transparent',
                      color: !useDuplication ? 'white' : 'var(--muted)',
                      fontWeight: 600
                    }}
                  >
                    📄 No Duplication
                  </button>
                </div>

              </div>

              {/* Pagination Controls */}
              {previewRows.length > pageSize && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', color: 'var(--muted)' }}>
                  <span>Page {page + 1} of {Math.ceil(previewRows.length / pageSize)}</span>
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
                    disabled={(page + 1) * pageSize >= previewRows.length} 
                    onClick={() => setPage(p => p + 1)} 
                    style={{ padding: '2px 6px', fontSize: '11px' }}
                  >
                    Next
                  </button>
                </div>
              )}
            </div>

            {/* Active Sort / Filter Badges Bar */}
            {(sortConfig.column || Object.keys(columnFilters).length > 0) && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap', paddingTop: '2px' }}>
                <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 600 }}>Active:</span>
                
                {/* Active Sort Badge */}
                {sortConfig.column && (
                  <span 
                    style={{ 
                      display: 'inline-flex', 
                      alignItems: 'center', 
                      gap: '4px', 
                      fontSize: '10.5px', 
                      background: 'var(--accent-soft)', 
                      color: 'var(--accent)', 
                      border: '1.5px solid var(--accent)', 
                      padding: '2px 8px', 
                      borderRadius: '4px',
                      fontWeight: 700 
                    }}
                  >
                    Sorted by: Col {((activeView === 'group_master' ? GROUP_MASTER_HEADERS : OUTPUT_HEADERS).indexOf(sortConfig.column) + 1)} {sortConfig.column} ({sortConfig.direction === 'asc' ? '▲ Ascending' : '▼ Descending'})
                    <button 
                      type="button" 
                      onClick={() => setSortConfig({ column: null, direction: null })} 
                      style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--accent)', display: 'flex' }}
                      title="Clear sort"
                    >
                      <X size={11} />
                    </button>
                  </span>
                )}

                {/* Active Filter Badges */}
                {Object.entries(columnFilters).map(([colKey, filterVal]) => {
                  const headers = activeView === 'group_master' ? GROUP_MASTER_HEADERS : OUTPUT_HEADERS;
                  const colIdx = headers.indexOf(colKey) + 1;
                  return (
                    <span 
                      key={colKey} 
                      style={{ 
                        display: 'inline-flex', 
                        alignItems: 'center', 
                        gap: '4px', 
                        fontSize: '10.5px', 
                        background: 'var(--accent-soft)', 
                        color: 'var(--accent)', 
                        border: '1px solid var(--accent)', 
                        padding: '2px 6px', 
                        borderRadius: '4px',
                        fontWeight: 600 
                      }}
                    >
                      {colIdx > 0 ? `Col ${colIdx} ` : ''}{colKey}: "{filterVal}"
                      <button 
                        type="button" 
                        onClick={() => updateColumnFilter(colKey, '')} 
                        style={{ background: 'none', border: 'none', padding: 0, cursor: 'pointer', color: 'var(--accent)', display: 'flex' }}
                      >
                        <X size={11} />
                      </button>
                    </span>
                  );
                })}
              </div>
            )}
          </div>

          {/* Master Table Container */}
          <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
            {rawRows.length === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'var(--muted)', gap: '10px' }}>
                <FileSpreadsheet size={44} style={{ opacity: 0.3 }} />
                <strong style={{ fontSize: '14px', color: 'var(--ink)' }}>No File Uploaded</strong>
                <span style={{ fontSize: '12px', maxWidth: '440px', textAlign: 'center', lineHeight: 1.4 }}>
                  Upload your raw syllabus / course master spreadsheet on the left, or download our ready-to-fill standard template.
                </span>
                <button
                  type="button"
                  onClick={handleDownloadTemplate}
                  className="secondary"
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px',
                    padding: '8px 16px',
                    fontSize: '12px',
                    fontWeight: 600,
                    marginTop: '4px',
                    cursor: 'pointer'
                  }}
                >
                  <Download size={14} /> Download Course Master Input Template (.xlsx)
                </button>
              </div>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px', whiteSpace: 'nowrap' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', position: 'sticky', top: 0, zIndex: 11 }}>
                    <th style={{ padding: '8px 10px', textAlign: 'center', borderBottom: '1.5px solid var(--line)', color: 'var(--muted)', width: '40px' }}>#</th>
                    {(activeView === 'group_master' ? GROUP_MASTER_HEADERS : OUTPUT_HEADERS).map((col, idx) => {
                      const isSorted = sortConfig.column === col;
                      return (
                        <th 
                          key={col} 
                          onClick={() => handleSort(col)}
                          title={`Click to sort by Col ${idx + 1} (${col}) - Asc / Desc / Default`}
                          style={{ 
                            padding: '8px 10px', 
                            textAlign: 'left', 
                            borderBottom: '1.5px solid var(--line)', 
                            color: isSorted ? 'var(--accent)' : 'var(--ink)', 
                            fontWeight: 700,
                            borderRight: '1px solid var(--line)',
                            background: isSorted ? 'var(--accent-soft)' : 'var(--bg)',
                            cursor: 'pointer',
                            userSelect: 'none',
                            transition: 'background 0.15s ease'
                          }}
                        >
                          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '8px' }}>
                            <div>
                              <span style={{ fontSize: '9px', color: isSorted ? 'var(--accent)' : 'var(--muted)', display: 'block', fontWeight: 600 }}>
                                Col {idx + 1}
                              </span>
                              <span>{col}</span>
                            </div>
                            <div style={{ display: 'flex', alignItems: 'center' }}>
                              {isSorted && sortConfig.direction === 'asc' && (
                                <ArrowUp size={13} color="var(--accent)" />
                              )}
                              {isSorted && sortConfig.direction === 'desc' && (
                                <ArrowDown size={13} color="var(--accent)" />
                              )}
                              {!isSorted && (
                                <ArrowUpDown size={11} style={{ opacity: 0.25, color: 'var(--ink)' }} />
                              )}
                            </div>
                          </div>
                        </th>
                      );
                    })}
                  </tr>

                  {/* Inline Column Filter Inputs Header Row */}
                  {showColumnFilterRow && (
                    <tr style={{ background: 'var(--panel)', position: 'sticky', top: '35px', zIndex: 10 }}>
                      <th style={{ padding: '4px 6px', textAlign: 'center', borderBottom: '1.5px solid var(--line)', borderRight: '1px solid var(--line)', background: 'var(--panel)' }}>
                        <ListFilter size={11} color="var(--muted)" style={{ margin: '0 auto' }} />
                      </th>
                      {(activeView === 'group_master' ? GROUP_MASTER_HEADERS : OUTPUT_HEADERS).map((col) => {
                        const isFiltered = !!columnFilters[col];
                        return (
                          <th 
                            key={`filter_${col}`} 
                            style={{ 
                              padding: '3px 6px', 
                              borderBottom: '1.5px solid var(--line)', 
                              borderRight: '1px solid var(--line)', 
                              background: isFiltered ? 'var(--accent-soft)' : 'var(--panel)'
                            }}
                          >
                            <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
                              <input 
                                type="text" 
                                placeholder={`Filter...`} 
                                value={columnFilters[col] || ''} 
                                onChange={(e) => updateColumnFilter(col, e.target.value)} 
                                style={{ 
                                  width: '100%', 
                                  padding: isFiltered ? '2px 18px 2px 5px' : '2px 5px', 
                                  fontSize: '10.5px', 
                                  borderRadius: '3px', 
                                  border: isFiltered ? '1.5px solid var(--accent)' : '1px solid var(--line)',
                                  background: 'var(--bg)',
                                  color: 'var(--ink)'
                                }} 
                              />
                              {isFiltered && (
                                <button 
                                  type="button" 
                                  onClick={() => updateColumnFilter(col, '')} 
                                  style={{ 
                                    position: 'absolute', 
                                    right: '4px', 
                                    background: 'none', 
                                    border: 'none', 
                                    padding: 0, 
                                    cursor: 'pointer', 
                                    color: 'var(--muted)',
                                    display: 'flex',
                                    alignItems: 'center'
                                  }}
                                  title="Clear filter"
                                >
                                  <X size={10} />
                                </button>
                              )}
                            </div>
                          </th>
                        );
                      })}
                    </tr>
                  )}
                </thead>
                <tbody>
                  {pagedRows.map((row, rowIdx) => {
                    const actualIdx = (page * pageSize) + rowIdx;
                    const headers = activeView === 'group_master' ? GROUP_MASTER_HEADERS : OUTPUT_HEADERS;
                    return (
                      <tr key={actualIdx} style={{ borderBottom: '1px solid var(--line)' }}>
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: 'var(--muted)', borderRight: '1px solid var(--line)', fontWeight: 600 }}>
                          {actualIdx + 1}
                        </td>
                        {headers.map(col => (
                          <td 
                            key={col} 
                            style={{ 
                              padding: '5px 8px', 
                              borderRight: '1px solid var(--line)',
                              color: (col === 'ImmidiateParentGroup' || col === 'SubGroupName') ? 'var(--accent)' : 'var(--ink)',
                              fontWeight: (col === 'ImmidiateParentGroup' || col === 'SubGroupName' || col === 'ParentGroupName') ? 700 : 400
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
          {rawRows.length > 0 && (
            <div style={{ padding: '8px 16px', background: 'var(--bg)', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
              <span>Showing {pagedRows.length} of {previewRows.length} rows • View: <strong>{activeView === 'group_master' ? 'Group Master (10 cols)' : 'Course Master (37 cols)'}</strong></span>
              <span>Dynamic group hierarchy and custom subgroups synchronized</span>
            </div>
          )}

        </main>
      </div>
    </div>
  );
}
