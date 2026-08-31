import React, { useState, useMemo, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { 
  ArrowLeft, 
  Upload, 
  Download, 
  FileSpreadsheet, 
  CheckCircle2, 
  AlertTriangle, 
  Search, 
  BookOpen, 
  Layers, 
  Copy, 
  FolderArchive, 
  RefreshCw, 
  HelpCircle, 
  Sparkles, 
  Settings2, 
  Sliders, 
  Plus, 
  Trash2, 
  Check, 
  FileCheck,
  ToggleLeft,
  ToggleRight,
  GitFork,
  TableProperties,
  Edit3,
  X,
  ListFilter,
  FolderPlus,
  Trash
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

const createDefaultConfigForGroup = (groupKey) => {
  const upper = String(groupKey || '').toUpperCase().trim();
  if (upper.includes('DSC') || upper.includes('MAJOR') || upper.includes('CORE') || upper.includes('DSE')) {
    return {
      pattern: 'group_subject',
      copies: 2,
      suffixStr: ', .',
      maxCredits: 4,
      maxMarks: 100,
      maxCourses: 1,
      minCourses: 1
    };
  } else {
    return {
      pattern: 'group_only',
      copies: 1,
      suffixStr: '',
      maxCredits: 3,
      maxMarks: 75,
      maxCourses: 1,
      minCourses: 1
    };
  }
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
  const { detectedGroups, uniqueSubjects, groupToSubjectsMap } = useMemo(() => {
    if (!rawRows.length) return { detectedGroups: [], uniqueSubjects: [], groupToSubjectsMap: {} };
    const groups = new Set();
    const subjects = new Set();
    const map = {};

    rawRows.forEach(row => {
      const g = getCell(row, 'Group Name', 'groupname', 'group', 'parentgroup').trim().toUpperCase();
      const s = getCell(row, 'Subject', 'subject').trim();
      if (g) {
        groups.add(g);
        if (!map[g]) map[g] = new Set();
        if (s) map[g].add(s);
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
      groupToSubjectsMap: finalMap
    };
  }, [rawRows, headerMap]);

  // Synchronize dynamic group configurations whenever a new file is uploaded
  useEffect(() => {
    if (!detectedGroups.length) return;
    setGroupConfigs(prev => {
      const next = { ...prev };
      detectedGroups.forEach(g => {
        if (!next[g]) {
          next[g] = createDefaultConfigForGroup(g);
        }
      });
      return next;
    });
  }, [detectedGroups]);

  // Helper to build default standard hierarchy automatically
  const autoPopulateHierarchy = (targetSubj = null) => {
    if (!detectedGroups.length) return;
    const activeSubj = targetSubj || (selectedSubjectFilter !== 'ALL' ? selectedSubjectFilter : uniqueSubjects[0]);
    const nextHierarchy = [];

    detectedGroups.forEach(gKey => {
      const cfg = groupConfigs[gKey] || createDefaultConfigForGroup(gKey);
      const subjsForThisGroup = groupToSubjectsMap[gKey] || uniqueSubjects;
      const copies = useDuplication ? Math.max(1, Number(cfg.copies) || 1) : 1;
      const suffixes = useDuplication ? (cfg.suffixStr || '').split(',').map(s => s.trim()) : [''];

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
    setStatus(`Auto-populated ${nextHierarchy.length} group(s) with buckets!`, 'success');
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

  const toggleSubjectInSubgroup = (groupId, subGroupId, subj) => {
    setGroupHierarchy(prev => prev.map(g => {
      if (g.id !== groupId) return g;
      return {
        ...g,
        subGroups: g.subGroups.map(sg => {
          if (sg.id !== subGroupId) return sg;
          const exists = (sg.subjects || []).includes(subj);
          return {
            ...sg,
            subjects: exists ? sg.subjects.filter(s => s !== subj) : [...(sg.subjects || []), subj]
          };
        })
      };
    }));
  };

  const selectAllForSubgroup = (groupId, subGroupId, groupName) => {
    const allSubjs = groupToSubjectsMap[groupName] || uniqueSubjects;
    setGroupHierarchy(prev => prev.map(g => {
      if (g.id !== groupId) return g;
      return {
        ...g,
        subGroups: g.subGroups.map(sg => {
          if (sg.id !== subGroupId) return sg;
          return { ...sg, subjects: [...allSubjs] };
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

      const eseMaxTh = getNumber(row, 'ESE Max - TH', 'esemaxth', 'eseth', 'esemax_th');
      const ccaMaxTh = getNumber(row, 'CCA Max - TH', 'ccamaxth', 'ccath', 'ccamax_th', 'cemaxth');
      const eseMaxPr = getNumber(row, 'ESE Max - PR', 'esemaxpr', 'esepr', 'esemax_pr');
      const ccaMaxPr = getNumber(row, 'CCA Max - PR', 'ccamaxpr', 'ccapr', 'ccamax_pr', 'cemaxpr');

      const hasThAssessments = eseMaxTh > 0 || ccaMaxTh > 0;
      const hasPrAssessments = eseMaxPr > 0 || ccaMaxPr > 0;

      let combinations = [];
      if (hasThAssessments) {
        combinations.push(['ESE', 'TH'], ['CE', 'TH']);
      }
      if (hasPrAssessments) {
        combinations.push(['ESE', 'PR'], ['CE', 'PR']);
      }
      if (combinations.length === 0) {
        combinations.push(['', '']);
      }

      const cfg = groupConfigs[groupKey] || createDefaultConfigForGroup(groupKey);
      const copies = useDuplication ? Math.max(1, Number(cfg.copies) || 1) : 1;
      const rawSuffixes = (cfg.suffixStr || '').split(',').map(s => s.trim());
      const suffixList = [];
      for (let i = 0; i < copies; i++) {
        suffixList.push(rawSuffixes[i] !== undefined ? rawSuffixes[i] : (i === 0 ? '' : ` ${i + 1}`));
      }

      combinations.forEach(([amMethod, atType]) => {
        const totalCredits = getNumber(row, 'Total Credits', 'totalcredits', 'credits', 'credit');
        const totalMarks = getNumber(row, 'Total Marks', 'totalmarks', 'marks');
        const minMarks = getNumber(row, 'Minimum Passing Marks', 'minimumpassingmarks', 'minpassingmarks', 'minmarks');

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
          newRow['GroupMaxCreditsforAdmission'] = cfg.maxCredits > 0 ? Number(cfg.maxCredits) : 0;
          newRow['GroupMaxMarks'] = cfg.maxMarks > 0 ? Number(cfg.maxMarks) : totalCredits;
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
          newRow['AMCredits'] = 0;
          newRow['AMEvaluationTemplate'] = 'Eight Level';
          newRow['ATEvaluationSystem'] = 'Marks System';
          newRow['ATEvaluationTemplate'] = 'Eight Level';

          if (atType === 'TH') {
            newRow['ATCredits'] = getNumber(row, 'Theory Credits', 'theorycredits', 'thcredits');
          } else if (atType === 'PR') {
            newRow['ATCredits'] = getNumber(row, 'Practical Credits', 'practicalcredits', 'prcredits');
          } else {
            newRow['ATCredits'] = 0;
          }

          if (amMethod === 'ESE' && atType === 'TH') {
            newRow['ATMaxMarks'] = eseMaxTh;
          } else if (amMethod === 'ESE' && atType === 'PR') {
            newRow['ATMaxMarks'] = eseMaxPr;
          } else if (amMethod === 'CE' && atType === 'TH') {
            newRow['ATMaxMarks'] = ccaMaxTh;
          } else if (amMethod === 'CE' && atType === 'PR') {
            newRow['ATMaxMarks'] = ccaMaxPr;
          } else {
            newRow['ATMaxMarks'] = 0;
          }

          newRow['ATMinMarks'] = 0;
          newRow['AMMaxMarks'] = newRow['ATMaxMarks'];
          newRow['AMMinMarks'] = newRow['ATMinMarks'];

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

  // 2. GROUP MASTER HIERARCHY TRANSFORMATION PIPELINE (10 COLUMNS)
  // Completely driven by user-defined groupHierarchy state or fallback to dynamic generation
  const generateGroupMasterRows = (forSubject = null) => {
    if (!rawRows.length) return [];
    const result = [];

    const addRow = (parent, sub, maxM = 100, minM = 0, maxC = 4, minC = 0, maxSub = 1, minSub = 1) => {
      result.push({
        UniqueProgramTermCode: '',
        ParentGroupName: parent,
        SubGroupName: sub,
        ParentGroupMinSubGroups: minSub,
        ParentGroupMaxSubGroups: maxSub,
        ParentGroupEvaluationSystem: 'Marks System',
        ParentGroupMaxMarks: maxM,
        ParentGroupMinMarks: minM,
        ParentGroupMaxCredits: maxC,
        ParentGroupMinCredits: minC
      });
    };

    if (groupHierarchy.length > 0) {
      // Use user-curated groupHierarchy
      groupHierarchy.forEach(g => {
        if (!g.subGroups || g.subGroups.length === 0) {
          addRow(g.groupName, g.groupName);
        } else {
          g.subGroups.forEach(sg => {
            addRow(g.groupName, sg.name);
            if (sg.pattern === 'group_subject') {
              let subjectsToRender = sg.subjects || [];
              if (forSubject) {
                // If exporting for a specific subject
                if (sg.name.endsWith('- 1')) {
                  subjectsToRender = [forSubject];
                } else {
                  subjectsToRender = (groupToSubjectsMap[g.groupName] || uniqueSubjects).filter(s => s.toLowerCase() !== forSubject.toLowerCase());
                }
              }
              subjectsToRender.forEach(subj => {
                addRow(sg.name, `${g.groupName} - ${subj}${sg.suffix || ''}`);
              });
            } else {
              addRow(sg.name, `${sg.name}${sg.suffix || ''}`);
            }
          });
        }
      });
    } else {
      // Fallback: build standard default structure dynamically
      detectedGroups.forEach(gKey => {
        const cfg = groupConfigs[gKey] || createDefaultConfigForGroup(gKey);
        const copies = useDuplication ? Math.max(1, Number(cfg.copies) || 1) : 1;
        const rawSuffixes = (cfg.suffixStr || '').split(',').map(s => s.trim());
        const subjsForThisGroup = groupToSubjectsMap[gKey] || uniqueSubjects;
        const activeSubj = forSubject || (selectedSubjectFilter !== 'ALL' ? selectedSubjectFilter : subjsForThisGroup[0]);

        if (copies === 1 && cfg.pattern === 'group_only') {
          addRow(gKey, gKey, cfg.maxMarks, 0, cfg.maxCredits, 0);
        } else {
          for (let copyIdx = 0; copyIdx < copies; copyIdx++) {
            const bucketName = `${gKey} - ${copyIdx + 1}`;
            const sfx = rawSuffixes[copyIdx] || '';
            addRow(gKey, bucketName, cfg.maxMarks, 0, cfg.maxCredits, 0);

            if (cfg.pattern === 'group_subject') {
              if (copyIdx === 0) {
                if (activeSubj) addRow(bucketName, `${gKey} - ${activeSubj}${sfx}`, cfg.maxMarks, 0, cfg.maxCredits, 0);
              } else {
                subjsForThisGroup
                  .filter(s => s.toLowerCase() !== (activeSubj || '').toLowerCase())
                  .forEach(s => {
                    addRow(bucketName, `${gKey} - ${s}${sfx}`, cfg.maxMarks, 0, cfg.maxCredits, 0);
                  });
              }
            } else {
              addRow(bucketName, `${gKey}${sfx}`, cfg.maxMarks, 0, cfg.maxCredits, 0);
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
      const parsed = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

      if (!parsed.length) throw new Error('No data found in uploaded sheet.');

      const headers = Object.keys(parsed[0]);
      const map = {};
      headers.forEach(h => {
        map[normalizeKey(h)] = h;
      });

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
    if (activeView === 'group_master') {
      const targetSubj = selectedSubjectFilter === 'ALL' ? null : selectedSubjectFilter;
      const generated = generateGroupMasterRows(targetSubj);
      if (!searchQuery) return generated;
      const q = searchQuery.toLowerCase();
      return generated.filter(r => 
        String(r.ParentGroupName || '').toLowerCase().includes(q) ||
        String(r.SubGroupName || '').toLowerCase().includes(q)
      );
    } else {
      const targetSubj = selectedSubjectFilter === 'ALL' ? null : selectedSubjectFilter;
      const targetGroup = selectedGroupFilter === 'ALL' ? null : selectedGroupFilter;
      const generated = generateProcessedRows(targetSubj, targetGroup);
      if (!searchQuery) return generated;
      const q = searchQuery.toLowerCase();
      return generated.filter(r => 
        String(r.CourseCode || '').toLowerCase().includes(q) ||
        String(r.CourseName || '').toLowerCase().includes(q) ||
        String(r.Subject || '').toLowerCase().includes(q) ||
        String(r.ImmidiateParentGroup || '').toLowerCase().includes(q)
      );
    }
  }, [rawRows, headerMap, groupConfigs, useDuplication, activeView, selectedSubjectFilter, selectedGroupFilter, searchQuery, groupHierarchy]);

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
                  <ListFilter size={18} color="var(--accent)" /> Assign Subjects to: <span style={{ color: 'var(--accent)' }}>{editingSubgroup.subGroupName}</span>
                </h3>
                <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                  Group: <strong>{editingSubgroup.groupName}</strong> • Pick subjects to include in this bucket
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
                  placeholder="Search subjects..." 
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

            {/* Modal Subject Checkboxes Grid */}
            <div style={{ flex: 1, overflowY: 'auto', padding: '16px 20px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
              {(groupToSubjectsMap[editingSubgroup.groupName] || uniqueSubjects)
                .filter(s => !subjectModalSearch || s.toLowerCase().includes(subjectModalSearch.toLowerCase()))
                .map(subj => {
                  const currGroup = groupHierarchy.find(g => g.id === editingSubgroup.groupId);
                  const currSubgroup = currGroup?.subGroups?.find(sg => sg.id === editingSubgroup.subGroupId);
                  const isAssigned = (currSubgroup?.subjects || []).includes(subj);
                  return (
                    <label 
                      key={subj} 
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
                        onChange={() => toggleSubjectInSubgroup(editingSubgroup.groupId, editingSubgroup.subGroupId, subj)} 
                      />
                      <span style={{ flex: 1, color: 'var(--ink)', fontWeight: isAssigned ? 600 : 400 }}>
                        {subj}
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
              <strong style={{ fontSize: '13px', color: 'var(--ink)' }}>Upload Course Master Sheet</strong>
              <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Extracts Group Names and lets you manage groups & buckets</span>
            </label>
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

          {/* FULLY CUSTOMIZABLE GROUP & SUBGROUP MANAGER */}
          <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '12px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3 style={{ margin: 0, fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700 }}>
                <GitFork size={15} color="var(--accent)" /> Group & Sub-Group Manager
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
                <p style={{ margin: '0 0 8px 0' }}>No custom groups added yet (Started blank as requested).</p>
                <div style={{ display: 'flex', justifyContent: 'center', gap: '6px' }}>
                  <button type="button" className="secondary" onClick={() => autoPopulateHierarchy()} style={{ fontSize: '11px', padding: '4px 10px' }}>
                    ⚡ Quick Auto-Fill Standard Groups
                  </button>
                  <button type="button" onClick={() => setShowNewGroupModal(true)} style={{ fontSize: '11px', padding: '4px 10px' }}>
                    <Plus size={12} style={{ display: 'inline', marginRight: '4px' }} /> Add Custom Group
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
                                {sg.pattern === 'group_subject' ? (
                                  <span style={{ fontSize: '10px', color: 'var(--muted)', marginLeft: '4px' }}>
                                    ({subjCount} subjects {sg.suffix ? `• Suffix: "${sg.suffix}"` : ''})
                                  </span>
                                ) : (
                                  <span style={{ fontSize: '10px', color: 'var(--muted)', marginLeft: '4px' }}>
                                    (Group Only)
                                  </span>
                                )}
                              </div>

                              <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                                {sg.pattern === 'group_subject' && (
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
                                    <Edit3 size={10} /> Subjects ({subjCount})
                                  </button>
                                )}
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
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 16px', borderBottom: '1px solid var(--line)', background: 'var(--bg)', flexWrap: 'wrap', gap: '10px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
              
              {/* Search Box */}
              <div style={{ position: 'relative', width: '180px' }}>
                <input 
                  type="text" 
                  placeholder={activeView === 'group_master' ? "Filter parents, subgroups..." : "Filter courses, codes..."} 
                  value={searchQuery} 
                  onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }} 
                  style={{ width: '100%', padding: '5px 8px 5px 26px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)' }} 
                />
                <Search size={13} color="var(--muted)" style={{ position: 'absolute', left: '8px', top: '7px' }} />
              </div>

              {/* Subject Filter Dropdown */}
              {uniqueSubjects.length > 0 && (
                <select 
                  value={selectedSubjectFilter} 
                  onChange={(e) => { setSelectedSubjectFilter(e.target.value); setPage(0); }}
                  style={{ padding: '5px 10px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', maxWidth: '170px' }}
                >
                  <option value="ALL">All Subjects ({uniqueSubjects.length})</option>
                  {uniqueSubjects.map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
              )}

              {/* Group Filter Dropdown (Course Master View) */}
              {activeView === 'course_master' && detectedGroups.length > 0 && (
                <select 
                  value={selectedGroupFilter} 
                  onChange={(e) => { setSelectedGroupFilter(e.target.value); setPage(0); }}
                  style={{ padding: '5px 10px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', maxWidth: '150px' }}
                >
                  <option value="ALL">All Groups ({detectedGroups.length})</option>
                  {detectedGroups.map(g => (
                    <option key={g} value={g}>{g}</option>
                  ))}
                </select>
              )}

              {/* Reset Filters Link */}
              {(selectedSubjectFilter !== 'ALL' || selectedGroupFilter !== 'ALL' || searchQuery) && (
                <button 
                  type="button" 
                  className="secondary" 
                  onClick={() => {
                    setSelectedSubjectFilter('ALL');
                    setSelectedGroupFilter('ALL');
                    setSearchQuery('');
                    setPage(0);
                  }}
                  style={{ padding: '3px 8px', fontSize: '11px', color: 'var(--danger)' }}
                >
                  Reset
                </button>
              )}

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
                    borderRadius: '4px' 
                  }}
                  title="Download only the currently filtered rows"
                >
                  <Download size={13} /> Download {activeView === 'group_master' ? 'Group Master' : 'Filtered'} ({previewRows.length})
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

          {/* Master Table Container */}
          <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
            {rawRows.length === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'var(--muted)', gap: '8px' }}>
                <FileSpreadsheet size={40} style={{ opacity: 0.3 }} />
                <strong style={{ fontSize: '14px', color: 'var(--ink)' }}>No File Uploaded</strong>
                <span style={{ fontSize: '12px' }}>Upload your raw syllabus / course master spreadsheet on the left to extract data and transform both masters.</span>
              </div>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px', whiteSpace: 'nowrap' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', position: 'sticky', top: 0, zIndex: 10 }}>
                    <th style={{ padding: '8px 10px', textAlign: 'center', borderBottom: '1.5px solid var(--line)', color: 'var(--muted)', width: '40px' }}>#</th>
                    {(activeView === 'group_master' ? GROUP_MASTER_HEADERS : OUTPUT_HEADERS).map((col, idx) => (
                      <th 
                        key={col} 
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
                        {col}
                      </th>
                    ))}
                  </tr>
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
