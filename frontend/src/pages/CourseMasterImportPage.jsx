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
  Check
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

const normalizeKey = (key) => String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');

const DEFAULT_GROUP_CONFIG = {
  DSC: {
    pattern: 'group_subject', // 'group_subject' ({Group} - {Subject}) | 'group_only' ({Group}) | 'custom'
    copies: 2,
    suffixStr: ', .', // Comma-separated suffixes for each repetition (e.g. '', '.')
    maxCredits: 4,
    maxMarks: 100,
    maxCourses: 1,
    minCourses: 1
  },
  MDC: {
    pattern: 'group_only',
    copies: 1,
    suffixStr: '',
    maxCredits: 3,
    maxMarks: 75,
    maxCourses: 1,
    minCourses: 1
  },
  VAC: {
    pattern: 'group_only',
    copies: 1,
    suffixStr: '',
    maxCredits: 3,
    maxMarks: 75,
    maxCourses: 1,
    minCourses: 1
  },
  DEFAULT: {
    pattern: 'group_only',
    copies: 1,
    suffixStr: '',
    maxCredits: 0,
    maxMarks: 0,
    maxCourses: 1,
    minCourses: 1
  }
};

export default function CourseMasterImportPage() {
  const [sourceFile, setSourceFile] = useState(null);
  const [rawRows, setRawRows] = useState([]);
  const [headerMap, setHeaderMap] = useState({});
  const [groupConfigs, setGroupConfigs] = useState(DEFAULT_GROUP_CONFIG);
  const [exportMode, setExportMode] = useState('zip'); // 'zip' | 'combined' | 'multitab'
  const [selectedSubjectFilter, setSelectedSubjectFilter] = useState('ALL');
  const [searchQuery, setSearchQuery] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('info');
  const [page, setPage] = useState(0);
  const pageSize = 50;

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

  // Extract all unique group names and subjects present in uploaded data
  const { detectedGroups, uniqueSubjects } = useMemo(() => {
    if (!rawRows.length) return { detectedGroups: [], uniqueSubjects: [] };
    const groups = new Set();
    const subjects = new Set();

    rawRows.forEach(row => {
      const g = getCell(row, 'Group Name', 'groupname', 'group', 'parentgroup');
      const s = getCell(row, 'Subject', 'subject');
      if (g) groups.add(g.trim().toUpperCase());
      if (s) subjects.add(s.trim());
    });

    return {
      detectedGroups: Array.from(groups).sort(),
      uniqueSubjects: Array.from(subjects).sort()
    };
  }, [rawRows, headerMap]);

  // Synchronize groupConfigs with any newly detected groups
  useEffect(() => {
    if (!detectedGroups.length) return;
    setGroupConfigs(prev => {
      const next = { ...prev };
      detectedGroups.forEach(g => {
        if (!next[g]) {
          if (g.startsWith('DSC')) {
            next[g] = { ...DEFAULT_GROUP_CONFIG.DSC };
          } else if (g === 'MDC' || g === 'VAC') {
            next[g] = { ...DEFAULT_GROUP_CONFIG[g] };
          } else {
            next[g] = { ...DEFAULT_GROUP_CONFIG.DEFAULT };
          }
        }
      });
      return next;
    });
  }, [detectedGroups]);

  // Apply Presets
  const applyPreset = (presetKey) => {
    if (presetKey === 'fyugp_dot') {
      // Preset 1: DSC - Subj & DSC - Subj., MDC, VAC
      const next = { ...groupConfigs };
      detectedGroups.forEach(g => {
        if (g.startsWith('DSC')) {
          next[g] = { pattern: 'group_subject', copies: 2, suffixStr: ', .', maxCredits: 4, maxMarks: 100, maxCourses: 1, minCourses: 1 };
        } else if (g === 'MDC' || g === 'VAC') {
          next[g] = { pattern: 'group_only', copies: 1, suffixStr: '', maxCredits: 3, maxMarks: 75, maxCourses: 1, minCourses: 1 };
        } else {
          next[g] = { pattern: 'group_only', copies: 1, suffixStr: '', maxCredits: 0, maxMarks: 0, maxCourses: 1, minCourses: 1 };
        }
      });
      setGroupConfigs(next);
      setStatus('Applied Preset: FYUGP Standard (DSC - Subject & DSC - Subject.)', 'success');
    } else if (presetKey === 'numbered_series') {
      // Preset 2: DSC 1/2, MDC 1/2, VAC 1/2
      const next = { ...groupConfigs };
      detectedGroups.forEach(g => {
        if (g.startsWith('DSC')) {
          next[g] = { pattern: 'group_only', copies: 2, suffixStr: ' 1,  2', maxCredits: 4, maxMarks: 100, maxCourses: 1, minCourses: 1 };
        } else if (g === 'MDC' || g === 'VAC') {
          next[g] = { pattern: 'group_only', copies: 2, suffixStr: ' 1,  2', maxCredits: 3, maxMarks: 75, maxCourses: 1, minCourses: 1 };
        } else {
          next[g] = { pattern: 'group_only', copies: 2, suffixStr: ' 1,  2', maxCredits: 0, maxMarks: 0, maxCourses: 1, minCourses: 1 };
        }
      });
      setGroupConfigs(next);
      setStatus('Applied Preset: Numbered Series (e.g. DSC 1/2, VAC 1/2, MDC 1/2)', 'success');
    } else if (presetKey === 'major_minor') {
      // Preset 3: DSC - Subj M1, DSC - Subj M2
      const next = { ...groupConfigs };
      detectedGroups.forEach(g => {
        if (g.startsWith('DSC')) {
          next[g] = { pattern: 'group_subject', copies: 2, suffixStr: ' M1,  M2', maxCredits: 4, maxMarks: 100, maxCourses: 1, minCourses: 1 };
        } else if (g === 'MDC' || g === 'VAC') {
          next[g] = { pattern: 'group_only', copies: 1, suffixStr: '', maxCredits: 3, maxMarks: 75, maxCourses: 1, minCourses: 1 };
        }
      });
      setGroupConfigs(next);
      setStatus('Applied Preset: Major / Minor Multipliers (DSC - Subj M1, DSC - Subj M2)', 'success');
    } else if (presetKey === 'clean_single') {
      // Preset 4: Clean Single (No Duplication)
      const next = { ...groupConfigs };
      detectedGroups.forEach(g => {
        next[g] = {
          pattern: g.startsWith('DSC') ? 'group_subject' : 'group_only',
          copies: 1,
          suffixStr: '',
          maxCredits: g.startsWith('DSC') ? 4 : (g === 'MDC' || g === 'VAC') ? 3 : 0,
          maxMarks: g.startsWith('DSC') ? 100 : (g === 'MDC' || g === 'VAC') ? 75 : 0,
          maxCourses: 1,
          minCourses: 1
        };
      });
      setGroupConfigs(next);
      setStatus('Applied Preset: Clean Single (No Duplication)', 'success');
    }
  };

  const updateGroupConfig = (group, field, value) => {
    setGroupConfigs(prev => ({
      ...prev,
      [group]: {
        ...(prev[group] || DEFAULT_GROUP_CONFIG.DEFAULT),
        [field]: value
      }
    }));
  };

  // Helper to parse comma-separated suffixes into array matching copy count
  const getSuffixList = (cfg, count) => {
    const rawParts = (cfg.suffixStr || '').split(',').map(s => s.trim());
    const list = [];
    for (let i = 0; i < count; i++) {
      if (rawParts[i] !== undefined && rawParts[i] !== '') {
        list.push(rawParts[i]);
      } else if (i === 0) {
        list.push('');
      } else {
        // Fallback sequence if fewer suffixes provided than copy count
        list.push(` ${i + 1}`);
      }
    }
    return list;
  };

  // Master transformation pipeline
  const generateProcessedRows = (forSubject = null) => {
    if (!rawRows.length) return [];
    const rowsForSubject = [];

    const subjectsToProcess = forSubject ? [forSubject] : (uniqueSubjects.length ? uniqueSubjects : [null]);

    subjectsToProcess.forEach(fileSubject => {
      rawRows.forEach(row => {
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

        // Fetch user custom configuration for this group
        const cfg = groupConfigs[groupKey] || (groupKey.startsWith('DSC') ? DEFAULT_GROUP_CONFIG.DSC : DEFAULT_GROUP_CONFIG.DEFAULT);
        const copies = Math.max(1, Number(cfg.copies) || 1);
        const suffixList = getSuffixList(cfg, copies);

        combinations.forEach(([amMethod, atType]) => {
          const totalCredits = getNumber(row, 'Total Credits', 'totalcredits', 'credits', 'credit');
          const totalMarks = getNumber(row, 'Total Marks', 'totalmarks', 'marks');
          const minMarks = getNumber(row, 'Minimum Passing Marks', 'minimumpassingmarks', 'minpassingmarks', 'minmarks');

          // Generate each repetition copy
          for (let copyIdx = 0; copyIdx < copies; copyIdx++) {
            const newRow = {};
            newRow['UniqueProgramTermCode'] = '';

            // Construct ImmidiateParentGroup with custom pattern and suffix
            let baseGroup = rawGroup || 'General';
            if (cfg.pattern === 'group_subject') {
              baseGroup = `${rawGroup.trim()} - ${subject.trim()}`;
            }

            const currentSuffix = suffixList[copyIdx] || '';
            const immidiateParentGroup = `${baseGroup}${currentSuffix}`;
            newRow['ImmidiateParentGroup'] = immidiateParentGroup;

            // Admission Limits
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
    });

    // Custom sorting priority
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

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setSourceFile(file);
    setIsProcessing(true);
    setStatus('Reading course master spreadsheet...', 'info');

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
      setStatus(`Loaded ${parsed.length} course rows across ${workbook.SheetNames.length} sheet(s)!`, 'success');
    } catch (err) {
      console.error('Course Master Ingestion Error:', err);
      setStatus(`Upload failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const executeExport = async () => {
    if (!rawRows.length) return alert('Please upload a course master sheet first.');
    setIsProcessing(true);
    setStatus('Generating custom course master workbooks...', 'info');

    try {
      if (exportMode === 'zip') {
        const zip = new JSZip();
        for (const subj of uniqueSubjects) {
          const rows = generateProcessedRows(subj);
          const aoa = [OUTPUT_HEADERS, ...rows.map(r => OUTPUT_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))];
          const wb = XLSX.utils.book_new();
          const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
          ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: OUTPUT_HEADERS.length - 1 } }) };
          ws['!cols'] = OUTPUT_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }));
          XLSX.utils.book_append_sheet(wb, ws, subj.slice(0, 30));
          const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
          zip.file(`${subj} Courses.xlsx`, out);
        }

        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const url = URL.createObjectURL(zipBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Course_Master_Custom_Subjects.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported ${uniqueSubjects.length} subject course workbooks into ZIP!`, 'success');

      } else if (exportMode === 'multitab') {
        const wb = XLSX.utils.book_new();
        for (const subj of uniqueSubjects) {
          const rows = generateProcessedRows(subj);
          const aoa = [OUTPUT_HEADERS, ...rows.map(r => OUTPUT_HEADERS.map(h => r[h] !== undefined ? r[h] : ''))];
          const ws = XLSX.utils.aoa_to_sheet(aoa, { dense: true });
          ws['!cols'] = OUTPUT_HEADERS.map(h => ({ wch: Math.max(h.length + 2, 12) }));
          XLSX.utils.book_append_sheet(wb, ws, subj.slice(0, 31).replace(/[:\\/?*[\]]/g, '_'));
        }
        const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Course_Master_Custom_All_Tabs.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported multi-tab workbook with ${uniqueSubjects.length} subject tabs!`, 'success');

      } else {
        const allRows = generateProcessedRows(null);
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
        a.download = `Course_Master_Custom_Combined.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported combined master XLSX with ${allRows.length} rows!`, 'success');
      }
    } catch (err) {
      console.error('Export Error:', err);
      setStatus(`Export failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const previewRows = useMemo(() => {
    const targetSubj = selectedSubjectFilter === 'ALL' ? null : selectedSubjectFilter;
    const generated = generateProcessedRows(targetSubj);
    if (!searchQuery) return generated;
    const q = searchQuery.toLowerCase();
    return generated.filter(r => 
      String(r.CourseCode || '').toLowerCase().includes(q) ||
      String(r.CourseName || '').toLowerCase().includes(q) ||
      String(r.Subject || '').toLowerCase().includes(q) ||
      String(r.ImmidiateParentGroup || '').toLowerCase().includes(q)
    );
  }, [rawRows, headerMap, groupConfigs, selectedSubjectFilter, searchQuery]);

  const pagedRows = useMemo(() => {
    const start = page * pageSize;
    return previewRows.slice(start, start + pageSize);
  }, [previewRows, page]);

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
            <BookOpen size={18} color="var(--accent)" /> Course Master Import & Dynamic Group Rules
            <span style={{ fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '12px', fontWeight: 700 }}>Customizable</span>
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
            {exportMode === 'zip' ? 'Download Subject ZIP' : exportMode === 'multitab' ? 'Export Multi-Tab XLSX' : 'Export Combined XLSX'}
          </button>
        </div>
      </header>

      {/* Main Workspace Layout */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden', padding: '16px 24px', gap: '16px' }}>
        
        {/* Left Settings Sidebar */}
        <aside style={{ width: '360px', display: 'flex', flexDirection: 'column', gap: '14px', overflowY: 'auto' }}>
          
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
              <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Supports raw syllabus matrix (.xlsx / .csv)</span>
            </label>
          </div>

          {/* Quick Presets Bar */}
          <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <div style={{ fontSize: '11.5px', textTransform: 'uppercase', color: 'var(--muted)', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '6px' }}>
              <Sparkles size={14} color="var(--accent)" /> Quick Configuration Presets
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '6px' }}>
              <button type="button" className="secondary" onClick={() => applyPreset('fyugp_dot')} style={{ fontSize: '11px', padding: '5px 8px', textAlign: 'left' }}>
                ⚡ DSC - Subj & .
              </button>
              <button type="button" className="secondary" onClick={() => applyPreset('numbered_series')} style={{ fontSize: '11px', padding: '5px 8px', textAlign: 'left' }}>
                ⚡ Numbers (1, 2)
              </button>
              <button type="button" className="secondary" onClick={() => applyPreset('major_minor')} style={{ fontSize: '11px', padding: '5px 8px', textAlign: 'left' }}>
                ⚡ Multiplier (M1, M2)
              </button>
              <button type="button" className="secondary" onClick={() => applyPreset('clean_single')} style={{ fontSize: '11px', padding: '5px 8px', textAlign: 'left' }}>
                ⚡ Single (No Dup)
              </button>
            </div>
          </div>

          {/* Group Duplication & Suffix Rules Builder */}
          <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '12px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3 style={{ margin: 0, fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700 }}>
                <Sliders size={15} color="var(--accent)" /> Group Duplication & Suffixes
              </h3>
              <span style={{ fontSize: '11px', color: 'var(--muted)' }}>{detectedGroups.length || 3} Group(s)</span>
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              {(detectedGroups.length ? detectedGroups : ['DSC', 'MDC', 'VAC']).map(groupKey => {
                const cfg = groupConfigs[groupKey] || (groupKey.startsWith('DSC') ? DEFAULT_GROUP_CONFIG.DSC : DEFAULT_GROUP_CONFIG.DEFAULT);
                return (
                  <div 
                    key={groupKey} 
                    style={{ 
                      padding: '10px', 
                      background: 'var(--bg)', 
                      border: '1px solid var(--line)', 
                      borderRadius: '6px',
                      display: 'flex',
                      flexDirection: 'column',
                      gap: '8px'
                    }}
                  >
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <strong style={{ fontSize: '12.5px', color: 'var(--accent)' }}>Group: {groupKey}</strong>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Copies:</span>
                        <input 
                          type="number" 
                          min={1} 
                          max={10} 
                          value={cfg.copies || 1} 
                          onChange={(e) => updateGroupConfig(groupKey, 'copies', Math.max(1, parseInt(e.target.value, 10) || 1))}
                          style={{ width: '48px', padding: '2px 4px', fontSize: '11px', textAlign: 'center' }} 
                        />
                      </div>
                    </div>

                    <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1.8fr', gap: '6px' }}>
                      <div className="form-group" style={{ margin: 0 }}>
                        <label style={{ fontSize: '10.5px' }}>Base Pattern</label>
                        <select 
                          value={cfg.pattern || 'group_only'} 
                          onChange={(e) => updateGroupConfig(groupKey, 'pattern', e.target.value)}
                          style={{ fontSize: '11px', padding: '3px 6px' }}
                        >
                          <option value="group_subject">{groupKey} - Subject</option>
                          <option value="group_only">{groupKey} Only</option>
                        </select>
                      </div>

                      <div className="form-group" style={{ margin: 0 }}>
                        <label style={{ fontSize: '10.5px' }}>Suffixes (comma-separated)</label>
                        <input 
                          type="text" 
                          placeholder="e.g. , . or 1, 2" 
                          value={cfg.suffixStr !== undefined ? cfg.suffixStr : ''} 
                          onChange={(e) => updateGroupConfig(groupKey, 'suffixStr', e.target.value)}
                          style={{ fontSize: '11px', padding: '3px 6px' }} 
                        />
                      </div>
                    </div>

                    {/* Preview Generated Suffixes */}
                    <div style={{ fontSize: '10.5px', color: 'var(--muted)', background: 'var(--panel)', padding: '4px 6px', borderRadius: '4px' }}>
                      Preview: {getSuffixList(cfg, cfg.copies).map((sfx, idx) => (
                        <span key={idx} style={{ fontWeight: 600, color: 'var(--ink)' }}>
                          {idx > 0 && ' | '}
                          {cfg.pattern === 'group_subject' ? `${groupKey} - Subj${sfx}` : `${groupKey}${sfx}`}
                        </span>
                      ))}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>

          {/* Export Output Package Format */}
          <div className="card" style={{ padding: '14px', margin: 0, display: 'flex', flexDirection: 'column', gap: '8px' }}>
            <h3 style={{ margin: 0, fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700 }}>
              <Settings2 size={15} color="var(--accent)" /> Output Package Format
            </h3>
            
            <select value={exportMode} onChange={(e) => setExportMode(e.target.value)} style={{ fontSize: '12px' }}>
              <option value="zip">📦 ZIP Archive (Separate .xlsx per Subject)</option>
              <option value="combined">📑 Combined Master Sheet (.xlsx)</option>
              <option value="multitab">📚 Multi-Tab Workbook (.xlsx)</option>
            </select>
          </div>

        </aside>

        {/* Center Live Validation & Preview Grid */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '10px', overflow: 'hidden' }}>
          
          {/* Table Header Bar */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 16px', borderBottom: '1px solid var(--line)', background: 'var(--bg)', flexWrap: 'wrap', gap: '10px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              
              {/* Search Box */}
              <div style={{ position: 'relative', width: '220px' }}>
                <input 
                  type="text" 
                  placeholder="Filter courses, codes..." 
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
                  style={{ padding: '5px 10px', fontSize: '12px', borderRadius: '4px', border: '1px solid var(--line)', maxWidth: '200px' }}
                >
                  <option value="ALL">All Subjects ({uniqueSubjects.length})</option>
                  {uniqueSubjects.map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
              )}

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

          {/* Master 37-Column Table Container */}
          <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
            {rawRows.length === 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', color: 'var(--muted)', gap: '8px' }}>
                <FileSpreadsheet size={40} style={{ opacity: 0.3 }} />
                <strong style={{ fontSize: '14px', color: 'var(--ink)' }}>No Course Master File Uploaded</strong>
                <span style={{ fontSize: '12px' }}>Upload your raw syllabus / course master spreadsheet on the left to transform.</span>
              </div>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px', whiteSpace: 'nowrap' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', position: 'sticky', top: 0, zIndex: 10 }}>
                    <th style={{ padding: '8px 10px', textAlign: 'center', borderBottom: '1.5px solid var(--line)', color: 'var(--muted)', width: '40px' }}>#</th>
                    {OUTPUT_HEADERS.map((col, idx) => (
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
                    return (
                      <tr key={actualIdx} style={{ borderBottom: '1px solid var(--line)' }}>
                        <td style={{ padding: '6px 8px', textAlign: 'center', color: 'var(--muted)', borderRight: '1px solid var(--line)', fontWeight: 600 }}>
                          {actualIdx + 1}
                        </td>
                        {OUTPUT_HEADERS.map(col => (
                          <td 
                            key={col} 
                            style={{ 
                              padding: '5px 8px', 
                              borderRight: '1px solid var(--line)',
                              color: col === 'ImmidiateParentGroup' ? 'var(--accent)' : 'var(--ink)',
                              fontWeight: col === 'ImmidiateParentGroup' ? 700 : 400
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
              <span>Showing {pagedRows.length} of {previewRows.length} rows across {uniqueSubjects.length} subjects</span>
              <span>All 37 master columns formatted with live custom group rules</span>
            </div>
          )}

        </main>
      </div>
    </div>
  );
}
