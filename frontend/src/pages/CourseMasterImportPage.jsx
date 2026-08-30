import React, { useState, useMemo } from 'react';
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

const normalizeKey = (key) => String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');

export default function CourseMasterImportPage() {
  const [sourceFile, setSourceFile] = useState(null);
  const [rawRows, setRawRows] = useState([]);
  const [headerMap, setHeaderMap] = useState({});
  const [useDuplication, setUseDuplication] = useState(true); // Toggle duplication logic
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

  const uniqueSubjects = useMemo(() => {
    if (!rawRows.length) return [];
    const subjects = new Set();
    rawRows.forEach(row => {
      const subj = getCell(row, 'Subject', 'subject');
      if (subj) subjects.add(subj);
    });
    return Array.from(subjects).sort();
  }, [rawRows, headerMap]);

  // Process rows for a specific subject or globally
  const generateProcessedRows = (forSubject = null) => {
    if (!rawRows.length) return [];
    const rowsForSubject = [];

    const subjectsToProcess = forSubject ? [forSubject] : (uniqueSubjects.length ? uniqueSubjects : [null]);

    subjectsToProcess.forEach(fileSubject => {
      rawRows.forEach(row => {
        const subject = getCell(row, 'Subject', 'subject');
        const groupName = getCell(row, 'Group Name', 'groupname', 'group', 'parentgroup');

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

        combinations.forEach(([amMethod, atType]) => {
          const newRow = {};

          newRow['UniqueProgramTermCode'] = '';

          // Logic for 'ImmidiateParentGroup' based on Group Name
          let immidiateParentGroup;
          const isDsc = groupName.trim().toUpperCase().startsWith('DSC');
          if (isDsc) {
            immidiateParentGroup = `${groupName.trim()} - ${subject.trim()}`;
          } else {
            immidiateParentGroup = groupName || 'General';
          }
          newRow['ImmidiateParentGroup'] = immidiateParentGroup;

          const totalCredits = getNumber(row, 'Total Credits', 'totalcredits', 'credits', 'credit');
          const totalMarks = getNumber(row, 'Total Marks', 'totalmarks', 'marks');
          const minMarks = getNumber(row, 'Minimum Passing Marks', 'minimumpassingmarks', 'minpassingmarks', 'minmarks');

          // Logic for GroupMax/MinCoursesforAdmission
          if (immidiateParentGroup === 'DSC 1' || immidiateParentGroup.startsWith('DSC - ')) {
            newRow['GroupMaxCoursesforAdmission'] = 1;
            newRow['GroupMinCoursesforAdmission'] = 1;
            newRow['GroupMaxCreditsforAdmission'] = 4;
            newRow['GroupMaxMarks'] = 100;
          } else if (immidiateParentGroup === 'MDC' || immidiateParentGroup === 'VAC') {
            newRow['GroupMaxCoursesforAdmission'] = 1;
            newRow['GroupMinCoursesforAdmission'] = 1;
            newRow['GroupMaxCreditsforAdmission'] = 3;
            newRow['GroupMaxMarks'] = 75;
          } else {
            newRow['GroupMaxCoursesforAdmission'] = 1;
            newRow['GroupMinCoursesforAdmission'] = 1;
            newRow['GroupMaxCreditsforAdmission'] = 0;
            newRow['GroupMaxMarks'] = totalCredits;
          }

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

          // Conditional assessment data
          newRow['AssessmentMethod'] = amMethod;
          newRow['AssessmentType'] = atType;
          newRow['AMEvaluationSystem'] = 'Marks System';
          newRow['AMCredits'] = 0;
          newRow['AMEvaluationTemplate'] = 'Eight Level';
          newRow['ATEvaluationSystem'] = 'Marks System';
          newRow['ATEvaluationTemplate'] = 'Eight Level';

          // ATCredits
          if (atType === 'TH') {
            newRow['ATCredits'] = getNumber(row, 'Theory Credits', 'theorycredits', 'thcredits');
          } else if (atType === 'PR') {
            newRow['ATCredits'] = getNumber(row, 'Practical Credits', 'practicalcredits', 'prcredits');
          } else {
            newRow['ATCredits'] = 0;
          }

          // ATMaxMarks
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

          // Duplication Logic (if enabled)
          if (useDuplication && isDsc) {
            const duplicatedRow = {
              ...newRow,
              ImmidiateParentGroup: `${immidiateParentGroup}.`,
              GroupMaxCoursesforAdmission: 1,
              GroupMinCoursesforAdmission: 1,
              GroupMaxCreditsforAdmission: 4,
              GroupMaxCredits: 4,
              GroupMaxMarks: 100
            };
            rowsForSubject.push(duplicatedRow);
          }
        });
      });
    });

    // Sort according to script priority
    rowsForSubject.sort((a, b) => {
      const order = (group) => {
        if (group === 'DSC 1') return 1;
        if (group === 'DSC 2') return 2;
        if (group.startsWith('DSC - ') && !group.endsWith('.')) return 3;
        if (group.startsWith('DSC - ') && group.endsWith('.')) return 4;
        if (group === 'MDC') return 5;
        if (group === 'VAC') return 6;
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
      setStatus(`Loaded ${parsed.length} raw course rows across ${workbook.SheetNames.length} sheet(s)!`, 'success');
    } catch (err) {
      console.error('Course Master Ingestion Error:', err);
      setStatus(`Upload failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Export handler supporting ZIP (per subject), Combined XLSX, or Multi-Tab XLSX
  const executeExport = async () => {
    if (!rawRows.length) return alert('Please upload a course master sheet first.');
    setIsProcessing(true);
    setStatus('Generating course master workbooks...', 'info');

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
        a.download = `Course_Master_Subjects_${useDuplication ? 'With_Duplication' : 'Standard'}.zip`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported ${uniqueSubjects.length} subject course workbooks into ZIP archive!`, 'success');

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
        a.download = `Course_Master_All_Tabs_${useDuplication ? 'With_Duplication' : 'Standard'}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported multi-tab Excel workbook with ${uniqueSubjects.length} subject sheets!`, 'success');

      } else {
        // Combined single sheet
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
        a.download = `Course_Master_Combined_${useDuplication ? 'With_Duplication' : 'Standard'}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        setStatus(`Exported combined 37-column Master XLSX with ${allRows.length} rows!`, 'success');
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
  }, [rawRows, headerMap, useDuplication, selectedSubjectFilter, searchQuery]);

  const pagedRows = useMemo(() => {
    const start = page * pageSize;
    return previewRows.slice(start, start + pageSize);
  }, [previewRows, page]);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', overflow: 'hidden', background: 'var(--bg)', color: 'var(--ink)' }}>
      {/* Header */}
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 24px', background: 'var(--panel)', borderBottom: '1px solid var(--line)', zIndex: 20 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}>
            <ArrowLeft size={16} /> Back to Portal
          </Link>
          <div style={{ height: '18px', width: '1px', background: 'var(--line)' }} />
          <h2 style={{ fontSize: '16px', margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
            <BookOpen size={18} color="var(--accent)" /> Course Master Import
            <span style={{ fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '12px', fontWeight: 700 }}>37 Master Columns</span>
          </h2>
        </div>

        {/* Status & Export Button */}
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

      {/* Main Workspace */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden', padding: '16px 24px', gap: '16px' }}>
        
        {/* Left Settings Sidebar */}
        <aside style={{ width: '320px', display: 'flex', flexDirection: 'column', gap: '14px', overflowY: 'auto' }}>
          
          {/* Upload Dropzone */}
          <div className="card" style={{ padding: '20px', margin: 0, textAlign: 'center', border: '1.5px dashed var(--accent)', background: 'var(--accent-soft)' }}>
            <input 
              type="file" 
              id="courseMasterFileInput" 
              accept=".xlsx,.xls,.csv" 
              onChange={handleFileUpload} 
              style={{ display: 'none' }} 
            />
            <label htmlFor="courseMasterFileInput" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '6px' }}>
              <Upload size={28} color="var(--accent)" />
              <strong style={{ fontSize: '13.5px', color: 'var(--ink)' }}>Upload Course Master Sheet</strong>
              <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Supports raw master course data</span>
            </label>
          </div>

          {/* User Prompt: Duplication Logic Toggle */}
          <div className="card" style={{ padding: '16px', margin: 0, display: 'flex', flexDirection: 'column', gap: '10px' }}>
            <h3 style={{ margin: 0, fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700 }}>
              <Copy size={15} color="var(--accent)" /> Duplication Mode
            </h3>
            <p style={{ margin: 0, fontSize: '12px', color: 'var(--muted)' }}>
              Choose whether to apply automated DSC 1 $\rightarrow$ DSC 2 and group duplication logic:
            </p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', marginTop: '4px' }}>
              <label style={{ display: 'flex', alignItems: 'flex-start', gap: '8px', cursor: 'pointer', fontSize: '12px', padding: '8px', borderRadius: '6px', border: '1px solid', borderColor: useDuplication ? 'var(--accent)' : 'var(--line)', background: useDuplication ? 'var(--accent-soft)' : 'transparent' }}>
                <input 
                  type="radio" 
                  name="duplicationToggle" 
                  checked={useDuplication} 
                  onChange={() => setUseDuplication(true)} 
                  style={{ marginTop: '2px' }}
                />
                <div>
                  <strong style={{ display: 'block', color: 'var(--ink)' }}>With Duplication (Recommended)</strong>
                  <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Duplicates DSC 1 $\rightarrow$ DSC 2 and appends alias groups for multi-choice allotment.</span>
                </div>
              </label>

              <label style={{ display: 'flex', alignItems: 'flex-start', gap: '8px', cursor: 'pointer', fontSize: '12px', padding: '8px', borderRadius: '6px', border: '1px solid', borderColor: !useDuplication ? 'var(--accent)' : 'var(--line)', background: !useDuplication ? 'var(--accent-soft)' : 'transparent' }}>
                <input 
                  type="radio" 
                  name="duplicationToggle" 
                  checked={!useDuplication} 
                  onChange={() => setUseDuplication(false)} 
                  style={{ marginTop: '2px' }}
                />
                <div>
                  <strong style={{ display: 'block', color: 'var(--ink)' }}>Without Duplication (Clean Master)</strong>
                  <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Outputs single master rows without creating DSC 2 or alias clones.</span>
                </div>
              </label>
            </div>
          </div>

          {/* Export Output Format */}
          <div className="card" style={{ padding: '16px', margin: 0, display: 'flex', flexDirection: 'column', gap: '10px' }}>
            <h3 style={{ margin: 0, fontSize: '13.5px', display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700 }}>
              <Settings2 size={15} color="var(--accent)" /> Output Package Format
            </h3>
            
            <div className="form-group" style={{ margin: 0 }}>
              <select value={exportMode} onChange={(e) => setExportMode(e.target.value)} style={{ fontSize: '12px' }}>
                <option value="zip">📦 ZIP Archive (Separate .xlsx per Subject)</option>
                <option value="combined">📑 Combined Master Sheet (.xlsx)</option>
                <option value="multitab">📚 Multi-Tab Workbook (.xlsx)</option>
              </select>
            </div>
          </div>

          {/* Summary Stats */}
          {rawRows.length > 0 && (
            <div className="card" style={{ padding: '16px', margin: 0, display: 'flex', flexDirection: 'column', gap: '8px' }}>
              <div style={{ fontSize: '11.5px', textTransform: 'uppercase', color: 'var(--muted)', fontWeight: 700 }}>Master Statistics</div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                <div style={{ padding: '8px', background: 'var(--bg)', borderRadius: '6px', border: '1px solid var(--line)', textAlign: 'center' }}>
                  <div style={{ fontSize: '18px', fontWeight: 800, color: 'var(--accent)' }}>{uniqueSubjects.length}</div>
                  <div style={{ fontSize: '10.5px', color: 'var(--muted)' }}>Subjects</div>
                </div>
                <div style={{ padding: '8px', background: 'var(--bg)', borderRadius: '6px', border: '1px solid var(--line)', textAlign: 'center' }}>
                  <div style={{ fontSize: '18px', fontWeight: 800, color: '#16a34a' }}>{previewRows.length}</div>
                  <div style={{ fontSize: '10.5px', color: 'var(--muted)' }}>Total Rows</div>
                </div>
              </div>
            </div>
          )}

        </aside>

        {/* Center Live Validation Table */}
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

          {/* Table Body */}
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
                              color: 'var(--ink)'
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
              <span>Showing {pagedRows.length} of {previewRows.length} rows • Duplication: <strong>{useDuplication ? 'ON (DSC 1 & 2 duplicates)' : 'OFF (Standard master)'}</strong></span>
              <span>All 37 standardized master columns formatted</span>
            </div>
          )}

        </main>
      </div>
    </div>
  );
}
