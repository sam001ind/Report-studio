import React, { useState, useMemo, useRef, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Search,
  Plus,
  Trash2,
  Edit2,
  Check,
  RotateCcw,
  Upload,
  Download,
  AlertCircle,
  CheckCircle2,
  Tag,
  Users,
  Heart,
  Globe,
  MapPin,
  GraduationCap,
  BookOpen,
  Building2,
  ChevronLeft,
  ChevronRight,
  Database,
  FileSpreadsheet,
  HelpCircle,
  X,
  SlidersHorizontal,
  Sparkles,
  Layers,
  ArrowUpDown
} from 'lucide-react';
import {
  LOOKUP_TAB_CONFIG,
  DEFAULT_LOOKUP_DATA,
  loadSavedLookups,
  saveLookupsToStorage,
  parseUploadedLookupSpreadsheet
} from '../data/admissionLookups';

const TAB_ICONS = {
  program: BookOpen,
  college: Building2,
  category: Tag,
  caste: Users,
  religion: Heart,
  country: Globe,
  state: MapPin,
  qualification: GraduationCap,
  stream: FileSpreadsheet
};

export default function MasterLookupDefinitionPage() {
  const [lookups, setLookups] = useState(() => loadSavedLookups());
  const [activeTab, setActiveTab] = useState('program');
  const [searchQuery, setSearchQuery] = useState('');
  const [pageSize, setPageSize] = useState(20);
  const [page, setPage] = useState(0);
  const [statusAlert, setStatusAlert] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showGuideModal, setShowGuideModal] = useState(false);

  // Simple lookup add state
  const [newName, setNewName] = useState('');
  const [newId, setNewId] = useState('');

  // Programme add state
  const [newProgRef, setNewProgRef] = useState('');
  const [newProgCode, setNewProgCode] = useState('');
  const [newProgWorkflow, setNewProgWorkflow] = useState('');
  const [newProgName, setNewProgName] = useState('');

  // College add state
  const [newClgId, setNewClgId] = useState('');
  const [newClgCode, setNewClgCode] = useState('');
  const [newClgName, setNewClgName] = useState('');
  const [newClgRef, setNewClgRef] = useState('');

  // Inline edit state
  const [editingIdx, setEditingIdx] = useState(null);
  const [editName, setEditName] = useState('');
  const [editId, setEditId] = useState('');
  const [editProgRef, setEditProgRef] = useState('');
  const [editProgCode, setEditProgCode] = useState('');
  const [editProgWorkflow, setEditProgWorkflow] = useState('');
  const [editProgName, setEditProgName] = useState('');
  const [editClgId, setEditClgId] = useState('');
  const [editClgCode, setEditClgCode] = useState('');
  const [editClgName, setEditClgName] = useState('');
  const [editClgRef, setEditClgRef] = useState('');

  const fileInputRef = useRef(null);

  // Load from localStorage on mount
  useEffect(() => {
    const saved = loadSavedLookups();
    setLookups(saved);
  }, []);

  const currentTabConfig = useMemo(() => {
    return LOOKUP_TAB_CONFIG.find(t => t.key === activeTab) || LOOKUP_TAB_CONFIG[0];
  }, [activeTab]);

  const currentList = useMemo(() => {
    return (lookups && Array.isArray(lookups[activeTab])) ? lookups[activeTab] : [];
  }, [lookups, activeTab]);

  const filteredList = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return currentList;

    if (activeTab === 'program') {
      return currentList.filter(item => {
        const ref = String(item.pgmName || '').toLowerCase();
        const code = String(item.code || '').toLowerCase();
        const wf = String(item.workflowId || '').toLowerCase();
        const name = String(item.name || '').toLowerCase();
        return ref.includes(q) || code.includes(q) || wf.includes(q) || name.includes(q);
      });
    }

    if (activeTab === 'college') {
      return currentList.filter(item => {
        const id = String(item.id || '').toLowerCase();
        const code = String(item.code || '').toLowerCase();
        const name = String(item.name || '').toLowerCase();
        const ref = String(item.clgRef || '').toLowerCase();
        return id.includes(q) || code.includes(q) || name.includes(q) || ref.includes(q);
      });
    }

    return currentList.filter(item => {
      const n = String(item.name || '').toLowerCase();
      const id = String(item.id || '').toLowerCase();
      return n.includes(q) || id.includes(q);
    });
  }, [currentList, searchQuery, activeTab]);

  const paginatedList = useMemo(() => {
    if (pageSize === 0) return filteredList;
    const start = page * pageSize;
    return filteredList.slice(start, start + pageSize);
  }, [filteredList, page, pageSize]);

  const totalEntriesCount = useMemo(() => {
    if (!lookups) return 0;
    return Object.values(lookups).reduce((sum, list) => sum + (Array.isArray(list) ? list.length : 0), 0);
  }, [lookups]);

  const totalPages = pageSize === 0 ? 1 : Math.ceil(filteredList.length / pageSize) || 1;

  const updateLookups = (newLookups) => {
    setLookups(newLookups);
    saveLookupsToStorage(newLookups);
  };

  const handleTabChange = (newTab) => {
    setActiveTab(newTab);
    setSearchQuery('');
    setEditingIdx(null);
    setPage(0);
    setNewName('');
    setNewId('');
    setNewProgRef('');
    setNewProgCode('');
    setNewProgWorkflow('');
    setNewProgName('');
    setNewClgId('');
    setNewClgCode('');
    setNewClgName('');
    setNewClgRef('');
    setStatusAlert(null);
  };

  const handleAdd = (e) => {
    e?.preventDefault();

    if (activeTab === 'program') {
      const pgmName = newProgRef.trim();
      const code = newProgCode.trim();
      const workflowId = newProgWorkflow.trim();
      const name = newProgName.trim();

      if (!pgmName || !code || !workflowId || !name) {
        setStatusAlert({
          type: 'warning',
          message: 'Please fill all 4 fields for Programme (Reference, Code, workFlowId, Full Name).'
        });
        return;
      }

      const exists = currentList.some(item => (item.pgmName || '').toLowerCase() === pgmName.toLowerCase());
      if (exists) {
        setStatusAlert({
          type: 'warning',
          message: `Programme "${pgmName}" already exists.`
        });
        return;
      }

      const updatedTabList = [{ pgmName, code, workflowId, name }, ...currentList];
      const updatedAll = { ...lookups, program: updatedTabList };
      updateLookups(updatedAll);
      setNewProgRef('');
      setNewProgCode('');
      setNewProgWorkflow('');
      setNewProgName('');
      setStatusAlert({
        type: 'success',
        message: `Added Programme "${pgmName}" (${code} / ${workflowId}).`
      });
      return;
    }

    if (activeTab === 'college') {
      const id = newClgId.trim();
      const code = newClgCode.trim() || id;
      const name = newClgName.trim();
      const clgRef = newClgRef.trim() || name;

      if (!id || !name) {
        setStatusAlert({
          type: 'warning',
          message: 'Please fill at least CollegeId and College Name.'
        });
        return;
      }

      const exists = currentList.some(item => (item.id || '') === id || (item.name || '').toLowerCase() === name.toLowerCase());
      if (exists) {
        setStatusAlert({
          type: 'warning',
          message: `College with ID ${id} or Name "${name}" already exists.`
        });
        return;
      }

      const updatedTabList = [{ id, code, name, clgRef }, ...currentList];
      const updatedAll = { ...lookups, college: updatedTabList };
      updateLookups(updatedAll);
      setNewClgId('');
      setNewClgCode('');
      setNewClgName('');
      setNewClgRef('');
      setStatusAlert({
        type: 'success',
        message: `Added College "${name}" (ID: ${id}).`
      });
      return;
    }

    // Simple lookup add
    const trimmedName = newName.trim();
    const trimmedId = newId.trim();

    if (!trimmedName || !trimmedId) {
      setStatusAlert({
        type: 'warning',
        message: 'Both Name and ID are required to add an entry.'
      });
      return;
    }

    const exists = currentList.some(item => (item.name || '').toLowerCase() === trimmedName.toLowerCase());
    if (exists) {
      setStatusAlert({
        type: 'warning',
        message: `Entry "${trimmedName}" already exists in ${currentTabConfig.label}.`
      });
      return;
    }

    const updatedTabList = [{ name: trimmedName, id: trimmedId }, ...currentList];
    const updatedAll = { ...lookups, [activeTab]: updatedTabList };
    updateLookups(updatedAll);
    setNewName('');
    setNewId('');
    setStatusAlert({
      type: 'success',
      message: `Added "${trimmedName}" -> "${trimmedId}" to ${currentTabConfig.label}.`
    });
  };

  const handleStartEdit = (item, idx) => {
    const realIdx = filteredList.indexOf(item);
    setEditingIdx(realIdx);
    if (activeTab === 'program') {
      setEditProgRef(item.pgmName || '');
      setEditProgCode(item.code || '');
      setEditProgWorkflow(item.workflowId || '');
      setEditProgName(item.name || '');
    } else if (activeTab === 'college') {
      setEditClgId(item.id || '');
      setEditClgCode(item.code || '');
      setEditClgName(item.name || '');
      setEditClgRef(item.clgRef || '');
    } else {
      setEditName(item.name || '');
      setEditId(item.id || '');
    }
  };

  const handleSaveEdit = (item) => {
    const realIdx = currentList.indexOf(item);
    if (realIdx === -1) return;

    const updatedTabList = [...currentList];

    if (activeTab === 'program') {
      const pgmName = editProgRef.trim();
      const code = editProgCode.trim();
      const workflowId = editProgWorkflow.trim();
      const name = editProgName.trim();
      if (!pgmName || !code || !workflowId || !name) {
        setStatusAlert({ type: 'warning', message: 'All 4 fields are required.' });
        return;
      }
      updatedTabList[realIdx] = { pgmName, code, workflowId, name };
    } else if (activeTab === 'college') {
      const id = editClgId.trim();
      const code = editClgCode.trim() || id;
      const name = editClgName.trim();
      const clgRef = editClgRef.trim() || name;
      if (!id || !name) {
        setStatusAlert({ type: 'warning', message: 'CollegeId and College Name are required.' });
        return;
      }
      updatedTabList[realIdx] = { id, code, name, clgRef };
    } else {
      const name = editName.trim();
      const id = editId.trim();
      if (!name || !id) {
        setStatusAlert({ type: 'warning', message: 'Name and ID are required.' });
        return;
      }
      updatedTabList[realIdx] = { name, id };
    }

    const updatedAll = { ...lookups, [activeTab]: updatedTabList };
    updateLookups(updatedAll);
    setEditingIdx(null);
    setStatusAlert({
      type: 'success',
      message: `Updated entry in ${currentTabConfig.label}.`
    });
  };

  const handleDelete = (item) => {
    const realIdx = currentList.indexOf(item);
    if (realIdx === -1) return;

    const label = activeTab === 'program' ? item.pgmName : (activeTab === 'college' ? item.name : item.name);
    if (!window.confirm(`Delete "${label}" from ${currentTabConfig.label}?`)) return;

    const updatedTabList = currentList.filter((_, i) => i !== realIdx);
    const updatedAll = { ...lookups, [activeTab]: updatedTabList };
    updateLookups(updatedAll);
    setEditingIdx(null);
    setStatusAlert({
      type: 'success',
      message: `Deleted "${label}" from ${currentTabConfig.label}.`
    });
  };

  const handleResetToDefaults = () => {
    if (window.confirm('Reset all defined IDs to official Kannur University defaults? Custom entries will be replaced with official definitions.')) {
      updateLookups(DEFAULT_LOOKUP_DATA);
      setEditingIdx(null);
      setPage(0);
      setStatusAlert({
        type: 'success',
        message: 'Reset all 9 master ID definitions to official Kannur University defaults.'
      });
    }
  };

  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setStatusAlert({
      type: 'info',
      message: `Reading "${file.name}"...`
    });

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const { imported, summary, total } = parseUploadedLookupSpreadsheet(workbook, activeTab);

        if (total === 0) {
          setStatusAlert({
            type: 'warning',
            message: `No matching lookup columns found in "${file.name}". Expected headers like "Programme Name (Reference)", "ProgramCode", "CollegeId", "College Name", etc.`
          });
          setIsProcessing(false);
          return;
        }

        const merged = { ...lookups };
        let newEntriesCount = 0;

        // Merge Programmes
        if (imported.program && imported.program.length > 0) {
          const existingRefs = new Set((merged.program || []).map(x => (x.pgmName || '').toLowerCase()));
          const toAdd = imported.program.filter(x => !existingRefs.has(x.pgmName.toLowerCase()));
          merged.program = [...toAdd, ...(merged.program || [])];
          newEntriesCount += toAdd.length;
        }

        // Merge Colleges
        if (imported.college && imported.college.length > 0) {
          const existingColls = new Set((merged.college || []).map(x => (x.name + '_' + x.id).toLowerCase()));
          const toAdd = imported.college.filter(x => !existingColls.has((x.name + '_' + x.id).toLowerCase()));
          merged.college = [...toAdd, ...(merged.college || [])];
          newEntriesCount += toAdd.length;
        }

        // Merge standard 2-column lists
        for (const key of Object.keys(imported)) {
          if (key === 'program' || key === 'college') continue;
          if (Array.isArray(imported[key]) && imported[key].length > 0) {
            const existingNames = new Set((merged[key] || []).map(x => (x.name || '').toLowerCase()));
            const toAdd = imported[key].filter(x => !existingNames.has(x.name.toLowerCase()));
            merged[key] = [...toAdd, ...(merged[key] || [])];
            newEntriesCount += toAdd.length;
          }
        }

        updateLookups(merged);
        const summaryText = Object.entries(summary)
          .filter(([_, count]) => count > 0)
          .map(([k, count]) => `${count} ${k}`)
          .join(', ');

        setStatusAlert({
          type: 'success',
          message: `Successfully imported ${total} rows (${newEntriesCount} new entries added): ${summaryText}. Saved to active definitions.`
        });
      } catch (err) {
        console.error('Lookup upload error:', err);
        setStatusAlert({
          type: 'error',
          message: `Failed to parse file: ${err.message}`
        });
      } finally {
        setIsProcessing(false);
        if (fileInputRef.current) fileInputRef.current.value = '';
      }
    };

    reader.onerror = () => {
      setStatusAlert({
        type: 'error',
        message: 'Failed to read file from disk.'
      });
      setIsProcessing(false);
    };

    reader.readAsArrayBuffer(file);
  };

  // Download Side-by-Side Programme & College Template (matching the user's exact uploaded table)
  const handleDownloadSideBySideTemplate = () => {
    const headers = [
      'Programme Name (Reference)',
      'ProgramCode',
      'workFlowId',
      'Program Name',
      '',
      'CollegeId',
      'CollegeCode',
      'College Name',
      'CLG Name'
    ];

    const sampleRows = [
      [
        'Five Year Integrated Masters Programme in  Clinical Psychology',
        'PFPSYGS25',
        '151',
        'Five Year Integrated Masters Programme in Clinical Psychology(with Credits)-Regular-FYIMP-2025-Clinical Psychology-FYIMP Year I SEMESTER I',
        '',
        '435',
        '435',
        'Department of Anthropology',
        'Department of Anthropology,Dr. Janaki Ammal Campus,Palayad'
      ],
      [
        'Five Year Integrated Masters Programme in Commerce',
        'PFCOMGS25',
        '148',
        'Five Year Integrated Masters Programme in Commerce(with Credits)-Regular-FYIMP-2025-Commerce-FYIMP Year I SEMESTER I',
        '',
        '432',
        '432',
        'School of Behavioural Sciences, Mangattuparamba',
        'Department of Behavioural Sciences,Kannur University, Mangattuparamba Campus,K U Campus P.O,Pin : 670567'
      ],
      [
        'Five Year Integrated Masters Programme in Computational science',
        'PFCLSGS25',
        '149',
        'Five Year Integrated Masters Programme in Computational science(with Credits)-Regular-FYIMP-2025-Computational Sciences-FYIMP Year I SEMESTER I',
        '',
        '430',
        '430',
        'Department of Information Technology, Mangattuparamba',
        'Jointly conducted by Dept of IT,Mathematics & Statistics,Mangattuparamba Campus, Kannur University Campus P. O, Kannur - 670 567'
      ]
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [
      { wch: 45 },
      { wch: 15 },
      { wch: 12 },
      { wch: 70 },
      { wch: 4 },
      { wch: 12 },
      { wch: 14 },
      { wch: 38 },
      { wch: 65 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Programme_and_College_Master');
    XLSX.writeFile(wb, 'Programme_and_College_Master_Template.xlsx');

    setStatusAlert({
      type: 'success',
      message: 'Downloaded Programme_and_College_Master_Template.xlsx (Side-by-side format).'
    });
  };

  // Export full defined workbook
  const handleExportAllToExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      // Sheet 1: Side-by-side Programme & College
      const maxRows = Math.max((lookups.program || []).length, (lookups.college || []).length);
      const combinedRows = [
        [
          'Programme Name (Reference)',
          'ProgramCode',
          'workFlowId',
          'Program Name',
          '',
          'CollegeId',
          'CollegeCode',
          'College Name',
          'CLG Name'
        ]
      ];

      for (let i = 0; i < maxRows; i++) {
        const p = (lookups.program || [])[i] || {};
        const c = (lookups.college || [])[i] || {};
        combinedRows.push([
          p.pgmName || '',
          p.code || '',
          p.workflowId || '',
          p.name || '',
          '',
          c.id || '',
          c.code || '',
          c.name || '',
          c.clgRef || ''
        ]);
      }

      const wsCombined = XLSX.utils.aoa_to_sheet(combinedRows);
      wsCombined['!cols'] = [
        { wch: 45 }, { wch: 15 }, { wch: 12 }, { wch: 70 }, { wch: 4 },
        { wch: 12 }, { wch: 14 }, { wch: 38 }, { wch: 65 }
      ];
      XLSX.utils.book_append_sheet(wb, wsCombined, 'Master_Programme_College');

      // Individual sheets for each lookup type
      for (const config of LOOKUP_TAB_CONFIG) {
        const list = lookups[config.key] || [];
        let rows = [];

        if (config.key === 'program') {
          rows = [
            ['Programme Name (Reference)', 'ProgramCode', 'workFlowId', 'Program Name'],
            ...list.map(x => [x.pgmName, x.code, x.workflowId, x.name])
          ];
        } else if (config.key === 'college') {
          rows = [
            ['CollegeId', 'CollegeCode', 'College Name', 'CLG Name'],
            ...list.map(x => [x.id, x.code, x.name, x.clgRef])
          ];
        } else {
          rows = [
            [config.nameHeader, config.idHeader],
            ...list.map(x => [x.name, x.id])
          ];
        }

        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, config.label.slice(0, 31));
      }

      XLSX.writeFile(wb, 'Defined_Admission_Master_Lookups.xlsx');
      setStatusAlert({
        type: 'success',
        message: 'Downloaded Defined_Admission_Master_Lookups.xlsx with all master entity definitions.'
      });
    } catch (err) {
      console.error('Export error:', err);
      setStatusAlert({
        type: 'error',
        message: `Export failed: ${err.message}`
      });
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-100 flex flex-col font-sans transition-colors duration-150">
      {/* Top Header */}
      <header className="bg-white dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 sticky top-0 z-20 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <Link
              to="/"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700/60 text-xs font-medium transition-colors"
            >
              <ArrowLeft className="w-3.5 h-3.5" /> Back to Dashboard
            </Link>
            <div className="h-4 w-px bg-slate-300 dark:bg-slate-600 hidden sm:block" />
            <div>
              <div className="flex items-center gap-2">
                <span className="p-1 rounded-md bg-emerald-50 dark:bg-emerald-950/60 border border-emerald-200 dark:border-emerald-800 text-emerald-600 dark:text-emerald-400">
                  <Database className="w-4 h-4" />
                </span>
                <h1 className="text-base sm:text-lg font-bold text-slate-900 dark:text-white">
                  Programme &amp; College Master Definitions
                </h1>
                <span className="text-[10px] font-semibold tracking-wider uppercase px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 dark:bg-emerald-900/60 dark:text-emerald-300 border border-emerald-300 dark:border-emerald-700">
                  {totalEntriesCount} Total Defined
                </span>
              </div>
              <p className="text-xs text-slate-500 dark:text-slate-400 hidden md:block">
                Manage, map, and import official university reference codes for Programmes, Colleges, Castes, Categories, and Allotment IDs.
              </p>
            </div>
          </div>

          <div className="flex items-center flex-wrap gap-2">
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              className="hidden"
            />
            
            <button
              onClick={() => fileInputRef.current?.click()}
              disabled={isProcessing}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-xs transition-colors disabled:opacity-50"
              title="Upload spreadsheet with Programme, College or lookup ID definitions"
            >
              <Upload className="w-3.5 h-3.5" />
              Upload Spreadsheet
            </button>

            <button
              onClick={handleDownloadSideBySideTemplate}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-medium transition-colors"
              title="Download Side-by-Side Programme & College Master Excel Template"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-teal-600 dark:text-teal-400" />
              Side-by-Side Template
            </button>

            <button
              onClick={handleExportAllToExcel}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-medium transition-colors"
              title="Export all defined master lookups to Excel"
            >
              <Download className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400" />
              Export All (.xlsx)
            </button>

            <button
              onClick={handleResetToDefaults}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-medium transition-colors"
              title="Reset to official Kannur University definitions"
            >
              <RotateCcw className="w-3.5 h-3.5 text-amber-500" />
              Reset Defaults
            </button>

            <button
              onClick={() => setShowGuideModal(true)}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 text-xs font-medium transition-colors"
            >
              <HelpCircle className="w-3.5 h-3.5 text-sky-500" />
              Guide
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 w-full flex-1 flex flex-col gap-6">
        {/* Status Alert Banner */}
        {statusAlert && (
          <div
            className={`p-3.5 rounded-xl border flex items-start justify-between gap-3 text-sm shadow-xs ${
              statusAlert.type === 'success'
                ? 'bg-emerald-50 dark:bg-emerald-950/40 border-emerald-200 dark:border-emerald-800 text-emerald-900 dark:text-emerald-200'
                : statusAlert.type === 'warning'
                ? 'bg-amber-50 dark:bg-amber-950/40 border-amber-200 dark:border-amber-800 text-amber-900 dark:text-amber-200'
                : statusAlert.type === 'error'
                ? 'bg-rose-50 dark:bg-rose-950/40 border-rose-200 dark:border-rose-800 text-rose-900 dark:text-rose-200'
                : 'bg-sky-50 dark:bg-sky-950/40 border-sky-200 dark:border-sky-800 text-sky-900 dark:text-sky-200'
            }`}
          >
            <div className="flex items-center gap-2">
              {statusAlert.type === 'success' && <CheckCircle2 className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0" />}
              {statusAlert.type === 'warning' && <AlertCircle className="w-4 h-4 text-amber-600 dark:text-amber-400 shrink-0" />}
              {statusAlert.type === 'error' && <AlertCircle className="w-4 h-4 text-rose-600 dark:text-rose-400 shrink-0" />}
              {statusAlert.type === 'info' && <Sparkles className="w-4 h-4 text-sky-600 dark:text-sky-400 shrink-0" />}
              <span>{statusAlert.message}</span>
            </div>
            <button
              onClick={() => setStatusAlert(null)}
              className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 p-0.5"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        )}

        {/* Hero Stat Overview */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4">
          <div className="p-4 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3.5">
            <div className="p-2.5 rounded-lg bg-teal-50 dark:bg-teal-950/50 text-teal-600 dark:text-teal-400 border border-teal-200 dark:border-teal-800">
              <BookOpen className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Programmes</p>
              <h3 className="text-xl font-bold text-slate-900 dark:text-white">{(lookups.program || []).length}</h3>
              <p className="text-[11px] text-slate-400">Reference ↔ Code ↔ WF</p>
            </div>
          </div>

          <div className="p-4 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3.5">
            <div className="p-2.5 rounded-lg bg-indigo-50 dark:bg-indigo-950/50 text-indigo-600 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-800">
              <Building2 className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Colleges</p>
              <h3 className="text-xl font-bold text-slate-900 dark:text-white">{(lookups.college || []).length}</h3>
              <p className="text-[11px] text-slate-400">ID ↔ Code ↔ Campus</p>
            </div>
          </div>

          <div className="p-4 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3.5">
            <div className="p-2.5 rounded-lg bg-amber-50 dark:bg-amber-950/50 text-amber-600 dark:text-amber-400 border border-amber-200 dark:border-amber-800">
              <Users className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Castes &amp; Categories</p>
              <h3 className="text-xl font-bold text-slate-900 dark:text-white">
                {((lookups.caste || []).length) + ((lookups.category || []).length)}
              </h3>
              <p className="text-[11px] text-slate-400">Social reservation codes</p>
            </div>
          </div>

          <div className="p-4 rounded-xl bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 shadow-xs flex items-center gap-3.5">
            <div className="p-2.5 rounded-lg bg-sky-50 dark:bg-sky-950/50 text-sky-600 dark:text-sky-400 border border-sky-200 dark:border-sky-800">
              <MapPin className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Geographic &amp; Other</p>
              <h3 className="text-xl font-bold text-slate-900 dark:text-white">
                {((lookups.state || []).length) + ((lookups.country || []).length) + ((lookups.religion || []).length) + ((lookups.qualification || []).length) + ((lookups.stream || []).length)}
              </h3>
              <p className="text-[11px] text-slate-400">States, Countries, etc.</p>
            </div>
          </div>
        </div>

        {/* Entity Tabs Container */}
        <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-xs overflow-hidden flex flex-col">
          {/* Tab Navigation */}
          <div className="border-b border-slate-200 dark:border-slate-700 bg-slate-50/70 dark:bg-slate-800/80 px-4 pt-3 flex items-center gap-1.5 overflow-x-auto scrollbar-thin">
            {LOOKUP_TAB_CONFIG.map(tab => {
              const Icon = TAB_ICONS[tab.key] || Tag;
              const count = (lookups && Array.isArray(lookups[tab.key])) ? lookups[tab.key].length : 0;
              const isActive = activeTab === tab.key;

              return (
                <button
                  key={tab.key}
                  onClick={() => handleTabChange(tab.key)}
                  className={`flex items-center gap-2 px-3.5 py-2.5 text-xs font-semibold rounded-t-lg transition-all border-b-2 whitespace-nowrap shrink-0 ${
                    isActive
                      ? 'border-emerald-600 text-emerald-600 dark:text-emerald-400 bg-white dark:bg-slate-800 shadow-2xs'
                      : 'border-transparent text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200 hover:bg-slate-100/60 dark:hover:bg-slate-700/50'
                  }`}
                >
                  <Icon className={`w-3.5 h-3.5 ${isActive ? 'text-emerald-600 dark:text-emerald-400' : 'text-slate-400'}`} />
                  <span>{tab.label}</span>
                  <span className={`px-1.5 py-0.5 rounded-full text-[10px] font-bold ${
                    isActive
                      ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950 dark:text-emerald-300'
                      : 'bg-slate-200 text-slate-600 dark:bg-slate-700 dark:text-slate-300'
                  }`}>
                    {count}
                  </span>
                </button>
              );
            })}
          </div>

          {/* Subheader: Description & Search */}
          <div className="p-4 bg-slate-50/40 dark:bg-slate-800/40 border-b border-slate-200 dark:border-slate-700 flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-sm font-bold text-slate-800 dark:text-slate-200 flex items-center gap-2">
                <span>{currentTabConfig.label} Directory</span>
                <span className="text-xs font-normal text-slate-400">({currentTabConfig.description})</span>
              </h2>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                Showing <strong>{filteredList.length}</strong> of <strong>{currentList.length}</strong> defined entries
              </p>
            </div>

            <div className="flex items-center gap-2 w-full sm:w-auto">
              <div className="relative flex-1 sm:w-72">
                <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }}
                  placeholder={`Search ${currentTabConfig.label.toLowerCase()}...`}
                  className="w-full pl-9 pr-8 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/30 focus:border-emerald-500"
                />
                {searchQuery && (
                  <button
                    onClick={() => setSearchQuery('')}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                  >
                    <X className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>

              <select
                value={pageSize}
                onChange={(e) => { setPageSize(Number(e.target.value)); setPage(0); }}
                className="px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-medium focus:outline-none"
              >
                <option value={15}>15 per page</option>
                <option value={25}>25 per page</option>
                <option value={50}>50 per page</option>
                <option value={100}>100 per page</option>
                <option value={0}>Show All</option>
              </select>
            </div>
          </div>

          {/* Quick Add Bar */}
          <div className="p-4 bg-emerald-50/30 dark:bg-emerald-950/20 border-b border-emerald-100 dark:border-emerald-900/50">
            <form onSubmit={handleAdd} className="flex flex-wrap items-center gap-2">
              <span className="text-xs font-bold text-emerald-800 dark:text-emerald-300 flex items-center gap-1.5 mr-1 shrink-0">
                <Plus className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
                Quick Add:
              </span>

              {activeTab === 'program' ? (
                <>
                  <input
                    type="text"
                    value={newProgRef}
                    onChange={(e) => setNewProgRef(e.target.value)}
                    placeholder="Programme Reference (e.g. Five Year Integrated...)"
                    className="flex-1 min-w-[200px] px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                  <input
                    type="text"
                    value={newProgCode}
                    onChange={(e) => setNewProgCode(e.target.value)}
                    placeholder="ProgramCode (e.g. PFPSYGS25)"
                    className="w-32 px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                  <input
                    type="text"
                    value={newProgWorkflow}
                    onChange={(e) => setNewProgWorkflow(e.target.value)}
                    placeholder="workFlowId (e.g. 151)"
                    className="w-28 px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                  <input
                    type="text"
                    value={newProgName}
                    onChange={(e) => setNewProgName(e.target.value)}
                    placeholder="Official Program Name (with Credits)-Regular-FYIMP..."
                    className="flex-1 min-w-[220px] px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                </>
              ) : activeTab === 'college' ? (
                <>
                  <input
                    type="text"
                    value={newClgId}
                    onChange={(e) => setNewClgId(e.target.value)}
                    placeholder="CollegeId (e.g. 435)"
                    className="w-28 px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                  <input
                    type="text"
                    value={newClgCode}
                    onChange={(e) => setNewClgCode(e.target.value)}
                    placeholder="CollegeCode (e.g. 435)"
                    className="w-28 px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                  <input
                    type="text"
                    value={newClgName}
                    onChange={(e) => setNewClgName(e.target.value)}
                    placeholder="College Name (e.g. Department of Anthropology)"
                    className="flex-1 min-w-[200px] px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                  <input
                    type="text"
                    value={newClgRef}
                    onChange={(e) => setNewClgRef(e.target.value)}
                    placeholder="CLG Name / Campus (e.g. Dr. Janaki Ammal Campus, Palayad)"
                    className="flex-1 min-w-[220px] px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                </>
              ) : (
                <>
                  <input
                    type="text"
                    value={newName}
                    onChange={(e) => setNewName(e.target.value)}
                    placeholder={`${currentTabConfig.nameLabel} (e.g. SEBC-Muslim, Mukkuvan, Kerala, etc.)`}
                    className="flex-1 min-w-[220px] px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                  <input
                    type="text"
                    value={newId}
                    onChange={(e) => setNewId(e.target.value)}
                    placeholder={`${currentTabConfig.idLabel} (e.g. 10, 320, 18, etc.)`}
                    className="w-48 px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                </>
              )}

              <button
                type="submit"
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-2xs transition-colors shrink-0"
              >
                <Plus className="w-3.5 h-3.5" />
                Add Entry
              </button>
            </form>
          </div>

          {/* Table Container */}
          <div className="overflow-x-auto flex-1">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                <tr className="bg-slate-100/80 dark:bg-slate-800/80 text-slate-600 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700 font-semibold uppercase tracking-wider text-[11px]">
                  <th className="py-2.5 px-3 w-12 text-center">#</th>
                  {activeTab === 'program' ? (
                    <>
                      <th className="py-2.5 px-3 min-w-[200px]">Programme Name (Reference)</th>
                      <th className="py-2.5 px-3 w-32">ProgramCode</th>
                      <th className="py-2.5 px-3 w-28">workFlowId</th>
                      <th className="py-2.5 px-3 min-w-[280px]">Program Name</th>
                    </>
                  ) : activeTab === 'college' ? (
                    <>
                      <th className="py-2.5 px-3 w-24">CollegeId</th>
                      <th className="py-2.5 px-3 w-28">CollegeCode</th>
                      <th className="py-2.5 px-3 min-w-[220px]">College Name</th>
                      <th className="py-2.5 px-3 min-w-[280px]">CLG Name (Campus / Reference)</th>
                    </>
                  ) : (
                    <>
                      <th className="py-2.5 px-3 min-w-[240px]">{currentTabConfig.nameHeader}</th>
                      <th className="py-2.5 px-3 w-44">{currentTabConfig.idHeader}</th>
                    </>
                  )}
                  <th className="py-2.5 px-3 w-24 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60">
                {paginatedList.length === 0 ? (
                  <tr>
                    <td colSpan={activeTab === 'program' || activeTab === 'college' ? 6 : 4} className="py-12 text-center text-slate-400 dark:text-slate-500">
                      <div className="flex flex-col items-center justify-center gap-2">
                        <FileSpreadsheet className="w-8 h-8 opacity-40 text-slate-400" />
                        <p className="text-sm font-medium">
                          {searchQuery ? `No entries matching "${searchQuery}"` : `No ${currentTabConfig.label} defined yet.`}
                        </p>
                        <p className="text-xs text-slate-400">
                          Add an entry using the quick add bar above or upload an Excel spreadsheet.
                        </p>
                      </div>
                    </td>
                  </tr>
                ) : (
                  paginatedList.map((item, idx) => {
                    const rowNumber = (pageSize === 0 ? 0 : page * pageSize) + idx + 1;
                    const isEditing = editingIdx === currentList.indexOf(item);

                    return (
                      <tr
                        key={idx}
                        className={`hover:bg-slate-50/80 dark:hover:bg-slate-700/40 transition-colors ${
                          isEditing ? 'bg-amber-50/60 dark:bg-amber-950/20' : ''
                        }`}
                      >
                        <td className="py-2.5 px-3 text-center text-slate-400 font-mono text-[11px]">
                          {rowNumber}
                        </td>

                        {activeTab === 'program' ? (
                          <>
                            <td className="py-2.5 px-3 font-medium text-slate-800 dark:text-slate-100">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editProgRef}
                                  onChange={(e) => setEditProgRef(e.target.value)}
                                  className="w-full px-2 py-1 text-xs rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                item.pgmName
                              )}
                            </td>
                            <td className="py-2.5 px-3 font-mono font-semibold text-emerald-600 dark:text-emerald-400">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editProgCode}
                                  onChange={(e) => setEditProgCode(e.target.value)}
                                  className="w-full px-2 py-1 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                <span className="px-1.5 py-0.5 rounded bg-emerald-50 dark:bg-emerald-950/50 border border-emerald-200 dark:border-emerald-800">
                                  {item.code}
                                </span>
                              )}
                            </td>
                            <td className="py-2.5 px-3 font-mono font-bold text-slate-700 dark:text-slate-300">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editProgWorkflow}
                                  onChange={(e) => setEditProgWorkflow(e.target.value)}
                                  className="w-full px-2 py-1 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                <span className="px-1.5 py-0.5 rounded bg-slate-100 dark:bg-slate-700 border border-slate-200 dark:border-slate-600">
                                  {item.workflowId}
                                </span>
                              )}
                            </td>
                            <td className="py-2.5 px-3 text-slate-600 dark:text-slate-300 text-[11px] leading-relaxed">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editProgName}
                                  onChange={(e) => setEditProgName(e.target.value)}
                                  className="w-full px-2 py-1 text-xs rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                item.name
                              )}
                            </td>
                          </>
                        ) : activeTab === 'college' ? (
                          <>
                            <td className="py-2.5 px-3 font-mono font-bold text-indigo-600 dark:text-indigo-400">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editClgId}
                                  onChange={(e) => setEditClgId(e.target.value)}
                                  className="w-full px-2 py-1 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                <span className="px-1.5 py-0.5 rounded bg-indigo-50 dark:bg-indigo-950/50 border border-indigo-200 dark:border-indigo-800">
                                  {item.id}
                                </span>
                              )}
                            </td>
                            <td className="py-2.5 px-3 font-mono text-slate-600 dark:text-slate-300">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editClgCode}
                                  onChange={(e) => setEditClgCode(e.target.value)}
                                  className="w-full px-2 py-1 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                item.code
                              )}
                            </td>
                            <td className="py-2.5 px-3 font-medium text-slate-800 dark:text-slate-100">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editClgName}
                                  onChange={(e) => setEditClgName(e.target.value)}
                                  className="w-full px-2 py-1 text-xs rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                item.name
                              )}
                            </td>
                            <td className="py-2.5 px-3 text-slate-600 dark:text-slate-300 text-[11px] leading-relaxed">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editClgRef}
                                  onChange={(e) => setEditClgRef(e.target.value)}
                                  className="w-full px-2 py-1 text-xs rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                item.clgRef
                              )}
                            </td>
                          </>
                        ) : (
                          <>
                            <td className="py-2.5 px-3 font-medium text-slate-800 dark:text-slate-100">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editName}
                                  onChange={(e) => setEditName(e.target.value)}
                                  className="w-full px-2 py-1 text-xs rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                item.name
                              )}
                            </td>
                            <td className="py-2.5 px-3 font-mono font-bold text-emerald-600 dark:text-emerald-400">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editId}
                                  onChange={(e) => setEditId(e.target.value)}
                                  className="w-full px-2 py-1 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                <span className="px-1.5 py-0.5 rounded bg-emerald-50 dark:bg-emerald-950/50 border border-emerald-200 dark:border-emerald-800">
                                  {item.id}
                                </span>
                              )}
                            </td>
                          </>
                        )}

                        <td className="py-2.5 px-3 text-right">
                          <div className="flex items-center justify-end gap-1">
                            {isEditing ? (
                              <>
                                <button
                                  onClick={() => handleSaveEdit(item)}
                                  className="p-1 rounded bg-emerald-600 hover:bg-emerald-700 text-white"
                                  title="Save changes"
                                >
                                  <Check className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  onClick={() => setEditingIdx(null)}
                                  className="p-1 rounded border border-slate-300 dark:border-slate-600 text-slate-500 hover:bg-slate-100 dark:hover:bg-slate-700"
                                  title="Cancel edit"
                                >
                                  <X className="w-3.5 h-3.5" />
                                </button>
                              </>
                            ) : (
                              <>
                                <button
                                  onClick={() => handleStartEdit(item, idx)}
                                  className="p-1 rounded text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-700"
                                  title="Edit entry"
                                >
                                  <Edit2 className="w-3.5 h-3.5" />
                                </button>
                                <button
                                  onClick={() => handleDelete(item)}
                                  className="p-1 rounded text-slate-400 hover:text-rose-600 dark:hover:text-rose-400 hover:bg-rose-50 dark:hover:bg-rose-950/40"
                                  title="Delete entry"
                                >
                                  <Trash2 className="w-3.5 h-3.5" />
                                </button>
                              </>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Pagination Footer */}
          {pageSize > 0 && totalPages > 1 && (
            <div className="p-3 bg-slate-50/70 dark:bg-slate-800/70 border-t border-slate-200 dark:border-slate-700 flex items-center justify-between text-xs text-slate-500">
              <span>
                Page <strong>{page + 1}</strong> of <strong>{totalPages}</strong>
              </span>
              <div className="flex items-center gap-1.5">
                <button
                  onClick={() => setPage(p => Math.max(0, p - 1))}
                  disabled={page === 0}
                  className="px-2.5 py-1 rounded border border-slate-300 dark:border-slate-600 disabled:opacity-40 hover:bg-white dark:hover:bg-slate-700 flex items-center gap-1"
                >
                  <ChevronLeft className="w-3.5 h-3.5" /> Prev
                </button>
                <button
                  onClick={() => setPage(p => Math.min(totalPages - 1, p + 1))}
                  disabled={page >= totalPages - 1}
                  className="px-2.5 py-1 rounded border border-slate-300 dark:border-slate-600 disabled:opacity-40 hover:bg-white dark:hover:bg-slate-700 flex items-center gap-1"
                >
                  Next <ChevronRight className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* Guide Modal */}
      {showGuideModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white dark:bg-slate-800 rounded-2xl max-w-3xl w-full border border-slate-200 dark:border-slate-700 shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
            <div className="p-4 sm:p-5 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="p-1.5 rounded-lg bg-emerald-50 dark:bg-emerald-950/60 text-emerald-600 dark:text-emerald-400">
                  <HelpCircle className="w-5 h-5" />
                </span>
                <h3 className="text-base font-bold text-slate-900 dark:text-white">
                  Programme &amp; College Master Definitions Guide
                </h3>
              </div>
              <button
                onClick={() => setShowGuideModal(false)}
                className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 p-1"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 sm:p-6 overflow-y-auto space-y-5 text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
              <div>
                <h4 className="text-sm font-bold text-slate-800 dark:text-slate-100 mb-1">
                  1. What does this tool do?
                </h4>
                <p>
                  This tool manages the master lookup definitions that power Kannur University admission and examination workflows. It resolves fuzzy and non-standard strings into official university IDs for:
                </p>
                <ul className="list-disc pl-5 mt-1.5 space-y-1">
                  <li><strong>Programmes</strong>: Reference Name ↔ ProgramCode ↔ workFlowId ↔ Full Name</li>
                  <li><strong>Colleges</strong>: CollegeId ↔ CollegeCode ↔ College Name ↔ Campus / CLG Name</li>
                  <li><strong>Social &amp; Geographical</strong>: Categories, Castes, Religions, Countries, and States</li>
                  <li><strong>Academic Qualifications</strong>: Qualifying examination and stream specializations</li>
                </ul>
              </div>

              <div>
                <h4 className="text-sm font-bold text-slate-800 dark:text-slate-100 mb-1">
                  2. Side-by-Side Excel Ingestion
                </h4>
                <p>
                  You can upload a single Excel worksheet containing both Programme columns on the left and College columns on the right (separated by an empty column or together). The tool auto-detects both headers simultaneously:
                </p>
                <div className="p-2.5 rounded-lg bg-slate-100 dark:bg-slate-900 font-mono text-[11px] overflow-x-auto mt-2 text-slate-700 dark:text-slate-300">
                  Programme Name (Reference) | ProgramCode | workFlowId | Program Name | [empty] | CollegeId | CollegeCode | College Name | CLG Name
                </div>
              </div>

              <div>
                <h4 className="text-sm font-bold text-slate-800 dark:text-slate-100 mb-1">
                  3. Live Synchronization Across Report Studio
                </h4>
                <p>
                  Any updates, additions, or spreadsheet imports saved here are stored in persistent browser storage (<code className="font-mono text-emerald-600">localStorage</code>). When you run <strong>Admission Import</strong> or <strong>Admission Report Maker - Row Wise</strong>, they instantly utilize your latest definitions!
                </p>
              </div>
            </div>

            <div className="p-4 border-t border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800/80 flex justify-end">
              <button
                onClick={() => setShowGuideModal(false)}
                className="px-4 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold"
              >
                Got It
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
