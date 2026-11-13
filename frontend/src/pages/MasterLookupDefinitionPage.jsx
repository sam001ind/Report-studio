import React, { useState, useMemo, useRef, useEffect, useCallback } from 'react';
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
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Columns2,
  Table,
  Filter,
  Copy,
  CheckCheck,
  Zap,
  ExternalLink,
  ShieldAlert
} from 'lucide-react';
import {
  LOOKUP_TAB_CONFIG,
  DEFAULT_LOOKUP_DATA,
  loadSavedLookups,
  saveLookupsToStorage,
  parseUploadedLookupSpreadsheet,
  buildLookupMaps,
  normalizeLookupKey
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
  const [viewMode, setViewMode] = useState('split'); // 'split' (Programme + College side-by-side) | 'tabbed' (All 9 categories)
  const [searchQuery, setSearchQuery] = useState('');
  const [pageSize, setPageSize] = useState(25);
  const [page, setPage] = useState(0);
  const [statusAlert, setStatusAlert] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showGuideModal, setShowGuideModal] = useState(false);
  const [showResetConfirmModal, setShowResetConfirmModal] = useState(false);
  const [copiedKey, setCopiedKey] = useState(null);

  // Sorting state: { column: 'name' | 'id' | 'code' | 'workflowId' | ..., direction: 'asc' | 'desc' }
  const [sortConfig, setSortConfig] = useState({ column: null, direction: 'asc' });

  // Quick Filter Chip (e.g. for Programme: 'all' | 'fyimp' | 'pg' | 'ug')
  const [filterChip, setFilterChip] = useState('all');

  // Multi-row Selection
  const [selectedRowKeys, setSelectedRowKeys] = useState(new Set());

  // Live String Resolver Simulator state
  const [showResolver, setShowResolver] = useState(true);
  const [resolverInput, setResolverInput] = useState('');

  // Side-by-Side Dual View independent searches
  const [progSearch, setProgSearch] = useState('');
  const [clgSearch, setClgSearch] = useState('');
  const [progPage, setProgPage] = useState(0);
  const [clgPage, setClgPage] = useState(0);
  const dualPageSize = 15;

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

  // Add Form Expansion Toggle
  const [isAddFormOpen, setIsAddFormOpen] = useState(false);

  // Inline edit state
  const [editingIdx, setEditingIdx] = useState(null);
  const [editingTargetTab, setEditingTargetTab] = useState(null);
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

  // Sync state changes to storage
  const updateLookups = useCallback((newLookups) => {
    setLookups(newLookups);
    saveLookupsToStorage(newLookups);
    setSelectedRowKeys(new Set());
  }, []);

  const currentTabConfig = useMemo(() => {
    return LOOKUP_TAB_CONFIG.find(t => t.key === activeTab) || LOOKUP_TAB_CONFIG[0];
  }, [activeTab]);

  const currentList = useMemo(() => {
    return (lookups && Array.isArray(lookups[activeTab])) ? lookups[activeTab] : [];
  }, [lookups, activeTab]);

  const totalEntriesCount = useMemo(() => {
    if (!lookups) return 0;
    return Object.values(lookups).reduce((sum, list) => sum + (Array.isArray(list) ? list.length : 0), 0);
  }, [lookups]);

  // Copy helper
  const handleCopy = (text, key) => {
    if (!text) return;
    navigator.clipboard?.writeText(String(text));
    setCopiedKey(key);
    setTimeout(() => {
      setCopiedKey(null);
    }, 1800);
  };

  // Live Resolver Engine
  const resolvedResult = useMemo(() => {
    const q = resolverInput.trim();
    if (!q) return null;
    const norm = normalizeLookupKey(q);
    const maps = buildLookupMaps(lookups);

    // 1. Program exact / normalized
    if (maps.programByName[norm] || maps.programByName[q.toLowerCase()]) {
      const p = maps.programByName[norm] || maps.programByName[q.toLowerCase()];
      return {
        entity: 'Programme',
        tabKey: 'program',
        matchType: 'Official Reference Match',
        title: p.pgmName || p.name,
        code: p.code,
        id: p.workflowId,
        detail: p.name,
        confidence: 'High (Exact)',
        color: '#059669',
        bgColor: 'rgba(5, 150, 105, 0.1)'
      };
    }
    if (maps.programByCode[q.toLowerCase()] || maps.programByCode[q]) {
      const p = maps.programByCode[q.toLowerCase()] || maps.programByCode[q];
      return {
        entity: 'Programme',
        tabKey: 'program',
        matchType: 'Program Code Match',
        title: p.pgmName,
        code: p.code,
        id: p.workflowId,
        detail: p.name,
        confidence: 'High (Code)',
        color: '#059669',
        bgColor: 'rgba(5, 150, 105, 0.1)'
      };
    }

    // 2. College exact / normalized
    if (maps.collegeByName[norm] || maps.collegeByName[q.toLowerCase()]) {
      const c = maps.collegeByName[norm] || maps.collegeByName[q.toLowerCase()];
      return {
        entity: 'College',
        tabKey: 'college',
        matchType: 'College / Campus Match',
        title: c.name,
        code: c.code,
        id: c.id,
        detail: c.clgRef,
        confidence: 'High (Exact)',
        color: '#4f46e5',
        bgColor: 'rgba(79, 70, 229, 0.1)'
      };
    }
    if (maps.collegeByCode[q] || maps.collegeById[q]) {
      const c = maps.collegeByCode[q] || maps.collegeById[q];
      return {
        entity: 'College',
        tabKey: 'college',
        matchType: 'College Code / ID Match',
        title: c.name,
        code: c.code,
        id: c.id,
        detail: c.clgRef,
        confidence: 'High (ID)',
        color: '#4f46e5',
        bgColor: 'rgba(79, 70, 229, 0.1)'
      };
    }

    // 3. State
    if (maps.stateByName[norm] || maps.stateByName[q.toLowerCase()]) {
      const s = maps.stateByName[norm] || maps.stateByName[q.toLowerCase()];
      return {
        entity: 'State',
        tabKey: 'state',
        matchType: 'State Master Match',
        title: s.name,
        id: s.id,
        detail: `corrStateId: ${s.id}`,
        confidence: 'High (State ID)',
        color: '#0284c7',
        bgColor: 'rgba(2, 132, 199, 0.1)'
      };
    }

    // 4. Caste
    if (maps.casteByName[norm] || maps.casteByName[q.toLowerCase()]) {
      const c = maps.casteByName[norm] || maps.casteByName[q.toLowerCase()];
      return {
        entity: 'Caste',
        tabKey: 'caste',
        matchType: 'Caste Master Match',
        title: c.name,
        id: c.id,
        detail: `Caste Reservation ID: ${c.id}`,
        confidence: 'High (Reservation ID)',
        color: '#d97706',
        bgColor: 'rgba(217, 119, 6, 0.1)'
      };
    }

    // 5. Category
    if (maps.categoryByName[norm] || maps.categoryByName[q.toLowerCase()]) {
      const cat = maps.categoryByName[norm] || maps.categoryByName[q.toLowerCase()];
      return {
        entity: 'Category',
        tabKey: 'category',
        matchType: 'Category Master Match',
        title: cat.name,
        id: cat.id,
        detail: `Category ID: ${cat.id}`,
        confidence: 'High (Category)',
        color: '#e11d48',
        bgColor: 'rgba(225, 29, 72, 0.1)'
      };
    }

    // 6. Religion
    if (maps.religionByName[norm] || maps.religionByName[q.toLowerCase()]) {
      const r = maps.religionByName[norm] || maps.religionByName[q.toLowerCase()];
      return {
        entity: 'Religion',
        tabKey: 'religion',
        matchType: 'Religion Match',
        title: r.name,
        id: r.id,
        detail: `Religion ID: ${r.id}`,
        confidence: 'High',
        color: '#9333ea',
        bgColor: 'rgba(147, 51, 234, 0.1)'
      };
    }

    // 7. Qualification
    if (maps.qualificationByName[norm] || maps.qualificationByName[q.toLowerCase()]) {
      const qu = maps.qualificationByName[norm] || maps.qualificationByName[q.toLowerCase()];
      return {
        entity: 'Qualification',
        tabKey: 'qualification',
        matchType: 'Examination Board Match',
        title: qu.name,
        id: qu.id,
        detail: `Faculty ID: ${qu.id}`,
        confidence: 'High',
        color: '#0d9488',
        bgColor: 'rgba(13, 148, 136, 0.1)'
      };
    }

    // 8. Stream
    if (maps.streamByName[norm] || maps.streamByName[q.toLowerCase()]) {
      const st = maps.streamByName[norm] || maps.streamByName[q.toLowerCase()];
      return {
        entity: 'Stream',
        tabKey: 'stream',
        matchType: 'Stream Specialization Match',
        title: st.name,
        id: st.id,
        detail: `Specialization ID: ${st.id}`,
        confidence: 'High',
        color: '#2563eb',
        bgColor: 'rgba(37, 99, 235, 0.1)'
      };
    }

    // 9. Fuzzy Keyword fallbacks
    for (const p of (lookups.program || [])) {
      if ((p.pgmName || '').toLowerCase().includes(q.toLowerCase()) || (p.name || '').toLowerCase().includes(q.toLowerCase())) {
        return {
          entity: 'Programme',
          tabKey: 'program',
          matchType: 'Fuzzy Keyword Match',
          title: p.pgmName,
          code: p.code,
          id: p.workflowId,
          detail: p.name,
          confidence: 'Medium (Keyword)',
          color: '#059669',
          bgColor: 'rgba(5, 150, 105, 0.1)'
        };
      }
    }

    for (const c of (lookups.college || [])) {
      if ((c.name || '').toLowerCase().includes(q.toLowerCase()) || (c.clgRef || '').toLowerCase().includes(q.toLowerCase())) {
        return {
          entity: 'College',
          tabKey: 'college',
          matchType: 'Fuzzy Keyword Match',
          title: c.name,
          code: c.code,
          id: c.id,
          detail: c.clgRef,
          confidence: 'Medium (Keyword)',
          color: '#4f46e5',
          bgColor: 'rgba(79, 70, 229, 0.1)'
        };
      }
    }

    return {
      entity: 'Unmapped',
      tabKey: null,
      matchType: 'No Master Match Found',
      title: q,
      id: 'N/A',
      detail: 'This string will use default fallback during admission import.',
      confidence: 'None',
      color: 'var(--muted)',
      bgColor: 'var(--bg)'
    };
  }, [resolverInput, lookups]);

  // Handle Tab Switch
  const handleTabChange = (newTab) => {
    setActiveTab(newTab);
    setSearchQuery('');
    setEditingIdx(null);
    setPage(0);
    setSortConfig({ column: null, direction: 'asc' });
    setFilterChip('all');
    setSelectedRowKeys(new Set());
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
    setIsAddFormOpen(false);
    setStatusAlert(null);
  };

  // Filtered and Sorted list for active tab
  const filteredList = useMemo(() => {
    let list = currentList;
    const q = searchQuery.trim().toLowerCase();

    // 1. Quick Chip Filtering
    if (activeTab === 'program') {
      if (filterChip === 'fyimp') {
        list = list.filter(p => (p.name || '').includes('FYIMP') || (p.code || '').startsWith('PF'));
      } else if (filterChip === 'regular') {
        list = list.filter(p => !(p.name || '').includes('FYIMP'));
      }
    } else if (activeTab === 'college') {
      if (filterChip === 'dept') {
        list = list.filter(c => (c.name || '').toLowerCase().includes('department'));
      } else if (filterChip === 'campus') {
        list = list.filter(c => (c.name || '').toLowerCase().includes('campus') || (c.clgRef || '').toLowerCase().includes('campus'));
      }
    } else if (activeTab === 'caste') {
      if (filterChip === 'sebc') {
        list = list.filter(c => Number(c.id) >= 300 || Number(c.id) <= 200);
      } else if (filterChip === 'sc') {
        list = list.filter(c => Number(c.id) >= 220 && Number(c.id) <= 250);
      }
    }

    // 2. Search Query Filtering
    if (q) {
      if (activeTab === 'program') {
        list = list.filter(item => {
          const ref = String(item.pgmName || '').toLowerCase();
          const code = String(item.code || '').toLowerCase();
          const wf = String(item.workflowId || '').toLowerCase();
          const name = String(item.name || '').toLowerCase();
          return ref.includes(q) || code.includes(q) || wf.includes(q) || name.includes(q);
        });
      } else if (activeTab === 'college') {
        list = list.filter(item => {
          const id = String(item.id || '').toLowerCase();
          const code = String(item.code || '').toLowerCase();
          const name = String(item.name || '').toLowerCase();
          const ref = String(item.clgRef || '').toLowerCase();
          return id.includes(q) || code.includes(q) || name.includes(q) || ref.includes(q);
        });
      } else {
        list = list.filter(item => {
          const n = String(item.name || '').toLowerCase();
          const id = String(item.id || '').toLowerCase();
          return n.includes(q) || id.includes(q);
        });
      }
    }

    // 3. Sorting
    if (sortConfig.column) {
      const col = sortConfig.column;
      const factor = sortConfig.direction === 'asc' ? 1 : -1;
      list = [...list].sort((a, b) => {
        let va = a[col] ?? '';
        let vb = b[col] ?? '';
        if (!isNaN(Number(va)) && !isNaN(Number(vb)) && va !== '' && vb !== '') {
          return (Number(va) - Number(vb)) * factor;
        }
        return String(va).localeCompare(String(vb)) * factor;
      });
    }

    return list;
  }, [currentList, searchQuery, activeTab, filterChip, sortConfig]);

  const paginatedList = useMemo(() => {
    if (pageSize === 0) return filteredList;
    const start = page * pageSize;
    return filteredList.slice(start, start + pageSize);
  }, [filteredList, page, pageSize]);

  const totalPages = pageSize === 0 ? 1 : Math.ceil(filteredList.length / pageSize) || 1;

  // Dual View Filtered Lists
  const dualProgList = useMemo(() => {
    const q = progSearch.trim().toLowerCase();
    const list = lookups.program || [];
    if (!q) return list;
    return list.filter(p => 
      (p.pgmName || '').toLowerCase().includes(q) ||
      (p.code || '').toLowerCase().includes(q) ||
      (p.workflowId || '').toLowerCase().includes(q) ||
      (p.name || '').toLowerCase().includes(q)
    );
  }, [lookups.program, progSearch]);

  const dualClgList = useMemo(() => {
    const q = clgSearch.trim().toLowerCase();
    const list = lookups.college || [];
    if (!q) return list;
    return list.filter(c => 
      (c.name || '').toLowerCase().includes(q) ||
      (c.id || '').toLowerCase().includes(q) ||
      (c.code || '').toLowerCase().includes(q) ||
      (c.clgRef || '').toLowerCase().includes(q)
    );
  }, [lookups.college, clgSearch]);

  const dualPaginatedProgs = useMemo(() => {
    return dualProgList.slice(progPage * dualPageSize, (progPage + 1) * dualPageSize);
  }, [dualProgList, progPage]);

  const dualPaginatedColleges = useMemo(() => {
    return dualClgList.slice(clgPage * dualPageSize, (clgPage + 1) * dualPageSize);
  }, [dualClgList, clgPage]);

  // Column Sort Handler
  const handleSort = (column) => {
    setSortConfig(prev => {
      if (prev.column === column) {
        if (prev.direction === 'asc') return { column, direction: 'desc' };
        return { column: null, direction: 'asc' };
      }
      return { column, direction: 'asc' };
    });
  };

  // Add Entry Handler
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
          message: 'Please fill all 4 fields for Programme (Reference Name, ProgramCode, workFlowId, Full Name).'
        });
        return;
      }

      const exists = (lookups.program || []).some(item => (item.pgmName || '').toLowerCase() === pgmName.toLowerCase());
      if (exists) {
        setStatusAlert({
          type: 'warning',
          message: `Programme reference "${pgmName}" already exists.`
        });
        return;
      }

      const updatedTabList = [{ pgmName, code, workflowId, name }, ...(lookups.program || [])];
      updateLookups({ ...lookups, program: updatedTabList });
      setNewProgRef('');
      setNewProgCode('');
      setNewProgWorkflow('');
      setNewProgName('');
      setIsAddFormOpen(false);
      setStatusAlert({
        type: 'success',
        message: `Successfully added Programme "${pgmName}" (Code: ${code}, Workflow: ${workflowId}).`
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

      const exists = (lookups.college || []).some(item => (item.id || '') === id || (item.name || '').toLowerCase() === name.toLowerCase());
      if (exists) {
        setStatusAlert({
          type: 'warning',
          message: `College with ID ${id} or Name "${name}" already exists.`
        });
        return;
      }

      const updatedTabList = [{ id, code, name, clgRef }, ...(lookups.college || [])];
      updateLookups({ ...lookups, college: updatedTabList });
      setNewClgId('');
      setNewClgCode('');
      setNewClgName('');
      setNewClgRef('');
      setIsAddFormOpen(false);
      setStatusAlert({
        type: 'success',
        message: `Successfully added College "${name}" (ID: ${id}).`
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
    updateLookups({ ...lookups, [activeTab]: updatedTabList });
    setNewName('');
    setNewId('');
    setIsAddFormOpen(false);
    setStatusAlert({
      type: 'success',
      message: `Added "${trimmedName}" -> "${trimmedId}" to ${currentTabConfig.label}.`
    });
  };

  // Inline Edit Handlers
  const handleStartEdit = (item, targetTab = activeTab) => {
    setEditingTargetTab(targetTab);
    const list = lookups[targetTab] || [];
    const realIdx = list.indexOf(item);
    setEditingIdx(realIdx);

    if (targetTab === 'program') {
      setEditProgRef(item.pgmName || '');
      setEditProgCode(item.code || '');
      setEditProgWorkflow(item.workflowId || '');
      setEditProgName(item.name || '');
    } else if (targetTab === 'college') {
      setEditClgId(item.id || '');
      setEditClgCode(item.code || '');
      setEditClgName(item.name || '');
      setEditClgRef(item.clgRef || '');
    } else {
      setEditName(item.name || '');
      setEditId(item.id || '');
    }
  };

  const handleSaveEdit = (targetTab = activeTab) => {
    const list = lookups[targetTab] || [];
    if (editingIdx === null || editingIdx < 0 || editingIdx >= list.length) return;

    const updatedTabList = [...list];

    if (targetTab === 'program') {
      const pgmName = editProgRef.trim();
      const code = editProgCode.trim();
      const workflowId = editProgWorkflow.trim();
      const name = editProgName.trim();
      if (!pgmName || !code || !workflowId || !name) {
        setStatusAlert({ type: 'warning', message: 'All 4 fields are required.' });
        return;
      }
      updatedTabList[editingIdx] = { pgmName, code, workflowId, name };
    } else if (targetTab === 'college') {
      const id = editClgId.trim();
      const code = editClgCode.trim() || id;
      const name = editClgName.trim();
      const clgRef = editClgRef.trim() || name;
      if (!id || !name) {
        setStatusAlert({ type: 'warning', message: 'CollegeId and College Name are required.' });
        return;
      }
      updatedTabList[editingIdx] = { id, code, name, clgRef };
    } else {
      const name = editName.trim();
      const id = editId.trim();
      if (!name || !id) {
        setStatusAlert({ type: 'warning', message: 'Name and ID are required.' });
        return;
      }
      updatedTabList[editingIdx] = { name, id };
    }

    updateLookups({ ...lookups, [targetTab]: updatedTabList });
    setEditingIdx(null);
    setEditingTargetTab(null);
    setStatusAlert({
      type: 'success',
      message: `Updated entry in ${targetTab.toUpperCase()}.`
    });
  };

  const handleDelete = (item, targetTab = activeTab) => {
    const list = lookups[targetTab] || [];
    const realIdx = list.indexOf(item);
    if (realIdx === -1) return;

    const label = targetTab === 'program' ? item.pgmName : (targetTab === 'college' ? item.name : item.name);
    if (!window.confirm(`Delete "${label}" from master definitions?`)) return;

    const updatedTabList = list.filter((_, i) => i !== realIdx);
    updateLookups({ ...lookups, [targetTab]: updatedTabList });
    setEditingIdx(null);
    setStatusAlert({
      type: 'success',
      message: `Deleted "${label}" from ${targetTab}.`
    });
  };

  // Bulk Delete
  const handleBulkDelete = () => {
    if (selectedRowKeys.size === 0) return;
    if (!window.confirm(`Delete ${selectedRowKeys.size} selected entries from ${currentTabConfig.label}?`)) return;

    const updatedTabList = currentList.filter((item, idx) => {
      const key = `${activeTab}_${idx}_${item.id || item.code || item.name}`;
      return !selectedRowKeys.has(key);
    });

    updateLookups({ ...lookups, [activeTab]: updatedTabList });
    setSelectedRowKeys(new Set());
    setStatusAlert({
      type: 'success',
      message: `Removed ${selectedRowKeys.size} entries from ${currentTabConfig.label}.`
    });
  };

  // Reset to Defaults
  const handleConfirmReset = () => {
    updateLookups(DEFAULT_LOOKUP_DATA);
    setEditingIdx(null);
    setPage(0);
    setShowResetConfirmModal(false);
    setStatusAlert({
      type: 'success',
      message: 'Reset all 9 master ID definitions to official Kannur University defaults.'
    });
  };

  // File Upload Ingestion
  const handleFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setStatusAlert({
      type: 'info',
      message: `Analyzing spreadsheet "${file.name}"...`
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
            message: `No matching master columns found in "${file.name}". Expected headers like "Programme Name (Reference)", "ProgramCode", "CollegeId", "College Name", etc.`
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
          message: `Successfully ingested ${total} rows (${newEntriesCount} new records added): ${summaryText}. Saved to active definitions.`
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

  // Download Side-by-Side Template
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
        'Five Year Integrated Masters Programme in Music',
        'PFMUSGA25',
        '153',
        'Five Year Integrated Masters Programme in Music(with Credits)-Regular-FYIMP-2025-Music-FYIMP Year I SEMESTER I',
        '',
        '127',
        '127',
        'Department of Music',
        'Department of Music'
      ]
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    ws['!cols'] = [
      { wch: 45 }, { wch: 15 }, { wch: 12 }, { wch: 70 }, { wch: 4 },
      { wch: 12 }, { wch: 14 }, { wch: 38 }, { wch: 65 }
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
        message: 'Downloaded Defined_Admission_Master_Lookups.xlsx containing all 9 master tables.'
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
    <div style={{ minHeight: '100vh', width: '100vw', background: 'var(--bg)', color: 'var(--ink)', display: 'flex', flexDirection: 'column', boxSizing: 'border-box' }}>
      {/* Hidden File Input */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileUpload}
        accept=".xlsx,.xls,.csv"
        style={{ display: 'none' }}
      />

      {/* Top Header Command Bar */}
      <header style={{
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        padding: '12px 24px',
        background: 'var(--panel)',
        borderBottom: '1px solid var(--line)',
        position: 'sticky',
        top: 0,
        zIndex: 50,
        flexWrap: 'wrap',
        gap: '12px'
      }}>
        {/* Left: Branding & Breadcrumbs */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
          <Link
            to="/"
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              padding: '6px 12px',
              borderRadius: '6px',
              border: '1px solid var(--line)',
              background: 'var(--bg)',
              color: 'var(--ink)',
              textDecoration: 'none',
              fontSize: '12px',
              fontWeight: 600
            }}
          >
            <ArrowLeft size={14} /> Back to Dashboard
          </Link>
          <div style={{ height: '18px', width: '1px', background: 'var(--line)' }} />
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{
                padding: '4px',
                borderRadius: '6px',
                background: 'var(--accent-soft)',
                color: 'var(--accent)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>
                <Database size={16} />
              </span>
              <h1 style={{ fontSize: '17px', fontWeight: 800, margin: 0, letterSpacing: '-0.2px' }}>
                Programme &amp; College Master Definitions
              </h1>
              <span style={{
                fontSize: '11px',
                fontWeight: 700,
                padding: '2px 8px',
                borderRadius: '12px',
                background: 'var(--accent-soft)',
                color: 'var(--accent)',
                border: '1px solid var(--line)'
              }}>
                {totalEntriesCount} Masters Loaded
              </span>
            </div>
            <p style={{ margin: '2px 0 0 0', fontSize: '11.5px', color: 'var(--muted)' }}>
              Canonical master reference directory for Kannur University admission codes, workflow IDs, campuses, and reservation quotas.
            </p>
          </div>
        </div>

        {/* Right: Actions & Tools */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
          {/* View Mode Toggle: Dual Split vs Tabbed */}
          <div style={{
            display: 'inline-flex',
            background: 'var(--bg)',
            border: '1px solid var(--line)',
            borderRadius: '6px',
            padding: '2px',
            gap: '2px'
          }}>
            <button
              onClick={() => setViewMode('split')}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '5px',
                padding: '5px 10px',
                fontSize: '11.5px',
                fontWeight: 600,
                borderRadius: '4px',
                background: viewMode === 'split' ? 'var(--accent)' : 'transparent',
                color: viewMode === 'split' ? '#fff' : 'var(--muted)',
                border: 'none',
                cursor: 'pointer'
              }}
            >
              <Columns2 size={13} /> Side-by-Side View
            </button>
            <button
              onClick={() => setViewMode('tabbed')}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: '5px',
                padding: '5px 10px',
                fontSize: '11.5px',
                fontWeight: 600,
                borderRadius: '4px',
                background: viewMode === 'tabbed' ? 'var(--accent)' : 'transparent',
                color: viewMode === 'tabbed' ? '#fff' : 'var(--muted)',
                border: 'none',
                cursor: 'pointer'
              }}
            >
              <Table size={13} /> All 9 Masters
            </button>
          </div>

          <button
            onClick={() => fileInputRef.current?.click()}
            disabled={isProcessing}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              padding: '6px 12px',
              borderRadius: '6px',
              background: 'var(--accent)',
              color: '#fff',
              fontSize: '12px',
              fontWeight: 600,
              border: 'none',
              cursor: 'pointer'
            }}
          >
            <Upload size={14} /> Upload Spreadsheet
          </button>

          <button
            onClick={handleDownloadSideBySideTemplate}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              padding: '6px 12px',
              borderRadius: '6px',
              background: 'var(--panel)',
              color: 'var(--ink)',
              fontSize: '12px',
              fontWeight: 600,
              border: '1px solid var(--line)',
              cursor: 'pointer'
            }}
          >
            <FileSpreadsheet size={14} color="var(--accent)" /> Side-by-Side Template
          </button>

          <button
            onClick={handleExportAllToExcel}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              padding: '6px 12px',
              borderRadius: '6px',
              background: 'var(--panel)',
              color: 'var(--ink)',
              fontSize: '12px',
              fontWeight: 600,
              border: '1px solid var(--line)',
              cursor: 'pointer'
            }}
          >
            <Download size={14} color="var(--accent)" /> Export All (.xlsx)
          </button>

          <button
            onClick={() => setShowResetConfirmModal(true)}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              padding: '6px 10px',
              borderRadius: '6px',
              background: 'var(--panel)',
              color: 'var(--ink)',
              fontSize: '12px',
              fontWeight: 600,
              border: '1px solid var(--line)',
              cursor: 'pointer'
            }}
          >
            <RotateCcw size={13} color="#d97706" /> Reset Defaults
          </button>

          <button
            onClick={() => setShowGuideModal(true)}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              padding: '6px 10px',
              borderRadius: '6px',
              background: 'var(--panel)',
              color: 'var(--ink)',
              fontSize: '12px',
              fontWeight: 600,
              border: '1px solid var(--line)',
              cursor: 'pointer'
            }}
          >
            <HelpCircle size={14} color="#0284c7" /> Guide
          </button>
        </div>
      </header>

      {/* Main Content Body */}
      <main style={{ maxWidth: '1440px', width: '100%', margin: '0 auto', padding: '20px 24px', flex: 1, display: 'flex', flexDirection: 'column', gap: '16px', boxSizing: 'border-box' }}>
        {/* Status Alert Banner */}
        {statusAlert && (
          <div style={{
            padding: '10px 16px',
            borderRadius: '8px',
            fontSize: '13px',
            fontWeight: 500,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            background: statusAlert.type === 'success' ? 'var(--accent-soft)' : statusAlert.type === 'warning' ? 'rgba(217, 119, 6, 0.12)' : 'rgba(220, 38, 38, 0.12)',
            color: statusAlert.type === 'success' ? 'var(--accent)' : statusAlert.type === 'warning' ? '#b45309' : '#b91c1c',
            border: `1px solid ${statusAlert.type === 'success' ? 'var(--accent)' : statusAlert.type === 'warning' ? '#f59e0b' : '#ef4444'}`
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              {statusAlert.type === 'success' ? <CheckCircle2 size={16} /> : <AlertCircle size={16} />}
              <span>{statusAlert.message}</span>
            </div>
            <button
              onClick={() => setStatusAlert(null)}
              style={{ background: 'transparent', border: 'none', color: 'inherit', padding: '2px', cursor: 'pointer' }}
            >
              <X size={15} />
            </button>
          </div>
        )}

        {/* Hero Interactive Stat Overview Cards */}
        <div style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
          gap: '14px'
        }}>
          {/* Card 1: Programmes */}
          <div
            onClick={() => { setViewMode('split'); handleTabChange('program'); }}
            style={{
              background: 'var(--panel)',
              border: (activeTab === 'program' || viewMode === 'split') ? '2px solid var(--accent)' : '1px solid var(--line)',
              borderRadius: '10px',
              padding: '14px 18px',
              cursor: 'pointer',
              boxShadow: '0 2px 8px rgba(0,0,0,0.03)',
              transition: 'all 0.15s ease'
            }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ padding: '8px', borderRadius: '8px', background: 'rgba(5, 150, 105, 0.12)', color: '#059669', display: 'flex' }}>
                <BookOpen size={18} />
              </div>
              <span style={{ fontSize: '10.5px', fontWeight: 700, textTransform: 'uppercase', color: '#059669', background: 'rgba(5, 150, 105, 0.1)', padding: '2px 8px', borderRadius: '10px' }}>
                Active Master
              </span>
            </div>
            <div style={{ marginTop: '10px' }}>
              <p style={{ margin: 0, fontSize: '11px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Programmes</p>
              <h3 style={{ margin: '4px 0 2px 0', fontSize: '22px', fontWeight: 800 }}>{(lookups.program || []).length}</h3>
              <p style={{ margin: 0, fontSize: '11.5px', color: 'var(--muted)' }}>Reference ↔ Code ↔ WF ID</p>
            </div>
          </div>

          {/* Card 2: Colleges */}
          <div
            onClick={() => { setViewMode('split'); handleTabChange('college'); }}
            style={{
              background: 'var(--panel)',
              border: (activeTab === 'college' || viewMode === 'split') ? '2px solid #4f46e5' : '1px solid var(--line)',
              borderRadius: '10px',
              padding: '14px 18px',
              cursor: 'pointer',
              boxShadow: '0 2px 8px rgba(0,0,0,0.03)',
              transition: 'all 0.15s ease'
            }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ padding: '8px', borderRadius: '8px', background: 'rgba(79, 70, 229, 0.12)', color: '#4f46e5', display: 'flex' }}>
                <Building2 size={18} />
              </div>
              <span style={{ fontSize: '10.5px', fontWeight: 700, textTransform: 'uppercase', color: '#4f46e5', background: 'rgba(79, 70, 229, 0.1)', padding: '2px 8px', borderRadius: '10px' }}>
                Active Master
              </span>
            </div>
            <div style={{ marginTop: '10px' }}>
              <p style={{ margin: 0, fontSize: '11px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Colleges &amp; Campuses</p>
              <h3 style={{ margin: '4px 0 2px 0', fontSize: '22px', fontWeight: 800 }}>{(lookups.college || []).length}</h3>
              <p style={{ margin: 0, fontSize: '11.5px', color: 'var(--muted)' }}>ID ↔ Code ↔ Campus Name</p>
            </div>
          </div>

          {/* Card 3: Social Quota */}
          <div
            onClick={() => { setViewMode('tabbed'); handleTabChange('caste'); }}
            style={{
              background: 'var(--panel)',
              border: (activeTab === 'caste' || activeTab === 'category') && viewMode === 'tabbed' ? '2px solid #d97706' : '1px solid var(--line)',
              borderRadius: '10px',
              padding: '14px 18px',
              cursor: 'pointer',
              boxShadow: '0 2px 8px rgba(0,0,0,0.03)',
              transition: 'all 0.15s ease'
            }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ padding: '8px', borderRadius: '8px', background: 'rgba(217, 119, 6, 0.12)', color: '#d97706', display: 'flex' }}>
                <Users size={18} />
              </div>
              <span style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)' }}>
                {(lookups.caste || []).length} Castes
              </span>
            </div>
            <div style={{ marginTop: '10px' }}>
              <p style={{ margin: 0, fontSize: '11px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Social Quota</p>
              <h3 style={{ margin: '4px 0 2px 0', fontSize: '22px', fontWeight: 800 }}>
                {((lookups.caste || []).length) + ((lookups.category || []).length)}
              </h3>
              <p style={{ margin: 0, fontSize: '11.5px', color: 'var(--muted)' }}>Castes &amp; Reservation Categories</p>
            </div>
          </div>

          {/* Card 4: Geography & Boards */}
          <div
            onClick={() => { setViewMode('tabbed'); handleTabChange('state'); }}
            style={{
              background: 'var(--panel)',
              border: (activeTab === 'state' || activeTab === 'country' || activeTab === 'qualification' || activeTab === 'stream') && viewMode === 'tabbed' ? '2px solid #0284c7' : '1px solid var(--line)',
              borderRadius: '10px',
              padding: '14px 18px',
              cursor: 'pointer',
              boxShadow: '0 2px 8px rgba(0,0,0,0.03)',
              transition: 'all 0.15s ease'
            }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ padding: '8px', borderRadius: '8px', background: 'rgba(2, 132, 199, 0.12)', color: '#0284c7', display: 'flex' }}>
                <MapPin size={18} />
              </div>
              <span style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)' }}>
                {(lookups.state || []).length} States
              </span>
            </div>
            <div style={{ marginTop: '10px' }}>
              <p style={{ margin: 0, fontSize: '11px', fontWeight: 600, color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.4px' }}>Geography &amp; Boards</p>
              <h3 style={{ margin: '4px 0 2px 0', fontSize: '22px', fontWeight: 800 }}>
                {((lookups.state || []).length) + ((lookups.country || []).length) + ((lookups.qualification || []).length) + ((lookups.stream || []).length)}
              </h3>
              <p style={{ margin: 0, fontSize: '11.5px', color: 'var(--muted)' }}>States, Boards, Streams, etc.</p>
            </div>
          </div>
        </div>

        {/* Live Lookup Resolver / Validator Sandbox */}
        <div style={{
          background: 'var(--panel)',
          border: '1px solid var(--line)',
          borderRadius: '10px',
          overflow: 'hidden',
          boxShadow: '0 2px 8px rgba(0,0,0,0.02)'
        }}>
          <div style={{
            padding: '10px 16px',
            background: 'var(--bg)',
            borderBottom: '1px solid var(--line)',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center'
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span style={{ color: '#d97706', display: 'flex' }}><Zap size={15} /></span>
              <strong style={{ fontSize: '12px', textTransform: 'uppercase', letterSpacing: '0.4px' }}>
                Interactive String Lookup Resolver &amp; Validator
              </strong>
              <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>
                (Simulates how strings in candidate spreadsheets map to university IDs)
              </span>
            </div>
            <button
              onClick={() => setShowResolver(!showResolver)}
              style={{ background: 'transparent', border: 'none', color: 'var(--accent)', cursor: 'pointer', fontSize: '11.5px', fontWeight: 600 }}
            >
              {showResolver ? 'Hide Simulator' : 'Show Simulator'}
            </button>
          </div>

          {showResolver && (
            <div style={{ padding: '14px 18px' }}>
              <div style={{ display: 'flex', gap: '10px', alignItems: 'center', flexWrap: 'wrap' }}>
                <div style={{ position: 'relative', flex: 1, minWidth: '280px' }}>
                  <Search size={15} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
                  <input
                    type="text"
                    value={resolverInput}
                    onChange={(e) => setResolverInput(e.target.value)}
                    placeholder="Type any test raw string (e.g. 'Karnataka', 'Kannur Music', 'PFPSYGS25', 'Mukkuvan', 'HSE')..."
                    style={{
                      width: '100%',
                      padding: '8px 12px 8px 32px',
                      borderRadius: '6px',
                      border: '1px solid var(--line)',
                      background: 'var(--bg)',
                      color: 'var(--ink)',
                      fontSize: '13px',
                      outline: 'none',
                      boxSizing: 'border-box'
                    }}
                  />
                  {resolverInput && (
                    <button
                      onClick={() => setResolverInput('')}
                      style={{ position: 'absolute', right: '8px', top: '50%', transform: 'translateY(-50%)', background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer' }}
                    >
                      <X size={14} />
                    </button>
                  )}
                </div>

                <div style={{ display: 'flex', gap: '6px', alignItems: 'center' }}>
                  <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 600 }}>Quick:</span>
                  {['Karnataka', 'Five Year Integrated Masters Programme in Music', 'Department of Music', 'Mukkuvan', 'SEBC-Muslim'].map(preset => (
                    <button
                      key={preset}
                      onClick={() => setResolverInput(preset)}
                      style={{
                        padding: '4px 8px',
                        borderRadius: '4px',
                        background: 'var(--bg)',
                        color: 'var(--ink)',
                        border: '1px solid var(--line)',
                        fontSize: '11px',
                        fontWeight: 500,
                        cursor: 'pointer'
                      }}
                    >
                      {preset.length > 20 ? preset.slice(0, 18) + '...' : preset}
                    </button>
                  ))}
                </div>
              </div>

              {resolvedResult && resolverInput.trim() && (
                <div style={{
                  marginTop: '12px',
                  padding: '10px 14px',
                  borderRadius: '8px',
                  background: 'var(--bg)',
                  border: '1px solid var(--line)',
                  display: 'flex',
                  justifyContent: 'space-between',
                  alignItems: 'center',
                  flexWrap: 'wrap',
                  gap: '8px',
                  fontSize: '12.5px'
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                    <span style={{
                      padding: '2px 8px',
                      borderRadius: '10px',
                      fontSize: '10.5px',
                      fontWeight: 700,
                      textTransform: 'uppercase',
                      color: resolvedResult.color,
                      background: resolvedResult.bgColor,
                      border: `1px solid ${resolvedResult.color}`
                    }}>
                      {resolvedResult.entity} Match
                    </span>
                    <strong style={{ color: 'var(--ink)' }}>{resolvedResult.title}</strong>
                    {resolvedResult.code && (
                      <span style={{ fontFamily: 'monospace', fontWeight: 700, padding: '2px 6px', borderRadius: '4px', background: 'var(--panel)', border: '1px solid var(--line)' }}>
                        Code: {resolvedResult.code}
                      </span>
                    )}
                    {resolvedResult.id && resolvedResult.id !== 'N/A' && (
                      <span style={{ fontFamily: 'monospace', fontWeight: 700, padding: '2px 6px', borderRadius: '4px', background: 'var(--panel)', border: '1px solid var(--line)' }}>
                        ID: {resolvedResult.id}
                      </span>
                    )}
                  </div>

                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--muted)', fontSize: '11.5px' }}>
                    <span style={{ color: resolvedResult.color, fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                      <CheckCircle2 size={13} /> {resolvedResult.matchType} ({resolvedResult.confidence})
                    </span>
                    {resolvedResult.tabKey && (
                      <button
                        onClick={() => {
                          setViewMode('tabbed');
                          handleTabChange(resolvedResult.tabKey);
                          setSearchQuery(resolvedResult.title);
                        }}
                        style={{
                          background: 'transparent',
                          border: 'none',
                          color: 'var(--accent)',
                          fontSize: '11.5px',
                          fontWeight: 600,
                          cursor: 'pointer',
                          display: 'inline-flex',
                          alignItems: 'center',
                          gap: '3px'
                        }}
                      >
                        Find in Directory <ExternalLink size={12} />
                      </button>
                    )}
                  </div>
                </div>
              )}
            </div>
          )}
        </div>

        {/* View Mode 1: DUAL SIDE-BY-SIDE VIEW (Programme & College) */}
        {viewMode === 'split' && (
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(480px, 1fr))',
            gap: '16px',
            alignItems: 'start'
          }}>
            {/* Left Pane: Programme Master */}
            <div style={{
              background: 'var(--panel)',
              border: '1px solid var(--line)',
              borderRadius: '10px',
              overflow: 'hidden',
              boxShadow: '0 2px 8px rgba(0,0,0,0.02)',
              display: 'flex',
              flexDirection: 'column'
            }}>
              {/* Header */}
              <div style={{
                padding: '12px 16px',
                background: 'var(--bg)',
                borderBottom: '1px solid var(--line)',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center'
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <div style={{ color: '#059669', display: 'flex' }}><BookOpen size={16} /></div>
                  <h2 style={{ fontSize: '13.5px', fontWeight: 700, margin: 0 }}>
                    Programme Master Directory
                    <span style={{ marginLeft: '8px', fontSize: '11px', background: 'rgba(5, 150, 105, 0.1)', color: '#059669', padding: '1px 6px', borderRadius: '10px' }}>
                      {(lookups.program || []).length}
                    </span>
                  </h2>
                </div>
                <button
                  onClick={() => { handleTabChange('program'); setIsAddFormOpen(true); }}
                  style={{
                    padding: '4px 10px',
                    borderRadius: '5px',
                    background: 'var(--accent)',
                    color: '#fff',
                    border: 'none',
                    fontSize: '11.5px',
                    fontWeight: 600,
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '4px'
                  }}
                >
                  <Plus size={13} /> Add
                </button>
              </div>

              {/* Search */}
              <div style={{ padding: '8px 12px', borderBottom: '1px solid var(--line)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <div style={{ position: 'relative', flex: 1 }}>
                  <Search size={13} style={{ position: 'absolute', left: '8px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
                  <input
                    type="text"
                    value={progSearch}
                    onChange={(e) => { setProgSearch(e.target.value); setProgPage(0); }}
                    placeholder="Search programmes by name, code (PF...), or workflow ID..."
                    style={{ width: '100%', padding: '6px 8px 6px 26px', fontSize: '12px', border: '1px solid var(--line)', borderRadius: '4px', background: 'var(--bg)', color: 'var(--ink)' }}
                  />
                  {progSearch && (
                    <button onClick={() => setProgSearch('')} style={{ position: 'absolute', right: '6px', top: '50%', transform: 'translateY(-50%)', background: 'transparent', border: 'none', cursor: 'pointer' }}>
                      <X size={12} />
                    </button>
                  )}
                </div>
                <span style={{ fontSize: '11px', color: 'var(--muted)' }}>{dualProgList.length} items</span>
              </div>

              {/* Table */}
              <div style={{ overflowX: 'auto', minHeight: '380px' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                  <thead>
                    <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)', color: 'var(--muted)', fontSize: '11px', textTransform: 'uppercase' }}>
                      <th style={{ padding: '8px 10px', width: '30px', textAlign: 'center' }}>#</th>
                      <th style={{ padding: '8px 10px' }}>Reference Name</th>
                      <th style={{ padding: '8px 10px', width: '100px' }}>Code</th>
                      <th style={{ padding: '8px 10px', width: '80px' }}>WF ID</th>
                      <th style={{ padding: '8px 10px', width: '60px', textAlign: 'right' }}>Act</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dualPaginatedProgs.map((item, idx) => {
                      const isEditing = editingTargetTab === 'program' && editingIdx === (lookups.program || []).indexOf(item);
                      return (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--line)', background: isEditing ? 'rgba(217, 119, 6, 0.08)' : 'transparent' }}>
                          <td style={{ padding: '8px 10px', textAlign: 'center', color: 'var(--muted)', fontFamily: 'monospace', fontSize: '11px' }}>
                            {progPage * dualPageSize + idx + 1}
                          </td>
                          <td style={{ padding: '8px 10px', fontWeight: 600 }}>
                            {isEditing ? (
                              <input type="text" value={editProgRef} onChange={(e) => setEditProgRef(e.target.value)} style={{ width: '100%', padding: '4px', fontSize: '11.5px' }} />
                            ) : (
                              <div>
                                <span>{item.pgmName}</span>
                                <div style={{ fontSize: '10.5px', color: 'var(--muted)', fontWeight: 400, marginTop: '2px' }}>{item.name}</div>
                              </div>
                            )}
                          </td>
                          <td style={{ padding: '8px 10px' }}>
                            {isEditing ? (
                              <input type="text" value={editProgCode} onChange={(e) => setEditProgCode(e.target.value)} style={{ width: '100%', padding: '4px', fontSize: '11.5px', fontFamily: 'monospace' }} />
                            ) : (
                              <button
                                onClick={() => handleCopy(item.code, `prog_c_${idx}`)}
                                style={{
                                  padding: '2px 6px',
                                  borderRadius: '4px',
                                  background: 'rgba(5, 150, 105, 0.1)',
                                  color: '#059669',
                                  border: '1px solid rgba(5, 150, 105, 0.25)',
                                  fontFamily: 'monospace',
                                  fontSize: '11px',
                                  fontWeight: 700,
                                  cursor: 'pointer',
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '4px'
                                }}
                              >
                                {copiedKey === `prog_c_${idx}` ? <CheckCheck size={12} /> : <Copy size={11} opacity={0.6} />}
                                {item.code}
                              </button>
                            )}
                          </td>
                          <td style={{ padding: '8px 10px' }}>
                            {isEditing ? (
                              <input type="text" value={editProgWorkflow} onChange={(e) => setEditProgWorkflow(e.target.value)} style={{ width: '100%', padding: '4px', fontSize: '11.5px', fontFamily: 'monospace' }} />
                            ) : (
                              <button
                                onClick={() => handleCopy(item.workflowId, `prog_w_${idx}`)}
                                style={{
                                  padding: '2px 6px',
                                  borderRadius: '4px',
                                  background: 'var(--bg)',
                                  color: 'var(--ink)',
                                  border: '1px solid var(--line)',
                                  fontFamily: 'monospace',
                                  fontSize: '11px',
                                  fontWeight: 700,
                                  cursor: 'pointer',
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '4px'
                                }}
                              >
                                {copiedKey === `prog_w_${idx}` ? <CheckCheck size={12} /> : <Copy size={11} opacity={0.6} />}
                                {item.workflowId}
                              </button>
                            )}
                          </td>
                          <td style={{ padding: '8px 10px', textAlign: 'right' }}>
                            <div style={{ display: 'flex', gap: '4px', justifyContent: 'flex-end' }}>
                              {isEditing ? (
                                <>
                                  <button onClick={() => handleSaveEdit('program')} style={{ padding: '4px', background: '#059669', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}><Check size={12} /></button>
                                  <button onClick={() => { setEditingIdx(null); setEditingTargetTab(null); }} style={{ padding: '4px', background: 'var(--bg)', color: 'var(--muted)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer' }}><X size={12} /></button>
                                </>
                              ) : (
                                <>
                                  <button onClick={() => handleStartEdit(item, 'program')} style={{ padding: '4px', background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer' }}><Edit2 size={12} /></button>
                                  <button onClick={() => handleDelete(item, 'program')} style={{ padding: '4px', background: 'transparent', border: 'none', color: '#dc2626', cursor: 'pointer' }}><Trash2 size={12} /></button>
                                </>
                              )}
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Pagination */}
              {dualProgList.length > dualPageSize && (
                <div style={{ padding: '8px 12px', borderTop: '1px solid var(--line)', background: 'var(--bg)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
                  <span>Page {progPage + 1} of {Math.ceil(dualProgList.length / dualPageSize)}</span>
                  <div style={{ display: 'flex', gap: '4px' }}>
                    <button onClick={() => setProgPage(p => Math.max(0, p - 1))} disabled={progPage === 0} style={{ padding: '3px 8px', fontSize: '11px', borderRadius: '4px', background: 'var(--panel)', color: 'var(--ink)', border: '1px solid var(--line)' }}>Prev</button>
                    <button onClick={() => setProgPage(p => Math.min(Math.ceil(dualProgList.length / dualPageSize) - 1, p + 1))} disabled={progPage >= Math.ceil(dualProgList.length / dualPageSize) - 1} style={{ padding: '3px 8px', fontSize: '11px', borderRadius: '4px', background: 'var(--panel)', color: 'var(--ink)', border: '1px solid var(--line)' }}>Next</button>
                  </div>
                </div>
              )}
            </div>

            {/* Right Pane: College Master */}
            <div style={{
              background: 'var(--panel)',
              border: '1px solid var(--line)',
              borderRadius: '10px',
              overflow: 'hidden',
              boxShadow: '0 2px 8px rgba(0,0,0,0.02)',
              display: 'flex',
              flexDirection: 'column'
            }}>
              {/* Header */}
              <div style={{
                padding: '12px 16px',
                background: 'var(--bg)',
                borderBottom: '1px solid var(--line)',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center'
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <div style={{ color: '#4f46e5', display: 'flex' }}><Building2 size={16} /></div>
                  <h2 style={{ fontSize: '13.5px', fontWeight: 700, margin: 0 }}>
                    College Master Directory
                    <span style={{ marginLeft: '8px', fontSize: '11px', background: 'rgba(79, 70, 229, 0.1)', color: '#4f46e5', padding: '1px 6px', borderRadius: '10px' }}>
                      {(lookups.college || []).length}
                    </span>
                  </h2>
                </div>
                <button
                  onClick={() => { handleTabChange('college'); setIsAddFormOpen(true); }}
                  style={{
                    padding: '4px 10px',
                    borderRadius: '5px',
                    background: '#4f46e5',
                    color: '#fff',
                    border: 'none',
                    fontSize: '11.5px',
                    fontWeight: 600,
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '4px'
                  }}
                >
                  <Plus size={13} /> Add
                </button>
              </div>

              {/* Search */}
              <div style={{ padding: '8px 12px', borderBottom: '1px solid var(--line)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <div style={{ position: 'relative', flex: 1 }}>
                  <Search size={13} style={{ position: 'absolute', left: '8px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
                  <input
                    type="text"
                    value={clgSearch}
                    onChange={(e) => { setClgSearch(e.target.value); setClgPage(0); }}
                    placeholder="Search colleges by name, ID (e.g. 127, 435), or campus..."
                    style={{ width: '100%', padding: '6px 8px 6px 26px', fontSize: '12px', border: '1px solid var(--line)', borderRadius: '4px', background: 'var(--bg)', color: 'var(--ink)' }}
                  />
                  {clgSearch && (
                    <button onClick={() => setClgSearch('')} style={{ position: 'absolute', right: '6px', top: '50%', transform: 'translateY(-50%)', background: 'transparent', border: 'none', cursor: 'pointer' }}>
                      <X size={12} />
                    </button>
                  )}
                </div>
                <span style={{ fontSize: '11px', color: 'var(--muted)' }}>{dualClgList.length} items</span>
              </div>

              {/* Table */}
              <div style={{ overflowX: 'auto', minHeight: '380px' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                  <thead>
                    <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)', color: 'var(--muted)', fontSize: '11px', textTransform: 'uppercase' }}>
                      <th style={{ padding: '8px 10px', width: '30px', textAlign: 'center' }}>#</th>
                      <th style={{ padding: '8px 10px', width: '80px' }}>ID</th>
                      <th style={{ padding: '8px 10px', width: '80px' }}>Code</th>
                      <th style={{ padding: '8px 10px' }}>College / Campus Name</th>
                      <th style={{ padding: '8px 10px', width: '60px', textAlign: 'right' }}>Act</th>
                    </tr>
                  </thead>
                  <tbody>
                    {dualPaginatedColleges.map((item, idx) => {
                      const isEditing = editingTargetTab === 'college' && editingIdx === (lookups.college || []).indexOf(item);
                      return (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--line)', background: isEditing ? 'rgba(217, 119, 6, 0.08)' : 'transparent' }}>
                          <td style={{ padding: '8px 10px', textAlign: 'center', color: 'var(--muted)', fontFamily: 'monospace', fontSize: '11px' }}>
                            {clgPage * dualPageSize + idx + 1}
                          </td>
                          <td style={{ padding: '8px 10px' }}>
                            {isEditing ? (
                              <input type="text" value={editClgId} onChange={(e) => setEditClgId(e.target.value)} style={{ width: '100%', padding: '4px', fontSize: '11.5px', fontFamily: 'monospace' }} />
                            ) : (
                              <button
                                onClick={() => handleCopy(item.id, `clg_id_${idx}`)}
                                style={{
                                  padding: '2px 6px',
                                  borderRadius: '4px',
                                  background: 'rgba(79, 70, 229, 0.1)',
                                  color: '#4f46e5',
                                  border: '1px solid rgba(79, 70, 229, 0.25)',
                                  fontFamily: 'monospace',
                                  fontSize: '11px',
                                  fontWeight: 700,
                                  cursor: 'pointer',
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '4px'
                                }}
                              >
                                {copiedKey === `clg_id_${idx}` ? <CheckCheck size={12} /> : <Copy size={11} opacity={0.6} />}
                                {item.id}
                              </button>
                            )}
                          </td>
                          <td style={{ padding: '8px 10px', fontFamily: 'monospace', color: 'var(--muted)' }}>
                            {isEditing ? (
                              <input type="text" value={editClgCode} onChange={(e) => setEditClgCode(e.target.value)} style={{ width: '100%', padding: '4px', fontSize: '11.5px', fontFamily: 'monospace' }} />
                            ) : (
                              item.code
                            )}
                          </td>
                          <td style={{ padding: '8px 10px', fontWeight: 600 }}>
                            {isEditing ? (
                              <input type="text" value={editClgName} onChange={(e) => setEditClgName(e.target.value)} style={{ width: '100%', padding: '4px', fontSize: '11.5px' }} />
                            ) : (
                              <div>
                                <span>{item.name}</span>
                                {item.clgRef && item.clgRef !== item.name && (
                                  <div style={{ fontSize: '10.5px', color: 'var(--muted)', fontWeight: 400, marginTop: '2px' }}>{item.clgRef}</div>
                                )}
                              </div>
                            )}
                          </td>
                          <td style={{ padding: '8px 10px', textAlign: 'right' }}>
                            <div style={{ display: 'flex', gap: '4px', justifyContent: 'flex-end' }}>
                              {isEditing ? (
                                <>
                                  <button onClick={() => handleSaveEdit('college')} style={{ padding: '4px', background: '#4f46e5', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}><Check size={12} /></button>
                                  <button onClick={() => { setEditingIdx(null); setEditingTargetTab(null); }} style={{ padding: '4px', background: 'var(--bg)', color: 'var(--muted)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer' }}><X size={12} /></button>
                                </>
                              ) : (
                                <>
                                  <button onClick={() => handleStartEdit(item, 'college')} style={{ padding: '4px', background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer' }}><Edit2 size={12} /></button>
                                  <button onClick={() => handleDelete(item, 'college')} style={{ padding: '4px', background: 'transparent', border: 'none', color: '#dc2626', cursor: 'pointer' }}><Trash2 size={12} /></button>
                                </>
                              )}
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Pagination */}
              {dualClgList.length > dualPageSize && (
                <div style={{ padding: '8px 12px', borderTop: '1px solid var(--line)', background: 'var(--bg)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
                  <span>Page {clgPage + 1} of {Math.ceil(dualClgList.length / dualPageSize)}</span>
                  <div style={{ display: 'flex', gap: '4px' }}>
                    <button onClick={() => setClgPage(p => Math.max(0, p - 1))} disabled={clgPage === 0} style={{ padding: '3px 8px', fontSize: '11px', borderRadius: '4px', background: 'var(--panel)', color: 'var(--ink)', border: '1px solid var(--line)' }}>Prev</button>
                    <button onClick={() => setClgPage(p => Math.min(Math.ceil(dualClgList.length / dualPageSize) - 1, p + 1))} disabled={clgPage >= Math.ceil(dualClgList.length / dualPageSize) - 1} style={{ padding: '3px 8px', fontSize: '11px', borderRadius: '4px', background: 'var(--panel)', color: 'var(--ink)', border: '1px solid var(--line)' }}>Next</button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* View Mode 2: TABBED MASTER DIRECTORY (All 9 Categories) */}
        {viewMode === 'tabbed' && (
          <div style={{
            background: 'var(--panel)',
            border: '1px solid var(--line)',
            borderRadius: '10px',
            overflow: 'hidden',
            boxShadow: '0 2px 8px rgba(0,0,0,0.02)',
            display: 'flex',
            flexDirection: 'column'
          }}>
            {/* Tab Ribbon */}
            <div style={{
              display: 'flex',
              gap: '4px',
              padding: '10px 16px 0 16px',
              background: 'var(--bg)',
              borderBottom: '1px solid var(--line)',
              overflowX: 'auto'
            }}>
              {LOOKUP_TAB_CONFIG.map(tab => {
                const Icon = TAB_ICONS[tab.key] || Tag;
                const count = (lookups && Array.isArray(lookups[tab.key])) ? lookups[tab.key].length : 0;
                const isActive = activeTab === tab.key;

                return (
                  <button
                    key={tab.key}
                    onClick={() => handleTabChange(tab.key)}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px',
                      padding: '8px 14px',
                      borderTopLeftRadius: '6px',
                      borderTopRightRadius: '6px',
                      border: 'none',
                      borderBottom: isActive ? '2.5px solid var(--accent)' : '2.5px solid transparent',
                      background: isActive ? 'var(--panel)' : 'transparent',
                      color: isActive ? 'var(--accent)' : 'var(--muted)',
                      fontWeight: isActive ? 700 : 500,
                      fontSize: '12px',
                      cursor: 'pointer',
                      whiteSpace: 'nowrap'
                    }}
                  >
                    <Icon size={14} />
                    <span>{tab.label}</span>
                    <span style={{
                      fontSize: '10px',
                      fontWeight: 700,
                      padding: '1px 6px',
                      borderRadius: '10px',
                      background: isActive ? 'var(--accent-soft)' : 'rgba(0,0,0,0.06)',
                      color: isActive ? 'var(--accent)' : 'var(--muted)'
                    }}>
                      {count}
                    </span>
                  </button>
                );
              })}
            </div>

            {/* Subheader: Search, Filters, Add Toggle */}
            <div style={{
              padding: '12px 16px',
              borderBottom: '1px solid var(--line)',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              flexWrap: 'wrap',
              gap: '10px'
            }}>
              <div>
                <strong style={{ fontSize: '13px' }}>{currentTabConfig.label} Directory</strong>
                <span style={{ fontSize: '11.5px', color: 'var(--muted)', marginLeft: '6px' }}>
                  ({currentTabConfig.description}) — Showing {filteredList.length} of {currentList.length}
                </span>
                {selectedRowKeys.size > 0 && (
                  <span style={{ marginLeft: '8px', fontSize: '11px', color: 'var(--accent)', fontWeight: 700 }}>
                    ({selectedRowKeys.size} selected)
                  </span>
                )}
              </div>

              <div style={{ display: 'flex', gap: '8px', alignItems: 'center', flexWrap: 'wrap' }}>
                <div style={{ position: 'relative', width: '220px' }}>
                  <Search size={13} style={{ position: 'absolute', left: '8px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }}
                    placeholder={`Search ${currentTabConfig.label.toLowerCase()}...`}
                    style={{ width: '100%', padding: '6px 8px 6px 26px', fontSize: '12px', border: '1px solid var(--line)', borderRadius: '4px', background: 'var(--bg)', color: 'var(--ink)' }}
                  />
                  {searchQuery && (
                    <button onClick={() => setSearchQuery('')} style={{ position: 'absolute', right: '6px', top: '50%', transform: 'translateY(-50%)', background: 'transparent', border: 'none', cursor: 'pointer' }}>
                      <X size={12} />
                    </button>
                  )}
                </div>

                <select
                  value={pageSize}
                  onChange={(e) => { setPageSize(Number(e.target.value)); setPage(0); }}
                  style={{ padding: '6px 8px', fontSize: '12px', border: '1px solid var(--line)', borderRadius: '4px', background: 'var(--bg)', color: 'var(--ink)' }}
                >
                  <option value={15}>15 / page</option>
                  <option value={25}>25 / page</option>
                  <option value={50}>50 / page</option>
                  <option value={100}>100 / page</option>
                  <option value={0}>Show All</option>
                </select>

                <button
                  onClick={() => setIsAddFormOpen(!isAddFormOpen)}
                  style={{
                    padding: '6px 12px',
                    borderRadius: '5px',
                    background: 'var(--accent)',
                    color: '#fff',
                    border: 'none',
                    fontSize: '12px',
                    fontWeight: 600,
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '4px'
                  }}
                >
                  <Plus size={13} /> {isAddFormOpen ? 'Close Form' : 'Add Entry'}
                </button>
              </div>
            </div>

            {/* Quick Filter Chips */}
            {(activeTab === 'program' || activeTab === 'college' || activeTab === 'caste') && (
              <div style={{ padding: '8px 16px', background: 'var(--bg)', borderBottom: '1px solid var(--line)', display: 'flex', gap: '6px', alignItems: 'center' }}>
                <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 600 }}>Filters:</span>
                {activeTab === 'program' && [
                  { id: 'all', label: 'All Programmes' },
                  { id: 'fyimp', label: 'Integrated Masters (FYIMP)' },
                  { id: 'regular', label: 'Regular PG / Other' }
                ].map(chip => (
                  <button
                    key={chip.id}
                    onClick={() => { setFilterChip(chip.id); setPage(0); }}
                    style={{
                      padding: '3px 8px',
                      borderRadius: '12px',
                      fontSize: '11px',
                      fontWeight: 500,
                      border: '1px solid var(--line)',
                      background: filterChip === chip.id ? 'var(--accent)' : 'var(--panel)',
                      color: filterChip === chip.id ? '#fff' : 'var(--ink)',
                      cursor: 'pointer'
                    }}
                  >
                    {chip.label}
                  </button>
                ))}
                {activeTab === 'college' && [
                  { id: 'all', label: 'All Colleges' },
                  { id: 'dept', label: 'University Departments' },
                  { id: 'campus', label: 'Campuses & Centers' }
                ].map(chip => (
                  <button
                    key={chip.id}
                    onClick={() => { setFilterChip(chip.id); setPage(0); }}
                    style={{
                      padding: '3px 8px',
                      borderRadius: '12px',
                      fontSize: '11px',
                      fontWeight: 500,
                      border: '1px solid var(--line)',
                      background: filterChip === chip.id ? 'var(--accent)' : 'var(--panel)',
                      color: filterChip === chip.id ? '#fff' : 'var(--ink)',
                      cursor: 'pointer'
                    }}
                  >
                    {chip.label}
                  </button>
                ))}
              </div>
            )}

            {/* Expandable Add Entry Form */}
            {isAddFormOpen && (
              <div style={{ padding: '14px 18px', background: 'var(--bg)', borderBottom: '1px solid var(--line)' }}>
                <form onSubmit={handleAdd}>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '10px', marginBottom: '10px' }}>
                    {activeTab === 'program' ? (
                      <>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>Reference Name *</label>
                          <input type="text" value={newProgRef} onChange={(e) => setNewProgRef(e.target.value)} placeholder="e.g. Five Year Integrated..." style={{ width: '100%', padding: '6px', fontSize: '12px' }} />
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>ProgramCode *</label>
                          <input type="text" value={newProgCode} onChange={(e) => setNewProgCode(e.target.value)} placeholder="e.g. PFPSYGS25" style={{ width: '100%', padding: '6px', fontSize: '12px', fontFamily: 'monospace' }} />
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>workFlowId *</label>
                          <input type="text" value={newProgWorkflow} onChange={(e) => setNewProgWorkflow(e.target.value)} placeholder="e.g. 151" style={{ width: '100%', padding: '6px', fontSize: '12px', fontFamily: 'monospace' }} />
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>Full Program Name *</label>
                          <input type="text" value={newProgName} onChange={(e) => setNewProgName(e.target.value)} placeholder="e.g. (with Credits)-Regular-FYIMP..." style={{ width: '100%', padding: '6px', fontSize: '12px' }} />
                        </div>
                      </>
                    ) : activeTab === 'college' ? (
                      <>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>CollegeId *</label>
                          <input type="text" value={newClgId} onChange={(e) => setNewClgId(e.target.value)} placeholder="e.g. 435" style={{ width: '100%', padding: '6px', fontSize: '12px', fontFamily: 'monospace' }} />
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>CollegeCode</label>
                          <input type="text" value={newClgCode} onChange={(e) => setNewClgCode(e.target.value)} placeholder="e.g. 435" style={{ width: '100%', padding: '6px', fontSize: '12px', fontFamily: 'monospace' }} />
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>College Name *</label>
                          <input type="text" value={newClgName} onChange={(e) => setNewClgName(e.target.value)} placeholder="e.g. Department of Anthropology" style={{ width: '100%', padding: '6px', fontSize: '12px' }} />
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>CLG Name (Campus)</label>
                          <input type="text" value={newClgRef} onChange={(e) => setNewClgRef(e.target.value)} placeholder="e.g. Dr. Janaki Ammal Campus, Palayad" style={{ width: '100%', padding: '6px', fontSize: '12px' }} />
                        </div>
                      </>
                    ) : (
                      <>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>{currentTabConfig.nameLabel} *</label>
                          <input type="text" value={newName} onChange={(e) => setNewName(e.target.value)} placeholder={`Enter ${currentTabConfig.nameLabel.toLowerCase()}...`} style={{ width: '100%', padding: '6px', fontSize: '12px' }} />
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 600, color: 'var(--muted)', marginBottom: '3px' }}>{currentTabConfig.idLabel} *</label>
                          <input type="text" value={newId} onChange={(e) => setNewId(e.target.value)} placeholder={`Enter ${currentTabConfig.idLabel.toLowerCase()}...`} style={{ width: '100%', padding: '6px', fontSize: '12px', fontFamily: 'monospace' }} />
                        </div>
                      </>
                    )}
                  </div>
                  <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
                    <button type="button" onClick={() => setIsAddFormOpen(false)} style={{ padding: '6px 12px', fontSize: '12px', background: 'var(--panel)', color: 'var(--ink)', border: '1px solid var(--line)', borderRadius: '4px' }}>Cancel</button>
                    <button type="submit" style={{ padding: '6px 14px', fontSize: '12px', background: 'var(--accent)', color: '#fff', border: 'none', borderRadius: '4px', fontWeight: 600 }}>Save to {currentTabConfig.label}</button>
                  </div>
                </form>
              </div>
            )}

            {/* Bulk Selection Bar */}
            {selectedRowKeys.size > 0 && (
              <div style={{ padding: '8px 16px', background: 'var(--accent-soft)', borderBottom: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '12px' }}>
                <span style={{ fontWeight: 600, color: 'var(--accent)' }}>{selectedRowKeys.size} item(s) selected</span>
                <div style={{ display: 'flex', gap: '6px' }}>
                  <button onClick={handleBulkDelete} style={{ padding: '4px 10px', fontSize: '11.5px', background: '#dc2626', color: '#fff', border: 'none', borderRadius: '4px' }}>Delete Selected</button>
                  <button onClick={() => setSelectedRowKeys(new Set())} style={{ padding: '4px 8px', fontSize: '11.5px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px' }}>Clear</button>
                </div>
              </div>
            )}

            {/* Table */}
            <div style={{ overflowX: 'auto', minHeight: '380px' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)', color: 'var(--muted)', fontSize: '11px', textTransform: 'uppercase' }}>
                    <th style={{ padding: '8px 10px', width: '30px', textAlign: 'center' }}>
                      <input
                        type="checkbox"
                        checked={paginatedList.length > 0 && paginatedList.every((item, idx) => selectedRowKeys.has(`${activeTab}_${(pageSize === 0 ? 0 : page * pageSize) + idx}_${item.id || item.code || item.name}`))}
                        onChange={(e) => {
                          const next = new Set(selectedRowKeys);
                          paginatedList.forEach((item, idx) => {
                            const key = `${activeTab}_${(pageSize === 0 ? 0 : page * pageSize) + idx}_${item.id || item.code || item.name}`;
                            if (e.target.checked) next.add(key);
                            else next.delete(key);
                          });
                          setSelectedRowKeys(next);
                        }}
                      />
                    </th>
                    <th style={{ padding: '8px 10px', width: '30px', textAlign: 'center' }}>#</th>

                    {activeTab === 'program' ? (
                      <>
                        <th onClick={() => handleSort('pgmName')} style={{ padding: '8px 10px', cursor: 'pointer' }}>
                          Programme Name (Reference) {sortConfig.column === 'pgmName' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                        <th onClick={() => handleSort('code')} style={{ padding: '8px 10px', width: '130px', cursor: 'pointer' }}>
                          ProgramCode {sortConfig.column === 'code' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                        <th onClick={() => handleSort('workflowId')} style={{ padding: '8px 10px', width: '110px', cursor: 'pointer' }}>
                          workFlowId {sortConfig.column === 'workflowId' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                        <th onClick={() => handleSort('name')} style={{ padding: '8px 10px', cursor: 'pointer' }}>
                          Full Program Name {sortConfig.column === 'name' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                      </>
                    ) : activeTab === 'college' ? (
                      <>
                        <th onClick={() => handleSort('id')} style={{ padding: '8px 10px', width: '110px', cursor: 'pointer' }}>
                          CollegeId {sortConfig.column === 'id' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                        <th onClick={() => handleSort('code')} style={{ padding: '8px 10px', width: '110px', cursor: 'pointer' }}>
                          CollegeCode {sortConfig.column === 'code' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                        <th onClick={() => handleSort('name')} style={{ padding: '8px 10px', cursor: 'pointer' }}>
                          College Name {sortConfig.column === 'name' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                        <th onClick={() => handleSort('clgRef')} style={{ padding: '8px 10px', cursor: 'pointer' }}>
                          CLG Name (Campus) {sortConfig.column === 'clgRef' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                      </>
                    ) : (
                      <>
                        <th onClick={() => handleSort('name')} style={{ padding: '8px 10px', cursor: 'pointer' }}>
                          {currentTabConfig.nameHeader} {sortConfig.column === 'name' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                        <th onClick={() => handleSort('id')} style={{ padding: '8px 10px', width: '160px', cursor: 'pointer' }}>
                          {currentTabConfig.idHeader} {sortConfig.column === 'id' ? (sortConfig.direction === 'asc' ? '↑' : '↓') : '↕'}
                        </th>
                      </>
                    )}

                    <th style={{ padding: '8px 10px', width: '70px', textAlign: 'right' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {paginatedList.map((item, idx) => {
                    const rowNumber = (pageSize === 0 ? 0 : page * pageSize) + idx + 1;
                    const isEditing = editingIdx === currentList.indexOf(item);
                    const rowKey = `${activeTab}_${rowNumber - 1}_${item.id || item.code || item.name}`;
                    const isSelected = selectedRowKeys.has(rowKey);

                    return (
                      <tr key={idx} style={{ borderBottom: '1px solid var(--line)', background: isEditing ? 'rgba(217, 119, 6, 0.08)' : isSelected ? 'var(--accent-soft)' : 'transparent' }}>
                        <td style={{ padding: '8px 10px', textAlign: 'center' }}>
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={(e) => {
                              const next = new Set(selectedRowKeys);
                              if (e.target.checked) next.add(rowKey);
                              else next.delete(rowKey);
                              setSelectedRowKeys(next);
                            }}
                          />
                        </td>
                        <td style={{ padding: '8px 10px', textAlign: 'center', color: 'var(--muted)', fontFamily: 'monospace', fontSize: '11px' }}>
                          {rowNumber}
                        </td>

                        {activeTab === 'program' ? (
                          <>
                            <td style={{ padding: '8px 10px', fontWeight: 600 }}>
                              {isEditing ? <input type="text" value={editProgRef} onChange={(e) => setEditProgRef(e.target.value)} style={{ width: '100%' }} /> : item.pgmName}
                            </td>
                            <td style={{ padding: '8px 10px' }}>
                              {isEditing ? <input type="text" value={editProgCode} onChange={(e) => setEditProgCode(e.target.value)} style={{ width: '100%', fontFamily: 'monospace' }} /> : (
                                <button onClick={() => handleCopy(item.code, `tab_c_${idx}`)} style={{ padding: '2px 6px', borderRadius: '4px', background: 'rgba(5, 150, 105, 0.1)', color: '#059669', border: '1px solid rgba(5, 150, 105, 0.25)', fontFamily: 'monospace', fontSize: '11px', fontWeight: 700, cursor: 'pointer' }}>
                                  {copiedKey === `tab_c_${idx}` ? 'Copied!' : item.code}
                                </button>
                              )}
                            </td>
                            <td style={{ padding: '8px 10px' }}>
                              {isEditing ? <input type="text" value={editProgWorkflow} onChange={(e) => setEditProgWorkflow(e.target.value)} style={{ width: '100%', fontFamily: 'monospace' }} /> : (
                                <button onClick={() => handleCopy(item.workflowId, `tab_w_${idx}`)} style={{ padding: '2px 6px', borderRadius: '4px', background: 'var(--bg)', color: 'var(--ink)', border: '1px solid var(--line)', fontFamily: 'monospace', fontSize: '11px', fontWeight: 700, cursor: 'pointer' }}>
                                  {copiedKey === `tab_w_${idx}` ? 'Copied!' : item.workflowId}
                                </button>
                              )}
                            </td>
                            <td style={{ padding: '8px 10px', fontSize: '11.5px', color: 'var(--muted)' }}>
                              {isEditing ? <input type="text" value={editProgName} onChange={(e) => setEditProgName(e.target.value)} style={{ width: '100%' }} /> : item.name}
                            </td>
                          </>
                        ) : activeTab === 'college' ? (
                          <>
                            <td style={{ padding: '8px 10px' }}>
                              {isEditing ? <input type="text" value={editClgId} onChange={(e) => setEditClgId(e.target.value)} style={{ width: '100%', fontFamily: 'monospace' }} /> : (
                                <button onClick={() => handleCopy(item.id, `tab_ci_${idx}`)} style={{ padding: '2px 6px', borderRadius: '4px', background: 'rgba(79, 70, 229, 0.1)', color: '#4f46e5', border: '1px solid rgba(79, 70, 229, 0.25)', fontFamily: 'monospace', fontSize: '11px', fontWeight: 700, cursor: 'pointer' }}>
                                  {copiedKey === `tab_ci_${idx}` ? 'Copied!' : item.id}
                                </button>
                              )}
                            </td>
                            <td style={{ padding: '8px 10px', fontFamily: 'monospace', color: 'var(--muted)' }}>
                              {isEditing ? <input type="text" value={editClgCode} onChange={(e) => setEditClgCode(e.target.value)} style={{ width: '100%', fontFamily: 'monospace' }} /> : item.code}
                            </td>
                            <td style={{ padding: '8px 10px', fontWeight: 600 }}>
                              {isEditing ? <input type="text" value={editClgName} onChange={(e) => setEditClgName(e.target.value)} style={{ width: '100%' }} /> : item.name}
                            </td>
                            <td style={{ padding: '8px 10px', fontSize: '11.5px', color: 'var(--muted)' }}>
                              {isEditing ? <input type="text" value={editClgRef} onChange={(e) => setEditClgRef(e.target.value)} style={{ width: '100%' }} /> : item.clgRef}
                            </td>
                          </>
                        ) : (
                          <>
                            <td style={{ padding: '8px 10px', fontWeight: 600 }}>
                              {isEditing ? <input type="text" value={editName} onChange={(e) => setEditName(e.target.value)} style={{ width: '100%' }} /> : item.name}
                            </td>
                            <td style={{ padding: '8px 10px' }}>
                              {isEditing ? <input type="text" value={editId} onChange={(e) => setEditId(e.target.value)} style={{ width: '100%', fontFamily: 'monospace' }} /> : (
                                <button onClick={() => handleCopy(item.id, `tab_id_${idx}`)} style={{ padding: '2px 6px', borderRadius: '4px', background: 'var(--bg)', color: 'var(--ink)', border: '1px solid var(--line)', fontFamily: 'monospace', fontSize: '11px', fontWeight: 700, cursor: 'pointer' }}>
                                  {copiedKey === `tab_id_${idx}` ? 'Copied!' : item.id}
                                </button>
                              )}
                            </td>
                          </>
                        )}

                        <td style={{ padding: '8px 10px', textAlign: 'right' }}>
                          <div style={{ display: 'flex', gap: '4px', justifyContent: 'flex-end' }}>
                            {isEditing ? (
                              <>
                                <button onClick={() => handleSaveEdit(activeTab)} style={{ padding: '4px', background: 'var(--accent)', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer' }}><Check size={12} /></button>
                                <button onClick={() => setEditingIdx(null)} style={{ padding: '4px', background: 'var(--bg)', color: 'var(--muted)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer' }}><X size={12} /></button>
                              </>
                            ) : (
                              <>
                                <button onClick={() => handleStartEdit(item, activeTab)} style={{ padding: '4px', background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer' }}><Edit2 size={12} /></button>
                                <button onClick={() => handleDelete(item, activeTab)} style={{ padding: '4px', background: 'transparent', border: 'none', color: '#dc2626', cursor: 'pointer' }}><Trash2 size={12} /></button>
                              </>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {/* Pagination */}
            {pageSize > 0 && totalPages > 1 && (
              <div style={{ padding: '10px 16px', borderTop: '1px solid var(--line)', background: 'var(--bg)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '12px', color: 'var(--muted)' }}>
                <span>Page {page + 1} of {totalPages} ({filteredList.length} total entries)</span>
                <div style={{ display: 'flex', gap: '4px' }}>
                  <button onClick={() => setPage(p => Math.max(0, p - 1))} disabled={page === 0} style={{ padding: '4px 10px', fontSize: '11.5px', borderRadius: '4px', background: 'var(--panel)', color: 'var(--ink)', border: '1px solid var(--line)' }}>Prev</button>
                  <button onClick={() => setPage(p => Math.min(totalPages - 1, p + 1))} disabled={page >= totalPages - 1} style={{ padding: '4px 10px', fontSize: '11.5px', borderRadius: '4px', background: 'var(--panel)', color: 'var(--ink)', border: '1px solid var(--line)' }}>Next</button>
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      {/* Reset Confirmation Modal */}
      {showResetConfirmModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0, 0, 0, 0.65)',
          backdropFilter: 'blur(4px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          padding: '20px'
        }}>
          <div style={{
            background: 'var(--panel)',
            border: '1px solid var(--line)',
            borderRadius: '12px',
            maxWidth: '440px',
            width: '100%',
            padding: '24px',
            boxShadow: '0 20px 50px rgba(0, 0, 0, 0.3)',
            display: 'flex',
            flexDirection: 'column',
            gap: '14px'
          }}>
            <div style={{ display: 'flex', gap: '12px', alignItems: 'center' }}>
              <div style={{ padding: '8px', borderRadius: '8px', background: 'rgba(217, 119, 6, 0.12)', color: '#d97706', display: 'flex' }}>
                <ShieldAlert size={22} />
              </div>
              <div>
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>Reset Master Definitions?</h3>
                <p style={{ margin: '2px 0 0 0', fontSize: '12px', color: 'var(--muted)' }}>Restore official university reference defaults</p>
              </div>
            </div>

            <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--muted)', lineHeight: '1.5' }}>
              This will reset all 9 master registries (Programmes, Colleges, Castes, Categories, Religions, States, Countries, Boards, and Streams) back to official Kannur University default masters. Any custom unexported records will be overwritten.
            </p>

            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', borderTop: '1px solid var(--line)', paddingTop: '12px' }}>
              <button
                onClick={() => setShowResetConfirmModal(false)}
                style={{ padding: '8px 14px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--panel)', color: 'var(--ink)', fontSize: '12.5px', fontWeight: 600, cursor: 'pointer' }}
              >
                Cancel
              </button>
              <button
                onClick={handleConfirmReset}
                style={{ padding: '8px 16px', borderRadius: '6px', border: 'none', background: '#d97706', color: '#fff', fontSize: '12.5px', fontWeight: 600, cursor: 'pointer' }}
              >
                Yes, Reset to University Defaults
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Guide Documentation Modal */}
      {showGuideModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0, 0, 0, 0.65)',
          backdropFilter: 'blur(4px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 1000,
          padding: '20px'
        }}>
          <div style={{
            background: 'var(--panel)',
            border: '1px solid var(--line)',
            borderRadius: '12px',
            maxWidth: '680px',
            width: '100%',
            maxHeight: '90vh',
            display: 'flex',
            flexDirection: 'column',
            overflow: 'hidden',
            boxShadow: '0 20px 50px rgba(0, 0, 0, 0.3)'
          }}>
            <div style={{
              padding: '16px 20px',
              borderBottom: '1px solid var(--line)',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
              background: 'var(--bg)'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ color: 'var(--accent)', display: 'flex' }}><HelpCircle size={18} /></span>
                <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700 }}>Programme &amp; College Master Definitions Guide</h3>
              </div>
              <button
                onClick={() => setShowGuideModal(false)}
                style={{ background: 'transparent', border: 'none', color: 'var(--muted)', cursor: 'pointer', padding: '2px' }}
              >
                <X size={16} />
              </button>
            </div>

            <div style={{ padding: '20px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '16px', fontSize: '12.5px', lineHeight: '1.6', color: 'var(--ink)' }}>
              <div>
                <strong style={{ fontSize: '13px', display: 'block', marginBottom: '4px' }}>1. What does this tool do?</strong>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  This tool manages the master lookup definitions that power Kannur University admission and examination workflows. It resolves fuzzy and non-standard strings into official university IDs for:
                </p>
                <ul style={{ margin: '6px 0 0 0', paddingLeft: '20px', color: 'var(--muted)' }}>
                  <li><strong>Programmes</strong>: Reference Name ↔ ProgramCode ↔ workFlowId ↔ Full Name</li>
                  <li><strong>Colleges</strong>: CollegeId ↔ CollegeCode ↔ College Name ↔ Campus / CLG Name</li>
                  <li><strong>Social &amp; Geographical</strong>: Categories, Castes, Religions, Countries, and States</li>
                  <li><strong>Academic Qualifications</strong>: Qualifying examination and stream specializations</li>
                </ul>
              </div>

              <div>
                <strong style={{ fontSize: '13px', display: 'block', marginBottom: '4px' }}>2. Side-by-Side Dual View &amp; Excel Ingestion</strong>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  Switching to <strong>"Side-by-Side View"</strong> displays the Programme Master and College Master registers concurrently on the same screen. You can also upload a single Excel worksheet containing both Programme columns on the left and College columns on the right (separated by an empty column or together). The tool auto-detects both headers simultaneously:
                </p>
                <div style={{ padding: '8px 12px', background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '6px', fontFamily: 'monospace', fontSize: '11px', marginTop: '6px' }}>
                  Programme Name (Reference) | ProgramCode | workFlowId | Program Name | [empty] | CollegeId | CollegeCode | College Name | CLG Name
                </div>
              </div>

              <div>
                <strong style={{ fontSize: '13px', display: 'block', marginBottom: '4px' }}>3. Interactive String Lookup Resolver &amp; Validator</strong>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  Use the live validator at the top to test any raw candidate string (e.g. <code>'Karnataka'</code>, <code>'Kannur IMPES'</code>, <code>'Music Dept'</code>). The simulator highlights exact matches, code matches, or keyword fallbacks instantly before you import student batches.
                </p>
              </div>

              <div>
                <strong style={{ fontSize: '13px', display: 'block', marginBottom: '4px' }}>4. Live Synchronization Across Report Studio</strong>
                <p style={{ margin: 0, color: 'var(--muted)' }}>
                  Any updates, additions, or spreadsheet imports saved here are stored in persistent browser storage (<code>localStorage</code>). When you run <strong>Admission Import</strong> or <strong>Admission Report Maker - Row Wise</strong>, they instantly utilize your latest definitions!
                </p>
              </div>
            </div>

            <div style={{ padding: '12px 20px', borderTop: '1px solid var(--line)', background: 'var(--bg)', display: 'flex', justifyContent: 'flex-end' }}>
              <button
                onClick={() => setShowGuideModal(false)}
                style={{ padding: '7px 16px', borderRadius: '6px', background: 'var(--accent)', color: '#fff', border: 'none', fontSize: '12.5px', fontWeight: 600, cursor: 'pointer' }}
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
