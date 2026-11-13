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
  Layers,
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
  ShieldCheck,
  Eye,
  RefreshCw
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
        color: 'emerald'
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
        color: 'emerald'
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
        color: 'indigo'
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
        color: 'indigo'
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
        detail: `State Code ID: ${s.id}`,
        confidence: 'High (Official)',
        color: 'sky'
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
        color: 'amber'
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
        color: 'rose'
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
        color: 'purple'
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
        color: 'teal'
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
        color: 'blue'
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
          color: 'emerald'
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
          color: 'indigo'
        };
      }
    }

    return {
      entity: 'Unmapped',
      tabKey: null,
      matchType: 'No Master Record Match',
      title: q,
      id: 'N/A',
      detail: 'This string will use default fallback during admission import.',
      confidence: 'None',
      color: 'slate'
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
        // If numeric strings, sort numerically
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
    <div className="min-h-screen bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-100 flex flex-col font-sans transition-colors duration-150">
      {/* Top Header Command Bar */}
      <header className="bg-white/95 dark:bg-slate-800/95 backdrop-blur-md border-b border-slate-200 dark:border-slate-700 sticky top-0 z-30 shadow-xs">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3 flex flex-wrap items-center justify-between gap-3">
          {/* Left: Branding & Breadcrumbs */}
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
                <span className="p-1.5 rounded-lg bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
                  <Database className="w-4 h-4" />
                </span>
                <h1 className="text-base sm:text-lg font-bold text-slate-900 dark:text-white tracking-tight">
                  Programme &amp; College Master Definitions
                </h1>
                <span className="text-[10px] font-bold tracking-wider uppercase px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 dark:bg-emerald-950/80 dark:text-emerald-300 border border-emerald-300 dark:border-emerald-800 shadow-2xs">
                  {totalEntriesCount} Masters Loaded
                </span>
              </div>
              <p className="text-xs text-slate-500 dark:text-slate-400 hidden md:block">
                Canonical master reference directory for Kannur University admission codes, workflow IDs, campuses, and reservation quotas.
              </p>
            </div>
          </div>

          {/* Right: Actions & Tools */}
          <div className="flex items-center flex-wrap gap-2">
            {/* View Mode Toggle: Dual Split vs Tabbed */}
            <div className="flex items-center bg-slate-100 dark:bg-slate-700/70 p-0.5 rounded-lg border border-slate-200 dark:border-slate-600">
              <button
                onClick={() => setViewMode('split')}
                className={`flex items-center gap-1.5 px-2.5 py-1 text-xs font-semibold rounded-md transition-all ${
                  viewMode === 'split'
                    ? 'bg-white dark:bg-slate-800 text-emerald-700 dark:text-emerald-300 shadow-xs'
                    : 'text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200'
                }`}
                title="View Programme and College master registries side-by-side"
              >
                <Columns2 className="w-3.5 h-3.5" />
                Side-by-Side View
              </button>
              <button
                onClick={() => setViewMode('tabbed')}
                className={`flex items-center gap-1.5 px-2.5 py-1 text-xs font-semibold rounded-md transition-all ${
                  viewMode === 'tabbed'
                    ? 'bg-white dark:bg-slate-800 text-emerald-700 dark:text-emerald-300 shadow-xs'
                    : 'text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200'
                }`}
                title="View full 9 entity lookup tabs (Program, College, Caste, State, etc.)"
              >
                <Table className="w-3.5 h-3.5" />
                All 9 Masters
              </button>
            </div>

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
              title="Upload spreadsheet with Programme, College or lookup definitions"
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
              onClick={() => setShowResetConfirmModal(true)}
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

      {/* Main Content Area */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-5 w-full flex-1 flex flex-col gap-5">
        {/* Status Alert Banner */}
        {statusAlert && (
          <div
            className={`p-3 rounded-xl border flex items-start justify-between gap-3 text-xs sm:text-sm shadow-xs transition-all ${
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

        {/* Hero Interactive Stat Overview Cards */}
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4">
          <div
            onClick={() => { setViewMode('split'); handleTabChange('program'); }}
            className={`p-3.5 sm:p-4 rounded-xl border cursor-pointer transition-all ${
              activeTab === 'program' || viewMode === 'split'
                ? 'bg-emerald-50/70 dark:bg-emerald-950/40 border-emerald-300 dark:border-emerald-700 shadow-sm ring-1 ring-emerald-500/30'
                : 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 hover:border-emerald-300 shadow-xs'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="p-2 rounded-lg bg-emerald-100 dark:bg-emerald-900/60 text-emerald-700 dark:text-emerald-300">
                <BookOpen className="w-4 h-4 sm:w-5 sm:h-5" />
              </div>
              <span className="text-[10px] font-bold uppercase tracking-wider text-emerald-600 dark:text-emerald-400 bg-emerald-100/60 dark:bg-emerald-900/40 px-2 py-0.5 rounded-full">
                Active Master
              </span>
            </div>
            <div className="mt-2.5">
              <p className="text-[11px] font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">Programmes</p>
              <h3 className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-white">{(lookups.program || []).length}</h3>
              <p className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5">Reference ↔ Code ↔ WF ID</p>
            </div>
          </div>

          <div
            onClick={() => { setViewMode('split'); handleTabChange('college'); }}
            className={`p-3.5 sm:p-4 rounded-xl border cursor-pointer transition-all ${
              activeTab === 'college' || viewMode === 'split'
                ? 'bg-indigo-50/70 dark:bg-indigo-950/40 border-indigo-300 dark:border-indigo-700 shadow-sm ring-1 ring-indigo-500/30'
                : 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 hover:border-indigo-300 shadow-xs'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="p-2 rounded-lg bg-indigo-100 dark:bg-indigo-900/60 text-indigo-700 dark:text-indigo-300">
                <Building2 className="w-4 h-4 sm:w-5 sm:h-5" />
              </div>
              <span className="text-[10px] font-bold uppercase tracking-wider text-indigo-600 dark:text-indigo-400 bg-indigo-100/60 dark:bg-indigo-900/40 px-2 py-0.5 rounded-full">
                Active Master
              </span>
            </div>
            <div className="mt-2.5">
              <p className="text-[11px] font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">Colleges &amp; Campuses</p>
              <h3 className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-white">{(lookups.college || []).length}</h3>
              <p className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5">ID ↔ Code ↔ Campus Name</p>
            </div>
          </div>

          <div
            onClick={() => { setViewMode('tabbed'); handleTabChange('caste'); }}
            className={`p-3.5 sm:p-4 rounded-xl border cursor-pointer transition-all ${
              (activeTab === 'caste' || activeTab === 'category') && viewMode === 'tabbed'
                ? 'bg-amber-50/70 dark:bg-amber-950/40 border-amber-300 dark:border-amber-700 shadow-sm ring-1 ring-amber-500/30'
                : 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 hover:border-amber-300 shadow-xs'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="p-2 rounded-lg bg-amber-100 dark:bg-amber-900/60 text-amber-700 dark:text-amber-300">
                <Users className="w-4 h-4 sm:w-5 sm:h-5" />
              </div>
              <span className="text-[10px] font-bold text-slate-400">
                {(lookups.caste || []).length} Castes
              </span>
            </div>
            <div className="mt-2.5">
              <p className="text-[11px] font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">Social Quota</p>
              <h3 className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-white">
                {((lookups.caste || []).length) + ((lookups.category || []).length)}
              </h3>
              <p className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5">Castes &amp; Reservation Categories</p>
            </div>
          </div>

          <div
            onClick={() => { setViewMode('tabbed'); handleTabChange('state'); }}
            className={`p-3.5 sm:p-4 rounded-xl border cursor-pointer transition-all ${
              (activeTab === 'state' || activeTab === 'country' || activeTab === 'qualification' || activeTab === 'stream') && viewMode === 'tabbed'
                ? 'bg-sky-50/70 dark:bg-sky-950/40 border-sky-300 dark:border-sky-700 shadow-sm ring-1 ring-sky-500/30'
                : 'bg-white dark:bg-slate-800 border-slate-200 dark:border-slate-700 hover:border-sky-300 shadow-xs'
            }`}
          >
            <div className="flex items-center justify-between">
              <div className="p-2 rounded-lg bg-sky-100 dark:bg-sky-900/60 text-sky-700 dark:text-sky-300">
                <MapPin className="w-4 h-4 sm:w-5 sm:h-5" />
              </div>
              <span className="text-[10px] font-bold text-slate-400">
                {(lookups.state || []).length} States
              </span>
            </div>
            <div className="mt-2.5">
              <p className="text-[11px] font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">Geography &amp; Boards</p>
              <h3 className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-white">
                {((lookups.state || []).length) + ((lookups.country || []).length) + ((lookups.qualification || []).length) + ((lookups.stream || []).length)}
              </h3>
              <p className="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5">States, Boards, Streams, etc.</p>
            </div>
          </div>
        </div>

        {/* Live Lookup Resolver / Validator Sandbox */}
        <div className="rounded-xl border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 shadow-xs overflow-hidden">
          <div className="px-4 py-2.5 bg-slate-50 dark:bg-slate-800/80 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <span className="p-1 rounded-md bg-amber-500/10 text-amber-600 dark:text-amber-400">
                <Zap className="w-3.5 h-3.5" />
              </span>
              <h3 className="text-xs font-bold text-slate-800 dark:text-slate-200 tracking-wide uppercase">
                Interactive String Lookup Resolver &amp; Validator
              </h3>
              <span className="text-[10px] text-slate-400 font-normal hidden sm:inline">
                Test how candidate rows from admission files map to university codes in real-time
              </span>
            </div>
            <button
              onClick={() => setShowResolver(!showResolver)}
              className="text-xs text-slate-500 hover:text-slate-800 dark:hover:text-slate-200 font-medium flex items-center gap-1"
            >
              {showResolver ? 'Collapse' : 'Expand'}
            </button>
          </div>

          {showResolver && (
            <div className="p-4 bg-gradient-to-b from-slate-50/50 to-white dark:from-slate-800/50 dark:to-slate-800">
              <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3">
                <div className="relative flex-1">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <input
                    type="text"
                    value={resolverInput}
                    onChange={(e) => setResolverInput(e.target.value)}
                    placeholder="Type any test raw string (e.g. 'Kannur Music', 'PFPSYGS25', 'Karnataka', 'Mukkuvan', 'HSE', 'Dept of IT')..."
                    className="w-full pl-9 pr-8 py-2 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/30 focus:border-emerald-500"
                  />
                  {resolverInput && (
                    <button
                      onClick={() => setResolverInput('')}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  )}
                </div>

                {/* Quick Presets */}
                <div className="flex items-center gap-1.5 overflow-x-auto text-[11px] shrink-0">
                  <span className="text-slate-400 font-medium">Quick test:</span>
                  {[
                    'Karnataka',
                    'Five Year Integrated Masters Programme in Music',
                    'Department of Music',
                    'Mukkuvan',
                    'SEBC-Muslim'
                  ].map(preset => (
                    <button
                      key={preset}
                      onClick={() => setResolverInput(preset)}
                      className="px-2 py-1 rounded bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600 transition-colors whitespace-nowrap"
                    >
                      {preset.length > 20 ? preset.slice(0, 18) + '...' : preset}
                    </button>
                  ))}
                </div>
              </div>

              {/* Real-time Match Output Badge */}
              {resolvedResult && resolverInput.trim() && (
                <div className="mt-3 p-3 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800/90 shadow-2xs flex flex-wrap items-center justify-between gap-3 text-xs">
                  <div className="flex items-center gap-2.5 flex-wrap">
                    <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider ${
                      resolvedResult.color === 'emerald' ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950 dark:text-emerald-300 border border-emerald-300' :
                      resolvedResult.color === 'indigo' ? 'bg-indigo-100 text-indigo-800 dark:bg-indigo-950 dark:text-indigo-300 border border-indigo-300' :
                      resolvedResult.color === 'amber' ? 'bg-amber-100 text-amber-800 dark:bg-amber-950 dark:text-amber-300 border border-amber-300' :
                      resolvedResult.color === 'sky' ? 'bg-sky-100 text-sky-800 dark:bg-sky-950 dark:text-sky-300 border border-sky-300' :
                      resolvedResult.color === 'rose' ? 'bg-rose-100 text-rose-800 dark:bg-rose-950 dark:text-rose-300 border border-rose-300' :
                      'bg-slate-100 text-slate-700 dark:bg-slate-700 dark:text-slate-300'
                    }`}>
                      {resolvedResult.entity} Match
                    </span>
                    <span className="font-semibold text-slate-800 dark:text-slate-200">
                      {resolvedResult.title}
                    </span>
                    {resolvedResult.code && (
                      <span className="font-mono px-2 py-0.5 rounded bg-emerald-50 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-300 font-bold border border-emerald-200 dark:border-emerald-800">
                        Code: {resolvedResult.code}
                      </span>
                    )}
                    {resolvedResult.id && resolvedResult.id !== 'N/A' && (
                      <span className="font-mono px-2 py-0.5 rounded bg-slate-100 dark:bg-slate-700 text-slate-800 dark:text-slate-200 font-bold border border-slate-200 dark:border-slate-600">
                        ID: {resolvedResult.id}
                      </span>
                    )}
                  </div>

                  <div className="flex items-center gap-2 text-slate-400 text-[11px]">
                    <span className="inline-flex items-center gap-1 text-emerald-600 dark:text-emerald-400 font-medium">
                      <CheckCircle2 className="w-3.5 h-3.5" />
                      {resolvedResult.matchType} ({resolvedResult.confidence})
                    </span>
                    {resolvedResult.tabKey && (
                      <button
                        onClick={() => {
                          setViewMode('tabbed');
                          handleTabChange(resolvedResult.tabKey);
                          setSearchQuery(resolvedResult.title);
                        }}
                        className="text-xs text-indigo-600 dark:text-indigo-400 hover:underline font-medium inline-flex items-center gap-0.5"
                      >
                        Find in Directory <ExternalLink className="w-3 h-3" />
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
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
            {/* Left Pane: Programme Master */}
            <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-xs overflow-hidden flex flex-col">
              {/* Pane Header */}
              <div className="p-3.5 bg-emerald-50/50 dark:bg-emerald-950/20 border-b border-emerald-100 dark:border-emerald-900/60 flex items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <span className="p-1.5 rounded-lg bg-emerald-500/10 text-emerald-600 dark:text-emerald-400">
                    <BookOpen className="w-4 h-4" />
                  </span>
                  <div>
                    <h2 className="text-sm font-bold text-slate-900 dark:text-white flex items-center gap-1.5">
                      Programme Master Directory
                      <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 dark:bg-emerald-900 dark:text-emerald-300">
                        {(lookups.program || []).length}
                      </span>
                    </h2>
                    <p className="text-[11px] text-slate-500 dark:text-slate-400">
                      Programme Name (Reference) ↔ ProgramCode ↔ workFlowId
                    </p>
                  </div>
                </div>

                <button
                  onClick={() => {
                    handleTabChange('program');
                    setIsAddFormOpen(true);
                  }}
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-2xs"
                >
                  <Plus className="w-3.5 h-3.5" /> Add
                </button>
              </div>

              {/* Pane Search */}
              <div className="p-2.5 bg-slate-50/50 dark:bg-slate-800/50 border-b border-slate-200 dark:border-slate-700 flex items-center gap-2">
                <div className="relative flex-1">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <input
                    type="text"
                    value={progSearch}
                    onChange={(e) => { setProgSearch(e.target.value); setProgPage(0); }}
                    placeholder="Search programmes by name, code (PF...), or workflow ID..."
                    className="w-full pl-8 pr-7 py-1 text-xs rounded border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                  {progSearch && (
                    <button
                      onClick={() => setProgSearch('')}
                      className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                    >
                      <X className="w-3 h-3" />
                    </button>
                  )}
                </div>
                <span className="text-[11px] text-slate-400 shrink-0 font-medium">
                  {dualProgList.length} items
                </span>
              </div>

              {/* Programme Table */}
              <div className="overflow-x-auto flex-1 min-h-[360px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead>
                    <tr className="bg-slate-100/70 dark:bg-slate-800/70 text-slate-600 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700 text-[11px] uppercase tracking-wider font-semibold">
                      <th className="py-2 px-2.5 w-10 text-center">#</th>
                      <th className="py-2 px-2.5 min-w-[160px]">Reference Name</th>
                      <th className="py-2 px-2 w-28">ProgramCode</th>
                      <th className="py-2 px-2 w-24">workFlowId</th>
                      <th className="py-2 px-2 text-right w-16">Act</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60">
                    {dualPaginatedProgs.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="py-12 text-center text-slate-400">
                          No programmes matching search query.
                        </td>
                      </tr>
                    ) : (
                      dualPaginatedProgs.map((item, idx) => {
                        const rowNum = progPage * dualPageSize + idx + 1;
                        const isEditing = editingTargetTab === 'program' && editingIdx === (lookups.program || []).indexOf(item);

                        return (
                          <tr key={idx} className="hover:bg-slate-50/80 dark:hover:bg-slate-700/40">
                            <td className="py-2 px-2 text-center text-slate-400 font-mono text-[11px]">
                              {rowNum}
                            </td>
                            <td className="py-2 px-2.5 font-medium text-slate-800 dark:text-slate-100">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editProgRef}
                                  onChange={(e) => setEditProgRef(e.target.value)}
                                  className="w-full px-1.5 py-0.5 text-xs rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                <div>
                                  <span>{item.pgmName}</span>
                                  <p className="text-[10px] text-slate-400 line-clamp-1">{item.name}</p>
                                </div>
                              )}
                            </td>
                            <td className="py-2 px-2">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editProgCode}
                                  onChange={(e) => setEditProgCode(e.target.value)}
                                  className="w-full px-1.5 py-0.5 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                <button
                                  onClick={() => handleCopy(item.code, `prog_code_${idx}`)}
                                  className="inline-flex items-center gap-1 font-mono text-[11px] px-1.5 py-0.5 rounded bg-emerald-50 dark:bg-emerald-950/50 text-emerald-700 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 font-bold hover:bg-emerald-100"
                                  title="Click to copy code"
                                >
                                  {copiedKey === `prog_code_${idx}` ? <CheckCheck className="w-3 h-3 text-emerald-600" /> : <Copy className="w-2.5 h-2.5 opacity-40" />}
                                  {item.code}
                                </button>
                              )}
                            </td>
                            <td className="py-2 px-2">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editProgWorkflow}
                                  onChange={(e) => setEditProgWorkflow(e.target.value)}
                                  className="w-full px-1.5 py-0.5 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                <button
                                  onClick={() => handleCopy(item.workflowId, `prog_wf_${idx}`)}
                                  className="inline-flex items-center gap-1 font-mono text-[11px] px-1.5 py-0.5 rounded bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 font-bold border border-slate-200 dark:border-slate-600 hover:bg-slate-200"
                                  title="Click to copy workflow ID"
                                >
                                  {copiedKey === `prog_wf_${idx}` ? <CheckCheck className="w-3 h-3 text-emerald-600" /> : <Copy className="w-2.5 h-2.5 opacity-40" />}
                                  {item.workflowId}
                                </button>
                              )}
                            </td>
                            <td className="py-2 px-2 text-right">
                              <div className="flex items-center justify-end gap-1">
                                {isEditing ? (
                                  <>
                                    <button
                                      onClick={() => handleSaveEdit('program')}
                                      className="p-1 rounded bg-emerald-600 hover:bg-emerald-700 text-white"
                                      title="Save edit"
                                    >
                                      <Check className="w-3 h-3" />
                                    </button>
                                    <button
                                      onClick={() => { setEditingIdx(null); setEditingTargetTab(null); }}
                                      className="p-1 rounded border border-slate-300 dark:border-slate-600 text-slate-500 hover:bg-slate-100"
                                      title="Cancel"
                                    >
                                      <X className="w-3 h-3" />
                                    </button>
                                  </>
                                ) : (
                                  <>
                                    <button
                                      onClick={() => handleStartEdit(item, 'program')}
                                      className="p-1 rounded text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
                                      title="Edit programme"
                                    >
                                      <Edit2 className="w-3 h-3" />
                                    </button>
                                    <button
                                      onClick={() => handleDelete(item, 'program')}
                                      className="p-1 rounded text-slate-400 hover:text-rose-600 dark:hover:text-rose-400"
                                      title="Delete programme"
                                    >
                                      <Trash2 className="w-3 h-3" />
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

              {/* Programme Pagination */}
              {dualProgList.length > dualPageSize && (
                <div className="p-2.5 bg-slate-50/70 dark:bg-slate-800/70 border-t border-slate-200 dark:border-slate-700 flex items-center justify-between text-xs text-slate-500">
                  <span>Page {progPage + 1} of {Math.ceil(dualProgList.length / dualPageSize)}</span>
                  <div className="flex items-center gap-1">
                    <button
                      onClick={() => setProgPage(p => Math.max(0, p - 1))}
                      disabled={progPage === 0}
                      className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-600 disabled:opacity-40"
                    >
                      Prev
                    </button>
                    <button
                      onClick={() => setProgPage(p => Math.min(Math.ceil(dualProgList.length / dualPageSize) - 1, p + 1))}
                      disabled={progPage >= Math.ceil(dualProgList.length / dualPageSize) - 1}
                      className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-600 disabled:opacity-40"
                    >
                      Next
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Right Pane: College Master */}
            <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-xs overflow-hidden flex flex-col">
              {/* Pane Header */}
              <div className="p-3.5 bg-indigo-50/50 dark:bg-indigo-950/20 border-b border-indigo-100 dark:border-indigo-900/60 flex items-center justify-between gap-3">
                <div className="flex items-center gap-2">
                  <span className="p-1.5 rounded-lg bg-indigo-500/10 text-indigo-600 dark:text-indigo-400">
                    <Building2 className="w-4 h-4" />
                  </span>
                  <div>
                    <h2 className="text-sm font-bold text-slate-900 dark:text-white flex items-center gap-1.5">
                      College Master Directory
                      <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-indigo-100 text-indigo-800 dark:bg-indigo-900 dark:text-indigo-300">
                        {(lookups.college || []).length}
                      </span>
                    </h2>
                    <p className="text-[11px] text-slate-500 dark:text-slate-400">
                      CollegeId ↔ CollegeCode ↔ College Name ↔ Campus / CLG Name
                    </p>
                  </div>
                </div>

                <button
                  onClick={() => {
                    handleTabChange('college');
                    setIsAddFormOpen(true);
                  }}
                  className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold shadow-2xs"
                >
                  <Plus className="w-3.5 h-3.5" /> Add
                </button>
              </div>

              {/* Pane Search */}
              <div className="p-2.5 bg-slate-50/50 dark:bg-slate-800/50 border-b border-slate-200 dark:border-slate-700 flex items-center gap-2">
                <div className="relative flex-1">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <input
                    type="text"
                    value={clgSearch}
                    onChange={(e) => { setClgSearch(e.target.value); setClgPage(0); }}
                    placeholder="Search colleges by name, ID (e.g. 127, 435), or campus..."
                    className="w-full pl-8 pr-7 py-1 text-xs rounded border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-indigo-500"
                  />
                  {clgSearch && (
                    <button
                      onClick={() => setClgSearch('')}
                      className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                    >
                      <X className="w-3 h-3" />
                    </button>
                  )}
                </div>
                <span className="text-[11px] text-slate-400 shrink-0 font-medium">
                  {dualClgList.length} items
                </span>
              </div>

              {/* College Table */}
              <div className="overflow-x-auto flex-1 min-h-[360px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead>
                    <tr className="bg-slate-100/70 dark:bg-slate-800/70 text-slate-600 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700 text-[11px] uppercase tracking-wider font-semibold">
                      <th className="py-2 px-2.5 w-10 text-center">#</th>
                      <th className="py-2 px-2 w-20">CollegeId</th>
                      <th className="py-2 px-2 w-20">Code</th>
                      <th className="py-2 px-2.5 min-w-[160px]">College / Campus Name</th>
                      <th className="py-2 px-2 text-right w-16">Act</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60">
                    {dualPaginatedColleges.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="py-12 text-center text-slate-400">
                          No colleges matching search query.
                        </td>
                      </tr>
                    ) : (
                      dualPaginatedColleges.map((item, idx) => {
                        const rowNum = clgPage * dualPageSize + idx + 1;
                        const isEditing = editingTargetTab === 'college' && editingIdx === (lookups.college || []).indexOf(item);

                        return (
                          <tr key={idx} className="hover:bg-slate-50/80 dark:hover:bg-slate-700/40">
                            <td className="py-2 px-2 text-center text-slate-400 font-mono text-[11px]">
                              {rowNum}
                            </td>
                            <td className="py-2 px-2">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editClgId}
                                  onChange={(e) => setEditClgId(e.target.value)}
                                  className="w-full px-1.5 py-0.5 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                <button
                                  onClick={() => handleCopy(item.id, `clg_id_${idx}`)}
                                  className="inline-flex items-center gap-1 font-mono text-[11px] px-1.5 py-0.5 rounded bg-indigo-50 dark:bg-indigo-950/50 text-indigo-700 dark:text-indigo-300 border border-indigo-200 dark:border-indigo-800 font-bold hover:bg-indigo-100"
                                  title="Click to copy CollegeId"
                                >
                                  {copiedKey === `clg_id_${idx}` ? <CheckCheck className="w-3 h-3 text-indigo-600" /> : <Copy className="w-2.5 h-2.5 opacity-40" />}
                                  {item.id}
                                </button>
                              )}
                            </td>
                            <td className="py-2 px-2 font-mono text-slate-600 dark:text-slate-300">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editClgCode}
                                  onChange={(e) => setEditClgCode(e.target.value)}
                                  className="w-full px-1.5 py-0.5 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                item.code
                              )}
                            </td>
                            <td className="py-2 px-2.5 font-medium text-slate-800 dark:text-slate-100">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={editClgName}
                                  onChange={(e) => setEditClgName(e.target.value)}
                                  className="w-full px-1.5 py-0.5 text-xs rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                />
                              ) : (
                                <div>
                                  <span>{item.name}</span>
                                  {item.clgRef && item.clgRef !== item.name && (
                                    <p className="text-[10px] text-slate-400 line-clamp-1">{item.clgRef}</p>
                                  )}
                                </div>
                              )}
                            </td>
                            <td className="py-2 px-2 text-right">
                              <div className="flex items-center justify-end gap-1">
                                {isEditing ? (
                                  <>
                                    <button
                                      onClick={() => handleSaveEdit('college')}
                                      className="p-1 rounded bg-indigo-600 hover:bg-indigo-700 text-white"
                                      title="Save edit"
                                    >
                                      <Check className="w-3 h-3" />
                                    </button>
                                    <button
                                      onClick={() => { setEditingIdx(null); setEditingTargetTab(null); }}
                                      className="p-1 rounded border border-slate-300 dark:border-slate-600 text-slate-500 hover:bg-slate-100"
                                      title="Cancel"
                                    >
                                      <X className="w-3 h-3" />
                                    </button>
                                  </>
                                ) : (
                                  <>
                                    <button
                                      onClick={() => handleStartEdit(item, 'college')}
                                      className="p-1 rounded text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
                                      title="Edit college"
                                    >
                                      <Edit2 className="w-3 h-3" />
                                    </button>
                                    <button
                                      onClick={() => handleDelete(item, 'college')}
                                      className="p-1 rounded text-slate-400 hover:text-rose-600 dark:hover:text-rose-400"
                                      title="Delete college"
                                    >
                                      <Trash2 className="w-3 h-3" />
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

              {/* College Pagination */}
              {dualClgList.length > dualPageSize && (
                <div className="p-2.5 bg-slate-50/70 dark:bg-slate-800/70 border-t border-slate-200 dark:border-slate-700 flex items-center justify-between text-xs text-slate-500">
                  <span>Page {clgPage + 1} of {Math.ceil(dualClgList.length / dualPageSize)}</span>
                  <div className="flex items-center gap-1">
                    <button
                      onClick={() => setClgPage(p => Math.max(0, p - 1))}
                      disabled={clgPage === 0}
                      className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-600 disabled:opacity-40"
                    >
                      Prev
                    </button>
                    <button
                      onClick={() => setClgPage(p => Math.min(Math.ceil(dualClgList.length / dualPageSize) - 1, p + 1))}
                      disabled={clgPage >= Math.ceil(dualClgList.length / dualPageSize) - 1}
                      className="px-2 py-0.5 rounded border border-slate-300 dark:border-slate-600 disabled:opacity-40"
                    >
                      Next
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* View Mode 2: TABBED MASTER DIRECTORY (All 9 Categories) */}
        {viewMode === 'tabbed' && (
          <div className="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-xs overflow-hidden flex flex-col">
            {/* Tab Navigation Ribbon */}
            <div className="border-b border-slate-200 dark:border-slate-700 bg-slate-50/80 dark:bg-slate-800/90 px-4 pt-3 flex items-center gap-1.5 overflow-x-auto scrollbar-thin">
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

            {/* Subheader: Search, Filters, Quick Add Button */}
            <div className="p-3.5 sm:p-4 bg-slate-50/40 dark:bg-slate-800/40 border-b border-slate-200 dark:border-slate-700 flex flex-wrap items-center justify-between gap-3">
              <div>
                <h2 className="text-sm font-bold text-slate-800 dark:text-slate-200 flex items-center gap-2">
                  <span>{currentTabConfig.label} Directory</span>
                  <span className="text-xs font-normal text-slate-400 hidden sm:inline">({currentTabConfig.description})</span>
                </h2>
                <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                  Showing <strong>{filteredList.length}</strong> of <strong>{currentList.length}</strong> defined entries
                  {selectedRowKeys.size > 0 && (
                    <span className="ml-2 text-emerald-600 dark:text-emerald-400 font-semibold">
                      ({selectedRowKeys.size} selected)
                    </span>
                  )}
                </p>
              </div>

              <div className="flex items-center flex-wrap gap-2 w-full sm:w-auto">
                {/* Search Bar */}
                <div className="relative flex-1 sm:w-64">
                  <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => { setSearchQuery(e.target.value); setPage(0); }}
                    placeholder={`Search ${currentTabConfig.label.toLowerCase()}...`}
                    className="w-full pl-8 pr-7 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                  />
                  {searchQuery && (
                    <button
                      onClick={() => setSearchQuery('')}
                      className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                    >
                      <X className="w-3 h-3" />
                    </button>
                  )}
                </div>

                {/* Page Size Select */}
                <select
                  value={pageSize}
                  onChange={(e) => { setPageSize(Number(e.target.value)); setPage(0); }}
                  className="px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-medium focus:outline-none"
                >
                  <option value={15}>15 / page</option>
                  <option value={25}>25 / page</option>
                  <option value={50}>50 / page</option>
                  <option value={100}>100 / page</option>
                  <option value={0}>Show All</option>
                </select>

                {/* Add Entry Toggle */}
                <button
                  onClick={() => setIsAddFormOpen(!isAddFormOpen)}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-2xs transition-colors shrink-0"
                >
                  <Plus className="w-3.5 h-3.5" />
                  {isAddFormOpen ? 'Close Form' : 'Add Master Entry'}
                </button>
              </div>
            </div>

            {/* Quick Filter Chips (for Programme, College, Caste) */}
            {(activeTab === 'program' || activeTab === 'college' || activeTab === 'caste') && (
              <div className="px-4 py-2 bg-slate-100/50 dark:bg-slate-800/60 border-b border-slate-200 dark:border-slate-700 flex items-center gap-1.5 overflow-x-auto text-[11px]">
                <span className="text-slate-400 font-medium mr-1 flex items-center gap-1">
                  <Filter className="w-3 h-3" /> Filters:
                </span>
                {activeTab === 'program' && [
                  { id: 'all', label: 'All Programmes' },
                  { id: 'fyimp', label: 'Integrated Masters (FYIMP)' },
                  { id: 'regular', label: 'Regular PG / Other' }
                ].map(chip => (
                  <button
                    key={chip.id}
                    onClick={() => { setFilterChip(chip.id); setPage(0); }}
                    className={`px-2.5 py-0.5 rounded-full font-medium transition-colors ${
                      filterChip === chip.id
                        ? 'bg-emerald-600 text-white'
                        : 'bg-white dark:bg-slate-700 text-slate-600 dark:text-slate-300 border border-slate-200 dark:border-slate-600 hover:bg-slate-50'
                    }`}
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
                    className={`px-2.5 py-0.5 rounded-full font-medium transition-colors ${
                      filterChip === chip.id
                        ? 'bg-indigo-600 text-white'
                        : 'bg-white dark:bg-slate-700 text-slate-600 dark:text-slate-300 border border-slate-200 dark:border-slate-600 hover:bg-slate-50'
                    }`}
                  >
                    {chip.label}
                  </button>
                ))}

                {activeTab === 'caste' && [
                  { id: 'all', label: 'All Castes' },
                  { id: 'sebc', label: 'OBC / SEBC Castes' },
                  { id: 'sc', label: 'Scheduled Castes' }
                ].map(chip => (
                  <button
                    key={chip.id}
                    onClick={() => { setFilterChip(chip.id); setPage(0); }}
                    className={`px-2.5 py-0.5 rounded-full font-medium transition-colors ${
                      filterChip === chip.id
                        ? 'bg-amber-600 text-white'
                        : 'bg-white dark:bg-slate-700 text-slate-600 dark:text-slate-300 border border-slate-200 dark:border-slate-600 hover:bg-slate-50'
                    }`}
                  >
                    {chip.label}
                  </button>
                ))}
              </div>
            )}

            {/* Expandable Add Entry Form */}
            {isAddFormOpen && (
              <div className="p-4 bg-emerald-50/40 dark:bg-emerald-950/30 border-b border-emerald-200 dark:border-emerald-900/60 transition-all">
                <form onSubmit={handleAdd} className="flex flex-col gap-3">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-emerald-800 dark:text-emerald-300 flex items-center gap-1.5">
                      <Plus className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
                      Add New {currentTabConfig.label} Entry:
                    </span>
                    <button
                      type="button"
                      onClick={() => setIsAddFormOpen(false)}
                      className="text-xs text-slate-400 hover:text-slate-600"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
                    {activeTab === 'program' ? (
                      <>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            Reference Name *
                          </label>
                          <input
                            type="text"
                            value={newProgRef}
                            onChange={(e) => setNewProgRef(e.target.value)}
                            placeholder="e.g. Five Year Integrated..."
                            className="w-full px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            ProgramCode *
                          </label>
                          <input
                            type="text"
                            value={newProgCode}
                            onChange={(e) => setNewProgCode(e.target.value)}
                            placeholder="e.g. PFPSYGS25"
                            className="w-full px-2.5 py-1.5 text-xs font-mono rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            workFlowId *
                          </label>
                          <input
                            type="text"
                            value={newProgWorkflow}
                            onChange={(e) => setNewProgWorkflow(e.target.value)}
                            placeholder="e.g. 151"
                            className="w-full px-2.5 py-1.5 text-xs font-mono rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            Official Full Program Name *
                          </label>
                          <input
                            type="text"
                            value={newProgName}
                            onChange={(e) => setNewProgName(e.target.value)}
                            placeholder="e.g. (with Credits)-Regular-FYIMP..."
                            className="w-full px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                      </>
                    ) : activeTab === 'college' ? (
                      <>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            CollegeId *
                          </label>
                          <input
                            type="text"
                            value={newClgId}
                            onChange={(e) => setNewClgId(e.target.value)}
                            placeholder="e.g. 435"
                            className="w-full px-2.5 py-1.5 text-xs font-mono rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            CollegeCode
                          </label>
                          <input
                            type="text"
                            value={newClgCode}
                            onChange={(e) => setNewClgCode(e.target.value)}
                            placeholder="e.g. 435"
                            className="w-full px-2.5 py-1.5 text-xs font-mono rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            College Name *
                          </label>
                          <input
                            type="text"
                            value={newClgName}
                            onChange={(e) => setNewClgName(e.target.value)}
                            placeholder="e.g. Department of Anthropology"
                            className="w-full px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                        <div>
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            CLG Name (Campus Details)
                          </label>
                          <input
                            type="text"
                            value={newClgRef}
                            onChange={(e) => setNewClgRef(e.target.value)}
                            placeholder="e.g. Dr. Janaki Ammal Campus, Palayad"
                            className="w-full px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                      </>
                    ) : (
                      <>
                        <div className="sm:col-span-2">
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            {currentTabConfig.nameLabel} *
                          </label>
                          <input
                            type="text"
                            value={newName}
                            onChange={(e) => setNewName(e.target.value)}
                            placeholder={`Enter ${currentTabConfig.nameLabel.toLowerCase()}...`}
                            className="w-full px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                        <div className="sm:col-span-2">
                          <label className="block text-[11px] font-semibold text-slate-600 dark:text-slate-300 mb-1">
                            {currentTabConfig.idLabel} *
                          </label>
                          <input
                            type="text"
                            value={newId}
                            onChange={(e) => setNewId(e.target.value)}
                            placeholder={`Enter ${currentTabConfig.idLabel.toLowerCase()}...`}
                            className="w-full px-2.5 py-1.5 text-xs font-mono rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-800 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                          />
                        </div>
                      </>
                    )}
                  </div>

                  <div className="flex justify-end gap-2 mt-1">
                    <button
                      type="button"
                      onClick={() => setIsAddFormOpen(false)}
                      className="px-3 py-1.5 rounded-lg border border-slate-300 dark:border-slate-600 text-slate-600 dark:text-slate-300 hover:bg-slate-100 text-xs font-medium"
                    >
                      Cancel
                    </button>
                    <button
                      type="submit"
                      className="px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-xs"
                    >
                      Save Entry to {currentTabConfig.label}
                    </button>
                  </div>
                </form>
              </div>
            )}

            {/* Floating Batch Action Toolbar */}
            {selectedRowKeys.size > 0 && (
              <div className="px-4 py-2 bg-emerald-50 dark:bg-emerald-950/60 border-b border-emerald-200 dark:border-emerald-800 flex items-center justify-between text-xs animate-in fade-in">
                <span className="font-semibold text-emerald-900 dark:text-emerald-200">
                  {selectedRowKeys.size} item(s) selected
                </span>
                <div className="flex items-center gap-2">
                  <button
                    onClick={handleBulkDelete}
                    className="inline-flex items-center gap-1 px-2.5 py-1 rounded bg-rose-600 hover:bg-rose-700 text-white font-semibold text-xs shadow-2xs"
                  >
                    <Trash2 className="w-3 h-3" /> Delete Selected
                  </button>
                  <button
                    onClick={() => setSelectedRowKeys(new Set())}
                    className="px-2 py-1 rounded border border-slate-300 text-slate-600 dark:text-slate-300 hover:bg-slate-100 text-xs"
                  >
                    Clear Selection
                  </button>
                </div>
              </div>
            )}

            {/* Table Container */}
            <div className="overflow-x-auto flex-1">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="bg-slate-100/80 dark:bg-slate-800/80 text-slate-600 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700 font-semibold uppercase tracking-wider text-[11px]">
                    <th className="py-2.5 px-3 w-10 text-center">
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
                        className="rounded text-emerald-600 focus:ring-emerald-500"
                      />
                    </th>
                    <th className="py-2.5 px-2 w-10 text-center">#</th>

                    {activeTab === 'program' ? (
                      <>
                        <th
                          onClick={() => handleSort('pgmName')}
                          className="py-2.5 px-3 min-w-[200px] cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>Programme Name (Reference)</span>
                            {sortConfig.column === 'pgmName' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                        <th
                          onClick={() => handleSort('code')}
                          className="py-2.5 px-3 w-32 cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>ProgramCode</span>
                            {sortConfig.column === 'code' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                        <th
                          onClick={() => handleSort('workflowId')}
                          className="py-2.5 px-3 w-28 cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>workFlowId</span>
                            {sortConfig.column === 'workflowId' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                        <th
                          onClick={() => handleSort('name')}
                          className="py-2.5 px-3 min-w-[280px] cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>Full Program Name</span>
                            {sortConfig.column === 'name' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                      </>
                    ) : activeTab === 'college' ? (
                      <>
                        <th
                          onClick={() => handleSort('id')}
                          className="py-2.5 px-3 w-28 cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>CollegeId</span>
                            {sortConfig.column === 'id' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                        <th
                          onClick={() => handleSort('code')}
                          className="py-2.5 px-3 w-28 cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>CollegeCode</span>
                            {sortConfig.column === 'code' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                        <th
                          onClick={() => handleSort('name')}
                          className="py-2.5 px-3 min-w-[220px] cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>College Name</span>
                            {sortConfig.column === 'name' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                        <th
                          onClick={() => handleSort('clgRef')}
                          className="py-2.5 px-3 min-w-[280px] cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>CLG Name (Campus Details)</span>
                            {sortConfig.column === 'clgRef' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                      </>
                    ) : (
                      <>
                        <th
                          onClick={() => handleSort('name')}
                          className="py-2.5 px-3 min-w-[240px] cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>{currentTabConfig.nameHeader}</span>
                            {sortConfig.column === 'name' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                        <th
                          onClick={() => handleSort('id')}
                          className="py-2.5 px-3 w-48 cursor-pointer hover:text-slate-900 dark:hover:text-white"
                        >
                          <div className="flex items-center gap-1">
                            <span>{currentTabConfig.idHeader}</span>
                            {sortConfig.column === 'id' ? (
                              sortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-emerald-600" /> : <ArrowDown className="w-3 h-3 text-emerald-600" />
                            ) : (
                              <ArrowUpDown className="w-3 h-3 opacity-30" />
                            )}
                          </div>
                        </th>
                      </>
                    )}

                    <th className="py-2.5 px-3 w-24 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60">
                  {paginatedList.length === 0 ? (
                    <tr>
                      <td colSpan={activeTab === 'program' || activeTab === 'college' ? 7 : 5} className="py-12 text-center text-slate-400 dark:text-slate-500">
                        <div className="flex flex-col items-center justify-center gap-2">
                          <FileSpreadsheet className="w-8 h-8 opacity-40 text-slate-400" />
                          <p className="text-sm font-medium">
                            {searchQuery ? `No entries matching "${searchQuery}"` : `No ${currentTabConfig.label} defined yet.`}
                          </p>
                          <p className="text-xs text-slate-400">
                            Add an entry using the quick add form above or upload an Excel spreadsheet.
                          </p>
                        </div>
                      </td>
                    </tr>
                  ) : (
                    paginatedList.map((item, idx) => {
                      const rowNumber = (pageSize === 0 ? 0 : page * pageSize) + idx + 1;
                      const isEditing = editingIdx === currentList.indexOf(item);
                      const rowKey = `${activeTab}_${rowNumber - 1}_${item.id || item.code || item.name}`;
                      const isSelected = selectedRowKeys.has(rowKey);

                      return (
                        <tr
                          key={idx}
                          className={`hover:bg-slate-50/80 dark:hover:bg-slate-700/40 transition-colors ${
                            isEditing ? 'bg-amber-50/70 dark:bg-amber-950/20' : isSelected ? 'bg-emerald-50/40 dark:bg-emerald-950/20' : ''
                          }`}
                        >
                          <td className="py-2.5 px-3 text-center">
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={(e) => {
                                const next = new Set(selectedRowKeys);
                                if (e.target.checked) next.add(rowKey);
                                else next.delete(rowKey);
                                setSelectedRowKeys(next);
                              }}
                              className="rounded text-emerald-600 focus:ring-emerald-500"
                            />
                          </td>
                          <td className="py-2.5 px-2 text-center text-slate-400 font-mono text-[11px]">
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
                              <td className="py-2.5 px-3">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    value={editProgCode}
                                    onChange={(e) => setEditProgCode(e.target.value)}
                                    className="w-full px-2 py-1 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                  />
                                ) : (
                                  <button
                                    onClick={() => handleCopy(item.code, `tab_code_${idx}`)}
                                    className="inline-flex items-center gap-1 font-mono font-semibold px-2 py-0.5 rounded bg-emerald-50 dark:bg-emerald-950/50 text-emerald-700 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 hover:bg-emerald-100"
                                    title="Click to copy code"
                                  >
                                    {copiedKey === `tab_code_${idx}` ? <CheckCheck className="w-3 h-3 text-emerald-600" /> : <Copy className="w-2.5 h-2.5 opacity-40" />}
                                    {item.code}
                                  </button>
                                )}
                              </td>
                              <td className="py-2.5 px-3">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    value={editProgWorkflow}
                                    onChange={(e) => setEditProgWorkflow(e.target.value)}
                                    className="w-full px-2 py-1 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                  />
                                ) : (
                                  <button
                                    onClick={() => handleCopy(item.workflowId, `tab_wf_${idx}`)}
                                    className="inline-flex items-center gap-1 font-mono font-bold px-2 py-0.5 rounded bg-slate-100 dark:bg-slate-700 text-slate-800 dark:text-slate-200 border border-slate-200 dark:border-slate-600 hover:bg-slate-200"
                                    title="Click to copy workflow ID"
                                  >
                                    {copiedKey === `tab_wf_${idx}` ? <CheckCheck className="w-3 h-3 text-emerald-600" /> : <Copy className="w-2.5 h-2.5 opacity-40" />}
                                    {item.workflowId}
                                  </button>
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
                              <td className="py-2.5 px-3">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    value={editClgId}
                                    onChange={(e) => setEditClgId(e.target.value)}
                                    className="w-full px-2 py-1 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                  />
                                ) : (
                                  <button
                                    onClick={() => handleCopy(item.id, `tab_clg_id_${idx}`)}
                                    className="inline-flex items-center gap-1 font-mono font-bold px-2 py-0.5 rounded bg-indigo-50 dark:bg-indigo-950/50 text-indigo-700 dark:text-indigo-300 border border-indigo-200 dark:border-indigo-800 hover:bg-indigo-100"
                                    title="Click to copy CollegeId"
                                  >
                                    {copiedKey === `tab_clg_id_${idx}` ? <CheckCheck className="w-3 h-3 text-indigo-600" /> : <Copy className="w-2.5 h-2.5 opacity-40" />}
                                    {item.id}
                                  </button>
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
                              <td className="py-2.5 px-3">
                                {isEditing ? (
                                  <input
                                    type="text"
                                    value={editId}
                                    onChange={(e) => setEditId(e.target.value)}
                                    className="w-full px-2 py-1 text-xs font-mono rounded border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 focus:outline-none"
                                  />
                                ) : (
                                  <button
                                    onClick={() => handleCopy(item.id, `tab_id_${idx}`)}
                                    className="inline-flex items-center gap-1 font-mono font-bold px-2 py-0.5 rounded bg-emerald-50 dark:bg-emerald-950/50 text-emerald-700 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 hover:bg-emerald-100"
                                    title="Click to copy ID"
                                  >
                                    {copiedKey === `tab_id_${idx}` ? <CheckCheck className="w-3 h-3 text-emerald-600" /> : <Copy className="w-2.5 h-2.5 opacity-40" />}
                                    {item.id}
                                  </button>
                                )}
                              </td>
                            </>
                          )}

                          <td className="py-2.5 px-3 text-right">
                            <div className="flex items-center justify-end gap-1">
                              {isEditing ? (
                                <>
                                  <button
                                    onClick={() => handleSaveEdit(activeTab)}
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
                                    onClick={() => handleStartEdit(item, activeTab)}
                                    className="p-1 rounded text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-700"
                                    title="Edit entry"
                                  >
                                    <Edit2 className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    onClick={() => handleDelete(item, activeTab)}
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
                  Page <strong>{page + 1}</strong> of <strong>{totalPages}</strong> ({filteredList.length} total entries)
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
        )}
      </main>

      {/* Reset Confirmation Modal */}
      {showResetConfirmModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-white dark:bg-slate-800 rounded-2xl max-w-md w-full border border-slate-200 dark:border-slate-700 shadow-2xl p-6 flex flex-col gap-4">
            <div className="flex items-center gap-3">
              <span className="p-2 rounded-xl bg-amber-500/10 text-amber-600 dark:text-amber-400">
                <AlertCircle className="w-6 h-6" />
              </span>
              <div>
                <h3 className="text-base font-bold text-slate-900 dark:text-white">
                  Reset Master Definitions?
                </h3>
                <p className="text-xs text-slate-500 dark:text-slate-400">
                  Restore official university reference defaults
                </p>
              </div>
            </div>

            <p className="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
              This will reset all 9 entity registries (Programmes, Colleges, Castes, Categories, Religions, States, Countries, Boards, and Streams) back to official Kannur University default masters. Any custom unexported records will be overwritten.
            </p>

            <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100 dark:border-slate-700">
              <button
                onClick={() => setShowResetConfirmModal(false)}
                className="px-3.5 py-2 rounded-lg border border-slate-300 dark:border-slate-600 text-slate-700 dark:text-slate-300 hover:bg-slate-100 text-xs font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleConfirmReset}
                className="px-4 py-2 rounded-lg bg-amber-600 hover:bg-amber-700 text-white text-xs font-semibold shadow-xs"
              >
                Yes, Reset to University Defaults
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Guide Documentation Modal */}
      {showGuideModal && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in">
          <div className="bg-white dark:bg-slate-800 rounded-2xl max-w-3xl w-full border border-slate-200 dark:border-slate-700 shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
            <div className="p-4 sm:p-5 border-b border-slate-200 dark:border-slate-700 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <span className="p-1.5 rounded-lg bg-emerald-500/10 text-emerald-600 dark:text-emerald-400">
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
                  2. Side-by-Side Dual View &amp; Excel Ingestion
                </h4>
                <p>
                  Switching to <strong>"Side-by-Side View"</strong> displays the Programme Master and College Master registers concurrently on the same screen. You can also upload a single Excel worksheet containing both Programme columns on the left and College columns on the right (separated by an empty column or together). The tool auto-detects both headers simultaneously:
                </p>
                <div className="p-2.5 rounded-lg bg-slate-100 dark:bg-slate-900 font-mono text-[11px] overflow-x-auto mt-2 text-slate-700 dark:text-slate-300">
                  Programme Name (Reference) | ProgramCode | workFlowId | Program Name | [empty] | CollegeId | CollegeCode | College Name | CLG Name
                </div>
              </div>

              <div>
                <h4 className="text-sm font-bold text-slate-800 dark:text-slate-100 mb-1">
                  3. Interactive String Lookup Resolver &amp; Validator
                </h4>
                <p>
                  Use the live validator at the top to test any raw candidate string (e.g. <code>'Karnataka'</code>, <code>'Kannur IMPES'</code>, <code>'Music Dept'</code>). The simulator highlights exact matches, code matches, or keyword fallbacks instantly before you import student batches.
                </p>
              </div>

              <div>
                <h4 className="text-sm font-bold text-slate-800 dark:text-slate-100 mb-1">
                  4. Live Synchronization Across Report Studio
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
