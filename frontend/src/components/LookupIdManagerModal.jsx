import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  X,
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
  FileSpreadsheet
} from 'lucide-react';
import {
  LOOKUP_TAB_CONFIG,
  DEFAULT_LOOKUP_DATA,
  parseUploadedLookupSpreadsheet,
  saveLookupsToStorage
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

export default function LookupIdManagerModal({
  isOpen,
  onClose,
  lookups,
  onSaveLookups
}) {
  const [activeTab, setActiveTab] = useState('program');
  const [searchQuery, setSearchQuery] = useState('');

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

  // Edit states
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

  const [page, setPage] = useState(0);
  const [statusAlert, setStatusAlert] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);

  const fileInputRef = useRef(null);
  const pageSize = 20;

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
    const start = page * pageSize;
    return filteredList.slice(start, start + pageSize);
  }, [filteredList, page, pageSize]);

  const totalEntriesCount = useMemo(() => {
    if (!lookups) return 0;
    return Object.values(lookups).reduce((sum, list) => sum + (Array.isArray(list) ? list.length : 0), 0);
  }, [lookups]);

  const totalPages = Math.ceil(filteredList.length / pageSize) || 1;

  if (!isOpen) return null;

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
          message: 'Please fill all 4 fields for Programme (Reference, Code, workFlowId, Name).'
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
      onSaveLookups(updatedAll);
      saveLookupsToStorage(updatedAll);
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
      const code = newClgCode.trim();
      const name = newClgName.trim();
      const clgRef = newClgRef.trim();

      if (!id || !code || !name) {
        setStatusAlert({
          type: 'warning',
          message: 'Please fill CollegeId, CollegeCode, and College Name.'
        });
        return;
      }

      const exists = currentList.some(item => (item.id === id || (item.name || '').toLowerCase() === name.toLowerCase()));
      if (exists) {
        setStatusAlert({
          type: 'warning',
          message: `College with ID ${id} or name "${name}" already exists.`
        });
        return;
      }

      const updatedTabList = [{ id, code, name, clgRef: clgRef || name }, ...currentList];
      const updatedAll = { ...lookups, college: updatedTabList };
      onSaveLookups(updatedAll);
      saveLookupsToStorage(updatedAll);
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
        message: 'Please provide both Name and ID before adding.'
      });
      return;
    }

    const exists = currentList.some(item => (item.name || '').toLowerCase() === trimmedName.toLowerCase());
    if (exists) {
      setStatusAlert({
        type: 'warning',
        message: `An entry for "${trimmedName}" already exists in ${currentTabConfig.label}.`
      });
      return;
    }

    const updatedTabList = [{ name: trimmedName, id: trimmedId }, ...currentList];
    const updatedAll = {
      ...lookups,
      [activeTab]: updatedTabList
    };

    onSaveLookups(updatedAll);
    saveLookupsToStorage(updatedAll);
    setNewName('');
    setNewId('');
    setStatusAlert({
      type: 'success',
      message: `Added "${trimmedName}" (${trimmedId}) to ${currentTabConfig.label}.`
    });
  };

  const startEdit = (item, actualIndex) => {
    setEditingIdx(actualIndex);
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

  const cancelEdit = () => {
    setEditingIdx(null);
    setEditName('');
    setEditId('');
    setEditProgRef('');
    setEditProgCode('');
    setEditProgWorkflow('');
    setEditProgName('');
    setEditClgId('');
    setEditClgCode('');
    setEditClgName('');
    setEditClgRef('');
  };

  const saveEdit = (actualIndex) => {
    const updatedTabList = [...currentList];

    if (activeTab === 'program') {
      const pgmName = editProgRef.trim();
      const code = editProgCode.trim();
      const workflowId = editProgWorkflow.trim();
      const name = editProgName.trim();

      if (!pgmName || !code || !workflowId || !name) {
        setStatusAlert({
          type: 'warning',
          message: 'Programme fields cannot be blank.'
        });
        return;
      }

      updatedTabList[actualIndex] = { pgmName, code, workflowId, name };
      const updatedAll = { ...lookups, program: updatedTabList };
      onSaveLookups(updatedAll);
      saveLookupsToStorage(updatedAll);
      cancelEdit();
      setStatusAlert({
        type: 'success',
        message: `Updated Programme "${pgmName}" successfully.`
      });
      return;
    }

    if (activeTab === 'college') {
      const id = editClgId.trim();
      const code = editClgCode.trim();
      const name = editClgName.trim();
      const clgRef = editClgRef.trim();

      if (!id || !code || !name) {
        setStatusAlert({
          type: 'warning',
          message: 'College fields cannot be blank.'
        });
        return;
      }

      updatedTabList[actualIndex] = { id, code, name, clgRef: clgRef || name };
      const updatedAll = { ...lookups, college: updatedTabList };
      onSaveLookups(updatedAll);
      saveLookupsToStorage(updatedAll);
      cancelEdit();
      setStatusAlert({
        type: 'success',
        message: `Updated College "${name}" successfully.`
      });
      return;
    }

    // Simple edit
    const trimmedName = editName.trim();
    const trimmedId = editId.trim();

    if (!trimmedName || !trimmedId) {
      setStatusAlert({
        type: 'warning',
        message: 'Name and ID cannot be blank.'
      });
      return;
    }

    updatedTabList[actualIndex] = { name: trimmedName, id: trimmedId };
    const updatedAll = {
      ...lookups,
      [activeTab]: updatedTabList
    };

    onSaveLookups(updatedAll);
    saveLookupsToStorage(updatedAll);
    cancelEdit();
    setStatusAlert({
      type: 'success',
      message: `Updated "${trimmedName}" successfully.`
    });
  };

  const handleDelete = (actualIndex, title) => {
    const updatedTabList = currentList.filter((_, idx) => idx !== actualIndex);
    const updatedAll = {
      ...lookups,
      [activeTab]: updatedTabList
    };

    onSaveLookups(updatedAll);
    saveLookupsToStorage(updatedAll);
    setStatusAlert({
      type: 'info',
      message: `Removed "${title}" from ${currentTabConfig.label}.`
    });
  };

  const handleResetToDefaults = () => {
    if (window.confirm('Reset all defined IDs to official Kannur University defaults? Custom entries will be replaced.')) {
      onSaveLookups(DEFAULT_LOOKUP_DATA);
      saveLookupsToStorage(DEFAULT_LOOKUP_DATA);
      setEditingIdx(null);
      setStatusAlert({
        type: 'success',
        message: 'Reset all 9 master ID definitions to official Kannur University defaults.'
      });
    }
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setIsProcessing(true);
      setStatusAlert({ type: 'info', message: `Reading "${file.name}"...` });

      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const { imported, total } = parseUploadedLookupSpreadsheet(workbook, activeTab);

      if (total === 0) {
        setStatusAlert({
          type: 'warning',
          message: 'No matching lookup columns found. Supported headers: [Programme Name (Reference), ProgramCode, workFlowId, Program Name], [CollegeId, CollegeCode, College Name, CLG Name], or [Name, ID] pairs.'
        });
        setIsProcessing(false);
        if (fileInputRef.current) fileInputRef.current.value = '';
        return;
      }

      const merged = { ...lookups };
      const summaryMsgParts = [];

      for (const [key, items] of Object.entries(imported)) {
        if (!items || items.length === 0) continue;
        const existing = [...(merged[key] || [])];

        if (key === 'program') {
          const map = new Map(existing.map((x, idx) => [(x.pgmName || '').toLowerCase(), idx]));
          for (const item of items) {
            const k = (item.pgmName || '').toLowerCase();
            if (map.has(k)) {
              existing[map.get(k)] = { ...existing[map.get(k)], ...item };
            } else {
              existing.push(item);
              map.set(k, existing.length - 1);
            }
          }
        } else if (key === 'college') {
          const map = new Map(existing.map((x, idx) => [(x.name + '_' + x.id).toLowerCase(), idx]));
          for (const item of items) {
            const k = (item.name + '_' + item.id).toLowerCase();
            if (map.has(k)) {
              existing[map.get(k)] = { ...existing[map.get(k)], ...item };
            } else {
              existing.push(item);
              map.set(k, existing.length - 1);
            }
          }
        } else {
          const map = new Map(existing.map((x, idx) => [(x.name || '').toLowerCase(), idx]));
          for (const item of items) {
            const lower = (item.name || '').toLowerCase();
            if (map.has(lower)) {
              existing[map.get(lower)] = { ...existing[map.get(lower)], id: item.id };
            } else {
              existing.push(item);
              map.set(lower, existing.length - 1);
            }
          }
        }

        merged[key] = existing;
        const tabConf = LOOKUP_TAB_CONFIG.find(t => t.key === key);
        summaryMsgParts.push(`${items.length} ${tabConf ? tabConf.label : key}`);
      }

      onSaveLookups(merged);
      saveLookupsToStorage(merged);

      setStatusAlert({
        type: 'success',
        message: `Successfully imported ${total} entries from "${file.name}": ${summaryMsgParts.join(', ')}.`
      });
    } catch (err) {
      console.error('Lookup upload error:', err);
      setStatusAlert({
        type: 'error',
        message: `Failed to import spreadsheet: ${err.message}`
      });
    } finally {
      setIsProcessing(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleExportAllToExcel = () => {
    try {
      const wb = XLSX.utils.book_new();

      LOOKUP_TAB_CONFIG.forEach(cfg => {
        const list = lookups[cfg.key] || [];
        let rows = [];

        if (cfg.key === 'program') {
          rows = list.map(item => ({
            'Programme Name (Reference)': item.pgmName || '',
            'ProgramCode': item.code || '',
            'workFlowId': item.workflowId || '',
            'Program Name': item.name || ''
          }));
        } else if (cfg.key === 'college') {
          rows = list.map(item => ({
            'CollegeId': item.id || '',
            'CollegeCode': item.code || '',
            'College Name': item.name || '',
            'CLG Name': item.clgRef || ''
          }));
        } else {
          rows = list.map(item => ({
            [cfg.nameHeader]: item.name || '',
            [cfg.idHeader]: item.id || ''
          }));
        }

        const ws = XLSX.utils.json_to_sheet(rows.length > 0 ? rows : [{}]);
        XLSX.utils.book_append_sheet(wb, ws, cfg.label);
      });

      XLSX.writeFile(wb, 'Defined_Admission_Master_Lookups.xlsx');
      setStatusAlert({
        type: 'success',
        message: 'Downloaded Defined_Admission_Master_Lookups.xlsx with all 9 master entity definitions.'
      });
    } catch (err) {
      console.error('Export error:', err);
      setStatusAlert({
        type: 'error',
        message: `Failed to export workbook: ${err.message}`
      });
    }
  };

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'rgba(0, 0, 0, 0.65)',
        backdropFilter: 'blur(4px)',
        WebkitBackdropFilter: 'blur(4px)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 2000,
        padding: '20px'
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        style={{
          backgroundColor: 'var(--panel)',
          border: '1px solid var(--line)',
          borderRadius: '12px',
          maxWidth: '1080px',
          width: '100%',
          maxHeight: '92vh',
          display: 'flex',
          flexDirection: 'column',
          overflow: 'hidden',
          boxShadow: '0 20px 50px rgba(0, 0, 0, 0.35)',
          color: 'var(--ink)'
        }}
      >
        {/* Top Header */}
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            padding: '16px 24px',
            borderBottom: '1px solid var(--line)',
            background: 'var(--bg)'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div
              style={{
                width: '38px',
                height: '38px',
                borderRadius: '8px',
                background: 'var(--accent-soft)',
                color: 'var(--accent)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}
            >
              <Database size={20} />
            </div>
            <div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <h3 style={{ margin: 0, fontSize: '17px', fontWeight: 700, color: 'var(--ink)' }}>
                  Define & Manage ID Lookups
                </h3>
                <span
                  style={{
                    fontSize: '11px',
                    fontWeight: 700,
                    padding: '2px 8px',
                    borderRadius: '12px',
                    background: 'var(--accent-soft)',
                    color: 'var(--accent)',
                    border: '1px solid var(--line)'
                  }}
                >
                  {totalEntriesCount} Total Defined
                </span>
              </div>
              <p style={{ margin: '3px 0 0', fontSize: '12px', color: 'var(--muted)' }}>
                Configure codes for Programmes, Colleges, Category, Caste, Religion, Country, State, Qualification, and Stream.
              </p>
            </div>
          </div>

          {/* Header Action Buttons */}
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <input
              type="file"
              ref={fileInputRef}
              onChange={handleFileUpload}
              accept=".xlsx,.xls,.csv"
              style={{ display: 'none' }}
            />
            <a
              href="/Admission_ID_Upload_Template.xlsx"
              download="Admission_ID_Upload_Template.xlsx"
              title="Download pre-formatted Excel template with sample IDs for all master entities"
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                padding: '7px 12px',
                fontSize: '12px',
                fontWeight: 600,
                background: 'var(--panel)',
                color: 'var(--ink)',
                border: '1px solid var(--line)',
                borderRadius: '6px',
                textDecoration: 'none'
              }}
            >
              <FileSpreadsheet size={14} color="var(--accent)" />
              <span>Download Template (.xlsx)</span>
            </a>

            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              disabled={isProcessing}
              title="Upload spreadsheet with lookup ID definitions"
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                padding: '7px 12px',
                fontSize: '12px',
                fontWeight: 600,
                background: 'var(--panel)',
                color: 'var(--ink)',
                border: '1px solid var(--line)',
                borderRadius: '6px'
              }}
            >
              <Upload size={14} color="var(--accent)" />
              <span>{isProcessing ? 'Importing...' : 'Upload IDs (.xlsx / .csv)'}</span>
            </button>

            <button
              type="button"
              onClick={handleExportAllToExcel}
              title="Export all defined IDs to Excel"
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                padding: '7px 12px',
                fontSize: '12px',
                fontWeight: 600,
                background: 'var(--panel)',
                color: 'var(--ink)',
                border: '1px solid var(--line)',
                borderRadius: '6px'
              }}
            >
              <Download size={14} color="#16a34a" />
              <span>Export All</span>
            </button>

            <button
              type="button"
              onClick={handleResetToDefaults}
              title="Reset to official Kannur University definitions"
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                padding: '7px 12px',
                fontSize: '12px',
                fontWeight: 600,
                background: 'var(--panel)',
                color: 'var(--danger)',
                border: '1px solid var(--line)',
                borderRadius: '6px'
              }}
            >
              <RotateCcw size={14} />
              <span>Reset</span>
            </button>

            <button
              type="button"
              onClick={onClose}
              title="Close Modal"
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                width: '32px',
                height: '32px',
                padding: 0,
                background: 'transparent',
                color: 'var(--muted)',
                border: 'none',
                borderRadius: '6px',
                cursor: 'pointer'
              }}
            >
              <X size={18} />
            </button>
          </div>
        </div>

        {/* Status Alert Banner */}
        {statusAlert && (
          <div
            style={{
              padding: '10px 24px',
              fontSize: '12px',
              fontWeight: 500,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              borderBottom: '1px solid var(--line)',
              background:
                statusAlert.type === 'error'
                  ? 'rgba(180, 35, 24, 0.12)'
                  : statusAlert.type === 'warning'
                  ? 'rgba(217, 119, 6, 0.12)'
                  : statusAlert.type === 'success'
                  ? 'rgba(22, 163, 74, 0.12)'
                  : 'var(--accent-soft)',
              color:
                statusAlert.type === 'error'
                  ? 'var(--danger)'
                  : statusAlert.type === 'warning'
                  ? '#b45309'
                  : statusAlert.type === 'success'
                  ? '#15803d'
                  : 'var(--accent)'
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
              {statusAlert.type === 'success' ? (
                <CheckCircle2 size={16} />
              ) : (
                <AlertCircle size={16} />
              )}
              <span>{statusAlert.message}</span>
            </div>
            <button
              type="button"
              onClick={() => setStatusAlert(null)}
              style={{
                background: 'transparent',
                border: 'none',
                color: 'inherit',
                cursor: 'pointer',
                padding: '2px',
                display: 'flex'
              }}
            >
              <X size={14} />
            </button>
          </div>
        )}

        {/* 9 Tabs Navigation Bar */}
        <div
          style={{
            display: 'flex',
            gap: '6px',
            padding: '10px 24px 0',
            borderBottom: '1px solid var(--line)',
            background: 'var(--bg)',
            overflowX: 'auto'
          }}
        >
          {LOOKUP_TAB_CONFIG.map(tab => {
            const Icon = TAB_ICONS[tab.key] || Tag;
            const count = (lookups && Array.isArray(lookups[tab.key])) ? lookups[tab.key].length : 0;
            const isActive = activeTab === tab.key;

            return (
              <button
                key={tab.key}
                type="button"
                onClick={() => handleTabChange(tab.key)}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '7px',
                  padding: '9px 14px',
                  fontSize: '12px',
                  fontWeight: isActive ? 700 : 500,
                  border: 'none',
                  borderBottom: isActive ? '2px solid var(--accent)' : '2px solid transparent',
                  borderRadius: '6px 6px 0 0',
                  background: isActive ? 'var(--panel)' : 'transparent',
                  color: isActive ? 'var(--accent)' : 'var(--muted)',
                  cursor: 'pointer',
                  whiteSpace: 'nowrap',
                  boxShadow: 'none'
                }}
              >
                <Icon size={14} color={isActive ? 'var(--accent)' : 'var(--muted)'} />
                <span>{tab.label}</span>
                <span
                  style={{
                    fontSize: '10.5px',
                    fontWeight: 700,
                    padding: '1px 6px',
                    borderRadius: '10px',
                    background: isActive ? 'var(--accent-soft)' : 'var(--line)',
                    color: isActive ? 'var(--accent)' : 'var(--muted)'
                  }}
                >
                  {count}
                </span>
              </button>
            );
          })}
        </div>

        {/* Sub-header with column mapping explanation */}
        <div
          style={{
            padding: '8px 24px',
            background: 'var(--bg)',
            borderBottom: '1px solid var(--line)',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            fontSize: '11.5px',
            color: 'var(--muted)',
            flexWrap: 'wrap',
            gap: '8px'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flexWrap: 'wrap' }}>
            <span style={{ fontWeight: 600, color: 'var(--ink)' }}>Output Column Mapping:</span>
            {activeTab === 'program' ? (
              <>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  Programme Name (Reference)
                </code>
                <span>↔</span>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  ProgramCode
                </code>
                <span>↔</span>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  workFlowId
                </code>
                <span>↔</span>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  Program Name
                </code>
              </>
            ) : activeTab === 'college' ? (
              <>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  CollegeId
                </code>
                <span>↔</span>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  CollegeCode
                </code>
                <span>↔</span>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  College Name
                </code>
                <span>↔</span>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  CLG Name
                </code>
              </>
            ) : (
              <>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  {currentTabConfig.nameHeader}
                </code>
                <span>↔</span>
                <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                  {currentTabConfig.idHeader}
                </code>
              </>
            )}
            <span style={{ marginLeft: '6px' }}>— {currentTabConfig.description}</span>
          </div>
          <div>
            Showing <strong>{filteredList.length}</strong> of <strong>{currentList.length}</strong> defined
          </div>
        </div>

        {/* Action Toolbar: Live Search & Add Entry */}
        <div
          style={{
            padding: '12px 24px',
            borderBottom: '1px solid var(--line)',
            background: 'var(--panel)',
            display: 'flex',
            gap: '12px',
            alignItems: 'center',
            flexWrap: 'wrap'
          }}
        >
          {/* Live Search */}
          <div style={{ position: 'relative', flex: '1', minWidth: '200px' }}>
            <Search
              size={14}
              color="var(--muted)"
              style={{ position: 'absolute', left: '10px', top: '10px' }}
            />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                setPage(0);
              }}
              placeholder={`Search ${currentTabConfig.label}...`}
              style={{
                width: '100%',
                padding: '7px 10px 7px 32px',
                fontSize: '12px',
                borderRadius: '6px',
                border: '1px solid var(--line)',
                background: 'var(--bg)',
                color: 'var(--ink)',
                outline: 'none'
              }}
            />
          </div>

          {/* Quick Add Entry Form: Tailored per tab type */}
          {activeTab === 'program' ? (
            <form
              onSubmit={handleAdd}
              style={{ display: 'flex', gap: '6px', alignItems: 'center', flexWrap: 'wrap' }}
            >
              <input
                type="text"
                value={newProgRef}
                onChange={(e) => setNewProgRef(e.target.value)}
                placeholder="Programme Name (Reference)"
                style={{
                  width: '180px',
                  padding: '7px 9px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <input
                type="text"
                value={newProgCode}
                onChange={(e) => setNewProgCode(e.target.value)}
                placeholder="ProgramCode (e.g. PFPSYGS25)"
                style={{
                  width: '140px',
                  padding: '7px 9px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <input
                type="text"
                value={newProgWorkflow}
                onChange={(e) => setNewProgWorkflow(e.target.value)}
                placeholder="workFlowId (e.g. 151)"
                style={{
                  width: '110px',
                  padding: '7px 9px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <input
                type="text"
                value={newProgName}
                onChange={(e) => setNewProgName(e.target.value)}
                placeholder="Full Program Name"
                style={{
                  width: '180px',
                  padding: '7px 9px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <button
                type="submit"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '5px',
                  padding: '7px 14px',
                  fontSize: '12px',
                  fontWeight: 600,
                  background: 'var(--accent)',
                  color: 'white',
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer',
                  whiteSpace: 'nowrap'
                }}
              >
                <Plus size={14} />
                <span>Add Programme</span>
              </button>
            </form>
          ) : activeTab === 'college' ? (
            <form
              onSubmit={handleAdd}
              style={{ display: 'flex', gap: '6px', alignItems: 'center', flexWrap: 'wrap' }}
            >
              <input
                type="text"
                value={newClgId}
                onChange={(e) => setNewClgId(e.target.value)}
                placeholder="CollegeId (e.g. 435)"
                style={{
                  width: '100px',
                  padding: '7px 9px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <input
                type="text"
                value={newClgCode}
                onChange={(e) => setNewClgCode(e.target.value)}
                placeholder="CollegeCode"
                style={{
                  width: '100px',
                  padding: '7px 9px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <input
                type="text"
                value={newClgName}
                onChange={(e) => setNewClgName(e.target.value)}
                placeholder="College Name"
                style={{
                  width: '190px',
                  padding: '7px 9px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <input
                type="text"
                value={newClgRef}
                onChange={(e) => setNewClgRef(e.target.value)}
                placeholder="CLG Name (Address/Campus)"
                style={{
                  width: '190px',
                  padding: '7px 9px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <button
                type="submit"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '5px',
                  padding: '7px 14px',
                  fontSize: '12px',
                  fontWeight: 600,
                  background: 'var(--accent)',
                  color: 'white',
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer',
                  whiteSpace: 'nowrap'
                }}
              >
                <Plus size={14} />
                <span>Add College</span>
              </button>
            </form>
          ) : (
            <form
              onSubmit={handleAdd}
              style={{ display: 'flex', gap: '8px', alignItems: 'center' }}
            >
              <input
                type="text"
                value={newName}
                onChange={(e) => setNewName(e.target.value)}
                placeholder={`New ${currentTabConfig.nameLabel}`}
                style={{
                  width: '210px',
                  padding: '7px 10px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <input
                type="text"
                value={newId}
                onChange={(e) => setNewId(e.target.value)}
                placeholder="ID"
                style={{
                  width: '95px',
                  padding: '7px 10px',
                  fontSize: '12px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  outline: 'none'
                }}
              />
              <button
                type="submit"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '5px',
                  padding: '7px 14px',
                  fontSize: '12px',
                  fontWeight: 600,
                  background: 'var(--accent)',
                  color: 'white',
                  border: 'none',
                  borderRadius: '6px',
                  cursor: 'pointer',
                  whiteSpace: 'nowrap'
                }}
              >
                <Plus size={14} />
                <span>Add Entry</span>
              </button>
            </form>
          )}
        </div>

        {/* Scrollable Table Content */}
        <div style={{ flex: 1, overflowY: 'auto', minHeight: '260px', maxHeight: '46vh' }}>
          {filteredList.length === 0 ? (
            <div style={{ padding: '40px 20px', textAlign: 'center', color: 'var(--muted)' }}>
              <Database size={36} color="var(--line)" style={{ margin: '0 auto 8px' }} />
              <p style={{ margin: 0, fontSize: '13px', fontWeight: 600 }}>
                {searchQuery ? `No entries matching "${searchQuery}"` : `No ${currentTabConfig.label} defined yet.`}
              </p>
              <span style={{ fontSize: '11.5px', marginTop: '4px', display: 'inline-block' }}>
                Use the inputs above to add an entry, or upload an Excel/CSV file to bulk import.
              </span>
            </div>
          ) : activeTab === 'program' ? (
            /* Programme Table */
            <table
              style={{
                width: '100%',
                borderCollapse: 'collapse',
                textAlign: 'left',
                fontSize: '12px'
              }}
            >
              <thead>
                <tr
                  style={{
                    position: 'sticky',
                    top: 0,
                    background: 'var(--bg)',
                    borderBottom: '2px solid var(--line)',
                    color: 'var(--muted)',
                    fontSize: '11px',
                    textTransform: 'uppercase',
                    fontWeight: 700,
                    zIndex: 5
                  }}
                >
                  <th style={{ padding: '8px 14px', width: '40px', textAlign: 'center' }}>#</th>
                  <th style={{ padding: '8px 14px', width: '220px' }}>Programme Name (Reference)</th>
                  <th style={{ padding: '8px 14px', width: '120px' }}>ProgramCode</th>
                  <th style={{ padding: '8px 14px', width: '90px' }}>workFlowId</th>
                  <th style={{ padding: '8px 14px' }}>Program Name</th>
                  <th style={{ padding: '8px 14px', width: '90px', textAlign: 'right' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {paginatedList.map((item, idx) => {
                  const actualIndex = currentList.indexOf(item);
                  const isEditing = editingIdx === actualIndex;

                  return (
                    <tr
                      key={`${item.code}-${item.workflowId}-${actualIndex}`}
                      style={{
                        borderBottom: '1px solid var(--line)',
                        background: isEditing ? 'var(--accent-soft)' : (idx % 2 === 0 ? 'var(--panel)' : 'var(--bg)')
                      }}
                    >
                      <td style={{ padding: '8px 14px', textAlign: 'center', color: 'var(--muted)', fontFamily: 'monospace', fontSize: '11px' }}>
                        {page * pageSize + idx + 1}
                      </td>

                      {/* Reference Name */}
                      <td style={{ padding: '8px 14px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editProgRef}
                            onChange={(e) => setEditProgRef(e.target.value)}
                            style={{ width: '100%', padding: '4px 6px', fontSize: '11.5px', borderRadius: '4px', border: '1px solid var(--accent)', background: 'var(--panel)', color: 'var(--ink)' }}
                          />
                        ) : (
                          <span style={{ fontWeight: 600, color: 'var(--ink)' }}>{item.pgmName}</span>
                        )}
                      </td>

                      {/* ProgramCode */}
                      <td style={{ padding: '8px 14px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editProgCode}
                            onChange={(e) => setEditProgCode(e.target.value)}
                            style={{ width: '100%', padding: '4px 6px', fontSize: '11.5px', fontFamily: 'monospace', borderRadius: '4px', border: '1px solid var(--accent)', background: 'var(--panel)', color: 'var(--ink)' }}
                          />
                        ) : (
                          <span style={{ padding: '2px 6px', borderRadius: '4px', fontFamily: 'monospace', fontWeight: 700, fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', border: '1px solid var(--line)', display: 'inline-block' }}>
                            {item.code}
                          </span>
                        )}
                      </td>

                      {/* workFlowId */}
                      <td style={{ padding: '8px 14px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editProgWorkflow}
                            onChange={(e) => setEditProgWorkflow(e.target.value)}
                            style={{ width: '100%', padding: '4px 6px', fontSize: '11.5px', fontFamily: 'monospace', borderRadius: '4px', border: '1px solid var(--accent)', background: 'var(--panel)', color: 'var(--ink)' }}
                          />
                        ) : (
                          <span style={{ padding: '2px 6px', borderRadius: '4px', fontFamily: 'monospace', fontWeight: 700, fontSize: '11px', background: 'rgba(22, 163, 74, 0.1)', color: '#16a34a', border: '1px solid rgba(22, 163, 74, 0.25)', display: 'inline-block' }}>
                            {item.workflowId}
                          </span>
                        )}
                      </td>

                      {/* Full Program Name */}
                      <td style={{ padding: '8px 14px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editProgName}
                            onChange={(e) => setEditProgName(e.target.value)}
                            style={{ width: '100%', padding: '4px 6px', fontSize: '11.5px', borderRadius: '4px', border: '1px solid var(--accent)', background: 'var(--panel)', color: 'var(--ink)' }}
                          />
                        ) : (
                          <span style={{ color: 'var(--muted)', fontSize: '11.5px' }}>{item.name}</span>
                        )}
                      </td>

                      {/* Actions */}
                      <td style={{ padding: '8px 14px', textAlign: 'right' }}>
                        {isEditing ? (
                          <div style={{ display: 'flex', gap: '4px', justifyContent: 'flex-end' }}>
                            <button
                              type="button"
                              onClick={() => saveEdit(actualIndex)}
                              title="Save changes"
                              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '26px', height: '26px', padding: 0, background: '#16a34a', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
                            >
                              <Check size={13} />
                            </button>
                            <button
                              type="button"
                              onClick={cancelEdit}
                              title="Cancel"
                              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '26px', height: '26px', padding: 0, background: 'var(--panel)', color: 'var(--muted)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer' }}
                            >
                              <X size={13} />
                            </button>
                          </div>
                        ) : (
                          <div style={{ display: 'flex', gap: '4px', justifyContent: 'flex-end' }}>
                            <button
                              type="button"
                              onClick={() => startEdit(item, actualIndex)}
                              title="Edit"
                              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '26px', height: '26px', padding: 0, background: 'var(--panel)', color: 'var(--accent)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer' }}
                            >
                              <Edit2 size={12} />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleDelete(actualIndex, item.pgmName)}
                              title="Delete"
                              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '26px', height: '26px', padding: 0, background: 'var(--panel)', color: 'var(--danger)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer' }}
                            >
                              <Trash2 size={12} />
                            </button>
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          ) : activeTab === 'college' ? (
            /* College Table */
            <table
              style={{
                width: '100%',
                borderCollapse: 'collapse',
                textAlign: 'left',
                fontSize: '12px'
              }}
            >
              <thead>
                <tr
                  style={{
                    position: 'sticky',
                    top: 0,
                    background: 'var(--bg)',
                    borderBottom: '2px solid var(--line)',
                    color: 'var(--muted)',
                    fontSize: '11px',
                    textTransform: 'uppercase',
                    fontWeight: 700,
                    zIndex: 5
                  }}
                >
                  <th style={{ padding: '8px 14px', width: '40px', textAlign: 'center' }}>#</th>
                  <th style={{ padding: '8px 14px', width: '90px' }}>CollegeId</th>
                  <th style={{ padding: '8px 14px', width: '90px' }}>CollegeCode</th>
                  <th style={{ padding: '8px 14px', width: '230px' }}>College Name</th>
                  <th style={{ padding: '8px 14px' }}>CLG Name (Address/Campus)</th>
                  <th style={{ padding: '8px 14px', width: '90px', textAlign: 'right' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {paginatedList.map((item, idx) => {
                  const actualIndex = currentList.indexOf(item);
                  const isEditing = editingIdx === actualIndex;

                  return (
                    <tr
                      key={`${item.id}-${item.code}-${actualIndex}`}
                      style={{
                        borderBottom: '1px solid var(--line)',
                        background: isEditing ? 'var(--accent-soft)' : (idx % 2 === 0 ? 'var(--panel)' : 'var(--bg)')
                      }}
                    >
                      <td style={{ padding: '8px 14px', textAlign: 'center', color: 'var(--muted)', fontFamily: 'monospace', fontSize: '11px' }}>
                        {page * pageSize + idx + 1}
                      </td>

                      {/* CollegeId */}
                      <td style={{ padding: '8px 14px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editClgId}
                            onChange={(e) => setEditClgId(e.target.value)}
                            style={{ width: '100%', padding: '4px 6px', fontSize: '11.5px', fontFamily: 'monospace', borderRadius: '4px', border: '1px solid var(--accent)', background: 'var(--panel)', color: 'var(--ink)' }}
                          />
                        ) : (
                          <span style={{ padding: '2px 6px', borderRadius: '4px', fontFamily: 'monospace', fontWeight: 700, fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', border: '1px solid var(--line)', display: 'inline-block' }}>
                            {item.id}
                          </span>
                        )}
                      </td>

                      {/* CollegeCode */}
                      <td style={{ padding: '8px 14px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editClgCode}
                            onChange={(e) => setEditClgCode(e.target.value)}
                            style={{ width: '100%', padding: '4px 6px', fontSize: '11.5px', fontFamily: 'monospace', borderRadius: '4px', border: '1px solid var(--accent)', background: 'var(--panel)', color: 'var(--ink)' }}
                          />
                        ) : (
                          <span style={{ padding: '2px 6px', borderRadius: '4px', fontFamily: 'monospace', fontWeight: 700, fontSize: '11px', background: 'rgba(22, 163, 74, 0.1)', color: '#16a34a', border: '1px solid rgba(22, 163, 74, 0.25)', display: 'inline-block' }}>
                            {item.code}
                          </span>
                        )}
                      </td>

                      {/* College Name */}
                      <td style={{ padding: '8px 14px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editClgName}
                            onChange={(e) => setEditClgName(e.target.value)}
                            style={{ width: '100%', padding: '4px 6px', fontSize: '11.5px', borderRadius: '4px', border: '1px solid var(--accent)', background: 'var(--panel)', color: 'var(--ink)' }}
                          />
                        ) : (
                          <span style={{ fontWeight: 600, color: 'var(--ink)' }}>{item.name}</span>
                        )}
                      </td>

                      {/* CLG Name */}
                      <td style={{ padding: '8px 14px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editClgRef}
                            onChange={(e) => setEditClgRef(e.target.value)}
                            style={{ width: '100%', padding: '4px 6px', fontSize: '11.5px', borderRadius: '4px', border: '1px solid var(--accent)', background: 'var(--panel)', color: 'var(--ink)' }}
                          />
                        ) : (
                          <span style={{ color: 'var(--muted)', fontSize: '11.5px' }}>{item.clgRef || item.name}</span>
                        )}
                      </td>

                      {/* Actions */}
                      <td style={{ padding: '8px 14px', textAlign: 'right' }}>
                        {isEditing ? (
                          <div style={{ display: 'flex', gap: '4px', justifyContent: 'flex-end' }}>
                            <button
                              type="button"
                              onClick={() => saveEdit(actualIndex)}
                              title="Save changes"
                              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '26px', height: '26px', padding: 0, background: '#16a34a', color: 'white', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
                            >
                              <Check size={13} />
                            </button>
                            <button
                              type="button"
                              onClick={cancelEdit}
                              title="Cancel"
                              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '26px', height: '26px', padding: 0, background: 'var(--panel)', color: 'var(--muted)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer' }}
                            >
                              <X size={13} />
                            </button>
                          </div>
                        ) : (
                          <div style={{ display: 'flex', gap: '4px', justifyContent: 'flex-end' }}>
                            <button
                              type="button"
                              onClick={() => startEdit(item, actualIndex)}
                              title="Edit"
                              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '26px', height: '26px', padding: 0, background: 'var(--panel)', color: 'var(--accent)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer' }}
                            >
                              <Edit2 size={12} />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleDelete(actualIndex, item.name)}
                              title="Delete"
                              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '26px', height: '26px', padding: 0, background: 'var(--panel)', color: 'var(--danger)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer' }}
                            >
                              <Trash2 size={12} />
                            </button>
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          ) : (
            /* Simple 2-column Table */
            <table
              style={{
                width: '100%',
                borderCollapse: 'collapse',
                textAlign: 'left',
                fontSize: '12px'
              }}
            >
              <thead>
                <tr
                  style={{
                    position: 'sticky',
                    top: 0,
                    background: 'var(--bg)',
                    borderBottom: '2px solid var(--line)',
                    color: 'var(--muted)',
                    fontSize: '11px',
                    textTransform: 'uppercase',
                    fontWeight: 700,
                    zIndex: 5
                  }}
                >
                  <th style={{ padding: '8px 16px', width: '50px', textAlign: 'center' }}>#</th>
                  <th style={{ padding: '8px 16px' }}>{currentTabConfig.nameLabel} ({currentTabConfig.nameHeader})</th>
                  <th style={{ padding: '8px 16px', width: '180px' }}>{currentTabConfig.idLabel} ({currentTabConfig.idHeader})</th>
                  <th style={{ padding: '8px 16px', width: '110px', textAlign: 'right' }}>Actions</th>
                </tr>
              </thead>
              <tbody>
                {paginatedList.map((item, idx) => {
                  const actualIndex = currentList.indexOf(item);
                  const isEditing = editingIdx === actualIndex;

                  return (
                    <tr
                      key={`${item.name}-${item.id}-${actualIndex}`}
                      style={{
                        borderBottom: '1px solid var(--line)',
                        background: isEditing ? 'var(--accent-soft)' : (idx % 2 === 0 ? 'var(--panel)' : 'var(--bg)')
                      }}
                    >
                      <td style={{ padding: '8px 16px', textAlign: 'center', color: 'var(--muted)', fontFamily: 'monospace', fontSize: '11px' }}>
                        {page * pageSize + idx + 1}
                      </td>

                      {/* Name Cell */}
                      <td style={{ padding: '8px 16px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editName}
                            onChange={(e) => setEditName(e.target.value)}
                            style={{
                              width: '100%',
                              padding: '4px 8px',
                              fontSize: '12px',
                              borderRadius: '4px',
                              border: '1px solid var(--accent)',
                              background: 'var(--panel)',
                              color: 'var(--ink)',
                              outline: 'none'
                            }}
                            autoFocus
                          />
                        ) : (
                          <span style={{ fontWeight: 600, color: 'var(--ink)' }}>
                            {item.name}
                          </span>
                        )}
                      </td>

                      {/* ID Cell */}
                      <td style={{ padding: '8px 16px' }}>
                        {isEditing ? (
                          <input
                            type="text"
                            value={editId}
                            onChange={(e) => setEditId(e.target.value)}
                            style={{
                              width: '120px',
                              padding: '4px 8px',
                              fontSize: '12px',
                              fontFamily: 'monospace',
                              borderRadius: '4px',
                              border: '1px solid var(--accent)',
                              background: 'var(--panel)',
                              color: 'var(--ink)',
                              outline: 'none'
                            }}
                          />
                        ) : (
                          <span
                            style={{
                              padding: '3px 8px',
                              borderRadius: '4px',
                              fontFamily: 'monospace',
                              fontWeight: 700,
                              fontSize: '11.5px',
                              background: 'var(--accent-soft)',
                              color: 'var(--accent)',
                              border: '1px solid var(--line)',
                              display: 'inline-block'
                            }}
                          >
                            {item.id}
                          </span>
                        )}
                      </td>

                      {/* Actions */}
                      <td style={{ padding: '8px 16px', textAlign: 'right' }}>
                        {isEditing ? (
                          <div style={{ display: 'flex', gap: '4px', justifyContent: 'flex-end' }}>
                            <button
                              type="button"
                              onClick={() => saveEdit(actualIndex)}
                              title="Save changes"
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                width: '28px',
                                height: '28px',
                                padding: 0,
                                background: '#16a34a',
                                color: 'white',
                                border: 'none',
                                borderRadius: '4px',
                                cursor: 'pointer'
                              }}
                            >
                              <Check size={14} />
                            </button>
                            <button
                              type="button"
                              onClick={cancelEdit}
                              title="Cancel"
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                width: '28px',
                                height: '28px',
                                padding: 0,
                                background: 'var(--panel)',
                                color: 'var(--muted)',
                                border: '1px solid var(--line)',
                                borderRadius: '4px',
                                cursor: 'pointer'
                              }}
                            >
                              <X size={14} />
                            </button>
                          </div>
                        ) : (
                          <div style={{ display: 'flex', gap: '4px', justifyContent: 'flex-end' }}>
                            <button
                              type="button"
                              onClick={() => startEdit(item, actualIndex)}
                              title="Edit"
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                width: '28px',
                                height: '28px',
                                padding: 0,
                                background: 'var(--panel)',
                                color: 'var(--accent)',
                                border: '1px solid var(--line)',
                                borderRadius: '4px',
                                cursor: 'pointer'
                              }}
                            >
                              <Edit2 size={13} />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleDelete(actualIndex, item.name)}
                              title="Delete"
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                width: '28px',
                                height: '28px',
                                padding: 0,
                                background: 'var(--panel)',
                                color: 'var(--danger)',
                                border: '1px solid var(--line)',
                                borderRadius: '4px',
                                cursor: 'pointer'
                              }}
                            >
                              <Trash2 size={13} />
                            </button>
                          </div>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>

        {/* Footer with Pagination & Done button */}
        <div
          style={{
            padding: '12px 24px',
            borderTop: '1px solid var(--line)',
            background: 'var(--bg)',
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px', color: 'var(--muted)' }}>
            <span>
              Page {page + 1} of {totalPages}
            </span>
            <button
              type="button"
              onClick={() => setPage(p => Math.max(0, p - 1))}
              disabled={page === 0}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                width: '26px',
                height: '26px',
                padding: 0,
                background: 'var(--panel)',
                color: 'var(--ink)',
                border: '1px solid var(--line)',
                borderRadius: '4px',
                cursor: page === 0 ? 'not-allowed' : 'pointer',
                opacity: page === 0 ? 0.4 : 1
              }}
            >
              <ChevronLeft size={14} />
            </button>
            <button
              type="button"
              onClick={() => setPage(p => Math.min(totalPages - 1, p + 1))}
              disabled={page >= totalPages - 1}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                width: '26px',
                height: '26px',
                padding: 0,
                background: 'var(--panel)',
                color: 'var(--ink)',
                border: '1px solid var(--line)',
                borderRadius: '4px',
                cursor: page >= totalPages - 1 ? 'not-allowed' : 'pointer',
                opacity: page >= totalPages - 1 ? 0.4 : 1
              }}
            >
              <ChevronRight size={14} />
            </button>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <span style={{ fontSize: '11px', color: 'var(--muted)' }}>
              All additions and edits save automatically to browser storage
            </span>
            <button
              type="button"
              onClick={onClose}
              style={{
                padding: '7px 18px',
                fontSize: '12px',
                fontWeight: 600,
                background: 'var(--accent)',
                color: 'white',
                border: 'none',
                borderRadius: '6px',
                cursor: 'pointer'
              }}
            >
              Done / Apply to Import
            </button>
          </div>
        </div>

      </div>
    </div>
  );
}
