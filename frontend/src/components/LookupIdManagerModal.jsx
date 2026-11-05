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
  ChevronLeft,
  ChevronRight,
  Database
} from 'lucide-react';
import {
  LOOKUP_TAB_CONFIG,
  DEFAULT_LOOKUP_DATA,
  parseUploadedLookupSpreadsheet,
  saveLookupsToStorage
} from '../data/admissionLookups';

const TAB_ICONS = {
  category: Tag,
  caste: Users,
  religion: Heart,
  country: Globe,
  state: MapPin,
  qualification: GraduationCap,
  stream: BookOpen
};

export default function LookupIdManagerModal({
  isOpen,
  onClose,
  lookups,
  onSaveLookups
}) {
  const [activeTab, setActiveTab] = useState('caste');
  const [searchQuery, setSearchQuery] = useState('');
  const [newName, setNewName] = useState('');
  const [newId, setNewId] = useState('');
  const [editingIdx, setEditingIdx] = useState(null);
  const [editName, setEditName] = useState('');
  const [editId, setEditId] = useState('');
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
    return currentList.filter(item => {
      const n = String(item.name || '').toLowerCase();
      const id = String(item.id || '').toLowerCase();
      return n.includes(q) || id.includes(q);
    });
  }, [currentList, searchQuery]);

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
    setStatusAlert(null);
  };

  const handleAdd = (e) => {
    e?.preventDefault();
    const trimmedName = newName.trim();
    const trimmedId = newId.trim();

    if (!trimmedName || !trimmedId) {
      setStatusAlert({
        type: 'warning',
        message: 'Please provide both Name and ID before adding.'
      });
      return;
    }

    const exists = currentList.some(item => item.name.toLowerCase() === trimmedName.toLowerCase());
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
    setEditName(item.name);
    setEditId(item.id);
  };

  const cancelEdit = () => {
    setEditingIdx(null);
    setEditName('');
    setEditId('');
  };

  const saveEdit = (actualIndex) => {
    const trimmedName = editName.trim();
    const trimmedId = editId.trim();

    if (!trimmedName || !trimmedId) {
      setStatusAlert({
        type: 'warning',
        message: 'Name and ID cannot be blank.'
      });
      return;
    }

    const updatedTabList = [...currentList];
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

  const handleDelete = (actualIndex, itemName) => {
    const updatedTabList = currentList.filter((_, idx) => idx !== actualIndex);
    const updatedAll = {
      ...lookups,
      [activeTab]: updatedTabList
    };

    onSaveLookups(updatedAll);
    saveLookupsToStorage(updatedAll);
    setStatusAlert({
      type: 'info',
      message: `Removed "${itemName}" from ${currentTabConfig.label}.`
    });
  };

  const handleResetToDefaults = () => {
    if (window.confirm('Reset all defined IDs to official Kannur University defaults? Custom entries will be replaced.')) {
      onSaveLookups(DEFAULT_LOOKUP_DATA);
      saveLookupsToStorage(DEFAULT_LOOKUP_DATA);
      setEditingIdx(null);
      setStatusAlert({
        type: 'success',
        message: 'Reset all 7 ID lookup definitions to official university defaults.'
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
          message: 'No matching lookup columns found. Upload a sheet matching: Caste/CasteId, categoryName/categoryId, religion/religionId, CountryName/CountryId, corrStateName/corrStateId, QualificationName/QualificationFacultyId, Stream/QualificationSpecializationId, or a 2-column Name/ID sheet.'
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
        const existingMap = new Map(existing.map((x, idx) => [x.name.toLowerCase(), idx]));

        for (const item of items) {
          const lower = item.name.toLowerCase();
          if (existingMap.has(lower)) {
            const idx = existingMap.get(lower);
            existing[idx] = { ...existing[idx], id: item.id };
          } else {
            existing.push(item);
            existingMap.set(lower, existing.length - 1);
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
        const rows = list.map(item => ({
          [cfg.nameHeader]: item.name,
          [cfg.idHeader]: item.id
        }));
        const ws = XLSX.utils.json_to_sheet(rows.length > 0 ? rows : [{ [cfg.nameHeader]: '', [cfg.idHeader]: '' }]);
        XLSX.utils.book_append_sheet(wb, ws, cfg.label);
      });

      const maxRows = Math.max(...LOOKUP_TAB_CONFIG.map(cfg => (lookups[cfg.key] || []).length));
      const masterRows = [];
      for (let i = 0; i < maxRows; i++) {
        const row = {};
        LOOKUP_TAB_CONFIG.forEach(cfg => {
          const item = (lookups[cfg.key] || [])[i];
          row[cfg.nameHeader] = item ? item.name : '';
          row[cfg.idHeader] = item ? item.id : '';
        });
        masterRows.push(row);
      }
      const masterWs = XLSX.utils.json_to_sheet(masterRows);
      XLSX.utils.book_append_sheet(wb, masterWs, 'All Defined IDs');

      XLSX.writeFile(wb, 'Defined_Admission_Lookup_IDs.xlsx');
      setStatusAlert({
        type: 'success',
        message: 'Downloaded Defined_Admission_Lookup_IDs.xlsx with all 7 entity ID mappings.'
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
          maxWidth: '960px',
          width: '100%',
          maxHeight: '90vh',
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
                Configure codes for Category, Caste, Religion, Country, State, Qualification, and Stream.
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

        {/* 7 Tabs Navigation Bar */}
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
            color: 'var(--muted)'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            <span style={{ fontWeight: 600, color: 'var(--ink)' }}>Output Column Mapping:</span>
            <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
              {currentTabConfig.nameHeader}
            </code>
            <span>↔</span>
            <code style={{ padding: '2px 6px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
              {currentTabConfig.idHeader}
            </code>
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
          <div style={{ position: 'relative', flex: '1', minWidth: '220px' }}>
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
              placeholder={`Search ${currentTabConfig.label} by name or ID...`}
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

          {/* Quick Add Entry Form */}
          <form
            onSubmit={handleAdd}
            style={{ display: 'flex', gap: '8px', alignItems: 'center' }}
          >
            <input
              type="text"
              value={newName}
              onChange={(e) => setNewName(e.target.value)}
              placeholder={`New ${currentTabConfig.nameLabel} (e.g. Mukkuvan)`}
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
              placeholder={`ID (e.g. 320)`}
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
        </div>

        {/* Scrollable Table Content */}
        <div style={{ flex: 1, overflowY: 'auto', minHeight: '260px', maxHeight: '45vh' }}>
          {filteredList.length === 0 ? (
            <div style={{ padding: '40px 20px', textAlign: 'center', color: 'var(--muted)' }}>
              <Database size={36} color="var(--line)" style={{ margin: '0 auto 8px' }} />
              <p style={{ margin: 0, fontSize: '13px', fontWeight: 600 }}>
                {searchQuery ? `No entries matching "${searchQuery}"` : `No ${currentTabConfig.label} defined yet.`}
              </p>
              <span style={{ fontSize: '11.5px', marginTop: '4px', display: 'inline-block' }}>
                Use the inputs above to add a new ID pair, or upload an Excel/CSV file to bulk import.
              </span>
            </div>
          ) : (
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
