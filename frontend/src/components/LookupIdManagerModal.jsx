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
  UploadCloud,
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

    // Check duplicate name in current tab
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
      const { imported, summary, total } = parseUploadedLookupSpreadsheet(workbook, activeTab);

      if (total === 0) {
        setStatusAlert({
          type: 'warning',
          message: 'No matching lookup columns found. Upload a sheet with columns matching any of: Caste/CasteId, categoryName/categoryId, religion/religionId, CountryName/CountryId, corrStateName/corrStateId, QualificationName/QualificationFacultyId, Stream/QualificationSpecializationId, or a 2-column Name/ID list.'
        });
        setIsProcessing(false);
        if (fileInputRef.current) fileInputRef.current.value = '';
        return;
      }

      // Merge imported into current lookups
      const merged = { ...lookups };
      const summaryMsgParts = [];

      for (const [key, items] of Object.entries(imported)) {
        if (!items || items.length === 0) continue;
        const existing = [...(merged[key] || [])];
        const existingMap = new Map(existing.map((x, idx) => [x.name.toLowerCase(), idx]));

        let addedCount = 0;
        let updatedCount = 0;

        for (const item of items) {
          const lower = item.name.toLowerCase();
          if (existingMap.has(lower)) {
            const idx = existingMap.get(lower);
            existing[idx] = { ...existing[idx], id: item.id };
            updatedCount++;
          } else {
            existing.push(item);
            existingMap.set(lower, existing.length - 1);
            addedCount++;
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

      // Create a master summary sheet and 7 individual sheets
      LOOKUP_TAB_CONFIG.forEach(cfg => {
        const list = lookups[cfg.key] || [];
        const rows = list.map(item => ({
          [cfg.nameHeader]: item.name,
          [cfg.idHeader]: item.id
        }));
        const ws = XLSX.utils.json_to_sheet(rows.length > 0 ? rows : [{ [cfg.nameHeader]: '', [cfg.idHeader]: '' }]);
        XLSX.utils.book_append_sheet(wb, ws, cfg.label);
      });

      // Also create a "Kannur IDs Master" sheet combining all columns side by side
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

  const totalEntriesCount = useMemo(() => {
    if (!lookups) return 0;
    return Object.values(lookups).reduce((sum, list) => sum + (Array.isArray(list) ? list.length : 0), 0);
  }, [lookups]);

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-in fade-in duration-200">
      <div className="relative w-full max-w-5xl bg-white dark:bg-slate-900 rounded-2xl shadow-2xl border border-slate-200 dark:border-slate-800 flex flex-col max-h-[90vh] overflow-hidden">
        
        {/* Top Header */}
        <div className="flex items-center justify-between px-6 py-4 border-b border-slate-100 dark:border-slate-800 bg-slate-50/70 dark:bg-slate-900/70">
          <div className="flex items-center space-x-3">
            <div className="p-2.5 bg-blue-100 dark:bg-blue-900/40 text-blue-700 dark:text-blue-300 rounded-xl">
              <Database className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <h2 className="text-xl font-bold text-slate-900 dark:text-white">
                  Define & Manage ID Lookups
                </h2>
                <span className="px-2 py-0.5 text-xs font-semibold rounded-full bg-blue-50 dark:bg-blue-950/60 text-blue-700 dark:text-blue-300 border border-blue-200/50 dark:border-blue-800/50">
                  {totalEntriesCount} Total Defined
                </span>
              </div>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">
                Define and map IDs for the 7 standard university master entities. Changes persist automatically and drive the 44-column output.
              </p>
            </div>
          </div>

          <div className="flex items-center space-x-2">
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
              title="Upload an Excel or CSV file containing ID definitions"
              className="flex items-center space-x-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700/60 transition shadow-sm"
            >
              <UploadCloud className="w-3.5 h-3.5 text-blue-600 dark:text-blue-400" />
              <span>{isProcessing ? 'Importing...' : 'Upload IDs (.xlsx / .csv)'}</span>
            </button>

            <button
              onClick={handleExportAllToExcel}
              title="Download all defined IDs into an Excel workbook"
              className="flex items-center space-x-1.5 px-3 py-1.5 text-xs font-medium text-slate-700 dark:text-slate-200 bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 rounded-lg hover:bg-slate-50 dark:hover:bg-slate-700/60 transition shadow-sm"
            >
              <Download className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400" />
              <span>Export All (.xlsx)</span>
            </button>

            <button
              onClick={handleResetToDefaults}
              title="Reset all lookups back to official Kannur University defaults"
              className="flex items-center space-x-1.5 px-3 py-1.5 text-xs font-medium text-rose-700 dark:text-rose-300 bg-rose-50 dark:bg-rose-950/40 border border-rose-200 dark:border-rose-900 rounded-lg hover:bg-rose-100 dark:hover:bg-rose-900/60 transition shadow-sm"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Reset Defaults</span>
            </button>

            <button
              onClick={onClose}
              className="p-1.5 text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 transition ml-2"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Status Alert Banner */}
        {statusAlert && (
          <div className={`px-6 py-2.5 text-xs flex items-center justify-between border-b ${
            statusAlert.type === 'error'
              ? 'bg-rose-50 dark:bg-rose-950/60 text-rose-800 dark:text-rose-200 border-rose-200 dark:border-rose-800'
              : statusAlert.type === 'warning'
              ? 'bg-amber-50 dark:bg-amber-950/60 text-amber-800 dark:text-amber-200 border-amber-200 dark:border-amber-800'
              : statusAlert.type === 'success'
              ? 'bg-emerald-50 dark:bg-emerald-950/60 text-emerald-800 dark:text-emerald-200 border-emerald-200 dark:border-emerald-800'
              : 'bg-blue-50 dark:bg-blue-950/60 text-blue-800 dark:text-blue-200 border-blue-200 dark:border-blue-800'
          }`}>
            <div className="flex items-center space-x-2">
              {statusAlert.type === 'success' ? (
                <CheckCircle2 className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0" />
              ) : (
                <AlertCircle className="w-4 h-4 shrink-0" />
              )}
              <span>{statusAlert.message}</span>
            </div>
            <button
              onClick={() => setStatusAlert(null)}
              className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        )}

        {/* 7 Tabs Bar */}
        <div className="flex items-center space-x-1 px-6 pt-3 border-b border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-900/40 overflow-x-auto no-scrollbar">
          {LOOKUP_TAB_CONFIG.map(tab => {
            const Icon = TAB_ICONS[tab.key] || Tag;
            const count = (lookups && Array.isArray(lookups[tab.key])) ? lookups[tab.key].length : 0;
            const isActive = activeTab === tab.key;

            return (
              <button
                key={tab.key}
                onClick={() => handleTabChange(tab.key)}
                className={`flex items-center space-x-2 px-3.5 py-2.5 text-xs font-medium rounded-t-xl transition-all whitespace-nowrap border-b-2 ${
                  isActive
                    ? 'border-blue-600 text-blue-600 dark:text-blue-400 bg-white dark:bg-slate-800 shadow-sm'
                    : 'border-transparent text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white hover:bg-white/50 dark:hover:bg-slate-800/50'
                }`}
              >
                <Icon className={`w-3.5 h-3.5 ${isActive ? 'text-blue-600 dark:text-blue-400' : 'text-slate-400'}`} />
                <span>{tab.label}</span>
                <span className={`px-1.5 py-0.2 text-[10px] rounded-full font-semibold ${
                  isActive
                    ? 'bg-blue-100 dark:bg-blue-900/60 text-blue-700 dark:text-blue-300'
                    : 'bg-slate-200 dark:bg-slate-800 text-slate-600 dark:text-slate-400'
                }`}>
                  {count}
                </span>
              </button>
            );
          })}
        </div>

        {/* Active Tab Sub-header & Description */}
        <div className="px-6 py-2.5 bg-blue-50/40 dark:bg-blue-950/20 border-b border-slate-200/60 dark:border-slate-800/60 flex items-center justify-between text-xs">
          <div className="flex items-center space-x-2">
            <span className="font-semibold text-slate-800 dark:text-slate-200">
              Output Header Mapping:
            </span>
            <code className="px-1.5 py-0.5 bg-white dark:bg-slate-800 rounded border border-slate-200 dark:border-slate-700 text-blue-600 dark:text-blue-400 font-mono text-[11px]">
              {currentTabConfig.nameHeader}
            </code>
            <span className="text-slate-400">↔</span>
            <code className="px-1.5 py-0.5 bg-white dark:bg-slate-800 rounded border border-slate-200 dark:border-slate-700 text-blue-600 dark:text-blue-400 font-mono text-[11px]">
              {currentTabConfig.idHeader}
            </code>
            <span className="text-slate-500 dark:text-slate-400 ml-2 hidden sm:inline">
              — {currentTabConfig.description}
            </span>
          </div>
          <span className="text-slate-500 dark:text-slate-400">
            {filteredList.length} of {currentList.length} shown
          </span>
        </div>

        {/* Action Bar: Search & Quick Add */}
        <div className="p-4 px-6 border-b border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 grid grid-cols-1 md:grid-cols-12 gap-3 items-center">
          {/* Search */}
          <div className="md:col-span-5 relative">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                setPage(0);
              }}
              placeholder={`Search ${currentTabConfig.label} by name or ID...`}
              className="w-full pl-9 pr-3 py-1.5 text-xs bg-slate-50 dark:bg-slate-800/80 border border-slate-200 dark:border-slate-700 rounded-lg text-slate-900 dark:text-white placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-blue-500"
            />
          </div>

          {/* Add New Entry Form */}
          <form onSubmit={handleAdd} className="md:col-span-7 flex items-center space-x-2">
            <input
              type="text"
              value={newName}
              onChange={(e) => setNewName(e.target.value)}
              placeholder={`New ${currentTabConfig.nameLabel} (e.g. Mukkuvan)`}
              className="flex-1 px-3 py-1.5 text-xs bg-slate-50 dark:bg-slate-800/80 border border-slate-200 dark:border-slate-700 rounded-lg text-slate-900 dark:text-white placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-blue-500"
            />
            <input
              type="text"
              value={newId}
              onChange={(e) => setNewId(e.target.value)}
              placeholder={`ID (e.g. 320)`}
              className="w-28 px-3 py-1.5 text-xs bg-slate-50 dark:bg-slate-800/80 border border-slate-200 dark:border-slate-700 rounded-lg text-slate-900 dark:text-white placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-blue-500"
            />
            <button
              type="submit"
              className="flex items-center space-x-1 px-3.5 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-lg shadow-sm transition whitespace-nowrap"
            >
              <Plus className="w-3.5 h-3.5" />
              <span>Add Entry</span>
            </button>
          </form>
        </div>

        {/* Entries Table */}
        <div className="flex-1 overflow-y-auto px-6 py-2">
          {filteredList.length === 0 ? (
            <div className="py-12 text-center">
              <Database className="w-10 h-10 text-slate-300 dark:text-slate-600 mx-auto mb-2" />
              <p className="text-sm font-medium text-slate-600 dark:text-slate-400">
                {searchQuery ? `No entries matching "${searchQuery}"` : `No ${currentTabConfig.label} defined yet.`}
              </p>
              <p className="text-xs text-slate-400 dark:text-slate-500 mt-1">
                Use the form above to add an entry, or upload a spreadsheet to define lookups in bulk.
              </p>
            </div>
          ) : (
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="border-b border-slate-200 dark:border-slate-800 text-[11px] font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider">
                  <th className="py-2.5 px-3 w-12 text-center">#</th>
                  <th className="py-2.5 px-3">{currentTabConfig.nameLabel} ({currentTabConfig.nameHeader})</th>
                  <th className="py-2.5 px-3 w-36">{currentTabConfig.idLabel} ({currentTabConfig.idHeader})</th>
                  <th className="py-2.5 px-3 w-28 text-right">Actions</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-800/60 text-xs">
                {paginatedList.map((item, idx) => {
                  const actualIndex = currentList.indexOf(item);
                  const isEditing = editingIdx === actualIndex;

                  return (
                    <tr
                      key={`${item.name}-${item.id}-${actualIndex}`}
                      className={`hover:bg-slate-50/80 dark:hover:bg-slate-800/50 transition ${
                        isEditing ? 'bg-blue-50/30 dark:bg-blue-950/30' : ''
                      }`}
                    >
                      <td className="py-2 px-3 text-center text-slate-400 font-mono text-[11px]">
                        {page * pageSize + idx + 1}
                      </td>

                      {/* Name Cell */}
                      <td className="py-2 px-3">
                        {isEditing ? (
                          <input
                            type="text"
                            value={editName}
                            onChange={(e) => setEditName(e.target.value)}
                            className="w-full px-2.5 py-1 text-xs bg-white dark:bg-slate-800 border border-blue-500 rounded text-slate-900 dark:text-white focus:outline-none"
                            autoFocus
                          />
                        ) : (
                          <span className="font-medium text-slate-800 dark:text-slate-200">
                            {item.name}
                          </span>
                        )}
                      </td>

                      {/* ID Cell */}
                      <td className="py-2 px-3">
                        {isEditing ? (
                          <input
                            type="text"
                            value={editId}
                            onChange={(e) => setEditId(e.target.value)}
                            className="w-full px-2.5 py-1 text-xs bg-white dark:bg-slate-800 border border-blue-500 rounded text-slate-900 dark:text-white font-mono focus:outline-none"
                          />
                        ) : (
                          <span className="px-2 py-0.5 font-mono text-xs font-semibold rounded bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border border-slate-200 dark:border-slate-700">
                            {item.id}
                          </span>
                        )}
                      </td>

                      {/* Actions Cell */}
                      <td className="py-2 px-3 text-right">
                        {isEditing ? (
                          <div className="flex items-center justify-end space-x-1">
                            <button
                              onClick={() => saveEdit(actualIndex)}
                              title="Save changes"
                              className="p-1 text-emerald-600 hover:text-emerald-700 hover:bg-emerald-50 dark:hover:bg-emerald-950/50 rounded transition"
                            >
                              <Check className="w-4 h-4" />
                            </button>
                            <button
                              onClick={cancelEdit}
                              title="Cancel editing"
                              className="p-1 text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 rounded transition"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          </div>
                        ) : (
                          <div className="flex items-center justify-end space-x-1">
                            <button
                              onClick={() => startEdit(item, actualIndex)}
                              title="Edit this entry"
                              className="p-1 text-slate-400 hover:text-blue-600 hover:bg-blue-50 dark:hover:bg-blue-950/40 rounded transition"
                            >
                              <Edit2 className="w-3.5 h-3.5" />
                            </button>
                            <button
                              onClick={() => handleDelete(actualIndex, item.name)}
                              title="Delete this entry"
                              className="p-1 text-slate-400 hover:text-rose-600 hover:bg-rose-50 dark:hover:bg-rose-950/40 rounded transition"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
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

        {/* Footer with Pagination & Close Button */}
        <div className="flex items-center justify-between px-6 py-3 border-t border-slate-200 dark:border-slate-800 bg-slate-50/70 dark:bg-slate-900/70">
          <div className="flex items-center space-x-2 text-xs text-slate-500 dark:text-slate-400">
            <span>
              Page {page + 1} of {totalPages}
            </span>
            <div className="flex items-center space-x-1 ml-2">
              <button
                onClick={() => setPage(p => Math.max(0, p - 1))}
                disabled={page === 0}
                className="p-1 rounded border border-slate-200 dark:border-slate-700 disabled:opacity-40 hover:bg-white dark:hover:bg-slate-800"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
              </button>
              <button
                onClick={() => setPage(p => Math.min(totalPages - 1, p + 1))}
                disabled={page >= totalPages - 1}
                className="p-1 rounded border border-slate-200 dark:border-slate-700 disabled:opacity-40 hover:bg-white dark:hover:bg-slate-800"
              >
                <ChevronRight className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>

          <div className="flex items-center space-x-2">
            <span className="text-[11px] text-slate-400 mr-2">
              All edits automatically save to browser storage
            </span>
            <button
              onClick={onClose}
              className="px-4 py-1.5 text-xs font-semibold text-white bg-blue-600 hover:bg-blue-700 rounded-lg shadow-sm transition"
            >
              Done / Apply to Import
            </button>
          </div>
        </div>

      </div>
    </div>
  );
}
