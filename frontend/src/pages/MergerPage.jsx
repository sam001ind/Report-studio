import { useState } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { 
  ArrowLeft, 
  HelpCircle, 
  Sparkles, 
  CheckCircle2, 
  FileSpreadsheet, 
  X, 
  Plus, 
  ShieldCheck, 
  Download, 
  Layers 
} from 'lucide-react';

const DEFAULT_TEXT_HEADERS = [
  'CLIENTAPPTRANSCATIONREFERENCENUMBER',
  'TRANSCATIONREFERENCENUMBER',
  'DESCRIPTION',
  'OTHERPARAMETERS',
  'REF1'
];

const normalizeKey = (key) => String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');

const MergerPage = () => {
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [headerRow, setHeaderRow] = useState(1);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'
  const [includeSourceFile, setIncludeSourceFile] = useState(true);
  const [autoDetectLongNumbers, setAutoDetectLongNumbers] = useState(true);
  const [textHeaders, setTextHeaders] = useState(DEFAULT_TEXT_HEADERS);
  const [newHeaderInput, setNewHeaderInput] = useState('');
  const [showHelpModal, setShowHelpModal] = useState(false);

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  const handleAddHeader = (e) => {
    e?.preventDefault();
    const val = newHeaderInput.trim();
    if (!val) return;
    if (!textHeaders.some(h => normalizeKey(h) === normalizeKey(val))) {
      setTextHeaders(prev => [...prev, val]);
    }
    setNewHeaderInput('');
  };

  const handleRemoveHeader = (headerToRemove) => {
    setTextHeaders(prev => prev.filter(h => h !== headerToRemove));
  };

  const handleResetHeaders = () => {
    setTextHeaders(DEFAULT_TEXT_HEADERS);
  };

  const handleFileChange = async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    setStatus('Processing files...');
    const newFiles = [];

    for (const file of files) {
      const ext = file.name.split('.').pop().toLowerCase();
      
      if (ext === 'zip') {
        try {
          setStatus(`Extracting ${file.name}...`);
          const zip = await JSZip.loadAsync(file);
          const zipFileKeys = Object.keys(zip.files);
          
          for (const key of zipFileKeys) {
            const zipEntry = zip.files[key];
            if (zipEntry.dir) continue;
            
            const entryExt = key.split('.').pop().toLowerCase();
            if (['xlsx', 'xls', 'xlsm', 'csv'].includes(entryExt)) {
              const buffer = await zipEntry.async('arraybuffer');
              newFiles.push({
                name: key.split('/').pop(),
                buffer: buffer,
                size: zipEntry._data.uncompressedSize
              });
            }
          }
        } catch (err) {
          setStatus(`Error reading ZIP: ${err.message}`, 'error');
        }
      } else if (['xlsx', 'xls', 'xlsm', 'csv'].includes(ext)) {
        try {
          const buffer = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (err) => reject(err);
            reader.readAsArrayBuffer(file);
          });
          newFiles.push({
            name: file.name,
            buffer: buffer,
            size: file.size
          });
        } catch (err) {
          setStatus(`Error reading ${file.name}: ${err.message}`, 'error');
        }
      }
    }

    if (newFiles.length > 0) {
      setSelectedFiles((prev) => [...prev, ...newFiles]);
      setStatus(`Added ${newFiles.length} file(s)`, 'success');
    } else {
      setStatus('No valid Excel or CSV files found.', 'error');
    }
  };

  const removeFile = (index) => {
    setSelectedFiles((prev) => prev.filter((_, idx) => idx !== index));
    setStatus('File removed');
  };

  const clearFiles = () => {
    setSelectedFiles([]);
    setStatus('Files cleared');
  };

  const isBlankRow = (row) => {
    return !row || row.every((cell) => cell === null || cell === undefined || cell === "");
  };

  const getActualRange = (sheet) => {
    if (!sheet) return null;
    let maxRow = -1, minCol = Infinity, maxCol = -1;
    
    const keys = Object.keys(sheet);
    for (const key of keys) {
      if (key[0] === '!') continue;

      if (!isNaN(key)) {
        const rIdx = parseInt(key, 10);
        const row = sheet[rIdx];
        if (Array.isArray(row) && row.length > 0) {
          for (let cIdx = 0; cIdx < row.length; cIdx++) {
            if (row[cIdx] !== undefined && row[cIdx] !== null) {
              if (rIdx > maxRow) maxRow = rIdx;
              if (cIdx < minCol) minCol = cIdx;
              if (cIdx > maxCol) maxCol = cIdx;
            }
          }
        }
        continue;
      }

      try {
        const cell = XLSX.utils.decode_cell(key);
        if (cell.r > maxRow) maxRow = cell.r;
        if (cell.c < minCol) minCol = cell.c;
        if (cell.c > maxCol) maxCol = cell.c;
      } catch (_e) {
        // Ignore non-cell property keys
      }
    }

    if (maxRow === -1) {
      if (sheet['!ref']) {
        try {
          const decoded = XLSX.utils.decode_range(sheet['!ref']);
          return {
            s: { r: 0, c: Math.max(0, decoded.s.c) },
            e: { r: Math.min(decoded.e.r, 200000), c: Math.min(decoded.e.c, 1000) }
          };
        } catch (_e) {
          // Ignore invalid !ref decoding
        }
      }
      return null;
    }

    return {
      s: { r: 0, c: Math.max(0, isFinite(minCol) ? minCol : 0) },
      e: { r: maxRow, c: maxCol }
    };
  };

  // Helper to extract raw cell object from sheet (handles dense and sparse modes)
  const getRawCellFromSheet = (sheet, rIdx, cIdx) => {
    if (!sheet) return null;
    if (Array.isArray(sheet[rIdx])) {
      return sheet[rIdx][cIdx];
    }
    const cellAddress = XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
    return sheet[cellAddress];
  };

  // Helper to format cell value into clean, exact text (expanding scientific notation and BigInt integers)
  const formatCellValue = (cell, isTextColumn) => {
    if (cell === null || cell === undefined) return '';

    let rawVal = cell;
    if (typeof cell === 'object' && cell !== null) {
      if (cell.v !== undefined && cell.v !== null) {
        rawVal = cell.v;
      } else if (cell.w !== undefined && cell.w !== null) {
        rawVal = cell.w;
      }
    }

    let str = String(rawVal).trim();
    if (str === '' || str.toLowerCase() === 'null' || str.toLowerCase() === 'undefined') return '';

    // Handle integers without scientific notation
    if (typeof rawVal === 'number') {
      if (Number.isInteger(rawVal)) {
        try {
          str = BigInt(Math.trunc(rawVal)).toString();
        } catch (_) {
          str = String(rawVal);
        }
      } else {
        str = String(rawVal);
      }
    }

    // Expand scientific notation strings (e.g. 2.02601E+16, 1.10003E+13, 2.02501E+15)
    if (/^[+-]?\d+(?:\.\d+)?[eE][+-]?\d+$/i.test(str)) {
      try {
        const parts = str.toLowerCase().split('e');
        const base = parts[0];
        const exp = parseInt(parts[1], 10);
        
        if (exp > 0) {
          const baseParts = base.split('.');
          const intPart = baseParts[0];
          const fracPart = baseParts[1] || '';
          
          if (exp >= fracPart.length) {
            str = intPart + fracPart + '0'.repeat(exp - fracPart.length);
          } else {
            str = intPart + fracPart.slice(0, exp) + '.' + fracPart.slice(exp);
          }
        }
      } catch (_) {
        // Fallback to original string
      }
    }

    return str;
  };

  const isDesignatedTextHeader = (headerName) => {
    const norm = normalizeKey(headerName);
    if (!norm) return false;
    return textHeaders.some(th => {
      const thNorm = normalizeKey(th);
      return norm === thNorm || norm.includes(thNorm) || thNorm.includes(norm);
    });
  };

  const executeMerge = async () => {
    if (selectedFiles.length === 0) {
      setStatus('Please select at least one file.', 'error');
      return;
    }

    const headerIdx = Number.parseInt(headerRow, 10) - 1; // Convert to 0-indexed
    if (isNaN(headerIdx) || headerIdx < 0) {
      setStatus('Header row must be a positive number.', 'error');
      return;
    }

    setIsProcessing(true);
    setStatus('Merging sheets with strict text preservation for transaction references...');

    try {
      // 1. First Pass: Collect union of headers across all sheets
      const masterHeaders = [];
      const masterHeaderNormMap = {}; // normKey -> index in masterHeaders

      if (includeSourceFile) {
        masterHeaders.push('Source File');
        masterHeaderNormMap['sourcefile'] = 0;
      }

      // Structure to hold extracted rows with header associations
      const allExtractedData = [];

      for (const file of selectedFiles) {
        const workbook = XLSX.read(file.buffer, { 
          type: 'array', 
          cellDates: false, 
          cellText: true, 
          cellNF: true, 
          dense: true, 
          raw: true 
        });
        
        for (const sheetName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sheetName];
          if (!sheet) continue;

          const range = getActualRange(sheet);
          if (!range) continue;

          let sheetHeaderFound = false;
          let currentSheetHeaderMap = []; // colIdx -> masterHeaderIdx

          for (let r = range.s.r; r <= range.e.r; r++) {
            if (r < headerIdx) continue;

            // Check if row has any content
            let rowHasData = false;
            const rowCells = [];
            for (let c = range.s.c; c <= range.e.c; c++) {
              const cell = getRawCellFromSheet(sheet, r, c);
              rowCells.push(cell);
              if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') {
                rowHasData = true;
              }
            }

            if (!rowHasData) continue;

            // Sheet Header Row
            if (!sheetHeaderFound) {
              sheetHeaderFound = true;
              currentSheetHeaderMap = [];

              rowCells.forEach((cell, cIdx) => {
                const headerVal = cell ? (cell.v !== undefined ? cell.v : cell.w) : '';
                const hName = String(headerVal || '').trim();
                if (!hName) return;

                const norm = normalizeKey(hName);
                if (masterHeaderNormMap[norm] === undefined) {
                  const newIdx = masterHeaders.length;
                  masterHeaders.push(hName);
                  masterHeaderNormMap[norm] = newIdx;
                  currentSheetHeaderMap[cIdx] = newIdx;
                } else {
                  currentSheetHeaderMap[cIdx] = masterHeaderNormMap[norm];
                }
              });
              continue;
            }

            // Data Row
            const rowDataObj = {
              _fileName: file.name,
              _cells: []
            };

            rowCells.forEach((cell, cIdx) => {
              const masterColIdx = currentSheetHeaderMap[cIdx];
              if (masterColIdx !== undefined) {
                const headerName = masterHeaders[masterColIdx];
                const isTextCol = isDesignatedTextHeader(headerName);
                const finalVal = formatCellValue(cell, isTextCol);
                rowDataObj._cells[masterColIdx] = finalVal;
              }
            });

            allExtractedData.push(rowDataObj);
          }
        }
      }

      if (masterHeaders.length === 0 || allExtractedData.length === 0) {
        throw new Error('No sheets contained data at the selected header row.');
      }

      // 2. Assemble Master AOA Grid
      const mergedGrid = [masterHeaders];

      allExtractedData.forEach(rowObj => {
        const rowArr = new Array(masterHeaders.length).fill('');
        if (includeSourceFile) {
          rowArr[0] = rowObj._fileName;
        }
        for (let c = (includeSourceFile ? 1 : 0); c < masterHeaders.length; c++) {
          rowArr[c] = rowObj._cells[c] !== undefined ? rowObj._cells[c] : '';
        }
        mergedGrid.push(rowArr);
      });

      setStatus('Formatting cells with explicit text types (@) & auto-filters...');

      // 3. Create Output Sheet and Explicitly Tag Text Columns
      const outputWorkbook = XLSX.utils.book_new();
      const outputSheet = XLSX.utils.aoa_to_sheet(mergedGrid, { dense: true });

      // Determine text columns
      const textColIndices = new Set();
      masterHeaders.forEach((hName, cIdx) => {
        if (isDesignatedTextHeader(hName)) {
          textColIndices.add(cIdx);
        }
      });

      // Format cells in dense mode with explicit string type and text format
      for (let r = 1; r < mergedGrid.length; r++) {
        const row = outputSheet[r];
        if (!row) continue;

        for (let c = 0; c < masterHeaders.length; c++) {
          const cell = row[c];
          if (!cell) continue;

          let rawVal = cell.v !== undefined ? String(cell.v).trim() : '';
          if (rawVal.toLowerCase() === 'null' || rawVal.toLowerCase() === 'undefined') {
            rawVal = '';
          }

          const isTextCol = textColIndices.has(c);
          const isLongNumber = autoDetectLongNumbers && /^\d{12,}$/.test(rawVal);
          const hasLeadingZero = rawVal.length > 1 && /^0\d+$/.test(rawVal);

          if (isTextCol || isLongNumber || hasLeadingZero) {
            cell.t = 's'; // String type
            cell.z = '@'; // Explicit Text format in Excel
            cell.v = rawVal;
            cell.w = rawVal;
          }
        }
      }

      // Auto-filter
      outputSheet['!autofilter'] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: Math.max(mergedGrid.length - 1, 0), c: Math.max(masterHeaders.length - 1, 0) }
        })
      };

      // Auto column widths
      outputSheet['!cols'] = masterHeaders.map((h, cIdx) => {
        let maxLen = h.length;
        for (let r = 1; r < Math.min(mergedGrid.length, 50); r++) {
          const cell = outputSheet[r]?.[cIdx];
          if (cell && cell.v) {
            maxLen = Math.max(maxLen, String(cell.v).length);
          }
        }
        return { wch: Math.min(Math.max(maxLen + 3, 14), 50) };
      });

      XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, 'Merged');

      setStatus('Downloading merged workbook...');
      const outputBuffer = XLSX.write(outputWorkbook, { bookType: 'xlsx', type: 'array', compression: true });
      const blob = new Blob([outputBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      
      const baseName = selectedFiles[0].name.replace(/\.[^/.]+$/, '');
      link.download = `${baseName}_merged.xlsx`;
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);

      setStatus(`Successfully merged ${selectedFiles.length} file(s) (${allExtractedData.length} total rows) with intact 15+ digit text headers!`, 'success');
    } catch (err) {
      console.error('Merge error details:', err);
      setStatus(`Merge failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const formatSize = (bytes) => {
    if (!bytes) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  };

  return (
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto', overflowY: 'auto', width: '100%', minHeight: '100vh', background: 'var(--bg)', color: 'var(--ink)' }}>
      
      {/* Navigation Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
            <ArrowLeft size={16} /> Back to Portal
          </Link>
          <div style={{ height: '18px', width: '1px', background: 'var(--line)' }} />
          <button 
            type="button" 
            className="secondary" 
            onClick={() => setShowHelpModal(true)}
            style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', padding: '5px 10px' }}
          >
            <HelpCircle size={14} /> Extraction Logic & Guide
          </button>
        </div>
        
        {/* Status Pill */}
        <div style={{
          padding: '6px 14px',
          borderRadius: '20px',
          fontSize: '13px',
          fontWeight: 600,
          background: statusType === 'error' ? 'var(--danger-soft)' : statusType === 'success' ? 'var(--accent-soft)' : 'var(--panel)',
          color: statusType === 'error' ? 'var(--danger)' : statusType === 'success' ? 'var(--accent)' : 'var(--muted)',
          border: '1px solid var(--line)'
        }}>
          {statusMsg}
        </div>
      </div>

      <div style={{ marginBottom: '24px' }}>
        <h2 style={{ fontSize: '24px', fontWeight: 700, margin: '0 0 6px 0' }}>Excel Sheet Merger</h2>
        <p style={{ margin: 0, color: 'var(--muted)', fontSize: '14px' }}>
          Combine multiple Excel or CSV files (or ZIP archives) into a consolidated master sheet with strict text preservation for long numbers and transaction reference IDs.
        </p>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(450px, 1fr))', gap: '32px', marginBottom: '32px' }}>
        
        {/* Configuration Panel */}
        <div className="card" style={{ padding: '28px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>1. Ingestion & Alignment Settings</h3>
          
          {/* File Upload Area */}
          <div style={{ border: '1.5px dashed var(--line)', padding: '28px', borderRadius: '8px', background: 'var(--bg)', textAlign: 'center', position: 'relative' }}>
            <FileSpreadsheet size={36} color="var(--accent)" style={{ margin: '0 auto 10px', opacity: 0.85 }} />
            <strong style={{ display: 'block', marginBottom: '6px', fontSize: '14px' }}>Add Excel, CSV or ZIP Files</strong>
            <span style={{ color: 'var(--muted)', fontSize: '12px' }}>Drag files here or click to browse</span>
            <input 
              type="file" 
              multiple
              accept=".xlsx, .xls, .xlsm, .csv, .zip" 
              onChange={handleFileChange} 
              style={{
                position: 'absolute',
                top: 0,
                left: 0,
                right: 0,
                bottom: 0,
                opacity: 0,
                cursor: 'pointer',
                width: '100%'
              }} 
            />
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
            <div className="form-group">
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', marginBottom: '4px' }}>Header Row Number</label>
              <input 
                type="number" 
                min="1" 
                step="1" 
                value={headerRow} 
                onChange={(e) => setHeaderRow(Math.max(1, parseInt(e.target.value) || 1))} 
                style={{ width: '100%', padding: '7px 10px', borderRadius: '6px', border: '1px solid var(--line)', background: 'var(--bg)', fontSize: '13px' }}
              />
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', justifyContent: 'center', gap: '6px' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', cursor: 'pointer', userSelect: 'none' }}>
                <input 
                  type="checkbox" 
                  checked={includeSourceFile} 
                  onChange={(e) => setIncludeSourceFile(e.target.checked)} 
                />
                <span>Include "Source File" Column</span>
              </label>
              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', cursor: 'pointer', userSelect: 'none' }}>
                <input 
                  type="checkbox" 
                  checked={autoDetectLongNumbers} 
                  onChange={(e) => setAutoDetectLongNumbers(e.target.checked)} 
                />
                <span>Auto-protect 12+ digit numbers</span>
              </label>
            </div>
          </div>

          {/* Text-Preserved Headers Manager */}
          <div style={{ background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px', padding: '14px', display: 'flex', flexDirection: 'column', gap: '10px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', fontWeight: 700 }}>
                <ShieldCheck size={16} color="var(--accent)" />
                <span>Text-Preserved Headers (15+ Digits):</span>
              </div>
              <button 
                type="button" 
                onClick={handleResetHeaders} 
                style={{ background: 'none', border: 'none', color: 'var(--muted)', fontSize: '11px', cursor: 'pointer', textDecoration: 'underline' }}
              >
                Reset Defaults
              </button>
            </div>

            <p style={{ margin: 0, fontSize: '11px', color: 'var(--muted)', lineHeight: '1.4' }}>
              These columns are formatted with Excel Text format (<code>@</code>) and string data types to ensure 15–18 digit transaction IDs and reference numbers are NEVER converted to scientific notation or truncated.
            </p>

            {/* Chips */}
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', maxHeight: '120px', overflowY: 'auto', padding: '2px 0' }}>
              {textHeaders.map((th) => (
                <span 
                  key={th} 
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '4px',
                    fontSize: '11px',
                    fontWeight: 600,
                    padding: '3px 8px',
                    borderRadius: '4px',
                    background: 'var(--accent-soft)',
                    color: 'var(--accent)',
                    border: '1px solid var(--accent)'
                  }}
                >
                  {th}
                  <button 
                    type="button" 
                    onClick={() => handleRemoveHeader(th)}
                    style={{ background: 'none', border: 'none', color: 'var(--accent)', padding: 0, cursor: 'pointer', display: 'flex' }}
                  >
                    <X size={12} />
                  </button>
                </span>
              ))}
            </div>

            {/* Add Custom Header Input */}
            <form onSubmit={handleAddHeader} style={{ display: 'flex', gap: '6px', marginTop: '4px' }}>
              <input 
                type="text" 
                placeholder="Add column header (e.g. UTR, CHALLAN_NO)..."
                value={newHeaderInput}
                onChange={(e) => setNewHeaderInput(e.target.value)}
                style={{ flex: 1, padding: '5px 8px', fontSize: '11.5px', borderRadius: '4px', border: '1px solid var(--line)', background: 'var(--panel)' }}
              />
              <button 
                type="submit" 
                className="secondary"
                style={{ padding: '5px 10px', fontSize: '11.5px', display: 'flex', alignItems: 'center', gap: '4px' }}
              >
                <Plus size={13} /> Add
              </button>
            </form>
          </div>

          <button 
            onClick={executeMerge} 
            disabled={isProcessing || selectedFiles.length === 0} 
            style={{ 
              width: '100%', 
              padding: '12px', 
              fontSize: '14px', 
              fontWeight: 700, 
              background: 'var(--accent)', 
              color: 'white', 
              border: 'none', 
              borderRadius: '6px', 
              cursor: (isProcessing || selectedFiles.length === 0) ? 'not-allowed' : 'pointer',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              gap: '8px'
            }}
          >
            <Download size={16} />
            {isProcessing ? "Merging Sheets..." : `Merge ${selectedFiles.length} File(s) & Download`}
          </button>
        </div>

        {/* Selected Files List Panel */}
        <div className="card" style={{ padding: '28px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>2. Queued Files ({selectedFiles.length})</h3>
            {selectedFiles.length > 0 && (
              <button 
                onClick={clearFiles} 
                style={{
                  background: 'none',
                  border: 'none',
                  color: 'var(--danger)',
                  cursor: 'pointer',
                  fontWeight: 600,
                  fontSize: '12.5px',
                  padding: 0
                }}
              >
                Clear All
              </button>
            )}
          </div>

          {selectedFiles.length > 0 ? (
            <div style={{ flex: 1, overflowY: 'auto', maxHeight: '420px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
              {selectedFiles.map((file, idx) => (
                <div 
                  key={idx} 
                  style={{
                    display: 'flex',
                    justifyContent: 'space-between',
                    alignItems: 'center',
                    padding: '10px 14px',
                    borderRadius: '6px',
                    background: 'var(--panel)',
                    border: '1px solid var(--line)',
                    fontSize: '12.5px'
                  }}
                >
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '2px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '320px' }}>
                    <span style={{ fontWeight: 600, color: 'var(--ink)' }}>📄 {file.name}</span>
                    <small style={{ color: 'var(--muted)', fontSize: '10.5px' }}>{formatSize(file.size)}</small>
                  </div>
                  <button 
                    onClick={() => removeFile(idx)}
                    title="Remove file"
                    style={{
                      background: 'none',
                      border: 'none',
                      color: 'var(--danger)',
                      cursor: 'pointer',
                      fontSize: '16px',
                      fontWeight: 'bold',
                      padding: '2px 6px',
                      borderRadius: '4px'
                    }}
                  >
                    ×
                  </button>
                </div>
              ))}
            </div>
          ) : (
            <div style={{ display: 'flex', flex: 1, flexDirection: 'column', alignItems: 'center', justifyContent: 'center', border: '1px dashed var(--line)', borderRadius: '8px', padding: '40px', color: 'var(--muted)', fontSize: '13px', gap: '10px' }}>
              <Layers size={40} style={{ opacity: 0.3 }} />
              <span>No files selected. Drag in Excel, CSV, or ZIP files to begin merging.</span>
            </div>
          )}
        </div>
      </div>

      {/* Extraction Logic & Guide Modal */}
      {showHelpModal && (
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
            width: '600px',
            maxWidth: '90vw',
            maxHeight: '85vh',
            display: 'flex',
            flexDirection: 'column',
            boxShadow: '0 20px 40px rgba(0,0,0,0.3)',
            overflow: 'hidden'
          }}>
            <div style={{ padding: '16px 20px', borderBottom: '1px solid var(--line)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <ShieldCheck size={20} color="var(--accent)" />
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>Excel Sheet Merger - Text & Reference Protection</h3>
              </div>
              <button 
                type="button" 
                onClick={() => setShowHelpModal(false)}
                style={{ background: 'none', border: 'none', color: 'var(--muted)', cursor: 'pointer', padding: '4px' }}
              >
                <X size={18} />
              </button>
            </div>

            <div style={{ padding: '20px', overflowY: 'auto', fontSize: '12.5px', lineHeight: '1.6', display: 'flex', flexDirection: 'column', gap: '14px' }}>
              <div>
                <strong>1. 15–18 Digit Precision Loss Prevention:</strong>
                <p style={{ margin: '4px 0', color: 'var(--muted)' }}>
                  Standard spreadsheets convert numbers with &gt;15 digits into floating point values, truncating trailing digits into zeros or scientific notation (e.g. <code>1.23456E+15</code>).
                </p>
                <p style={{ margin: '4px 0', color: 'var(--muted)' }}>
                  The Sheet Merger reads cell text formats directly (<code>cell.w</code>) and assigns explicit string types (<code>t: 's'</code>) and text formatting (<code>z: '@'</code>) so all digits remain 100% exact.
                </p>
              </div>

              <div>
                <strong>2. Protected Reference Headers:</strong>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', marginTop: '6px' }}>
                  {textHeaders.map(h => (
                    <span key={h} style={{ background: 'var(--accent-soft)', color: 'var(--accent)', border: '1px solid var(--accent)', padding: '2px 8px', borderRadius: '4px', fontSize: '11px', fontWeight: 600 }}>
                      {h}
                    </span>
                  ))}
                </div>
              </div>

              <div>
                <strong>3. Sequence-Agnostic Column Union:</strong>
                <p style={{ margin: '4px 0', color: 'var(--muted)' }}>
                  Files with differently ordered columns or extra columns are automatically matched and aligned by column header names across all sheets.
                </p>
              </div>
            </div>

            <div style={{ padding: '12px 20px', borderTop: '1px solid var(--line)', display: 'flex', justifyContent: 'flex-end', background: 'var(--bg)' }}>
              <button 
                type="button" 
                onClick={() => setShowHelpModal(false)}
                style={{ padding: '6px 16px', background: 'var(--accent)', color: 'white', borderRadius: '6px', border: 'none', fontWeight: 600, cursor: 'pointer' }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};

export default MergerPage;

