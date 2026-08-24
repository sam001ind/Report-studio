import { useState } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

const MergerPage = () => {
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [headerRow, setHeaderRow] = useState(1);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
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
                name: key.split('/').pop(), // Get filename only
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

  const denseAoaToSheet = (aoa) => {
    const ws = { '!dense': true };
    let maxCol = -1;

    for (let r = 0; r < aoa.length; r++) {
      const row = aoa[r];
      if (!row) continue;
      const denseRow = new Array(row.length);
      if (row.length > maxCol) maxCol = row.length;

      for (let c = 0; c < row.length; c++) {
        const val = row[c];
        if (val === null || val === undefined) continue;
        if (typeof val === 'number') {
          denseRow[c] = { v: val, t: 'n' };
        } else if (typeof val === 'boolean') {
          denseRow[c] = { v: val, t: 'b' };
        } else if (val instanceof Date) {
          denseRow[c] = { v: val, t: 'd' };
        } else {
          denseRow[c] = { v: String(val), t: 's' };
        }
      }
      ws[r] = denseRow;
    }

    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(aoa.length - 1, 0), c: Math.max(maxCol - 1, 0) }
    });

    return ws;
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
    setStatus('Merging sheets...');

    try {
      const mergedRows = [];
      let headerWritten = false;
      let targetHeaderLength = 0;

      for (const file of selectedFiles) {
        // Read file array buffer directly with dense mode
        const workbook = XLSX.read(file.buffer, { type: 'array', cellDates: true, dense: true });
        
        for (const sheetName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sheetName];
          if (!sheet) continue;

          const range = getActualRange(sheet);
          if (!range) continue;

          const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            range: range,
            raw: true,
            defval: null
          });

          if (!rows || rows.length === 0) continue;

          let sheetHeaderFound = false;

          for (let rIdx = 0; rIdx < rows.length; rIdx++) {
            if (rIdx < headerIdx) continue;

            const rowCells = rows[rIdx];

            // Skip empty rows
            if (isBlankRow(rowCells)) continue;

            // First non-empty row in this sheet (at or after headerIdx)
            if (!sheetHeaderFound) {
              sheetHeaderFound = true;

              if (!headerWritten) {
                // Write header: prepend "Source File" column
                mergedRows.push(["Source File", ...rowCells]);
                targetHeaderLength = rowCells.length + 1;
                headerWritten = true;
              }
              continue;
            }

            // Write Data Row: prepend current file name
            mergedRows.push([file.name, ...rowCells]);
          }
        }
      }

      if (!headerWritten) {
        throw new Error('No sheets contained data at the selected header row.');
      }

      setStatus('Creating output workbook...');
      const outputWorkbook = XLSX.utils.book_new();
      const outputSheet = denseAoaToSheet(mergedRows);

      outputSheet["!autofilter"] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: Math.max(mergedRows.length - 1, 0), c: Math.max(targetHeaderLength - 1, 0) }
        })
      };

      XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, "Merged");

      setStatus('Downloading merged workbook...');
      const outputBuffer = XLSX.write(outputWorkbook, { bookType: "xlsx", type: "array", dense: true, compression: true });
      const blob = new Blob([outputBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      
      const baseName = selectedFiles[0].name.replace(/\.[^/.]+$/, "");
      link.download = `${baseName}_merged.xlsx`;
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

      setStatus(`Successfully merged ${selectedFiles.length} file(s)!`, 'success');
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
        <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
          ← Back to Portal
        </Link>
        
        {/* Status Pill */}
        <div style={{
          padding: '6px 14px',
          borderRadius: '20px',
          fontSize: '13px',
          fontWeight: 600,
          background: statusType === 'error' ? 'var(--danger)' : statusType === 'success' ? 'var(--accent-soft)' : 'var(--panel)',
          color: statusType === 'error' ? 'white' : statusType === 'success' ? 'var(--accent)' : 'var(--muted)',
          border: '1px solid var(--line)'
        }}>
          {statusMsg}
        </div>
      </div>

      <h2>Excel Sheet Merger</h2>
      <p className="subtitle">Choose multiple Excel or CSV files (or a single .zip file containing them) to merge into a single workbook.</p>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(450px, 1fr))', gap: '32px', marginBottom: '32px' }}>
        {/* Configuration Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0 }}>Input Files</h3>
          
          {/* File Upload Area */}
          <div style={{ border: '1px dashed var(--line)', padding: '32px', borderRadius: '8px', background: 'var(--bg)', textAlign: 'center', position: 'relative' }}>
            <strong style={{ display: 'block', marginBottom: '8px' }}>Add Excel, CSV or ZIP Files</strong>
            <span style={{ color: 'var(--muted)', fontSize: '13px' }}>Drag here or click to browse</span>
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

          <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '16px' }}>
            <div className="form-group">
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', marginBottom: '4px' }}>Header Row Number</label>
              <input 
                type="number" 
                min="1" 
                step="1" 
                value={headerRow} 
                onChange={(e) => setHeaderRow(Math.max(1, parseInt(e.target.value) || 1))} 
                style={{ width: '100px' }}
              />
            </div>
          </div>

          <button 
            onClick={executeMerge} 
            disabled={isProcessing || selectedFiles.length === 0} 
            style={{ width: '100%', padding: '14px', fontSize: '15px', marginTop: '10px' }}
          >
            {isProcessing ? "Merging Sheets..." : "Merge Sheets & Download"}
          </button>
        </div>

        {/* Selected Files List Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0 }}>Files to Merge ({selectedFiles.length})</h3>
            {selectedFiles.length > 0 && (
              <button 
                onClick={clearFiles} 
                style={{
                  background: 'none',
                  border: 'none',
                  color: 'var(--danger)',
                  cursor: 'pointer',
                  fontWeight: 600,
                  fontSize: '13px',
                  padding: 0
                }}
              >
                Clear All
              </button>
            )}
          </div>

          {selectedFiles.length > 0 ? (
            <div style={{ flex: 1, overflowY: 'auto', maxHeight: '300px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
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
                    fontSize: '13px'
                  }}
                >
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '2px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '300px' }}>
                    <span style={{ fontWeight: 500 }}>📄 {file.name}</span>
                    <small style={{ color: 'var(--muted)', fontSize: '10px' }}>{formatSize(file.size)}</small>
                  </div>
                  <button 
                    onClick={() => removeFile(idx)}
                    style={{
                      background: 'none',
                      border: 'none',
                      color: 'var(--danger)',
                      cursor: 'pointer',
                      fontSize: '16px',
                      fontWeight: 'bold',
                      padding: '0 4px'
                    }}
                  >
                    ×
                  </button>
                </div>
              ))}
            </div>
          ) : (
            <div style={{ display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', border: '1px dashed var(--line)', borderRadius: '8px', padding: '40px', color: 'var(--muted)', fontSize: '14px' }}>
              No files selected. Drag in Excel, CSV, or ZIP files to begin.
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default MergerPage;
