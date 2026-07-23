import React, { useState } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';

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

  const handleFileChange = (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    setSelectedFiles((prev) => [...prev, ...files]);
    setStatus(`${files.length} file(s) added`, 'success');
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

  const executeMerge = async () => {
    if (selectedFiles.length === 0) {
      setStatus('Please select at least one Excel file.', 'error');
      return;
    }

    const headerIdx = Number.parseInt(headerRow, 10) - 1; // Convert to 0-indexed
    if (isNaN(headerIdx) || headerIdx < 0) {
      setStatus('Header row must be a positive number.', 'error');
      return;
    }

    setIsProcessing(true);
    setStatus('Reading and merging sheets...');

    try {
      const mergedRows = [];
      let headerWritten = false;
      let targetHeaderLength = 0;

      for (const file of selectedFiles) {
        const fileData = await new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = (e) => resolve(e.target.result);
          reader.onerror = (err) => reject(err);
          reader.readAsArrayBuffer(file);
        });

        const workbook = XLSX.read(fileData, { type: 'array', cellDates: true });
        
        for (const sheetName of workbook.SheetNames) {
          const sheet = workbook.Sheets[sheetName];
          const rows = XLSX.utils.sheet_to_json(sheet, {
            header: 1,
            raw: true,
            defval: null,
            blankrows: true
          });

          if (rows.length === 0) continue;

          for (let rIdx = 0; rIdx < rows.length; rIdx++) {
            if (rIdx < headerIdx) continue;

            const rowCells = rows[rIdx];

            // Skip empty rows
            if (isBlankRow(rowCells)) continue;

            // Handle Header Row
            if (rIdx === headerIdx) {
              if (headerWritten) continue;

              // Write header: prepend "Source File" column
              mergedRows.push(["Source File", ...rowCells]);
              targetHeaderLength = rowCells.length;
              headerWritten = true;
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
      const outputSheet = XLSX.utils.aoa_to_sheet(mergedRows, { cellDates: true });

      // Apply autofilter to all columns
      outputSheet["!autofilter"] = {
        ref: XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: Math.max(mergedRows.length - 1, 0), c: Math.max(targetHeaderLength, 0) }
        })
      };

      XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, "Merged");

      setStatus('Downloading merged workbook...');
      const outputBuffer = XLSX.write(outputWorkbook, { bookType: "xlsx", type: "array", cellDates: true });
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
      setStatus(`Merge failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
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
      <p className="subtitle">Choose multiple Excel workbooks to merge them into a single file with a prepended "Source File" column tracking origins.</p>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(450px, 1fr))', gap: '32px', marginBottom: '32px' }}>
        {/* Configuration Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0 }}>Input Files</h3>
          
          {/* File Upload Area */}
          <div style={{ border: '1px dashed var(--line)', padding: '32px', borderRadius: '8px', background: 'var(--bg)', textAlign: 'center', position: 'relative' }}>
            <strong style={{ display: 'block', marginBottom: '8px' }}>Add Excel Files</strong>
            <span style={{ color: 'var(--muted)', fontSize: '13px' }}>Drag here or click to add multiple files</span>
            <input 
              type="file" 
              multiple
              accept=".xlsx, .xls, .xlsm" 
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
            <h3 style={{ margin: 0 }}>Selected Files ({selectedFiles.length})</h3>
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
                  <span style={{ fontWeight: 500, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '300px' }}>
                    📄 {file.name}
                  </span>
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
              No files selected. Add some files to begin.
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default MergerPage;
