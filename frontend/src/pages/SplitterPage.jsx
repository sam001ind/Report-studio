import { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { readSpreadsheetWorkbook } from '../utils/excelParser';

const SplitterPage = () => {
  
  // File and sheet states
  const [selectedFile, setSelectedFile] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  
  // Configuration states
  const [splitMethod, setSplitMethod] = useState('rows'); // 'rows' | 'column'
  const [rowsPerFile, setRowsPerFile] = useState(10);
  const [headers, setHeaders] = useState([]);
  const [selectedHeader, setSelectedHeader] = useState('');
  const [skipBlankRows, setSkipBlankRows] = useState(true);
  const [copyColumnWidths, setCopyColumnWidths] = useState(true);

  // Status and processing states
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Output/Result states
  const [runSummary, setRunSummary] = useState(null);

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  // Populate headers when sheet changes
  useEffect(() => {
    if (workbook && selectedSheet) {
      const sheet = workbook.Sheets[selectedSheet];
      if (sheet) {
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
        if (rows.length > 0) {
          const firstRow = rows[0];
          const parsedHeaders = firstRow.map(h => String(h || '').trim()).filter(h => h !== '');
          setHeaders(parsedHeaders);
          setSelectedHeader(parsedHeaders[0] || '');
        }
      }
    } else {
      setHeaders([]);
      setSelectedHeader('');
    }
  }, [workbook, selectedSheet]);

  const handleFileChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setSelectedFile(file);
    setStatus('Reading file...');
    setRunSummary(null);

    try {
      const { workbook: wb, sheetNames: sheets } = await readSpreadsheetWorkbook(file);
      setWorkbook(wb);
      setSheetNames(sheets);
      setSelectedSheet(sheets[0] || '');
      setStatus('Workbook loaded successfully', 'success');
    } catch (err) {
      setStatus(`Error reading Excel: ${err.message}`, 'error');
      setWorkbook(null);
      setSheetNames([]);
      setSelectedSheet('');
    }
  };

  const isBlankRow = (row) => {
    return !row || row.every((cell) => cell === null || cell === undefined || cell === "");
  };

  const normalizeRows = (sheet, skipBlanks) => {
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: null,
      blankrows: !skipBlanks
    });

    if (rows.length === 0) {
      return { header: [], dataRows: [] };
    }

    const header = rows[0];
    const body = rows.slice(1);
    const dataRows = skipBlanks ? body.filter((row) => !isBlankRow(row)) : body;

    return { header, dataRows };
  };

  const executeSplit = async () => {
    if (!workbook || !selectedSheet) {
      setStatus('Select an Excel file and sheet first.', 'error');
      return;
    }

    if (splitMethod === 'rows') {
      const rowsLimit = Number.parseInt(rowsPerFile, 10);
      if (isNaN(rowsLimit) || rowsLimit < 1) {
        setStatus('Rows per file must be at least 1.', 'error');
        return;
      }
    } else {
      if (!selectedHeader) {
        setStatus('Select a column header to split by.', 'error');
        return;
      }
    }

    setIsProcessing(true);
    setStatus('Splitting sheet...');

    try {
      const sourceSheet = workbook.Sheets[selectedSheet];
      if (!sourceSheet) throw new Error(`Sheet not found: ${selectedSheet}`);

      const { header, dataRows } = normalizeRows(sourceSheet, skipBlankRows);
      if (header.length === 0) throw new Error("The selected sheet has no header row.");
      if (dataRows.length === 0) throw new Error("The selected sheet has no data rows below the header.");

      const generatedFiles = [];
      const zip = new JSZip();
      const baseName = selectedFile.name.replace(/\.[^/.]+$/, "").replace(/\s+/g, "_");

      if (splitMethod === 'rows') {
        // Mode 1: Split by Row count limit
        const rowsLimit = Number.parseInt(rowsPerFile, 10);
        const lotCount = Math.ceil(dataRows.length / rowsLimit);

        for (let index = 0; index < lotCount; index += 1) {
          const firstRowIndex = index * rowsLimit;
          const lotRows = dataRows.slice(firstRowIndex, firstRowIndex + rowsLimit);
          
          const outputWorkbook = XLSX.utils.book_new();
          const outputSheet = XLSX.utils.aoa_to_sheet([header, ...lotRows], { cellDates: true });

          if (copyColumnWidths && sourceSheet["!cols"]) {
            outputSheet["!cols"] = sourceSheet["!cols"].slice();
          }

          outputSheet["!autofilter"] = {
            ref: XLSX.utils.encode_range({
              s: { r: 0, c: 0 },
              e: { r: Math.max(lotRows.length, 1), c: Math.max(header.length - 1, 0) }
            })
          };

          XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, selectedSheet.slice(0, 31) || "Sheet1");

          const excelBuffer = XLSX.write(outputWorkbook, { bookType: "xlsx", type: "array", cellDates: true });
          const fileName = `${baseName}_lot_${String(index + 1).padStart(2, "0")}.xlsx`;
          
          zip.file(fileName, excelBuffer);

          generatedFiles.push({
            fileName,
            buffer: excelBuffer,
            dataRows: lotRows.length,
            range: `${firstRowIndex + 1}-${firstRowIndex + lotRows.length}`
          });
        }
      } else {
        // Mode 2: Split by unique values of selected Header column
        const headerColIdx = header.findIndex(h => String(h || '').trim() === selectedHeader);
        if (headerColIdx === -1) throw new Error(`Column not found in sheet: ${selectedHeader}`);

        // Group rows
        const groupedData = {};
        dataRows.forEach(row => {
          const val = String(row[headerColIdx] || 'Unassigned').trim();
          if (!groupedData[val]) groupedData[val] = [];
          groupedData[val].push(row);
        });

        Object.keys(groupedData).forEach((colValue) => {
          const groupRows = groupedData[colValue];
          const outputWorkbook = XLSX.utils.book_new();
          const outputSheet = XLSX.utils.aoa_to_sheet([header, ...groupRows], { cellDates: true });

          if (copyColumnWidths && sourceSheet["!cols"]) {
            outputSheet["!cols"] = sourceSheet["!cols"].slice();
          }

          outputSheet["!autofilter"] = {
            ref: XLSX.utils.encode_range({
              s: { r: 0, c: 0 },
              e: { r: Math.max(groupRows.length, 1), c: Math.max(header.length - 1, 0) }
            })
          };

          XLSX.utils.book_append_sheet(outputWorkbook, outputSheet, selectedSheet.slice(0, 31) || "Sheet1");

          const excelBuffer = XLSX.write(outputWorkbook, { bookType: "xlsx", type: "array", cellDates: true });
          const safeColVal = String(colValue).replace(/[/\\?%*:|"<>]/g, "_").trim() || "Unassigned";
          const fileName = `${safeColVal}.xlsx`;

          zip.file(fileName, excelBuffer);

          generatedFiles.push({
            fileName,
            buffer: excelBuffer,
            dataRows: groupRows.length,
            range: `Matching: ${colValue}`
          });
        });
      }

      // Generate ZIP blob and download it
      setStatus('Generating ZIP archive...');
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      
      const downloadUrl = URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `${baseName}_lots.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

      setRunSummary({
        fileName: selectedFile.name,
        sheetName: selectedSheet,
        sourceRows: dataRows.length,
        filesCreated: generatedFiles.length,
        rowsPerFile: splitMethod === 'rows' ? rowsPerFile : 'Dynamic (grouped)',
        generatedFiles
      });

      setStatus(`Successfully created and downloaded ${generatedFiles.length} files!`, 'success');
    } catch (err) {
      setStatus(`Error: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const downloadSingleFile = (file) => {
    const blob = new Blob([file.buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.download = file.fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(downloadUrl);
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

      <h2>Excel Lot Splitter</h2>
      <p className="subtitle">Split a large Excel worksheet into multiple files (either by row count or by unique values of a header) and download them in a ZIP.</p>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(450px, 1fr))', gap: '32px', marginBottom: '32px' }}>
        {/* Settings Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0 }}>Configuration</h3>
          
          {/* File Picker */}
          <div style={{ border: '1px dashed var(--line)', padding: '24px', borderRadius: '8px', background: 'var(--bg)', textAlign: 'center', position: 'relative' }}>
            <strong style={{ display: 'block', marginBottom: '8px' }}>Choose Excel Workbook</strong>
            {selectedFile ? (
              <div style={{ color: 'var(--accent)', fontWeight: 600 }}>📄 {selectedFile.name}</div>
            ) : (
              <span style={{ color: 'var(--muted)', fontSize: '13px' }}>Drag here or click to browse</span>
            )}
            <input 
              type="file" 
              accept=".xlsx, .xls, .xlsm, .csv" 
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

          {/* Splitting Method Toggle */}
          <div className="form-group">
            <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', marginBottom: '6px', display: 'block' }}>Split Method</label>
            <div style={{ display: 'flex', gap: '10px' }}>
              <button 
                onClick={() => setSplitMethod('rows')}
                style={{
                  flex: 1,
                  padding: '10px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: splitMethod === 'rows' ? 'var(--accent)' : 'var(--panel)',
                  color: splitMethod === 'rows' ? 'white' : 'var(--ink)',
                  cursor: 'pointer',
                  fontWeight: 600,
                  fontSize: '13px'
                }}
              >
                By Row Count
              </button>
              <button 
                onClick={() => setSplitMethod('column')}
                style={{
                  flex: 1,
                  padding: '10px',
                  borderRadius: '6px',
                  border: '1px solid var(--line)',
                  background: splitMethod === 'column' ? 'var(--accent)' : 'var(--panel)',
                  color: splitMethod === 'column' ? 'white' : 'var(--ink)',
                  cursor: 'pointer',
                  fontWeight: 600,
                  fontSize: '13px'
                }}
              >
                By Column Value
              </button>
            </div>
          </div>

          {/* Form Options */}
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
            <div className="form-group">
              <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', marginBottom: '4px' }}>Select Sheet</label>
              <select 
                value={selectedSheet} 
                onChange={(e) => setSelectedSheet(e.target.value)} 
                disabled={sheetNames.length === 0}
                style={{ width: '100%' }}
              >
                {sheetNames.map((name) => (
                  <option key={name} value={name}>{name}</option>
                ))}
              </select>
            </div>

            {splitMethod === 'rows' ? (
              <div className="form-group">
                <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', marginBottom: '4px' }}>Rows Per File</label>
                <input 
                  type="number" 
                  min="1" 
                  step="1" 
                  value={rowsPerFile} 
                  onChange={(e) => setRowsPerFile(Math.max(1, parseInt(e.target.value) || 1))} 
                  style={{ width: '100%' }}
                />
              </div>
            ) : (
              <div className="form-group">
                <label style={{ fontSize: '12px', fontWeight: 600, color: 'var(--muted)', marginBottom: '4px' }}>Split Column (Header)</label>
                <select 
                  value={selectedHeader} 
                  onChange={(e) => setSelectedHeader(e.target.value)} 
                  disabled={headers.length === 0}
                  style={{ width: '100%' }}
                >
                  {headers.length === 0 ? (
                    <option value="">-- No Columns --</option>
                  ) : (
                    headers.map((h) => (
                      <option key={h} value={h}>{h}</option>
                    ))
                  )}
                </select>
              </div>
            )}
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '14px' }}>
              <input 
                type="checkbox" 
                checked={skipBlankRows} 
                onChange={(e) => setSkipBlankRows(e.target.checked)} 
              />
              Skip blank rows
            </label>

            <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '14px' }}>
              <input 
                type="checkbox" 
                checked={copyColumnWidths} 
                onChange={(e) => setCopyColumnWidths(e.target.checked)} 
              />
              Keep column widths
            </label>
          </div>

          <button 
            onClick={executeSplit} 
            disabled={isProcessing || !workbook} 
            style={{ width: '100%', padding: '14px', fontSize: '15px', marginTop: '10px' }}
          >
            {isProcessing ? "Processing lots..." : "Split & Download ZIP"}
          </button>
        </div>

        {/* Run Summary Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0 }}>Run Summary</h3>
          {runSummary ? (
            <>
              <p style={{ margin: 0, fontSize: '14px', color: 'var(--muted)' }}>
                <strong>Source:</strong> {runSummary.fileName} • <strong>Sheet:</strong> {runSummary.sheetName}
              </p>
              
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '16px', borderBottom: '1px solid var(--line)', paddingBottom: '16px' }}>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span style={{ fontSize: '11px', color: 'var(--muted)', textTransform: 'uppercase' }}>Source rows</span>
                  <strong style={{ fontSize: '20px', color: 'var(--accent)' }}>{new Intl.NumberFormat('en-IN').format(runSummary.sourceRows)}</strong>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span style={{ fontSize: '11px', color: 'var(--muted)', textTransform: 'uppercase' }}>Files Created</span>
                  <strong style={{ fontSize: '20px', color: 'var(--accent)' }}>{runSummary.filesCreated}</strong>
                </div>
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span style={{ fontSize: '11px', color: 'var(--muted)', textTransform: 'uppercase' }}>Rows each / Split By</span>
                  <strong style={{ fontSize: '18px', color: 'var(--accent)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                    {splitMethod === 'rows' ? runSummary.rowsPerFile : `Column: ${selectedHeader}`}
                  </strong>
                </div>
              </div>

              {/* Table list of generated files */}
              <div style={{ flex: 1, overflowY: 'auto', maxHeight: '220px' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                  <thead>
                    <tr>
                      <th style={{ textAlign: 'left', padding: '6px', borderBottom: '1px solid var(--line)' }}>File</th>
                      <th style={{ textAlign: 'left', padding: '6px', borderBottom: '1px solid var(--line)' }}>Rows</th>
                      <th style={{ textAlign: 'left', padding: '6px', borderBottom: '1px solid var(--line)' }}>Source details</th>
                    </tr>
                  </thead>
                  <tbody>
                    {runSummary.generatedFiles.map((file, idx) => (
                      <tr key={idx} style={{ borderBottom: '1px solid var(--bg)' }}>
                        <td style={{ padding: '6px' }}>
                          <button 
                            onClick={() => downloadSingleFile(file)} 
                            style={{
                              background: 'none',
                              border: 'none',
                              color: 'var(--accent)',
                              textDecoration: 'underline',
                              cursor: 'pointer',
                              padding: 0,
                              fontWeight: 500,
                              fontSize: '12px',
                              textAlign: 'left'
                            }}
                          >
                            {file.fileName}
                          </button>
                        </td>
                        <td style={{ padding: '6px' }}>{file.dataRows}</td>
                        <td style={{ padding: '6px', color: 'var(--muted)' }}>{file.range}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          ) : (
            <div style={{ display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', border: '1px dashed var(--line)', borderRadius: '8px', padding: '40px', color: 'var(--muted)', fontSize: '14px' }}>
              No split run performed yet. Select configuration and split.
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default SplitterPage;
