import React, { useState, useMemo } from 'react';
import { useNavigate } from 'react-router-dom';
import * as XLSX from 'xlsx';
import {
  ArrowLeft,
  Upload,
  Database,
  FileSpreadsheet,
  Download,
  Search,
  CheckCircle2,
  AlertCircle,
  FileCode,
  Layers,
  Table as TableIcon,
  Copy,
  Terminal,
  RefreshCw,
  Info
} from 'lucide-react';

const DatabaseDumpExtractorPage = () => {
  const navigate = useNavigate();

  const [dumpFile, setDumpFile] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Upload a PostgreSQL or SQL database dump / backup file (.dump, .sql, .dmp, .bak, .txt)');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Parsed Tables: { [tableName: string]: { cols: string[], rows: any[][] } }
  const [tablesData, setTablesData] = useState({});
  const [activeTable, setActiveTable] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [copiedCmd, setCopiedCmd] = useState(false);

  // Unescape PostgreSQL COPY string
  const unescapePgVal = (val) => {
    if (val === '\\N') return '';
    if (!val.includes('\\')) return val;
    return val.replace(/\\n/g, '\n').replace(/\\r/g, '\r').replace(/\\t/g, '\t').replace(/\\\\/g, '\\');
  };

  // Parse SQL INSERT INTO values string
  const parseSqlValues = (valuesStr) => {
    const rows = [];
    let currentRow = [];
    let currentVal = [];
    let inQuotes = false;
    let quoteChar = null;
    let escaped = false;
    let inTuple = false;

    const n = valuesStr.length;
    for (let i = 0; i < n; i++) {
      const ch = valuesStr[i];

      if (!inTuple) {
        if (ch === '(') {
          inTuple = true;
          currentRow = [];
          currentVal = [];
        }
        continue;
      }

      if (inQuotes) {
        if (escaped) {
          currentVal.push(ch);
          escaped = false;
        } else if (ch === '\\') {
          escaped = true;
        } else if (ch === quoteChar) {
          if (i + 1 < n && valuesStr[i + 1] === quoteChar) {
            currentVal.push(quoteChar);
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          currentVal.push(ch);
        }
      } else {
        if (ch === "'" || ch === '"' || ch === '`') {
          inQuotes = true;
          quoteChar = ch;
        } else if (ch === ',') {
          const valStr = currentVal.join('').trim();
          currentRow.push(valStr.toUpperCase() === 'NULL' ? '' : valStr);
          currentVal = [];
        } else if (ch === ')') {
          const valStr = currentVal.join('').trim();
          currentRow.push(valStr.toUpperCase() === 'NULL' ? '' : valStr);
          rows.push(currentRow);
          currentRow = [];
          currentVal = [];
          inTuple = false;
        } else {
          currentVal.push(ch);
        }
      }
    }

    return rows;
  };

  // Main File Upload Handler
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setDumpFile(file);
    setIsProcessing(true);
    setStatusMsg(`Scanning and extracting dump file: ${file.name} (${(file.size / (1024 * 1024)).toFixed(2)} MB)...`);
    setStatusType('normal');

    try {
      const text = await file.text();
      const lines = text.split(/\r?\n/);

      const parsedTables = {};
      const tableSchemas = {};

      let currentCopyTable = null;
      let inCreateTable = false;
      let currentCreateTable = '';
      let currentCreateCols = [];

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];

        // 1. Inside COPY block
        if (currentCopyTable !== null) {
          if (line.trim() === '\\.') {
            currentCopyTable = null;
            continue;
          }
          const rawParts = line.split('\t');
          const cleanedRow = rawParts.map(unescapePgVal);
          parsedTables[currentCopyTable].rows.push(cleanedRow);
          continue;
        }

        const stripped = line.trim();
        if (!stripped || stripped.startsWith('--') || stripped.startsWith('/*')) continue;

        // 2. CREATE TABLE detection
        if (!inCreateTable) {
          const mCreate = stripped.match(/^CREATE\s+TABLE(?:\s+IF\s+NOT\s+EXISTS)?\s+(?:ONLY\s+)?(?:[a-zA-Z0-9_]+\.)?([a-zA-Z0-9_"`]+)\s*\(/i);
          if (mCreate) {
            inCreateTable = true;
            currentCreateTable = mCreate[1].replace(/["`]/g, '').trim();
            currentCreateCols = [];
            if (stripped.endsWith(');')) {
              inCreateTable = false;
              tableSchemas[currentCreateTable] = currentCreateCols;
            }
            continue;
          }
        } else {
          if (stripped.startsWith(');')) {
            inCreateTable = false;
            tableSchemas[currentCreateTable] = currentCreateCols;
            continue;
          }
          const mCol = stripped.match(/^\s*([a-zA-Z0-9_"`]+)\s+([a-zA-Z0-9_]+(?:\([^)]*\))?)/);
          if (mCol) {
            const cname = mCol[1].replace(/["`]/g, '').trim();
            if (!['CONSTRAINT', 'PRIMARY', 'KEY', 'FOREIGN', 'UNIQUE', 'CHECK'].includes(cname.toUpperCase())) {
              currentCreateCols.push(cname);
            }
          }
          continue;
        }

        // 3. COPY block detection (e.g. COPY studentmaster_fyugp (col1, col2) FROM stdin;)
        const mCopy = stripped.match(/^COPY\s+(?:ONLY\s+)?(?:[a-zA-Z0-9_]+\.)?([a-zA-Z0-9_"`]+)\s*(?:\(([^)]+)\))?\s+FROM\s+stdin;/i);
        if (mCopy) {
          const tname = mCopy[1].replace(/["`]/g, '').trim();
          const colsStr = mCopy[2];
          const cols = colsStr ? colsStr.split(',').map(c => c.trim().replace(/["`]/g, '')) : (tableSchemas[tname] || []);

          if (!parsedTables[tname]) {
            parsedTables[tname] = { cols, rows: [] };
          } else if (cols.length > 0 && parsedTables[tname].cols.length === 0) {
            parsedTables[tname].cols = cols;
          }

          currentCopyTable = tname;
          continue;
        }

        // 4. INSERT INTO detection
        const mInsert = stripped.match(/^INSERT\s+INTO\s+(?:[a-zA-Z0-9_]+\.)?([a-zA-Z0-9_"`]+)\s*(?:\(([^)]+)\))?\s+VALUES\s*(.*);?$/i);
        if (mInsert) {
          const tname = mInsert[1].replace(/["`]/g, '').trim();
          const colsStr = mInsert[2];
          const valuesPart = mInsert[3];

          const cols = colsStr ? colsStr.split(',').map(c => c.trim().replace(/["`]/g, '')) : (tableSchemas[tname] || []);

          if (!parsedTables[tname]) {
            parsedTables[tname] = { cols, rows: [] };
          } else if (cols.length > 0 && parsedTables[tname].cols.length === 0) {
            parsedTables[tname].cols = cols;
          }

          if (valuesPart) {
            const rows = parseSqlValues(valuesPart);
            parsedTables[tname].rows.push(...rows);
          }
          continue;
        }
      }

      const tableNames = Object.keys(parsedTables);
      if (tableNames.length === 0) {
        setStatusMsg('No table data (COPY or INSERT INTO) found in the uploaded dump file.');
        setStatusType('error');
        setTablesData({});
        setActiveTable('');
      } else {
        // Ensure columns exist for all tables
        tableNames.forEach(tname => {
          const t = parsedTables[tname];
          if ((!t.cols || t.cols.length === 0) && t.rows.length > 0) {
            t.cols = t.rows[0].map((_, idx) => `Column_${idx + 1}`);
          }
        });

        setTablesData(parsedTables);
        setActiveTable(tableNames[0]);
        const totalRows = tableNames.reduce((sum, n) => sum + parsedTables[n].rows.length, 0);
        setStatusMsg(`Successfully extracted ${tableNames.length} table(s) with ${totalRows.toLocaleString()} total records!`);
        setStatusType('success');
      }
    } catch (err) {
      console.error(err);
      setStatusMsg(`Error parsing dump file: ${err.message}`);
      setStatusType('error');
    } finally {
      setIsProcessing(false);
    }
  };

  const tableNames = Object.keys(tablesData);
  const currentTableData = activeTable && tablesData[activeTable] ? tablesData[activeTable] : { cols: [], rows: [] };

  // Filtered rows for active table
  const filteredRows = useMemo(() => {
    if (!currentTableData.rows || currentTableData.rows.length === 0) return [];
    if (!searchQuery.trim()) return currentTableData.rows;

    const q = searchQuery.toLowerCase().trim();
    return currentTableData.rows.filter(row => {
      return row.some(cell => cell && String(cell).toLowerCase().includes(q));
    });
  }, [currentTableData.rows, searchQuery]);

  // Export current table or all tables to XLSX
  const exportToExcel = (exportAll = true) => {
    if (tableNames.length === 0) return;

    const wb = XLSX.utils.book_new();
    const targetTables = exportAll ? tableNames : [activeTable];

    targetTables.forEach(tname => {
      const t = tablesData[tname];
      if (!t || (!t.rows.length && !t.cols.length)) return;

      // Ensure every cell is written with exact text string format to preserve leading zeroes and IDs
      const aoa = [
        t.cols,
        ...t.rows.map(r => r.map(c => (c === null || c === undefined ? '' : String(c))))
      ];

      const ws = XLSX.utils.aoa_to_sheet(aoa);

      // Auto-fit column widths
      const colWidths = t.cols.map((col, cIdx) => {
        let maxL = String(col).length;
        for (let r = 0; r < Math.min(200, t.rows.length); r++) {
          const val = t.rows[r][cIdx];
          if (val) maxL = Math.max(maxL, String(val).length);
        }
        return { wch: Math.min(50, Math.max(12, maxL + 3)) };
      });
      ws['!cols'] = colWidths;

      // Safe sheet name (max 31 chars)
      const safeTitle = tname.slice(0, 31).replace(/[:\\/?*[\]]/g, '_');
      XLSX.utils.book_append_sheet(wb, ws, safeTitle);
    });

    const baseName = dumpFile ? dumpFile.name.replace(/\.[^.]+$/, '') : 'database_dump';
    const fileName = `${baseName}_extracted.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  // Export active table to CSV
  const exportToCsv = () => {
    if (!activeTable || !currentTableData.rows.length) return;

    const aoa = [
      currentTableData.cols,
      ...currentTableData.rows.map(r => r.map(c => (c === null || c === undefined ? '' : String(c))))
    ];

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const csvContent = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${activeTable}_extracted.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const copyCliCommand = () => {
    const cmd = `python3 /Users/abhijitht/.gemini/antigravity/scratch/report-studio/extract_dump_to_xlsx.py "path/to/your_dump.dump"`;
    navigator.clipboard.writeText(cmd);
    setCopiedCmd(true);
    setTimeout(() => setCopiedCmd(false), 2000);
  };

  return (
    <div className="studio-container" style={{ padding: '24px 32px', maxWidth: '1440px', margin: '0 auto' }}>
      {/* Top Navigation */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px' }}>
        <button
          onClick={() => navigate('/')}
          className="btn-secondary"
          style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '8px 14px', fontSize: '13px' }}
        >
          <ArrowLeft size={16} /> Back to Dashboard
        </button>

        <button
          onClick={copyCliCommand}
          className="btn-secondary"
          style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px' }}
          title="Copy Python CLI extraction command for terminal usage"
        >
          <Terminal size={14} /> {copiedCmd ? 'Command Copied!' : 'Copy CLI Script Command'}
        </button>
      </div>

      {/* Header Banner */}
      <div style={{ marginBottom: '24px' }}>
        <h1 style={{ fontSize: '26px', fontWeight: 800, margin: '0 0 6px 0', color: 'var(--ink)' }}>
          Database Dump & Backup to Excel Extractor
        </h1>
        <p style={{ margin: 0, fontSize: '14px', color: 'var(--muted)' }}>
          Extract database dumps (PostgreSQL <code>COPY</code>, MySQL <code>INSERT INTO</code>, SQL backups) into fully intact Excel (<code>.xlsx</code>) files with all data, IDs, leading zeroes, and text formats preserved.
        </p>
      </div>

      {/* Upload Dropzone */}
      <div className="card" style={{ padding: '24px', border: '1.5px dashed var(--border)', background: 'var(--panel)', borderRadius: '12px', marginBottom: '24px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '14px' }}>
          <div style={{ background: 'rgba(23,107,135,0.1)', padding: '10px', borderRadius: '8px', color: 'var(--accent)' }}>
            <Database size={24} />
          </div>
          <div>
            <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 700 }}>Upload Database Dump File</h3>
            <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
              Supports PostgreSQL, MySQL, MariaDB, SQLite dump and backup files (<code>.dump</code>, <code>.sql</code>, <code>.dmp</code>, <code>.bak</code>, <code>.txt</code>)
            </span>
          </div>
        </div>

        <label
          style={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '32px',
            borderRadius: '8px',
            background: 'var(--bg)',
            border: '1px dashed var(--accent)',
            cursor: 'pointer',
            textAlign: 'center'
          }}
        >
          <Upload size={28} style={{ color: 'var(--accent)', marginBottom: '10px' }} />
          <span style={{ fontWeight: 600, fontSize: '14px' }}>
            {dumpFile ? dumpFile.name : 'Click or Drag to Upload Database Dump File'}
          </span>
          <span style={{ fontSize: '12px', color: 'var(--muted)', marginTop: '6px' }}>
            {dumpFile ? `${(dumpFile.size / (1024 * 1024)).toFixed(2)} MB • Ready` : 'Extracts full table schemas, column names, records, and text fields'}
          </span>
          <input
            type="file"
            accept=".dump,.sql,.dmp,.bak,.txt"
            onChange={handleFileUpload}
            style={{ display: 'none' }}
          />
        </label>
      </div>

      {/* Status Bar */}
      {statusMsg && (
        <div
          style={{
            padding: '12px 18px',
            marginBottom: '24px',
            borderRadius: '8px',
            fontSize: '13.5px',
            display: 'flex',
            alignItems: 'center',
            gap: '10px',
            background: statusType === 'error' ? 'rgba(239, 68, 68, 0.1)' : statusType === 'success' ? 'rgba(34, 197, 94, 0.1)' : 'var(--panel)',
            color: statusType === 'error' ? '#ef4444' : statusType === 'success' ? '#16a34a' : 'var(--ink)',
            border: `1px solid ${statusType === 'error' ? '#ef4444' : statusType === 'success' ? '#22c55e' : 'var(--border)'}`
          }}
        >
          {statusType === 'success' && <CheckCircle2 size={18} />}
          {statusType === 'error' && <AlertCircle size={18} />}
          {statusType === 'normal' && <Info size={18} />}
          <span>{statusMsg}</span>
        </div>
      )}

      {/* Extracted Data Section */}
      {tableNames.length > 0 && (
        <div className="card" style={{ padding: '24px', borderRadius: '12px', background: 'var(--panel)' }}>
          {/* Table Tabs */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '14px', marginBottom: '18px' }}>
            <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' }}>
              <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--muted)', marginRight: '6px' }}>Tables:</span>
              {tableNames.map(tname => (
                <button
                  key={tname}
                  onClick={() => { setActiveTable(tname); setSearchQuery(''); }}
                  style={{
                    padding: '6px 14px',
                    fontSize: '12.5px',
                    fontWeight: 600,
                    borderRadius: '6px',
                    cursor: 'pointer',
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '6px',
                    background: activeTable === tname ? 'var(--accent)' : 'var(--bg)',
                    color: activeTable === tname ? '#ffffff' : 'var(--ink)',
                    border: `1px solid ${activeTable === tname ? 'var(--accent)' : 'var(--border)'}`
                  }}
                >
                  <TableIcon size={14} />
                  <span>{tname}</span>
                  <span style={{
                    fontSize: '11px',
                    background: activeTable === tname ? 'rgba(255,255,255,0.25)' : 'rgba(0,0,0,0.06)',
                    padding: '1px 6px',
                    borderRadius: '10px'
                  }}>
                    {tablesData[tname].rows.length.toLocaleString()}
                  </span>
                </button>
              ))}
            </div>

            {/* Export Actions */}
            <div style={{ display: 'flex', gap: '10px' }}>
              <button
                onClick={exportToCsv}
                className="btn-secondary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 14px', fontSize: '13px' }}
              >
                <Download size={15} /> Export Table CSV
              </button>

              <button
                onClick={() => exportToExcel(true)}
                className="btn-primary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '8px 16px', fontSize: '13px', fontWeight: 600 }}
              >
                <FileSpreadsheet size={16} /> Export All to Excel (.xlsx)
              </button>
            </div>
          </div>

          {/* Active Table Details & Search */}
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px', flexWrap: 'wrap', gap: '12px' }}>
            <div style={{ fontSize: '13.5px', color: 'var(--muted)' }}>
              Showing table <strong>{activeTable}</strong>: <strong>{currentTableData.rows.length.toLocaleString()}</strong> records across <strong>{currentTableData.cols.length}</strong> columns
            </div>

            <div style={{ position: 'relative', width: '300px' }}>
              <Search size={15} style={{ position: 'absolute', left: '10px', top: '10px', color: 'var(--muted)' }} />
              <input
                type="text"
                placeholder="Search table rows..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                style={{ width: '100%', padding: '8px 10px 8px 32px', fontSize: '13px', borderRadius: '6px', border: '1px solid var(--border)' }}
              />
            </div>
          </div>

          {/* Table Container */}
          <div style={{ overflowX: 'auto', border: '1px solid var(--border)', borderRadius: '8px', maxHeight: '550px' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
              <thead style={{ position: 'sticky', top: 0, zIndex: 10 }}>
                <tr style={{ background: '#176B87', color: '#ffffff' }}>
                  <th style={{ padding: '10px 8px', textAlign: 'center', width: '50px', borderRight: '1px solid rgba(255,255,255,0.1)' }}>#</th>
                  {currentTableData.cols.map((col, cIdx) => (
                    <th key={cIdx} style={{ padding: '10px 10px', whiteSpace: 'nowrap', borderRight: '1px solid rgba(255,255,255,0.1)', fontWeight: 600 }}>
                      {col}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filteredRows.slice(0, 300).map((row, rIdx) => (
                  <tr
                    key={rIdx}
                    style={{
                      borderBottom: '1px solid var(--border)',
                      background: rIdx % 2 === 0 ? 'transparent' : 'rgba(0,0,0,0.015)'
                    }}
                  >
                    <td style={{ padding: '8px', textAlign: 'center', color: 'var(--muted)', fontWeight: 600, borderRight: '1px solid var(--border)' }}>
                      {rIdx + 1}
                    </td>
                    {currentTableData.cols.map((_, cIdx) => {
                      const val = row[cIdx];
                      return (
                        <td
                          key={cIdx}
                          style={{
                            padding: '8px 10px',
                            borderRight: '1px solid var(--border)',
                            whiteSpace: 'nowrap',
                            maxWidth: '300px',
                            overflow: 'hidden',
                            textOverflow: 'ellipsis'
                          }}
                          title={val || ''}
                        >
                          {val !== undefined && val !== null && val !== '' ? (
                            String(val)
                          ) : (
                            <span style={{ color: 'var(--muted)', opacity: 0.4 }}>—</span>
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <div style={{ marginTop: '12px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '12px', color: 'var(--muted)' }}>
            <span>
              Previewing first <strong>{Math.min(300, filteredRows.length)}</strong> of <strong>{filteredRows.length.toLocaleString()}</strong> rows in browser
            </span>
            <span>
              Click <strong>"Export All to Excel (.xlsx)"</strong> above to download all records intact.
            </span>
          </div>
        </div>
      )}

      {/* Terminal CLI Instructions Card */}
      <div className="card" style={{ marginTop: '24px', padding: '20px 24px', borderRadius: '12px', background: 'var(--panel)', border: '1px solid var(--border)' }}>
        <h4 style={{ margin: '0 0 10px 0', fontSize: '14.5px', display: 'flex', alignItems: 'center', gap: '8px' }}>
          <Terminal size={16} /> Python CLI Tool for High-Volume Dumps
        </h4>
        <p style={{ margin: '0 0 12px 0', fontSize: '13px', color: 'var(--muted)' }}>
          For massive database dumps exceeding 100+ megabytes, you can also run the bundled Python CLI extractor directly from your terminal:
        </p>
        <div style={{ background: '#0f172a', color: '#f8fafc', padding: '12px 16px', borderRadius: '8px', fontFamily: 'monospace', fontSize: '12.5px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <code>python3 /Users/abhijitht/.gemini/antigravity/scratch/report-studio/extract_dump_to_xlsx.py "/path/to/dumpfile.dump"</code>
          <button
            onClick={copyCliCommand}
            style={{ background: 'rgba(255,255,255,0.15)', color: '#fff', border: 'none', padding: '4px 10px', borderRadius: '4px', cursor: 'pointer', fontSize: '12px' }}
          >
            {copiedCmd ? 'Copied' : 'Copy'}
          </button>
        </div>
      </div>
    </div>
  );
};

export default DatabaseDumpExtractorPage;
