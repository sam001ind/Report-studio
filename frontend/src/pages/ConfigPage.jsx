import React, { useState } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';

const ConfigPage = ({ dataset, setDataset, setStats }) => {
  const [uploadStatus, setUploadStatus] = useState('Click or Drag to Upload CSV / Excel / JSON');
  
  // Data Management State
  const [combineCols, setCombineCols] = useState([]);
  const [combineNewName, setCombineNewName] = useState('');
  
  const [filterCol, setFilterCol] = useState('');
  const [filterOp, setFilterOp] = useState('equals');
  const [filterVal, setFilterVal] = useState('');
  
  const [activeFilters, setActiveFilters] = useState([]);
  const [configName, setConfigName] = useState('');

  const [filteredRows, setFilteredRows] = useState(dataset.rows || []);
  const [columns, setColumns] = useState(dataset.columns || []);
  const [sourceRows, setSourceRows] = useState(dataset.rows || []);

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setUploadStatus(`Loaded: ${file.name}`);
    const ext = file.name.split('.').pop().toLowerCase();
    
    const processData = (data) => {
      if (!data || data.length === 0) return alert("File is empty.");
      
      let columns = Object.keys(data[0] || {}).map(c => c.trim());
      
      const records = data.map(row => {
         const newRow = {};
         for (let col of columns) {
            newRow[col] = (row[col] === null || row[col] === undefined) ? "" : row[col];
         }
         return newRow;
      });
      
      setColumns(columns);
      setSourceRows(records);
      setFilteredRows(records);
      
      setDataset({ columns, rows: records });
      setStats({ rows: records.length, cols: columns.length });
      
      if (columns.length > 0) {
        setCombineCols([columns[0], columns[0]]);
        setFilterCol(columns[0]);
      }
    };
    
    if (ext === 'csv') {
       Papa.parse(file, {
          header: true,
          skipEmptyLines: true,
          complete: (results) => {
             processData(results.data);
          },
          error: (err) => {
             alert("Error parsing CSV: " + err.message);
          }
       });
    } else if (ext === 'xlsx' || ext === 'xls') {
       const reader = new FileReader();
       reader.onload = (evt) => {
          try {
             const bstr = evt.target.result;
             const workbook = XLSX.read(bstr, {type: 'binary'});
             const sheetName = workbook.SheetNames[0];
             const sheet = workbook.Sheets[sheetName];
             const data = XLSX.utils.sheet_to_json(sheet, {defval: ""});
             processData(data);
          } catch(err) {
             alert("Error parsing Excel: " + err.message);
          }
       };
       reader.readAsBinaryString(file);
    } else if (ext === 'json') {
       const reader = new FileReader();
       reader.onload = (evt) => {
          try {
             const data = JSON.parse(evt.target.result);
             if (!Array.isArray(data)) return alert("JSON must be an array of objects.");
             processData(data);
          } catch(err) {
             alert("Error parsing JSON: " + err.message);
          }
       };
       reader.readAsText(file);
    } else {
       alert("Unsupported file format. Please upload CSV, Excel, or JSON.");
    }
  };

  const handleCombineColumns = () => {
    if (combineCols.length < 2 || !combineNewName) return alert("Select columns and enter a new name.");
    if (columns.includes(combineNewName)) return alert("Column name already exists!");

    const newCols = [...columns, combineNewName];
    const newSource = sourceRows.map(row => {
      const combinedVal = combineCols.map(c => row[c] || '').join(' ').trim();
      return { ...row, [combineNewName]: combinedVal };
    });
    
    setColumns(newCols);
    setSourceRows(newSource);
    applyFilters(activeFilters, newSource, newCols);
    
    setCombineNewName('');
  };

  const applyFilters = (filters, sourceData, currentColumns = columns) => {
    let result = [...sourceData];
    
    filters.forEach(f => {
      result = result.filter(row => {
        const cellVal = String(row[f.col] || '').toLowerCase();
        const testVal = String(f.val).toLowerCase();
        
        if (f.op === 'equals') return cellVal === testVal;
        if (f.op === 'contains') return cellVal.includes(testVal);
        if (f.op === 'greater') return parseFloat(cellVal) > parseFloat(testVal);
        if (f.op === 'less') return parseFloat(cellVal) < parseFloat(testVal);
        if (f.op === 'not_blank') return cellVal.trim() !== '';
        return true;
      });
    });
    
    setFilteredRows(result);
    setDataset({ columns: currentColumns, rows: result });
    setStats({ rows: result.length, cols: currentColumns.length });
  };

  const handleAddFilter = () => {
    if (!filterCol || (!filterVal && filterOp !== 'not_blank')) return alert("Enter a filter value.");
    const newFilter = { id: Date.now(), col: filterCol, op: filterOp, val: filterVal };
    const newFilters = [...activeFilters, newFilter];
    setActiveFilters(newFilters);
    applyFilters(newFilters, sourceRows);
    setFilterVal('');
  };

  const removeFilter = (id) => {
    const newFilters = activeFilters.filter(f => f.id !== id);
    setActiveFilters(newFilters);
    applyFilters(newFilters, sourceRows);
  };

  const saveReportConfig = () => {
    if (!configName) return alert("Enter a configuration name.");
    const saved = JSON.parse(localStorage.getItem('rs_configs') || '[]');
    saved.push({
      id: `cfg_${Date.now()}`,
      name: configName,
      columns,
      filters: activeFilters
    });
    localStorage.setItem('rs_configs', JSON.stringify(saved));
    alert("Configuration Saved!");
  };

  const handleExportCsv = () => {
    if (filteredRows.length === 0) return alert("No data to export.");
    const header = columns.join(",");
    const csvContent = filteredRows.map(row => 
      columns.map(col => `"${String(row[col] || '').replace(/"/g, '""')}"`).join(",")
    ).join("\n");
    const fullCsv = header + "\n" + csvContent;
    
    const dataStr = "data:text/csv;charset=utf-8," + encodeURIComponent(fullCsv);
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", (configName || "filtered_data") + ".csv");
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  };

  return (
    <div style={styles.page}>
      <h2>Report Configuration</h2>
      <p className="subtitle">Upload data, combine fields, apply filters, and save your data configuration.</p>

      <div className="card">
        <h3 style={{ marginTop: 0 }}>Step 1: Upload Dataset</h3>
        <div style={styles.fileDrop} onClick={() => document.getElementById('fileInput').click()}>
          <h3>{uploadStatus}</h3>
          <p style={{ color: 'var(--muted)', fontSize: '14px' }}>
            Data is parsed securely inside your browser. No data is uploaded to a server.
          </p>
          <input 
            type="file" 
            id="fileInput" 
            accept=".csv, .xlsx, .xls, .json" 
            style={{ display: 'none' }}
            onChange={handleFileUpload}
          />
        </div>
      </div>
      
      {columns.length > 0 && (
        <div id="dataMgmtSection">
          
          <div className="card">
            <h3 style={{ marginTop: 0 }}>Step 2: Data Management</h3>
            
            <div style={{ marginBottom: '24px' }}>
              <label style={{ fontWeight: 600, display: 'block', marginBottom: '8px' }}>Combine Columns (2 or more)</label>
              <div className="form-row" style={{ marginBottom: 0, alignItems: 'center' }}>
                {combineCols.map((c, i) => (
                  <React.Fragment key={i}>
                    <select value={c} onChange={e => {
                      const newCols = [...combineCols];
                      newCols[i] = e.target.value;
                      setCombineCols(newCols);
                    }}>
                      {columns.map(col => <option key={col} value={col}>{col}</option>)}
                    </select>
                    {i < combineCols.length - 1 && <span style={{ fontWeight: 'bold' }}>+</span>}
                  </React.Fragment>
                ))}
                
                <button onClick={() => setCombineCols([...combineCols, columns[0]])} className="secondary" style={{ padding: '8px', fontSize: '12px' }}>+ Add</button>
                {combineCols.length > 2 && (
                  <button onClick={() => setCombineCols(combineCols.slice(0, -1))} className="danger" style={{ padding: '8px', fontSize: '12px', marginLeft: '-4px' }}>- Drop</button>
                )}
                
                <span style={{ fontWeight: 'bold', margin: '0 8px' }}>=</span>
                
                <div className="form-group" style={{ margin: 0, minWidth: '150px' }}>
                  <input type="text" value={combineNewName} onChange={e => setCombineNewName(e.target.value)} placeholder="New Column Name" />
                </div>
                <button onClick={handleCombineColumns} className="secondary">Combine</button>
              </div>
            </div>

            <div>
              <label style={{ fontWeight: 600, display: 'block', marginBottom: '8px' }}>Add Filters</label>
              <div className="form-row" style={{ marginBottom: 0 }}>
                <select value={filterCol} onChange={e => setFilterCol(e.target.value)}>
                  {columns.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
                <select value={filterOp} onChange={e => setFilterOp(e.target.value)}>
                  <option value="equals">Equals</option>
                  <option value="contains">Contains</option>
                  <option value="greater">Greater Than</option>
                  <option value="less">Less Than</option>
                  <option value="not_blank">Is Not Blank</option>
                </select>
                <div className="form-group">
                  <input type="text" value={filterVal} onChange={e => setFilterVal(e.target.value)} placeholder="Value" disabled={filterOp === 'not_blank'} />
                </div>
                <button onClick={handleAddFilter} className="secondary">Add Filter</button>
              </div>
              
              <div className="tag-list" style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginTop: '12px' }}>
                {activeFilters.map(f => (
                  <div key={f.id} style={{ background: 'var(--accent-soft)', color: 'var(--accent)', padding: '4px 12px', borderRadius: '100px', fontSize: '12px', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '8px' }}>
                    {f.col} {f.op} {f.val}
                    <button onClick={() => removeFilter(f.id)} style={{ background: 'none', border: 'none', padding: 0, color: 'inherit', fontSize: '16px', cursor: 'pointer' }}>×</button>
                  </div>
                ))}
              </div>
            </div>
          </div>

          <div className="card">
            <h3 style={{ marginTop: 0 }}>Step 3: Save Configuration</h3>
            <div className="form-row">
              <div className="form-group" style={{ flex: 2 }}>
                <label>Configuration Name</label>
                <input type="text" value={configName} onChange={e => setConfigName(e.target.value)} placeholder="e.g., Monthly Sales Filtered" />
              </div>
              <button onClick={saveReportConfig}>Save Config</button>
            </div>
          </div>

          <div className="card">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <h3 style={{ margin: 0 }}>Data Preview (Filtered) - {filteredRows.length} Rows</h3>
              <button onClick={handleExportCsv} className="secondary" style={{ borderColor: 'var(--accent)', color: 'var(--accent)', padding: '6px 12px' }}>
                ⬇️ Export CSV
              </button>
            </div>
            <div style={{ overflowX: 'auto', maxHeight: '300px' }}>
              <table className="data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                <thead>
                  <tr>
                    {columns.map(c => <th key={c} style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid var(--line)', background: '#f3f6f5', fontWeight: 600, textTransform: 'uppercase', fontSize: '11px' }}>{c}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.slice(0, 50).map((row, idx) => (
                    <tr key={idx}>
                      {columns.map(c => <td key={c} style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid var(--line)' }}>{row[c]}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
              {filteredRows.length > 50 && <p style={{ textAlign: 'center', color: 'var(--muted)', marginTop: '12px', fontSize: '12px' }}>Showing top 50 rows.</p>}
            </div>
          </div>
          
        </div>
      )}
    </div>
  );
};

const styles = {
  page: {
    padding: '40px',
    maxWidth: '1400px',
    margin: '0 auto',
  },
  fileDrop: {
    border: '2px dashed var(--line)',
    borderRadius: '12px',
    padding: '40px',
    textAlign: 'center',
    backgroundColor: '#fafafa',
    cursor: 'pointer',
    transition: 'border-color 0.2s',
  }
};

export default ConfigPage;
