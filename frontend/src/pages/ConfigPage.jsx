import React, { useState, useEffect } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import { supabase } from '../supabaseClient';
import { useAuth } from '../context/AuthContext';

const ConfigPage = ({ dataset, setDataset, setStats, initialConfig }) => {
  const { user } = useAuth();
  const [uploadStatus, setUploadStatus] = useState('Click or Drag to Upload Dataset (CSV, XLSX, JSON)');
  const [isDragging, setIsDragging] = useState(false);
  
  const [sourceRows, setSourceRows] = useState(dataset.rows || []);
  const [sourceCols, setSourceCols] = useState(dataset.columns || []);
  
  const [pipelineSteps, setPipelineSteps] = useState([]);
  const [currentPipelineCols, setCurrentPipelineCols] = useState(dataset.columns || []);
  const [filteredRows, setFilteredRows] = useState(dataset.rows || []);
  
  const [configName, setConfigName] = useState('');
  const [showAddStep, setShowAddStep] = useState(false);

  // Load from initialConfig
  useEffect(() => {
    if (initialConfig && initialConfig.config_data) {
      setConfigName(initialConfig.name);
      const loadedSteps = initialConfig.config_data.pipeline || [];
      
      // Migrate old format
      if (initialConfig.config_data.filters && loadedSteps.length === 0) {
        setPipelineSteps(initialConfig.config_data.filters.map(f => ({ ...f, type: 'filter' })));
      } else {
        setPipelineSteps(loadedSteps);
      }
    }
  }, [initialConfig]);

  // Rerun pipeline whenever steps or source data changes
  useEffect(() => {
    runPipeline(pipelineSteps, sourceRows, sourceCols);
  }, [pipelineSteps, sourceRows, sourceCols]);

  const processFile = (file) => {
    if (!file) return;

    setUploadStatus(`Loaded: ${file.name}`);
    const ext = file.name.split('.').pop().toLowerCase();
    
    const processData = (data) => {
      if (!data || data.length === 0) return alert("File is empty.");
      
      let columns = Object.keys(data[0] || {}).map(c => c.trim());
      
      const records = data.map(row => {
         const newRow = {};
         for (let col of columns) {
            const origKey = Object.keys(row).find(k => k.trim() === col) || col;
            let val = row[origKey];
            newRow[col] = (val === null || val === undefined) ? "" : String(val);
         }
         return newRow;
      });
      
      setSourceCols(columns);
      setSourceRows(records);
      
      // The useEffect will automatically run the pipeline and update dataset/stats
    };
    
    if (ext === 'csv') {
       Papa.parse(file, { header: true, skipEmptyLines: true, complete: (results) => processData(results.data), error: (err) => alert("Error parsing CSV: " + err.message) });
    } else if (ext === 'xlsx' || ext === 'xls') {
       const reader = new FileReader();
       reader.onload = (evt) => {
          try {
             const arrayBuffer = evt.target.result;
             const workbook = XLSX.read(arrayBuffer, {type: 'array'});
             const sheetName = workbook.SheetNames[0];
             const sheet = workbook.Sheets[sheetName];
             const data = XLSX.utils.sheet_to_json(sheet, {defval: ""});
             processData(data);
          } catch(err) {
             alert("Error parsing Excel: " + err.message);
          }
       };
       reader.readAsArrayBuffer(file);
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
       alert("Unsupported file format.");
    }
  };

  const handleFileUpload = (e) => processFile(e.target.files[0]);
  const handleDragOver = (e) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = (e) => { e.preventDefault(); setIsDragging(false); };
  const handleDrop = (e) => { e.preventDefault(); setIsDragging(false); if (e.dataTransfer.files?.length > 0) processFile(e.dataTransfer.files[0]); };

  const runPipeline = (steps, initialRows, initialCols) => {
    let currentData = [...initialRows];
    let currentCols = [...initialCols];

    steps.forEach(step => {
      if (step.type === 'filter') {
        currentData = currentData.filter(row => {
          const cellVal = String(row[step.col] || '').toLowerCase();
          const testVal = String(step.val).toLowerCase();
          if (step.op === 'equals') return cellVal === testVal;
          if (step.op === 'contains') return cellVal.includes(testVal);
          if (step.op === 'greater') return parseFloat(cellVal) > parseFloat(testVal);
          if (step.op === 'less') return parseFloat(cellVal) < parseFloat(testVal);
          if (step.op === 'not_blank') return cellVal.trim() !== '';
          return true;
        });
      } else if (step.type === 'combine') {
        currentData = currentData.map(row => {
          const combinedVal = step.cols.map(c => row[c] || '').join(' ').trim();
          return { ...row, [step.newName]: combinedVal };
        });
        if (!currentCols.includes(step.newName)) currentCols.push(step.newName);
      } else if (step.type === 'calc') {
        const checkCondition = (row) => {
          const cellVal = String(row[step.condCol] || '').toLowerCase();
          const testVal = String(step.condVal).toLowerCase();
          if (step.condOp === 'equals') return cellVal === testVal;
          if (step.condOp === 'contains') return cellVal.includes(testVal);
          if (step.condOp === 'greater') return parseFloat(cellVal) > parseFloat(testVal);
          if (step.condOp === 'less') return parseFloat(cellVal) < parseFloat(testVal);
          if (step.condOp === 'not_blank') return cellVal.trim() !== '';
          return false;
        };

        if (step.outputMode === 'ROW_BY_ROW') {
          currentData = currentData.map(row => {
            let val = 0;
            if (checkCondition(row)) {
              if (step.op === 'SUM_IF') val = parseFloat(row[step.targetCol]) || 0;
              if (step.op === 'COUNT_IF') val = 1;
            }
            return { ...row, [step.newColName]: val };
          });
          if (!currentCols.includes(step.newColName)) currentCols.push(step.newColName);
        } 
        else if (step.outputMode === 'GLOBAL_AGG') {
          let total = 0;
          currentData.forEach(row => {
            if (checkCondition(row)) {
              if (step.op === 'SUM_IF') total += (parseFloat(row[step.targetCol]) || 0);
              if (step.op === 'COUNT_IF') total += 1;
            }
          });
          currentData = currentData.map(row => ({ ...row, [step.newColName]: total }));
          if (!currentCols.includes(step.newColName)) currentCols.push(step.newColName);
        }
        else if (step.outputMode === 'PIVOT') {
          const groups = {};
          currentData.forEach(row => {
            const groupKey = row[step.groupCol] || 'Unknown';
            if (!groups[groupKey]) {
              groups[groupKey] = { [step.groupCol]: groupKey, [step.newColName]: 0 };
            }
            if (checkCondition(row)) {
              if (step.op === 'SUM_IF') groups[groupKey][step.newColName] += (parseFloat(row[step.targetCol]) || 0);
              if (step.op === 'COUNT_IF') groups[groupKey][step.newColName] += 1;
            }
          });
          currentData = Object.values(groups);
          currentCols = [step.groupCol, step.newColName];
        }
      }
    });

    setCurrentPipelineCols(currentCols);
    setFilteredRows(currentData);
    setDataset({ columns: currentCols, rows: currentData });
    setStats({ rows: currentData.length, cols: currentCols.length });
  };

  const addStep = (type) => {
    const newStep = { id: Date.now(), type };
    if (type === 'filter') {
      newStep.col = currentPipelineCols[0] || '';
      newStep.op = 'equals';
      newStep.val = '';
    } else if (type === 'combine') {
      newStep.cols = [currentPipelineCols[0] || '', currentPipelineCols[0] || ''];
      newStep.newName = `Combined_${Date.now().toString().slice(-4)}`;
    } else if (type === 'calc') {
      newStep.op = 'SUM_IF';
      newStep.targetCol = currentPipelineCols[0] || '';
      newStep.condCol = currentPipelineCols[0] || '';
      newStep.condOp = 'equals';
      newStep.condVal = '';
      newStep.outputMode = 'ROW_BY_ROW';
      newStep.groupCol = currentPipelineCols[0] || '';
      newStep.newColName = `Calc_${Date.now().toString().slice(-4)}`;
    }
    setPipelineSteps([...pipelineSteps, newStep]);
    setShowAddStep(false);
  };

  const updateStep = (id, key, val) => {
    setPipelineSteps(pipelineSteps.map(s => s.id === id ? { ...s, [key]: val } : s));
  };

  const removeStep = (id) => {
    setPipelineSteps(pipelineSteps.filter(s => s.id !== id));
  };

  const saveReportConfig = async () => {
    if (!configName) return alert("Enter a configuration name.");
    
    const configData = {
      pipeline: pipelineSteps
    };
    
    const { data, error } = await supabase
      .from('configs')
      .insert([
        { name: configName, config_data: configData, user_id: user.id }
      ]);
      
    if (error) {
      console.error(error);
      alert("Error saving config to Supabase: " + error.message);
    } else {
      alert("Configuration Saved to Cloud!");
    }
  };

  const handleExportCsv = () => {
    if (filteredRows.length === 0) return alert("No data to export.");
    const header = currentPipelineCols.join(",");
    const csvContent = filteredRows.map(row => 
      currentPipelineCols.map(col => `"${String(row[col] || '').replace(/"/g, '""')}"`).join(",")
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
      <p className="subtitle">Build an interactive data pipeline to clean, filter, and transform your data.</p>

      {/* STEP 1 */}
      <div className="card">
        <h3 style={{ marginTop: 0 }}>Step 1: Data Source</h3>
        
        {initialConfig && sourceRows.length === 0 && (
          <div style={{ padding: '16px', background: 'var(--accent-soft)', color: 'var(--accent)', borderRadius: '8px', marginBottom: '16px', fontWeight: 600 }}>
            Editing saved configuration: "{initialConfig.name}". Please upload your dataset to apply these rules.
          </div>
        )}

        <div 
          style={{
            ...styles.fileDrop,
            borderColor: isDragging ? 'var(--accent)' : 'var(--line)',
            backgroundColor: isDragging ? 'rgba(0, 240, 255, 0.05)' : 'var(--panel)',
          }} 
          onClick={() => document.getElementById('fileInput').click()}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
        >
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
      
      {/* STEP 2 - PIPELINE */}
      {sourceCols.length > 0 && (
        <div id="dataMgmtSection">
          
          <div className="card">
            <h3 style={{ marginTop: 0 }}>Step 2: Data Pipeline</h3>
            <p style={{ color: 'var(--muted)', fontSize: '13px', marginBottom: '24px' }}>
              Add transformation steps. They run from top to bottom.
            </p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', marginBottom: '24px' }}>
              {pipelineSteps.map((step, idx) => (
                <div key={step.id} style={{ border: '1px solid var(--line)', borderRadius: '8px', padding: '16px', background: 'var(--bg)', position: 'relative' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '12px' }}>
                    <strong style={{ color: 'var(--accent)', fontSize: '14px', textTransform: 'uppercase' }}>
                      {idx + 1}. {step.type}
                    </strong>
                    <button onClick={() => removeStep(step.id)} style={{ background: 'none', border: 'none', color: 'var(--danger)', cursor: 'pointer', padding: 0 }}>
                      🗑️
                    </button>
                  </div>

                  {step.type === 'filter' && (
                    <div className="form-row" style={{ marginBottom: 0 }}>
                      <select value={step.col} onChange={e => updateStep(step.id, 'col', e.target.value)}>
                        {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                      <select value={step.op} onChange={e => updateStep(step.id, 'op', e.target.value)}>
                        <option value="equals">Equals</option>
                        <option value="contains">Contains</option>
                        <option value="greater">Greater Than</option>
                        <option value="less">Less Than</option>
                        <option value="not_blank">Is Not Blank</option>
                      </select>
                      <input type="text" value={step.val} onChange={e => updateStep(step.id, 'val', e.target.value)} placeholder="Value" disabled={step.op === 'not_blank'} />
                    </div>
                  )}

                  {step.type === 'combine' && (
                    <div className="form-row" style={{ marginBottom: 0, alignItems: 'center' }}>
                      <select value={step.cols[0]} onChange={e => updateStep(step.id, 'cols', [e.target.value, step.cols[1]])}>
                        {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                      <span style={{ fontWeight: 'bold' }}>+</span>
                      <select value={step.cols[1]} onChange={e => updateStep(step.id, 'cols', [step.cols[0], e.target.value])}>
                        {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                      <span style={{ fontWeight: 'bold', margin: '0 8px' }}>=</span>
                      <input type="text" value={step.newName} onChange={e => updateStep(step.id, 'newName', e.target.value)} placeholder="New Column Name" />
                    </div>
                  )}

                  {step.type === 'calc' && (
                    <div>
                      <div className="form-row" style={{ flexWrap: 'wrap', marginBottom: '10px' }}>
                        <select value={step.op} onChange={e => updateStep(step.id, 'op', e.target.value)}>
                          <option value="SUM_IF">SUM IF</option>
                          <option value="COUNT_IF">COUNT IF</option>
                        </select>
                        
                        {step.op === 'SUM_IF' && (
                          <select value={step.targetCol} onChange={e => updateStep(step.id, 'targetCol', e.target.value)}>
                            <option value="" disabled>Target Column</option>
                            {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        )}
                        
                        <span style={{ margin: '8px 4px', fontWeight: 'bold' }}>WHERE</span>
                        
                        <select value={step.condCol} onChange={e => updateStep(step.id, 'condCol', e.target.value)}>
                          {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                        
                        <select value={step.condOp} onChange={e => updateStep(step.id, 'condOp', e.target.value)}>
                          <option value="equals">Equals</option>
                          <option value="contains">Contains</option>
                          <option value="greater">Greater Than</option>
                          <option value="less">Less Than</option>
                          <option value="not_blank">Not Blank</option>
                        </select>
                        
                        <input type="text" value={step.condVal} onChange={e => updateStep(step.id, 'condVal', e.target.value)} placeholder="Condition Value" disabled={step.condOp === 'not_blank'} />
                      </div>
                      
                      <div className="form-row" style={{ flexWrap: 'wrap', alignItems: 'center', marginBottom: 0 }}>
                        <label style={{ marginRight: '10px', fontWeight: 600 }}>Output:</label>
                        <select value={step.outputMode} onChange={e => updateStep(step.id, 'outputMode', e.target.value)}>
                          <option value="ROW_BY_ROW">Row-by-Row</option>
                          <option value="GLOBAL_AGG">Global Aggregate</option>
                          <option value="PIVOT">Group By (Pivot)</option>
                        </select>
                        
                        {step.outputMode === 'PIVOT' && (
                          <>
                            <span style={{ margin: '0 8px' }}>GROUP BY:</span>
                            <select value={step.groupCol} onChange={e => updateStep(step.id, 'groupCol', e.target.value)}>
                              {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                            </select>
                          </>
                        )}
                        
                        <span style={{ margin: '0 8px', fontWeight: 'bold' }}>=</span>
                        <input type="text" value={step.newColName} onChange={e => updateStep(step.id, 'newColName', e.target.value)} placeholder="Result Col Name" />
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>

            <div style={{ position: 'relative' }}>
              <button className="secondary" onClick={() => setShowAddStep(!showAddStep)} style={{ width: '100%', borderStyle: 'dashed' }}>
                + Add Pipeline Step
              </button>
              
              {showAddStep && (
                <div style={{ position: 'absolute', top: '100%', left: 0, width: '100%', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '8px', padding: '8px', zIndex: 10, boxShadow: 'var(--shadow)', marginTop: '8px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <button className="secondary" onClick={() => addStep('filter')} style={{ textAlign: 'left', border: 'none' }}>🛡️ Filter Data</button>
                  <button className="secondary" onClick={() => addStep('combine')} style={{ textAlign: 'left', border: 'none' }}>🔗 Combine Columns</button>
                  <button className="secondary" onClick={() => addStep('calc')} style={{ textAlign: 'left', border: 'none' }}>🧮 Advanced Calculation</button>
                </div>
              )}
            </div>
          </div>

          {/* STEP 3 */}
          <div className="card">
            <h3 style={{ marginTop: 0 }}>Step 3: Save Configuration</h3>
            <div className="form-row">
              <div className="form-group" style={{ flex: 2 }}>
                <label>Configuration Name</label>
                <input type="text" value={configName} onChange={e => setConfigName(e.target.value)} placeholder="e.g., Monthly Sales Rules" />
              </div>
              <button onClick={saveReportConfig}>Save Config</button>
            </div>
          </div>

          {/* LIVE PREVIEW */}
          <div className="card">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <h3 style={{ margin: 0 }}>Pipeline Output Preview - {filteredRows.length} Rows</h3>
              <button onClick={handleExportCsv} className="secondary" style={{ borderColor: 'var(--accent)', color: 'var(--accent)', padding: '6px 12px' }}>
                ⬇️ Export CSV
              </button>
            </div>
            <div style={{ overflowX: 'auto', maxHeight: '400px' }}>
              <table className="data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                <thead>
                  <tr>
                    {currentPipelineCols.map(c => <th key={c} style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid var(--line)', background: '#f3f6f5', fontWeight: 600, textTransform: 'uppercase', fontSize: '11px', position: 'sticky', top: 0 }}>{c}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.slice(0, 50).map((row, idx) => (
                    <tr key={idx}>
                      {currentPipelineCols.map(c => <td key={c} style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid var(--line)' }}>{row[c]}</td>)}
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
    backgroundColor: 'var(--panel)',
    cursor: 'pointer',
    transition: 'border-color 0.2s',
  }
};

export default ConfigPage;
