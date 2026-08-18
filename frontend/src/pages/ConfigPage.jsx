import { useState, useEffect, useCallback } from 'react';
import { supabase } from '../supabaseClient';
import { useAuth } from '../context/AuthContext';
import { readSpreadsheetFile } from '../utils/excelParser';
import { autoDetectDatasetColumns, suggestArchetype, TEMPLATE_ARCHETYPES } from '../utils/templateEngine';
import { Sparkles, FileText, CalendarRange, Tag, TableProperties, ArrowRight, UploadCloud } from 'lucide-react';

const ConfigPage = ({ dataset, setDataset, setStats, initialConfig, onNavigateToTemplate }) => {
  const { user } = useAuth();
  const [uploadStatus, setUploadStatus] = useState('Click or Drag to Upload Spreadsheet (.xlsx, .xls, .csv, .zip)');
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  
  const [sourceRows, setSourceRows] = useState(dataset.rows || []);
  const [sourceCols, setSourceCols] = useState(dataset.columns || []);
  
  const [pipelineSteps, setPipelineSteps] = useState([]);
  const [currentPipelineCols, setCurrentPipelineCols] = useState(dataset.columns || []);
  const [filteredRows, setFilteredRows] = useState(dataset.rows || []);
  
  const [configName, setConfigName] = useState('');
  const [showAddStep, setShowAddStep] = useState(false);

  const detectedArchetype = sourceCols.length > 0 ? suggestArchetype(autoDetectDatasetColumns(sourceCols)) : 'NOMINAL_ROLL';

  const runPipeline = useCallback((steps, initialRows, initialCols) => {
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
      } else if (step.type === 'extract_date') {
        currentData = currentData.map(row => {
          const val = row[step.col] || '';
          let extracted = '';
          if (val) {
             const d = new Date(val);
             if (!isNaN(d)) {
               if (step.part === 'day_of_week') {
                 extracted = d.toLocaleDateString(undefined, { weekday: 'long' });
               } else if (step.part === 'day_of_month') {
                 extracted = d.getDate();
               } else if (step.part === 'month') {
                 extracted = d.toLocaleDateString(undefined, { month: 'long' });
               } else if (step.part === 'year') {
                 extracted = d.getFullYear();
               }
             }
          }
          return { ...row, [step.newName]: extracted };
        });
        if (!currentCols.includes(step.newName)) currentCols.push(step.newName);
      }
    });

    setCurrentPipelineCols(currentCols);
    setFilteredRows(currentData);
    setDataset({ columns: currentCols, rows: currentData });
    setStats({ rows: currentData.length, cols: currentCols.length });
  }, [setDataset, setStats]);

  // Load from initialConfig
  useEffect(() => {
    if (initialConfig && initialConfig.config_data) {
      setConfigName(initialConfig.name);
      const loadedSteps = initialConfig.config_data.pipeline || [];
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
  }, [pipelineSteps, sourceRows, sourceCols, runPipeline]);

  const processFile = async (file) => {
    if (!file) return;

    setUploadStatus(`Reading: ${file.name}...`);
    setIsProcessing(true);

    try {
      const { rows, columns } = await readSpreadsheetFile(file);
      if (!rows || rows.length === 0) {
        alert('File is empty.');
        setIsProcessing(false);
        return;
      }

      setUploadStatus(`Loaded: ${file.name} (${rows.length} rows, ${columns.length} columns)`);
      setSourceCols(columns);
      setSourceRows(rows);
    } catch (err) {
      console.error(err);
      alert('Error parsing file: ' + err.message);
      setUploadStatus('Failed to read file. Please try another .xlsx or .csv file.');
    } finally {
      setIsProcessing(false);
    }
  };

  const handleFileUpload = (e) => processFile(e.target.files[0]);
  const handleDragOver = (e) => { e.preventDefault(); setIsDragging(true); };
  const handleDragLeave = (e) => { e.preventDefault(); setIsDragging(false); };
  const handleDrop = (e) => { 
    e.preventDefault(); 
    setIsDragging(false); 
    if (e.dataTransfer.files?.length > 0) processFile(e.dataTransfer.files[0]); 
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
    } else if (type === 'extract_date') {
      newStep.col = currentPipelineCols[0] || '';
      newStep.part = 'day_of_week';
      newStep.newName = `DatePart_${Date.now().toString().slice(-4)}`;
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
    if (!configName) return alert('Enter a configuration name.');
    
    const configData = {
      pipeline: pipelineSteps,
      createdAt: new Date().toISOString()
    };
    
    const { error } = await supabase
      .from('configs')
      .insert([
        { name: configName, config_data: configData, user_id: user.id }
      ]);
      
    if (error) {
      console.error(error);
      alert('Error saving config to Supabase: ' + error.message);
    } else {
      alert('Configuration Saved to Cloud!');
    }
  };

  const handleExportCsv = () => {
    if (filteredRows.length === 0) return alert('No data to export.');
    const header = currentPipelineCols.join(',');
    const csvContent = filteredRows.map(row => 
      currentPipelineCols.map(col => `"${String(row[col] || '').replace(/"/g, '""')}"`).join(',')
    ).join('\n');
    const fullCsv = header + '\n' + csvContent;
    
    const dataStr = 'data:text/csv;charset=utf-8,' + encodeURIComponent(fullCsv);
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute('href', dataStr);
    downloadAnchorNode.setAttribute('download', (configName || 'filtered_data') + '.csv');
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  };

  return (
    <div style={styles.page}>
      <h2>Report Configuration & Dataset Ingestion</h2>
      <p className="subtitle">Upload any university or examination Excel spreadsheet, clean data, and launch customizable report templates.</p>

      {/* STEP 1: UPLOAD DATASET */}
      <div className="card">
        <h3 style={{ marginTop: 0 }}>Step 1: Upload Excel / Spreadsheet Data</h3>
        
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
            cursor: isProcessing ? 'wait' : 'pointer'
          }} 
          onClick={() => !isProcessing && document.getElementById('fileInput').click()}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
        >
          <UploadCloud size={44} color="var(--accent)" style={{ margin: '0 auto 12px', opacity: 0.8 }} />
          <h3 style={{ margin: '0 0 6px 0', fontSize: '17px' }}>{uploadStatus}</h3>
          <p style={{ color: 'var(--muted)', fontSize: '13.5px', margin: 0 }}>
            Supports binary <code>.xlsx</code>, <code>.xls</code>, <code>.csv</code>, <code>.tsv</code>, or <code>.zip</code> multi-sheet archives.
          </p>
          <input 
            type="file" 
            id="fileInput" 
            accept=".csv, .xlsx, .xls, .xlsm, .tsv, .zip" 
            style={{ display: 'none' }}
            onChange={handleFileUpload}
            disabled={isProcessing}
          />
        </div>
      </div>

      {/* RECOMMENDED REPORT TEMPLATES BANNER */}
      {sourceRows.length > 0 && (
        <div className="card" style={{ background: 'linear-gradient(135deg, rgba(23,107,135,0.06), #fff)', border: '1.5px solid var(--accent)', padding: '24px', marginBottom: '24px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px' }}>
            <div>
              <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Sparkles size={20} /> Choose or Customize a Report Template
              </h3>
              <p style={{ margin: '4px 0 0 0', color: 'var(--muted)', fontSize: '13px' }}>
                {sourceRows.length} records loaded. Pick a layout to open directly in the Report Template Studio:
              </p>
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: '14px' }}>
            {Object.values(TEMPLATE_ARCHETYPES).map(arch => {
              const isRecommended = arch.id === detectedArchetype;
              return (
                <div 
                  key={arch.id}
                  onClick={() => onNavigateToTemplate && onNavigateToTemplate(arch.id)}
                  style={{
                    padding: '16px',
                    borderRadius: '10px',
                    border: isRecommended ? '2px solid var(--accent)' : '1px solid var(--line)',
                    background: 'white',
                    cursor: 'pointer',
                    display: 'flex',
                    flexDirection: 'column',
                    justifyContent: 'space-between',
                    boxShadow: '0 2px 8px rgba(0,0,0,0.04)',
                    transition: 'all 0.2s ease'
                  }}
                >
                  <div>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                        {arch.id === 'NOMINAL_ROLL' && <FileText size={18} color="var(--accent)" />}
                        {arch.id === 'QP_STATEMENT' && <CalendarRange size={18} color="var(--accent)" />}
                        {arch.id === 'QP_COVER_LABEL' && <Tag size={18} color="var(--accent)" />}
                        {arch.id === 'CUSTOM_TABULAR' && <TableProperties size={18} color="var(--accent)" />}
                        <strong style={{ fontSize: '14px' }}>{arch.name}</strong>
                      </div>
                      {isRecommended && (
                        <span style={{ fontSize: '11px', fontWeight: 700, padding: '2px 8px', borderRadius: '12px', background: 'var(--accent-soft)', color: 'var(--accent)' }}>
                          Recommended
                        </span>
                      )}
                    </div>
                    <p style={{ margin: '0 0 14px 0', fontSize: '12px', color: 'var(--muted)', lineHeight: '1.4' }}>
                      {arch.description}
                    </p>
                  </div>

                  <button 
                    className="button"
                    style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '6px', fontSize: '12.5px', padding: '7px 12px', width: '100%' }}
                  >
                    Open in Template Studio <ArrowRight size={14} />
                  </button>
                </div>
              );
            })}
          </div>
        </div>
      )}
      
      {/* STEP 2 - PIPELINE */}
      {sourceCols.length > 0 && (
        <div id="dataMgmtSection">
          
          <div className="card">
            <h3 style={{ marginTop: 0 }}>Step 2: Data Pipeline (Optional Filtering & Calculations)</h3>
            <p style={{ color: 'var(--muted)', fontSize: '13px', marginBottom: '24px' }}>
              Add transformation steps if you need to filter rows, combine fields, or perform aggregations before generating reports.
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
                        <option value="not_blank">Not Blank</option>
                      </select>
                      {step.op !== 'not_blank' && (
                        <input type="text" value={step.val} onChange={e => updateStep(step.id, 'val', e.target.value)} placeholder="Filter value..." />
                      )}
                    </div>
                  )}

                  {step.type === 'combine' && (
                    <div className="form-row" style={{ marginBottom: 0 }}>
                      <select value={step.cols[0]} onChange={e => updateStep(step.id, 'cols', [e.target.value, step.cols[1]])}>
                        {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                      <select value={step.cols[1]} onChange={e => updateStep(step.id, 'cols', [step.cols[0], e.target.value])}>
                        {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                      <input type="text" value={step.newName} onChange={e => updateStep(step.id, 'newName', e.target.value)} placeholder="New Column Name" />
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
                  <button className="secondary" onClick={() => addStep('extract_date')} style={{ textAlign: 'left', border: 'none' }}>📅 Extract Date Part</button>
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
                <input type="text" value={configName} onChange={e => setConfigName(e.target.value)} placeholder="e.g., Examination Data Pipeline" />
              </div>
              <button onClick={saveReportConfig}>Save Config</button>
            </div>
          </div>

          {/* LIVE PREVIEW */}
          <div className="card">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
              <h3 style={{ margin: 0 }}>Data Output Preview - {filteredRows.length} Rows</h3>
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
