import { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { supabase } from '../supabaseClient';
import { readSpreadsheetFile } from '../utils/excelParser';
import { autoDetectDatasetColumns, suggestArchetype, TEMPLATE_ARCHETYPES } from '../utils/templateEngine';
import { 
  Sparkles, 
  FileText, 
  CalendarRange, 
  Tag, 
  TableProperties, 
  ArrowRight, 
  UploadCloud, 
  Filter, 
  Link2, 
  Calculator, 
  Calendar, 
  Trash2, 
  Eye
} from 'lucide-react';

const ConfigPage = ({ _onDataLoaded, dataset = { columns: [], rows: [] }, setDataset, setStats, initialConfig, onNavigateToTemplate }) => {
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
  const [showPasteModal, setShowPasteModal] = useState(false);
  const [pastedText, setPastedText] = useState('');

  const detectedArchetype = sourceCols.length > 0 ? suggestArchetype(autoDetectDatasetColumns(sourceCols)) : 'NOMINAL_ROLL';

  const runPipeline = useCallback((steps, initialRows, initialCols) => {
    let currentData = [...initialRows];
    let currentCols = [...initialCols];

    steps.forEach(step => {
      if (step.type === 'filter') {
        currentData = currentData.filter(row => {
          const cellVal = String(row[step.col] || '').trim().toLowerCase();
          const testVal = String(step.val || '').trim().toLowerCase();
          let match = true;

          if (step.op === 'equals') match = cellVal === testVal;
          else if (step.op === 'contains') match = cellVal.includes(testVal);
          else if (step.op === 'starts_with') match = cellVal.startsWith(testVal);
          else if (step.op === 'ends_with') match = cellVal.endsWith(testVal);
          else if (step.op === 'greater') match = parseFloat(cellVal) > parseFloat(testVal);
          else if (step.op === 'less') match = parseFloat(cellVal) < parseFloat(testVal);
          else if (step.op === 'not_blank') match = cellVal !== '';
          else if (step.op === 'is_blank') match = cellVal === '';

          return step.action === 'exclude' ? !match : match;
        });
      } 
      else if (step.type === 'combine') {
        const sep = step.separator !== undefined ? step.separator : ' - ';
        const prefix = step.prefix || '';
        const suffix = step.suffix || '';
        const newCol = step.newName || `Combined_${step.id.toString().slice(-4)}`;

        currentData = currentData.map(row => {
          const parts = (step.cols || []).map(c => {
            let val = String(row[c] || '').trim();
            if (step.formatDots && val) {
              val = val.replace(/\.+$/, '').trim() + '.';
            }
            return val;
          }).filter(v => v !== '');

          let combinedVal = parts.join(sep).trim();
          if (combinedVal) {
            combinedVal = `${prefix}${combinedVal}${suffix}`;
          }
          return { ...row, [newCol]: combinedVal };
        });

        if (!currentCols.includes(newCol)) currentCols.push(newCol);
      } 
      else if (step.type === 'calc') {
        const newCol = step.newColName || `Calc_${step.id.toString().slice(-4)}`;
        const op = step.op || 'SUM';

        const checkCondition = (row) => {
          if (op !== 'SUM_IF' && op !== 'COUNT_IF') return true;
          const cellVal = String(row[step.condCol] || '').toLowerCase();
          const testVal = String(step.condVal || '').toLowerCase();
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
            const tVal = parseFloat(row[step.targetCol]) || 0;
            const sVal = parseFloat(row[step.secondCol]) || parseFloat(step.constantVal) || 0;

            if (op === 'ADD') val = tVal + sVal;
            else if (op === 'SUBTRACT') val = tVal - sVal;
            else if (op === 'MULTIPLY') val = tVal * sVal;
            else if (op === 'DIVIDE') val = sVal !== 0 ? (tVal / sVal) : 0;
            else if (op === 'SUM_IF') val = checkCondition(row) ? tVal : 0;
            else if (op === 'COUNT_IF') val = checkCondition(row) ? 1 : 0;
            else if (op === 'SUM') val = tVal;
            else if (op === 'COUNT') val = 1;

            return { ...row, [newCol]: val };
          });
          if (!currentCols.includes(newCol)) currentCols.push(newCol);
        } 
        else if (step.outputMode === 'GLOBAL_AGG') {
          let total = 0;
          let count = 0;

          currentData.forEach(row => {
            const tVal = parseFloat(row[step.targetCol]) || 0;
            if (checkCondition(row)) {
              if (op === 'SUM' || op === 'SUM_IF') total += tVal;
              if (op === 'COUNT' || op === 'COUNT_IF') total += 1;
              if (op === 'AVERAGE') {
                total += tVal;
                count++;
              }
            }
          });

          if (op === 'AVERAGE' && count > 0) total = total / count;

          currentData = currentData.map(row => ({ ...row, [newCol]: total }));
          if (!currentCols.includes(newCol)) currentCols.push(newCol);
        }
        else if (step.outputMode === 'PIVOT') {
          const groups = {};
          const grpCol = step.groupCol || currentCols[0] || 'Group';

          currentData.forEach(row => {
            const groupKey = row[grpCol] || 'Unknown';
            if (!groups[groupKey]) {
              groups[groupKey] = { [grpCol]: groupKey, [newCol]: 0, _cnt: 0 };
            }
            const tVal = parseFloat(row[step.targetCol]) || 0;
            if (checkCondition(row)) {
              if (op === 'SUM' || op === 'SUM_IF') groups[groupKey][newCol] += tVal;
              if (op === 'COUNT' || op === 'COUNT_IF') groups[groupKey][newCol] += 1;
              if (op === 'AVERAGE') {
                groups[groupKey][newCol] += tVal;
                groups[groupKey]._cnt += 1;
              }
            }
          });

          if (op === 'AVERAGE') {
            Object.values(groups).forEach(g => {
              if (g._cnt > 0) g[newCol] = Math.round((g[newCol] / g._cnt) * 100) / 100;
              delete g._cnt;
            });
          }

          currentData = Object.values(groups);
          currentCols = [grpCol, newCol];
        }
      } 
      else if (step.type === 'extract_date') {
        const newCol = step.newName || `Date_${step.id.toString().slice(-4)}`;

        currentData = currentData.map(row => {
          const val = row[step.col] || '';
          let extracted = '';
          if (val) {
            const d = new Date(val);
            if (!isNaN(d.getTime())) {
              if (step.part === 'day_of_week') {
                extracted = d.toLocaleDateString(undefined, { weekday: 'long' });
              } else if (step.part === 'day_of_month') {
                extracted = String(d.getDate()).padStart(2, '0');
              } else if (step.part === 'month') {
                extracted = d.toLocaleDateString(undefined, { month: 'long' });
              } else if (step.part === 'year') {
                extracted = String(d.getFullYear());
              } else if (step.part === 'formatted_date') {
                extracted = d.toLocaleDateString('en-GB'); // DD/MM/YYYY
              }
            } else {
              extracted = val;
            }
          }
          return { ...row, [newCol]: extracted };
        });

        if (!currentCols.includes(newCol)) currentCols.push(newCol);
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

  const handlePasteDataSubmit = () => {
    if (!pastedText.trim()) return alert('Please paste your spreadsheet text.');
    try {
      const wb = XLSX.read(pastedText, { type: 'string', raw: true, dense: true });
      if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if (json && json.length > 0) {
          const cols = Object.keys(json[0]);
          setUploadStatus(`Loaded from Pasted Text (${json.length} rows, ${cols.length} columns)`);
          setSourceCols(cols);
          setSourceRows(json);
          setShowPasteModal(false);
          setPastedText('');
          return;
        }
      }
      alert('Could not parse valid tabular data from pasted text. Please verify the header row.');
    } catch (err) {
      alert('Error parsing text: ' + err.message);
    }
  };

  const addStep = (type) => {
    const newStep = { id: Date.now(), type };

    if (type === 'filter') {
      newStep.col = currentPipelineCols[0] || '';
      newStep.op = 'equals';
      newStep.val = '';
      newStep.action = 'keep'; // 'keep' | 'exclude'
    } 
    else if (type === 'combine') {
      newStep.cols = [currentPipelineCols[0] || '', currentPipelineCols[1] || currentPipelineCols[0] || ''];
      newStep.separator = ' - '; // In between
      newStep.prefix = '';       // Before
      newStep.suffix = '';       // After
      newStep.formatDots = false;
      newStep.newName = `Combined_${Date.now().toString().slice(-4)}`;
    } 
    else if (type === 'calc') {
      newStep.op = 'SUM'; // 'SUM' | 'COUNT' | 'AVERAGE' | 'ADD' | 'SUBTRACT' | 'MULTIPLY' | 'DIVIDE' | 'SUM_IF' | 'COUNT_IF'
      newStep.targetCol = currentPipelineCols[0] || '';
      newStep.secondCol = currentPipelineCols[1] || '';
      newStep.constantVal = '';
      newStep.condCol = currentPipelineCols[0] || '';
      newStep.condOp = 'equals';
      newStep.condVal = '';
      newStep.outputMode = 'ROW_BY_ROW'; // 'ROW_BY_ROW' | 'GLOBAL_AGG' | 'PIVOT'
      newStep.groupCol = currentPipelineCols[0] || '';
      newStep.newColName = `Total_${currentPipelineCols[0] || 'Calc'}`;
    } 
    else if (type === 'extract_date') {
      newStep.col = currentPipelineCols[0] || '';
      newStep.part = 'day_of_week';
      newStep.newName = `Day_${Date.now().toString().slice(-4)}`;
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

    let authUserId = null;
    try {
      const { data: authData } = await supabase.auth.getUser();
      if (authData?.user?.id) {
        authUserId = authData.user.id;
      }
    } catch {
      // not logged in with real auth
    }

    let savedToCloud = false;
    if (authUserId) {
      const { error } = await supabase
        .from('configs')
        .insert([
          { name: configName, config_data: configData, user_id: authUserId }
        ]);
      if (!error) {
        savedToCloud = true;
      }
    }

    // Always persist to local workspace / storage
    try {
      const localConfigs = JSON.parse(localStorage.getItem('saved_configs') || '[]');
      const newConfig = { id: `cfg_${Date.now()}`, name: configName, config_data: configData, created_at: new Date().toISOString() };
      const updated = [newConfig, ...localConfigs.filter(c => c.name !== configName)];
      localStorage.setItem('saved_configs', JSON.stringify(updated));
    } catch (e) {
      console.warn('LocalStorage error:', e);
    }

    if (savedToCloud) {
      alert('Configuration Saved to Cloud and Workspace!');
    } else {
      alert('Configuration Saved Successfully to your Workspace!');
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
      <p className="subtitle">Upload your examination spreadsheets, build data transformation pipelines, and launch customizable report templates.</p>

      {/* STEP 1: UPLOAD DATASET */}
      <div className="card">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px' }}>
          <h3 style={{ margin: 0 }}>Step 1: Upload Excel / Spreadsheet Data</h3>
          <button
            type="button"
            className="button secondary"
            onClick={() => setShowPasteModal(true)}
            style={{ fontSize: '12.5px', padding: '6px 12px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
          >
            📋 Paste Raw Text (TSV / CSV)
          </button>
        </div>
        
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

        {/* PASTE TEXT DATA MODAL */}
        {showPasteModal && (
          <div style={{
            position: 'fixed',
            top: 0, left: 0, right: 0, bottom: 0,
            background: 'rgba(0,0,0,0.5)',
            zIndex: 1000,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px'
          }}>
            <div style={{
              background: 'white',
              borderRadius: '12px',
              maxWidth: '800px',
              width: '100%',
              padding: '24px',
              boxShadow: '0 20px 50px rgba(0,0,0,0.2)'
            }}>
              <h3 style={{ margin: '0 0 8px 0', fontSize: '18px', fontWeight: 800 }}>Paste Raw TSV / CSV / Excel Data</h3>
              <p style={{ margin: '0 0 14px 0', fontSize: '13px', color: 'var(--muted)' }}>
                Paste raw tab-separated or comma-separated text copied from Excel, Google Sheets, or ERP export files:
              </p>
              <textarea
                rows={12}
                value={pastedText}
                onChange={(e) => setPastedText(e.target.value)}
                placeholder="Paste tabular rows here (including header row)..."
                style={{
                  width: '100%',
                  fontFamily: 'monospace',
                  fontSize: '12px',
                  padding: '12px',
                  borderRadius: '8px',
                  border: '1.5px solid var(--line)',
                  marginBottom: '16px',
                  boxSizing: 'border-box'
                }}
              />
              <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
                <button
                  type="button"
                  className="button secondary"
                  onClick={() => setShowPasteModal(false)}
                >
                  Cancel
                </button>
                <button
                  type="button"
                  className="button"
                  onClick={handlePasteDataSubmit}
                >
                  Load Data
                </button>
              </div>
            </div>
          </div>
        )}
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
            <h3 style={{ marginTop: 0 }}>Step 2: Data Pipeline (Filtering, Calculations & Concatenation)</h3>
            <p style={{ color: 'var(--muted)', fontSize: '13px', marginBottom: '24px' }}>
              Add transformation steps to clean, concatenate with custom symbols/text, extract dates, or perform math and aggregations.
            </p>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', marginBottom: '24px' }}>
              {pipelineSteps.map((step, idx) => (
                <div key={step.id} style={{ border: '1.5px solid var(--line)', borderRadius: '10px', padding: '18px', background: 'var(--bg)', position: 'relative' }}>
                  
                  {/* Step Header */}
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                      {step.type === 'filter' && <Filter size={16} color="var(--accent)" />}
                      {step.type === 'combine' && <Link2 size={16} color="var(--accent)" />}
                      {step.type === 'calc' && <Calculator size={16} color="var(--accent)" />}
                      {step.type === 'extract_date' && <Calendar size={16} color="var(--accent)" />}
                      
                      <strong style={{ color: 'var(--accent)', fontSize: '14px', textTransform: 'uppercase' }}>
                        {idx + 1}. {step.type === 'calc' ? 'Advanced Calculation / Aggregation' : step.type === 'combine' ? 'Combine & Concatenate' : step.type === 'extract_date' ? 'Date Extraction' : 'Filter Data'}
                      </strong>
                    </div>
                    <button 
                      onClick={() => removeStep(step.id)} 
                      style={{ background: 'none', border: 'none', color: 'var(--danger)', cursor: 'pointer', padding: '4px' }}
                      title="Remove Step"
                    >
                      <Trash2 size={16} />
                    </button>
                  </div>

                  {/* 1. FILTER STEP */}
                  {step.type === 'filter' && (
                    <div>
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: '10px', marginBottom: '8px' }}>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Filter Column:</label>
                          <select value={step.col} onChange={e => updateStep(step.id, 'col', e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                            {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Operator:</label>
                          <select value={step.op} onChange={e => updateStep(step.id, 'op', e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                            <option value="equals">Equals (=)</option>
                            <option value="contains">Contains</option>
                            <option value="starts_with">Starts With</option>
                            <option value="ends_with">Ends With</option>
                            <option value="greater">Greater Than (&gt;)</option>
                            <option value="less">Less Than (&lt;)</option>
                            <option value="not_blank">Not Blank (Has Value)</option>
                            <option value="is_blank">Is Blank (Empty)</option>
                          </select>
                        </div>
                        {step.op !== 'not_blank' && step.op !== 'is_blank' && (
                          <div>
                            <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Match Value:</label>
                            <input type="text" value={step.val} onChange={e => updateStep(step.id, 'val', e.target.value)} placeholder="Value to match..." style={{ width: '100%', fontSize: '13px' }} />
                          </div>
                        )}
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Action:</label>
                          <select value={step.action || 'keep'} onChange={e => updateStep(step.id, 'action', e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                            <option value="keep">Keep Matching Rows</option>
                            <option value="exclude">Exclude Matching Rows</option>
                          </select>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* 2. COMBINE / CONCATENATE STEP WITH PREFIX, SEPARATOR, SUFFIX */}
                  {step.type === 'combine' && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                      <p style={{ margin: '0 0 6px 0', fontSize: '12px', color: 'var(--muted)' }}>
                        Concatenate two or more columns with custom text or symbols in between, before (prefix), or after (suffix).
                      </p>

                      {/* Columns to Combine */}
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: '10px' }}>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>First Column:</label>
                          <select 
                            value={step.cols?.[0] || ''} 
                            onChange={e => {
                              const newCols = [...(step.cols || [])];
                              newCols[0] = e.target.value;
                              updateStep(step.id, 'cols', newCols);
                            }}
                            style={{ width: '100%', fontSize: '13px' }}
                          >
                            {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>

                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Second Column:</label>
                          <select 
                            value={step.cols?.[1] || ''} 
                            onChange={e => {
                              const newCols = [...(step.cols || [])];
                              newCols[1] = e.target.value;
                              updateStep(step.id, 'cols', newCols);
                            }}
                            style={{ width: '100%', fontSize: '13px' }}
                          >
                            <option value="">-- None --</option>
                            {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>

                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Third Column (Optional):</label>
                          <select 
                            value={step.cols?.[2] || ''} 
                            onChange={e => {
                              const newCols = [...(step.cols || [])];
                              newCols[2] = e.target.value;
                              updateStep(step.id, 'cols', newCols);
                            }}
                            style={{ width: '100%', fontSize: '13px' }}
                          >
                            <option value="">-- None --</option>
                            {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>
                      </div>

                      {/* Custom In-between Separator, Prefix, and Suffix */}
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(140px, 1fr))', gap: '10px', background: 'white', padding: '12px', borderRadius: '8px', border: '1px solid var(--line)' }}>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Prefix (Before):</label>
                          <input 
                            type="text" 
                            value={step.prefix || ''} 
                            onChange={e => updateStep(step.id, 'prefix', e.target.value)} 
                            placeholder="e.g. [ or Course: " 
                            style={{ width: '100%', fontSize: '12.5px' }} 
                          />
                        </div>

                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>In-between Separator / Symbol:</label>
                          <input 
                            type="text" 
                            value={step.separator !== undefined ? step.separator : ' - '} 
                            onChange={e => updateStep(step.id, 'separator', e.target.value)} 
                            placeholder="e.g.  -  or  •  or , " 
                            style={{ width: '100%', fontSize: '12.5px', fontWeight: 700 }} 
                          />
                        </div>

                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Suffix (After):</label>
                          <input 
                            type="text" 
                            value={step.suffix || ''} 
                            onChange={e => updateStep(step.id, 'suffix', e.target.value)} 
                            placeholder="e.g. ] or ." 
                            style={{ width: '100%', fontSize: '12.5px' }} 
                          />
                        </div>

                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>New Output Column Name:</label>
                          <input 
                            type="text" 
                            value={step.newName || ''} 
                            onChange={e => updateStep(step.id, 'newName', e.target.value)} 
                            placeholder="New Column Name" 
                            style={{ width: '100%', fontSize: '12.5px', fontWeight: 600 }} 
                          />
                        </div>
                      </div>

                      {/* Live Pattern Preview */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px', color: 'var(--ink)' }}>
                        <span style={{ fontWeight: 700, color: 'var(--accent)' }}>Pattern Preview:</span>
                        <code style={{ background: 'var(--accent-soft)', padding: '2px 8px', borderRadius: '4px', color: 'var(--accent)', fontWeight: 700 }}>
                          {step.prefix || ''}{step.cols?.[0] || 'COL_A'}{step.separator !== undefined ? step.separator : ' - '}{step.cols?.[1] || 'COL_B'}{step.suffix || ''}
                        </code>
                      </div>
                    </div>
                  )}

                  {/* 3. CALCULATION & AGGREGATION STEP */}
                  {step.type === 'calc' && (
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                      <p style={{ margin: '0 0 6px 0', fontSize: '12px', color: 'var(--muted)' }}>
                        Perform mathematical arithmetic (Add, Subtract, Multiply, Divide) or statistical aggregations (Sum, Count, Average, Sum-If, Count-If).
                      </p>

                      {/* Operation and Mode Selection */}
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: '10px' }}>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Operation / Math Function:</label>
                          <select value={step.op || 'SUM'} onChange={e => updateStep(step.id, 'op', e.target.value)} style={{ width: '100%', fontSize: '13px', fontWeight: 700 }}>
                            <optgroup label="Aggregations">
                              <option value="SUM">SUM (Total of Column)</option>
                              <option value="COUNT">COUNT (Number of Rows)</option>
                              <option value="AVERAGE">AVERAGE (Mean Value)</option>
                              <option value="SUM_IF">SUM_IF (Conditional Sum)</option>
                              <option value="COUNT_IF">COUNT_IF (Conditional Count)</option>
                            </optgroup>
                            <optgroup label="Row-by-Row Arithmetic">
                              <option value="ADD">ADD (Col A + Col B / Constant)</option>
                              <option value="SUBTRACT">SUBTRACT (Col A - Col B / Constant)</option>
                              <option value="MULTIPLY">MULTIPLY (Col A × Col B / Constant)</option>
                              <option value="DIVIDE">DIVIDE (Col A ÷ Col B / Constant)</option>
                            </optgroup>
                          </select>
                        </div>

                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Target Column:</label>
                          <select value={step.targetCol || ''} onChange={e => updateStep(step.id, 'targetCol', e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                            {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>

                        {/* Secondary column for math operations */}
                        {(['ADD', 'SUBTRACT', 'MULTIPLY', 'DIVIDE'].includes(step.op)) && (
                          <div>
                            <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Second Column (or enter constant):</label>
                            <div style={{ display: 'flex', gap: '6px' }}>
                              <select value={step.secondCol || ''} onChange={e => updateStep(step.id, 'secondCol', e.target.value)} style={{ flex: 1, fontSize: '13px' }}>
                                <option value="">-- Use Constant --</option>
                                {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                              </select>
                              {!step.secondCol && (
                                <input 
                                  type="number" 
                                  value={step.constantVal || ''} 
                                  onChange={e => updateStep(step.id, 'constantVal', e.target.value)} 
                                  placeholder="Value..." 
                                  style={{ width: '70px', fontSize: '13px' }} 
                                />
                              )}
                            </div>
                          </div>
                        )}

                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Output Mode:</label>
                          <select value={step.outputMode || 'ROW_BY_ROW'} onChange={e => updateStep(step.id, 'outputMode', e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                            <option value="ROW_BY_ROW">Row-by-Row (New Column in Each Row)</option>
                            <option value="GLOBAL_AGG">Global Total (Dataset Total Column)</option>
                            <option value="PIVOT">Pivot Summary (Grouped Totals Table)</option>
                          </select>
                        </div>

                        {step.outputMode === 'PIVOT' && (
                          <div>
                            <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Group By Column:</label>
                            <select value={step.groupCol || ''} onChange={e => updateStep(step.id, 'groupCol', e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                              {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                            </select>
                          </div>
                        )}

                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>New Output Column Name:</label>
                          <input type="text" value={step.newColName || ''} onChange={e => updateStep(step.id, 'newColName', e.target.value)} placeholder="New Column Name" style={{ width: '100%', fontSize: '13px', fontWeight: 600 }} />
                        </div>
                      </div>

                      {/* Condition Builder for SUM_IF and COUNT_IF */}
                      {(step.op === 'SUM_IF' || step.op === 'COUNT_IF') && (
                        <div style={{ background: 'white', padding: '12px', borderRadius: '8px', border: '1px solid var(--line)' }}>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '6px', color: 'var(--accent)' }}>
                            Condition Rule (Only compute when condition matches):
                          </label>
                          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(160px, 1fr))', gap: '10px' }}>
                            <select value={step.condCol || ''} onChange={e => updateStep(step.id, 'condCol', e.target.value)} style={{ fontSize: '12.5px' }}>
                              {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                            </select>
                            <select value={step.condOp || 'equals'} onChange={e => updateStep(step.id, 'condOp', e.target.value)} style={{ fontSize: '12.5px' }}>
                              <option value="equals">Equals</option>
                              <option value="contains">Contains</option>
                              <option value="greater">Greater Than</option>
                              <option value="less">Less Than</option>
                              <option value="not_blank">Not Blank</option>
                            </select>
                            {step.condOp !== 'not_blank' && (
                              <input type="text" value={step.condVal || ''} onChange={e => updateStep(step.id, 'condVal', e.target.value)} placeholder="Condition value..." style={{ fontSize: '12.5px' }} />
                            )}
                          </div>
                        </div>
                      )}
                    </div>
                  )}

                  {/* 4. DATE EXTRACTION STEP */}
                  {step.type === 'extract_date' && (
                    <div>
                      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(180px, 1fr))', gap: '10px' }}>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Date Column:</label>
                          <select value={step.col} onChange={e => updateStep(step.id, 'col', e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                            {sourceCols.map(c => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>Extract Part:</label>
                          <select value={step.part} onChange={e => updateStep(step.id, 'part', e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                            <option value="day_of_week">Day of Week (e.g. Wednesday)</option>
                            <option value="day_of_month">Day Number (e.g. 09)</option>
                            <option value="month">Month Name (e.g. September)</option>
                            <option value="year">Year (e.g. 2026)</option>
                            <option value="formatted_date">DD/MM/YYYY Format</option>
                          </select>
                        </div>
                        <div>
                          <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '4px' }}>New Column Name:</label>
                          <input type="text" value={step.newName} onChange={e => updateStep(step.id, 'newName', e.target.value)} placeholder="New Column Name" style={{ width: '100%', fontSize: '13px', fontWeight: 600 }} />
                        </div>
                      </div>
                    </div>
                  )}

                </div>
              ))}
            </div>

            {/* ADD STEP DROPDOWN BUTTONS */}
            <div style={{ position: 'relative' }}>
              <button 
                className="secondary" 
                onClick={() => setShowAddStep(!showAddStep)} 
                style={{ width: '100%', borderStyle: 'dashed', padding: '12px', fontSize: '13px', fontWeight: 700 }}
              >
                + Add Pipeline Transformation Step
              </button>
              
              {showAddStep && (
                <div style={{ position: 'absolute', top: '100%', left: 0, width: '100%', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '8px', padding: '8px', zIndex: 10, boxShadow: '0 8px 24px rgba(0,0,0,0.12)', marginTop: '8px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                  <button className="button secondary" onClick={() => addStep('combine')} style={{ textAlign: 'left', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', padding: '10px' }}>
                    <Link2 size={16} color="var(--accent)" />
                    <div>
                      <strong>Combine & Concatenate</strong>
                      <div style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 400 }}>Join columns with custom symbol/prefix/suffix</div>
                    </div>
                  </button>

                  <button className="button secondary" onClick={() => addStep('calc')} style={{ textAlign: 'left', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', padding: '10px' }}>
                    <Calculator size={16} color="var(--accent)" />
                    <div>
                      <strong>Advanced Calculation</strong>
                      <div style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 400 }}>Math (+, -, ×, ÷), Sum, Count, Average, Sum-If</div>
                    </div>
                  </button>

                  <button className="button secondary" onClick={() => addStep('filter')} style={{ textAlign: 'left', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', padding: '10px' }}>
                    <Filter size={16} color="var(--accent)" />
                    <div>
                      <strong>Filter Data</strong>
                      <div style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 400 }}>Keep or exclude matching records</div>
                    </div>
                  </button>

                  <button className="button secondary" onClick={() => addStep('extract_date')} style={{ textAlign: 'left', display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12.5px', padding: '10px' }}>
                    <Calendar size={16} color="var(--accent)" />
                    <div>
                      <strong>Extract Date Part</strong>
                      <div style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 400 }}>Extract Day of week, Month, Year</div>
                    </div>
                  </button>
                </div>
              )}
            </div>
          </div>

          {/* STEP 3: SAVE CONFIGURATION */}
          <div className="card">
            <h3 style={{ marginTop: 0 }}>Step 3: Save Data Pipeline Configuration</h3>
            <div className="form-row">
              <div className="form-group" style={{ flex: 2 }}>
                <label>Configuration Name</label>
                <input type="text" value={configName} onChange={e => setConfigName(e.target.value)} placeholder="e.g., Examination Data Pipeline" />
              </div>
              <button onClick={saveReportConfig}>Save Config</button>
            </div>
          </div>

          {/* LIVE DATA PREVIEW */}
          <div className="card">
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '8px' }}>
              <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Eye size={18} color="var(--accent)" /> Output Preview — {filteredRows.length} Rows, {currentPipelineCols.length} Columns
              </h3>
              <button onClick={handleExportCsv} className="button secondary" style={{ fontSize: '12px', padding: '6px 12px' }}>
                ⬇️ Export Cleaned CSV
              </button>
            </div>
            <div style={{ overflowX: 'auto', maxHeight: '400px', border: '1px solid var(--line)', borderRadius: '8px' }}>
              <table className="data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                <thead>
                  <tr>
                    {currentPipelineCols.map(c => <th key={c} style={{ padding: '12px', textAlign: 'left', borderBottom: '1px solid var(--line)', background: '#f3f6f5', fontWeight: 700, textTransform: 'uppercase', fontSize: '11px', position: 'sticky', top: 0 }}>{c}</th>)}
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.slice(0, 50).map((row, idx) => (
                    <tr key={idx}>
                      {currentPipelineCols.map(c => <td key={c} style={{ padding: '10px 12px', textAlign: 'left', borderBottom: '1px solid var(--line)' }}>{row[c]}</td>)}
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
