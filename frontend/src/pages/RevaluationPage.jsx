import { useState, useEffect, useCallback } from 'react';
import { useNavigate, Link } from 'react-router-dom';
import { supabase } from '../supabaseClient';
import { useAuth } from '../context/AuthContext';
import { readSpreadsheetFile } from '../utils/excelParser';

const RevaluationPage = () => {
  const navigate = useNavigate();
  const { user } = useAuth();
  const [appFiles, setAppFiles] = useState([]);
  const [resultFile, setResultFile] = useState(null);
  const [mergedData, setMergedData] = useState([]);
  const [mergedCols, setMergedCols] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  
  // Database Saved Datasets States
  const [savedDatasets, setSavedDatasets] = useState([]);
  const [loadingDb, setLoadingDb] = useState(false);
  const [showSaveModal, setShowSaveModal] = useState(false);
  const [saveName, setSaveName] = useState('');
  const [isSaving, setIsSaving] = useState(false);

  // Fetch saved datasets from configs table
  const fetchSavedDatasets = useCallback(async () => {
    if (!user) return;
    setLoadingDb(true);
    try {
      // Fetch only metadata and rowCount to prevent downloading massive data arrays on page load
      const { data, error } = await supabase
        .from('configs')
        .select('id, name, created_at, config_data->rowCount')
        .eq('config_data->>isDataset', 'true')
        .eq('user_id', user.id);
      if (error) throw error;
      setSavedDatasets(data || []);
    } catch (err) {
      console.error("Error fetching datasets:", err);
    } finally {
      setLoadingDb(false);
    }
  }, [user]);

  useEffect(() => {
    fetchSavedDatasets();
  }, [fetchSavedDatasets]);

  // Parse helper
  const parseFile = async (file) => {
    try {
      const { rows } = await readSpreadsheetFile(file);
      return rows;
    } catch (err) {
      throw new Error(`Failed to read ${file.name}: ${err.message}`, { cause: err });
    }
  };

  const handleAppFilesChange = (e) => {
    const selected = Array.from(e.target.files);
    setAppFiles((prev) => [...prev, ...selected]);
  };

  const removeAppFile = (index) => {
    setAppFiles((prev) => prev.filter((_, i) => i !== index));
  };

  const findRowKey = (row, possibleNames) => {
    if (!row) return null;
    const keys = Object.keys(row);
    for (const name of possibleNames) {
      const found = keys.find(k => k.toLowerCase().replace(/[\s_-]/g, '') === name.toLowerCase().replace(/[\s_-]/g, ''));
      if (found) return found;
    }
    for (const name of possibleNames) {
      const found = keys.find(k => k.toLowerCase().includes(name.toLowerCase()));
      if (found) return found;
    }
    return null;
  };

  const matchesCondition = (row) => {
    const keys = Object.keys(row);
    const assessmentKeys = keys.filter(k => k.toLowerCase().includes('assessment'));
    
    let hasTH = false;
    let hasESE = false;
    
    for (const k of assessmentKeys) {
      const val = String(row[k]).trim().toUpperCase();
      if (k.toLowerCase().includes('type')) {
        if (val === 'TH') hasTH = true;
      } else if (k.toLowerCase().includes('method')) {
        if (val === 'ESE') hasESE = true;
      } else {
        if (val === 'TH') hasTH = true;
        if (val === 'ESE') hasESE = true;
      }
    }
    
    if (!hasTH || !hasESE) {
      for (const k of keys) {
        const val = String(row[k]).trim().toUpperCase();
        if (val === 'TH') hasTH = true;
        if (val === 'ESE') hasESE = true;
      }
    }
    
    return hasTH && hasESE;
  };

  const handleMerge = async () => {
    if (appFiles.length === 0) return alert("Upload at least one Application Report.");
    if (!resultFile) return alert("Upload the Result sheet.");
    
    setIsProcessing(true);
    try {
      let stackedApps = [];
      
      // Parse all Application Reports and stack them
      for (const file of appFiles) {
        const parsed = await parseFile(file);
        stackedApps = stackedApps.concat(parsed);
      }
      
      // Clean Keys (in case of whitespace)
      stackedApps = stackedApps.map(row => {
        const newRow = {};
        for(let key in row) newRow[key.trim()] = row[key];
        return newRow;
      });

      // Parse Result
      let resultData = await parseFile(resultFile);
      resultData = resultData.map(row => {
        const newRow = {};
        for(let key in row) newRow[key.trim()] = row[key];
        return newRow;
      });

      // Locate column keys dynamically
      const appPrnKey = findRowKey(stackedApps[0], ['PRN number', 'PRN', 'PRNNo', 'PRN No']);
      const appCourseCodeKey = findRowKey(stackedApps[0], ['course code', 'course_code', 'Course Code', 'Course Cod']);
      const resPrnKey = findRowKey(resultData[0], ['PRN', 'PRN number', 'PRNNo', 'PRN No']);
      const resCourseCodeKey = findRowKey(resultData[0], ['Course Cod', 'Course Code', 'CourseCode', 'course_code']);

      if (!appPrnKey || !appCourseCodeKey) {
        setIsProcessing(false);
        return alert(`Error: Could not locate 'PRN number' or 'course code' columns in Application Reports.`);
      }
      if (!resPrnKey || !resCourseCodeKey) {
        setIsProcessing(false);
        return alert(`Error: Could not locate 'PRN' or 'Course Code' columns in Result Sheet.`);
      }

      // Create a dictionary for fast Result lookup (only including matched Assessment conditions)
      const resultLookup = {};
      resultData.forEach(row => {
        if (!matchesCondition(row)) return;
        
        const prnVal = String(row[resPrnKey] || '').trim().toLowerCase();
        const courseCodeVal = String(row[resCourseCodeKey] || '').trim().toLowerCase();
        
        if (prnVal && courseCodeVal) {
          const matchKey = prnVal + "_" + courseCodeVal;
          resultLookup[matchKey] = row;
        }
      });

      // Left Join Application data with Result data on matchKey (PRN + Course Code)
      const merged = stackedApps.map(appRow => {
        const prnVal = String(appRow[appPrnKey] || '').trim().toLowerCase();
        const courseCodeVal = String(appRow[appCourseCodeKey] || '').trim().toLowerCase();
        const matchKey = prnVal + "_" + courseCodeVal;
        
        const resRow = resultLookup[matchKey] || {};
        
        // Merge - result fields override/supplement application fields
        return { ...appRow, ...resRow };
      });

      // Filter by TransactionStatus === 'Paid'
      const statusKey = findRowKey(merged[0] || {}, ['TransactionStatus', 'Transaction Status', 'PaymentStatus', 'Payment Status']);
      let filteredMerged = merged;
      if (statusKey) {
        filteredMerged = merged.filter(row => String(row[statusKey] || '').trim().toLowerCase() === 'paid');
      } else {
        // Fallback: Filter rows containing the value 'paid' in any column if exact header isn't found
        filteredMerged = merged.filter(row => 
          Object.values(row).some(val => String(val).trim().toLowerCase() === 'paid')
        );
      }

      if (filteredMerged.length === 0) {
        setIsProcessing(false);
        return alert("Merged dataset is empty after filtering for Paid transaction status.");
      }
      
      const columns = Array.from(new Set(filteredMerged.flatMap(Object.keys)));
      setMergedData(filteredMerged);
      setMergedCols(columns);
    } catch (err) {
      alert("Error processing files: " + err.message);
    }
    setIsProcessing(false);
  };

  const handleExportCsv = () => {
    if (mergedData.length === 0) return;
    const header = mergedCols.join(",");
    const csvContent = mergedData.map(row => 
      mergedCols.map(col => `"${String(row[col] || '').replace(/"/g, '""')}"`).join(",")
    ).join("\n");
    
    const dataStr = "data:text/csv;charset=utf-8," + encodeURIComponent(header + "\n" + csvContent);
    const a = document.createElement('a');
    a.href = dataStr;
    a.download = "Revaluation_Summary.csv";
    a.click();
  };

  const handleSendToStudio = () => {
    if (mergedData.length === 0) return;
    navigate('/studio', { 
      state: { 
        activePage: 'template', 
        dataset: { columns: mergedCols, rows: mergedData },
        stats: { rows: mergedData.length, cols: mergedCols.length } 
      } 
    });
  };

  const handleSaveToDatabase = async () => {
    if (!saveName.trim()) return alert("Please enter a name for the dataset.");
    if (!user) return alert("You must be logged in to save datasets.");

    setIsSaving(true);
    const payload = {
      isDataset: true,
      rowCount: mergedData.length,
      dataset: {
        columns: mergedCols,
        rows: mergedData
      },
      createdAt: new Date().toISOString()
    };

    const { error } = await supabase
      .from('configs')
      .insert([
        {
          name: saveName.trim(),
          config_data: payload,
          user_id: user.id
        }
      ]);

    setIsSaving(false);
    if (error) {
      alert("Error saving dataset: " + error.message);
    } else {
      alert("Dataset saved successfully!");
      setShowSaveModal(false);
      setSaveName('');
      fetchSavedDatasets(); // Update saved datasets list instantly
    }
  };

  const handleLoadSavedDataset = async (datasetInfo) => {
    if (!datasetInfo) return;
    setLoadingDb(true);
    try {
      // Fetch full dataset JSON on demand
      const { data, error } = await supabase
        .from('configs')
        .select('config_data')
        .eq('id', datasetInfo.id)
        .single();
        
      if (error) throw error;
      
      const configData = data?.config_data;
      const datasetObj = configData?.dataset;
      
      if (datasetObj) {
        setMergedCols(datasetObj.columns || []);
        setMergedData(datasetObj.rows || []);
        alert(`Loaded database dataset "${datasetInfo.name}" successfully! See preview below.`);
      } else {
        alert("Selected config is not a valid dataset.");
      }
    } catch (err) {
      alert("Error loading dataset: " + err.message);
    } finally {
      setLoadingDb(false);
    }
  };

  const handleDeleteSavedDataset = async (id) => {
    if (!window.confirm("Are you sure you want to delete this dataset?")) return;
    const { error } = await supabase.from('configs').delete().eq('id', id);
    if (error) {
      alert("Error deleting dataset: " + error.message);
    } else {
      fetchSavedDatasets();
    }
  };

  return (
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto', overflowY: 'auto', width: '100%', minHeight: '100vh', background: 'var(--bg)', color: 'var(--ink)' }}>
      {/* Navigation Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
        <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
          ← Back to Portal
        </Link>
      </div>

      <h2>Revaluation Process</h2>
      <p className="subtitle">Upload multiple Application Reports and a single Result Sheet. They will be stacked and merged automatically using the <strong>PRN number</strong>.</p>
      
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(350px, 1fr))', gap: '24px', marginBottom: '32px' }}>
        {/* Application Reports Box (Multi-upload) */}
        <div style={{ border: '1px dashed var(--line)', padding: '24px', borderRadius: '8px', background: 'var(--panel)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
          <strong style={{ display: 'block' }}>Application Reports (Upload multiple)</strong>
          <input 
            type="file" 
            multiple 
            accept=".csv, .xlsx, .xls" 
            onChange={handleAppFilesChange}
            id="multi-app-upload"
            style={{ display: 'none' }}
          />
          <label 
            htmlFor="multi-app-upload" 
            style={{ 
              display: 'inline-block', 
              padding: '12px 20px', 
              background: 'var(--accent-soft)', 
              color: 'var(--accent)', 
              borderRadius: '6px', 
              cursor: 'pointer', 
              textAlign: 'center', 
              fontWeight: '600',
              border: '1px solid var(--accent)'
            }}
          >
            Select Files
          </label>
          {appFiles.length > 0 && (
            <div style={{ marginTop: '12px', display: 'flex', flexDirection: 'column', gap: '6px', maxHeight: '180px', overflowY: 'auto' }}>
              {appFiles.map((file, idx) => (
                <div key={idx} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--bg)', padding: '6px 12px', borderRadius: '4px', border: '1px solid var(--line)', fontSize: '13px' }}>
                  <span style={{ textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap', maxWidth: '80%' }}>📄 {file.name}</span>
                  <button 
                    onClick={() => removeAppFile(idx)} 
                    style={{ background: 'transparent', color: 'var(--danger)', border: 'none', padding: '0 4px', fontSize: '16px', cursor: 'pointer' }}
                  >
                    ×
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Result Sheet Box (Single-upload) */}
        <div style={{ border: '1px dashed var(--line)', padding: '24px', borderRadius: '8px', background: 'var(--panel)', display: 'flex', flexDirection: 'column', gap: '12px', justifyContent: 'center' }}>
          <strong style={{ display: 'block' }}>Result Excel Sheet (Single file)</strong>
          {resultFile ? (
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--bg)', padding: '12px', borderRadius: '4px', border: '1px solid var(--line)' }}>
              <span style={{ fontWeight: 600, color: 'var(--accent)' }}>✅ {resultFile.name}</span>
              <button 
                onClick={() => setResultFile(null)} 
                style={{ background: 'transparent', color: 'var(--danger)', border: 'none', padding: '0 4px', fontSize: '16px', cursor: 'pointer' }}
              >
                ×
              </button>
            </div>
          ) : (
            <>
              <input 
                type="file" 
                accept=".csv, .xlsx, .xls" 
                onChange={e => setResultFile(e.target.files[0])}
                id="result-upload"
                style={{ display: 'none' }}
              />
              <label 
                htmlFor="result-upload" 
                style={{ 
                  display: 'inline-block', 
                  padding: '12px 20px', 
                  background: 'var(--accent-soft)', 
                  color: 'var(--accent)', 
                  borderRadius: '6px', 
                  cursor: 'pointer', 
                  textAlign: 'center', 
                  fontWeight: '600',
                  border: '1px solid var(--accent)'
                }}
              >
                Select Result Sheet
              </label>
            </>
          )}
        </div>

        {/* Database Saved Datasets */}
        <div style={{ border: '1px solid var(--line)', padding: '24px', borderRadius: '8px', background: 'var(--panel)', display: 'flex', flexDirection: 'column', gap: '12px' }}>
          <strong style={{ display: 'block' }}>Saved Tables in Database</strong>
          {loadingDb ? (
            <div style={{ fontSize: '13px', color: 'var(--muted)' }}>Loading saved tables...</div>
          ) : savedDatasets.length === 0 ? (
            <div style={{ fontSize: '13px', color: 'var(--muted)', background: 'var(--bg)', padding: '12px', borderRadius: '6px', textAlign: 'center' }}>No saved tables found.</div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', maxHeight: '180px', overflowY: 'auto' }}>
              {savedDatasets.map(d => {
                const rowCount = d.rowCount !== undefined ? d.rowCount : (d.config_data?.rowCount !== undefined ? d.config_data.rowCount : (d.config_data?.dataset?.rows?.length || 0));
                return (
                  <div key={d.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', background: 'var(--bg)', padding: '8px 12px', borderRadius: '6px', border: '1px solid var(--line)', fontSize: '12px' }}>
                    <div style={{ textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap', maxWidth: '60%' }}>
                      <span style={{ fontWeight: 600, display: 'block' }}>💾 {d.name}</span>
                      <span style={{ color: 'var(--muted)', fontSize: '11px' }}>{rowCount} rows</span>
                    </div>
                    <div style={{ display: 'flex', gap: '6px' }}>
                      <button onClick={() => handleLoadSavedDataset(d)} className="secondary" style={{ padding: '4px 8px', fontSize: '11px', borderColor: 'var(--accent)', color: 'var(--accent)' }}>Load</button>
                      <button onClick={() => handleDeleteSavedDataset(d.id)} className="danger" style={{ padding: '4px 8px', fontSize: '11px' }}>Delete</button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      </div>

      <button onClick={handleMerge} disabled={isProcessing} style={{ width: '100%', marginBottom: '32px', padding: '16px', fontSize: '16px' }}>
        {isProcessing ? "Merging Data..." : "Combine and Process"}
      </button>

      {mergedData.length > 0 && (
        <div className="card">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
            <h3 style={{ margin: 0 }}>Merged Result - {mergedData.length} Rows</h3>
            <div style={{ display: 'flex', gap: '12px' }}>
              <button onClick={handleExportCsv} className="secondary" style={{ borderColor: 'var(--accent)', color: 'var(--accent)' }}>⬇️ Export CSV</button>
              <button onClick={() => setShowSaveModal(true)} className="secondary" style={{ borderColor: 'var(--accent)', color: 'var(--accent)' }}>💾 Save to Database</button>
              <button onClick={handleSendToStudio}>🎨 Send to Template Creator</button>
            </div>
          </div>
          
          <div style={{ overflowX: 'auto', maxHeight: '400px' }}>
            <table className="data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
              <thead>
                <tr>
                  {mergedCols.map(c => <th key={c} style={{ padding: '8px', textAlign: 'left', borderBottom: '1px solid var(--line)', background: '#f3f6f5', position: 'sticky', top: 0 }}>{c}</th>)}
                </tr>
              </thead>
              <tbody>
                {mergedData.slice(0, 50).map((row, idx) => (
                  <tr key={idx}>
                    {mergedCols.map(c => <td key={c} style={{ padding: '8px', borderBottom: '1px solid var(--line)' }}>{row[c]}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Save Modal */}
      {showSaveModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0,0,0,0.5)',
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          zIndex: 1000,
          backdropFilter: 'blur(4px)'
        }}>
          <div className="card" style={{ width: '400px', background: 'var(--panel)', padding: '24px', borderRadius: '12px', border: '1px solid var(--line)' }}>
            <h3 style={{ marginTop: 0 }}>Save Combined Dataset</h3>
            <div className="form-group" style={{ marginBottom: '20px' }}>
              <label>Dataset Name</label>
              <input 
                type="text" 
                placeholder="e.g. Revaluation June 2026" 
                value={saveName} 
                onChange={e => setSaveName(e.target.value)} 
                style={{ width: '100%', marginTop: '4px' }}
              />
            </div>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '12px' }}>
              <button className="secondary" onClick={() => setShowSaveModal(false)} disabled={isSaving}>Cancel</button>
              <button onClick={handleSaveToDatabase} disabled={isSaving}>
                {isSaving ? "Saving..." : "Save"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default RevaluationPage;
