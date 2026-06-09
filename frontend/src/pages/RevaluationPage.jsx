import React, { useState } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';

const RevaluationPage = ({ setDataset, setStats, setActivePage }) => {
  const [files, setFiles] = useState({
    app1: null,
    app2: null,
    app3: null,
    result: null
  });
  const [mergedData, setMergedData] = useState([]);
  const [mergedCols, setMergedCols] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);

  // Parse helper
  const parseFile = (file) => {
    return new Promise((resolve, reject) => {
      const ext = file.name.split('.').pop().toLowerCase();
      if (ext === 'csv') {
        Papa.parse(file, { 
          header: true, 
          skipEmptyLines: true, 
          complete: (res) => resolve(res.data), 
          error: reject 
        });
      } else if (ext === 'xlsx' || ext === 'xls') {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const workbook = XLSX.read(e.target.result, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            resolve(data);
          } catch (err) { reject(err); }
        };
        reader.readAsArrayBuffer(file);
      } else {
        reject(new Error("Unsupported format"));
      }
    });
  };

  const handleMerge = async () => {
    if (!files.app1 && !files.app2 && !files.app3) return alert("Upload at least one Application Report.");
    if (!files.result) return alert("Upload the Result sheet.");
    
    setIsProcessing(true);
    try {
      let stackedApps = [];
      
      // Parse App Reports
      if (files.app1) stackedApps = stackedApps.concat(await parseFile(files.app1));
      if (files.app2) stackedApps = stackedApps.concat(await parseFile(files.app2));
      if (files.app3) stackedApps = stackedApps.concat(await parseFile(files.app3));
      
      // Clean Keys (in case of whitespace)
      stackedApps = stackedApps.map(row => {
        const newRow = {};
        for(let key in row) newRow[key.trim()] = row[key];
        return newRow;
      });

      // Parse Result
      let resultData = await parseFile(files.result);
      resultData = resultData.map(row => {
        const newRow = {};
        for(let key in row) newRow[key.trim()] = row[key];
        return newRow;
      });

      // Check for PRN number
      if (stackedApps.length > 0 && !Object.keys(stackedApps[0]).includes("PRN number")) {
         setIsProcessing(false);
         return alert("Error: Could not find 'PRN number' column in Application Reports.");
      }
      if (resultData.length > 0 && !Object.keys(resultData[0]).includes("PRN number")) {
         setIsProcessing(false);
         return alert("Error: Could not find 'PRN number' column in Result Sheet.");
      }

      // Create a dictionary for fast Result lookup
      const resultLookup = {};
      resultData.forEach(row => {
        const prn = String(row["PRN number"]).trim();
        resultLookup[prn] = row;
      });

      // Left Join Application data with Result data
      const merged = stackedApps.map(appRow => {
        const prn = String(appRow["PRN number"]).trim();
        const resRow = resultLookup[prn] || {};
        
        // Merge. To avoid key collision, we can prefix result columns or just overwrite.
        // Let's overwrite so Result columns take precedence if names conflict.
        return { ...appRow, ...resRow };
      });

      if (merged.length === 0) {
        setIsProcessing(false);
        return alert("Merged dataset is empty.");
      }
      
      const columns = Array.from(new Set(merged.flatMap(Object.keys)));
      setMergedData(merged);
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
    setDataset({ columns: mergedCols, rows: mergedData });
    setStats({ rows: mergedData.length, cols: mergedCols.length });
    setActivePage('template');
  };

  const FileUploader = ({ label, id }) => (
    <div style={{ border: '1px dashed var(--line)', padding: '24px', borderRadius: '8px', textAlign: 'center', background: 'var(--panel)' }}>
      <strong style={{ display: 'block', marginBottom: '8px' }}>{label}</strong>
      {files[id] ? (
        <div style={{ color: 'var(--accent)', fontWeight: 600 }}>{files[id].name}</div>
      ) : (
        <input type="file" accept=".csv, .xlsx, .xls" onChange={e => setFiles({ ...files, [id]: e.target.files[0] })} />
      )}
    </div>
  );

  return (
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto', overflowY: 'auto' }}>
      <h2>Revaluation Process</h2>
      <p className="subtitle">Upload your 3 Application Reports and the Result Sheet. They will be merged automatically using the <strong>PRN number</strong>.</p>
      
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '24px', marginBottom: '24px' }}>
        <FileUploader label="Application Report 1" id="app1" />
        <FileUploader label="Application Report 2" id="app2" />
        <FileUploader label="Application Report 3" id="app3" />
        <FileUploader label="Result Excel Sheet" id="result" />
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
    </div>
  );
};

export default RevaluationPage;
