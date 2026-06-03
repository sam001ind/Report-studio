import React, { useState, useCallback } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

const SchedulerPage = () => {
  const [loading, setLoading] = useState(false);
  const [loadingText, setLoadingText] = useState("");
  const [stats, setStats] = useState({ rawRows: 0, students: 0, papers: 0 });
  const [rules, setRules] = useState([]);
  const [finalSlots, setFinalSlots] = useState([]);
  const [hierarchicalMap, setHierarchicalMap] = useState(new Map());
  
  // Ref for holding massive datasets without triggering endless re-renders
  const dataRef = React.useRef({
    masterDataset: [],
    uniquePrns: new Set(),
    uniquePapers: new Map(),
    enrollmentMap: new Map()
  });

  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setLoadingText("Scanning binary data packages...");
    
    try {
      const extension = file.name.split('.').pop().toLowerCase();
      let rawData = [];

      if (extension === 'zip') {
        const jszip = new JSZip();
        const unzipped = await jszip.loadAsync(file);
        for (const filename of Object.keys(unzipped.files)) {
          if (!unzipped.files[filename].dir && (filename.endsWith('.xlsx') || filename.endsWith('.xls') || filename.endsWith('.csv'))) {
            const blob = await unzipped.files[filename].async("blob");
            const sheetRows = await extractSpreadsheetRows(blob);
            rawData = rawData.concat(sheetRows);
          }
        }
      } else {
        rawData = await extractSpreadsheetRows(file);
      }
      
      alignAndConsolidateData(rawData);
    } catch (err) {
      console.error(err);
      alert("Processing Exception: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  const extractSpreadsheetRows = (fileBlob) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          let rows = [];
          workbook.SheetNames.forEach(name => {
            const sheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: "" });
            rows = rows.concat(sheetRows);
          });
          resolve(rows);
        } catch (ex) { reject(ex); }
      };
      reader.readAsArrayBuffer(fileBlob);
    });
  };

  const alignAndConsolidateData = (rawData) => {
    setLoadingText("Deduplicating entries and capturing valid parent groups...");
    
    const filteringMap = new Set();
    const consolidated = [];
    const uniquePrns = new Set();
    const uniquePapers = new Map();
    const newHierarchicalMap = new Map();
    const enrollmentMap = new Map();

    const strictValidMainGroups = ["AEC", "MDC", "SEC", "VAC", "DSC"];

    for (let i = 0; i < rawData.length; i++) {
      const r = rawData[i];
      const prn = r["PRN"] ? String(r["PRN"]).trim() : "";
      let code = r["PaperCode"] ? String(r["PaperCode"]).trim() : "";
      const name = r["PaperName"] ? String(r["PaperName"]).trim() : `Course Code ${code}`;

      if (!prn || !code) continue;

      // Strip trailing formatting artifacts/dots
      code = code.replace(/[.,\s_\-]+$/, ""); 

      const trackingKey = `${prn}_${code}`;
      if (!filteringMap.has(trackingKey)) {
        filteringMap.add(trackingKey);
        uniquePrns.add(prn);
        uniquePapers.set(code, name);
        consolidated.push({ prn, code, name, raw: r });
        
        if (!enrollmentMap.has(code)) enrollmentMap.set(code, new Set());
        enrollmentMap.get(code).add(prn);

        // Strict Two-Tier Regex Sorter Pass
        strictValidMainGroups.forEach(stemKey => {
          const stemIndex = code.indexOf(stemKey);
          if (stemIndex !== -1) {
            if (!newHierarchicalMap.has(stemKey)) {
              newHierarchicalMap.set(stemKey, new Set());
            }
            const remainingPart = code.substring(stemIndex + stemKey.length);
            const alphaSubMatch = remainingPart.match(/[A-Z]+/); 
            if (alphaSubMatch && alphaSubMatch[0].length > 0) {
              newHierarchicalMap.get(stemKey).add(alphaSubMatch[0]);
            }
          }
        });
      }
    }

    dataRef.current = {
      masterDataset: consolidated,
      uniquePrns,
      uniquePapers,
      enrollmentMap
    };

    setStats({
      rawRows: consolidated.length,
      students: uniquePrns.size,
      papers: uniquePapers.size
    });

    setHierarchicalMap(newHierarchicalMap);
  };

  const addRule = () => {
    setRules([...rules, { id: Date.now(), type: 'pool', mainGroup: '', subGroups: [] }]);
  };

  const updateRule = (id, field, value) => {
    setRules(rules.map(r => r.id === id ? { ...r, [field]: value } : r));
  };
  
  const removeRule = (id) => {
    setRules(rules.filter(r => r.id !== id));
  };

  const matchesTwoTierRule = (paperCode, rule) => {
    const mainIdx = paperCode.indexOf(rule.mainGroup);
    if (mainIdx === -1) return false;
    if (rule.subGroups.includes("ALL")) return true;
    const textAfterMain = paperCode.substring(mainIdx + rule.mainGroup.length);
    return rule.subGroups.some(sub => textAfterMain.startsWith(sub));
  };

  const buildTimetable = () => {
    setLoading(true);
    setLoadingText("Resolving non-overlapping graph matrices...");
    
    // Use timeout to allow UI to update loading state before heavy calculation
    setTimeout(() => {
      try {
        const { uniquePapers, enrollmentMap } = dataRef.current;
        let remainingPapers = Array.from(uniquePapers.keys());
        const structuredSlots = []; 

        // PASS 1: UNIVERSAL CONSOLIDATED GROUPS
        rules.filter(r => r.type === 'universal-group').forEach(rule => {
          if (!rule.mainGroup) return;
          const matches = remainingPapers.filter(code => matchesTwoTierRule(code, rule));
          if (matches.length > 0) {
            structuredSlots.push(matches);
            remainingPapers = remainingPapers.filter(c => !matches.includes(c));
          }
        });

        // PASS 2: ELECTIVE PARALLEL POOLS
        rules.filter(r => r.type === 'pool').forEach(rule => {
          if (!rule.mainGroup) return;
          const matches = remainingPapers.filter(code => matchesTwoTierRule(code, rule));
          if (matches.length > 0) {
            const conflictVerifiedPool = [];
            matches.forEach(code => {
              const codeStudents = enrollmentMap.get(code);
              let safeToPool = true;
              for(const allocated of conflictVerifiedPool) {
                const allocatedStudents = enrollmentMap.get(allocated);
                for(const prn of codeStudents) { 
                  if(allocatedStudents.has(prn)) { safeToPool = false; break; } 
                }
                if(!safeToPool) break;
              }
              if(safeToPool) { 
                conflictVerifiedPool.push(code); 
                remainingPapers = remainingPapers.filter(c => c !== code); 
              }
            });
            if(conflictVerifiedPool.length > 0) structuredSlots.push(conflictVerifiedPool);
          }
        });

        // PASS 3: CORE MAJOR DOMAIN PAPERS
        remainingPapers.sort((a, b) => enrollmentMap.get(b).size - enrollmentMap.get(a).size);
        remainingPapers.forEach(code => {
          const currentStudents = enrollmentMap.get(code);
          let placed = false;

          for (let i = 0; i < structuredSlots.length; i++) {
            const targetSlot = structuredSlots[i];
            
            // Do not merge into strict rule slots
            if (targetSlot.some(c => rules.some(rule => rule.mainGroup && matchesTwoTierRule(c, rule)))) continue;

            let conflict = false;
            for (const existingCode of targetSlot) {
              const existingStudents = enrollmentMap.get(existingCode);
              for (const student of currentStudents) { 
                if (existingStudents.has(student)) { conflict = true; break; } 
              }
              if (conflict) break;
            }
            if (!conflict) { 
              targetSlot.push(code); 
              placed = true; 
              break; 
            }
          }
          if (!placed) structuredSlots.push([code]);
        });

        setFinalSlots(structuredSlots);
      } catch (e) {
        console.error(e);
        alert("Scheduling Error: " + e.message);
      } finally {
        setLoading(false);
      }
    }, 50);
  };

  return (
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto' }}>
      <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '8px', marginBottom: '24px' }}>
        ← Back to Portal
      </Link>
      <h2 style={{ marginTop: 0 }}>Timetable Scheduler</h2>
      <p className="subtitle">Isolate structural blocks, map calendar execution dates, and generate sorted venue logs.</p>

      <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: '32px', marginBottom: '32px' }}>
        
        {/* Step 1: Upload */}
        <div style={{ background: 'white', padding: '24px', borderRadius: '12px', border: '1px solid var(--line)', boxShadow: 'var(--shadow)' }}>
          <h3 style={{ marginTop: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span>📁</span> Step 1: Drop & Import ZIP or Student Excel Dataset
          </h3>
          <p style={{ color: 'var(--muted)', fontSize: '14px', marginBottom: '16px' }}>
            Import your master dataset compressed folder (.zip) or workbook (.xlsx) directly using the restored stable file reader engine.
          </p>
          
          <label style={{
            display: 'block',
            border: '2px dashed var(--line)',
            borderRadius: '12px',
            padding: '40px',
            textAlign: 'center',
            cursor: 'pointer',
            background: 'var(--bg)',
            transition: 'all 0.2s'
          }}>
            <input type="file" style={{ display: 'none' }} accept=".zip,.xlsx,.xls,.csv" onChange={handleFileUpload} />
            <div style={{ fontSize: '32px', marginBottom: '8px' }}>📥</div>
            <div style={{ fontWeight: 600 }}>Click to browse or drop ZIP/Excel dataset here</div>
            <div style={{ fontSize: '12px', color: 'var(--muted)', marginTop: '4px' }}>Processes and consolidates structures safely in the browser</div>
          </label>

          {loading && (
            <div style={{ marginTop: '16px', padding: '12px', background: 'var(--accent-soft)', color: 'var(--accent)', borderRadius: '8px', fontWeight: 500, display: 'flex', gap: '8px', alignItems: 'center' }}>
              <span>⚙️</span> {loadingText}
            </div>
          )}
        </div>

        {/* Overview Stats */}
        <div style={{ background: 'white', padding: '24px', borderRadius: '12px', border: '1px solid var(--line)', boxShadow: 'var(--shadow)', display: 'flex', flexDirection: 'column', justifyContent: 'space-between' }}>
          <div>
            <h3 style={{ marginTop: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
              <span>📊</span> Structural Overview
            </h3>
            <div style={{ display: 'flex', justifyContent: 'space-between', padding: '12px 0', borderBottom: '1px solid var(--line)' }}>
              <span style={{ color: 'var(--muted)' }}>Total Valid Records:</span>
              <strong style={{ fontFamily: 'monospace' }}>{stats.rawRows || '-'}</strong>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', padding: '12px 0', borderBottom: '1px solid var(--line)' }}>
              <span style={{ color: 'var(--muted)' }}>Unique Active Students:</span>
              <strong style={{ fontFamily: 'monospace' }}>{stats.students || '-'}</strong>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', padding: '12px 0' }}>
              <span style={{ color: 'var(--muted)' }}>Identified Distinct Papers:</span>
              <strong style={{ fontFamily: 'monospace' }}>{stats.papers || '-'}</strong>
            </div>
          </div>
          
          <button 
            className="primary" 
            style={{ width: '100%', marginTop: '24px', opacity: stats.rawRows === 0 ? 0.5 : 1 }}
            disabled={stats.rawRows === 0}
            onClick={buildTimetable}
          >
            ⚡ Build Rule-Compliant Timetable
          </button>
        </div>
      </div>

      {/* Step 2: Rules */}
      <div style={{ background: 'white', padding: '24px', borderRadius: '12px', border: '1px solid var(--line)', boxShadow: 'var(--shadow)', marginBottom: '32px', opacity: stats.rawRows === 0 ? 0.5 : 1, pointerEvents: stats.rawRows === 0 ? 'none' : 'auto' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
          <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span>⚙️</span> Step 2: Configure Two-Tier Allocation Checklists
          </h3>
          <button className="secondary" onClick={addRule}>+ Add New Custom Rule</button>
        </div>
        <p style={{ color: 'var(--muted)', fontSize: '14px', marginBottom: '24px' }}>
          Click **Add New Custom Rule** to insert a configuration block. Select your discovered Main Group and pick which subgroup patterns should bundle together into isolated slots.
        </p>

        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(300px, 1fr))', gap: '16px' }}>
          {rules.map(rule => (
            <div key={rule.id} style={{ 
              padding: '16px', 
              borderRadius: '12px', 
              border: `1px solid ${rule.type === 'universal-group' ? '#fcd34d' : '#6ee7b7'}`,
              background: rule.type === 'universal-group' ? '#fffbeb' : '#ecfdf5' 
            }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: '12px' }}>
                <select 
                  value={rule.type} 
                  onChange={e => updateRule(rule.id, 'type', e.target.value)}
                  style={{ fontWeight: 600, padding: '4px 8px', borderRadius: '4px', border: '1px solid var(--line)' }}
                >
                  <option value="pool">Elective Parallel Pool</option>
                  <option value="universal-group">Universal Consolidated Day</option>
                </select>
                <button onClick={() => removeRule(rule.id)} style={{ background: 'transparent', border: 'none', color: '#ef4444', cursor: 'pointer', fontWeight: 'bold' }}>✕</button>
              </div>

              <div className="form-group" style={{ marginBottom: '12px' }}>
                <label style={{ fontSize: '11px', textTransform: 'uppercase', color: 'var(--muted)' }}>1. Main Group</label>
                <select 
                  value={rule.mainGroup} 
                  onChange={e => {
                    updateRule(rule.id, 'mainGroup', e.target.value);
                    updateRule(rule.id, 'subGroups', []); // Reset subgroups on change
                  }}
                  style={{ width: '100%', padding: '6px', borderRadius: '6px', border: '1px solid var(--line)' }}
                >
                  <option value="">-- Select Main Group --</option>
                  {Array.from(hierarchicalMap.keys()).sort().map(key => (
                    <option key={key} value={key}>{key}</option>
                  ))}
                </select>
              </div>

              <div className="form-group" style={{ marginBottom: 0 }}>
                <label style={{ fontSize: '11px', textTransform: 'uppercase', color: 'var(--muted)' }}>2. Refine Subgroups (Multi-Select)</label>
                <div style={{ background: 'white', border: '1px solid var(--line)', borderRadius: '6px', padding: '8px', maxHeight: '120px', overflowY: 'auto' }}>
                  {!rule.mainGroup ? (
                    <div style={{ fontSize: '12px', color: 'var(--muted)', fontStyle: 'italic' }}>Select a main group first...</div>
                  ) : (
                    <>
                      <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px', fontWeight: 'bold', color: 'var(--accent)', borderBottom: '1px solid var(--line)', paddingBottom: '4px', marginBottom: '4px' }}>
                        <input 
                          type="checkbox" 
                          checked={rule.subGroups.includes('ALL')}
                          onChange={(e) => updateRule(rule.id, 'subGroups', e.target.checked ? ['ALL'] : [])}
                        />
                        [Select All Subgroups]
                      </label>
                      {Array.from(hierarchicalMap.get(rule.mainGroup) || []).sort().map(sub => (
                        <label key={sub} style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px', color: 'var(--ink)' }}>
                          <input 
                            type="checkbox" 
                            checked={rule.subGroups.includes('ALL') || rule.subGroups.includes(sub)}
                            onChange={(e) => {
                              if (rule.subGroups.includes('ALL')) return; // handled by select all
                              const current = rule.subGroups;
                              const updated = e.target.checked 
                                ? [...current, sub] 
                                : current.filter(s => s !== sub);
                              updateRule(rule.id, 'subGroups', updated);
                            }}
                          />
                          <span style={{ fontFamily: 'monospace', background: 'var(--bg)', padding: '0 4px', border: '1px solid var(--line)', borderRadius: '4px' }}>{sub}</span>
                        </label>
                      ))}
                    </>
                  )}
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Output Grid */}
      {finalSlots.length > 0 && (
        <div style={{ background: 'white', padding: '24px', borderRadius: '12px', border: '1px solid var(--line)', boxShadow: 'var(--shadow)' }}>
          <h3 style={{ marginTop: 0 }}>🗓️ Master Plan Slots</h3>
          <p style={{ color: 'var(--muted)', fontSize: '14px', marginBottom: '24px' }}>Generated non-overlapping conflict-free slots. Assign formal calendar dates below.</p>
          
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(250px, 1fr))', gap: '16px' }}>
            {finalSlots.map((slotPapers, index) => (
              <div key={index} style={{ border: '1px solid var(--line)', borderRadius: '8px', overflow: 'hidden' }}>
                <div style={{ background: 'var(--bg)', padding: '8px 12px', borderBottom: '1px solid var(--line)', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <strong>Day {index + 1} Slot</strong>
                  <input type="date" style={{ padding: '4px', borderRadius: '4px', border: '1px solid var(--line)', fontSize: '12px' }} />
                </div>
                <div style={{ padding: '12px', maxHeight: '200px', overflowY: 'auto' }}>
                  {slotPapers.map(code => (
                    <div key={code} style={{ fontSize: '12px', fontFamily: 'monospace', padding: '4px 0', borderBottom: '1px dashed var(--line)' }}>
                      {code} <br/>
                      <span style={{ fontSize: '10px', color: 'var(--muted)', whiteSpace: 'normal', fontFamily: 'var(--font-family)' }}>
                        {dataRef.current.uniquePapers.get(code)}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

export default SchedulerPage;
