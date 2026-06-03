import React, { useState, useEffect } from 'react';

const PAGE_SIZES = {
  A4: { width: 794, height: 1123 },
  A3: { width: 1123, height: 1588 },
  Letter: { width: 816, height: 1056 }
};

const GeneratePage = ({ dataset }) => {
  const [templates, setTemplates] = useState([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState('');
  const [mode, setMode] = useState('all'); // 'all' = Mail Merge, 'summary' = Table List
  const [generatedPages, setGeneratedPages] = useState([]);
  const [isGenerating, setIsGenerating] = useState(false);

  useEffect(() => {
    const saved = JSON.parse(localStorage.getItem('rs_templates') || '[]');
    setTemplates(saved);
    if (saved.length > 0) setSelectedTemplateId(saved[saved.length - 1].id);
  }, []);

  const handleGenerate = () => {
    if (!selectedTemplateId) return alert("Select a template first.");
    if (!dataset.rows || dataset.rows.length === 0) return alert("No active data. Go back to Report Config to upload data.");
    
    setIsGenerating(true);
    
    // Use timeout to allow UI to show loading state
    setTimeout(() => {
      const template = templates.find(t => t.id === selectedTemplateId);
      
      let width = PAGE_SIZES[template.pageSize].width;
      let height = PAGE_SIZES[template.pageSize].height;
      if (template.orientation === 'landscape') [width, height] = [height, width];
      
      const pages = [];
      
      if (mode === 'all') {
        // Mail Merge: One page per row
        dataset.rows.forEach(row => {
          pages.push({
            id: `page_${Math.random()}`,
            width,
            height,
            template,
            rowData: row,
            isSummary: false
          });
        });
      } else {
        // Summary: All rows in one template (data injected into first data table)
        pages.push({
          id: `page_summary`,
          width,
          height,
          template,
          rowData: dataset.rows,
          isSummary: true
        });
      }
      
      setGeneratedPages(pages);
      setIsGenerating(false);
    }, 50);
  };

  const handleExportJson = () => {
    if (generatedPages.length === 0) return alert("Generate a report first.");
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(generatedPages));
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", "report_data.json");
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  };

  const handleExportHtml = () => {
    if (generatedPages.length === 0) return alert("Generate a report first.");
    const container = document.querySelector('.print-container');
    if (!container) return;
    
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>Report Output</title>
        <style>
          body { font-family: Inter, sans-serif; background: #e5e7eb; padding: 40px; margin: 0; }
          .preview-page { background: white; margin: 0 auto 24px; position: relative; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border: 1px solid #ccc; overflow: hidden; page-break-after: always; }
          table { width: 100%; border-collapse: collapse; }
          @media print {
            body { background: white; padding: 0; }
            .preview-page { box-shadow: none; border: none; margin: 0; }
            @page { margin: 0; }
          }
        </style>
      </head>
      <body>
        ${container.innerHTML}
      </body>
      </html>
    `;
    
    const dataStr = "data:text/html;charset=utf-8," + encodeURIComponent(htmlContent);
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", "report_output.html");
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  };

  return (
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto', display: 'flex', flexDirection: 'column', height: '100%' }}>
      
      <div className="no-print">
        <h2>Generate Report</h2>
        <p className="subtitle">Combine your Active Dataset with a saved Template to generate your output.</p>

        <div className="card" style={{ display: 'flex', gap: '12px', alignItems: 'flex-end', background: 'white', padding: '20px', borderRadius: '12px' }}>
          <div className="form-group">
            <label>1. Active Dataset</label>
            <input type="text" disabled value={`${dataset.rows.length} rows loaded`} style={{ background: '#f5f5f5', color: 'var(--muted)' }} />
          </div>
          <div className="form-group">
            <label>2. Visual Template</label>
            <select value={selectedTemplateId} onChange={e => setSelectedTemplateId(e.target.value)}>
              <option value="">-- Select Template --</option>
              {templates.map(t => <option key={t.id} value={t.id}>{t.name} ({t.pageSize})</option>)}
            </select>
          </div>
          <div className="form-group">
            <label>3. Mode</label>
            <select value={mode} onChange={e => setMode(e.target.value)}>
              <option value="summary">Summary Report (All rows in table)</option>
              <option value="all">Mail Merge (One page per row)</option>
            </select>
          </div>
          <button onClick={handleGenerate}>{isGenerating ? 'Generating...' : 'Generate Output'}</button>
          
          {generatedPages.length > 0 && (
            <div style={{ display: 'flex', gap: '8px' }}>
              <button onClick={handleExportJson} className="secondary" style={{ borderColor: 'var(--accent)', color: 'var(--accent)' }}>
                {'{}'} JSON
              </button>
              <button onClick={handleExportHtml} className="secondary" style={{ borderColor: 'var(--accent)', color: 'var(--accent)' }}>
                🌐 HTML
              </button>
              <button onClick={() => window.print()} className="secondary" style={{ borderColor: 'var(--accent)', color: 'var(--accent)' }}>
                🖨️ PDF / Print
              </button>
            </div>
          )}
        </div>
      </div>

      {generatedPages.length === 0 ? (
        <div className="no-print" style={{ background: '#e5e7eb', padding: '40px', borderRadius: '12px', flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--muted)' }}>
          Select your configs and click Generate.
        </div>
      ) : (
        <div className="print-container" style={{ background: '#e5e7eb', padding: '40px', borderRadius: '12px', flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '24px', overflowY: 'auto' }}>
          {generatedPages.map((page, idx) => (
            <div 
              key={page.id} 
              className="preview-page" 
              style={{ 
                width: `${page.width}px`, 
                height: `${page.height}px`,
                background: 'white',
                position: 'relative',
                boxShadow: 'var(--shadow)',
                border: '1px solid var(--line)',
                overflow: 'hidden',
                flexShrink: 0,
                pageBreakAfter: 'always'
              }}
            >
              {page.template.bgImage && (
                <img src={page.template.bgImage} alt="bg" style={{ position: 'absolute', top: 0, left: 0, width: '100%', height: '100%', zIndex: 0, pointerEvents: 'none' }} />
              )}
              
              {page.template.elements.map(el => {
                
                // Determine Content
                let displayContent = el.content;
                if (el.type === 'field') {
                   displayContent = page.isSummary ? `[${el.content}]` : (page.rowData[el.content] || '');
                }

                return (
                  <div
                    key={el.id}
                    style={{
                      position: 'absolute',
                      left: el.x,
                      top: el.y,
                      width: el.width ? `${el.width}px` : 'auto',
                      height: el.height ? `${el.height}px` : 'auto',
                      fontSize: `${el.size || 16}px`,
                      fontWeight: el.bold ? 'bold' : 'normal',
                      fontStyle: el.italic ? 'italic' : 'normal',
                      textDecoration: el.underline ? 'underline' : 'none',
                      textAlign: el.align || 'left',
                      borderWidth: el.borderWidth ? `${el.borderWidth}px` : 0,
                      borderStyle: el.borderStyle || 'solid',
                      borderColor: 'var(--ink)',
                      zIndex: 10
                    }}
                  >
                     <div style={{ width: '100%', height: '100%' }}>
                       {el.type === 'text' || el.type === 'field' ? (
                         displayContent
                       ) : el.type === 'image' ? (
                         <img src={el.content} style={{ width: '100%', height: '100%', objectFit: 'contain' }} alt="Logo" />
                       ) : el.type === 'box' || el.type === 'line' ? (
                         <div style={{ width: '100%', height: '100%' }} />
                       ) : el.type === 'table' ? (
                         <table className="data-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: `${el.size}px` }}>
                           {el.tableType === 'data' && (
                             <thead>
                               <tr>
                                 {(el.columns || []).map(c => <th key={c} style={{ border: `${el.borderWidth}px ${el.borderStyle} var(--ink)`, padding: `${el.cellHeight}px` }}>{c}</th>)}
                               </tr>
                             </thead>
                           )}
                           <tbody>
                             {el.tableType === 'vertical' ? (
                               (el.columns || []).map(c => (
                                 <tr key={c}>
                                   <th style={{ border: `${el.borderWidth}px ${el.borderStyle} var(--ink)`, padding: `${el.cellHeight}px`, width: '40%', background: '#f3f6f5' }}>{c}</th>
                                   <td style={{ border: `${el.borderWidth}px ${el.borderStyle} var(--ink)`, padding: `${el.cellHeight}px` }}>
                                     {page.isSummary ? `[${c}]` : (page.rowData[c] || '')}
                                   </td>
                                 </tr>
                               ))
                             ) : el.tableType === 'data' && page.isSummary ? (
                               // SUMMARY MODE: DUMP ALL DATA
                               (Array.isArray(page.rowData) ? page.rowData : []).map((row, rIdx) => (
                                 <tr key={rIdx}>
                                   {(el.columns || []).map(c => <td key={c} style={{ border: `${el.borderWidth}px ${el.borderStyle} var(--ink)`, padding: `${el.cellHeight}px` }}>{row[c]}</td>)}
                                 </tr>
                               ))
                             ) : el.tableType === 'data' && !page.isSummary ? (
                               <tr>
                                 {(el.columns || []).map(c => <td key={c} style={{ border: `${el.borderWidth}px ${el.borderStyle} var(--ink)`, padding: `${el.cellHeight}px` }}>{page.rowData[c]}</td>)}
                               </tr>
                             ) : (
                               // BLANK ROWS
                               Array.from({ length: el.emptyRows || 5 }).map((_, rIdx) => (
                                 <tr key={rIdx}>
                                   {(el.columns || []).map(c => (
                                     <td key={c} style={{ border: `${el.borderWidth}px ${el.borderStyle} var(--ink)`, padding: `${el.cellHeight}px` }}>&nbsp;</td>
                                   ))}
                                 </tr>
                               ))
                             )}
                           </tbody>
                         </table>
                       ) : null}
                     </div>
                  </div>
                );
              })}
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default GeneratePage;
