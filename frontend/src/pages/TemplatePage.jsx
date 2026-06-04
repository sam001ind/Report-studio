import React, { useState, useRef, useEffect } from 'react';
import mammoth from 'mammoth';
import { supabase } from '../supabaseClient';
import { useAuth } from '../context/AuthContext';

const PAGE_SIZES = {
  A4: { width: 794, height: 1123 },
  A3: { width: 1123, height: 1588 },
  Letter: { width: 816, height: 1056 }
};

const TemplatePage = ({ dataset, initialTemplate }) => {
  const { user } = useAuth();
  const [template, setTemplate] = useState(initialTemplate?.layout_data || {
    pageSize: 'A4',
    orientation: 'portrait',
    zoomScale: 0.75,
    snapToGrid: false,
    elements: [],
    bgImage: null,
    name: ''
  });
  
  const [selectedId, setSelectedId] = useState(null);
  const [previewRowIndex, setPreviewRowIndex] = useState(-1);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  
  // Dragging / Resizing State
  const actionRef = useRef(null);

  // Auto-save mechanism
  useEffect(() => {
    // We could save to localStorage here if we wanted persistent drafts
  }, [template]);

  useEffect(() => {
    if (initialTemplate) {
      setTemplate(initialTemplate.layout_data);
    }
  }, [initialTemplate]);

  const handleMouseDown = (e, id, type = 'drag') => {
    e.stopPropagation();
    setSelectedId(id);
    const el = template.elements.find(el => el.id === id);
    if (!el) return;
    
    actionRef.current = {
      type,
      id,
      startX: e.clientX,
      startY: e.clientY,
      initialX: el.x,
      initialY: el.y,
      initialW: el.width || 0,
      initialH: el.height || 0
    };
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  const handleMouseMove = (e) => {
    if (!actionRef.current) return;
    
    const { type, id, startX, startY, initialX, initialY, initialW, initialH } = actionRef.current;
    
    let dx = (e.clientX - startX) / template.zoomScale;
    let dy = (e.clientY - startY) / template.zoomScale;
    
    // Grid Snap
    if (template.snapToGrid) {
      dx = Math.round(dx / 10) * 10;
      dy = Math.round(dy / 10) * 10;
    }

    setTemplate(prev => ({
      ...prev,
      elements: prev.elements.map(el => {
        if (el.id !== id) return el;
        
        let newX = el.x;
        let newY = el.y;
        let newW = el.width;
        let newH = el.height;

        if (type === 'drag') {
          newX = initialX + dx;
          newY = initialY + dy;
          if (template.snapToGrid) {
            newX = Math.round(newX / 10) * 10;
            newY = Math.round(newY / 10) * 10;
          }
        } else if (type === 'resize-se') {
          newW = Math.max(20, initialW + dx);
          newH = Math.max(20, initialH + dy);
        } else if (type === 'resize-e') {
          newW = Math.max(20, initialW + dx);
        } else if (type === 'resize-s') {
          newH = Math.max(20, initialH + dy);
        }

        return { ...el, x: newX, y: newY, width: newW, height: newH };
      })
    }));
  };

  const handleMouseUp = () => {
    actionRef.current = null;
    document.removeEventListener('mousemove', handleMouseMove);
    document.removeEventListener('mouseup', handleMouseUp);
  };

  const generateId = () => 'el_' + Math.random().toString(36).substr(2, 9);

  // --- ADD ELEMENTS --- //
  const addElement = (elementData) => {
    setTemplate(prev => ({ ...prev, elements: [...prev.elements, elementData] }));
  };

  const addTextElement = () => addElement({ id: generateId(), type: 'text', content: 'Double click to edit', x: 50, y: 50, size: 16, align: 'left', bold: false, italic: false, underline: false });
  const addFieldElement = () => {
    if (!dataset.columns || dataset.columns.length === 0) return alert("Upload data first to see available fields.");
    addElement({ id: generateId(), type: 'field', content: dataset.columns[0], x: 50, y: 100, size: 16, align: 'left', bold: true, italic: false, underline: false });
  };
  const addTableElement = () => {
    if (!dataset.columns || dataset.columns.length === 0) return alert("Upload data first.");
    const defaultCols = dataset.columns.slice(0, Math.min(3, dataset.columns.length));
    addElement({ id: generateId(), type: 'table', columns: defaultCols, x: 50, y: 150, size: 12, align: 'left', width: 600, cellHeight: 12, borderWidth: 1, borderStyle: 'solid', tableType: 'data', emptyRows: 5 });
  };
  const addLineElement = () => addElement({ id: generateId(), type: 'line', x: 50, y: 200, width: 400, height: 10, borderWidth: 2, borderStyle: 'solid' });
  const addBoxElement = () => addElement({ id: generateId(), type: 'box', x: 50, y: 250, width: 300, height: 150, borderWidth: 1, borderStyle: 'solid' });

  // --- FILE IMPORTS --- //
  const handleLogoUpload = (e) => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      addElement({ id: generateId(), type: 'image', content: evt.target.result, x: 50, y: 50, width: 150 });
    };
    reader.readAsDataURL(file);
  };

  const handleBgUpload = (e) => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => setTemplate(prev => ({ ...prev, bgImage: evt.target.result }));
    reader.readAsDataURL(file);
  };

  const handleDocxImport = (e) => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      mammoth.extractRawText({ arrayBuffer: evt.target.result })
        .then(result => {
          const paragraphs = result.value.split('\n').map(p => p.trim()).filter(p => p.length > 0);
          if (paragraphs.length === 0) return alert("No readable text found.");
          
          let currentY = 50;
          const newElements = paragraphs.map(p => {
             const el = { id: generateId(), type: 'text', content: p, x: 50, y: currentY, size: 14, align: 'left', width: 600, bold: false, italic: false, underline: false };
             currentY += 40;
             return el;
          });
          setTemplate(prev => ({ ...prev, elements: [...prev.elements, ...newElements] }));
          alert(`Imported ${paragraphs.length} text blocks.`);
        })
        .catch(err => alert("Error parsing Word Doc: " + err.message));
    };
    reader.readAsArrayBuffer(file);
  };

  const handleImportTemplate = (e) => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const imported = JSON.parse(evt.target.result);
        if (imported.elements && Array.isArray(imported.elements)) {
          setTemplate(imported);
          alert("Template imported successfully!");
        } else {
          alert("Invalid template JSON structure.");
        }
      } catch (err) {
        alert("Error parsing JSON: " + err.message);
      }
    };
    reader.readAsText(file);
  };

  const handleHtmlImport = (e) => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const htmlString = evt.target.result;
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlString, 'text/html');
      
      let currentY = 50;
      const newElements = [];
      
      const processNode = (node) => {
        if (node.nodeType === Node.ELEMENT_NODE) {
          if (node.tagName === 'IMG') {
            newElements.push({ id: generateId(), type: 'image', content: node.src, x: 50, y: currentY, width: 200 });
            currentY += 150;
          } else if (['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'DIV', 'LI'].includes(node.tagName)) {
             const hasBlockChild = Array.from(node.children).some(c => ['P', 'DIV', 'H1', 'H2', 'UL', 'OL', 'LI', 'TABLE'].includes(c.tagName));
             
             let size = 14;
             let bold = false;
             if (node.tagName === 'H1') { size = 28; bold = true; }
             if (node.tagName === 'H2') { size = 24; bold = true; }
             if (node.tagName === 'H3') { size = 20; bold = true; }
             
             if (!hasBlockChild) {
                const text = node.textContent.trim();
                if (text) {
                  newElements.push({ id: generateId(), type: 'text', content: text, x: 50, y: currentY, size, align: 'left', width: 600, bold, italic: false, underline: false });
                  currentY += Math.max(40, size * 2);
                }
             } else {
                Array.from(node.childNodes).forEach(processNode);
             }
          } else {
            Array.from(node.childNodes).forEach(processNode);
          }
        } else if (node.nodeType === Node.TEXT_NODE) {
           const text = node.textContent.trim();
           if (text && node.parentNode.tagName === 'BODY') {
             newElements.push({ id: generateId(), type: 'text', content: text, x: 50, y: currentY, size: 14, align: 'left', width: 600, bold: false, italic: false, underline: false });
             currentY += 40;
           }
        }
      };
      
      processNode(doc.body);
      
      if (newElements.length > 0) {
        setTemplate(prev => ({ ...prev, elements: [...prev.elements, ...newElements] }));
        alert(`Imported ${newElements.length} HTML elements.`);
      } else {
        alert("No readable elements found in HTML.");
      }
    };
    reader.readAsText(file);
  };

  const handleUnifiedImport = (e) => {
    const file = e.target.files[0];
    if(!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    
    if (ext === 'docx') handleDocxImport(e);
    else if (ext === 'html') handleHtmlImport(e);
    else if (ext === 'json') handleImportTemplate(e);
    else alert("Unsupported format. Please upload .docx, .html, or .json");
  };

  const saveTemplate = async () => {
    if (!template.name) return alert("Enter a template name.");
    const layoutData = {
      ...template,
      createdAt: new Date().toISOString()
    };
    const { data, error } = await supabase
      .from('templates')
      .insert([{ name: template.name, layout_data: layoutData, user_id: user.id }]);
      
    if (error) {
      console.error(error);
      alert("Error saving to Supabase: " + error.message);
    } else {
      alert("Template Saved to Cloud!");
    }
  };

  const handleExportTemplate = () => {
    if (template.elements.length === 0) return alert("Nothing to export.");
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(template));
    const downloadAnchorNode = document.createElement('a');
    downloadAnchorNode.setAttribute("href", dataStr);
    downloadAnchorNode.setAttribute("download", (template.name || "template") + ".json");
    document.body.appendChild(downloadAnchorNode);
    downloadAnchorNode.click();
    downloadAnchorNode.remove();
  };

  // --- PROPERTIES BAR --- //
  const selectedElement = template.elements.find(el => el.id === selectedId);
  const updateSelected = (updates) => {
    if (!selectedId) return;
    setTemplate(prev => ({
      ...prev,
      elements: prev.elements.map(el => el.id === selectedId ? { ...el, ...updates } : el)
    }));
  };

  const deleteSelected = () => {
    if (!selectedId) return;
    setTemplate(prev => ({
      ...prev,
      elements: prev.elements.filter(el => el.id !== selectedId)
    }));
    setSelectedId(null);
  };

  const toggleStyle = (styleProp) => {
    if (!selectedElement) return;
    updateSelected({ [styleProp]: !selectedElement[styleProp] });
  };

  // --- RENDER LOGIC --- //
  let width = PAGE_SIZES[template.pageSize].width;
  let height = PAGE_SIZES[template.pageSize].height;
  if (template.orientation === 'landscape') [width, height] = [height, width];

  return (
    <div style={{ padding: '24px', height: '100%', display: 'flex', flexDirection: 'column' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
        <div>
          <h2 style={{ marginTop: 0 }}>Template Creator</h2>
          <p className="subtitle" style={{ marginBottom: '16px' }}>Set page size, design layout, format text, and place data fields exactly where you want.</p>
        </div>
        <button onClick={() => setIsSidebarOpen(!isSidebarOpen)} className="secondary" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          {isSidebarOpen ? '▶ Collapse Tools' : '◀ Expand Tools'}
        </button>
      </div>

      <div 
        className="builder-layout" 
        onClick={() => setSelectedId(null)}
        style={{ display: 'grid', gridTemplateColumns: isSidebarOpen ? '350px 1fr' : '0px 1fr', gap: isSidebarOpen ? '32px' : '0px', flex: 1, minHeight: 0, transition: 'all 0.3s ease' }}
      >
        
        {/* TOOLBOX SIDEBAR */}
        <div className="builder-tools" onClick={e => e.stopPropagation()} style={{ overflowY: 'auto', paddingRight: '12px', paddingBottom: '24px', opacity: isSidebarOpen ? 1 : 0, pointerEvents: isSidebarOpen ? 'auto' : 'none', display: isSidebarOpen ? 'block' : 'none' }}>
          <div className="form-row" style={{ marginBottom: 0 }}>
             <div className="form-group">
               <label>Page Size</label>
               <select value={template.pageSize} onChange={e => setTemplate({...template, pageSize: e.target.value})}>
                 <option value="A4">A4</option>
                 <option value="A3">A3</option>
                 <option value="Letter">US Letter</option>
               </select>
             </div>
             <div className="form-group">
               <label>Orientation</label>
               <select value={template.orientation} onChange={e => setTemplate({...template, orientation: e.target.value})}>
                 <option value="portrait">Portrait</option>
                 <option value="landscape">Landscape</option>
               </select>
             </div>
             <div className="form-group">
               <label>Zoom</label>
               <select value={template.zoomScale} onChange={e => setTemplate({...template, zoomScale: parseFloat(e.target.value)})}>
                 <option value="1">100%</option>
                 <option value="0.75">75%</option>
                 <option value="0.5">50%</option>
               </select>
             </div>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginTop: '12px', fontSize: '13px', fontWeight: 500 }}>
             <input type="checkbox" id="toggleGrid" checked={template.snapToGrid} onChange={e => setTemplate({...template, snapToGrid: e.target.checked})} />
             <label htmlFor="toggleGrid" style={{ cursor: 'pointer' }}>Show Ruler Grid & Snap</label>
          </div>

          <div className="form-group" style={{ background: 'var(--bg)', padding: '12px', borderRadius: '8px', border: '1px solid var(--line)', marginTop: '16px' }}>
             <label>Live Data Preview (From Config)</label>
             <select value={previewRowIndex} onChange={e => setPreviewRowIndex(parseInt(e.target.value))}>
               <option value="-1">Show Placeholders</option>
               {dataset.rows.slice(0, 20).map((row, i) => (
                 <option key={i} value={i}>Row {i + 1}</option>
               ))}
             </select>
          </div>

          <hr style={{ border: 0, borderTop: '1px solid var(--line)', margin: '8px 0' }} />

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '8px' }}>
             <button onClick={addTextElement} className="secondary" style={{ padding: '8px', fontSize: '12px' }}>+ Text</button>
             <button onClick={addFieldElement} className="secondary" style={{ padding: '8px', fontSize: '12px', borderColor: 'var(--accent)', color: 'var(--accent)' }}>+ Field</button>
             <button onClick={addTableElement} className="secondary" style={{ padding: '8px', fontSize: '12px', borderColor: 'var(--accent)', color: 'var(--accent)' }}>+ Table</button>
             
             <button onClick={addBoxElement} className="secondary" style={{ padding: '8px', fontSize: '12px' }}>+ Box</button>
             <button onClick={addLineElement} className="secondary" style={{ padding: '8px', fontSize: '12px' }}>+ Line</button>
             <label className="secondary" style={{ padding: '8px', fontSize: '12px', textAlign: 'center', cursor: 'pointer', borderRadius: '8px', fontWeight: 600, border: '1px solid var(--line)' }}>
               + Logo <input type="file" accept="image/*" onChange={handleLogoUpload} style={{ display: 'none' }} />
             </label>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '8px', marginTop: '8px' }}>
             <label className="secondary" style={{ padding: '8px', fontSize: '12px', textAlign: 'center', cursor: 'pointer', borderRadius: '8px', fontWeight: 600, border: '1px solid var(--line)' }}>
               + Form Background Image <input type="file" accept="image/*" onChange={handleBgUpload} style={{ display: 'none' }} />
             </label>
             <label className="secondary" style={{ padding: '8px', fontSize: '12px', textAlign: 'center', cursor: 'pointer', borderRadius: '8px', fontWeight: 600, border: '1px solid var(--accent)', color: 'var(--accent)' }}>
               + Import File (.docx, .html, .json) <input type="file" accept=".docx,.html,.json" onChange={handleUnifiedImport} style={{ display: 'none' }} />
             </label>
             {template.bgImage && (
               <button onClick={() => setTemplate(prev => ({...prev, bgImage: null}))} className="danger" style={{ padding: '8px', fontSize: '12px' }}>Clear Background</button>
             )}
          </div>

          <hr style={{ border: 0, borderTop: '1px solid var(--line)', margin: '8px 0' }} />

          {/* DYNAMIC PROPERTIES PANEL */}
          {selectedElement ? (
            <div style={{ background: '#fafafa', padding: '12px', borderRadius: '8px', border: '1px solid var(--line)' }}>
              <h4 style={{ margin: '0 0 12px', color: 'var(--accent)' }}>Element Properties</h4>
              
              {(selectedElement.type === 'text' || selectedElement.type === 'field') && (
                <div className="form-group">
                  <label>{selectedElement.type === 'text' ? 'Text Content' : 'Data Column'}</label>
                  {selectedElement.type === 'text' ? (
                    <input type="text" value={selectedElement.content} onChange={e => updateSelected({ content: e.target.value })} />
                  ) : (
                    <select value={selectedElement.content} onChange={e => updateSelected({ content: e.target.value })}>
                      {dataset.columns.map(col => <option key={col} value={col}>{col}</option>)}
                    </select>
                  )}
                </div>
              )}

              {selectedElement.type === 'table' && (
                <>
                  <div className="form-group" style={{ marginBottom: '12px' }}>
                    <label>Table Format</label>
                    <select value={selectedElement.tableType} onChange={e => updateSelected({ tableType: e.target.value })}>
                      <option value="data">Data List (Rows)</option>
                      <option value="vertical">Form (Key-Value)</option>
                      <option value="blank">Blank Grid</option>
                    </select>
                  </div>
                  {selectedElement.tableType === 'blank' && (
                    <div className="form-group" style={{ marginBottom: '12px' }}>
                      <label>Empty Rows</label>
                      <input type="number" value={selectedElement.emptyRows} onChange={e => updateSelected({ emptyRows: parseInt(e.target.value) || 1 })} />
                    </div>
                  )}
                  <div className="form-group">
                    <label>Table Columns</label>
                    <div className="column-checkbox-list" style={{ background: 'white' }}>
                      {dataset.columns.map(col => (
                        <label key={col}>
                          <input 
                            type="checkbox" 
                            checked={(selectedElement.columns || []).includes(col)}
                            onChange={(e) => {
                              const cols = selectedElement.columns || [];
                              updateSelected({ columns: e.target.checked ? [...cols, col] : cols.filter(c => c !== col) });
                            }} 
                          /> 
                          {col}
                        </label>
                      ))}
                    </div>
                  </div>
                </>
              )}

              {/* Typography Controls */}
              {['text', 'field', 'table'].includes(selectedElement.type) && (
                <div style={{ marginTop: '12px', display: 'flex', gap: '4px' }}>
                  <button onClick={() => toggleStyle('bold')} className={`secondary style-toggle-btn ${selectedElement.bold ? 'active' : ''}`}><b>B</b></button>
                  <button onClick={() => toggleStyle('italic')} className={`secondary style-toggle-btn ${selectedElement.italic ? 'active' : ''}`}><i>I</i></button>
                  <button onClick={() => toggleStyle('underline')} className={`secondary style-toggle-btn ${selectedElement.underline ? 'active' : ''}`}><u>U</u></button>
                  <select value={selectedElement.align || 'left'} onChange={e => updateSelected({ align: e.target.value })} style={{ flex: 1, padding: '4px 8px' }}>
                    <option value="left">Left Align</option>
                    <option value="center">Center Align</option>
                    <option value="right">Right Align</option>
                  </select>
                </div>
              )}

              {/* Dimension Controls */}
              <div className="form-row" style={{ marginTop: '12px', marginBottom: 0 }}>
                {['box', 'line', 'image', 'table', 'text'].includes(selectedElement.type) && (
                  <div className="form-group">
                    <label>Width (px)</label>
                    <input type="number" value={selectedElement.width || ''} onChange={e => updateSelected({ width: parseInt(e.target.value) || undefined })} placeholder="Auto" />
                  </div>
                )}
                {['box', 'line', 'image'].includes(selectedElement.type) && (
                  <div className="form-group">
                    <label>Height (px)</label>
                    <input type="number" value={selectedElement.height || ''} onChange={e => updateSelected({ height: parseInt(e.target.value) || undefined })} placeholder="Auto" />
                  </div>
                )}
                {['table'].includes(selectedElement.type) && (
                  <div className="form-group">
                    <label>Cell Pad (px)</label>
                    <input type="number" value={selectedElement.cellHeight || 12} onChange={e => updateSelected({ cellHeight: parseInt(e.target.value) || 12 })} />
                  </div>
                )}
              </div>

              {/* Styling Controls */}
              <div className="form-row" style={{ marginTop: '12px', marginBottom: 0 }}>
                {['text', 'field', 'table'].includes(selectedElement.type) && (
                  <div className="form-group">
                    <label>Font Size</label>
                    <input type="number" value={selectedElement.size || 16} onChange={e => updateSelected({ size: parseInt(e.target.value) || 16 })} />
                  </div>
                )}
                {['box', 'line', 'table'].includes(selectedElement.type) && (
                  <div className="form-group">
                    <label>Border (px)</label>
                    <input type="number" value={selectedElement.borderWidth || 0} onChange={e => updateSelected({ borderWidth: parseInt(e.target.value) || 0 })} />
                  </div>
                )}
                {['box', 'line', 'table'].includes(selectedElement.type) && (
                  <div className="form-group">
                    <label>Border Style</label>
                    <select value={selectedElement.borderStyle || 'solid'} onChange={e => updateSelected({ borderStyle: e.target.value })}>
                      <option value="solid">Solid</option>
                      <option value="dashed">Dashed</option>
                      <option value="dotted">Dotted</option>
                      <option value="none">None</option>
                    </select>
                  </div>
                )}
              </div>

              <button className="danger" onClick={deleteSelected} style={{ width: '100%', marginTop: '16px' }}>Delete Element</button>
            </div>
          ) : (
            <div style={{ fontSize: '13px', color: 'var(--muted)', textAlign: 'center', padding: '20px 0' }}>
              Select an element to edit properties.
            </div>
          )}

          <hr style={{ border: 0, borderTop: '1px solid var(--line)', margin: '8px 0' }} />
          
          <div className="form-group">
            <label>Template Name</label>
            <input type="text" value={template.name} onChange={e => setTemplate({...template, name: e.target.value})} placeholder="e.g., Corporate Invoice" />
          </div>
          <div style={{ display: 'flex', gap: '8px' }}>
             <button onClick={saveTemplate} style={{ flex: 1 }}>Save</button>
             <button onClick={handleExportTemplate} className="secondary" style={{ flex: 1, borderColor: 'var(--accent)', color: 'var(--accent)' }}>Export</button>
          </div>
        </div>

        {/* CANVAS WORKSPACE */}
        <div className="canvas-container" style={{ overflow: 'auto', background: '#e5e7eb', padding: '40px', borderRadius: '12px', textAlign: 'center', height: '100%', boxSizing: 'border-box' }}>
          <div 
            className={`paper-wrapper ${template.snapToGrid ? 'show-grid' : ''}`}
            style={{ 
              width: `${width}px`, 
              height: `${height}px`, 
              transform: `scale(${template.zoomScale})`, 
              transformOrigin: 'top center',
              marginBottom: `-${height * (1 - template.zoomScale)}px`
            }}
          >
             <div className="ruler-corner"></div>
             <div className="ruler-x"></div>
             <div className="ruler-y"></div>
             
             <div className={`paper ${template.snapToGrid ? 'show-grid' : ''}`}>
               {template.bgImage && (
                 <img src={template.bgImage} className="page-bg" alt="Form Background" />
               )}

               {template.elements.map(el => {
                 
                 // Get display content based on Live Data Preview
                 let displayContent = el.content;
                 if (el.type === 'field') {
                    if (previewRowIndex >= 0 && dataset.rows[previewRowIndex]) {
                      displayContent = dataset.rows[previewRowIndex][el.content] || '';
                    } else {
                      displayContent = `[${el.content}]`;
                    }
                 }

                 const isSelected = selectedId === el.id;

                 return (
                   <div
                     key={el.id}
                     className={`drag-element ${isSelected ? 'selected' : ''} ${el.type === 'field' && previewRowIndex === -1 ? 'field-type' : ''}`}
                     onMouseDown={(e) => handleMouseDown(e, el.id, 'drag')}
                     style={{
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
                       borderColor: 'var(--ink)'
                     }}
                   >
                     {/* Resize Handles */}
                     <div className="resize-handle se" onMouseDown={(e) => handleMouseDown(e, el.id, 'resize-se')} />
                     <div className="resize-handle e" onMouseDown={(e) => handleMouseDown(e, el.id, 'resize-e')} />
                     <div className="resize-handle s" onMouseDown={(e) => handleMouseDown(e, el.id, 'resize-s')} />

                     <div className="content-wrap">
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
                                     {previewRowIndex >= 0 ? dataset.rows[previewRowIndex]?.[c] : `[${c}]`}
                                   </td>
                                 </tr>
                               ))
                             ) : el.tableType === 'data' && previewRowIndex >= 0 ? (
                               <tr>
                                 {(el.columns || []).map(c => <td key={c} style={{ border: `${el.borderWidth}px ${el.borderStyle} var(--ink)`, padding: `${el.cellHeight}px` }}>{dataset.rows[previewRowIndex]?.[c]}</td>)}
                               </tr>
                             ) : (
                               // Fill blank rows
                               Array.from({ length: el.tableType === 'data' ? 3 : (el.emptyRows || 5) }).map((_, rIdx) => (
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
          </div>
        </div>

      </div>
    </div>
  );
};

export default TemplatePage;
