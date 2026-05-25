import React, { useState, useRef, useEffect } from 'react';

const PAGE_SIZES = {
  A4: { width: 794, height: 1123 },
  A3: { width: 1123, height: 1588 },
  Letter: { width: 816, height: 1056 }
};

const TemplatePage = () => {
  const [template, setTemplate] = useState({
    pageSize: 'A4',
    orientation: 'portrait',
    zoomScale: 0.75,
    elements: [],
    bgImage: null
  });
  
  const [selectedId, setSelectedId] = useState(null);
  
  // Dragging state
  const draggingRef = useRef(null);
  
  const handleMouseDown = (e, id) => {
    e.stopPropagation();
    setSelectedId(id);
    const el = template.elements.find(el => el.id === id);
    if (!el) return;
    
    draggingRef.current = {
      id,
      startX: e.clientX,
      startY: e.clientY,
      initialX: el.x,
      initialY: el.y
    };
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  const handleMouseMove = (e) => {
    if (!draggingRef.current) return;
    
    const { id, startX, startY, initialX, initialY } = draggingRef.current;
    const dx = (e.clientX - startX) / template.zoomScale;
    const dy = (e.clientY - startY) / template.zoomScale;
    
    setTemplate(prev => ({
      ...prev,
      elements: prev.elements.map(el => 
        el.id === id ? { ...el, x: initialX + dx, y: initialY + dy } : el
      )
    }));
  };

  const handleMouseUp = () => {
    draggingRef.current = null;
    document.removeEventListener('mousemove', handleMouseMove);
    document.removeEventListener('mouseup', handleMouseUp);
  };

  const addTextElement = () => {
    setTemplate(prev => ({
      ...prev,
      elements: [
        ...prev.elements,
        { id: `el_${Date.now()}`, type: 'text', content: 'New Text', x: 50, y: 50, size: 16, fontFamily: "Inter, sans-serif" }
      ]
    }));
  };

  const selectedElement = template.elements.find(el => el.id === selectedId);

  let width = PAGE_SIZES[template.pageSize].width;
  let height = PAGE_SIZES[template.pageSize].height;
  if (template.orientation === 'landscape') {
    [width, height] = [height, width];
  }

  return (
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto' }}>
      <h2>Template Creator</h2>
      <p className="subtitle">Set page size, design layout, and format text.</p>

      <div style={{ display: 'flex', gap: '24px', alignItems: 'flex-start' }}>
        
        {/* Toolbox */}
        <div style={{ width: '340px', background: 'white', padding: '20px', borderRadius: '12px', border: '1px solid var(--line)', boxShadow: 'var(--shadow)', display: 'flex', flexDirection: 'column', gap: '16px' }}>
          <div className="form-row" style={{ marginBottom: 0 }}>
             <div className="form-group">
               <label>Page Size</label>
               <select value={template.pageSize} onChange={e => setTemplate({...template, pageSize: e.target.value})}>
                 <option value="A4">A4</option>
                 <option value="Letter">Letter</option>
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
          
          <button onClick={addTextElement} className="secondary">+ Add Text</button>

          <hr style={{ border: 0, borderTop: '1px solid var(--line)', margin: '8px 0' }} />

          {selectedElement ? (
            <div style={{ background: '#fafafa', padding: '12px', borderRadius: '8px', border: '1px solid var(--line)' }}>
              <h4 style={{ margin: '0 0 12px', color: 'var(--accent)' }}>Properties</h4>
              <div className="form-group">
                <label>Text Content</label>
                <input 
                  type="text" 
                  value={selectedElement.content} 
                  onChange={e => setTemplate(prev => ({
                    ...prev,
                    elements: prev.elements.map(el => el.id === selectedId ? { ...el, content: e.target.value } : el)
                  }))} 
                />
              </div>
            </div>
          ) : (
            <div style={{ fontSize: '13px', color: 'var(--muted)', textAlign: 'center', padding: '20px 0' }}>
              Select an element to edit properties.
            </div>
          )}
        </div>

        {/* Canvas Area */}
        <div style={{ flex: 1, background: '#e5e7eb', padding: '40px', borderRadius: '12px', minHeight: '70vh', overflow: 'auto' }}>
          <div style={{
             width: `${width}px`, 
             height: `${height}px`, 
             transform: `scale(${template.zoomScale})`, 
             transformOrigin: 'top left',
             background: 'white', 
             boxShadow: 'var(--shadow)', 
             border: '1px solid var(--line)', 
             position: 'relative',
             margin: '0 auto'
          }}>
             {template.elements.map(el => (
               <div
                 key={el.id}
                 onMouseDown={(e) => handleMouseDown(e, el.id)}
                 style={{
                   position: 'absolute',
                   left: el.x,
                   top: el.y,
                   fontSize: el.size,
                   fontFamily: el.fontFamily,
                   cursor: 'grab',
                   userSelect: 'none',
                   padding: '4px',
                   border: selectedId === el.id ? '1px solid var(--accent)' : '1px solid transparent',
                   backgroundColor: selectedId === el.id ? 'rgba(23, 107, 135, 0.05)' : 'transparent',
                 }}
               >
                 {el.content}
               </div>
             ))}
          </div>
        </div>
      </div>
    </div>
  );
};

export default TemplatePage;
