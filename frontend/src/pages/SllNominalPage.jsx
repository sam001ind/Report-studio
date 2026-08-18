
// Bulletproof Universal Excel, CSV, HTML-table & ZIP parser
const parseUploadedSpreadsheet = async (file) => {
  let buffers = [];

  if (file.name.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file);
    const validFiles = Object.keys(zip.files).filter(name => 
      !name.startsWith("__MACOSX/") && 
      !name.startsWith(".") && 
      !zip.files[name].dir &&
      (name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".csv"))
    );

    if (validFiles.length === 0) {
      throw new Error("No .xlsx, .xls, or .csv spreadsheets found inside the uploaded ZIP archive.");
    }

    for (const name of validFiles) {
      const buf = await zip.files[name].async("arraybuffer");
      buffers.push({ name, buffer: buf });
    }
  } else {
    const buf = await file.arrayBuffer();
    buffers.push({ name: file.name, buffer: buf });
  }

  let allRows = [];
  for (const item of buffers) {
    let wb;
    try {
      // Attempt 1: Binary array buffer (True .xlsx / .xls)
      wb = XLSX.read(new Uint8Array(item.buffer), { type: "array", cellDates: true });
    } catch (err1) {
      try {
        // Attempt 2: UTF-8 string parse (HTML table or CSV saved as .xls/.xlsx)
        const text = new TextDecoder("utf-8").decode(item.buffer);
        wb = XLSX.read(text, { type: "string", raw: true });
      } catch (err2) {
        // Attempt 3: Windows-1252 / ISO-8859 parse
        const text = new TextDecoder("windows-1252").decode(item.buffer);
        wb = XLSX.read(text, { type: "string", raw: true });
      }
    }

    if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
      const firstSheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
      if (json && json.length > 0) {
        allRows = allRows.concat(json);
      }
    }
  }

  if (allRows.length === 0) {
    throw new Error("No data rows found in the uploaded file(s).");
  }

  return allRows;
};

import { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import JSZip from 'jszip';
import { Settings, Download, Eye, FileText, Plus, Trash2, AlignCenter, AlignLeft, AlignRight, Bold } from 'lucide-react';

const SllNominalPage = () => {
  
  // File and sheet states
  const [selectedFile, setSelectedFile] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  
  // Column mapping states (0-indexed indices mapping to excel columns)
  const [headers, setHeaders] = useState([]);
  const [columnMapping, setColumnMapping] = useState({
    programme: 0,
    name: 2,
    seatNo: 3,
    venueCode: 4,
    venueName: 5,
    courseCode: 6,
    courseTitle: 7
  });

  // Dynamic Header Rich Lines (Structured Rich Text Editor for PDF)
  const [headerLines, setHeaderLines] = useState([
    { text: 'Kannur University', fontSize: 16, isBold: true, align: 'center', yOffset: 40 },
    { text: 'Examination Branch', fontSize: 12, isBold: true, align: 'center', yOffset: 60 },
    { text: 'I Semester Private Registration 2025 -2028 Admission - November 2025', fontSize: 10, isBold: true, align: 'center', yOffset: 78 }
  ]);

  // Dynamic Table Columns styling editor (Rich Text styling for table grid)
  const [tableColumns, setTableColumns] = useState([
    { id: 'slNo', label: 'Sl No', width: 35, align: 'center', fontSize: 9 },
    { id: 'seatNo', label: 'Seat No', width: 70, align: 'center', fontSize: 9 },
    { id: 'name', label: 'Name', width: 140, align: 'left', fontSize: 9 },
    { id: 'courses', label: 'Courses', width: 210, align: 'left', fontSize: 9.5 },
    { id: 'remark', label: 'Remark', width: 60, align: 'left', fontSize: 9 }
  ]);

  // Processing states
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Generated Venue Data State
  const [comboGroups, setComboGroups] = useState({});
  const [activeComboTab, setActiveComboTab] = useState('');

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  // Extract headers on sheet change
  useEffect(() => {
    if (workbook && selectedSheet) {
      const sheet = workbook.Sheets[selectedSheet];
      if (sheet) {
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
        if (rows.length > 0) {
          // Store raw headers for mapping dropdowns
          const firstRow = rows[0].map((cell, idx) => cell ? String(cell).trim() : `Column ${idx + 1}`);
          setHeaders(firstRow);

          // Try to auto-detect columns
          setColumnMapping(prev => {
            const autoMap = { ...prev };
            firstRow.forEach((name, idx) => {
              const lower = name.toLowerCase();
              if (lower.includes('programme') || lower.includes('program')) autoMap.programme = idx;
              if (lower.includes('name')) autoMap.name = idx;
              if (lower.includes('seat') || lower.includes('reg') || lower.includes('register') || lower.includes('roll')) autoMap.seatNo = idx;
              if (lower.includes('venuecode') || lower.includes('centercode') || lower.includes('collegecode')) autoMap.venueCode = idx;
              if (lower.includes('venuename') || lower.includes('centername') || lower.includes('collegename') || (lower.includes('venue') && !lower.includes('code'))) autoMap.venueName = idx;
              if (lower.includes('coursecode') || lower.includes('subjectcode') || (lower.includes('course') && !lower.includes('title') && !lower.includes('name'))) autoMap.courseCode = idx;
              if (lower.includes('coursetitle') || lower.includes('subjectname') || lower.includes('coursename') || lower.includes('title')) autoMap.courseTitle = idx;
            });
            return autoMap;
          });
        }
      }
    } else {
      setHeaders([]);
    }
  }, [workbook, selectedSheet]);

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setSelectedFile(file);
    setStatus('Reading workbook...');
    setComboGroups({});
    setActiveComboTab('');

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target.result;
        const wb = XLSX.read(data, { type: 'array', cellDates: true });
        setWorkbook(wb);
        setSheetNames(wb.SheetNames);
        setSelectedSheet(wb.SheetNames[0] || '');
        setStatus('Workbook loaded successfully', 'success');
      } catch (err) {
        setStatus(`Error reading Excel: ${err.message}`, 'error');
        setWorkbook(null);
        setSheetNames([]);
        setSelectedSheet('');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Group data by [Programme Name] + [VenueCode - VenueName]
  const processNominalRoll = () => {
    if (!workbook || !selectedSheet) {
      setStatus('Please upload and select an Excel sheet first.', 'error');
      return;
    }

    const sheet = workbook.Sheets[selectedSheet];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
    if (rows.length < 2) {
      setStatus('Sheet does not contain sufficient data rows.', 'error');
      return;
    }

    setIsProcessing(true);
    setStatus('Grouping nominal rolls...');

    try {
      const dataRows = rows.slice(1);
      const groups = {};

      dataRows.forEach((row) => {
        // Skip blank rows
        const isBlank = row.every(cell => cell === null || cell === undefined || cell === '');
        if (isBlank) return;

        const programme = String(row[columnMapping.programme] || 'Unassigned Programme').trim();
        
        // Merge Venue Code and Venue Name into a single "Code - Name" string
        const vCode = String(row[columnMapping.venueCode] || '').trim();
        const vName = String(row[columnMapping.venueName] || '').trim();
        const venue = vCode && vName ? `${vCode} - ${vName}` : (vCode || vName || 'Unassigned Venue');

        const seatNo = String(row[columnMapping.seatNo] || '').trim();
        const name = String(row[columnMapping.name] || '').trim();
        const cCode = String(row[columnMapping.courseCode] || '').trim();
        const cTitle = String(row[columnMapping.courseTitle] || '').trim();

        if (!venue || !seatNo) return; // Skip if no venue or register number

        const comboKey = `${programme} - ${venue}`;

        if (!groups[comboKey]) {
          groups[comboKey] = {
            programme,
            venue,
            students: {}
          };
        }

        // Group by seatNo (register number) to perform row spanning
        const studentKey = `${seatNo}_${name}`;
        if (!groups[comboKey].students[studentKey]) {
          groups[comboKey].students[studentKey] = {
            seatNo,
            name,
            courses: []
          };
        }

        groups[comboKey].students[studentKey].courses.push({
          code: cCode,
          title: cTitle
        });
      });

      // Convert inner dictionary to sorted arrays of students
      const processedGroups = {};
      Object.keys(groups).forEach(comboKey => {
        const { programme, venue, students } = groups[comboKey];
        const studentList = Object.values(students);
        // Sort students by Seat No
        studentList.sort((a, b) => a.seatNo.localeCompare(b.seatNo, undefined, { numeric: true, sensitivity: 'base' }));
        processedGroups[comboKey] = {
          programme,
          venue,
          students: studentList
        };
      });

      setComboGroups(processedGroups);
      const comboKeys = Object.keys(processedGroups);
      if (comboKeys.length > 0) {
        setActiveComboTab(comboKeys[0]);
      }
      
      setStatus(`Nominal rolls processed! Found ${comboKeys.length} Programme-Venue combinations.`, 'success');
    } catch (err) {
      setStatus(`Error processing data: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Generate jsPDF instance for a single combo group
  const generateVenuePDF = (programmeName, venueName, students) => {
    const doc = new jsPDF({
      orientation: 'portrait',
      unit: 'pt',
      format: 'a4'
    });

    // 1. Dynamic Header Block
    headerLines.forEach((line) => {
      if (!line.text.trim()) return;
      doc.setFont('Helvetica', line.isBold ? 'bold' : 'normal');
      doc.setFontSize(line.fontSize);
      
      let xPos = 297; // center
      if (line.align === 'left') xPos = 40;
      if (line.align === 'right') xPos = 550; // a4 width is 595 pt
      
      doc.text(line.text, xPos, line.yOffset, { align: line.align });
    });

    // 2. Metadata Info
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(11);
    doc.text(`Programme: ${programmeName}`, 40, 110);
    doc.text(`Venue: ${venueName}`, 40, 128);

    // 3. Prepare Table Data (Applying Row-Spanning / Merging)
    const tableBody = [];
    let slNo = 1;

    students.forEach((student) => {
      const courseCount = student.courses.length;
      
      student.courses.forEach((course, index) => {
        const row = [];
        
        // Only output student profile details on the first course row (enable row spanning)
        if (index === 0) {
          row.push({ content: slNo.toString(), rowSpan: courseCount, styles: { valign: 'middle', halign: tableColumns[0].align } });
          row.push({ content: student.seatNo, rowSpan: courseCount, styles: { valign: 'middle', halign: tableColumns[1].align } });
          row.push({ content: student.name, rowSpan: courseCount, styles: { valign: 'middle', halign: tableColumns[2].align } });
        }
        
        // Course detail column (Course Code - Course Title)
        const courseStr = course.code ? `${course.code} - ${course.title}` : course.title;
        row.push({ content: courseStr, styles: { valign: 'top', halign: tableColumns[3].align } });
        
        // Remark column
        if (index === 0) {
          row.push({ content: '', rowSpan: courseCount, styles: { valign: 'middle', halign: tableColumns[4].align } });
        }

        tableBody.push(row);
      });
      slNo += 1;
    });

    // Extract dynamic headers configured by the user
    const headLabels = tableColumns.map(col => col.label);

    // Map column styles dynamically
    const colStyles = {};
    tableColumns.forEach((col, idx) => {
      colStyles[idx] = { 
        cellWidth: col.width,
        halign: col.align,
        fontSize: col.fontSize
      };
    });

    // 4. Generate Autotable
    autoTable(doc, {
      startY: 145,
      head: [headLabels],
      body: tableBody,
      theme: 'grid',
      margin: { left: 40, right: 40, bottom: 40 },
      styles: {
        fontSize: 9,
        cellPadding: 6,
        lineColor: [180, 180, 180],
        lineWidth: 0.5,
        textColor: [30, 30, 30]
      },
      headStyles: {
        fillColor: [240, 240, 240],
        textColor: [0, 0, 0],
        fontStyle: 'bold',
        halign: 'center',
        lineWidth: 1
      },
      columnStyles: colStyles
    });

    return doc;
  };

  // Download single combination PDF
  const downloadSinglePDF = (comboKey) => {
    const group = comboGroups[comboKey];
    if (!group) return;
    const doc = generateVenuePDF(group.programme, group.venue, group.students);
    const safeProg = String(group.programme).replace(/[^a-zA-Z0-9_-]/g, "_").replace(/_+/g, "_");
    const safeVenue = String(group.venue).replace(/[^a-zA-Z0-9_-]/g, "_").replace(/_+/g, "_");
    doc.save(`${safeProg}_${safeVenue}.pdf`);
  };

  // Batch download all PDFs in a single ZIP
  const downloadAllAsZip = async () => {
    const comboKeys = Object.keys(comboGroups);
    if (comboKeys.length === 0) return;

    setStatus('Creating ZIP archive...', 'normal');
    const zip = new JSZip();

    try {
      comboKeys.forEach((comboKey) => {
        const group = comboGroups[comboKey];
        const doc = generateVenuePDF(group.programme, group.venue, group.students);
        
        // Output PDF to Blob
        const pdfBlob = doc.output('blob');
        const safeProg = String(group.programme).replace(/[^a-zA-Z0-9_-]/g, "_").replace(/_+/g, "_");
        const safeVenue = String(group.venue).replace(/[^a-zA-Z0-9_-]/g, "_").replace(/_+/g, "_");
        zip.file(`${safeProg}_${safeVenue}.pdf`, pdfBlob);
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const downloadUrl = URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `Nominal_Rolls_by_Programme_Venue.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

      setStatus(`ZIP archive downloaded successfully!`, 'success');
    } catch (err) {
      setStatus(`Failed to generate ZIP: ${err.message}`, 'error');
    }
  };

  // Header Rich Text Editor Helper Actions
  const updateHeaderLine = (index, key, value) => {
    const updated = [...headerLines];
    updated[index][key] = value;
    setHeaderLines(updated);
  };

  const addHeaderLine = () => {
    const lastLine = headerLines[headerLines.length - 1];
    const newY = lastLine ? lastLine.yOffset + 18 : 40;
    setHeaderLines([...headerLines, { text: '', fontSize: 10, isBold: false, align: 'center', yOffset: newY }]);
  };

  const removeHeaderLine = (index) => {
    const updated = headerLines.filter((_, idx) => idx !== index);
    setHeaderLines(updated);
  };

  // Table Column Editor Actions
  const updateTableCol = (index, key, value) => {
    const updated = [...tableColumns];
    updated[index][key] = value;
    setTableColumns(updated);
  };

  return (
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto', overflowY: 'auto', width: '100%', minHeight: '100vh', background: 'var(--bg)', color: 'var(--ink)' }}>
      {/* Navigation Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
        <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
          ← Back to Portal
        </Link>
        
        {/* Status Pill */}
        <div style={{
          padding: '6px 14px',
          borderRadius: '20px',
          fontSize: '13px',
          fontWeight: 600,
          background: statusType === 'error' ? 'var(--danger)' : statusType === 'success' ? 'var(--accent-soft)' : 'var(--panel)',
          color: statusType === 'error' ? 'white' : statusType === 'success' ? 'var(--accent)' : 'var(--muted)',
          border: '1px solid var(--line)'
        }}>
          {statusMsg}
        </div>
      </div>

      <h2>Venue-Wise Nominal Roll</h2>
      <p className="subtitle">Compile nominal roll lists grouped by unique combinations of Programme + Venue. Seat numbers with multiple subjects are automatically merged.</p>
<p className="subtitle" style={{ marginTop: '8px', color: 'var(--muted)' }}>Upload one or more <strong>Event wise Pre exam data</strong> files (Excel <code>.xlsx</code>, <code>.xls</code>, <code>.xlsm</code>, <code>.csv</code>) or a <code>.zip</code> containing multiple Excel files. If a ZIP is uploaded, all Excel files will be merged before processing.</p>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(450px, 1fr))', gap: '32px', marginBottom: '32px' }}>
        {/* Settings Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}><Settings size={20} /> Settings</h3>
          
          {/* File Picker */}
          <div style={{ border: '1px dashed var(--line)', padding: '24px', borderRadius: '8px', background: 'var(--bg)', textAlign: 'center', position: 'relative' }}>
            <strong style={{ display: 'block', marginBottom: '8px' }}>Select Nominal Roll Excel</strong>
            {selectedFile ? (
              <div style={{ color: 'var(--accent)', fontWeight: 600 }}>📄 {selectedFile.name}</div>
            ) : (
              <span style={{ color: 'var(--muted)', fontSize: '13px' }}>Drag here or click to browse</span>
            )}
            <input 
              type="file" 
              accept=".xlsx, .xls, .xlsm, .csv, .zip" 
              onChange={handleFileChange} 
              style={{
                position: 'absolute',
                top: 0,
                left: 0,
                right: 0,
                bottom: 0,
                opacity: 0,
                cursor: 'pointer',
                width: '100%'
              }} 
            />
          </div>

          {/* Configuration Form */}
          <div className="form-group">
            <label>Select Sheet</label>
            <select 
              value={selectedSheet} 
              onChange={(e) => setSelectedSheet(e.target.value)} 
              disabled={sheetNames.length === 0}
            >
              {sheetNames.map((name) => (
                <option key={name} value={name}>{name}</option>
              ))}
            </select>
          </div>

          {/* Column Mappings */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)' }}>
            <strong style={{ fontSize: '13px', display: 'block', marginBottom: '12px' }}>Excel Column Mappings</strong>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', fontSize: '12px' }}>
              <div className="form-group">
                <label>Programme Column (A)</label>
                <select 
                  value={columnMapping.programme} 
                  onChange={(e) => setColumnMapping({ ...columnMapping, programme: parseInt(e.target.value) })}
                  disabled={headers.length === 0}
                >
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Venue Code Column</label>
                <select 
                  value={columnMapping.venueCode} 
                  onChange={(e) => setColumnMapping({ ...columnMapping, venueCode: parseInt(e.target.value) })}
                  disabled={headers.length === 0}
                >
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Venue Name Column</label>
                <select 
                  value={columnMapping.venueName} 
                  onChange={(e) => setColumnMapping({ ...columnMapping, venueName: parseInt(e.target.value) })}
                  disabled={headers.length === 0}
                >
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Seat No / Reg No (D)</label>
                <select 
                  value={columnMapping.seatNo} 
                  onChange={(e) => setColumnMapping({ ...columnMapping, seatNo: parseInt(e.target.value) })}
                  disabled={headers.length === 0}
                >
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Student Name (C)</label>
                <select 
                  value={columnMapping.name} 
                  onChange={(e) => setColumnMapping({ ...columnMapping, name: parseInt(e.target.value) })}
                  disabled={headers.length === 0}
                >
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Course Code (G)</label>
                <select 
                  value={columnMapping.courseCode} 
                  onChange={(e) => setColumnMapping({ ...columnMapping, courseCode: parseInt(e.target.value) })}
                  disabled={headers.length === 0}
                >
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Course Title (H)</label>
                <select 
                  value={columnMapping.courseTitle} 
                  onChange={(e) => setColumnMapping({ ...columnMapping, courseTitle: parseInt(e.target.value) })}
                  disabled={headers.length === 0}
                >
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
            </div>
          </div>

          {/* Structured PDF Header Rich Text Style Manager */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)', display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <strong style={{ fontSize: '13px' }}>PDF Header Layout Editor</strong>
              <button 
                onClick={addHeaderLine} 
                className="secondary" 
                style={{ padding: '4px 8px', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px' }}
              >
                <Plus size={12} /> Add Line
              </button>
            </div>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              {headerLines.map((line, idx) => (
                <div key={idx} style={{ borderBottom: '1px solid var(--line)', paddingBottom: '12px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                    <span style={{ fontSize: '11px', color: 'var(--muted)', width: '45px' }}>Line {idx+1}</span>
                    <input 
                      type="text" 
                      value={line.text} 
                      onChange={(e) => updateHeaderLine(idx, 'text', e.target.value)} 
                      placeholder={`Header Line Text ${idx+1}`}
                      style={{ flex: 1, padding: '6px' }}
                    />
                    <button 
                      onClick={() => removeHeaderLine(idx)} 
                      style={{ padding: '6px', background: 'transparent', border: 'none', color: 'var(--danger)', cursor: 'pointer' }}
                      disabled={headerLines.length <= 1}
                    >
                      <Trash2 size={14} />
                    </button>
                  </div>

                  {/* Inline formatting controls */}
                  <div style={{ display: 'flex', gap: '16px', alignItems: 'center', paddingLeft: '53px' }}>
                    {/* Size selector */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                      <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Size:</span>
                      <input 
                        type="number" 
                        value={line.fontSize} 
                        onChange={(e) => updateHeaderLine(idx, 'fontSize', parseInt(e.target.value) || 10)} 
                        style={{ width: '45px', padding: '3px 4px', fontSize: '11px' }}
                        min="6"
                        max="32"
                      />
                      <span style={{ fontSize: '11px', color: 'var(--muted)' }}>pt</span>
                    </div>

                    {/* Bold Toggle */}
                    <button
                      onClick={() => updateHeaderLine(idx, 'isBold', !line.isBold)}
                      style={{
                        padding: '4px 8px',
                        background: line.isBold ? 'var(--accent)' : 'var(--panel)',
                        color: line.isBold ? 'white' : 'var(--ink)',
                        border: '1px solid var(--line)',
                        borderRadius: '4px',
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '2px',
                        fontSize: '11px'
                      }}
                    >
                      <Bold size={11} /> {line.isBold ? 'Bold' : 'Normal'}
                    </button>

                    {/* Alignment */}
                    <div style={{ display: 'flex', gap: '2px', border: '1px solid var(--line)', borderRadius: '4px', overflow: 'hidden' }}>
                      <button
                        onClick={() => updateHeaderLine(idx, 'align', 'left')}
                        style={{ padding: '4px 8px', background: line.align === 'left' ? 'var(--accent)' : 'var(--panel)', color: line.align === 'left' ? 'white' : 'var(--ink)', border: 'none', cursor: 'pointer' }}
                      >
                        <AlignLeft size={11} />
                      </button>
                      <button
                        onClick={() => updateHeaderLine(idx, 'align', 'center')}
                        style={{ padding: '4px 8px', background: line.align === 'center' ? 'var(--accent)' : 'var(--panel)', color: line.align === 'center' ? 'white' : 'var(--ink)', border: 'none', cursor: 'pointer' }}
                      >
                        <AlignCenter size={11} />
                      </button>
                      <button
                        onClick={() => updateHeaderLine(idx, 'align', 'right')}
                        style={{ padding: '4px 8px', background: line.align === 'right' ? 'var(--accent)' : 'var(--panel)', color: line.align === 'right' ? 'white' : 'white', border: 'none', cursor: 'pointer' }}
                      >
                        <AlignRight size={11} />
                      </button>
                    </div>

                    {/* Vertical Position Offset slider */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '6px', flex: 1 }}>
                      <span style={{ fontSize: '11px', color: 'var(--muted)', whiteSpace: 'nowrap' }}>Pos-Y:</span>
                      <input 
                        type="range" 
                        value={line.yOffset} 
                        onChange={(e) => updateHeaderLine(idx, 'yOffset', parseInt(e.target.value))} 
                        min="20"
                        max="140"
                        style={{ flex: 1, height: '4px', cursor: 'pointer' }}
                      />
                      <span style={{ fontSize: '11px', color: 'var(--muted)', width: '25px', textAlign: 'right' }}>{line.yOffset}</span>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Table Grid Columns Rich Text Editor */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)', display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <strong style={{ fontSize: '13px' }}>Table Column & Layout Editor</strong>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              {tableColumns.map((col, idx) => (
                <div key={col.id} style={{ display: 'flex', flexDirection: 'column', gap: '8px', borderBottom: '1px solid var(--line)', paddingBottom: '12px' }}>
                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                    <span style={{ fontSize: '11px', color: 'var(--muted)', width: '60px' }}>Col {idx+1}:</span>
                    <input 
                      type="text" 
                      value={col.label} 
                      onChange={(e) => updateTableCol(idx, 'label', e.target.value)} 
                      placeholder="Header Label"
                      style={{ flex: 2, padding: '6px', fontSize: '11px' }}
                    />
                    
                    {/* Width setting */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '2px' }}>
                      <span style={{ fontSize: '10px', color: 'var(--muted)' }}>Width:</span>
                      <input 
                        type="number" 
                        value={col.width} 
                        onChange={(e) => updateTableCol(idx, 'width', parseInt(e.target.value) || 20)} 
                        style={{ width: '45px', padding: '4px', fontSize: '11px' }}
                        min="10"
                        max="300"
                      />
                    </div>
                  </div>

                  <div style={{ display: 'flex', gap: '16px', alignItems: 'center', paddingLeft: '68px' }}>
                    {/* Font size */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                      <span style={{ fontSize: '10px', color: 'var(--muted)' }}>Font Size:</span>
                      <input 
                        type="number" 
                        value={col.fontSize} 
                        onChange={(e) => updateTableCol(idx, 'fontSize', parseFloat(e.target.value) || 8)} 
                        style={{ width: '45px', padding: '3px 4px', fontSize: '11px' }}
                        step="0.5"
                        min="6"
                        max="16"
                      />
                      <span style={{ fontSize: '10px', color: 'var(--muted)' }}>pt</span>
                    </div>

                    {/* Alignment */}
                    <div style={{ display: 'flex', gap: '2px', border: '1px solid var(--line)', borderRadius: '4px', overflow: 'hidden' }}>
                      <button
                        onClick={() => updateTableCol(idx, 'align', 'left')}
                        style={{ padding: '3px 6px', background: col.align === 'left' ? 'var(--accent)' : 'var(--panel)', color: col.align === 'left' ? 'white' : 'var(--ink)', border: 'none', cursor: 'pointer' }}
                      >
                        <AlignLeft size={10} />
                      </button>
                      <button
                        onClick={() => updateTableCol(idx, 'align', 'center')}
                        style={{ padding: '3px 6px', background: col.align === 'center' ? 'var(--accent)' : 'var(--panel)', color: col.align === 'center' ? 'white' : 'var(--ink)', border: 'none', cursor: 'pointer' }}
                      >
                        <AlignCenter size={10} />
                      </button>
                      <button
                        onClick={() => updateTableCol(idx, 'align', 'right')}
                        style={{ padding: '3px 6px', background: col.align === 'right' ? 'var(--accent)' : 'var(--panel)', color: col.align === 'right' ? 'white' : 'white', border: 'none', cursor: 'pointer' }}
                      >
                        <AlignRight size={10} />
                      </button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>

          <button 
            onClick={processNominalRoll} 
            disabled={isProcessing || !workbook} 
            style={{ width: '100%', padding: '14px', fontSize: '15px' }}
          >
            {isProcessing ? "Processing Data..." : "Run Split & Layout Process"}
          </button>
        </div>

        {/* Live Preview Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}><Eye size={20} /> Live Preview</h3>
            {Object.keys(comboGroups).length > 0 && (
              <button 
                onClick={downloadAllAsZip} 
                className="secondary" 
                style={{ padding: '8px 16px', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px', borderColor: 'var(--accent)', color: 'var(--accent)' }}
              >
                <Download size={16} /> Download All (ZIP)
              </button>
            )}
          </div>

          {Object.keys(comboGroups).length > 0 ? (
            <>
              {/* Tab Selector */}
              <div style={{ display: 'flex', gap: '8px', overflowX: 'auto', paddingBottom: '8px', borderBottom: '1px solid var(--line)' }}>
                {Object.keys(comboGroups).map((comboKey) => (
                  <button
                    key={comboKey}
                    onClick={() => setActiveComboTab(comboKey)}
                    style={{
                      padding: '6px 12px',
                      borderRadius: '6px',
                      border: '1px solid var(--line)',
                      background: activeComboTab === comboKey ? 'var(--accent)' : 'var(--panel)',
                      color: activeComboTab === comboKey ? 'white' : 'var(--ink)',
                      whiteSpace: 'nowrap',
                      fontSize: '12px',
                      cursor: 'pointer',
                      fontWeight: activeComboTab === comboKey ? 600 : 400
                    }}
                  >
                    {comboKey} ({comboGroups[comboKey].students.length} Studs)
                  </button>
                ))}
              </div>

              {/* Active Venue Layout Preview */}
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <h4 style={{ margin: 0, color: 'var(--accent)' }}>{comboGroups[activeComboTab].programme}</h4>
                    <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Venue: {comboGroups[activeComboTab].venue}</span>
                  </div>
                  <button 
                    onClick={() => downloadSinglePDF(activeComboTab)} 
                    style={{ padding: '6px 12px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}
                  >
                    <FileText size={14} /> Download PDF
                  </button>
                </div>

                {/* Simulated Sheet Table Grid */}
                <div style={{ flex: 1, overflowY: 'auto', maxHeight: '400px', border: '1px solid var(--line)', borderRadius: '6px' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px', textAlign: 'left', background: 'var(--bg)' }}>
                    <thead>
                      <tr style={{ background: 'var(--panel)', borderBottom: '2px solid var(--line)', color: 'var(--ink)' }}>
                        {tableColumns.map((col, idx) => (
                          <th 
                            key={col.id}
                            style={{ 
                              padding: '10px 8px', 
                              borderRight: idx < tableColumns.length - 1 ? '1px solid var(--line)' : 'none', 
                              textAlign: col.align, 
                              width: `${col.width}px` 
                            }}
                          >
                            {col.label}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {comboGroups[activeComboTab].students.map((student, sIdx) => {
                        return student.courses.map((course, cIdx) => {
                          const isFirst = cIdx === 0;
                          return (
                            <tr key={`${sIdx}-${cIdx}`} style={{ borderBottom: '1px solid var(--line)' }}>
                              {isFirst ? (
                                <>
                                  <td 
                                    rowSpan={student.courses.length} 
                                    style={{ 
                                      padding: '8px', 
                                      borderRight: '1px solid var(--line)', 
                                      textAlign: tableColumns[0].align, 
                                      verticalAlign: 'middle', 
                                      background: 'rgba(255,255,255,0.01)',
                                      fontSize: `${tableColumns[0].fontSize}px`
                                    }}
                                  >
                                    {sIdx + 1}
                                  </td>
                                  <td 
                                    rowSpan={student.courses.length} 
                                    style={{ 
                                      padding: '8px', 
                                      borderRight: '1px solid var(--line)', 
                                      textAlign: tableColumns[1].align, 
                                      verticalAlign: 'middle', 
                                      fontWeight: 600,
                                      fontSize: `${tableColumns[1].fontSize}px`
                                    }}
                                  >
                                    {student.seatNo}
                                  </td>
                                  <td 
                                    rowSpan={student.courses.length} 
                                    style={{ 
                                      padding: '8px', 
                                      borderRight: '1px solid var(--line)', 
                                      verticalAlign: 'middle',
                                      textAlign: tableColumns[2].align,
                                      fontSize: `${tableColumns[2].fontSize}px`
                                    }}
                                  >
                                    {student.name}
                                  </td>
                                </>
                              ) : null}
                              <td 
                                style={{ 
                                  padding: '8px', 
                                  borderRight: '1px solid var(--line)', 
                                  color: 'var(--ink)',
                                  textAlign: tableColumns[3].align,
                                  fontSize: `${tableColumns[3].fontSize}px`
                                }}
                              >
                                <span style={{ fontWeight: 600, color: 'var(--accent)' }}>{course.code}</span>
                                {course.code && course.title ? ' - ' : ''}
                                <span>{course.title}</span>
                              </td>
                              {isFirst ? (
                                <td 
                                  rowSpan={student.courses.length} 
                                  style={{ 
                                    padding: '8px', 
                                    verticalAlign: 'middle',
                                    textAlign: tableColumns[4].align,
                                    fontSize: `${tableColumns[4].fontSize}px`
                                  }}
                                >
                                  {/* Remark Column */}
                                </td>
                              ) : null}
                            </tr>
                          );
                        });
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          ) : (
            <div style={{ display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', border: '1px dashed var(--line)', borderRadius: '8px', padding: '60px', color: 'var(--muted)', fontSize: '14px', textAlign: 'center' }}>
              No split run processed yet. Upload your nominal roll Excel sheet, review your column mappings, and click "Run Split & Layout Process".
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default SllNominalPage;
