import React, { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import JSZip from 'jszip';
import { useAuth } from '../context/AuthContext';
import { Settings, Download, Eye, FileText, Plus, Trash2, AlignCenter, AlignLeft, AlignRight, Bold, CalendarRange } from 'lucide-react';

const QpStatementPage = () => {
  const { user } = useAuth();
  
  // File and sheet states
  const [selectedFile, setSelectedFile] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  
  // Column headers and mapping
  const [headers, setHeaders] = useState([]);
  const [columnMapping, setColumnMapping] = useState({
    eventName: 0,
    centerCode: 1,
    centerName: 2,
    venueCode: 3,
    venueName: 4,
    programName: 5,
    examDate: 6,
    startTime: 7,
    endTime: 8,
    courseCode: 9,
    courseName: 10,
    studentCount: 11
  });

  // Report styling/header configuration (Structured PDF Editor)
  const [headerLines, setHeaderLines] = useState([
    { text: 'Kannur University', fontSize: 16, isBold: true, align: 'center', yOffset: 40 },
    { text: 'Examination Branch', fontSize: 12, isBold: true, align: 'center', yOffset: 60 },
    { text: 'QUESTION PAPER REQUIREMENT STATEMENT', fontSize: 11, isBold: true, align: 'center', yOffset: 80 }
  ]);

  // Table customization parameters
  const [tableColumns, setTableColumns] = useState([
    { id: 'date', label: 'Exam Date & Session', width: 110, align: 'center', fontSize: 9 },
    { id: 'courseCode', label: 'Course Code', width: 75, align: 'center', fontSize: 9 },
    { id: 'courseName', label: 'Course Title', width: 230, align: 'left', fontSize: 9 },
    { id: 'studentCount', label: 'QP Count (Qty)', width: 60, align: 'center', fontSize: 9.5 },
    { id: 'remark', label: 'Remarks / Sign', width: 60, align: 'center', fontSize: 9 }
  ]);

  // Report generation mode: 'consolidated' (Date-wise summary) vs 'venue' (Venue-wise separate packing slips)
  const [reportMode, setReportMode] = useState('consolidated');

  // Processing results
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('normal');

  // Grouped datasets
  const [consolidatedData, setConsolidatedData] = useState([]);
  const [venueSlips, setVenueSlips] = useState({});
  const [activeVenueTab, setActiveVenueTab] = useState('');
  const [eventNameVal, setEventNameVal] = useState('');

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  // Header auto-detection
  useEffect(() => {
    if (workbook && selectedSheet) {
      const sheet = workbook.Sheets[selectedSheet];
      if (sheet) {
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
        if (rows.length > 0) {
          const firstRow = rows[0].map((cell, idx) => cell ? String(cell).trim() : `Column ${idx + 1}`);
          setHeaders(firstRow);

          // Get default Event Name from first data row
          if (rows.length > 1 && rows[1][0]) {
            setEventNameVal(String(rows[1][0]).trim());
          }

          // Auto map columns based on exact image headers
          const autoMap = { ...columnMapping };
          firstRow.forEach((name, idx) => {
            const lower = name.toLowerCase().replace(/[\s_-]/g, '');
            if (lower.includes('event')) autoMap.eventName = idx;
            if (lower.includes('centercode')) autoMap.centerCode = idx;
            if (lower.includes('centername')) autoMap.centerName = idx;
            if (lower.includes('venuecode')) autoMap.venueCode = idx;
            if (lower.includes('venuename') || (lower.includes('venue') && !lower.includes('code') && !lower.includes('id'))) autoMap.venueName = idx;
            if (lower.includes('program')) autoMap.programName = idx;
            if (lower.includes('date')) autoMap.examDate = idx;
            if (lower.includes('starttime') || lower.includes('start')) autoMap.startTime = idx;
            if (lower.includes('endtime') || lower.includes('end')) autoMap.endTime = idx;
            if (lower.includes('coursecode') || (lower.includes('subject') && lower.includes('code'))) autoMap.courseCode = idx;
            if (lower.includes('coursename') || lower.includes('coursetitle') || lower.includes('subjectname')) autoMap.courseName = idx;
            if (lower.includes('studentcount') || lower.includes('count') || lower.includes('qty') || lower.includes('strength')) autoMap.studentCount = idx;
          });
          setColumnMapping(autoMap);
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
    setConsolidatedData([]);
    setVenueSlips({});
    setActiveVenueTab('');

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

  // Compile QP Statement
  const processQpStatement = () => {
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
    setStatus('Grouping question paper statements...');

    try {
      const dataRows = rows.slice(1);
      const tempConsolidated = {};
      const tempVenueSlips = {};

      dataRows.forEach((row) => {
        const isBlank = row.every(cell => cell === null || cell === undefined || cell === '');
        if (isBlank) return;

        const event = String(row[columnMapping.eventName] || eventNameVal || '').trim();
        const vCode = String(row[columnMapping.venueCode] || '').trim();
        const vName = String(row[columnMapping.venueName] || '').trim();
        const venue = vCode && vName ? `${vCode} - ${vName}` : (vCode || vName || 'Unassigned Venue');
        
        const dateRaw = row[columnMapping.examDate];
        let dateStr = '';
        if (dateRaw instanceof Date) {
          dateStr = dateRaw.toLocaleDateString('en-GB'); // DD/MM/YYYY
        } else {
          dateStr = String(dateRaw || '').trim();
        }

        const start = String(row[columnMapping.startTime] || '').trim();
        const end = String(row[columnMapping.endTime] || '').trim();
        const session = start && end ? `${start} - ${end}` : (start || end || 'Session');

        const cCode = String(row[columnMapping.courseCode] || '').trim();
        const cName = String(row[columnMapping.courseName] || '').trim();
        const count = parseInt(row[columnMapping.studentCount]) || 0;

        if (!cCode || !dateStr) return; // Skip invalid records

        // A. Process Consolidated Summary (Date + Session + Course)
        const conKey = `${dateStr} | ${session} | ${cCode}`;
        if (!tempConsolidated[conKey]) {
          tempConsolidated[conKey] = {
            date: dateStr,
            session,
            courseCode: cCode,
            courseName: cName,
            totalQPs: 0,
            venues: {}
          };
        }
        tempConsolidated[conKey].totalQPs += count;
        if (venue) {
          tempConsolidated[conKey].venues[venue] = (tempConsolidated[conKey].venues[venue] || 0) + count;
        }

        // B. Process Venue-Wise Packing Slips (Venue -> Date + Session + Course)
        if (venue) {
          if (!tempVenueSlips[venue]) {
            tempVenueSlips[venue] = [];
          }
          tempVenueSlips[venue].push({
            date: dateStr,
            session,
            courseCode: cCode,
            courseName: cName,
            studentCount: count
          });
        }
      });

      // Format Consolidated array
      const formattedCon = Object.values(tempConsolidated);
      // Sort by Date, then Session, then Course Code
      formattedCon.sort((a, b) => a.date.localeCompare(b.date) || a.session.localeCompare(b.session) || a.courseCode.localeCompare(b.courseCode));
      setConsolidatedData(formattedCon);

      // Sort and compile Venue lists
      const sortedVenueSlips = {};
      Object.keys(tempVenueSlips).forEach((ven) => {
        const slips = tempVenueSlips[ven];
        slips.sort((a, b) => a.date.localeCompare(b.date) || a.session.localeCompare(b.session));
        sortedVenueSlips[ven] = slips;
      });
      setVenueSlips(sortedVenueSlips);

      const venues = Object.keys(sortedVenueSlips);
      if (venues.length > 0) {
        setActiveVenueTab(venues[0]);
      }

      setStatus(`QP Statement compiled! Found ${formattedCon.length} records across ${venues.length} venues.`, 'success');
    } catch (err) {
      setStatus(`Error compiling: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Generate jsPDF document for Consolidated summary
  const generateConsolidatedPDF = () => {
    const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    
    // Header lines
    headerLines.forEach((line) => {
      doc.setFont('Helvetica', line.isBold ? 'bold' : 'normal');
      doc.setFontSize(line.fontSize);
      let xPos = 297;
      if (line.align === 'left') xPos = 40;
      if (line.align === 'right') xPos = 550;
      doc.text(line.text, xPos, line.yOffset, { align: line.align });
    });

    // Metadata
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(11);
    doc.text(`Event Name: ${eventNameVal}`, 40, 110);
    doc.text(`Report Type: Consolidated QP Statement (All Venues)`, 40, 126);

    const body = consolidatedData.map((row) => [
      `${row.date} (${row.session})`,
      row.courseCode,
      row.courseName,
      row.totalQPs.toString(),
      '' // Remarks
    ]);

    const headLabels = tableColumns.map(c => c.label);
    const colStyles = {};
    tableColumns.forEach((col, idx) => {
      colStyles[idx] = { cellWidth: col.width, halign: col.align, fontSize: col.fontSize };
    });

    autoTable(doc, {
      startY: 145,
      head: [headLabels],
      body,
      theme: 'grid',
      margin: { left: 40, right: 40, bottom: 40 },
      styles: { fontSize: 9, cellPadding: 6, lineColor: [180, 180, 180], lineWidth: 0.5, textColor: [30, 30, 30] },
      headStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center', lineWidth: 1 },
      columnStyles: colStyles
    });

    return doc;
  };

  // Generate jsPDF for a Single Venue Slip
  const generateVenuePDF = (venueName, slips) => {
    const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });

    headerLines.forEach((line) => {
      doc.setFont('Helvetica', line.isBold ? 'bold' : 'normal');
      doc.setFontSize(line.fontSize);
      let xPos = 297;
      if (line.align === 'left') xPos = 40;
      if (line.align === 'right') xPos = 550;
      doc.text(line.text, xPos, line.yOffset, { align: line.align });
    });

    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(11);
    doc.text(`Event Name: ${eventNameVal}`, 40, 110);
    doc.text(`Dispatch Venue: ${venueName}`, 40, 126);

    const body = slips.map((row) => [
      `${row.date} (${row.session})`,
      row.courseCode,
      row.courseName,
      row.studentCount.toString(),
      '' // Remarks
    ]);

    const headLabels = tableColumns.map(c => c.label);
    const colStyles = {};
    tableColumns.forEach((col, idx) => {
      colStyles[idx] = { cellWidth: col.width, halign: col.align, fontSize: col.fontSize };
    });

    autoTable(doc, {
      startY: 145,
      head: [headLabels],
      body,
      theme: 'grid',
      margin: { left: 40, right: 40, bottom: 40 },
      styles: { fontSize: 9, cellPadding: 6, lineColor: [180, 180, 180], lineWidth: 0.5, textColor: [30, 30, 30] },
      headStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center', lineWidth: 1 },
      columnStyles: colStyles
    });

    return doc;
  };

  // Downloads
  const downloadConsolidatedPDF = () => {
    const doc = generateConsolidatedPDF();
    doc.save(`${eventNameVal || 'Event'}_Consolidated_QP_Statement.pdf`);
  };

  const downloadVenuePDF = (venueName) => {
    const slips = venueSlips[venueName];
    if (!slips) return;
    const doc = generateVenuePDF(venueName, slips);
    const safeName = venueName.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
    doc.save(`${eventNameVal || 'Event'}_${safeName}_QP_Slip.pdf`);
  };

  const downloadAllVenueSlipsAsZip = async () => {
    const venues = Object.keys(venueSlips);
    if (venues.length === 0) return;

    setStatus('Creating ZIP archive...', 'normal');
    const zip = new JSZip();

    try {
      venues.forEach((venueName) => {
        const slips = venueSlips[venueName];
        const doc = generateVenuePDF(venueName, slips);
        const pdfBlob = doc.output('blob');
        const safeName = venueName.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
        zip.file(`${eventNameVal || 'Event'}_${safeName}_QP_Slip.pdf`, pdfBlob);
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const downloadUrl = URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `${eventNameVal || 'Event'}_Venue_QP_Slips.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

      setStatus(`ZIP archive downloaded successfully!`, 'success');
    } catch (err) {
      setStatus(`Failed to generate ZIP: ${err.message}`, 'error');
    }
  };

  // Editor Actions
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

      <h2>QP Statement Report</h2>
      <p className="subtitle">Group and summarize question paper strengths by exam date, course, or venue code from master exam registers.</p>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(450px, 1fr))', gap: '32px', marginBottom: '32px' }}>
        {/* Settings Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}><Settings size={20} /> Settings</h3>
          
          {/* File Picker */}
          <div style={{ border: '1px dashed var(--line)', padding: '24px', borderRadius: '8px', background: 'var(--bg)', textAlign: 'center', position: 'relative' }}>
            <strong style={{ display: 'block', marginBottom: '8px' }}>Select QP Details Excel</strong>
            {selectedFile ? (
              <div style={{ color: 'var(--accent)', fontWeight: 600 }}>📄 {selectedFile.name}</div>
            ) : (
              <span style={{ color: 'var(--muted)', fontSize: '13px' }}>Drag here or click to browse</span>
            )}
            <input 
              type="file" 
              accept=".xlsx, .xls, .xlsm, .csv" 
              onChange={handleFileChange} 
              style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, opacity: 0, cursor: 'pointer', width: '100%' }} 
            />
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
            <div className="form-group">
              <label>Select Sheet</label>
              <select value={selectedSheet} onChange={(e) => setSelectedSheet(e.target.value)} disabled={sheetNames.length === 0}>
                {sheetNames.map((name) => <option key={name} value={name}>{name}</option>)}
              </select>
            </div>
            <div className="form-group">
              <label>Event Name</label>
              <input type="text" value={eventNameVal} onChange={(e) => setEventNameVal(e.target.value)} placeholder="e.g. I Semester Nov 2025" />
            </div>
          </div>

          {/* Column Mappings */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)' }}>
            <strong style={{ fontSize: '13px', display: 'block', marginBottom: '12px' }}>Excel Column Mappings</strong>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', fontSize: '12px' }}>
              <div className="form-group">
                <label>Event Name</label>
                <select value={columnMapping.eventName} onChange={(e) => setColumnMapping({ ...columnMapping, eventName: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Venue Code</label>
                <select value={columnMapping.venueCode} onChange={(e) => setColumnMapping({ ...columnMapping, venueCode: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Venue Name</label>
                <select value={columnMapping.venueName} onChange={(e) => setColumnMapping({ ...columnMapping, venueName: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Exam Date</label>
                <select value={columnMapping.examDate} onChange={(e) => setColumnMapping({ ...columnMapping, examDate: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Start Time (FN/AN)</label>
                <select value={columnMapping.startTime} onChange={(e) => setColumnMapping({ ...columnMapping, startTime: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>End Time</label>
                <select value={columnMapping.endTime} onChange={(e) => setColumnMapping({ ...columnMapping, endTime: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Course Code</label>
                <select value={columnMapping.courseCode} onChange={(e) => setColumnMapping({ ...columnMapping, courseCode: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Course Name</label>
                <select value={columnMapping.courseName} onChange={(e) => setColumnMapping({ ...columnMapping, courseName: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Student Count</label>
                <select value={columnMapping.studentCount} onChange={(e) => setColumnMapping({ ...columnMapping, studentCount: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
            </div>
          </div>

          {/* Header Layout Editor */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)', display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <strong style={{ fontSize: '13px' }}>PDF Header Layout Editor</strong>
              <button onClick={addHeaderLine} className="secondary" style={{ padding: '4px 8px', fontSize: '11px', display: 'flex', alignItems: 'center', gap: '4px' }}><Plus size={12} /> Add Line</button>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              {headerLines.map((line, idx) => (
                <div key={idx} style={{ borderBottom: '1px solid var(--line)', paddingBottom: '12px', display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                    <span style={{ fontSize: '11px', color: 'var(--muted)', width: '45px' }}>Line {idx+1}</span>
                    <input type="text" value={line.text} onChange={(e) => updateHeaderLine(idx, 'text', e.target.value)} placeholder={`Header Line Text ${idx+1}`} style={{ flex: 1, padding: '6px' }} />
                    <button onClick={() => removeHeaderLine(idx)} style={{ padding: '6px', background: 'transparent', border: 'none', color: 'var(--danger)', cursor: 'pointer' }} disabled={headerLines.length <= 1}><Trash2 size={14} /></button>
                  </div>
                  <div style={{ display: 'flex', gap: '16px', alignItems: 'center', paddingLeft: '53px' }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                      <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Size:</span>
                      <input type="number" value={line.fontSize} onChange={(e) => updateHeaderLine(idx, 'fontSize', parseInt(e.target.value) || 10)} style={{ width: '45px', padding: '3px 4px', fontSize: '11px' }} min="6" max="32" />
                    </div>
                    <button onClick={() => updateHeaderLine(idx, 'isBold', !line.isBold)} style={{ padding: '4px 8px', background: line.isBold ? 'var(--accent)' : 'var(--panel)', color: line.isBold ? 'white' : 'var(--ink)', border: '1px solid var(--line)', borderRadius: '4px', cursor: 'pointer', fontSize: '11px' }}><Bold size={11} /> Bold</button>
                    <div style={{ display: 'flex', gap: '6px', alignItems: 'center', flex: 1 }}>
                      <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Pos-Y:</span>
                      <input type="range" value={line.yOffset} onChange={(e) => updateHeaderLine(idx, 'yOffset', parseInt(e.target.value))} min="20" max="140" style={{ flex: 1 }} />
                      <span style={{ fontSize: '11px', color: 'var(--muted)' }}>{line.yOffset}</span>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>

          <button onClick={processQpStatement} disabled={isProcessing || !workbook} style={{ width: '100%', padding: '14px', fontSize: '15px' }}>
            {isProcessing ? "Processing Data..." : "Generate QP Statements"}
          </button>
        </div>

        {/* Live Preview Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}><Eye size={20} /> Live Preview</h3>
            
            <div style={{ display: 'flex', gap: '8px' }}>
              <select value={reportMode} onChange={(e) => setReportMode(e.target.value)} style={{ padding: '6px 12px', fontSize: '13px' }}>
                <option value="consolidated">Consolidated Summary</option>
                <option value="venue">Venue-Wise Slips</option>
              </select>

              {reportMode === 'consolidated' && consolidatedData.length > 0 && (
                <button onClick={downloadConsolidatedPDF} className="secondary" style={{ padding: '8px 16px', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <Download size={16} /> Download PDF
                </button>
              )}

              {reportMode === 'venue' && Object.keys(venueSlips).length > 0 && (
                <button onClick={downloadAllVenueSlipsAsZip} className="secondary" style={{ padding: '8px 16px', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <Download size={16} /> Download All (ZIP)
                </button>
              )}
            </div>
          </div>

          {reportMode === 'consolidated' ? (
            consolidatedData.length > 0 ? (
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div>
                  <h4 style={{ margin: 0, color: 'var(--accent)' }}>Consolidated Question Paper Statement</h4>
                  <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Shows total required print copies grouped by exam date & sessions</span>
                </div>

                <div style={{ flex: 1, overflowY: 'auto', maxHeight: '550px', border: '1px solid var(--line)', borderRadius: '6px' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px', textAlign: 'left', background: 'var(--bg)' }}>
                    <thead>
                      <tr style={{ background: 'var(--panel)', borderBottom: '2px solid var(--line)', color: 'var(--ink)' }}>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)' }}>Exam Date & Session</th>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)' }}>Course Code</th>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)' }}>Course Title</th>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)', textAlign: 'center' }}>Total QPs</th>
                        <th style={{ padding: '10px 8px' }}>Venue Distribution</th>
                      </tr>
                    </thead>
                    <tbody>
                      {consolidatedData.map((row, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--line)' }}>
                          <td style={{ padding: '8px', borderRight: '1px solid var(--line)', fontWeight: 600 }}>{row.date} ({row.session})</td>
                          <td style={{ padding: '8px', borderRight: '1px solid var(--line)', color: 'var(--accent)', fontWeight: 600 }}>{row.courseCode}</td>
                          <td style={{ padding: '8px', borderRight: '1px solid var(--line)' }}>{row.courseName}</td>
                          <td style={{ padding: '8px', borderRight: '1px solid var(--line)', textAlign: 'center', fontWeight: 700 }}>{row.totalQPs}</td>
                          <td style={{ padding: '8px', color: 'var(--muted)', fontSize: '10px' }}>
                            {Object.entries(row.venues).map(([v, count]) => `${v} (Qty: ${count})`).join(', ')}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', border: '1px dashed var(--line)', borderRadius: '8px', padding: '60px', color: 'var(--muted)', fontSize: '14px', textAlign: 'center' }}>
                Upload QP Excel sheets and click "Generate QP Statements" to review consolidated values.
              </div>
            )
          ) : (
            Object.keys(venueSlips).length > 0 ? (
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '16px' }}>
                {/* Tabs */}
                <div style={{ display: 'flex', gap: '8px', overflowX: 'auto', paddingBottom: '8px', borderBottom: '1px solid var(--line)' }}>
                  {Object.keys(venueSlips).map((venName) => (
                    <button
                      key={venName}
                      onClick={() => setActiveVenueTab(venName)}
                      style={{
                        padding: '6px 12px',
                        borderRadius: '6px',
                        border: '1px solid var(--line)',
                        background: activeVenueTab === venName ? 'var(--accent)' : 'var(--panel)',
                        color: activeVenueTab === venName ? 'white' : 'var(--ink)',
                        whiteSpace: 'nowrap',
                        fontSize: '12px',
                        cursor: 'pointer',
                        fontWeight: activeVenueTab === venName ? 600 : 400
                      }}
                    >
                      {venName}
                    </button>
                  ))}
                </div>

                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <h4 style={{ margin: 0, color: 'var(--accent)' }}>Venue QP Slip</h4>
                    <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Dispatch target: {activeVenueTab}</span>
                  </div>
                  <button onClick={() => downloadVenuePDF(activeVenueTab)} style={{ padding: '6px 12px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                    <FileText size={14} /> Download PDF
                  </button>
                </div>

                <div style={{ flex: 1, overflowY: 'auto', maxHeight: '400px', border: '1px solid var(--line)', borderRadius: '6px' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11px', textAlign: 'left', background: 'var(--bg)' }}>
                    <thead>
                      <tr style={{ background: 'var(--panel)', borderBottom: '2px solid var(--line)', color: 'var(--ink)' }}>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)' }}>Exam Date & Session</th>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)' }}>Course Code</th>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)' }}>Course Title</th>
                        <th style={{ padding: '10px 8px', textAlign: 'center' }}>Required QPs</th>
                      </tr>
                    </thead>
                    <tbody>
                      {venueSlips[activeVenueTab].map((row, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--line)' }}>
                          <td style={{ padding: '8px', borderRight: '1px solid var(--line)' }}>{row.date} ({row.session})</td>
                          <td style={{ padding: '8px', borderRight: '1px solid var(--line)', color: 'var(--accent)', fontWeight: 600 }}>{row.courseCode}</td>
                          <td style={{ padding: '8px', borderRight: '1px solid var(--line)' }}>{row.courseName}</td>
                          <td style={{ padding: '8px', textAlign: 'center', fontWeight: 700 }}>{row.studentCount}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', border: '1px dashed var(--line)', borderRadius: '8px', padding: '60px', color: 'var(--muted)', fontSize: '14px', textAlign: 'center' }}>
                Upload QP Excel sheets and click "Generate QP Statements" to review venue-wise slips.
              </div>
            )
          )}
        </div>
      </div>
    </div>
  );
};

export default QpStatementPage;
