import React, { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import JSZip from 'jszip';
import { useAuth } from '../context/AuthContext';
import { Settings, Download, Eye, FileText, Plus, Trash2 } from 'lucide-react';

const QpLabelPage = () => {
  const { user } = useAuth();
  
  // File and sheet states
  const [selectedFile, setSelectedFile] = useState(null);
  const [workbook, setWorkbook] = useState(null);
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  
  // Column headers and mapping
  const [headers, setHeaders] = useState([]);
  const [columnMapping, setColumnMapping] = useState({
    centreCode: 0,
    centreName: 1,
    courseCode: 2,
    courseName: 3,
    day: 4,
    date: 5,
    time: 6
  });

  // Editor states for manual override defaults
  const [defaultDay, setDefaultDay] = useState('Monday');
  const [defaultDate, setDefaultDate] = useState('2025-10-13');
  const [defaultTime, setDefaultTime] = useState('10:00:00 - 11:30:00');
  const [examName, setExamName] = useState('Second Semester Degree (Private Registration) Regular Examinations April 2025');

  // Processing states
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('normal');

  // Grouped datasets
  const [labelGroups, setLabelGroups] = useState([]);
  const [activeLabelIndex, setActiveLabelIndex] = useState(0);

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

          // Try to auto-detect columns
          const autoMap = { ...columnMapping };
          firstRow.forEach((name, idx) => {
            const lower = name.toLowerCase().replace(/[\s_-]/g, '');
            if (lower.includes('centrecode') || lower.includes('centercode') || lower.includes('venuecode')) autoMap.centreCode = idx;
            if (lower.includes('centrename') || lower.includes('venue') || lower.includes('collegename')) autoMap.centreName = idx;
            if (lower.includes('coursecode') || (lower.includes('subject') && lower.includes('code'))) autoMap.courseCode = idx;
            if (lower.includes('coursename') || lower.includes('coursetitle') || lower.includes('subjectname') || lower.includes('subjecttitle')) autoMap.courseName = idx;
            if (lower.includes('day')) autoMap.day = idx;
            if (lower.includes('date')) autoMap.date = idx;
            if (lower.includes('time') || lower.includes('session')) autoMap.time = idx;
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
    setLabelGroups([]);
    setActiveLabelIndex(0);

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

  // Compile QP Labels
  const processQpLabels = () => {
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
    setStatus('Compiling QP Labels...');

    try {
      const dataRows = rows.slice(1);
      const groups = {};

      dataRows.forEach((row) => {
        const isBlank = row.every(cell => cell === null || cell === undefined || cell === '');
        if (isBlank) return;

        // Read mapped values
        const cCode = String(row[columnMapping.centreCode] || '').trim();
        const cName = String(row[columnMapping.centreName] || '').trim();
        const courseCode = String(row[columnMapping.courseCode] || '').trim();
        const courseName = String(row[columnMapping.courseName] || '').trim();

        const rowDay = row[columnMapping.day] ? String(row[columnMapping.day]).trim() : defaultDay;
        
        let rowDate = defaultDate;
        const rawDate = row[columnMapping.date];
        if (rawDate instanceof Date) {
          rowDate = rawDate.toLocaleDateString('en-CA'); // YYYY-MM-DD
        } else if (rawDate) {
          rowDate = String(rawDate).trim();
        }

        const rowTime = row[columnMapping.time] ? String(row[columnMapping.time]).trim() : defaultTime;

        if (!cCode || !courseCode) return; // Skip invalid records

        // Group key: unique Centre + Course
        const key = `${cCode}_${courseCode}`;

        if (!groups[key]) {
          groups[key] = {
            centreCode: cCode,
            centreName: cName,
            courseCode,
            courseName,
            day: rowDay,
            date: rowDate,
            time: rowTime,
            studentCount: 0
          };
        }
        groups[key].studentCount += 1;
      });

      const list = Object.values(groups);
      // Sort by Centre Code, then Course Code
      list.sort((a, b) => a.centreCode.localeCompare(b.centreCode) || a.courseCode.localeCompare(b.courseCode));
      
      setLabelGroups(list);
      setActiveLabelIndex(0);
      setStatus(`Compiled ${list.length} QP Labels successfully!`, 'success');
    } catch (err) {
      setStatus(`Error compiling labels: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Draw a Single QP Label Page on jsPDF
  const drawLabelPage = (doc, data, isFirstPage = true) => {
    if (!isFirstPage) {
      doc.addPage();
    }

    // A. Logo and University Headers
    // Placeholders for emblem
    doc.setDrawColor(200, 50, 50); // Red
    doc.setLineWidth(1.5);
    doc.ellipse(80, 70, 15, 20); // University Emblem outer ring
    doc.ellipse(80, 70, 10, 15);
    doc.line(75, 70, 85, 70);
    
    doc.setTextColor(200, 30, 30); // Red Title
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(16);
    doc.text('Kannur University', 110, 55);

    doc.setTextColor(30, 30, 30); // Dark text
    doc.setFontSize(11);
    doc.text('കണ്ണൂർ സർവകലാശാല', 110, 70);

    doc.setFont('Helvetica', 'normal');
    doc.setFontSize(9);
    doc.text('Thavakkara, Civil Station P.O, Kannur', 110, 82);
    doc.text('Reaccredited by NAAC with \'B++\' Grade', 110, 93);

    // B. Examination Subheading
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(11);
    doc.text('(Examination Branch)', 297, 125, { align: 'center' });

    doc.setFontSize(10.5);
    doc.text(examName, 297, 142, { align: 'center' });

    // C. Grid Layout Table (Exact Replica)
    // Table starting at X: 40pt, Y: 160pt. Total Width: 515pt.
    const startX = 40;
    const startY = 160;
    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(1);

    // Grid Row 1: CENTRE CODE AND NAME
    doc.rect(startX, startY, 515, 22);
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('CENTRE CODE AND NAME', startX + 8, startY + 14);
    doc.text(data.centreCode, startX + 160, startY + 14);
    doc.setFont('Helvetica', 'bold');
    doc.text(data.centreName, startX + 220, startY + 14);

    // Grid Verticals for Row 1
    doc.line(startX + 150, startY, startX + 150, startY + 22);
    doc.line(startX + 210, startY, startX + 210, startY + 22);

    // Grid Row 2: DAY, DATE, TIME
    const row2Y = startY + 22;
    doc.rect(startX, row2Y, 515, 22);
    doc.setFont('Helvetica', 'bold');
    doc.text('DAY', startX + 8, row2Y + 14);
    doc.setFont('Helvetica', 'normal');
    doc.text(data.day, startX + 160, row2Y + 14);

    doc.setFont('Helvetica', 'bold');
    doc.text('DATE', startX + 220, row2Y + 14);
    doc.setFont('Helvetica', 'normal');
    doc.text(data.date, startX + 280, row2Y + 14);

    doc.setFont('Helvetica', 'bold');
    doc.text('TIME', startX + 390, row2Y + 14);
    doc.setFont('Helvetica', 'normal');
    doc.text(data.time, startX + 440, row2Y + 14);

    // Grid Verticals for Row 2
    doc.line(startX + 150, row2Y, startX + 150, row2Y + 22); // Day column end
    doc.line(startX + 210, row2Y, startX + 210, row2Y + 22); // Date header start
    doc.line(startX + 270, row2Y, startX + 270, row2Y + 22); // Date value start
    doc.line(startX + 380, row2Y, startX + 380, row2Y + 22); // Time header start
    doc.line(startX + 430, row2Y, startX + 430, row2Y + 22); // Time value start

    // Grid Row 3: SUBJECT
    const row3Y = row2Y + 22;
    doc.rect(startX, row3Y, 515, 22);
    doc.setFont('Helvetica', 'bold');
    doc.text('SUBJECT', startX + 8, row3Y + 14);
    doc.setFont('Helvetica', 'bold');
    doc.text(`${data.courseCode} - ${data.courseName}`, startX + 160, row3Y + 14);

    // Vertical line for Row 3
    doc.line(startX + 150, row3Y, startX + 150, row3Y + 22);

    // Grid Row 4: NO. OF COPIES, COVER NUMBER
    const row4Y = row3Y + 22;
    doc.rect(startX, row4Y, 515, 22);
    doc.setFont('Helvetica', 'bold');
    doc.text('NO. OF COPIES', startX + 8, row4Y + 14);
    
    // Write strength as help context, but keep blank box clean
    doc.setFont('Helvetica', 'normal');
    doc.setFontSize(9);
    doc.text(`NC: ${data.studentCount}`, startX + 160, row4Y + 14);

    doc.setFont('Helvetica', 'bold');
    doc.text('COVER NUMBER', startX + 220, row4Y + 14);

    // Verticals for Row 4
    doc.line(startX + 150, row4Y, startX + 150, row4Y + 22);
    doc.line(startX + 210, row4Y, startX + 210, row4Y + 22);
    doc.line(startX + 380, row4Y, startX + 380, row4Y + 22);

    // D. Decorative Double Dotted/Dashed Line Break
    const lineBreakY = row4Y + 45;
    doc.setLineDash([3, 3], 0);
    doc.setDrawColor(120, 120, 120);
    doc.line(startX, lineBreakY, startX + 515, lineBreakY);
    doc.line(startX, lineBreakY + 4, startX + 515, lineBreakY + 4);
    doc.setLineDash([], 0); // Restore solid line format

    // E. Certificate Subheading
    const certY = lineBreakY + 30;
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(16);
    doc.setTextColor(0, 0, 0);
    doc.text('CERTIFICATE', 297, certY, { align: 'center' });

    // F. Certificate Description Body
    doc.setFont('Helvetica', 'normal');
    doc.setFontSize(11);
    doc.text('We hereby certify that we have examined this cover and satisfied ourselves that the seals are intact and that it was opened at', 40, certY + 30);
    doc.line(40, certY + 54, 210, certY + 54);
    doc.text('A.M/P.M in our presence.', 215, certY + 50);

    // G. Signatures Block
    const signY = certY + 90;
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('INVIGILATOR', 40, signY);
    doc.text('ADDL.CHIEF SUPERINTENDENT', 370, signY);

    // Invigilator signature lines
    doc.setFont('Helvetica', 'normal');
    doc.text('(1) ________________________________', 45, signY + 30);
    doc.text('(2) ________________________________', 45, signY + 55);

    // Place and Date left alignment, Chief Superintendent right alignment
    const footerY = signY + 100;
    doc.setFont('Helvetica', 'bold');
    doc.text('PLACE:', 40, footerY);
    doc.text('DATE:', 40, footerY + 25);
    doc.text('CHIEF SUPERINTENDENT', 390, footerY + 25);
  };

  // Download all combined in one PDF
  const downloadAllCombinedPDF = () => {
    if (labelGroups.length === 0) return;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    
    labelGroups.forEach((data, index) => {
      drawLabelPage(doc, data, index === 0);
    });

    doc.save(`${eventNameValPrefix()}_Question_Paper_Labels.pdf`);
  };

  // Download active index label PDF
  const downloadSinglePDF = (index) => {
    const data = labelGroups[index];
    if (!data) return;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    drawLabelPage(doc, data, true);

    const safeProg = data.courseCode.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
    const safeVenue = data.centreCode.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
    doc.save(`QP_Label_${safeVenue}_${safeProg}.pdf`);
  };

  // Batch ZIP downloads
  const downloadAllAsZip = async () => {
    if (labelGroups.length === 0) return;
    setStatus('Creating ZIP archive...', 'normal');
    const zip = new JSZip();

    try {
      labelGroups.forEach((data) => {
        const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
        drawLabelPage(doc, data, true);
        const pdfBlob = doc.output('blob');
        const safeProg = data.courseCode.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
        const safeVenue = data.centreCode.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
        zip.file(`QP_Label_${safeVenue}_${safeProg}.pdf`, pdfBlob);
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const downloadUrl = URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `QP_Labels_Package.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

      setStatus(`ZIP archive downloaded successfully!`, 'success');
    } catch (err) {
      setStatus(`Failed to generate ZIP: ${err.message}`, 'error');
    }
  };

  const eventNameValPrefix = () => {
    return examName.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_").substring(0, 30);
  };

  return (
    <div style={{ padding: '40px', maxWidth: '1400px', margin: '0 auto', overflowY: 'auto', width: '100%', minHeight: '100vh', background: 'var(--bg)', color: 'var(--ink)' }}>
      {/* Navigation Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px' }}>
        <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
          ← Back to Portal
        </Link>
        
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

      <h2>QP Label Generator</h2>
      <p className="subtitle">Compile print-ready covers and packet labels for examination question papers grouped by center and subject combinations.</p>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(450px, 1fr))', gap: '32px', marginBottom: '32px' }}>
        {/* Settings Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}><Settings size={20} /> Settings</h3>
          
          <div style={{ border: '1px dashed var(--line)', padding: '24px', borderRadius: '8px', background: 'var(--bg)', textAlign: 'center', position: 'relative' }}>
            <strong style={{ display: 'block', marginBottom: '8px' }}>Select Nominal Roll Excel</strong>
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
              <label>Examination Subtitle</label>
              <input type="text" value={examName} onChange={(e) => setExamName(e.target.value)} placeholder="e.g. Second Semester Degree..." />
            </div>
          </div>

          {/* Mappings */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)' }}>
            <strong style={{ fontSize: '13px', display: 'block', marginBottom: '12px' }}>Excel Column Mappings</strong>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', fontSize: '12px' }}>
              <div className="form-group">
                <label>Centre Code</label>
                <select value={columnMapping.centreCode} onChange={(e) => setColumnMapping({ ...columnMapping, centreCode: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Centre Name</label>
                <select value={columnMapping.centreName} onChange={(e) => setColumnMapping({ ...columnMapping, centreName: parseInt(e.target.value) })} disabled={headers.length === 0}>
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
                <label>Day Column</label>
                <select value={columnMapping.day} onChange={(e) => setColumnMapping({ ...columnMapping, day: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Date Column</label>
                <select value={columnMapping.date} onChange={(e) => setColumnMapping({ ...columnMapping, date: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Time Column</label>
                <select value={columnMapping.time} onChange={(e) => setColumnMapping({ ...columnMapping, time: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
            </div>
          </div>

          {/* Schedule Override Defaults */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)', display: 'flex', flexDirection: 'column', gap: '10px' }}>
            <strong style={{ fontSize: '13px' }}>Schedule Override Defaults (If missing in Excel)</strong>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '12px', fontSize: '11px' }}>
              <div className="form-group">
                <label>Default Day</label>
                <input type="text" value={defaultDay} onChange={(e) => setDefaultDay(e.target.value)} />
              </div>
              <div className="form-group">
                <label>Default Date</label>
                <input type="text" value={defaultDate} onChange={(e) => setDefaultDate(e.target.value)} />
              </div>
              <div className="form-group">
                <label>Default Time</label>
                <input type="text" value={defaultTime} onChange={(e) => setDefaultTime(e.target.value)} />
              </div>
            </div>
          </div>

          <button onClick={processQpLabels} disabled={isProcessing || !workbook} style={{ width: '100%', padding: '14px', fontSize: '15px' }}>
            {isProcessing ? "Processing Data..." : "Generate QP Labels"}
          </button>
        </div>

        {/* Live Preview Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}><Eye size={20} /> Live Preview</h3>
            
            {labelGroups.length > 0 && (
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={downloadAllCombinedPDF} style={{ padding: '8px 16px', fontSize: '13px' }}>
                  Download Combined PDF
                </button>
                <button onClick={downloadAllAsZip} className="secondary" style={{ padding: '8px 16px', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px', borderColor: 'var(--line)', color: 'var(--ink)' }}>
                  Download ZIP
                </button>
              </div>
            )}
          </div>

          {labelGroups.length > 0 ? (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', flex: 1 }}>
              {/* Pagination controls */}
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid var(--line)', paddingBottom: '10px' }}>
                <div style={{ fontSize: '13px' }}>
                  Label <strong>{activeLabelIndex + 1}</strong> of <strong>{labelGroups.length}</strong>
                </div>
                <div style={{ display: 'flex', gap: '8px' }}>
                  <button 
                    onClick={() => setActiveLabelIndex(Math.max(0, activeLabelIndex - 1))}
                    disabled={activeLabelIndex === 0}
                    style={{ padding: '4px 10px', fontSize: '12px' }}
                  >
                    Prev
                  </button>
                  <button 
                    onClick={() => setActiveLabelIndex(Math.min(labelGroups.length - 1, activeLabelIndex + 1))}
                    disabled={activeLabelIndex === labelGroups.length - 1}
                    style={{ padding: '4px 10px', fontSize: '12px' }}
                  >
                    Next
                  </button>
                  <button onClick={() => downloadSinglePDF(activeLabelIndex)} style={{ padding: '4px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}>
                    <FileText size={12} /> Download PDF
                  </button>
                </div>
              </div>

              {/* Visual simulated label rendering */}
              <div style={{ 
                flex: 1, 
                border: '1px solid var(--line)', 
                borderRadius: '8px', 
                background: 'white', 
                color: 'black', 
                padding: '30px', 
                fontSize: '11px',
                fontFamily: 'Helvetica, Arial, sans-serif',
                boxShadow: '0 4px 10px rgba(0,0,0,0.05)',
                display: 'flex',
                flexDirection: 'column',
                gap: '24px'
              }}>
                {/* Header */}
                <div style={{ display: 'flex', gap: '16px' }}>
                  <div style={{ width: '40px', height: '50px', border: '1.5px solid #d32f2f', borderRadius: '50%', display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: '18px', color: '#d32f2f', fontWeight: 'bold' }}>KU</div>
                  <div>
                    <h3 style={{ margin: 0, fontSize: '18px', color: '#d32f2f', fontWeight: 'bold' }}>Kannur University</h3>
                    <h4 style={{ margin: '2px 0 0 0', fontSize: '12px', color: '#333' }}>കണ്ണൂർ സർവകലാശാല</h4>
                    <span style={{ fontSize: '9px', color: '#666', display: 'block' }}>Thavakkara, Civil Station P.O, Kannur</span>
                    <span style={{ fontSize: '9px', color: '#666', display: 'block' }}>Reaccredited by NAAC with 'B++' Grade</span>
                  </div>
                </div>

                <div style={{ textAlign: 'center', display: 'flex', flexDirection: 'column', gap: '4px' }}>
                  <strong style={{ fontSize: '11px' }}>(Examination Branch)</strong>
                  <strong style={{ fontSize: '11px' }}>{examName}</strong>
                </div>

                {/* Simulated table grid */}
                <div style={{ border: '1px solid black', display: 'flex', flexDirection: 'column' }}>
                  <div style={{ display: 'flex', borderBottom: '1px solid black' }}>
                    <div style={{ width: '130px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>CENTRE CODE AND NAME</div>
                    <div style={{ width: '60px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>{labelGroups[activeLabelIndex].centreCode}</div>
                    <div style={{ flex: 1, padding: '6px', fontWeight: 'bold' }}>{labelGroups[activeLabelIndex].centreName}</div>
                  </div>
                  
                  <div style={{ display: 'flex', borderBottom: '1px solid black' }}>
                    <div style={{ width: '60px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>DAY</div>
                    <div style={{ width: '130px', padding: '6px', borderRight: '1px solid black' }}>{labelGroups[activeLabelIndex].day}</div>
                    <div style={{ width: '60px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>DATE</div>
                    <div style={{ width: '110px', padding: '6px', borderRight: '1px solid black' }}>{labelGroups[activeLabelIndex].date}</div>
                    <div style={{ width: '60px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>TIME</div>
                    <div style={{ flex: 1, padding: '6px' }}>{labelGroups[activeLabelIndex].time}</div>
                  </div>

                  <div style={{ display: 'flex', borderBottom: '1px solid black' }}>
                    <div style={{ width: '130px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>SUBJECT</div>
                    <div style={{ flex: 1, padding: '6px', fontWeight: 'bold' }}>
                      {labelGroups[activeLabelIndex].courseCode} - {labelGroups[activeLabelIndex].courseName}
                    </div>
                  </div>

                  <div style={{ display: 'flex' }}>
                    <div style={{ width: '130px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>NO. OF COPIES</div>
                    <div style={{ width: '60px', padding: '6px', borderRight: '1px solid black', color: '#888' }}>NC: {labelGroups[activeLabelIndex].studentCount}</div>
                    <div style={{ width: '130px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>COVER NUMBER</div>
                    <div style={{ flex: 1, padding: '6px' }}></div>
                  </div>
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', borderTop: '2px dashed #999', paddingTop: '16px' }}>
                  <h4 style={{ textAlign: 'center', margin: '0 0 10px 0', fontSize: '14px', letterSpacing: '0.5px' }}>CERTIFICATE</h4>
                  <p style={{ margin: 0, lineHeight: '1.6' }}>
                    We hereby certify that we have examined this cover and satisfied ourselves that the seals are intact and that it was opened at ________________________________A.M/P.M in our presence.
                  </p>
                </div>

                <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '10px', fontWeight: 'bold' }}>
                  <div>
                    <span>INVIGILATOR</span>
                    <div style={{ marginTop: '20px', fontWeight: 'normal' }}>(1) ________________________________</div>
                    <div style={{ marginTop: '10px', fontWeight: 'normal' }}>(2) ________________________________</div>
                  </div>
                  <div>
                    <span>ADDL.CHIEF SUPERINTENDENT</span>
                  </div>
                </div>

                <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '16px', fontWeight: 'bold' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
                    <span>PLACE:</span>
                    <span>DATE:</span>
                  </div>
                  <div style={{ alignSelf: 'flex-end' }}>
                    <span>CHIEF SUPERINTENDENT</span>
                  </div>
                </div>
              </div>
            </div>
          ) : (
            <div style={{ display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', border: '1px dashed var(--line)', borderRadius: '8px', padding: '60px', color: 'var(--muted)', fontSize: '14px', textAlign: 'center' }}>
              Upload your nominal roll Excel sheet, review your mappings, and click "Generate QP Labels" to render visual slips.
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default QpLabelPage;
