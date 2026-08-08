import React, { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import JSZip from 'jszip';
import { useAuth } from '../context/AuthContext';
import { Settings, Download, Eye, FileText, Plus, Trash2, HelpCircle } from 'lucide-react';
import { logoBase64 } from '../assets/logoBase64';

const QpLabelPage = () => {
  const { user } = useAuth();
  
  // File states
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [consolidatedRows, setConsolidatedRows] = useState([]);
  
  // Column headers and mapping
  const [headers, setHeaders] = useState([]);
  const [columnMapping, setColumnMapping] = useState({
    centreCode: 3,
    centreName: 4,
    courseCode: 5,
    courseName: 6,
    date: -1,
    courseStartTime: -1,
    courseEndTime: -1
  });

  // Editor states
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

  // Helper function to read Excel rows as Promise
  const readExcelFile = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = e.target.result;
          const wb = XLSX.read(data, { type: 'array', cellDates: true });
          const firstSheet = wb.SheetNames[0];
          const sheet = wb.Sheets[firstSheet];
          if (!sheet) {
            resolve([]);
            return;
          }
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
          resolve(rows);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = (err) => reject(err);
      reader.readAsArrayBuffer(file);
    });
  };

  const handleFileChange = async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    setSelectedFiles(files);
    setStatus(`Reading ${files.length} file(s)...`);
    setLabelGroups([]);
    setActiveLabelIndex(0);

    try {
      const allFilesRows = await Promise.all(files.map(file => readExcelFile(file)));
      
      let masterRows = [];
      let firstHeaders = [];

      allFilesRows.forEach((rows, index) => {
        if (rows.length > 0) {
          if (masterRows.length === 0) {
            firstHeaders = rows[0].map((cell, idx) => cell ? String(cell).trim() : `Column ${idx + 1}`);
            masterRows = [...rows];
          } else {
            // Append data rows excluding headers
            masterRows = masterRows.concat(rows.slice(1));
          }
        }
      });

      if (masterRows.length > 0) {
        setHeaders(firstHeaders);
        setConsolidatedRows(masterRows);

        // Auto-detect columns based on first file's headers
        const autoMap = { ...columnMapping };
        let foundVenueCode = false;
        let foundVenueName = false;

        firstHeaders.forEach((name, idx) => {
          const lower = name.toLowerCase().replace(/[\s_-]/g, '');
          
          // Prioritize Venue Code / Centre Code
          if (lower === 'venuecode' || lower === 'centrecode' || lower === 'centercode') {
            autoMap.centreCode = idx;
            foundVenueCode = true;
          } else if (!foundVenueCode && (lower.includes('venuecode') || lower.includes('centrecode') || lower.includes('centercode'))) {
            autoMap.centreCode = idx;
            foundVenueCode = true;
          } else if (!foundVenueCode && lower.includes('venueid')) {
            autoMap.centreCode = idx;
          }

          // Prioritize Venue Name / Centre Name
          if (lower === 'venuename' || lower === 'centrename' || lower === 'collegename') {
            autoMap.centreName = idx;
            foundVenueName = true;
          } else if (!foundVenueName && (lower.includes('venuename') || lower.includes('centrename') || lower.includes('collegename'))) {
            autoMap.centreName = idx;
            foundVenueName = true;
          } else if (!foundVenueName && lower.includes('venue') && !lower.includes('code') && !lower.includes('id')) {
            autoMap.centreName = idx;
          }

          if (lower.includes('coursecode') || (lower.includes('subject') && lower.includes('code'))) autoMap.courseCode = idx;
          if (lower.includes('coursename') || lower.includes('coursetitle') || lower.includes('subjectname') || lower.includes('subjecttitle')) autoMap.courseName = idx;
          if (lower.includes('date')) autoMap.date = idx;
          if (lower.includes('starttime') || (lower.includes('start') && lower.includes('time'))) autoMap.courseStartTime = idx;
          if (lower.includes('endtime') || (lower.includes('end') && lower.includes('time'))) autoMap.courseEndTime = idx;
        });
        setColumnMapping(autoMap);

        // Try auto-setting Exam Name from first column first row
        if (masterRows.length > 1 && masterRows[1][0] && String(masterRows[1][0]).length > 10) {
          setExamName(String(masterRows[1][0]).trim());
        }

        setStatus(`Loaded and merged ${files.length} file(s) with ${masterRows.length - 1} rows.`, 'success');
      } else {
        setStatus('No data rows found in uploaded files.', 'error');
      }
    } catch (err) {
      setStatus(`Error reading/merging files: ${err.message}`, 'error');
      setConsolidatedRows([]);
      setHeaders([]);
    }
  };

  // Compile QP Labels
  const processQpLabels = () => {
    if (consolidatedRows.length < 2) {
      setStatus('No uploaded data to compile labels from.', 'error');
      return;
    }

    setIsProcessing(true);
    setStatus('Compiling QP Labels...');

    try {
      const dataRows = consolidatedRows.slice(1);
      const groups = {};

      dataRows.forEach((row) => {
        const isBlank = row.every(cell => cell === null || cell === undefined || cell === '');
        if (isBlank) return;

        const cCode = String(row[columnMapping.centreCode] || '').trim();
        const cName = String(row[columnMapping.centreName] || '').trim();
        const courseCode = String(row[columnMapping.courseCode] || '').trim();
        const courseName = String(row[columnMapping.courseName] || '').trim();

        let rowDate = '';
        let dateObj = null;
        if (columnMapping.date !== -1) {
          const rawDate = row[columnMapping.date];
          if (rawDate instanceof Date) {
            dateObj = rawDate;
            rowDate = rawDate.toLocaleDateString('en-CA');
          } else if (rawDate) {
            rowDate = String(rawDate).trim();
            const parsed = Date.parse(rowDate);
            if (!isNaN(parsed)) {
              dateObj = new Date(parsed);
            }
          }
        }

        // Derive day from date
        let rowDay = '';
        if (dateObj) {
          const daysOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          rowDay = daysOfWeek[dateObj.getDay()];
        }

        // Combine course start time and end time
        const startTimeVal = columnMapping.courseStartTime !== -1 && row[columnMapping.courseStartTime] ? String(row[columnMapping.courseStartTime]).trim() : '';
        const endTimeVal = columnMapping.courseEndTime !== -1 && row[columnMapping.courseEndTime] ? String(row[columnMapping.courseEndTime]).trim() : '';
        let rowTime = '';
        if (startTimeVal && endTimeVal) {
          rowTime = `${startTimeVal} - ${endTimeVal}`;
        } else if (startTimeVal) {
          rowTime = startTimeVal;
        } else if (endTimeVal) {
          rowTime = endTimeVal;
        }

        if (!cCode || !courseCode) return;

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

  // Draw a Single QP Label Page on jsPDF (Landscape format)
  const drawLabelPage = (doc, data, isFirstPage = true) => {
    if (!isFirstPage) {
      doc.addPage();
    }

    // A4 Landscape Width: 842 pt, Height: 595 pt.
    // Center alignment point: X = 421 pt.

    // A. Logo Header (Centered)
    const logoWidth = 320;
    const logoHeight = 90;
    const logoX = (842 - logoWidth) / 2;
    doc.addImage(logoBase64, 'PNG', logoX, 25, logoWidth, logoHeight);

    // B. Examination Subheading
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(11);
    doc.text('(Examination Branch)', 421, 130, { align: 'center' });

    doc.setFontSize(10.5);
    doc.text(examName, 421, 147, { align: 'center' });

    // C. Grid Layout Table (Landscape dimensions)
    const startX = 121;
    const startY = 165;
    const tableWidth = 600;
    const rowHeight = 26;

    doc.setDrawColor(0, 0, 0);
    doc.setLineWidth(1);

    // Grid Row 1: CENTRE CODE AND NAME
    doc.rect(startX, startY, tableWidth, rowHeight);
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('CENTRE CODE AND NAME', startX + 8, startY + 17);
    doc.text(data.centreCode, startX + 160, startY + 17);
    doc.text(data.centreName, startX + 230, startY + 17);

    // Grid Verticals for Row 1
    doc.line(startX + 150, startY, startX + 150, startY + rowHeight);
    doc.line(startX + 220, startY, startX + 220, startY + rowHeight);

    // Grid Row 2: DAY, DATE, TIME
    const row2Y = startY + rowHeight;
    doc.rect(startX, row2Y, tableWidth, rowHeight);
    doc.text('DAY', startX + 8, row2Y + 17);
    doc.setFont('Helvetica', 'normal');
    doc.text(data.day, startX + 68, row2Y + 17);

    doc.setFont('Helvetica', 'bold');
    doc.text('DATE', startX + 188, row2Y + 17);
    doc.setFont('Helvetica', 'normal');
    doc.text(data.date, startX + 238, row2Y + 17);

    doc.setFont('Helvetica', 'bold');
    doc.text('TIME', startX + 368, row2Y + 17);
    doc.setFont('Helvetica', 'normal');
    doc.text(data.time, startX + 418, row2Y + 17);

    // Grid Verticals for Row 2
    doc.line(startX + 60, row2Y, startX + 60, row2Y + rowHeight);
    doc.line(startX + 180, row2Y, startX + 180, row2Y + rowHeight);
    doc.line(startX + 230, row2Y, startX + 230, row2Y + rowHeight);
    doc.line(startX + 360, row2Y, startX + 360, row2Y + rowHeight);
    doc.line(startX + 410, row2Y, startX + 410, row2Y + rowHeight);

    // Grid Row 3: SUBJECT
    const row3Y = row2Y + rowHeight;
    doc.rect(startX, row3Y, tableWidth, rowHeight);
    doc.setFont('Helvetica', 'bold');
    doc.text('SUBJECT', startX + 8, row3Y + 17);
    doc.text(`${data.courseCode} - ${data.courseName}`, startX + 160, row3Y + 17);

    // Vertical line for Row 3
    doc.line(startX + 150, row3Y, startX + 150, row3Y + rowHeight);

    // Grid Row 4: NO. OF COPIES, COVER NUMBER
    const row4Y = row3Y + rowHeight;
    doc.rect(startX, row4Y, tableWidth, rowHeight);
    doc.text('NO. OF COPIES', startX + 8, row4Y + 17);

    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('COVER NUMBER', startX + 230, row4Y + 17);

    // Verticals for Row 4
    doc.line(startX + 150, row4Y, startX + 150, row4Y + rowHeight);
    doc.line(startX + 220, row4Y, startX + 220, row4Y + rowHeight);
    doc.line(startX + 400, row4Y, startX + 400, row4Y + rowHeight);

    // D. Decorative Double Dotted Line Break
    const lineBreakY = row4Y + 40;
    doc.setLineDash([3, 3], 0);
    doc.setDrawColor(120, 120, 120);
    doc.line(startX, lineBreakY, startX + tableWidth, lineBreakY);
    doc.line(startX, lineBreakY + 4, startX + tableWidth, lineBreakY + 4);
    doc.setLineDash([], 0);

    // E. Certificate Subheading
    const certY = lineBreakY + 28;
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(16);
    doc.setTextColor(0, 0, 0);
    doc.text('CERTIFICATE', 421, certY, { align: 'center' });

    // F. Certificate Description Body
    doc.setFont('Helvetica', 'normal');
    doc.setFontSize(11);
    doc.text('We hereby certify that we have examined this cover and satisfied ourselves that the seals are intact and that it was opened at', startX, certY + 30);
    doc.text('______________________________________ A.M/P.M in our presence.', startX, certY + 52);

    // G. Signatures Block
    const signY = certY + 95;
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(10);
    doc.text('INVIGILATOR', startX, signY);
    doc.text('ADDL.CHIEF SUPERINTENDENT', startX + tableWidth, signY, { align: 'right' });

    // Invigilator signature lines
    doc.setFont('Helvetica', 'normal');
    doc.text('(1) ________________________________', startX + 5, signY + 25);
    doc.text('(2) ________________________________', startX + 5, signY + 48);

    // Place and Date left alignment, Chief Superintendent right alignment
    const footerY = signY + 85;
    doc.setFont('Helvetica', 'bold');
    doc.text('PLACE:', startX, footerY);
    doc.text('DATE:', startX, footerY + 22);
    doc.text('CHIEF SUPERINTENDENT', startX + tableWidth, footerY + 22, { align: 'right' });
  };

  // Download all combined in one Landscape PDF
  const downloadAllCombinedPDF = () => {
    if (labelGroups.length === 0) return;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    
    labelGroups.forEach((data, index) => {
      drawLabelPage(doc, data, index === 0);
    });

    doc.save(`${eventNameValPrefix()}_Question_Paper_Labels_Landscape.pdf`);
  };

  // Download active index label PDF
  const downloadSinglePDF = (index) => {
    const data = labelGroups[index];
    if (!data) return;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    drawLabelPage(doc, data, true);

    const safeProg = data.courseCode.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
    const safeVenue = data.centreCode.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
    doc.save(`QP_Label_Landscape_${safeVenue}_${safeProg}.pdf`);
  };

  // Batch ZIP downloads
  const downloadAllAsZip = async () => {
    if (labelGroups.length === 0) return;
    setStatus('Creating ZIP archive...', 'normal');
    const zip = new JSZip();

    try {
      labelGroups.forEach((data) => {
        const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
        drawLabelPage(doc, data, true);
        const pdfBlob = doc.output('blob');
        const safeProg = data.courseCode.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
        const safeVenue = data.centreCode.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
        zip.file(`QP_Label_Landscape_${safeVenue}_${safeProg}.pdf`, pdfBlob);
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const downloadUrl = URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `QP_Labels_Landscape_Package.zip`;
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

      <h2>QP Label Generator (Landscape)</h2>
      <p className="subtitle">Compile print-ready envelope covers and packet labels grouped by unique center and subject combinations.</p>

      {/* Instructional Banner */}
      <div style={{
        padding: '16px 20px',
        borderRadius: '8px',
        background: 'var(--accent-soft)',
        border: '1px solid var(--accent)',
        color: 'var(--ink)',
        fontSize: '13.5px',
        lineHeight: '1.6',
        marginBottom: '28px',
        display: 'flex',
        alignItems: 'flex-start',
        gap: '12px'
      }}>
        <HelpCircle size={20} color="var(--accent)" style={{ flexShrink: 0, marginTop: '2px' }} />
        <div>
          <strong>📌 Instructions:</strong> Please upload the <strong>Course-wise Venue-wise Date-wise Nominal Roll / Report</strong> Excel sheet(s).
          You can upload <strong>multiple reports at once</strong> (by holding Ctrl/Cmd during selection). The generator will automatically merge them into a single consolidated master list before compiling the labels.
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(450px, 1fr))', gap: '32px', marginBottom: '32px' }}>
        {/* Settings Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}><Settings size={20} /> Settings</h3>
          
          <div style={{ border: '1px dashed var(--line)', padding: '24px', borderRadius: '8px', background: 'var(--bg)', textAlign: 'center', position: 'relative' }}>
            <strong style={{ display: 'block', marginBottom: '8px' }}>Upload Nominal Roll Reports (Supports Multiple)</strong>
            {selectedFiles.length > 0 ? (
              <div style={{ color: 'var(--accent)', fontWeight: 600 }}>
                {selectedFiles.length === 1 ? `📄 ${selectedFiles[0].name}` : `📂 ${selectedFiles.length} files selected`}
              </div>
            ) : (
              <span style={{ color: 'var(--muted)', fontSize: '13px' }}>Drag here or click to browse multiple reports</span>
            )}
            <input 
              type="file" 
              multiple 
              accept=".xlsx, .xls, .xlsm, .csv" 
              onChange={handleFileChange} 
              style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, opacity: 0, cursor: 'pointer', width: '100%' }} 
            />
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr', gap: '16px' }}>
            <div className="form-group">
              <label>Examination Title Subtitle</label>
              <input type="text" value={examName} onChange={(e) => setExamName(e.target.value)} placeholder="e.g. Second Semester Degree..." />
            </div>
          </div>

          {/* Mappings */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)' }}>
            <strong style={{ fontSize: '13px', display: 'block', marginBottom: '12px' }}>Excel Column Mappings</strong>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', fontSize: '12px' }}>
              <div className="form-group">
                <label>Centre Code / Venue Code</label>
                <select value={columnMapping.centreCode} onChange={(e) => setColumnMapping({ ...columnMapping, centreCode: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Centre Name / Venue Name</label>
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
                <label>Date Column</label>
                <select value={columnMapping.date} onChange={(e) => setColumnMapping({ ...columnMapping, date: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  <option value="-1">- Keep Blank -</option>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Course Start Time</label>
                <select value={columnMapping.courseStartTime} onChange={(e) => setColumnMapping({ ...columnMapping, courseStartTime: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  <option value="-1">- Keep Blank -</option>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group" style={{ gridColumn: 'span 2' }}>
                <label>Course End Time</label>
                <select value={columnMapping.courseEndTime} onChange={(e) => setColumnMapping({ ...columnMapping, courseEndTime: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  <option value="-1">- Keep Blank -</option>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
            </div>
          </div>

          <button onClick={processQpLabels} disabled={isProcessing || consolidatedRows.length === 0} style={{ width: '100%', padding: '14px', fontSize: '15px' }}>
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

              {/* Visual simulated label rendering (Landscape design) */}
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
                gap: '20px'
              }}>
                {/* Header Logo Centered */}
                <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '4px' }}>
                  <img src={logoBase64} alt="Kannur University Logo" style={{ maxWidth: '300px', height: 'auto' }} />
                </div>

                <div style={{ textAlign: 'center', display: 'flex', flexDirection: 'column', gap: '2px' }}>
                  <strong style={{ fontSize: '10px' }}>(Examination Branch)</strong>
                  <strong style={{ fontSize: '10px' }}>{examName}</strong>
                </div>

                {/* Simulated table grid (Landscape spacing) */}
                <div style={{ border: '1px solid black', display: 'flex', flexDirection: 'column' }}>
                  <div style={{ display: 'flex', borderBottom: '1px solid black' }}>
                    <div style={{ width: '150px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>CENTRE CODE AND NAME</div>
                    <div style={{ width: '70px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>{labelGroups[activeLabelIndex].centreCode}</div>
                    <div style={{ flex: 1, padding: '6px', fontWeight: 'bold' }}>{labelGroups[activeLabelIndex].centreName}</div>
                  </div>
                  
                  <div style={{ display: 'flex', borderBottom: '1px solid black' }}>
                    <div style={{ width: '60px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>DAY</div>
                    <div style={{ width: '120px', padding: '6px', borderRight: '1px solid black' }}>{labelGroups[activeLabelIndex].day}</div>
                    <div style={{ width: '50px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>DATE</div>
                    <div style={{ width: '130px', padding: '6px', borderRight: '1px solid black' }}>{labelGroups[activeLabelIndex].date}</div>
                    <div style={{ width: '50px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>TIME</div>
                    <div style={{ flex: 1, padding: '6px' }}>{labelGroups[activeLabelIndex].time}</div>
                  </div>

                  <div style={{ display: 'flex', borderBottom: '1px solid black' }}>
                    <div style={{ width: '150px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>SUBJECT</div>
                    <div style={{ flex: 1, padding: '6px', fontWeight: 'bold' }}>
                      {labelGroups[activeLabelIndex].courseCode} - {labelGroups[activeLabelIndex].courseName}
                    </div>
                  </div>

                  <div style={{ display: 'flex' }}>
                    <div style={{ width: '150px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>NO. OF COPIES</div>
                    <div style={{ width: '70px', padding: '6px', borderRight: '1px solid black' }}></div>
                    <div style={{ width: '150px', padding: '6px', borderRight: '1px solid black', fontWeight: 'bold' }}>COVER NUMBER</div>
                    <div style={{ flex: 1, padding: '6px' }}></div>
                  </div>
                </div>

                <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', borderTop: '2px dashed #999', paddingTop: '12px' }}>
                  <h4 style={{ textAlign: 'center', margin: '0 0 6px 0', fontSize: '13px', letterSpacing: '0.5px' }}>CERTIFICATE</h4>
                  <p style={{ margin: 0, lineHeight: '1.5' }}>
                    We hereby certify that we have examined this cover and satisfied ourselves that the seals are intact and that it was opened at ______________________________________ A.M/P.M in our presence.
                  </p>
                </div>

                <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '5px', fontWeight: 'bold' }}>
                  <div>
                    <span>INVIGILATOR</span>
                    <div style={{ marginTop: '12px', fontWeight: 'normal' }}>(1) ________________________________</div>
                    <div style={{ marginTop: '6px', fontWeight: 'normal' }}>(2) ________________________________</div>
                  </div>
                  <div>
                    <span>ADDL.CHIEF SUPERINTENDENT</span>
                  </div>
                </div>

                <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '10px', fontWeight: 'bold' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
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
              Upload your nominal roll Excel sheets, review your mappings, and click "Generate QP Labels" to render visual slips.
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default QpLabelPage;
