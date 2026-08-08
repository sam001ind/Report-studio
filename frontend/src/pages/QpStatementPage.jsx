import React, { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import JSZip from 'jszip';
import { useAuth } from '../context/AuthContext';
import { Settings, Download, Eye, FileText, Plus, Trash2, AlignCenter, AlignLeft, AlignRight, Bold, HelpCircle } from 'lucide-react';

const QpStatementPage = () => {
  const { user } = useAuth();
  
  // File and sheet states (supports multiple files)
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [consolidatedRows, setConsolidatedRows] = useState([]);
  
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

  // Dynamic Header Lines matching screenshot formatting
  const [headerLines, setHeaderLines] = useState([
    { text: 'Kannur University', fontSize: 16, isBold: true, align: 'center', yOffset: 40 },
    { text: '(Examination Branch)', fontSize: 12, isBold: true, align: 'center', yOffset: 60 },
    { text: 'QP Statement for 1st Semester Degree Private Registration Regular/Supplementary Examination', fontSize: 10, isBold: true, align: 'center', yOffset: 78 },
    { text: 'November 2025', fontSize: 10, isBold: true, align: 'center', yOffset: 92 }
  ]);

  // Default PDF Columns matching screenshot: SL No, Date, Course, NC, QP, LP
  const [tableColumns, setTableColumns] = useState([
    { id: 'slNo', label: 'SL No', width: 35, align: 'center', fontSize: 9.5 },
    { id: 'date', label: 'Date', width: 110, align: 'left', fontSize: 9.5 },
    { id: 'course', label: 'Course', width: 220, align: 'left', fontSize: 9.5 },
    { id: 'nc', label: 'NC', width: 40, align: 'left', fontSize: 9.5 },
    { id: 'qp', label: 'QP', width: 55, align: 'center', fontSize: 9.5 },
    { id: 'lp', label: 'LP', width: 55, align: 'center', fontSize: 9.5 }
  ]);

  // Report generation mode: 'venue' (Separate slip per center) vs 'consolidated' (Date-wise summary)
  const [reportMode, setReportMode] = useState('venue');

  // Processing states
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

  // Helper function to format Date to 'YYYY-MM-DD DayName'
  const formatDateWithDay = (dateVal) => {
    if (!dateVal) return '';
    let dateObj;

    // Handle serial dates or direct Dates
    if (dateVal instanceof Date) {
      dateObj = dateVal;
    } else if (typeof dateVal === 'number') {
      // Excel Serial Date conversion
      dateObj = new Date((dateVal - 25569) * 86400 * 1000);
    } else {
      dateObj = new Date(dateVal);
    }

    if (isNaN(dateObj.getTime())) {
      return String(dateVal).trim();
    }

    const yyyy = dateObj.getFullYear();
    const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    const dd = String(dateObj.getDate()).padStart(2, '0');

    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const dayName = days[dateObj.getDay()];

    return `${yyyy}-${mm}-${dd} ${dayName}`;
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
    setConsolidatedData([]);
    setVenueSlips({});
    setActiveVenueTab('');

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

        // Get default Event Name from first data row
        if (masterRows.length > 1 && masterRows[1][0] && String(masterRows[1][0]).length > 10) {
          setEventNameVal(String(masterRows[1][0]).trim());
          
          // Update the 3rd header line with the detected exam name
          const updatedHeaders = [...headerLines];
          if (updatedHeaders[2]) {
            updatedHeaders[2].text = String(masterRows[1][0]).trim();
            setHeaderLines(updatedHeaders);
          }
        }

        const autoMap = { ...columnMapping };
        firstHeaders.forEach((name, idx) => {
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

  // Compile QP Statement
  const processQpStatement = () => {
    if (consolidatedRows.length < 2) {
      setStatus('No uploaded data to compile statements from.', 'error');
      return;
    }

    setIsProcessing(true);
    setStatus('Grouping question paper statements...');

    try {
      const dataRows = consolidatedRows.slice(1);
      const tempConsolidated = {};
      const tempVenueSlips = {};

      dataRows.forEach((row) => {
        const isBlank = row.every(cell => cell === null || cell === undefined || cell === '');
        if (isBlank) return;

        const event = String(row[columnMapping.eventName] || eventNameVal || '').trim();
        
        // Match Center Name styling (NK - Naher Arts and Science College, Kanhirode)
        const cCode = String(row[columnMapping.centerCode] || '').trim();
        const cName = String(row[columnMapping.centerName] || '').trim();
        const venue = cCode && cName ? `${cCode} - ${cName}` : (cCode || cName || 'Unassigned Venue');
        
        const dateRaw = row[columnMapping.examDate];
        const dateFormatted = formatDateWithDay(dateRaw);

        const cCodeVal = String(row[columnMapping.courseCode] || '').trim();
        const cNameVal = String(row[columnMapping.courseName] || '').trim();
        const count = parseInt(row[columnMapping.studentCount]) || 0;

        if (!cCodeVal || !dateFormatted) return;

        // A. Process Consolidated Summary (Date + Course)
        const conKey = `${dateFormatted} | ${cCodeVal}`;
        if (!tempConsolidated[conKey]) {
          tempConsolidated[conKey] = {
            date: dateFormatted,
            courseCode: cCodeVal,
            courseName: cNameVal,
            totalQPs: 0,
            venues: {}
          };
        }
        tempConsolidated[conKey].totalQPs += count;
        
        // Add venue count for breakdown block
        if (!tempConsolidated[conKey].venues[venue]) {
          tempConsolidated[conKey].venues[venue] = 0;
        }
        tempConsolidated[conKey].venues[venue] += count;

        // B. Process Venue Packing Slips (Separate Slip per Venue)
        if (!tempVenueSlips[venue]) {
          tempVenueSlips[venue] = [];
        }
        tempVenueSlips[venue].push({
          date: dateFormatted,
          courseCode: cCodeVal,
          courseName: cNameVal,
          studentCount: count
        });
      });

      // Format Consolidated Summary Table data
      const sortedConsolidated = Object.values(tempConsolidated);
      // Sort by Date, then Course Code
      sortedConsolidated.sort((a, b) => a.date.localeCompare(b.date) || a.courseCode.localeCompare(b.courseCode));
      setConsolidatedData(sortedConsolidated);

      // Sort courses inside each Venue packing slip by date
      Object.keys(tempVenueSlips).forEach((venue) => {
        tempVenueSlips[venue].sort((a, b) => a.date.localeCompare(b.date) || a.courseCode.localeCompare(b.courseCode));
      });
      setVenueSlips(tempVenueSlips);

      const venueList = Object.keys(tempVenueSlips);
      if (venueList.length > 0) {
        setActiveVenueTab(venueList[0]);
      }

      setStatus(`Processed ${sortedConsolidated.length} date-wise blocks across ${venueList.length} venues.`, 'success');
    } catch (err) {
      setStatus(`Compilation failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Helper: Draw Header on jsPDF document page
  const drawPageHeaders = (doc, titleText = '') => {
    headerLines.forEach((line) => {
      doc.setFont('Helvetica', line.isBold ? 'bold' : 'normal');
      doc.setFontSize(line.fontSize);
      doc.text(line.text, line.align === 'center' ? 297 : line.align === 'right' ? 550 : 45, line.yOffset, { align: line.align });
    });

    if (titleText) {
      doc.setFont('Helvetica', 'bold');
      doc.setFontSize(11);
      doc.text(titleText, 45, 115);
    }

    doc.setDrawColor(180, 180, 180);
    doc.setLineWidth(1);
    doc.line(45, 122, 550, 122); // Elegant horizontal line separation
  };

  // Generate and Download Date-Wise Consolidated PDF
  const downloadConsolidatedPDF = () => {
    if (consolidatedData.length === 0) return;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });

    // Draw header and content using autoTable
    drawPageHeaders(doc, 'Date-wise Consolidated Printing Summary');

    const tableRows = consolidatedData.map((row, idx) => [
      idx + 1,
      row.date,
      `${row.courseCode}\n${row.courseName}`,
      row.totalQPs,
      '', // Blank QP cell
      ''  // Blank LP cell
    ]);

    const cols = tableColumns.map(c => ({ header: c.label, dataKey: c.id }));
    const widths = {};
    tableColumns.forEach(c => { widths[c.id] = c.width; });

    autoTable(doc, {
      startY: 135,
      columns: cols,
      body: tableRows,
      theme: 'grid',
      styles: {
        fontSize: 8.5,
        cellPadding: 6,
        valign: 'middle'
      },
      headStyles: {
        fillColor: [60, 60, 60],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        halign: 'center'
      },
      columnStyles: {
        slNo: { halign: 'center', cellWidth: widths.slNo },
        date: { cellWidth: widths.date },
        course: { cellWidth: widths.course },
        nc: { halign: 'center', cellWidth: widths.nc },
        qp: { cellWidth: widths.qp },
        lp: { cellWidth: widths.lp }
      },
      didDrawPage: (data) => {
        // Draw header on subsequent pages
        if (data.pageNumber > 1) {
          drawPageHeaders(doc, 'Date-wise Consolidated Printing Summary (Continued)');
        }
      }
    });

    doc.save(`${eventNameValPrefix()}_Consolidated_QP_Statement.pdf`);
  };

  // Generate and Download single Venue Packing Slip PDF
  const downloadVenuePDF = (venueName) => {
    const slipData = venueSlips[venueName];
    if (!slipData) return;

    const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
    drawPageHeaders(doc, `QP Envelope Packing Slip - Center: ${venueName}`);

    const tableRows = slipData.map((row, idx) => [
      idx + 1,
      row.date,
      `${row.courseCode}\n${row.courseName}`,
      row.studentCount,
      '', // Blank QP
      ''  // Blank LP
    ]);

    const cols = tableColumns.map(c => ({ header: c.label, dataKey: c.id }));
    const widths = {};
    tableColumns.forEach(c => { widths[c.id] = c.width; });

    autoTable(doc, {
      startY: 135,
      columns: cols,
      body: tableRows,
      theme: 'grid',
      styles: {
        fontSize: 8.5,
        cellPadding: 6,
        valign: 'middle'
      },
      headStyles: {
        fillColor: [60, 60, 60],
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        halign: 'center'
      },
      columnStyles: {
        slNo: { halign: 'center', cellWidth: widths.slNo },
        date: { cellWidth: widths.date },
        course: { cellWidth: widths.course },
        nc: { halign: 'center', cellWidth: widths.nc },
        qp: { cellWidth: widths.qp },
        lp: { cellWidth: widths.lp }
      }
    });

    const safeVenue = venueName.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
    doc.save(`QP_Packing_Slip_${safeVenue}.pdf`);
  };

  // Download all Venue Packing Slips compressed in a single ZIP
  const downloadAllVenueSlipsZip = async () => {
    const venues = Object.keys(venueSlips);
    if (venues.length === 0) return;

    setStatus('Creating ZIP archive...', 'normal');
    const zip = new JSZip();

    try {
      venues.forEach((venue) => {
        const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });
        drawPageHeaders(doc, `QP Envelope Packing Slip - Center: ${venue}`);

        const slipData = venueSlips[venue];
        const tableRows = slipData.map((row, idx) => [
          idx + 1,
          row.date,
          `${row.courseCode}\n${row.courseName}`,
          row.studentCount,
          '',
          ''
        ]);

        const cols = tableColumns.map(c => ({ header: c.label, dataKey: c.id }));
        const widths = {};
        tableColumns.forEach(c => { widths[c.id] = c.width; });

        autoTable(doc, {
          startY: 135,
          columns: cols,
          body: tableRows,
          theme: 'grid',
          styles: {
            fontSize: 8.5,
            cellPadding: 6,
            valign: 'middle'
          },
          headStyles: {
            fillColor: [60, 60, 60],
            textColor: [255, 255, 255],
            fontStyle: 'bold',
            halign: 'center'
          },
          columnStyles: {
            slNo: { halign: 'center', cellWidth: widths.slNo },
            date: { cellWidth: widths.date },
            course: { cellWidth: widths.course },
            nc: { halign: 'center', cellWidth: widths.nc },
            qp: { cellWidth: widths.qp },
            lp: { cellWidth: widths.lp }
          }
        });

        const pdfBlob = doc.output('blob');
        const safeVenue = venue.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
        zip.file(`QP_Packing_Slip_${safeVenue}.pdf`, pdfBlob);
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const downloadUrl = URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `QP_Packing_Slips_Collection.zip`;
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
    return eventNameVal.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_").substring(0, 35);
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

      <h2>QP Statement Report</h2>
      <p className="subtitle">Group, merge, and generate unified date-wise printing summaries and separate center envelope packing slips.</p>

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
          <strong>📌 Instructions:</strong> Please upload the <strong>Course-wise Venue-wise Date-wise Report</strong> Excel sheet(s).
          You can upload <strong>multiple reports at once</strong> (by holding Ctrl/Cmd during selection). The generator will automatically merge them into a single consolidated master list before compiling the statements.
        </div>
      </div>

      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(450px, 1fr))', gap: '32px', marginBottom: '32px' }}>
        {/* Settings Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}><Settings size={20} /> Settings</h3>
          
          <div style={{ border: '1px dashed var(--line)', padding: '24px', borderRadius: '8px', background: 'var(--bg)', textAlign: 'center', position: 'relative' }}>
            <strong style={{ display: 'block', marginBottom: '8px' }}>Upload Reports (Supports Multiple)</strong>
            {selectedFiles.length > 0 ? (
              <div style={{ color: 'var(--accent)', fontWeight: 600 }}>
                {selectedFiles.length === 1 ? `📄 ${selectedFiles[0].name}` : `📂 ${selectedFiles.length} files selected`}
              </div>
            ) : (
              <span style={{ color: 'var(--muted)', fontSize: '13px' }}>Drag here or click to browse multiple files</span>
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
              <label>Report Compilation Mode</label>
              <select value={reportMode} onChange={(e) => setReportMode(e.target.value)}>
                <option value="venue">Separate slips per Center/Venue</option>
                <option value="consolidated">Unified Date-wise Summary</option>
              </select>
            </div>
          </div>

          {/* Mappings */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)' }}>
            <strong style={{ fontSize: '13px', display: 'block', marginBottom: '12px' }}>Excel Column Mappings</strong>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', fontSize: '12px' }}>
              <div className="form-group">
                <label>Event Name Column</label>
                <select value={columnMapping.eventName} onChange={(e) => setColumnMapping({ ...columnMapping, eventName: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Center Code</label>
                <select value={columnMapping.centerCode} onChange={(e) => setColumnMapping({ ...columnMapping, centerCode: parseInt(e.target.value) })} disabled={headers.length === 0}>
                  {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i+1})</option>)}
                </select>
              </div>
              <div className="form-group">
                <label>Center Name</label>
                <select value={columnMapping.centerName} onChange={(e) => setColumnMapping({ ...columnMapping, centerName: parseInt(e.target.value) })} disabled={headers.length === 0}>
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
                <label>Program Name</label>
                <select value={columnMapping.programName} onChange={(e) => setColumnMapping({ ...columnMapping, programName: parseInt(e.target.value) })} disabled={headers.length === 0}>
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
                <label>Student Count (NC)</label>
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

          {/* Table Customizer */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)', display: 'flex', flexDirection: 'column', gap: '14px' }}>
            <strong style={{ fontSize: '13px' }}>Table Column Width Customizer (A4 printable width is 515pt)</strong>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              {tableColumns.map((col, idx) => (
                <div key={col.id} style={{ display: 'flex', gap: '12px', alignItems: 'center', fontSize: '11px' }}>
                  <span style={{ fontWeight: 600, width: '60px' }}>{col.label}</span>
                  <div style={{ display: 'flex', gap: '4px', alignItems: 'center' }}>
                    <span>Width:</span>
                    <input type="number" value={col.width} onChange={(e) => updateTableCol(idx, 'width', parseInt(e.target.value) || 20)} style={{ width: '45px', padding: '3px 4px' }} />
                  </div>
                  <div style={{ display: 'flex', gap: '4px', alignItems: 'center' }}>
                    <span>Font:</span>
                    <input type="number" value={col.fontSize} step="0.5" onChange={(e) => updateTableCol(idx, 'fontSize', parseFloat(e.target.value) || 9)} style={{ width: '45px', padding: '3px 4px' }} />
                  </div>
                  <div style={{ display: 'flex', gap: '4px', alignItems: 'center' }}>
                    <span>Align:</span>
                    <select value={col.align} onChange={(e) => updateTableCol(idx, 'align', e.target.value)} style={{ padding: '3px 4px', fontSize: '11px' }}>
                      <option value="left">Left</option>
                      <option value="center">Center</option>
                      <option value="right">Right</option>
                    </select>
                  </div>
                </div>
              ))}
              <div style={{
                fontSize: '11px',
                padding: '8px 12px',
                borderRadius: '4px',
                background: tableColumns.reduce((s, c) => s + c.width, 0) === 515 ? 'var(--accent-soft)' : 'var(--panel)',
                color: tableColumns.reduce((s, c) => s + c.width, 0) === 515 ? 'var(--accent)' : 'var(--danger)',
                fontWeight: 600
              }}>
                Current sum: {tableColumns.reduce((s, c) => s + c.width, 0)} pt / Target: 515 pt
              </div>
            </div>
          </div>

          <button onClick={processQpStatement} disabled={isProcessing || consolidatedRows.length === 0} style={{ width: '100%', padding: '14px', fontSize: '15px' }}>
            {isProcessing ? "Processing Data..." : "Generate QP Statement"}
          </button>
        </div>

        {/* Live Preview Panel */}
        <div className="card" style={{ padding: '32px', display: 'flex', flexDirection: 'column', gap: '20px', margin: 0 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}><Eye size={20} /> Live Preview</h3>
            
            {reportMode === 'consolidated' && consolidatedData.length > 0 && (
              <button onClick={downloadConsolidatedPDF} style={{ padding: '8px 16px', fontSize: '13px' }}>
                Download Consolidated PDF
              </button>
            )}

            {reportMode === 'venue' && Object.keys(venueSlips).length > 0 && (
              <div style={{ display: 'flex', gap: '8px' }}>
                <button onClick={() => downloadVenuePDF(activeVenueTab)} style={{ padding: '8px 16px', fontSize: '13px' }}>
                  Download Slip PDF
                </button>
                <button onClick={downloadAllVenueSlipsZip} className="secondary" style={{ padding: '8px 16px', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px', borderColor: 'var(--line)', color: 'var(--ink)' }}>
                  Download All ZIP
                </button>
              </div>
            )}
          </div>

          {/* Consolidated Mode Preview */}
          {reportMode === 'consolidated' && (
            consolidatedData.length > 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', flex: 1 }}>
                <div style={{
                  padding: '24px',
                  border: '1px solid var(--line)',
                  borderRadius: '8px',
                  background: 'white',
                  color: 'black',
                  fontFamily: 'Helvetica, Arial, sans-serif',
                  boxShadow: '0 4px 10px rgba(0,0,0,0.05)',
                  overflowX: 'auto'
                }}>
                  {/* Header Preview */}
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', marginBottom: '20px', borderBottom: '1px solid #ccc', paddingBottom: '12px' }}>
                    {headerLines.map((line, i) => (
                      <span key={i} style={{ 
                        fontSize: `${line.fontSize}px`, 
                        fontWeight: line.isBold ? 'bold' : 'normal',
                        textAlign: line.align,
                        width: '100%',
                        display: 'block',
                        marginBottom: '4px'
                      }}>{line.text}</span>
                    ))}
                    <strong style={{ fontSize: '11px', width: '100%', textAlign: 'left', marginTop: '10px' }}>Date-wise Consolidated Printing Summary</strong>
                  </div>

                  {/* Table Preview */}
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '9px' }}>
                    <thead>
                      <tr style={{ background: '#3c3c3c', color: 'white', fontWeight: 'bold' }}>
                        {tableColumns.map(col => (
                          <th key={col.id} style={{ border: '1px solid #ccc', padding: '6px', width: `${col.width}px`, textAlign: col.align }}>{col.label}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {consolidatedData.map((row, idx) => (
                        <tr key={idx}>
                          <td style={{ border: '1px solid #ccc', padding: '6px', textAlign: 'center' }}>{idx + 1}</td>
                          <td style={{ border: '1px solid #ccc', padding: '6px' }}>{row.date}</td>
                          <td style={{ border: '1px solid #ccc', padding: '6px', whiteSpace: 'pre-line' }}><strong>{row.courseCode}</strong><br />{row.courseName}</td>
                          <td style={{ border: '1px solid #ccc', padding: '6px', textAlign: 'center' }}>{row.totalQPs}</td>
                          <td style={{ border: '1px solid #ccc', padding: '6px' }}></td>
                          <td style={{ border: '1px solid #ccc', padding: '6px' }}></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', border: '1px dashed var(--line)', borderRadius: '8px', padding: '60px', color: 'var(--muted)', fontSize: '14px', textAlign: 'center' }}>
                Upload nominal roll reports, map columns, and click "Generate QP Statement" to render summary table preview.
              </div>
            )
          )}

          {/* Venue Mode Preview */}
          {reportMode === 'venue' && (
            Object.keys(venueSlips).length > 0 ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', flex: 1 }}>
                {/* Tabs */}
                <div style={{ display: 'flex', gap: '8px', overflowX: 'auto', paddingBottom: '8px', borderBottom: '1px solid var(--line)' }}>
                  {Object.keys(venueSlips).map((venue) => (
                    <button
                      key={venue}
                      onClick={() => setActiveVenueTab(venue)}
                      className={activeVenueTab === venue ? 'primary' : 'secondary'}
                      style={{ padding: '6px 12px', fontSize: '11px', whiteSpace: 'nowrap', borderColor: activeVenueTab === venue ? 'var(--accent)' : 'var(--line)', background: activeVenueTab === venue ? 'var(--accent)' : 'transparent', color: activeVenueTab === venue ? 'white' : 'var(--ink)' }}
                    >
                      {venue.split('-')[0].trim()}
                    </button>
                  ))}
                </div>

                <div style={{
                  padding: '24px',
                  border: '1px solid var(--line)',
                  borderRadius: '8px',
                  background: 'white',
                  color: 'black',
                  fontFamily: 'Helvetica, Arial, sans-serif',
                  boxShadow: '0 4px 10px rgba(0,0,0,0.05)',
                  overflowX: 'auto'
                }}>
                  {/* Header Preview */}
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', marginBottom: '20px', borderBottom: '1px solid #ccc', paddingBottom: '12px' }}>
                    {headerLines.map((line, i) => (
                      <span key={i} style={{ 
                        fontSize: `${line.fontSize}px`, 
                        fontWeight: line.isBold ? 'bold' : 'normal',
                        textAlign: line.align,
                        width: '100%',
                        display: 'block',
                        marginBottom: '4px'
                      }}>{line.text}</span>
                    ))}
                    <strong style={{ fontSize: '11px', width: '100%', textAlign: 'left', marginTop: '10px' }}>QP Envelope Packing Slip - Center: {activeVenueTab}</strong>
                  </div>

                  {/* Table Preview */}
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '9px' }}>
                    <thead>
                      <tr style={{ background: '#3c3c3c', color: 'white', fontWeight: 'bold' }}>
                        {tableColumns.map(col => (
                          <th key={col.id} style={{ border: '1px solid #ccc', padding: '6px', width: `${col.width}px`, textAlign: col.align }}>{col.label}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {venueSlips[activeVenueTab] && venueSlips[activeVenueTab].map((row, idx) => (
                        <tr key={idx}>
                          <td style={{ border: '1px solid #ccc', padding: '6px', textAlign: 'center' }}>{idx + 1}</td>
                          <td style={{ border: '1px solid #ccc', padding: '6px' }}>{row.date}</td>
                          <td style={{ border: '1px solid #ccc', padding: '6px', whiteSpace: 'pre-line' }}><strong>{row.courseCode}</strong><br />{row.courseName}</td>
                          <td style={{ border: '1px solid #ccc', padding: '6px', textAlign: 'center' }}>{row.studentCount}</td>
                          <td style={{ border: '1px solid #ccc', padding: '6px' }}></td>
                          <td style={{ border: '1px solid #ccc', padding: '6px' }}></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', flex: 1, alignItems: 'center', justifyContent: 'center', border: '1px dashed var(--line)', borderRadius: '8px', padding: '60px', color: 'var(--muted)', fontSize: '14px', textAlign: 'center' }}>
                Upload nominal roll reports, map columns, and click "Generate QP Statement" to render separate venue packing slips.
              </div>
            )
          )}
        </div>
      </div>
    </div>
  );
};

export default QpStatementPage;
