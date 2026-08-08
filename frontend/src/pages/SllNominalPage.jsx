import React, { useState, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import 'jspdf-autotable';
import JSZip from 'jszip';
import { useAuth } from '../context/AuthContext';
import { FileStack, Download, Eye, Settings, FileText } from 'lucide-react';

const SllNominalPage = () => {
  const { user } = useAuth();
  
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
    venue: 5,
    courseCode: 6,
    courseTitle: 7
  });

  // Custom Header Text Configuration
  const [headerConfig, setHeaderConfig] = useState({
    title: 'Kannur University',
    subtitle: 'Examination Branch',
    session: 'I Semester Private Registration 2025 -2028 Admission - November 2025'
  });

  // Processing states
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Generated Venue Data State
  const [venueGroups, setVenueGroups] = useState({});
  const [activeVenueTab, setActiveVenueTab] = useState('');
  const [programmeName, setProgrammeName] = useState('');

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  // Extract headers and programme name on sheet change
  useEffect(() => {
    if (workbook && selectedSheet) {
      const sheet = workbook.Sheets[selectedSheet];
      if (sheet) {
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
        if (rows.length > 0) {
          // Store raw headers for mapping dropdowns
          const firstRow = rows[0].map((cell, idx) => cell ? String(cell).trim() : `Column ${idx + 1}`);
          setHeaders(firstRow);

          // Get default programme name from A2 (index 0, row 1)
          if (rows.length > 1 && rows[1][0]) {
            setProgrammeName(String(rows[1][0]).trim());
          } else {
            setProgrammeName('Private Registration');
          }

          // Try to auto-detect columns
          const autoMap = { ...columnMapping };
          firstRow.forEach((name, idx) => {
            const lower = name.toLowerCase();
            if (lower.includes('programme') || lower.includes('program')) autoMap.programme = idx;
            if (lower.includes('name')) autoMap.name = idx;
            if (lower.includes('seat') || lower.includes('reg') || lower.includes('register') || lower.includes('roll')) autoMap.seatNo = idx;
            if (lower.includes('venue') || lower.includes('center') || lower.includes('college')) autoMap.venue = idx;
            if (lower.includes('coursecode') || lower.includes('subjectcode') || (lower.includes('course') && !lower.includes('title') && !lower.includes('name'))) autoMap.courseCode = idx;
            if (lower.includes('coursetitle') || lower.includes('subjectname') || lower.includes('coursename') || lower.includes('title')) autoMap.courseTitle = idx;
          });
          setColumnMapping(autoMap);
        }
      }
    } else {
      setHeaders([]);
      setProgrammeName('');
    }
  }, [workbook, selectedSheet]);

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setSelectedFile(file);
    setStatus('Reading workbook...');
    setVenueGroups({});
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

  // Group data by Venue and Student
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

        const venue = String(row[columnMapping.venue] || 'Unassigned Venue').trim();
        const seatNo = String(row[columnMapping.seatNo] || '').trim();
        const name = String(row[columnMapping.name] || '').trim();
        const cCode = String(row[columnMapping.courseCode] || '').trim();
        const cTitle = String(row[columnMapping.courseTitle] || '').trim();

        if (!venue || !seatNo) return; // Skip if no venue or register number

        if (!groups[venue]) {
          groups[venue] = {};
        }

        // Group by seatNo (register number) to perform row spanning
        const studentKey = `${seatNo}_${name}`;
        if (!groups[venue][studentKey]) {
          groups[venue][studentKey] = {
            seatNo,
            name,
            courses: []
          };
        }

        groups[venue][studentKey].courses.push({
          code: cCode,
          title: cTitle
        });
      });

      // Convert inner dictionary to sorted arrays of students
      const processedGroups = {};
      Object.keys(groups).forEach(venue => {
        const studentList = Object.values(groups[venue]);
        // Sort students by Seat No
        studentList.sort((a, b) => a.seatNo.localeCompare(b.seatNo, undefined, { numeric: true, sensitivity: 'base' }));
        processedGroups[venue] = studentList;
      });

      setVenueGroups(processedGroups);
      const venues = Object.keys(processedGroups);
      if (venues.length > 0) {
        setActiveVenueTab(venues[0]);
      }
      
      setStatus(`Nominal rolls processed! Found ${venues.length} venues.`, 'success');
    } catch (err) {
      setStatus(`Error processing data: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Generate jsPDF instance for a single venue
  const generateVenuePDF = (venueName, students) => {
    const doc = new jsPDF({
      orientation: 'portrait',
      unit: 'pt',
      format: 'a4'
    });

    // 1. Header Block
    doc.setFont('Helvetica', 'bold');
    doc.setFontSize(16);
    doc.text(headerConfig.title, 297, 40, { align: 'center' });

    doc.setFontSize(12);
    doc.text(headerConfig.subtitle, 297, 60, { align: 'center' });

    doc.setFontSize(10);
    doc.text(headerConfig.session, 297, 78, { align: 'center' });

    // 2. Metadata Info
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
          row.push({ content: slNo.toString(), rowSpan: courseCount, styles: { valign: 'middle', halign: 'center' } });
          row.push({ content: student.seatNo, rowSpan: courseCount, styles: { valign: 'middle', halign: 'center' } });
          row.push({ content: student.name, rowSpan: courseCount, styles: { valign: 'middle' } });
        }
        
        // Course detail column (Course Code - Course Title)
        const courseStr = course.code ? `${course.code} - ${course.title}` : course.title;
        row.push({ content: courseStr, styles: { valign: 'top' } });
        
        // Remark column
        if (index === 0) {
          row.push({ content: '', rowSpan: courseCount });
        }

        tableBody.push(row);
      });
      slNo += 1;
    });

    // 4. Generate Autotable
    doc.autoTable({
      startY: 145,
      head: [['Sl No', 'Seat No', 'Name', 'Courses', 'Remark']],
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
      columnStyles: {
        0: { cellWidth: 35 },  // Sl No
        1: { cellWidth: 70 },  // Seat No
        2: { cellWidth: 140 }, // Name
        3: { cellWidth: 210 }, // Course Code + Title
        4: { cellWidth: 60 }   // Remark
      },
      didParseCell: (data) => {
        // Match line spacing and styles
        if (data.section === 'body' && data.column.index === 3) {
          data.cell.styles.fontSize = 9.5;
        }
      }
    });

    return doc;
  };

  // Download single venue PDF
  const downloadSinglePDF = (venueName) => {
    const students = venueGroups[venueName];
    if (!students) return;
    const doc = generateVenuePDF(venueName, students);
    const safeName = venueName.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
    doc.save(`${programmeName}_${safeName}.pdf`);
  };

  // Batch download all venue PDFs in a single ZIP
  const downloadAllAsZip = async () => {
    const venues = Object.keys(venueGroups);
    if (venues.length === 0) return;

    setStatus('Creating ZIP archive...', 'normal');
    const zip = new JSZip();

    try {
      venues.forEach((venueName) => {
        const students = venueGroups[venueName];
        const doc = generateVenuePDF(venueName, students);
        
        // Output PDF to ArrayBuffer
        const pdfBlob = doc.output('blob');
        const safeName = venueName.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").replace(/\s+/g, "_");
        zip.file(`${programmeName}_${safeName}.pdf`, pdfBlob);
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const downloadUrl = URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `${programmeName}_Venue_Nominal_Rolls.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

      setStatus(`ZIP archive downloaded successfully!`, 'success');
    } catch (err) {
      setStatus(`Failed to generate ZIP: ${err.message}`, 'error');
    }
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
      <p className="subtitle">Compile venue-wise nominal roll lists from a master register. Rows for the same seat/register number will be automatically merged, just like in Kannur University VBA logs.</p>

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
              accept=".xlsx, .xls, .xlsm, .csv" 
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
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '16px' }}>
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
            <div className="form-group">
              <label>Programme Title</label>
              <input 
                type="text" 
                value={programmeName} 
                onChange={(e) => setProgrammeName(e.target.value)} 
                placeholder="A2 Programme Value"
                disabled={!workbook}
              />
            </div>
          </div>

          {/* Column Mappings */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)' }}>
            <strong style={{ fontSize: '13px', display: 'block', marginBottom: '12px' }}>Excel Column Mappings</strong>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', fontSize: '12px' }}>
              <div className="form-group">
                <label>Venue (F)</label>
                <select 
                  value={columnMapping.venue} 
                  onChange={(e) => setColumnMapping({ ...columnMapping, venue: parseInt(e.target.value) })}
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

          {/* Header Text Configurations */}
          <div style={{ border: '1px solid var(--line)', padding: '16px', borderRadius: '8px', background: 'var(--bg)' }}>
            <strong style={{ fontSize: '13px', display: 'block', marginBottom: '12px' }}>Report Headers</strong>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', fontSize: '12px' }}>
              <div className="form-group">
                <label>Header Line 1 (University)</label>
                <input 
                  type="text" 
                  value={headerConfig.title} 
                  onChange={(e) => setHeaderConfig({ ...headerConfig, title: e.target.value })} 
                />
              </div>
              <div className="form-group">
                <label>Header Line 2 (Branch)</label>
                <input 
                  type="text" 
                  value={headerConfig.subtitle} 
                  onChange={(e) => setHeaderConfig({ ...headerConfig, subtitle: e.target.value })} 
                />
              </div>
              <div className="form-group">
                <label>Header Line 3 (Session Detail)</label>
                <input 
                  type="text" 
                  value={headerConfig.session} 
                  onChange={(e) => setHeaderConfig({ ...headerConfig, session: e.target.value })} 
                />
              </div>
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
            {Object.keys(venueGroups).length > 0 && (
              <button 
                onClick={downloadAllAsZip} 
                className="secondary" 
                style={{ padding: '8px 16px', fontSize: '13px', display: 'flex', alignItems: 'center', gap: '6px', borderColor: 'var(--accent)', color: 'var(--accent)' }}
              >
                <Download size={16} /> Download All (ZIP)
              </button>
            )}
          </div>

          {Object.keys(venueGroups).length > 0 ? (
            <>
              {/* Tab Selector */}
              <div style={{ display: 'flex', gap: '8px', overflowX: 'auto', paddingBottom: '8px', borderBottom: '1px solid var(--line)' }}>
                {Object.keys(venueGroups).map((venueName) => (
                  <button
                    key={venueName}
                    onClick={() => setActiveVenueTab(venueName)}
                    style={{
                      padding: '6px 12px',
                      borderRadius: '6px',
                      border: '1px solid var(--line)',
                      background: activeVenueTab === venueName ? 'var(--accent)' : 'var(--panel)',
                      color: activeVenueTab === venueName ? 'white' : 'var(--ink)',
                      whiteSpace: 'nowrap',
                      fontSize: '12px',
                      cursor: 'pointer',
                      fontWeight: activeVenueTab === venueName ? 600 : 400
                    }}
                  >
                    {venueName} ({venueGroups[venueName].length} Studs)
                  </button>
                ))}
              </div>

              {/* Active Venue Layout Preview */}
              <div style={{ flex: 1, display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                  <div>
                    <h4 style={{ margin: 0, color: 'var(--accent)' }}>{activeVenueTab}</h4>
                    <span style={{ fontSize: '12px', color: 'var(--muted)' }}>Programme: {programmeName}</span>
                  </div>
                  <button 
                    onClick={() => downloadSinglePDF(activeVenueTab)} 
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
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)', textAlign: 'center', width: '40px' }}>Sl No</th>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)', textAlign: 'center', width: '80px' }}>Seat No</th>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)', width: '130px' }}>Name</th>
                        <th style={{ padding: '10px 8px', borderRight: '1px solid var(--line)' }}>Courses</th>
                        <th style={{ padding: '10px 8px', width: '70px' }}>Remark</th>
                      </tr>
                    </thead>
                    <tbody>
                      {venueGroups[activeVenueTab].map((student, sIdx) => {
                        return student.courses.map((course, cIdx) => {
                          const isFirst = cIdx === 0;
                          return (
                            <tr key={`${sIdx}-${cIdx}`} style={{ borderBottom: '1px solid var(--line)' }}>
                              {isFirst ? (
                                <>
                                  <td 
                                    rowSpan={student.courses.length} 
                                    style={{ padding: '8px', borderRight: '1px solid var(--line)', textAlign: 'center', verticalAlign: 'middle', background: 'rgba(255,255,255,0.01)' }}
                                  >
                                    {sIdx + 1}
                                  </td>
                                  <td 
                                    rowSpan={student.courses.length} 
                                    style={{ padding: '8px', borderRight: '1px solid var(--line)', textAlign: 'center', verticalAlign: 'middle', fontWeight: 600 }}
                                  >
                                    {student.seatNo}
                                  </td>
                                  <td 
                                    rowSpan={student.courses.length} 
                                    style={{ padding: '8px', borderRight: '1px solid var(--line)', verticalAlign: 'middle' }}
                                  >
                                    {student.name}
                                  </td>
                                </>
                              ) : null}
                              <td style={{ padding: '8px', borderRight: '1px solid var(--line)', color: 'var(--ink)' }}>
                                <span style={{ fontWeight: 600, color: 'var(--accent)' }}>{course.code}</span>
                                {course.code && course.title ? ' - ' : ''}
                                <span>{course.title}</span>
                              </td>
                              {isFirst ? (
                                <td 
                                  rowSpan={student.courses.length} 
                                  style={{ padding: '8px', verticalAlign: 'middle' }}
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
