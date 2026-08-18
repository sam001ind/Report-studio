import { useState, useEffect } from 'react';
import { supabase } from '../supabaseClient';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import JSZip from 'jszip';
import { FileText, Archive, Printer, Sparkles } from 'lucide-react';

const hexToRgb = (hex) => {
  let c = String(hex || '#000000').replace(/^#/, '');
  if (c.length === 3) c = c.split('').map(x => x + x).join('');
  const num = parseInt(c, 16);
  if (isNaN(num)) return [0, 0, 0];
  return [(num >> 16) & 255, (num >> 8) & 255, num & 255];
};

const GeneratePage = ({ dataset = { columns: [], rows: [] } }) => {
  const [templates, setTemplates] = useState([]);
  const [selectedTemplateId, setSelectedTemplateId] = useState('');
  const [generatedPages, setGeneratedPages] = useState([]);
  const [isGenerating, setIsGenerating] = useState(false);
  const [zoomScale, setZoomScale] = useState(0.75);

  useEffect(() => {
    const fetchTemplates = async () => {
      const { data, error } = await supabase.from('templates').select('*');
      if (error) {
        console.error('Error fetching templates:', error);
      } else if (data && data.length > 0) {
        const mapped = data.map(row => ({ id: row.id, name: row.name, ...row.layout_data }));
        setTemplates(mapped);
        setSelectedTemplateId(mapped[mapped.length - 1].id);
      }
    };
    fetchTemplates();
  }, []);

  const handleGenerate = () => {
    const template = templates.find(t => t.id === selectedTemplateId) || templates[0] || {
      archetype: 'NOMINAL_ROLL',
      name: 'Default Nominal Roll Template',
      config: {
        pageSize: 'A4',
        orientation: 'portrait',
        headersList: [
          { id: 'h1', text: 'Kannur University', size: 15, bold: true, align: 'center', color: '#000000' },
          { id: 'h2', text: '(Examination Branch)', size: 11, bold: true, align: 'center', color: '#111111' },
          { id: 'h3', text: 'IV Semester Private Registration 2024 -2027 Admission', size: 10, bold: true, align: 'center', color: '#111111' },
          { id: 'h4', text: 'April 2026', size: 9.5, bold: false, align: 'center', color: '#222222' }
        ],
        tableTheme: { headerBg: '#f1f5f9', headerColor: '#000000', fontSize: 8, borderColor: '#64748b' }
      }
    };

    setIsGenerating(true);

    setTimeout(() => {
      const rows = dataset.rows && dataset.rows.length > 0 ? dataset.rows : [
        {
          'Programme': 'Bachelor of Business Administration (BBA)',
          'Venue': 'GA - Sree Narayana Guru College of Advanced Studies, Thottada',
          'Seat No': '4PR24BB001',
          'Candidate Name': 'ADITHYA K',
          'Course Code': 'KU4VACBBA200',
          'Course Title': 'Disaster Management'
        },
        {
          'Programme': 'Bachelor of Business Administration (BBA)',
          'Venue': 'GA - Sree Narayana Guru College of Advanced Studies, Thottada',
          'Seat No': '4PR24BB001',
          'Candidate Name': 'ADITHYA K',
          'Course Code': 'KU4SECBBA201',
          'Course Title': 'Soft Skills & Personality Development'
        },
        {
          'Programme': 'Bachelor of Business Administration (BBA)',
          'Venue': 'GA - Sree Narayana Guru College of Advanced Studies, Thottada',
          'Seat No': '4PR24BB002',
          'Candidate Name': 'ANANYA RAJEEV',
          'Course Code': 'KU4VACBBA200',
          'Course Title': 'Disaster Management'
        }
      ];

      const config = template.config || {};
      const arch = template.archetype || 'NOMINAL_ROLL';

      // Group rows for Nominal Roll
      const groups = {};
      rows.forEach(r => {
        const prog = r['Programme'] || r['Program'] || r['programme'] || 'General Programme';
        const venue = r['Venue'] || r['Venue Name'] || r['venue'] || 'GA - Sree Narayana Guru College, Thottada';
        const seat = r['Seat No'] || r['Register No'] || r['seatNo'] || r['regNo'] || '4PR24BB001';
        const name = r['Candidate Name'] || r['Name'] || r['studentName'] || 'Candidate';
        const cCode = r['Course Code'] || r['Paper Code'] || r['courseCode'] || '';
        const cTitle = r['Course Title'] || r['Paper Name'] || r['courseTitle'] || 'Subject';

        const gKey = `${prog} • ${venue}`;
        if (!groups[gKey]) {
          groups[gKey] = {
            programme: prog,
            venueLabel: venue,
            candidatesMap: {}
          };
        }

        if (!groups[gKey].candidatesMap[seat]) {
          groups[gKey].candidatesMap[seat] = {
            seatNo: seat,
            studentName: name,
            courses: []
          };
        }

        const crsDisplay = cCode && cTitle ? `${cCode} - ${cTitle}` : (cCode || cTitle);
        if (!groups[gKey].candidatesMap[seat].courses.some(c => c.display === crsDisplay)) {
          groups[gKey].candidatesMap[seat].courses.push({ display: crsDisplay });
        }
      });

      const pages = Object.keys(groups).map((gKey, idx) => ({
        id: `page_${idx}`,
        template,
        config,
        archetype: arch,
        group: {
          ...groups[gKey],
          candidates: Object.values(groups[gKey].candidatesMap)
        }
      }));

      setGeneratedPages(pages);
      setIsGenerating(false);
    }, 100);
  };

  const handleDownloadPdf = () => {
    if (generatedPages.length === 0) return alert('Please click Generate Output first.');
    const activePage = generatedPages[0];
    const isLandscape = activePage.config?.orientation === 'landscape';
    const doc = new jsPDF(isLandscape ? 'l' : 'p', 'mm', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();

    generatedPages.forEach((p, pIdx) => {
      if (pIdx > 0) doc.addPage();
      let startY = 14;

      // Draw Headers
      (p.config?.headersList || []).forEach(h => {
        doc.setFont('helvetica', h.bold ? 'bold' : 'normal');
        doc.setFontSize(h.size || 11);
        const [r, g, b] = hexToRgb(h.color || '#000000');
        doc.setTextColor(r, g, b);
        doc.text(h.text, pageWidth / 2, startY, { align: 'center' });
        startY += (h.size || 11) * 0.42 + 2;
      });

      // Group Box
      startY += 2;
      const boxW = pageWidth - 28;
      doc.setDrawColor(180, 180, 180);
      doc.setFillColor(248, 249, 250);
      doc.roundedRect(14, startY, boxW, 14, 2, 2, 'FD');

      doc.setFontSize(9);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(20, 20, 20);
      doc.text(`Programme: ${p.group.programme}`, 18, startY + 5);
      doc.setTextColor(23, 107, 135);
      doc.text(`Venue: ${p.group.venueLabel}`, 18, startY + 10);

      // Table
      const tableBody = [];
      p.group.candidates.forEach((cand, cIdx) => {
        const cCount = Math.max(1, cand.courses.length);
        cand.courses.forEach((crs, crsIdx) => {
          if (crsIdx === 0) {
            tableBody.push([
              { content: String(cIdx + 1), rowSpan: cCount, styles: { halign: 'center', valign: 'middle' } },
              { content: cand.seatNo, rowSpan: cCount, styles: { halign: 'center', valign: 'middle', fontStyle: 'bold' } },
              { content: cand.studentName, rowSpan: cCount, styles: { halign: 'left', valign: 'middle', fontStyle: 'bold' } },
              { content: crs.display, styles: { halign: 'left', valign: 'middle' } },
              { content: '', rowSpan: cCount, styles: { halign: 'center', valign: 'middle' } }
            ]);
          } else {
            tableBody.push([{ content: crs.display, styles: { halign: 'left', valign: 'middle' } }]);
          }
        });
      });

      autoTable(doc, {
        startY: startY + 18,
        head: [['Sl No', 'Register Number', 'Candidate Name', 'Courses', 'Remarks']],
        body: tableBody,
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: 8, cellPadding: 2, textColor: [0, 0, 0] },
        headStyles: { fillColor: [241, 245, 249], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center' }
      });
    });

    doc.save('Consolidated_Report_Output.pdf');
  };

  const handleDownloadZip = async () => {
    if (generatedPages.length === 0) return alert('Please click Generate Output first.');
    const zip = new JSZip();
    
    generatedPages.forEach((p, pIdx) => {
      const isLandscape = p.config?.orientation === 'landscape';
      const doc = new jsPDF(isLandscape ? 'l' : 'p', 'mm', 'a4');
      const pageWidth = doc.internal.pageSize.getWidth();

      let startY = 14;
      (p.config?.headersList || []).forEach(h => {
        doc.setFont('helvetica', h.bold ? 'bold' : 'normal');
        doc.setFontSize(h.size || 11);
        const [r, g, b] = hexToRgb(h.color || '#000000');
        doc.setTextColor(r, g, b);
        doc.text(h.text, pageWidth / 2, startY, { align: 'center' });
        startY += (h.size || 11) * 0.42 + 2;
      });

      const boxW = pageWidth - 28;
      doc.setDrawColor(180, 180, 180);
      doc.setFillColor(248, 249, 250);
      doc.roundedRect(14, startY, boxW, 14, 2, 2, 'FD');

      doc.setFontSize(9);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(20, 20, 20);
      doc.text(`Programme: ${p.group.programme}`, 18, startY + 5);
      doc.setTextColor(23, 107, 135);
      doc.text(`Venue: ${p.group.venueLabel}`, 18, startY + 10);

      const tableBody = [];
      p.group.candidates.forEach((cand, cIdx) => {
        const cCount = Math.max(1, cand.courses.length);
        cand.courses.forEach((crs, crsIdx) => {
          if (crsIdx === 0) {
            tableBody.push([
              { content: String(cIdx + 1), rowSpan: cCount, styles: { halign: 'center', valign: 'middle' } },
              { content: cand.seatNo, rowSpan: cCount, styles: { halign: 'center', valign: 'middle', fontStyle: 'bold' } },
              { content: cand.studentName, rowSpan: cCount, styles: { halign: 'left', valign: 'middle', fontStyle: 'bold' } },
              { content: crs.display, styles: { halign: 'left', valign: 'middle' } },
              { content: '', rowSpan: cCount, styles: { halign: 'center', valign: 'middle' } }
            ]);
          } else {
            tableBody.push([{ content: crs.display, styles: { halign: 'left', valign: 'middle' } }]);
          }
        });
      });

      autoTable(doc, {
        startY: startY + 18,
        head: [['Sl No', 'Register Number', 'Candidate Name', 'Courses', 'Remarks']],
        body: tableBody,
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: 8, cellPadding: 2, textColor: [0, 0, 0] },
        headStyles: { fillColor: [241, 245, 249], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center' }
      });

      const pdfBlob = doc.output('blob');
      const safeVenue = p.group.venueLabel.replace(/[/\\?%*:|"<>•]/g, '_').slice(0, 40);
      zip.file(`Individual_PDFs/Page_${pIdx + 1}_${safeVenue}.pdf`, pdfBlob);
    });

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Individual_Reports_Package.zip';
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div style={{ padding: '32px', maxWidth: '1400px', margin: '0 auto', display: 'flex', flexDirection: 'column', height: '100%' }}>
      
      <div className="no-print">
        <h2 style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <Sparkles size={22} color="var(--accent)" /> Generate Reports & Master PDF
        </h2>
        <p className="subtitle">Produce exact matching reports with live multi-page preview, master PDF compiling, and ZIP packaging.</p>

        <div className="card" style={{ display: 'flex', gap: '14px', alignItems: 'flex-end', background: 'white', padding: '20px', borderRadius: '12px', flexWrap: 'wrap', marginBottom: '24px' }}>
          <div className="form-group" style={{ minWidth: '180px' }}>
            <label style={{ fontSize: '12px', fontWeight: 700 }}>1. Active Dataset</label>
            <input type="text" disabled value={`${dataset.rows?.length || 0} rows loaded`} style={{ background: '#f8fafc', color: 'var(--ink)', fontWeight: 600 }} />
          </div>
          
          <div className="form-group" style={{ minWidth: '240px', flex: 1 }}>
            <label style={{ fontSize: '12px', fontWeight: 700 }}>2. Report Template</label>
            <select value={selectedTemplateId} onChange={e => setSelectedTemplateId(e.target.value)} style={{ width: '100%' }}>
              <option value="">-- Select Template --</option>
              {templates.map(t => <option key={t.id} value={t.id}>{t.name} ({t.archetype || 'Custom'})</option>)}
            </select>
          </div>

          <div className="form-group" style={{ width: '120px' }}>
            <label style={{ fontSize: '12px', fontWeight: 700 }}>3. Zoom</label>
            <select value={zoomScale} onChange={e => setZoomScale(parseFloat(e.target.value))}>
              <option value="1">100%</option>
              <option value="0.75">75%</option>
              <option value="0.5">50%</option>
            </select>
          </div>

          <button className="button" onClick={handleGenerate} style={{ padding: '10px 20px', fontSize: '13px' }}>
            {isGenerating ? 'Generating...' : '⚡ Generate Output'}
          </button>
          
          {generatedPages.length > 0 && (
            <div style={{ display: 'flex', gap: '8px', flexWrap: 'wrap' }}>
              <button className="button secondary" onClick={handleDownloadPdf} style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '8px 14px' }}>
                <FileText size={15} /> Download Master PDF
              </button>
              <button className="button" onClick={handleDownloadZip} style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '8px 16px' }}>
                <Archive size={15} /> Individual PDFs (.zip)
              </button>
              <button className="button secondary" onClick={() => window.print()} style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '8px 14px' }}>
                <Printer size={15} /> Print
              </button>
            </div>
          )}
        </div>
      </div>

      {generatedPages.length === 0 ? (
        <div className="no-print" style={{ background: '#f8fafc', padding: '60px 40px', borderRadius: '12px', border: '1px dashed #cbd5e1', textAlign: 'center', color: 'var(--muted)' }}>
          <Sparkles size={32} color="var(--accent)" style={{ margin: '0 auto 12px', opacity: 0.7 }} />
          <h3 style={{ margin: '0 0 6px 0', fontSize: '16px' }}>No report generated yet</h3>
          <p style={{ margin: 0, fontSize: '13px' }}>Click <strong>"Generate Output"</strong> above to produce and preview the exact reports.</p>
        </div>
      ) : (
        <div className="print-container" style={{ background: '#e2e8f0', padding: '32px', borderRadius: '12px', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '32px', overflowY: 'auto' }}>
          {generatedPages.map((page, idx) => (
            <div 
              key={page.id || idx} 
              className="preview-page" 
              style={{ 
                width: page.config?.orientation === 'landscape' ? '1123px' : '794px', 
                minHeight: '1000px',
                background: 'white',
                padding: page.config?.orientation === 'landscape' ? '32px 40px' : '40px 48px',
                position: 'relative',
                boxShadow: '0 10px 25px rgba(0,0,0,0.1)',
                border: '1px solid #cbd5e1',
                borderRadius: '6px',
                pageBreakAfter: 'always',
                transform: `scale(${zoomScale})`,
                transformOrigin: 'top center',
                marginBottom: `-${1000 * (1 - zoomScale) - 24}px`
              }}
            >
              {/* Dynamic Headers */}
              <div style={{ textAlign: 'center', marginBottom: '20px' }}>
                {(page.config?.headersList || []).map((h, hIdx) => (
                  <div 
                    key={h.id || hIdx}
                    style={{
                      fontSize: `${h.size || 11}pt`,
                      fontWeight: h.bold ? 800 : 400,
                      color: h.color || '#000000',
                      textAlign: h.align || 'center',
                      marginBottom: '3px'
                    }}
                  >
                    {h.text}
                  </div>
                ))}
              </div>

              {/* Program & Venue Info Box */}
              <div style={{ background: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: '6px', padding: '10px 14px', marginBottom: '16px' }}>
                <div style={{ fontSize: '13px', fontWeight: 800, color: '#0f172a', marginBottom: '4px' }}>
                  Programme: {page.group.programme}
                </div>
                <div style={{ fontSize: '13px', fontWeight: 800, color: '#0284c7' }}>
                  Venue: {page.group.venueLabel}
                </div>
              </div>

              {/* Candidate Table with merged sub-row courses */}
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '8.5pt' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', color: '#000', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000' }}>
                    <th style={{ border: '1px solid #64748b', padding: '8px 4px', width: '45px', textAlign: 'center' }}>Sl No</th>
                    <th style={{ border: '1px solid #64748b', padding: '8px 6px', width: '140px', textAlign: 'center' }}>Register Number</th>
                    <th style={{ border: '1px solid #64748b', padding: '8px 10px', width: '180px', textAlign: 'left' }}>Candidate Name</th>
                    <th style={{ border: '1px solid #64748b', padding: '8px 10px', textAlign: 'left' }}>Courses</th>
                    <th style={{ border: '1px solid #64748b', padding: '8px 6px', width: '90px', textAlign: 'center' }}>Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  {page.group.candidates.map((cand, candIdx) => {
                    const courses = cand.courses || [];
                    return courses.map((crs, crsIdx) => (
                      <tr key={`${candIdx}_${crsIdx}`} style={{ borderBottom: '1px solid #64748b' }}>
                        {crsIdx === 0 && (
                          <>
                            <td rowSpan={courses.length} style={{ border: '1px solid #64748b', padding: '6px 4px', textAlign: 'center', verticalAlign: 'middle' }}>{candIdx + 1}</td>
                            <td rowSpan={courses.length} style={{ border: '1px solid #64748b', padding: '6px 8px', textAlign: 'center', verticalAlign: 'middle', fontWeight: 700 }}>{cand.seatNo}</td>
                            <td rowSpan={courses.length} style={{ border: '1px solid #64748b', padding: '6px 10px', verticalAlign: 'middle', fontWeight: 700 }}>{cand.studentName}</td>
                          </>
                        )}
                        <td style={{ border: '1px solid #64748b', padding: '5px 8px', verticalAlign: 'middle' }}>{crs.display}</td>
                        {crsIdx === 0 && (
                          <td rowSpan={courses.length} style={{ border: '1px solid #64748b', padding: '6px', textAlign: 'center', verticalAlign: 'middle' }}></td>
                        )}
                      </tr>
                    ));
                  })}
                </tbody>
              </table>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default GeneratePage;
