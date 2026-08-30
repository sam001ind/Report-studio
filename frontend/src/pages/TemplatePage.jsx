import { useState, useEffect, useMemo, useRef } from 'react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { 
  FileText, 
  TableProperties, 
  Edit3, 
  Sliders, 
  Download, 
  Printer, 
  Archive, 
  Save, 
  Sparkles, 
  Layers, 
  ChevronLeft, 
  ChevronRight,
  Plus,
  Trash2,
  Image as ImageIcon,
  Bold,
  Italic,
  AlignLeft,
  AlignCenter,
  AlignRight,
  ArrowUp,
  ArrowDown,
  ArrowLeft,
  ArrowRight as ArrowRightIcon,
  FileSpreadsheet
} from 'lucide-react';
import { supabase } from '../supabaseClient';
import { logoBase64 } from '../assets/logoBase64';
import { DEFAULT_REPORT_TEMPLATE, autoDetectDatasetColumns } from '../utils/templateEngine';

const hexToRgb = (hex) => {
  let c = String(hex || '#000000').replace(/^#/, '');
  if (c.length === 3) c = c.split('').map(x => x + x).join('');
  const num = parseInt(c, 16);
  if (isNaN(num)) return [0, 0, 0];
  return [(num >> 16) & 255, (num >> 8) & 255, num & 255];
};

const TemplatePage = ({ dataset = { columns: [], rows: [] }, initialTemplate }) => {
  const fileInputRef = useRef(null);

  // Template Configuration
  const [templateConfig, setTemplateConfig] = useState(() => {
    if (initialTemplate?.config) {
      return JSON.parse(JSON.stringify(initialTemplate.config));
    }
    return JSON.parse(JSON.stringify(DEFAULT_REPORT_TEMPLATE));
  });

  const [columnMappings, setColumnMappings] = useState(() => autoDetectDatasetColumns(dataset?.columns || []));
  const [templateName, setTemplateName] = useState(initialTemplate?.name || 'Custom Tabular Report');
  const [isProcessing, setIsProcessing] = useState(false);
  const [activePreviewIndex, setActivePreviewIndex] = useState(0);

  // Active Tool Tab: 'columns' | 'headers' | 'page' | 'grouping' | 'styling'
  const [activeToolTab, setActiveToolTab] = useState('columns');

  // Auto-detect columns and optionally build initial tableColumns when dataset changes
  useEffect(() => {
    if (dataset?.columns?.length > 0) {
      const detected = autoDetectDatasetColumns(dataset.columns);
      setColumnMappings(detected);

      // If templateColumns is empty or only default, populate from dataset columns
      setTemplateConfig(prev => {
        if (!prev.tableColumns || prev.tableColumns.length === 0) {
          const autoCols = dataset.columns.slice(0, 8).map((colName, idx) => ({
            id: `col_${Date.now()}_${idx}`,
            label: colName,
            field: colName,
            width: 30,
            align: idx === 0 ? 'center' : 'left',
            bold: idx === 0
          }));
          return {
            ...prev,
            tableColumns: [
              { id: `col_slno_${Date.now()}`, label: 'Sl No', field: 'slNo', width: 12, align: 'center', bold: true },
              ...autoCols
            ]
          };
        }
        return prev;
      });
    }
  }, [dataset]);

  // Load from initialTemplate
  useEffect(() => {
    if (initialTemplate) {
      if (initialTemplate.config) {
        setTemplateConfig(initialTemplate.config);
      }
      if (initialTemplate.columnMappings) {
        setColumnMappings(initialTemplate.columnMappings);
      }
      if (initialTemplate.name) {
        setTemplateName(initialTemplate.name);
      }
    }
  }, [initialTemplate]);

  const updateConfig = (updates) => {
    setTemplateConfig(prev => ({
      ...prev,
      ...updates
    }));
  };

  // Top Page Headers Helpers
  const addHeaderLine = () => {
    const newHeader = {
      id: `h_${Date.now()}`,
      text: 'New Institutional Header Line',
      size: 11,
      bold: true,
      italic: false,
      align: 'center',
      color: '#000000',
      font: 'helvetica'
    };
    updateConfig({
      headersList: [...(templateConfig.headersList || []), newHeader]
    });
  };

  const updateHeaderLine = (id, updates) => {
    const list = (templateConfig.headersList || []).map(h => h.id === id ? { ...h, ...updates } : h);
    updateConfig({ headersList: list });
  };

  const removeHeaderLine = (id) => {
    const list = (templateConfig.headersList || []).filter(h => h.id !== id);
    updateConfig({ headersList: list });
  };

  const moveHeaderLine = (index, direction) => {
    const list = [...(templateConfig.headersList || [])];
    const targetIdx = index + direction;
    if (targetIdx < 0 || targetIdx >= list.length) return;
    const temp = list[index];
    list[index] = list[targetIdx];
    list[targetIdx] = temp;
    updateConfig({ headersList: list });
  };

  // Table Columns Helpers
  const addTableColumn = () => {
    const cols = templateConfig.tableColumns || [];
    const newCol = {
      id: `col_${Date.now()}`,
      label: `Column ${cols.length + 1}`,
      field: dataset?.columns?.[0] || 'blank',
      width: 30,
      align: 'left',
      bold: false
    };
    updateConfig({ tableColumns: [...cols, newCol] });
  };

  const importAllDatasetColumns = () => {
    if (!dataset?.columns?.length) return alert('Please upload an Excel spreadsheet in Report Config first.');
    const imported = [
      { id: `col_slno_${Date.now()}`, label: 'Sl No', field: 'slNo', width: 12, align: 'center', bold: true },
      ...dataset.columns.map((colName, idx) => ({
        id: `col_${Date.now()}_${idx}`,
        label: colName,
        field: colName,
        width: 30,
        align: 'left',
        bold: false
      }))
    ];
    updateConfig({ tableColumns: imported });
  };

  const updateTableColumn = (id, updates) => {
    const cols = (templateConfig.tableColumns || []).map(c => c.id === id ? { ...c, ...updates } : c);
    updateConfig({ tableColumns: cols });
  };

  const removeTableColumn = (id) => {
    const cols = (templateConfig.tableColumns || []).filter(c => c.id !== id);
    updateConfig({ tableColumns: cols });
  };

  const moveTableColumn = (index, direction) => {
    const cols = [...(templateConfig.tableColumns || [])];
    const targetIdx = index + direction;
    if (targetIdx < 0 || targetIdx >= cols.length) return;
    const temp = cols[index];
    cols[index] = cols[targetIdx];
    cols[targetIdx] = temp;
    updateConfig({ tableColumns: cols });
  };

  // Logo Upload Handler
  const handleLogoUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      updateConfig({
        logo: {
          ...templateConfig.logo,
          show: true,
          src: evt.target.result
        }
      });
    };
    reader.readAsDataURL(file);
  };

  // Realistic Mock Fallback when no dataset is loaded
  const sampleRows = useMemo(() => {
    return [
      { 'slNo': 1, 'seatNo': '4PR24BB001', 'name': 'ADITHYA K', 'course': 'KU4VACBBA200 - Disaster Management', 'status': 'Registered', 'Department': 'Management Studies', 'Venue': 'Sree Narayana Guru College, Thottada' },
      { 'slNo': 2, 'seatNo': '4PR24BB002', 'name': 'ANANYA RAJEEV', 'course': 'KU4SECBBA201 - Soft Skills & Personality', 'status': 'Registered', 'Department': 'Management Studies', 'Venue': 'Sree Narayana Guru College, Thottada' },
      { 'slNo': 3, 'seatNo': '4PR24BB003', 'name': 'FARHAN MOHAMMED', 'course': 'KU4MDCBBA202 - Digital Marketing & E-Commerce', 'status': 'Registered', 'Department': 'Commerce', 'Venue': 'Govt. Brennen College, Thalassery' },
      { 'slNo': 4, 'seatNo': '4PR24BB004', 'name': 'GAYATHRI S', 'course': 'KU4AECBBA203 - Business Ethics', 'status': 'Registered', 'Department': 'Commerce', 'Venue': 'Govt. Brennen College, Thalassery' },
      { 'slNo': 5, 'seatNo': '4PR24BB005', 'name': 'HARIKRISHNAN P', 'course': 'KU4VACBBA200 - Disaster Management', 'status': 'Registered', 'Department': 'Management Studies', 'Venue': 'Sree Narayana Guru College, Thottada' }
    ];
  }, []);

  // Processed Data & Grouping Engine
  const processedData = useMemo(() => {
    const rawRows = dataset?.rows?.length ? dataset.rows : sampleRows;
    const groupByCol = templateConfig.groupBy;

    if (!groupByCol || groupByCol === 'none') {
      return {
        groups: { 'All Records': rawRows },
        groupKeys: ['All Records'],
        totalCount: 1,
        totalRows: rawRows.length
      };
    }

    const groups = {};
    rawRows.forEach(row => {
      const gVal = String(row[groupByCol] !== undefined && row[groupByCol] !== null ? row[groupByCol] : 'Unassigned').trim();
      if (!groups[gVal]) {
        groups[gVal] = [];
      }
      groups[gVal].push(row);
    });

    const groupKeys = Object.keys(groups).sort();
    return {
      groups,
      groupKeys,
      totalCount: groupKeys.length,
      totalRows: rawRows.length
    };
  }, [dataset, sampleRows, templateConfig.groupBy]);

  const effectiveGroupKey = processedData.groupKeys[activePreviewIndex] || processedData.groupKeys[0] || 'All Records';
  const currentPreviewRows = processedData.groups[effectiveGroupKey] || [];
  const activeCols = templateConfig.tableColumns || DEFAULT_REPORT_TEMPLATE.tableColumns;

  // Draw Dynamic Headers on PDF
  const drawDynamicHeaders = (doc, pageWidth) => {
    let startY = 12;

    // 1. Draw Logo if enabled
    if (templateConfig.logo?.show && templateConfig.logo?.src) {
      try {
        const logoW = templateConfig.logo.width || 18;
        const logoH = logoW;
        let logoX = (pageWidth / 2) - (logoW / 2);

        if (templateConfig.logo.position === 'left') {
          logoX = 16;
        } else if (templateConfig.logo.position === 'right') {
          logoX = pageWidth - 16 - logoW;
        }

        doc.addImage(templateConfig.logo.src, 'PNG', logoX, startY, logoW, logoH);
        if (templateConfig.logo.position === 'top' || templateConfig.logo.position === 'center') {
          startY += logoH + 3;
        }
      } catch (e) {
        console.warn('Could not draw logo:', e);
      }
    }

    // 2. Draw Header Lines
    (templateConfig.headersList || []).forEach(h => {
      if (!h.text || !h.text.trim()) return;

      doc.setFont(h.font || 'helvetica', h.bold ? (h.italic ? 'bolditalic' : 'bold') : (h.italic ? 'italic' : 'normal'));
      doc.setFontSize(h.size || 11);

      const [r, g, b] = hexToRgb(h.color || '#000000');
      doc.setTextColor(r, g, b);

      let textX = pageWidth / 2;
      let alignOption = 'center';
      if (h.align === 'left') {
        textX = 16;
        alignOption = 'left';
      } else if (h.align === 'right') {
        textX = pageWidth - 16;
        alignOption = 'right';
      }

      const lines = doc.splitTextToSize(h.text, pageWidth - 32);
      doc.text(lines, textX, startY + (h.size * 0.35), { align: alignOption });
      startY += lines.length * (h.size * 0.42) + 2;
    });

    return startY + 4;
  };

  // Generate Vector PDF for any dataset group
  const generatePdfForGroup = (groupRows, groupTitle, targetDoc = null) => {
    const isLandscape = templateConfig.orientation === 'landscape';
    const doc = targetDoc || new jsPDF(isLandscape ? 'l' : 'p', 'mm', (templateConfig.pageSize || 'a4').toLowerCase());
    const pageWidth = doc.internal.pageSize.getWidth();

    let startY = drawDynamicHeaders(doc, pageWidth);

    // Optional Group Header Banner
    if (templateConfig.groupBy && templateConfig.groupBy !== 'none' && groupTitle) {
      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(23, 107, 135);
      doc.text(`${templateConfig.groupBy}: ${groupTitle}`, 16, startY);
      startY += 6;
    }

    const cols = templateConfig.tableColumns || DEFAULT_REPORT_TEMPLATE.tableColumns;
    const tableBody = groupRows.map((row, idx) => {
      return cols.map(col => {
        if (col.field === 'slNo') return idx + 1;
        if (col.field === 'blank') return '';
        if (col.field === 'seatNo') return row[columnMappings.seatNo] || row['seatNo'] || row['Register Number'] || row['Register No'] || '';
        if (col.field === 'name') return row[columnMappings.name] || row['name'] || row['Candidate Name'] || '';
        return row[col.field] !== undefined ? row[col.field] : (row[col.label] || '');
      });
    });

    const [hBgR, hBgG, hBgB] = hexToRgb(templateConfig.tableTheme?.headerBg || '#f1f5f9');
    const [hTxtR, hTxtG, hTxtB] = hexToRgb(templateConfig.tableTheme?.headerColor || '#000000');
    const [bdrR, bdrG, bdrB] = hexToRgb(templateConfig.tableTheme?.borderColor || '#64748b');

    const colStylesObj = {};
    cols.forEach((c, i) => {
      colStylesObj[i] = {
        cellWidth: isLandscape ? Math.round((c.width || 25) * 1.3) : (c.width || 25),
        halign: c.align || 'left',
        valign: 'middle'
      };
      if (c.bold) colStylesObj[i].fontStyle = 'bold';
    });

    autoTable(doc, {
      startY,
      head: [cols.map(c => c.label)],
      body: tableBody,
      theme: 'grid',
      styles: { 
        font: 'helvetica', 
        fontSize: templateConfig.tableTheme?.fontSize || 8.5, 
        valign: 'middle', 
        cellPadding: 2.5, 
        textColor: [0, 0, 0], 
        lineColor: [bdrR, bdrG, bdrB], 
        lineWidth: 0.2 
      },
      headStyles: { 
        fillColor: [hBgR, hBgG, hBgB], 
        textColor: [hTxtR, hTxtG, hTxtB], 
        fontStyle: 'bold', 
        halign: 'center', 
        valign: 'middle', 
        fontSize: 9, 
        lineColor: [bdrR, bdrG, bdrB], 
        lineWidth: 0.25 
      },
      columnStyles: colStylesObj
    });

    return doc;
  };

  // --- MULTI-FORMAT EXPORT HANDLERS --- //

  // 1. Single Page Vector PDF
  const handleDownloadSinglePdf = () => {
    if (!currentPreviewRows.length) return alert('No data available to export.');
    const doc = generatePdfForGroup(currentPreviewRows, effectiveGroupKey);
    const safeName = templateName.replace(/[^a-zA-Z0-9_-]/g, '_');
    doc.save(`${safeName}_Page_${activePreviewIndex + 1}.pdf`);
  };

  // 2. Master Consolidated Vector PDF
  const handleDownloadConsolidatedPdf = () => {
    if (!processedData.groupKeys.length) return alert('No records found to export.');
    setIsProcessing(true);

    try {
      const isLandscape = templateConfig.orientation === 'landscape';
      const doc = new jsPDF(isLandscape ? 'l' : 'p', 'mm', (templateConfig.pageSize || 'a4').toLowerCase());

      processedData.groupKeys.forEach((gKey, idx) => {
        if (idx > 0) doc.addPage();
        generatePdfForGroup(processedData.groups[gKey], gKey, doc);
      });

      const safeName = templateName.replace(/[^a-zA-Z0-9_-]/g, '_');
      doc.save(`Master_Consolidated_${safeName}.pdf`);
    } catch (err) {
      console.error(err);
      alert(`Error generating PDF: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  // 3. Export Styled Excel Spreadsheet (.xlsx)
  const handleExportExcel = () => {
    const rawRows = dataset?.rows?.length ? dataset.rows : sampleRows;
    if (!rawRows.length) return alert('No data to export to Excel.');

    try {
      const wb = XLSX.utils.book_new();
      const cols = templateConfig.tableColumns || DEFAULT_REPORT_TEMPLATE.tableColumns;

      const appendSheetForGroup = (sheetRows, sheetTitle) => {
        const sheetData = [];

        // 1. Top Header Rows
        (templateConfig.headersList || []).forEach(h => {
          if (h.text && h.text.trim()) {
            sheetData.push([h.text.trim()]);
          }
        });
        if (templateConfig.groupBy && templateConfig.groupBy !== 'none' && sheetTitle) {
          sheetData.push([`${templateConfig.groupBy}: ${sheetTitle}`]);
        }
        sheetData.push([]); // blank row

        // 2. Column Headers
        sheetData.push(cols.map(c => c.label));

        // 3. Data Rows
        sheetRows.forEach((row, idx) => {
          const rowData = cols.map(c => {
            if (c.field === 'slNo') return idx + 1;
            if (c.field === 'blank') return '';
            if (c.field === 'seatNo') return row[columnMappings.seatNo] || row['seatNo'] || row['Register Number'] || row['Register No'] || '';
            if (c.field === 'name') return row[columnMappings.name] || row['name'] || row['Candidate Name'] || '';
            return row[c.field] !== undefined ? row[c.field] : (row[c.label] || '');
          });
          sheetData.push(rowData);
        });

        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        ws['!cols'] = cols.map(c => ({ wch: Math.max(String(c.label).length + 4, 16) }));
        const safeSheetName = (sheetTitle || 'Report Data').replace(/[/\\?%*:|"<>]/g, '_').slice(0, 31);
        XLSX.utils.book_append_sheet(wb, ws, safeSheetName);
      };

      if (templateConfig.groupBy && templateConfig.groupBy !== 'none') {
        processedData.groupKeys.forEach(gKey => {
          appendSheetForGroup(processedData.groups[gKey], gKey);
        });
      } else {
        appendSheetForGroup(rawRows, 'Report Data');
      }

      const safeName = templateName.replace(/[^a-zA-Z0-9_-]/g, '_');
      XLSX.writeFile(wb, `${safeName}_Report.xlsx`);
    } catch (err) {
      console.error(err);
      alert(`Error exporting Excel: ${err.message}`);
    }
  };

  // 4. Export CSV (.csv)
  const handleExportCsv = () => {
    const rawRows = dataset?.rows?.length ? dataset.rows : sampleRows;
    if (!rawRows.length) return alert('No data to export.');
    const cols = templateConfig.tableColumns || DEFAULT_REPORT_TEMPLATE.tableColumns;
    
    const headerRow = cols.map(c => `"${c.label.replace(/"/g, '""')}"`).join(',');
    const dataRows = rawRows.map((row, idx) => {
      return cols.map(c => {
        let val;
        if (c.field === 'slNo') val = idx + 1;
        else if (c.field === 'blank') val = '';
        else if (c.field === 'seatNo') val = row[columnMappings.seatNo] || row['seatNo'] || '';
        else if (c.field === 'name') val = row[columnMappings.name] || row['name'] || '';
        else val = row[c.field] !== undefined ? row[c.field] : (row[c.label] || '');
        return `"${String(val).replace(/"/g, '""')}"`;
      }).join(',');
    }).join('\n');

    const csvContent = `${headerRow}\n${dataRows}`;
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${templateName.replace(/[^a-zA-Z0-9_-]/g, '_')}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // 5. Individual Files in ZIP (PDFs & Excel files per group)
  const handleDownloadZip = async () => {
    if (!processedData.groupKeys.length) return alert('No data to export.');
    setIsProcessing(true);

    try {
      const zip = new JSZip();
      const sanitize = (str) => String(str || '').replace(/[/\\?%*:|"<>•]/g, '_').replace(/\s+/g, ' ').trim();

      processedData.groupKeys.forEach((gKey) => {
        const groupRows = processedData.groups[gKey];
        const doc = generatePdfForGroup(groupRows, gKey);
        const pdfBlob = doc.output('blob');
        const safeGroup = sanitize(gKey).slice(0, 50);
        zip.file(`PDF_Reports/${safeGroup}_Report.pdf`, pdfBlob);
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${templateName.replace(/[^a-zA-Z0-9_-]/g, '_')}_Individual_PDFs.zip`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      alert(`Error creating ZIP: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  // Save Template to Cloud and Local Workspace
  const handleSaveTemplate = async () => {
    if (!templateName.trim()) return alert('Please enter a template name.');
    
    const templateData = {
      name: templateName,
      archetype: 'CUSTOM_TABULAR',
      config: templateConfig,
      columnMappings,
      createdAt: new Date().toISOString()
    };

    let authUserId = null;
    try {
      const { data: authData } = await supabase.auth.getUser();
      if (authData?.user?.id) {
        authUserId = authData.user.id;
      }
    } catch {
      // not logged in
    }

    let savedToCloud = false;
    if (authUserId) {
      const { error } = await supabase
        .from('templates')
        .insert([{ name: templateName, layout_data: templateData, user_id: authUserId }]);
      if (!error) {
        savedToCloud = true;
      }
    }

    // Always persist to local workspace / storage
    try {
      const localTemplates = JSON.parse(localStorage.getItem('saved_templates') || '[]');
      const newTemplate = { id: `tpl_${Date.now()}`, name: templateName, layout_data: templateData, created_at: new Date().toISOString() };
      const updated = [newTemplate, ...localTemplates.filter(t => t.name !== templateName)];
      localStorage.setItem('saved_templates', JSON.stringify(updated));
    } catch (e) {
      console.warn('LocalStorage error:', e);
    }

    if (savedToCloud) {
      alert('Template saved to Cloud and Local Workspace successfully!');
    } else {
      alert('Template saved successfully to your Local Workspace!');
    }
  };

  return (
    <div style={{ padding: '24px 32px 80px', maxWidth: '1600px', margin: '0 auto', fontFamily: 'var(--font-family)' }}>
      
      {/* Top Title & Save Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '24px', flexWrap: 'wrap', gap: '16px' }}>
        <div>
          <h1 style={{ margin: 0, fontSize: '24px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '10px' }}>
            <Sparkles size={24} color="var(--accent)" /> Report Template Studio
          </h1>
          <p style={{ margin: '4px 0 0 0', color: 'var(--muted)', fontSize: '13.5px' }}>
            Design custom reports for any Excel spreadsheet with rich headers, customizable columns, grouping, and multi-format exports (.pdf, .xlsx, .csv).
          </p>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
          <input 
            type="text"
            value={templateName}
            onChange={(e) => setTemplateName(e.target.value)}
            placeholder="Template Name..."
            style={{ padding: '8px 14px', fontSize: '13px', fontWeight: 600, borderRadius: '8px', border: '1.5px solid var(--line)', width: '260px' }}
          />
          <button 
            className="button"
            onClick={handleSaveTemplate}
            style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '8px 16px' }}
          >
            <Save size={15} /> Save Template
          </button>
        </div>
      </div>

      {/* Main Studio Two-Column Grid: Config Controls (Left) + WYSIWYG A4 Preview (Right) */}
      <div style={{ display: 'grid', gridTemplateColumns: '460px 1fr', gap: '24px' }}>
        
        {/* LEFT COLUMN: Customization Sub-Tabs */}
        <div>
          
          {/* Sub-Tab Navigation Bar */}
          <div style={{ display: 'flex', background: 'var(--panel)', borderRadius: '8px', padding: '4px', marginBottom: '16px', gap: '4px', border: '1px solid var(--line)', flexWrap: 'wrap' }}>
            <button
              onClick={() => setActiveToolTab('columns')}
              style={{
                flex: 1,
                minWidth: '80px',
                padding: '8px 4px',
                fontSize: '11.5px',
                fontWeight: 700,
                borderRadius: '6px',
                border: 'none',
                cursor: 'pointer',
                background: activeToolTab === 'columns' ? 'var(--accent)' : 'transparent',
                color: activeToolTab === 'columns' ? 'white' : 'var(--ink)'
              }}
            >
              📊 Table Columns
            </button>
            <button
              onClick={() => setActiveToolTab('headers')}
              style={{
                flex: 1,
                minWidth: '80px',
                padding: '8px 4px',
                fontSize: '11.5px',
                fontWeight: 700,
                borderRadius: '6px',
                border: 'none',
                cursor: 'pointer',
                background: activeToolTab === 'headers' ? 'var(--accent)' : 'transparent',
                color: activeToolTab === 'headers' ? 'white' : 'var(--ink)'
              }}
            >
              📑 Page Headers
            </button>
            <button
              onClick={() => setActiveToolTab('page')}
              style={{
                flex: 1,
                minWidth: '80px',
                padding: '8px 4px',
                fontSize: '11.5px',
                fontWeight: 700,
                borderRadius: '6px',
                border: 'none',
                cursor: 'pointer',
                background: activeToolTab === 'page' ? 'var(--accent)' : 'transparent',
                color: activeToolTab === 'page' ? 'white' : 'var(--ink)'
              }}
            >
              🖼️ Page & Logo
            </button>
            <button
              onClick={() => setActiveToolTab('grouping')}
              style={{
                flex: 1,
                minWidth: '80px',
                padding: '8px 4px',
                fontSize: '11.5px',
                fontWeight: 700,
                borderRadius: '6px',
                border: 'none',
                cursor: 'pointer',
                background: activeToolTab === 'grouping' ? 'var(--accent)' : 'transparent',
                color: activeToolTab === 'grouping' ? 'white' : 'var(--ink)'
              }}
            >
              🗂️ Grouping
            </button>
            <button
              onClick={() => setActiveToolTab('styling')}
              style={{
                flex: 1,
                minWidth: '80px',
                padding: '8px 4px',
                fontSize: '11.5px',
                fontWeight: 700,
                borderRadius: '6px',
                border: 'none',
                cursor: 'pointer',
                background: activeToolTab === 'styling' ? 'var(--accent)' : 'transparent',
                color: activeToolTab === 'styling' ? 'white' : 'var(--ink)'
              }}
            >
              🎨 Table Colors
            </button>
          </div>

          {/* TAB 1: TABLE COLUMN HEADERS & LABELS (ADD / REMOVE / REORDER) */}
          {activeToolTab === 'columns' && (
            <div className="card" style={{ padding: '20px', marginBottom: '20px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px', flexWrap: 'wrap', gap: '8px' }}>
                <div>
                  <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <TableProperties size={16} /> Table Columns ({activeCols.length})
                  </h3>
                  <p style={{ margin: '2px 0 0', fontSize: '11.5px', color: 'var(--muted)' }}>
                    Customize the columns and headers printed in the report table.
                  </p>
                </div>
                <div style={{ display: 'flex', gap: '6px' }}>
                  {dataset?.columns?.length > 0 && (
                    <button
                      className="button secondary"
                      onClick={importAllDatasetColumns}
                      style={{ fontSize: '11.5px', padding: '5px 8px' }}
                      title="Import all columns from uploaded Excel"
                    >
                      Import All
                    </button>
                  )}
                  <button
                    className="button"
                    onClick={addTableColumn}
                    style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '12px', padding: '5px 10px' }}
                  >
                    <Plus size={14} /> Add Column
                  </button>
                </div>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', maxHeight: '560px', overflowY: 'auto', paddingRight: '4px' }}>
                {activeCols.map((col, idx) => (
                  <div key={col.id || idx} style={{ padding: '12px', background: 'var(--bg)', borderRadius: '8px', border: '1px solid var(--line)' }}>
                    
                    {/* Top Row: Column # & Move/Delete */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                      <span style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>
                        Column #{idx + 1}
                      </span>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <button 
                          disabled={idx === 0} 
                          onClick={() => moveTableColumn(idx, -1)} 
                          style={{ background: 'none', border: 'none', cursor: idx === 0 ? 'default' : 'pointer', color: idx === 0 ? '#ccc' : 'var(--ink)', padding: '2px' }}
                          title="Move Left"
                        >
                          <ArrowLeft size={14} />
                        </button>
                        <button 
                          disabled={idx === activeCols.length - 1} 
                          onClick={() => moveTableColumn(idx, 1)} 
                          style={{ background: 'none', border: 'none', cursor: idx === activeCols.length - 1 ? 'default' : 'pointer', color: idx === activeCols.length - 1 ? '#ccc' : 'var(--ink)', padding: '2px' }}
                          title="Move Right"
                        >
                          <ArrowRightIcon size={14} />
                        </button>
                        <button 
                          onClick={() => removeTableColumn(col.id)} 
                          style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--danger)', padding: '2px', marginLeft: '4px' }}
                          title="Remove Column"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </div>

                    {/* Header Label Input */}
                    <div style={{ marginBottom: '8px' }}>
                      <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '3px' }}>Header Title:</label>
                      <input 
                        type="text"
                        value={col.label || ''}
                        onChange={(e) => updateTableColumn(col.id, { label: e.target.value })}
                        placeholder="Column Header Text..."
                        style={{ width: '100%', fontSize: '13px', fontWeight: 700 }}
                      />
                    </div>

                    {/* Field Binding & Alignment */}
                    <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr', gap: '8px', marginBottom: '8px' }}>
                      <div>
                        <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '3px' }}>Data Content:</label>
                        <select
                          value={col.field || 'blank'}
                          onChange={(e) => updateTableColumn(col.id, { field: e.target.value })}
                          style={{ width: '100%', fontSize: '12px' }}
                        >
                          <optgroup label="Standard Fields">
                            <option value="slNo">Serial No (1, 2, 3...)</option>
                            <option value="blank">Blank / Signature Box</option>
                          </optgroup>
                          {dataset.columns?.length > 0 ? (
                            <optgroup label="Excel Columns">
                              {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                            </optgroup>
                          ) : (
                            <optgroup label="Sample Columns">
                              <option value="seatNo">Register Number</option>
                              <option value="name">Candidate Name</option>
                              <option value="course">Course / Subject</option>
                              <option value="status">Status</option>
                              <option value="Department">Department</option>
                              <option value="Venue">Venue</option>
                            </optgroup>
                          )}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '3px' }}>Alignment:</label>
                        <div style={{ display: 'flex', gap: '3px' }}>
                          {['left', 'center', 'right'].map(align => (
                            <button
                              key={align}
                              type="button"
                              onClick={() => updateTableColumn(col.id, { align })}
                              style={{
                                flex: 1,
                                padding: '4px',
                                textTransform: 'capitalize',
                                fontSize: '11px',
                                borderRadius: '4px',
                                border: '1px solid var(--line)',
                                background: col.align === align ? 'var(--accent)' : 'white',
                                color: col.align === align ? 'white' : 'var(--ink)',
                                cursor: 'pointer'
                              }}
                            >
                              {align}
                            </button>
                          ))}
                        </div>
                      </div>
                    </div>

                    {/* Width & Bold Controls */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Width:</span>
                        <input 
                          type="number"
                          min="10"
                          max="150"
                          value={col.width || 30}
                          onChange={(e) => updateTableColumn(col.id, { width: parseInt(e.target.value, 10) || 30 })}
                          style={{ width: '56px', padding: '3px 6px', fontSize: '11.5px', textAlign: 'center' }}
                        />
                        <span style={{ fontSize: '11px', color: 'var(--muted)' }}>mm</span>
                      </div>

                      <label style={{ display: 'flex', alignItems: 'center', gap: '4px', fontSize: '11.5px', cursor: 'pointer' }}>
                        <input 
                          type="checkbox"
                          checked={!!col.bold}
                          onChange={(e) => updateTableColumn(col.id, { bold: e.target.checked })}
                          style={{ accentColor: 'var(--accent)' }}
                        />
                        Bold Text
                      </label>
                    </div>

                  </div>
                ))}
              </div>
            </div>
          )}

          {/* TAB 2: PAGE TITLE & INSTITUTIONAL HEADERS */}
          {activeToolTab === 'headers' && (
            <div className="card" style={{ padding: '20px', marginBottom: '20px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
                <div>
                  <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Edit3 size={16} /> Page Title & Sub-Headers ({templateConfig.headersList?.length || 0})
                  </h3>
                  <p style={{ margin: '2px 0 0', fontSize: '11.5px', color: 'var(--muted)' }}>
                    Add, remove, reorder, and format institutional headers at the top of the report.
                  </p>
                </div>
                <button
                  className="button"
                  onClick={addHeaderLine}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '12px', padding: '5px 10px' }}
                >
                  <Plus size={14} /> Add Line
                </button>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '14px', maxHeight: '550px', overflowY: 'auto', paddingRight: '4px' }}>
                {(templateConfig.headersList || []).map((h, idx) => (
                  <div 
                    key={h.id} 
                    style={{ 
                      padding: '12px', 
                      background: 'var(--bg)', 
                      borderRadius: '8px', 
                      border: '1px solid var(--line)' 
                    }}
                  >
                    {/* Top Row: Line Label & Move/Delete */}
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
                      <span style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase' }}>
                        Header #{idx + 1}
                      </span>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <button 
                          disabled={idx === 0} 
                          onClick={() => moveHeaderLine(idx, -1)} 
                          style={{ background: 'none', border: 'none', cursor: idx === 0 ? 'default' : 'pointer', color: idx === 0 ? '#ccc' : 'var(--ink)', padding: '2px' }}
                          title="Move Up"
                        >
                          <ArrowUp size={14} />
                        </button>
                        <button 
                          disabled={idx === templateConfig.headersList.length - 1} 
                          onClick={() => moveHeaderLine(idx, 1)} 
                          style={{ background: 'none', border: 'none', cursor: idx === templateConfig.headersList.length - 1 ? 'default' : 'pointer', color: idx === templateConfig.headersList.length - 1 ? '#ccc' : 'var(--ink)', padding: '2px' }}
                          title="Move Down"
                        >
                          <ArrowDown size={14} />
                        </button>
                        <button 
                          onClick={() => removeHeaderLine(h.id)} 
                          style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--danger)', padding: '2px', marginLeft: '4px' }}
                          title="Delete Line"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </div>

                    {/* Text Input */}
                    <input 
                      type="text"
                      value={h.text || ''}
                      onChange={(e) => updateHeaderLine(h.id, { text: e.target.value })}
                      placeholder="Header text..."
                      style={{ width: '100%', fontSize: '13px', fontWeight: h.bold ? 700 : 400, fontStyle: h.italic ? 'italic' : 'normal', marginBottom: '8px' }}
                    />

                    {/* Styling Controls Toolbar */}
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '8px' }}>
                      
                      {/* Font Size */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Size:</span>
                        <input 
                          type="number"
                          min="8"
                          max="32"
                          value={h.size || 11}
                          onChange={(e) => updateHeaderLine(h.id, { size: parseFloat(e.target.value) || 11 })}
                          style={{ width: '52px', padding: '4px 6px', fontSize: '12px', textAlign: 'center' }}
                        />
                        <span style={{ fontSize: '11px', color: 'var(--muted)' }}>pt</span>
                      </div>

                      {/* Color Picker */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <span style={{ fontSize: '11px', color: 'var(--muted)' }}>Color:</span>
                        <input 
                          type="color"
                          value={h.color || '#000000'}
                          onChange={(e) => updateHeaderLine(h.id, { color: e.target.value })}
                          style={{ width: '28px', height: '28px', border: 'none', borderRadius: '4px', cursor: 'pointer', background: 'transparent' }}
                        />
                      </div>

                      {/* Bold / Italic Toggles */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <button
                          type="button"
                          onClick={() => updateHeaderLine(h.id, { bold: !h.bold })}
                          style={{
                            padding: '4px 8px',
                            borderRadius: '4px',
                            border: '1px solid var(--line)',
                            background: h.bold ? 'var(--accent)' : 'white',
                            color: h.bold ? 'white' : 'var(--ink)',
                            cursor: 'pointer'
                          }}
                        >
                          <Bold size={13} />
                        </button>
                        <button
                          type="button"
                          onClick={() => updateHeaderLine(h.id, { italic: !h.italic })}
                          style={{
                            padding: '4px 8px',
                            borderRadius: '4px',
                            border: '1px solid var(--line)',
                            background: h.italic ? 'var(--accent)' : 'white',
                            color: h.italic ? 'white' : 'var(--ink)',
                            cursor: 'pointer'
                          }}
                        >
                          <Italic size={13} />
                        </button>
                      </div>

                      {/* Alignment */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        <button
                          type="button"
                          onClick={() => updateHeaderLine(h.id, { align: 'left' })}
                          style={{
                            padding: '4px 6px',
                            borderRadius: '4px',
                            border: '1px solid var(--line)',
                            background: h.align === 'left' ? 'var(--accent)' : 'white',
                            color: h.align === 'left' ? 'white' : 'var(--ink)',
                            cursor: 'pointer'
                          }}
                        >
                          <AlignLeft size={13} />
                        </button>
                        <button
                          type="button"
                          onClick={() => updateHeaderLine(h.id, { align: 'center' })}
                          style={{
                            padding: '4px 6px',
                            borderRadius: '4px',
                            border: '1px solid var(--line)',
                            background: h.align === 'center' || !h.align ? 'var(--accent)' : 'white',
                            color: h.align === 'center' || !h.align ? 'white' : 'var(--ink)',
                            cursor: 'pointer'
                          }}
                        >
                          <AlignCenter size={13} />
                        </button>
                        <button
                          type="button"
                          onClick={() => updateHeaderLine(h.id, { align: 'right' })}
                          style={{
                            padding: '4px 6px',
                            borderRadius: '4px',
                            border: '1px solid var(--line)',
                            background: h.align === 'right' ? 'var(--accent)' : 'white',
                            color: h.align === 'right' ? 'white' : 'var(--ink)',
                            cursor: 'pointer'
                          }}
                        >
                          <AlignRight size={13} />
                        </button>
                      </div>

                    </div>

                  </div>
                ))}
              </div>
            </div>
          )}

          {/* TAB 3: PAGE ORIENTATION & LOGO CONTROLS */}
          {activeToolTab === 'page' && (
            <div className="card" style={{ padding: '20px', marginBottom: '20px' }}>
              <h3 style={{ margin: '0 0 16px 0', fontSize: '15px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Sliders size={16} /> Page Orientation & Custom Logo
              </h3>

              {/* Orientation Buttons */}
              <div style={{ marginBottom: '18px' }}>
                <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>Page Orientation:</label>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
                  <button
                    type="button"
                    onClick={() => updateConfig({ orientation: 'portrait' })}
                    style={{
                      padding: '10px',
                      borderRadius: '8px',
                      border: templateConfig.orientation === 'portrait' || !templateConfig.orientation ? '2px solid var(--accent)' : '1px solid var(--line)',
                      background: templateConfig.orientation === 'portrait' || !templateConfig.orientation ? 'var(--accent-soft)' : 'white',
                      fontWeight: 700,
                      fontSize: '13px',
                      color: templateConfig.orientation === 'portrait' || !templateConfig.orientation ? 'var(--accent)' : 'var(--ink)',
                      cursor: 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '6px'
                    }}
                  >
                    📄 Portrait
                  </button>
                  <button
                    type="button"
                    onClick={() => updateConfig({ orientation: 'landscape' })}
                    style={{
                      padding: '10px',
                      borderRadius: '8px',
                      border: templateConfig.orientation === 'landscape' ? '2px solid var(--accent)' : '1px solid var(--line)',
                      background: templateConfig.orientation === 'landscape' ? 'var(--accent-soft)' : 'white',
                      fontWeight: 700,
                      fontSize: '13px',
                      color: templateConfig.orientation === 'landscape' ? 'var(--accent)' : 'var(--ink)',
                      cursor: 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '6px'
                    }}
                  >
                    📜 Landscape
                  </button>
                </div>
              </div>

              {/* Page Size */}
              <div style={{ marginBottom: '18px' }}>
                <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>Page Standard Size:</label>
                <select
                  value={templateConfig.pageSize || 'A4'}
                  onChange={(e) => updateConfig({ pageSize: e.target.value })}
                  style={{ width: '100%', fontSize: '13px' }}
                >
                  <option value="A4">A4 (210 × 297 mm)</option>
                  <option value="A3">A3 (297 × 420 mm)</option>
                  <option value="Letter">US Letter (8.5 × 11 in)</option>
                </select>
              </div>

              {/* Logo Upload & Positioning */}
              <div style={{ padding: '14px', background: 'var(--bg)', borderRadius: '8px', border: '1px solid var(--line)' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '13px', fontWeight: 700, marginBottom: '10px' }}>
                  <input 
                    type="checkbox"
                    checked={!!templateConfig.logo?.show}
                    onChange={(e) => updateConfig({ logo: { ...templateConfig.logo, show: e.target.checked } })}
                    style={{ width: '16px', height: '16px', accentColor: 'var(--accent)' }}
                  />
                  Show Logo / Crest on Page
                </label>

                {templateConfig.logo?.show && (
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginTop: '10px' }}>
                    
                    {/* Upload button */}
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      <button
                        type="button"
                        className="button secondary"
                        onClick={() => fileInputRef.current?.click()}
                        style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12px', padding: '6px 12px' }}
                      >
                        <ImageIcon size={14} /> Upload Custom Logo
                      </button>
                      <input 
                        type="file" 
                        ref={fileInputRef} 
                        accept="image/*" 
                        onChange={handleLogoUpload} 
                        style={{ display: 'none' }} 
                      />

                      <button
                        type="button"
                        className="button secondary"
                        onClick={() => updateConfig({ logo: { ...templateConfig.logo, src: logoBase64 } })}
                        style={{ fontSize: '11px', padding: '6px 10px' }}
                        title="Reset to default Crest"
                      >
                        Default Crest
                      </button>
                    </div>

                    {/* Logo Position */}
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                      <span style={{ fontSize: '12px', fontWeight: 600 }}>Position:</span>
                      <div style={{ display: 'flex', gap: '4px' }}>
                        {['left', 'center', 'right', 'top'].map(pos => (
                          <button
                            key={pos}
                            type="button"
                            onClick={() => updateConfig({ logo: { ...templateConfig.logo, position: pos } })}
                            style={{
                              padding: '4px 8px',
                              fontSize: '11px',
                              textTransform: 'capitalize',
                              borderRadius: '4px',
                              border: '1px solid var(--line)',
                              background: templateConfig.logo?.position === pos ? 'var(--accent)' : 'white',
                              color: templateConfig.logo?.position === pos ? 'white' : 'var(--ink)',
                              cursor: 'pointer'
                            }}
                          >
                            {pos}
                          </button>
                        ))}
                      </div>
                    </div>

                    {/* Logo Width */}
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                      <span style={{ fontSize: '12px', fontWeight: 600 }}>Size (Width):</span>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                        <input 
                          type="range"
                          min="10"
                          max="40"
                          value={templateConfig.logo?.width || 18}
                          onChange={(e) => updateConfig({ logo: { ...templateConfig.logo, width: parseInt(e.target.value, 10) } })}
                          style={{ width: '90px' }}
                        />
                        <span style={{ fontSize: '11px', fontWeight: 700 }}>{templateConfig.logo?.width || 18} mm</span>
                      </div>
                    </div>

                  </div>
                )}
              </div>

            </div>
          )}

          {/* TAB 4: GROUPING & TOTALS */}
          {activeToolTab === 'grouping' && (
            <div className="card" style={{ padding: '20px', marginBottom: '20px' }}>
              <h3 style={{ margin: '0 0 16px 0', fontSize: '15px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Layers size={16} /> Grouping & Page Break Rules
              </h3>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                <div>
                  <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '6px' }}>Group Records By Column:</label>
                  <select
                    value={templateConfig.groupBy || 'none'}
                    onChange={(e) => {
                      updateConfig({ groupBy: e.target.value });
                      setActivePreviewIndex(0);
                    }}
                    style={{ width: '100%', fontSize: '13px' }}
                  >
                    <option value="none">No Grouping (Single Continuous Table)</option>
                    {dataset.columns?.length > 0 ? (
                      dataset.columns.map(c => <option key={c} value={c}>Group by: {c}</option>)
                    ) : (
                      <>
                        <option value="Department">Group by: Department</option>
                        <option value="Venue">Group by: Venue</option>
                      </>
                    )}
                  </select>
                  <p style={{ margin: '4px 0 0', fontSize: '11.5px', color: 'var(--muted)' }}>
                    When grouping is enabled, separate pages and sub-files are generated for each unique group value.
                  </p>
                </div>

                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '13px', fontWeight: 600 }}>
                  <input 
                    type="checkbox"
                    checked={!!templateConfig.showTotalCount}
                    onChange={(e) => updateConfig({ showTotalCount: e.target.checked })}
                    style={{ width: '16px', height: '16px', accentColor: 'var(--accent)' }}
                  />
                  Show Total Record Count Footer
                </label>
              </div>
            </div>
          )}

          {/* TAB 5: TABLE STYLING & COLORS */}
          {activeToolTab === 'styling' && (
            <div className="card" style={{ padding: '20px', marginBottom: '20px' }}>
              <h3 style={{ margin: '0 0 16px 0', fontSize: '15px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Sliders size={16} /> Table Colors & Font Size
              </h3>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                
                {/* Table Header Color */}
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <span style={{ fontSize: '12px', fontWeight: 600 }}>Table Header Background:</span>
                  <input 
                    type="color"
                    value={templateConfig.tableTheme?.headerBg || '#f1f5f9'}
                    onChange={(e) => updateConfig({ tableTheme: { ...templateConfig.tableTheme, headerBg: e.target.value } })}
                    style={{ width: '32px', height: '32px', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
                  />
                </div>

                {/* Table Header Text Color */}
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <span style={{ fontSize: '12px', fontWeight: 600 }}>Table Header Text Color:</span>
                  <input 
                    type="color"
                    value={templateConfig.tableTheme?.headerColor || '#000000'}
                    onChange={(e) => updateConfig({ tableTheme: { ...templateConfig.tableTheme, headerColor: e.target.value } })}
                    style={{ width: '32px', height: '32px', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
                  />
                </div>

                {/* Table Border Color */}
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <span style={{ fontSize: '12px', fontWeight: 600 }}>Table Border Color:</span>
                  <input 
                    type="color"
                    value={templateConfig.tableTheme?.borderColor || '#64748b'}
                    onChange={(e) => updateConfig({ tableTheme: { ...templateConfig.tableTheme, borderColor: e.target.value } })}
                    style={{ width: '32px', height: '32px', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
                  />
                </div>

                {/* Table Font Size */}
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <span style={{ fontSize: '12px', fontWeight: 600 }}>Table Content Font Size:</span>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <input 
                      type="number"
                      min="7"
                      max="14"
                      step="0.5"
                      value={templateConfig.tableTheme?.fontSize || 8.5}
                      onChange={(e) => updateConfig({ tableTheme: { ...templateConfig.tableTheme, fontSize: parseFloat(e.target.value) || 8.5 } })}
                      style={{ width: '56px', padding: '4px 6px', fontSize: '12px', textAlign: 'center' }}
                    />
                    <span style={{ fontSize: '11px', color: 'var(--muted)' }}>pt</span>
                  </div>
                </div>

              </div>
            </div>
          )}

        </div>

        {/* RIGHT COLUMN: Live WYSIWYG A4 Preview and Export Toolbar */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
          
          {/* Multi-Format Export Toolbar */}
          <div className="card" style={{ padding: '16px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
            
            {/* Page Navigator when Grouping */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <button 
                className="button secondary"
                disabled={activePreviewIndex <= 0}
                onClick={() => setActivePreviewIndex(p => Math.max(0, p - 1))}
                style={{ padding: '7px 12px', fontSize: '12.5px' }}
              >
                <ChevronLeft size={16} /> Prev
              </button>
              <span style={{ fontSize: '13px', fontWeight: 700, color: 'var(--ink)' }}>
                Page {activePreviewIndex + 1} of {Math.max(1, processedData.totalCount)}
              </span>
              <button 
                className="button secondary"
                disabled={activePreviewIndex >= (processedData.totalCount - 1)}
                onClick={() => setActivePreviewIndex(p => Math.min(processedData.totalCount - 1, p + 1))}
                style={{ padding: '7px 12px', fontSize: '12.5px' }}
              >
                Next <ChevronRight size={16} />
              </button>
            </div>

            {/* Export Buttons */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
              
              <button 
                className="button secondary"
                onClick={handleDownloadSinglePdf}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12px', padding: '7px 12px' }}
                title="Download active page vector PDF"
              >
                <Download size={14} /> Page PDF
              </button>

              <button 
                className="button secondary"
                onClick={handleDownloadConsolidatedPdf}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12px', padding: '7px 12px' }}
                title="Download consolidated all-pages vector PDF"
              >
                <FileText size={14} /> Master PDF
              </button>

              <button 
                className="button"
                onClick={handleExportExcel}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12px', padding: '7px 14px', background: '#107c41', borderColor: '#107c41' }}
                title="Download formatted Excel Spreadsheet (.xlsx) with header rows and custom columns"
              >
                <FileSpreadsheet size={15} /> Export Excel (.xlsx)
              </button>

              <button 
                className="button secondary"
                onClick={handleExportCsv}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12px', padding: '7px 12px' }}
                title="Export clean CSV data"
              >
                <Download size={14} /> CSV
              </button>

              <button 
                className="button secondary"
                onClick={handleDownloadZip}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12px', padding: '7px 12px' }}
                title="Download group-wise split PDFs in ZIP"
              >
                <Archive size={14} /> ZIP
              </button>

              <button 
                className="button secondary"
                onClick={() => window.print()}
                title="Direct Print A4"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12px', padding: '7px 12px' }}
              >
                <Printer size={14} /> Print
              </button>

            </div>

          </div>

          {/* Rendered Live Paper Display (Adapts to Portrait or Landscape) */}
          <div 
            className="card print-container" 
            style={{ 
              padding: templateConfig.orientation === 'landscape' ? '32px 40px' : '40px 48px', 
              background: 'white', 
              border: '1px solid #cbd5e1', 
              borderRadius: '8px', 
              boxShadow: '0 8px 30px rgba(0,0,0,0.08)', 
              minHeight: '750px',
              maxWidth: templateConfig.orientation === 'landscape' ? '1123px' : '794px',
              margin: '0 auto',
              width: '100%',
              transition: 'all 0.3s ease'
            }}
          >
            
            {/* DYNAMIC HEADER RENDERING ON LIVE PAPER */}
            <div style={{ textAlign: 'center', marginBottom: '20px', position: 'relative' }}>
              
              {/* Optional Logo */}
              {templateConfig.logo?.show && templateConfig.logo?.src && (
                <div style={{
                  display: 'flex',
                  justifyContent: templateConfig.logo.position === 'left' ? 'flex-start' : templateConfig.logo.position === 'right' ? 'flex-end' : 'center',
                  marginBottom: '10px'
                }}>
                  <img 
                    src={templateConfig.logo.src} 
                    alt="Logo" 
                    style={{ height: `${(templateConfig.logo.width || 18) * 3}px`, objectFit: 'contain' }} 
                  />
                </div>
              )}

              {/* Dynamic Headers List */}
              {(templateConfig.headersList || []).map((h, idx) => (
                <div 
                  key={h.id || idx}
                  style={{
                    fontSize: `${h.size || 11}pt`,
                    fontWeight: h.bold ? 800 : 400,
                    fontStyle: h.italic ? 'italic' : 'normal',
                    color: h.color || '#000000',
                    textAlign: h.align || 'center',
                    marginBottom: '4px',
                    lineHeight: '1.3'
                  }}
                >
                  {h.text}
                </div>
              ))}

              {/* Optional Group Title Banner */}
              {templateConfig.groupBy && templateConfig.groupBy !== 'none' && (
                <div style={{ 
                  marginTop: '12px', 
                  padding: '6px 12px', 
                  background: 'var(--accent-soft)', 
                  color: 'var(--accent)', 
                  fontWeight: 800, 
                  borderRadius: '6px', 
                  fontSize: '13px',
                  display: 'inline-block'
                }}>
                  {templateConfig.groupBy}: {effectiveGroupKey}
                </div>
              )}

            </div>

            {/* LIVE CUSTOM TABULAR REPORT TABLE */}
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: `${templateConfig.tableTheme?.fontSize || 8.5}pt` }}>
                <thead>
                  <tr style={{ background: templateConfig.tableTheme?.headerBg || '#f1f5f9', color: templateConfig.tableTheme?.headerColor || '#000000', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000' }}>
                    {activeCols.map(c => (
                      <th 
                        key={c.id} 
                        style={{ 
                          border: `1px solid ${templateConfig.tableTheme?.borderColor || '#64748b'}`, 
                          padding: '8px 6px', 
                          textAlign: c.align || 'left', 
                          fontWeight: 700 
                        }}
                      >
                        {c.label}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {currentPreviewRows.slice(0, 40).map((row, idx) => (
                    <tr key={idx} style={{ borderBottom: `1px solid ${templateConfig.tableTheme?.borderColor || '#64748b'}` }}>
                      {activeCols.map(col => {
                        let content;
                        if (col.field === 'slNo') content = idx + 1;
                        else if (col.field === 'blank') content = '';
                        else if (col.field === 'seatNo') content = row[columnMappings.seatNo] || row['seatNo'] || row['Register Number'] || row['Register No'] || '';
                        else if (col.field === 'name') content = row[columnMappings.name] || row['name'] || row['Candidate Name'] || '';
                        else content = row[col.field] !== undefined ? row[col.field] : (row[col.label] || '');

                        return (
                          <td 
                            key={col.id} 
                            style={{ 
                              border: `1px solid ${templateConfig.tableTheme?.borderColor || '#64748b'}`, 
                              padding: '6px 8px', 
                              textAlign: col.align || 'left', 
                              fontWeight: col.bold ? 700 : 400 
                            }}
                          >
                            {content}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>

              {/* Optional Record Count Footer */}
              {templateConfig.showTotalCount && (
                <div style={{ marginTop: '12px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '11.5px', color: 'var(--muted)', fontWeight: 600 }}>
                  <span>
                    {currentPreviewRows.length > 40 ? `⚡ Showing preview of 40 of ${currentPreviewRows.length} records. (Full ${currentPreviewRows.length} records included in PDF/Excel export)` : ''}
                  </span>
                  <span>
                    Total Records: {currentPreviewRows.length} {templateConfig.groupBy !== 'none' ? `(in this group)` : ''}
                  </span>
                </div>
              )}
            </div>

          </div>

        </div>

      </div>

    </div>
  );
};

export default TemplatePage;
