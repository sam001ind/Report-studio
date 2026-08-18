import { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import JSZip from 'jszip';
import { 
  FileText, 
  CalendarRange, 
  Tag, 
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
  ArrowRight as ArrowRightIcon
} from 'lucide-react';
import { supabase } from '../supabaseClient';
import { logoBase64 } from '../assets/logoBase64';
import { TEMPLATE_ARCHETYPES, autoDetectDatasetColumns, suggestArchetype } from '../utils/templateEngine';

const hexToRgb = (hex) => {
  let c = String(hex || '#000000').replace(/^#/, '');
  if (c.length === 3) c = c.split('').map(x => x + x).join('');
  const num = parseInt(c, 16);
  if (isNaN(num)) return [0, 0, 0];
  return [(num >> 16) & 255, (num >> 8) & 255, num & 255];
};

const TemplatePage = ({ dataset = { columns: [], rows: [] }, initialTemplate }) => {
  const fileInputRef = useRef(null);
  
  // Selected Archetype
  const [activeArchetype, setActiveArchetype] = useState(() => {
    if (initialTemplate?.archetype) return initialTemplate.archetype;
    if (dataset?.columns?.length > 0) {
      const detected = autoDetectDatasetColumns(dataset.columns);
      return suggestArchetype(detected);
    }
    return 'NOMINAL_ROLL';
  });

  // Archetype Configurations
  const [archetypeConfigs, setArchetypeConfigs] = useState(() => {
    return {
      NOMINAL_ROLL: JSON.parse(JSON.stringify(TEMPLATE_ARCHETYPES.NOMINAL_ROLL.defaultConfig)),
      QP_STATEMENT: JSON.parse(JSON.stringify(TEMPLATE_ARCHETYPES.QP_STATEMENT.defaultConfig)),
      QP_COVER_LABEL: JSON.parse(JSON.stringify(TEMPLATE_ARCHETYPES.QP_COVER_LABEL.defaultConfig)),
      CUSTOM_TABULAR: JSON.parse(JSON.stringify(TEMPLATE_ARCHETYPES.CUSTOM_TABULAR.defaultConfig))
    };
  });

  // Column Mappings
  const [columnMappings, setColumnMappings] = useState(() => autoDetectDatasetColumns(dataset?.columns || []));

  const [isSidebarOpen, _setIsSidebarOpen] = useState(true);
  const [templateName, setTemplateName] = useState(initialTemplate?.name || 'My Custom University Report');
  const [isProcessing, setIsProcessing] = useState(false);
  const [activePreviewIndex, setActivePreviewIndex] = useState(0);

  // Active Tool Tab: 'headers' | 'columns' | 'data' | 'page' | 'styling'
  const [activeToolTab, setActiveToolTab] = useState('columns');

  // Update column detection when dataset changes
  useEffect(() => {
    if (dataset?.columns?.length > 0) {
      const detected = autoDetectDatasetColumns(dataset.columns);
      setColumnMappings(detected);
      
      // Auto-populate session name if found
      if (detected.session && dataset.rows?.length > 0) {
        const sampleSess = dataset.rows.find(r => r[detected.session])?.[detected.session];
        if (sampleSess && typeof sampleSess === 'string' && sampleSess.trim()) {
          setArchetypeConfigs(prev => {
            const copy = JSON.parse(JSON.stringify(prev));
            ['NOMINAL_ROLL', 'QP_STATEMENT'].forEach(arch => {
              if (copy[arch]?.headersList?.length >= 4) {
                copy[arch].headersList[3].text = sampleSess.trim();
              }
            });
            return copy;
          });
        }
      }
    }
  }, [dataset]);

  // Load from initialTemplate
  useEffect(() => {
    if (initialTemplate) {
      if (initialTemplate.archetype) {
        setActiveArchetype(initialTemplate.archetype);
      }
      if (initialTemplate.config) {
        setArchetypeConfigs(prev => ({
          ...prev,
          [initialTemplate.archetype || 'NOMINAL_ROLL']: initialTemplate.config
        }));
      }
      if (initialTemplate.columnMappings) {
        setColumnMappings(initialTemplate.columnMappings);
      }
      if (initialTemplate.name) {
        setTemplateName(initialTemplate.name);
      }
    }
  }, [initialTemplate]);

  const currentConfig = archetypeConfigs[activeArchetype] || archetypeConfigs.NOMINAL_ROLL;

  const updateCurrentConfig = (updates) => {
    setArchetypeConfigs(prev => ({
      ...prev,
      [activeArchetype]: {
        ...prev[activeArchetype],
        ...updates
      }
    }));
  };

  // Top Page Headers Helpers
  const addHeaderLine = () => {
    const newHeader = {
      id: `h_${Date.now()}`,
      text: 'New Header Line',
      size: 11,
      bold: true,
      italic: false,
      align: 'center',
      color: '#000000',
      font: 'helvetica'
    };
    updateCurrentConfig({
      headersList: [...(currentConfig.headersList || []), newHeader]
    });
  };

  const updateHeaderLine = (id, updates) => {
    const list = (currentConfig.headersList || []).map(h => h.id === id ? { ...h, ...updates } : h);
    updateCurrentConfig({ headersList: list });
  };

  const removeHeaderLine = (id) => {
    const list = (currentConfig.headersList || []).filter(h => h.id !== id);
    updateCurrentConfig({ headersList: list });
  };

  const moveHeaderLine = (index, direction) => {
    const list = [...(currentConfig.headersList || [])];
    const targetIdx = index + direction;
    if (targetIdx < 0 || targetIdx >= list.length) return;
    const temp = list[index];
    list[index] = list[targetIdx];
    list[targetIdx] = temp;
    updateCurrentConfig({ headersList: list });
  };

  // Table Column Headers Helpers
  const addTableColumn = () => {
    const cols = currentConfig.tableColumns || [];
    const newCol = {
      id: `col_${Date.now()}`,
      label: `Column ${cols.length + 1}`,
      field: 'blank',
      width: 25,
      align: 'center',
      isSpan: true
    };
    updateCurrentConfig({ tableColumns: [...cols, newCol] });
  };

  const updateTableColumn = (id, updates) => {
    const cols = (currentConfig.tableColumns || []).map(c => c.id === id ? { ...c, ...updates } : c);
    updateCurrentConfig({ tableColumns: cols });
  };

  const removeTableColumn = (id) => {
    const cols = (currentConfig.tableColumns || []).filter(c => c.id !== id);
    updateCurrentConfig({ tableColumns: cols });
  };

  const moveTableColumn = (index, direction) => {
    const cols = [...(currentConfig.tableColumns || [])];
    const targetIdx = index + direction;
    if (targetIdx < 0 || targetIdx >= cols.length) return;
    const temp = cols[index];
    cols[index] = cols[targetIdx];
    cols[targetIdx] = temp;
    updateCurrentConfig({ tableColumns: cols });
  };

  // Logo Upload Handler
  const handleLogoUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      updateCurrentConfig({
        logo: {
          ...currentConfig.logo,
          show: true,
          src: evt.target.result
        }
      });
    };
    reader.readAsDataURL(file);
  };

  // Mock Fallback Data Generator when no dataset is loaded
  const getMockDataForArchetype = useCallback((archetype, formatDots = false) => {
    if (archetype === 'NOMINAL_ROLL') {
      const prog = 'Bachelor of Business Administration (BBA)';
      const venueLabel = 'GA - Sree Narayana Guru College of Advanced Studies, Thottada';
      const c1 = formatDots ? 'KU4VACBBA200. - Disaster Management.' : 'KU4VACBBA200 - Disaster Management';
      const c2 = formatDots ? 'KU4SECBBA201. - Soft Skills & Personality Development.' : 'KU4SECBBA201 - Soft Skills & Personality Development';
      const c3 = formatDots ? 'KU4MDCBBA202. - Digital Marketing & E-Commerce.' : 'KU4MDCBBA202 - Digital Marketing & E-Commerce';
      const c4 = formatDots ? 'KU4AECBBA203. - Business Ethics & Corporate Governance.' : 'KU4AECBBA203 - Business Ethics & Corporate Governance';

      const groupKey = `${prog} • ${venueLabel}`;
      return {
        groups: {
          [groupKey]: {
            programme: prog,
            venueCode: 'GA',
            venueName: 'Sree Narayana Guru College of Advanced Studies, Thottada',
            venueLabel,
            candidates: [
              {
                seatNo: '4PR24BB001',
                studentName: 'ADITHYA K',
                rawRow: {},
                courses: [
                  { display: c1 },
                  { display: c2 },
                  { display: c3 },
                  { display: c4 }
                ]
              },
              {
                seatNo: '4PR24BB002',
                studentName: 'ANANYA RAJEEV',
                rawRow: {},
                courses: [
                  { display: c1 },
                  { display: c2 },
                  { display: c3 }
                ]
              },
              {
                seatNo: '4PR24BB003',
                studentName: 'FARHAN MOHAMMED',
                rawRow: {},
                courses: [
                  { display: c1 },
                  { display: c3 },
                  { display: c4 }
                ]
              }
            ],
            totalCandidates: 3
          }
        },
        groupKeys: [groupKey],
        totalCount: 1
      };
    }

    if (archetype === 'QP_STATEMENT') {
      const venueLabel = 'GA - Sree Narayana Guru College of Advanced Studies, Thottada';
      return {
        groups: {
          [venueLabel]: {
            venueLabel,
            items: [
              { date: '2026-09-09', courseDisplay: 'KU4VACBBA200 - Disaster Management', studentCount: 28 },
              { date: '2026-09-11', courseDisplay: 'KU4SECBBA201 - Soft Skills & Personality Development', studentCount: 28 },
              { date: '2026-09-15', courseDisplay: 'KU4MDCBBA202 - Digital Marketing & E-Commerce', studentCount: 25 },
              { date: '2026-09-18', courseDisplay: 'KU4AECBBA203 - Business Ethics & Corporate Governance', studentCount: 28 }
            ]
          }
        },
        groupKeys: [venueLabel],
        totalCount: 1
      };
    }

    if (archetype === 'QP_COVER_LABEL') {
      const labels = [
        {
          venueCode: 'GA',
          venueName: 'Sree Narayana Guru College of Advanced Studies, Thottada',
          day: 'Wednesday',
          date: '2026-09-09',
          timeRange: '01:30 PM - 03:00 PM',
          subject: 'KU4VACBBA200 - Disaster Management',
          studentCount: '28',
          coverNumber: '1'
        },
        {
          venueCode: 'GA',
          venueName: 'Sree Narayana Guru College of Advanced Studies, Thottada',
          day: 'Friday',
          date: '2026-09-11',
          timeRange: '01:30 PM - 03:00 PM',
          subject: 'KU4SECBBA201 - Soft Skills & Personality Development',
          studentCount: '28',
          coverNumber: '2'
        }
      ];
      return {
        groups: { allLabels: labels },
        groupKeys: ['allLabels'],
        totalCount: labels.length
      };
    }

    return {
      groups: {
        default: [
          { 'Sl No': 1, 'Register No': '4PR24BB001', 'Candidate Name': 'ADITHYA K', 'Course': 'Disaster Management', 'Status': 'Registered' },
          { 'Sl No': 2, 'Register No': '4PR24BB002', 'Candidate Name': 'ANANYA RAJEEV', 'Course': 'Disaster Management', 'Status': 'Registered' }
        ]
      },
      groupKeys: ['default'],
      totalCount: 1
    };
  }, []);

  // --- DATA GROUPING AND STRUCTURING ENGINE --- //
  const processedData = useMemo(() => {
    const rows = dataset?.rows || [];
    if (!rows.length) {
      return getMockDataForArchetype(activeArchetype, currentConfig.formatCodeDotNameDot);
    }

    const getVal = (row, colKey, fallback = '') => {
      if (colKey && row[colKey] !== undefined && row[colKey] !== null) return String(row[colKey]).trim();
      return fallback;
    };

    if (activeArchetype === 'NOMINAL_ROLL') {
      const groups = {};
      const formatDots = currentConfig.formatCodeDotNameDot;

      rows.forEach(row => {
        const prog = getVal(row, columnMappings.programme, 'General Programme');
        const vCode = getVal(row, columnMappings.venueCode, '');
        const vName = getVal(row, columnMappings.venueName, 'Unassigned Venue');
        const seatNo = getVal(row, columnMappings.seatNo, 'N/A');
        const candName = getVal(row, columnMappings.name, 'Candidate');
        const cCode = getVal(row, columnMappings.courseCode, '');
        const cTitle = getVal(row, columnMappings.courseTitle, '');

        const venueLabel = vCode ? `${vCode} - ${vName}` : vName;
        const groupKey = currentConfig.groupBy === 'venue' ? venueLabel : `${prog} • ${venueLabel}`;

        if (!groups[groupKey]) {
          groups[groupKey] = {
            programme: prog,
            venueCode: vCode,
            venueName: vName,
            venueLabel,
            candidatesMap: {}
          };
        }

        if (!groups[groupKey].candidatesMap[seatNo]) {
          groups[groupKey].candidatesMap[seatNo] = {
            seatNo,
            studentName: candName,
            rawRow: row,
            courses: []
          };
        }

        if (cCode || cTitle) {
          let codeStr = cCode;
          let titleStr = cTitle;
          if (formatDots) {
            codeStr = codeStr.replace(/\.+$/, '').trim();
            titleStr = titleStr.replace(/\.+$/, '').trim();
            if (codeStr) codeStr += '.';
            if (titleStr) titleStr += '.';
          } else {
            codeStr = codeStr.replace(/\.+$/, '').trim();
            titleStr = titleStr.replace(/\.+$/, '').trim();
          }

          const display = codeStr && titleStr ? `${codeStr} - ${titleStr}` : (codeStr || titleStr);
          if (!groups[groupKey].candidatesMap[seatNo].courses.some(c => c.display === display)) {
            groups[groupKey].candidatesMap[seatNo].courses.push({ code: codeStr, title: titleStr, display });
          }
        }
      });

      const finalized = {};
      Object.keys(groups).sort().forEach(gKey => {
        const g = groups[gKey];
        const candidateList = Object.values(g.candidatesMap).map(cand => ({
          ...cand,
          courses: cand.courses && cand.courses.length > 0 ? cand.courses : [{ display: '—' }]
        })).sort((a, b) => 
          a.seatNo.localeCompare(b.seatNo, undefined, { numeric: true, sensitivity: 'base' })
        );
        finalized[gKey] = {
          ...g,
          candidates: candidateList,
          totalCandidates: candidateList.length
        };
      });

      const groupKeys = Object.keys(finalized);
      return { groups: finalized, groupKeys, totalCount: groupKeys.length };
    }

    if (activeArchetype === 'QP_STATEMENT') {
      const groups = {};
      const formatDots = currentConfig.formatCodeDotNameDot;

      rows.forEach(row => {
        const vCode = getVal(row, columnMappings.venueCode, '');
        const vName = getVal(row, columnMappings.venueName, 'Unassigned Venue');
        const venueLabel = vCode ? `${vCode} - ${vName}` : vName;
        const examDate = getVal(row, columnMappings.examDate, 'N/A');
        let cCode = getVal(row, columnMappings.courseCode, '');
        let cTitle = getVal(row, columnMappings.courseTitle, '');
        const countVal = parseInt(getVal(row, columnMappings.count, '1'), 10) || 1;

        if (formatDots) {
          cCode = cCode.replace(/\.+$/, '').trim();
          cTitle = cTitle.replace(/\.+$/, '').trim();
          if (cCode) cCode += '.';
          if (cTitle) cTitle += '.';
        } else {
          cCode = cCode.replace(/\.+$/, '').trim();
          cTitle = cTitle.replace(/\.+$/, '').trim();
        }

        const courseDisplay = cCode && cTitle ? `${cCode} - ${cTitle}` : (cCode || cTitle || 'Course');
        const groupKey = venueLabel;

        if (!groups[groupKey]) {
          groups[groupKey] = {
            venueLabel,
            items: []
          };
        }

        const existing = groups[groupKey].items.find(it => it.date === examDate && it.courseDisplay === courseDisplay);
        if (existing) {
          existing.studentCount += countVal;
        } else {
          groups[groupKey].items.push({
            date: examDate,
            courseCode: cCode,
            courseTitle: cTitle,
            courseDisplay,
            studentCount: countVal
          });
        }
      });

      Object.keys(groups).forEach(k => {
        groups[k].items.sort((a, b) => a.date.localeCompare(b.date));
      });

      const groupKeys = Object.keys(groups).sort();
      return { groups, groupKeys, totalCount: groupKeys.length };
    }

    if (activeArchetype === 'QP_COVER_LABEL') {
      const labels = [];
      const seen = new Set();

      rows.forEach(row => {
        const vCode = getVal(row, columnMappings.venueCode, 'GA');
        const vName = getVal(row, columnMappings.venueName, 'Examination Centre');
        const examDate = getVal(row, columnMappings.examDate, '2026-09-09');
        let cCode = getVal(row, columnMappings.courseCode, 'COURSE01');
        let cTitle = getVal(row, columnMappings.courseTitle, 'Course Name');
        const sTime = getVal(row, columnMappings.startTime, '01:30 PM');
        const eTime = getVal(row, columnMappings.endTime, '03:00 PM');
        const countVal = getVal(row, columnMappings.count, '');

        let dayName = 'Wednesday';
        if (examDate && examDate !== 'N/A') {
          const d = new Date(examDate);
          if (!isNaN(d.getTime())) {
            dayName = d.toLocaleDateString('en-US', { weekday: 'long' });
          }
        }

        const subjectStr = cCode && cTitle ? `${cCode} - ${cTitle}` : (cCode || cTitle);
        const dedupeKey = `${vCode}|${vName}|${examDate}|${subjectStr}`;

        if (!seen.has(dedupeKey)) {
          seen.add(dedupeKey);
          labels.push({
            venueCode: vCode,
            venueName: vName,
            day: dayName,
            date: examDate,
            timeRange: sTime && eTime ? `${sTime} - ${eTime}` : (sTime || eTime || '01:30 PM - 03:00 PM'),
            subject: subjectStr,
            studentCount: countVal,
            coverNumber: ''
          });
        }
      });

      return { groups: { allLabels: labels }, groupKeys: ['allLabels'], totalCount: labels.length };
    }

    return { groups: { default: rows }, groupKeys: ['default'], totalCount: rows.length };
  }, [dataset, activeArchetype, currentConfig, columnMappings, getMockDataForArchetype]);

  const effectiveGroupKey = processedData.groupKeys[activePreviewIndex] || processedData.groupKeys[0] || '';
  const currentPreviewGroup = processedData.groups[effectiveGroupKey];

  // Helper to draw headers on PDF dynamically
  const drawDynamicHeaders = (doc, pageWidth) => {
    let startY = 12;

    // 1. Draw Logo if enabled
    if (currentConfig.logo?.show && currentConfig.logo?.src) {
      try {
        const logoW = currentConfig.logo.width || 18;
        const logoH = logoW;
        let logoX = (pageWidth / 2) - (logoW / 2);

        if (currentConfig.logo.position === 'left') {
          logoX = 16;
        } else if (currentConfig.logo.position === 'right') {
          logoX = pageWidth - 16 - logoW;
        }

        doc.addImage(currentConfig.logo.src, 'PNG', logoX, startY, logoW, logoH);
        if (currentConfig.logo.position === 'top' || currentConfig.logo.position === 'center') {
          startY += logoH + 3;
        }
      } catch (e) {
        console.warn('Could not draw logo:', e);
      }
    }

    // 2. Draw Each Custom Header Line
    (currentConfig.headersList || []).forEach(h => {
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

  // --- PDF GENERATION ENGINE --- //
  const generatePdfForArchetype = (groupData, targetDoc = null) => {
    const isLandscape = currentConfig.orientation === 'landscape';
    const doc = targetDoc || new jsPDF(isLandscape ? 'l' : 'p', 'mm', (currentConfig.pageSize || 'a4').toLowerCase());
    const pageWidth = doc.internal.pageSize.getWidth();

    if (activeArchetype === 'NOMINAL_ROLL') {
      const g = groupData;
      if (!g) return doc;

      const headerEndY = drawDynamicHeaders(doc, pageWidth);

      const progText = `Programme: ${g.programme || 'N/A'}`;
      const venueText = `Venue: ${g.venueLabel || 'N/A'}`;

      doc.setFontSize(9);
      doc.setFont('helvetica', 'bold');
      
      const boxW = pageWidth - 28;
      const progLines = doc.splitTextToSize(progText, boxW - 10);
      const venueLines = doc.splitTextToSize(venueText, boxW - 10);
      const boxHeight = Math.max(13, (progLines.length + venueLines.length) * 4.2 + 4);

      doc.setDrawColor(180, 180, 180);
      doc.setFillColor(248, 249, 250);
      doc.roundedRect(14, headerEndY, boxW, boxHeight, 2, 2, 'FD');

      let textY = headerEndY + 4;
      doc.setTextColor(20, 20, 20);
      doc.text(progLines, 18, textY);
      textY += progLines.length * 4.2;
      
      doc.setTextColor(23, 107, 135);
      doc.text(venueLines, 18, textY);

      const cols = currentConfig.tableColumns || TEMPLATE_ARCHETYPES.NOMINAL_ROLL.defaultConfig.tableColumns;
      const tableBody = [];

      g.candidates.forEach((cand, candIdx) => {
        const cCount = Math.max(1, cand.courses.length);
        cand.courses.forEach((crs, crsIdx) => {
          const rowCells = [];
          cols.forEach(col => {
            let content = '';
            const styles = { halign: col.align || 'left', valign: 'middle' };
            if (col.bold) styles.fontStyle = 'bold';

            if (col.field === 'slNo') content = String(candIdx + 1);
            else if (col.field === 'seatNo') content = cand.seatNo;
            else if (col.field === 'name') content = cand.studentName;
            else if (col.field === 'courses') content = crs.display;
            else if (col.field === 'remarks' || col.field === 'blank') content = '';
            else if (cand.rawRow && cand.rawRow[col.field]) content = String(cand.rawRow[col.field]);

            if (col.field === 'courses' || !col.isSpan) {
              rowCells.push({ content, styles });
            } else {
              if (crsIdx === 0) {
                rowCells.push({ content, rowSpan: cCount, styles });
              }
            }
          });
          tableBody.push(rowCells);
        });
      });

      const [hBgR, hBgG, hBgB] = hexToRgb(currentConfig.tableTheme?.headerBg || '#f1f5f9');
      const [hTxtR, hTxtG, hTxtB] = hexToRgb(currentConfig.tableTheme?.headerColor || '#000000');
      const [bdrR, bdrG, bdrB] = hexToRgb(currentConfig.tableTheme?.borderColor || '#64748b');

      const colStylesObj = {};
      cols.forEach((c, i) => {
        colStylesObj[i] = {
          cellWidth: isLandscape ? Math.round(c.width * 1.3) : c.width,
          halign: c.align || 'left',
          valign: 'middle'
        };
        if (c.bold) colStylesObj[i].fontStyle = 'bold';
      });

      autoTable(doc, {
        startY: headerEndY + boxHeight + 4,
        head: [cols.map(c => c.label)],
        body: tableBody,
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: currentConfig.tableTheme?.fontSize || 8, valign: 'middle', cellPadding: { top: 1.5, bottom: 1.5, left: 2.5, right: 2.5 }, minCellHeight: 5, textColor: [0, 0, 0], lineColor: [bdrR, bdrG, bdrB], lineWidth: 0.18 },
        headStyles: { fillColor: [hBgR, hBgG, hBgB], textColor: [hTxtR, hTxtG, hTxtB], fontSize: 8.5, fontStyle: 'bold', halign: 'center', valign: 'middle', cellPadding: { top: 2.5, bottom: 2.5, left: 2, right: 2 }, lineColor: [bdrR, bdrG, bdrB], lineWidth: 0.18 },
        columnStyles: colStylesObj
      });
      return doc;
    }

    if (activeArchetype === 'QP_STATEMENT') {
      const g = groupData;
      if (!g) return doc;

      const headerEndY = drawDynamicHeaders(doc, pageWidth);

      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text(`${currentConfig.centerPrefix} ${g.venueLabel}`, pageWidth / 2, headerEndY, { align: 'center' });

      const cols = currentConfig.tableColumns || TEMPLATE_ARCHETYPES.QP_STATEMENT.defaultConfig.tableColumns;
      const tableBody = g.items.map((it, idx) => {
        return cols.map(col => {
          if (col.field === 'slNo') return idx + 1;
          if (col.field === 'date') return it.date;
          if (col.field === 'course') return it.courseDisplay;
          if (col.field === 'count') return it.studentCount;
          if (col.field === 'qp' || col.field === 'lp' || col.field === 'blank') return '';
          return it[col.field] || '';
        });
      });

      const [hBgR, hBgG, hBgB] = hexToRgb(currentConfig.tableTheme?.headerBg || '#f1f5f9');
      const [hTxtR, hTxtG, hTxtB] = hexToRgb(currentConfig.tableTheme?.headerColor || '#000000');
      const [bdrR, bdrG, bdrB] = hexToRgb(currentConfig.tableTheme?.borderColor || '#000000');

      const colStylesObj = {};
      cols.forEach((c, i) => {
        colStylesObj[i] = {
          cellWidth: isLandscape ? Math.round(c.width * 1.3) : c.width,
          halign: c.align || 'left',
          valign: 'middle'
        };
        if (c.bold) colStylesObj[i].fontStyle = 'bold';
      });

      autoTable(doc, {
        startY: headerEndY + 6,
        head: [cols.map(c => c.label)],
        body: tableBody,
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: currentConfig.tableTheme?.fontSize || 8.5, valign: 'middle', cellPadding: 3, textColor: [0, 0, 0], lineColor: [bdrR, bdrG, bdrB], lineWidth: 0.2 },
        headStyles: { fillColor: [hBgR, hBgG, hBgB], textColor: [hTxtR, hTxtG, hTxtB], fontStyle: 'bold', halign: 'center', valign: 'middle', fontSize: 9, lineColor: [bdrR, bdrG, bdrB], lineWidth: 0.3 },
        columnStyles: colStylesObj
      });
      return doc;
    }

    if (activeArchetype === 'QP_COVER_LABEL') {
      const lbl = groupData;
      if (!lbl) return doc;

      const headerEndY = drawDynamicHeaders(doc, pageWidth);

      const boxBody = [
        [
          { content: currentConfig.centerPrefix, styles: { fontStyle: 'bold', cellWidth: 42 } },
          { content: lbl.venueCode || '', styles: { fontStyle: 'bold', halign: 'center', cellWidth: 16 } },
          { content: lbl.venueName || '', colSpan: 4, styles: { fontStyle: 'bold' } }
        ],
        [
          { content: 'DAY', styles: { fontStyle: 'bold', cellWidth: 20 } },
          { content: lbl.day || '', styles: { cellWidth: 32 } },
          { content: 'DATE', styles: { fontStyle: 'bold', cellWidth: 18 } },
          { content: lbl.date || '', styles: { cellWidth: 36 } },
          { content: 'TIME', styles: { fontStyle: 'bold', cellWidth: 16 } },
          { content: lbl.timeRange || '', styles: { cellWidth: 48 } }
        ],
        [
          { content: 'SUBJECT', styles: { fontStyle: 'bold', cellWidth: 20 } },
          { content: lbl.subject || '', colSpan: 5, styles: { fontStyle: 'bold' } }
        ],
        [
          { content: 'NO. OF COPIES', styles: { fontStyle: 'bold', cellWidth: 42 } },
          { content: lbl.studentCount || '', colSpan: 2, styles: { fontStyle: 'bold', halign: 'center' } },
          { content: 'COVER NUMBER', styles: { fontStyle: 'bold', cellWidth: 42 } },
          { content: lbl.coverNumber || '', colSpan: 2, styles: { fontStyle: 'bold', halign: 'center' } }
        ]
      ];

      autoTable(doc, {
        startY: headerEndY + 2,
        body: boxBody,
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: 8.5, cellPadding: 3, textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.35, valign: 'middle' },
        columnStyles: {
          0: { fontStyle: 'bold' }
        }
      });

      const finalY = doc.lastAutoTable.finalY || 100;
      doc.setLineDashPattern([2, 2], 0);
      doc.setDrawColor(120, 120, 120);
      doc.line(14, finalY + 5, pageWidth - 14, finalY + 5);
      doc.line(14, finalY + 7, pageWidth - 14, finalY + 7);
      doc.setLineDashPattern([], 0);

      doc.setFontSize(12);
      doc.setFont('helvetica', 'bold');
      doc.text('CERTIFICATE', pageWidth / 2, finalY + 16, { align: 'center' });

      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      const certLines = doc.splitTextToSize(currentConfig.certificateText, pageWidth - 32);
      doc.text(certLines, 16, finalY + 24);

      let sigY = finalY + 44;
      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'bold');
      doc.text('INVIGILATOR', 16, sigY);
      doc.text('ADDL.CHIEF SUPERINTENDENT', pageWidth - 70, sigY);

      sigY += 9;
      doc.setFont('helvetica', 'normal');
      doc.text('(1) ___________________________________', 16, sigY);
      sigY += 7;
      doc.text('(2) ___________________________________', 16, sigY);

      sigY += 12;
      doc.setFont('helvetica', 'bold');
      doc.text('PLACE:', 16, sigY);
      sigY += 7;
      doc.text('DATE:', 16, sigY);
      doc.text('CHIEF SUPERINTENDENT', pageWidth - 65, sigY);

      return doc;
    }

    return doc;
  };

  // --- EXPORT HANDLERS --- //
  const handleDownloadSinglePdf = () => {
    if (!currentPreviewGroup) return;
    const doc = generatePdfForArchetype(
      activeArchetype === 'QP_COVER_LABEL' 
        ? processedData.groups.allLabels[activePreviewIndex] 
        : currentPreviewGroup
    );
    const safeName = templateName.replace(/[^a-zA-Z0-9_-]/g, '_');
    doc.save(`${safeName}_Page_${activePreviewIndex + 1}.pdf`);
  };

  const handleDownloadConsolidatedPdf = () => {
    if (!processedData.groupKeys.length) return;
    setIsProcessing(true);

    try {
      const isLandscape = currentConfig.orientation === 'landscape';
      const doc = new jsPDF(isLandscape ? 'l' : 'p', 'mm', (currentConfig.pageSize || 'a4').toLowerCase());

      if (activeArchetype === 'QP_COVER_LABEL') {
        const labels = processedData.groups.allLabels || [];
        labels.forEach((lbl, idx) => {
          if (idx > 0) doc.addPage();
          generatePdfForArchetype(lbl, doc);
        });
      } else {
        processedData.groupKeys.forEach((gKey, idx) => {
          if (idx > 0) doc.addPage();
          generatePdfForArchetype(processedData.groups[gKey], doc);
        });
      }

      const safeName = templateName.replace(/[^a-zA-Z0-9_-]/g, '_');
      doc.save(`Master_Consolidated_${safeName}.pdf`);
    } catch (err) {
      console.error(err);
      alert(`Error: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownloadZip = async () => {
    if (!processedData.groupKeys.length) return;
    setIsProcessing(true);

    try {
      const zip = new JSZip();
      const sanitize = (str) => String(str || '').replace(/[/\\?%*:|"<>•]/g, '_').replace(/\s+/g, ' ').trim();

      if (activeArchetype === 'QP_COVER_LABEL') {
        const labels = processedData.groups.allLabels || [];
        labels.forEach((lbl, idx) => {
          const doc = generatePdfForArchetype(lbl);
          const pdfBlob = doc.output('blob');
          const safeVenue = sanitize(lbl.venueName).slice(0, 40);
          const safeSubject = sanitize(lbl.subject).slice(0, 40);
          zip.file(`QP_Cover_Labels/${safeVenue}/${lbl.date}_${safeSubject}_Label_${idx + 1}.pdf`, pdfBlob);
        });
      } else if (activeArchetype === 'NOMINAL_ROLL') {
        processedData.groupKeys.forEach((gKey) => {
          const g = processedData.groups[gKey];
          const doc = generatePdfForArchetype(g);
          const pdfBlob = doc.output('blob');
          const safeProg = sanitize(g.programme).slice(0, 50);
          const safeVenue = sanitize(g.venueLabel).slice(0, 60);
          zip.file(`Individual_PDFs/${safeProg}/${safeVenue}_Nominal_Roll.pdf`, pdfBlob);
        });
      } else {
        processedData.groupKeys.forEach((gKey) => {
          const g = processedData.groups[gKey];
          const doc = generatePdfForArchetype(g);
          const pdfBlob = doc.output('blob');
          const safeVenue = sanitize(g.venueLabel || gKey).slice(0, 60);
          zip.file(`Statements/${safeVenue}_Statement.pdf`, pdfBlob);
        });
      }

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

  const handleSaveTemplate = async () => {
    if (!templateName.trim()) return alert('Please enter a template name.');
    
    const templateData = {
      name: templateName,
      archetype: activeArchetype,
      config: currentConfig,
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

  const activeCols = currentConfig.tableColumns || TEMPLATE_ARCHETYPES[activeArchetype]?.defaultConfig?.tableColumns || [];

  return (
    <div style={{ padding: '24px 32px 80px', maxWidth: '1600px', margin: '0 auto', fontFamily: 'var(--font-family)' }}>
      
      {/* Top Title & Archetype Switcher Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '16px' }}>
        <div>
          <h1 style={{ margin: 0, fontSize: '24px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '10px' }}>
            <Sparkles size={24} color="var(--accent)" /> Report Template Studio
          </h1>
          <p style={{ margin: '4px 0 0 0', color: 'var(--muted)', fontSize: '13.5px' }}>
            Create and customize report templates with Rich Text headers, customizable table columns, logo upload, and dynamic Excel column binding.
          </p>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
          <input 
            type="text"
            value={templateName}
            onChange={(e) => setTemplateName(e.target.value)}
            placeholder="Template Name..."
            style={{ padding: '8px 14px', fontSize: '13px', fontWeight: 600, borderRadius: '8px', border: '1.5px solid var(--line)', width: '240px' }}
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

      {/* Preset Archetypes Selector Bar */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))', gap: '12px', marginBottom: '24px' }}>
        {Object.values(TEMPLATE_ARCHETYPES).map((arch) => {
          const isSelected = activeArchetype === arch.id;
          return (
            <div
              key={arch.id}
              onClick={() => {
                setActiveArchetype(arch.id);
                setActivePreviewIndex(0);
              }}
              style={{
                padding: '16px',
                borderRadius: '10px',
                border: isSelected ? '2px solid var(--accent)' : '1px solid var(--line)',
                background: isSelected ? 'var(--accent-soft)' : 'var(--panel)',
                cursor: 'pointer',
                transition: 'all 0.2s ease',
                boxShadow: isSelected ? '0 4px 12px rgba(23,107,135,0.12)' : 'none'
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '6px' }}>
                <div style={{ 
                  background: isSelected ? 'var(--accent)' : '#e2e8f0', 
                  color: isSelected ? 'white' : '#475569', 
                  padding: '6px', 
                  borderRadius: '6px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center'
                }}>
                  {arch.id === 'NOMINAL_ROLL' && <FileText size={18} />}
                  {arch.id === 'QP_STATEMENT' && <CalendarRange size={18} />}
                  {arch.id === 'QP_COVER_LABEL' && <Tag size={18} />}
                  {arch.id === 'CUSTOM_TABULAR' && <TableProperties size={18} />}
                </div>
                <strong style={{ fontSize: '14.5px', color: isSelected ? 'var(--accent)' : 'var(--ink)' }}>{arch.name}</strong>
              </div>
              <p style={{ margin: 0, fontSize: '12px', color: 'var(--muted)', lineHeight: '1.4' }}>
                {arch.description}
              </p>
            </div>
          );
        })}
      </div>

      {/* Main Studio Two-Column Grid: Config Controls (Left) + WYSIWYG A4 Preview (Right) */}
      <div style={{ display: 'grid', gridTemplateColumns: isSidebarOpen ? '440px 1fr' : '0px 1fr', gap: isSidebarOpen ? '24px' : '0px', transition: 'all 0.3s ease' }}>
        
        {/* LEFT COLUMN: Customization Panels with Tabbed Controls */}
        <div style={{ display: isSidebarOpen ? 'block' : 'none' }}>
          
          {/* Sub-Tab Navigation for Studio Controls */}
          <div style={{ display: 'flex', background: 'var(--panel)', borderRadius: '8px', padding: '4px', marginBottom: '16px', gap: '4px', border: '1px solid var(--line)', flexWrap: 'wrap' }}>
            <button
              onClick={() => setActiveToolTab('columns')}
              style={{
                flex: 1,
                minWidth: '80px',
                padding: '8px 4px',
                fontSize: '11px',
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
                fontSize: '11px',
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
              onClick={() => setActiveToolTab('data')}
              style={{
                flex: 1,
                minWidth: '80px',
                padding: '8px 4px',
                fontSize: '11px',
                fontWeight: 700,
                borderRadius: '6px',
                border: 'none',
                cursor: 'pointer',
                background: activeToolTab === 'data' ? 'var(--accent)' : 'transparent',
                color: activeToolTab === 'data' ? 'white' : 'var(--ink)'
              }}
            >
              🔗 Data Binding
            </button>
            <button
              onClick={() => setActiveToolTab('page')}
              style={{
                flex: 1,
                minWidth: '80px',
                padding: '8px 4px',
                fontSize: '11px',
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
              onClick={() => setActiveToolTab('styling')}
              style={{
                flex: 1,
                minWidth: '80px',
                padding: '8px 4px',
                fontSize: '11px',
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

          {/* TAB: TABLE COLUMN HEADERS & LABELS (ADD / REMOVE / REORDER) */}
          {activeToolTab === 'columns' && (
            <div className="card" style={{ padding: '20px', marginBottom: '20px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
                <div>
                  <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <TableProperties size={16} /> Table Columns ({activeCols.length})
                  </h3>
                  <p style={{ margin: '2px 0 0', fontSize: '11.5px', color: 'var(--muted)' }}>
                    Add, remove, reorder, and rename columns printed in the report table.
                  </p>
                </div>
                <button
                  className="button"
                  onClick={addTableColumn}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', fontSize: '12px', padding: '5px 10px' }}
                >
                  <Plus size={14} /> Add Column
                </button>
              </div>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', maxHeight: '550px', overflowY: 'auto', paddingRight: '4px' }}>
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
                    <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr', gap: '8px' }}>
                      <div>
                        <label style={{ display: 'block', fontSize: '11px', fontWeight: 700, marginBottom: '3px' }}>Data Content:</label>
                        <select
                          value={col.field || 'blank'}
                          onChange={(e) => updateTableColumn(col.id, { field: e.target.value })}
                          style={{ width: '100%', fontSize: '12px' }}
                        >
                          <optgroup label="Standard Fields">
                            <option value="slNo">Serial No (1, 2, 3...)</option>
                            <option value="seatNo">Register / Seat No</option>
                            <option value="name">Candidate Name</option>
                            <option value="courses">Merged Courses (Sub-rows)</option>
                            <option value="date">Exam Date</option>
                            <option value="course">Course Name</option>
                            <option value="count">Candidate Count</option>
                            <option value="blank">Blank / Signature Box</option>
                          </optgroup>
                          {dataset.columns?.length > 0 && (
                            <optgroup label="Excel Columns">
                              {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
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

                  </div>
                ))}
              </div>
            </div>
          )}

          {/* TAB 1: DYNAMIC PAGE HEADERS & RICH TEXT BUILDER */}
          {activeToolTab === 'headers' && (
            <div className="card" style={{ padding: '20px', marginBottom: '20px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '14px' }}>
                <div>
                  <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Edit3 size={16} /> Page Title & Sub-Headers ({currentConfig.headersList?.length || 0})
                  </h3>
                  <p style={{ margin: '2px 0 0', fontSize: '11.5px', color: 'var(--muted)' }}>
                    Add, remove, reorder, and style institutional headers at the top of the report.
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
                {(currentConfig.headersList || []).map((h, idx) => (
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
                          disabled={idx === currentConfig.headersList.length - 1} 
                          onClick={() => moveHeaderLine(idx, 1)} 
                          style={{ background: 'none', border: 'none', cursor: idx === currentConfig.headersList.length - 1 ? 'default' : 'pointer', color: idx === currentConfig.headersList.length - 1 ? '#ccc' : 'var(--ink)', padding: '2px' }}
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

          {/* TAB 2: PAGE ORIENTATION & LOGO CONTROLS */}
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
                    onClick={() => updateCurrentConfig({ orientation: 'portrait' })}
                    style={{
                      padding: '10px',
                      borderRadius: '8px',
                      border: currentConfig.orientation === 'portrait' || !currentConfig.orientation ? '2px solid var(--accent)' : '1px solid var(--line)',
                      background: currentConfig.orientation === 'portrait' || !currentConfig.orientation ? 'var(--accent-soft)' : 'white',
                      fontWeight: 700,
                      fontSize: '13px',
                      color: currentConfig.orientation === 'portrait' || !currentConfig.orientation ? 'var(--accent)' : 'var(--ink)',
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
                    onClick={() => updateCurrentConfig({ orientation: 'landscape' })}
                    style={{
                      padding: '10px',
                      borderRadius: '8px',
                      border: currentConfig.orientation === 'landscape' ? '2px solid var(--accent)' : '1px solid var(--line)',
                      background: currentConfig.orientation === 'landscape' ? 'var(--accent-soft)' : 'white',
                      fontWeight: 700,
                      fontSize: '13px',
                      color: currentConfig.orientation === 'landscape' ? 'var(--accent)' : 'var(--ink)',
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
                  value={currentConfig.pageSize || 'A4'}
                  onChange={(e) => updateCurrentConfig({ pageSize: e.target.value })}
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
                    checked={!!currentConfig.logo?.show}
                    onChange={(e) => updateCurrentConfig({ logo: { ...currentConfig.logo, show: e.target.checked } })}
                    style={{ width: '16px', height: '16px', accentColor: 'var(--accent)' }}
                  />
                  Show Logo / Crest on Page
                </label>

                {currentConfig.logo?.show && (
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
                        onClick={() => updateCurrentConfig({ logo: { ...currentConfig.logo, src: logoBase64 } })}
                        style={{ fontSize: '11px', padding: '6px 10px' }}
                        title="Reset to default Kannur University Crest"
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
                            onClick={() => updateCurrentConfig({ logo: { ...currentConfig.logo, position: pos } })}
                            style={{
                              padding: '4px 8px',
                              fontSize: '11px',
                              textTransform: 'capitalize',
                              borderRadius: '4px',
                              border: '1px solid var(--line)',
                              background: currentConfig.logo?.position === pos ? 'var(--accent)' : 'white',
                              color: currentConfig.logo?.position === pos ? 'white' : 'var(--ink)',
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
                          value={currentConfig.logo?.width || 18}
                          onChange={(e) => updateCurrentConfig({ logo: { ...currentConfig.logo, width: parseInt(e.target.value, 10) } })}
                          style={{ width: '90px' }}
                        />
                        <span style={{ fontSize: '11px', fontWeight: 700 }}>{currentConfig.logo?.width || 18} mm</span>
                      </div>
                    </div>

                  </div>
                )}
              </div>

            </div>
          )}

          {/* TAB 3: DYNAMIC EXCEL DATA BINDING */}
          {activeToolTab === 'data' && (
            <div className="card" style={{ padding: '20px', marginBottom: '20px' }}>
              <h3 style={{ margin: '0 0 16px 0', fontSize: '15px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Layers size={16} /> Excel Spreadsheet Column Binding
              </h3>

              {dataset.columns.length === 0 ? (
                <div style={{ padding: '14px', background: '#f8fafc', borderRadius: '8px', border: '1px dashed #cbd5e1', fontSize: '13px', color: 'var(--muted)' }}>
                  Upload an Excel file in Report Config to bind your actual spreadsheet columns.
                </div>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                  
                  {activeArchetype === 'NOMINAL_ROLL' && (
                    <>
                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Register / Seat Number:</label>
                        <select 
                          value={columnMappings.seatNo || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, seatNo: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Candidate Name:</label>
                        <select 
                          value={columnMappings.name || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, name: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Course Code (Paper Code):</label>
                        <select 
                          value={columnMappings.courseCode || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, courseCode: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Course Title (Paper Name):</label>
                        <select 
                          value={columnMappings.courseTitle || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, courseTitle: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Programme Name:</label>
                        <select 
                          value={columnMappings.programme || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, programme: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Venue Name / Code:</label>
                        <select 
                          value={columnMappings.venueName || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, venueName: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>
                    </>
                  )}

                  {activeArchetype === 'QP_STATEMENT' && (
                    <>
                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Examination Date:</label>
                        <select 
                          value={columnMappings.examDate || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, examDate: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Course / Subject Code:</label>
                        <select 
                          value={columnMappings.courseCode || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, courseCode: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Course Title:</label>
                        <select 
                          value={columnMappings.courseTitle || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, courseTitle: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Candidate Count (NC):</label>
                        <select 
                          value={columnMappings.count || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, count: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Venue / Centre:</label>
                        <select 
                          value={columnMappings.venueName || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, venueName: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>
                    </>
                  )}

                  {activeArchetype === 'QP_COVER_LABEL' && (
                    <>
                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Centre Code:</label>
                        <select 
                          value={columnMappings.venueCode || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, venueCode: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Centre Name:</label>
                        <select 
                          value={columnMappings.venueName || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, venueName: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Exam Date:</label>
                        <select 
                          value={columnMappings.examDate || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, examDate: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>

                      <div>
                        <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Subject / Paper:</label>
                        <select 
                          value={columnMappings.courseTitle || columnMappings.courseCode || ''} 
                          onChange={(e) => setColumnMappings(p => ({ ...p, courseTitle: e.target.value }))}
                          style={{ width: '100%', fontSize: '13px' }}
                        >
                          <option value="">-- Select Column --</option>
                          {dataset.columns.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </div>
                    </>
                  )}

                </div>
              )}
            </div>
          )}

          {/* TAB 4: TABLE STYLING & FORMATTING */}
          {activeToolTab === 'styling' && (
            <div className="card" style={{ padding: '20px', marginBottom: '20px' }}>
              <h3 style={{ margin: '0 0 16px 0', fontSize: '15px', fontWeight: 800, color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Sliders size={16} /> Table & Text Formatting
              </h3>

              <div style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '13px', fontWeight: 600 }}>
                  <input 
                    type="checkbox"
                    checked={!!currentConfig.formatCodeDotNameDot}
                    onChange={(e) => updateCurrentConfig({ formatCodeDotNameDot: e.target.checked })}
                    style={{ width: '16px', height: '16px', accentColor: 'var(--accent)' }}
                  />
                  Format: CODE. - NAME. (with trailing dots)
                </label>

                {/* Table Header Color */}
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <span style={{ fontSize: '12px', fontWeight: 600 }}>Table Header Background:</span>
                  <input 
                    type="color"
                    value={currentConfig.tableTheme?.headerBg || '#f1f5f9'}
                    onChange={(e) => updateCurrentConfig({ tableTheme: { ...currentConfig.tableTheme, headerBg: e.target.value } })}
                    style={{ width: '32px', height: '32px', border: 'none', borderRadius: '4px', cursor: 'pointer' }}
                  />
                </div>

                {/* Table Header Text Color */}
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <span style={{ fontSize: '12px', fontWeight: 600 }}>Table Header Text Color:</span>
                  <input 
                    type="color"
                    value={currentConfig.tableTheme?.headerColor || '#000000'}
                    onChange={(e) => updateCurrentConfig({ tableTheme: { ...currentConfig.tableTheme, headerColor: e.target.value } })}
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
                      value={currentConfig.tableTheme?.fontSize || 8}
                      onChange={(e) => updateCurrentConfig({ tableTheme: { ...currentConfig.tableTheme, fontSize: parseFloat(e.target.value) || 8 } })}
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
          
          {/* Navigation & Export Toolbar */}
          <div className="card" style={{ padding: '16px 20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
            
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

            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
              <button 
                className="button secondary"
                onClick={handleDownloadSinglePdf}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '7px 14px' }}
              >
                <Download size={15} /> Page PDF
              </button>

              <button 
                className="button secondary"
                onClick={handleDownloadConsolidatedPdf}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '7px 14px' }}
              >
                <FileText size={15} /> Master PDF
              </button>

              <button 
                className="button"
                onClick={handleDownloadZip}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '7px 16px' }}
              >
                <Archive size={15} /> Individual PDFs (.zip)
              </button>

              <button 
                className="button secondary"
                onClick={() => window.print()}
                title="Direct Print A4"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '7px 14px' }}
              >
                <Printer size={15} /> Print A4
              </button>
            </div>

          </div>

          {/* Rendered Live Paper Display (Adapts to Portrait or Landscape) */}
          <div 
            className="card print-container" 
            style={{ 
              padding: currentConfig.orientation === 'landscape' ? '32px 40px' : '40px 48px', 
              background: 'white', 
              border: '1px solid #cbd5e1', 
              borderRadius: '8px', 
              boxShadow: '0 8px 30px rgba(0,0,0,0.08)', 
              minHeight: '750px',
              maxWidth: currentConfig.orientation === 'landscape' ? '1123px' : '794px',
              margin: '0 auto',
              width: '100%',
              transition: 'all 0.3s ease'
            }}
          >
            
            {/* DYNAMIC HEADER RENDERING ON LIVE PAPER */}
            <div style={{ textAlign: 'center', marginBottom: '20px', position: 'relative' }}>
              
              {/* Optional Logo */}
              {currentConfig.logo?.show && currentConfig.logo?.src && (
                <div style={{
                  display: 'flex',
                  justifyContent: currentConfig.logo.position === 'left' ? 'flex-start' : currentConfig.logo.position === 'right' ? 'flex-end' : 'center',
                  marginBottom: '10px'
                }}>
                  <img 
                    src={currentConfig.logo.src} 
                    alt="Logo" 
                    style={{ height: `${(currentConfig.logo.width || 18) * 3}px`, objectFit: 'contain' }} 
                  />
                </div>
              )}

              {/* Dynamic Headers List */}
              {(currentConfig.headersList || []).map((h, idx) => (
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

              {activeArchetype === 'QP_STATEMENT' && currentPreviewGroup && (
                <div style={{ fontSize: '12pt', fontWeight: 800, color: '#000', marginTop: '6px' }}>
                  {currentConfig.centerPrefix} {effectiveGroupKey}
                </div>
              )}
            </div>

            {/* ARCHETYPE 1: NOMINAL ROLL LIVE VIEW */}
            {activeArchetype === 'NOMINAL_ROLL' && (
              <div>
                {currentPreviewGroup && (
                  <div style={{ background: '#f8fafc', border: '1px solid #cbd5e1', borderRadius: '8px', padding: '12px 18px', marginBottom: '20px' }}>
                    <div style={{ fontSize: '13.5px', fontWeight: 800, color: '#0f172a', marginBottom: '4px' }}>
                      Programme: {currentPreviewGroup.programme || 'General Programme'}
                    </div>
                    <div style={{ fontSize: '13.5px', fontWeight: 800, color: '#0284c7' }}>
                      Venue: {currentPreviewGroup.venueLabel || 'Examination Venue'}
                    </div>
                  </div>
                )}

                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: `${currentConfig.tableTheme?.fontSize || 8}pt` }}>
                  <thead>
                    <tr style={{ background: currentConfig.tableTheme?.headerBg || '#f1f5f9', color: currentConfig.tableTheme?.headerColor || '#000000', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000' }}>
                      {activeCols.map(c => (
                        <th 
                          key={c.id} 
                          style={{ 
                            border: `1px solid ${currentConfig.tableTheme?.borderColor || '#64748b'}`, 
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
                    {currentPreviewGroup?.candidates?.map((cand, candIdx) => {
                      const courses = cand.courses || [];
                      return courses.map((crs, crsIdx) => (
                        <tr key={`${candIdx}_${crsIdx}`} style={{ borderBottom: `1px solid ${currentConfig.tableTheme?.borderColor || '#64748b'}` }}>
                          {activeCols.map(col => {
                            let content = '';
                            if (col.field === 'slNo') content = candIdx + 1;
                            else if (col.field === 'seatNo') content = cand.seatNo;
                            else if (col.field === 'name') content = cand.studentName;
                            else if (col.field === 'courses') content = crs.display;
                            else if (col.field === 'remarks' || col.field === 'blank') content = '';
                            else if (cand.rawRow && cand.rawRow[col.field]) content = cand.rawRow[col.field];

                            if (col.field === 'courses' || !col.isSpan) {
                              return (
                                <td 
                                  key={col.id} 
                                  style={{ 
                                    border: `1px solid ${currentConfig.tableTheme?.borderColor || '#64748b'}`, 
                                    padding: '5px 8px', 
                                    verticalAlign: 'middle',
                                    textAlign: col.align || 'left',
                                    fontWeight: col.bold ? 700 : 400
                                  }}
                                >
                                  {content}
                                </td>
                              );
                            } else {
                              if (crsIdx === 0) {
                                return (
                                  <td 
                                    key={col.id} 
                                    rowSpan={courses.length} 
                                    style={{ 
                                      border: `1px solid ${currentConfig.tableTheme?.borderColor || '#64748b'}`, 
                                      padding: '6px 8px', 
                                      textAlign: col.align || 'left', 
                                      verticalAlign: 'middle', 
                                      fontWeight: col.bold || col.field === 'seatNo' || col.field === 'name' ? 700 : 400 
                                    }}
                                  >
                                    {content}
                                  </td>
                                );
                              }
                              return null;
                            }
                          })}
                        </tr>
                      ));
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* ARCHETYPE 2: QP STATEMENT LIVE VIEW */}
            {activeArchetype === 'QP_STATEMENT' && (
              <div>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: `${currentConfig.tableTheme?.fontSize || 8.5}pt` }}>
                  <thead>
                    <tr style={{ background: currentConfig.tableTheme?.headerBg || '#f1f5f9', color: currentConfig.tableTheme?.headerColor || '#000000', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000' }}>
                      {activeCols.map(c => (
                        <th key={c.id} style={{ border: `1px solid ${currentConfig.tableTheme?.borderColor || '#000000'}`, padding: '8px 6px', textAlign: c.align || 'left', fontWeight: 700 }}>
                          {c.label}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {currentPreviewGroup?.items?.map((it, idx) => (
                      <tr key={idx} style={{ borderBottom: `1px solid ${currentConfig.tableTheme?.borderColor || '#000000'}` }}>
                        {activeCols.map(col => {
                          let content;
                          if (col.field === 'slNo') content = idx + 1;
                          else if (col.field === 'date') content = it.date;
                          else if (col.field === 'course') content = it.courseDisplay;
                          else if (col.field === 'count') content = it.studentCount;
                          else if (col.field === 'qp' || col.field === 'lp' || col.field === 'blank') content = '';
                          else content = it[col.field] || '';

                          return (
                            <td key={col.id} style={{ border: `1px solid ${currentConfig.tableTheme?.borderColor || '#000000'}`, padding: '6px 8px', textAlign: col.align || 'left', fontWeight: col.bold || col.field === 'count' ? 700 : 400 }}>
                              {content}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {/* ARCHETYPE 3: QP COVER ENVELOPE LABEL LIVE VIEW */}
            {activeArchetype === 'QP_COVER_LABEL' && (() => {
              const lbl = processedData.groups.allLabels?.[activePreviewIndex] || {
                venueCode: 'GA',
                venueName: 'Sree Narayana Guru College of Advanced Studies, Thottada',
                day: 'Wednesday',
                date: '2026-09-09',
                timeRange: '01:30 PM - 03:00 PM',
                subject: 'KU4VACBBA200 - Disaster Management',
                studentCount: '',
                coverNumber: ''
              };

              return (
                <div style={{ maxWidth: '800px', margin: '0 auto' }}>
                  {/* Border Table Box */}
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', border: '1.5px solid #000', marginBottom: '20px' }}>
                    <tbody>
                      <tr style={{ borderBottom: '1px solid #000' }}>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px', fontWeight: 800, width: '180px' }}>{currentConfig.centerPrefix}</td>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px', fontWeight: 800, textAlign: 'center', width: '60px' }}>{lbl.venueCode}</td>
                        <td colSpan={4} style={{ padding: '8px 12px', fontWeight: 800 }}>{lbl.venueName}</td>
                      </tr>
                      <tr style={{ borderBottom: '1px solid #000' }}>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px', fontWeight: 800 }}>DAY</td>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px' }}>{lbl.day}</td>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px', fontWeight: 800, width: '60px' }}>DATE</td>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px' }}>{lbl.date}</td>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px', fontWeight: 800, width: '60px' }}>TIME</td>
                        <td style={{ padding: '8px 10px' }}>{lbl.timeRange}</td>
                      </tr>
                      <tr style={{ borderBottom: '1px solid #000' }}>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px', fontWeight: 800 }}>SUBJECT</td>
                        <td colSpan={5} style={{ padding: '8px 12px', fontWeight: 800 }}>{lbl.subject}</td>
                      </tr>
                      <tr>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px', fontWeight: 800 }}>NO. OF COPIES</td>
                        <td colSpan={2} style={{ borderRight: '1px solid #000', padding: '8px 10px', textAlign: 'center', fontWeight: 700 }}>{lbl.studentCount}</td>
                        <td style={{ borderRight: '1px solid #000', padding: '8px 10px', fontWeight: 800 }}>COVER NUMBER</td>
                        <td colSpan={2} style={{ padding: '8px 10px', textAlign: 'center', fontWeight: 700 }}>{lbl.coverNumber}</td>
                      </tr>
                    </tbody>
                  </table>

                  {/* Certificate Section */}
                  <div style={{ borderTop: '2px dashed #94a3b8', paddingTop: '16px', textAlign: 'center', marginBottom: '14px' }}>
                    <div style={{ fontSize: '16px', fontWeight: 800, color: '#000', letterSpacing: '1px', marginBottom: '12px' }}>
                      CERTIFICATE
                    </div>
                    <p style={{ fontSize: '13px', lineHeight: '1.6', textAlign: 'left', margin: '0 0 30px 0', color: '#111' }}>
                      {currentConfig.certificateText}
                    </p>
                  </div>

                  {/* Signatures Row */}
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', fontSize: '12.5px', fontWeight: 700, color: '#000', marginTop: '20px' }}>
                    <div>
                      <div style={{ marginBottom: '10px' }}>INVIGILATOR</div>
                      <div style={{ fontWeight: 500, marginBottom: '6px' }}>(1) ____________________________________</div>
                      <div style={{ fontWeight: 500 }}>(2) ____________________________________</div>
                    </div>
                    <div style={{ textAlign: 'right' }}>
                      <div>ADDL.CHIEF SUPERINTENDENT</div>
                    </div>
                  </div>

                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-end', fontSize: '12.5px', fontWeight: 700, color: '#000', marginTop: '30px' }}>
                    <div>
                      <div style={{ marginBottom: '6px' }}>PLACE:</div>
                      <div>DATE:</div>
                    </div>
                    <div style={{ textAlign: 'right' }}>
                      <div>CHIEF SUPERINTENDENT</div>
                    </div>
                  </div>

                </div>
              );
            })()}

            {/* ARCHETYPE 4: CUSTOM TABULAR LIVE VIEW */}
            {activeArchetype === 'CUSTOM_TABULAR' && (
              <div>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: `${currentConfig.tableTheme?.fontSize || 8.5}pt` }}>
                  <thead>
                    <tr style={{ background: currentConfig.tableTheme?.headerBg || '#f1f5f9', color: currentConfig.tableTheme?.headerColor || '#000000', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000' }}>
                      {activeCols.map(c => (
                        <th key={c.id} style={{ border: `1px solid ${currentConfig.tableTheme?.borderColor || '#64748b'}`, padding: '8px 6px', textAlign: c.align || 'left' }}>
                          {c.label}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {dataset.rows.slice(0, 25).map((row, idx) => (
                      <tr key={idx} style={{ borderBottom: `1px solid ${currentConfig.tableTheme?.borderColor || '#64748b'}` }}>
                        {activeCols.map(col => {
                          let content;
                          if (col.field === 'slNo') content = idx + 1;
                          else if (col.field === 'seatNo') content = row[columnMappings.seatNo] || row['seatNo'] || '';
                          else if (col.field === 'name') content = row[columnMappings.name] || row['name'] || '';
                          else if (col.field === 'blank') content = '';
                          else content = row[col.field] || '';

                          return (
                            <td key={col.id} style={{ border: `1px solid ${currentConfig.tableTheme?.borderColor || '#64748b'}`, padding: '6px 8px', textAlign: col.align || 'left' }}>
                              {content}
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

          </div>

        </div>

      </div>

    </div>
  );
};

export default TemplatePage;
