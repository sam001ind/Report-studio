import { useState, useEffect, useMemo } from 'react';
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
  ChevronRight
} from 'lucide-react';
import { supabase } from '../supabaseClient';
import { useAuth } from '../context/AuthContext';
import { logoBase64 } from '../assets/logoBase64';
import { TEMPLATE_ARCHETYPES, autoDetectDatasetColumns, suggestArchetype } from '../utils/templateEngine';

const TemplatePage = ({ dataset = { columns: [], rows: [] }, initialTemplate }) => {
  const { user } = useAuth();
  
  // Selected Archetype / Mode: 'NOMINAL_ROLL' | 'QP_STATEMENT' | 'QP_COVER_LABEL' | 'CUSTOM_TABULAR'
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
      NOMINAL_ROLL: { ...TEMPLATE_ARCHETYPES.NOMINAL_ROLL.defaultConfig },
      QP_STATEMENT: { ...TEMPLATE_ARCHETYPES.QP_STATEMENT.defaultConfig },
      QP_COVER_LABEL: { ...TEMPLATE_ARCHETYPES.QP_COVER_LABEL.defaultConfig },
      CUSTOM_TABULAR: { ...TEMPLATE_ARCHETYPES.CUSTOM_TABULAR.defaultConfig }
    };
  });

  // Column Mappings
  const [columnMappings, setColumnMappings] = useState(() => autoDetectDatasetColumns(dataset?.columns || []));

  const [isSidebarOpen, _setIsSidebarOpen] = useState(true);
  const [templateName, setTemplateName] = useState(initialTemplate?.name || 'My Custom University Report');
  const [isProcessing, setIsProcessing] = useState(false);
  const [_statusMsg, setStatusMsg] = useState('Ready');
  const [activePreviewIndex, setActivePreviewIndex] = useState(0);

  // Update column detection when dataset changes
  useEffect(() => {
    if (dataset?.columns?.length > 0) {
      const detected = autoDetectDatasetColumns(dataset.columns);
      setColumnMappings(detected);
      
      // Auto-populate session name if found
      if (detected.session && dataset.rows?.length > 0) {
        const sampleSess = dataset.rows.find(r => r[detected.session])?.[detected.session];
        if (sampleSess && typeof sampleSess === 'string' && sampleSess.trim()) {
          setArchetypeConfigs(prev => ({
            ...prev,
            NOMINAL_ROLL: { ...prev.NOMINAL_ROLL, sessionName: sampleSess.trim() },
            QP_STATEMENT: { ...prev.QP_STATEMENT, sessionName: sampleSess.trim() },
            QP_COVER_LABEL: { ...prev.QP_COVER_LABEL, examTitle: `Fourth Semester Degree (Private Registration) Regular Examinations ${sampleSess.trim()}` }
          }));
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

  // --- DATA GROUPING AND STRUCTURING ENGINE --- //
  const processedData = useMemo(() => {
    const rows = dataset?.rows || [];
    if (!rows.length) return { groups: {}, groupKeys: [], totalCount: 0 };

    const getVal = (row, colKey, fallback = '') => {
      if (!colKey || !row[colKey]) return fallback;
      return String(row[colKey]).trim();
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
          // Deduplicate courses
          if (!groups[groupKey].candidatesMap[seatNo].courses.some(c => c.display === display)) {
            groups[groupKey].candidatesMap[seatNo].courses.push({ code: codeStr, title: titleStr, display });
          }
        }
      });

      const finalized = {};
      Object.keys(groups).sort().forEach(gKey => {
        const g = groups[gKey];
        const candidateList = Object.values(g.candidatesMap).sort((a, b) => 
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

        // Aggregate same date & course or push
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

      // Sort items by date
      Object.keys(groups).forEach(k => {
        groups[k].items.sort((a, b) => a.date.localeCompare(b.date));
      });

      const groupKeys = Object.keys(groups).sort();
      return { groups, groupKeys, totalCount: groupKeys.length };
    }

    if (activeArchetype === 'QP_COVER_LABEL') {
      // Each unique Venue + Date + Course combination generates 1 Label
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

    // CUSTOM_TABULAR
    return { groups: { default: rows }, groupKeys: ['default'], totalCount: rows.length };
  }, [dataset, activeArchetype, currentConfig, columnMappings]);

  // Current Active Preview Data
  const effectiveGroupKey = processedData.groupKeys[activePreviewIndex] || processedData.groupKeys[0] || '';
  const currentPreviewGroup = processedData.groups[effectiveGroupKey];

  // --- PDF GENERATION ENGINE --- //
  const generatePdfForArchetype = (groupData, targetDoc = null) => {
    const doc = targetDoc || new jsPDF('p', 'mm', 'a4');

    if (activeArchetype === 'NOMINAL_ROLL') {
      const g = groupData;
      if (!g) return doc;

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.text(currentConfig.universityName, 105, 14, { align: 'center' });

      doc.setFontSize(10.5);
      doc.setFont('helvetica', 'bold');
      doc.text(currentConfig.branchName, 105, 19.5, { align: 'center' });

      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.text(currentConfig.examTitle, 105, 25, { align: 'center' });

      if (currentConfig.sessionName) {
        doc.setFontSize(9.5);
        doc.setFont('helvetica', 'normal');
        doc.text(currentConfig.sessionName, 105, 30, { align: 'center' });
      }

      const progText = `Programme: ${g.programme || 'N/A'}`;
      const venueText = `Venue: ${g.venueLabel || 'N/A'}`;

      doc.setFontSize(9);
      doc.setFont('helvetica', 'bold');
      
      const progLines = doc.splitTextToSize(progText, 172);
      const venueLines = doc.splitTextToSize(venueText, 172);
      const boxHeight = Math.max(13, (progLines.length + venueLines.length) * 4.2 + 4);

      doc.setDrawColor(180, 180, 180);
      doc.setFillColor(248, 249, 250);
      doc.roundedRect(14, 34, 182, boxHeight, 2, 2, 'FD');

      let textY = 38;
      doc.setTextColor(20, 20, 20);
      doc.text(progLines, 18, textY);
      textY += progLines.length * 4.2;
      
      doc.setTextColor(23, 107, 135);
      doc.text(venueLines, 18, textY);

      const tableBody = [];
      g.candidates.forEach((cand, candIdx) => {
        const cCount = Math.max(1, cand.courses.length);
        if (cand.courses.length === 0) {
          tableBody.push([
            { content: String(candIdx + 1), styles: { halign: 'center', valign: 'middle' } },
            { content: cand.seatNo, styles: { halign: 'center', valign: 'middle', fontStyle: 'bold' } },
            { content: cand.studentName, styles: { halign: 'left', valign: 'middle', fontStyle: 'bold' } },
            { content: '—', styles: { halign: 'left', valign: 'middle' } },
            { content: '', styles: { halign: 'center', valign: 'middle' } }
          ]);
        } else {
          cand.courses.forEach((crs, crsIdx) => {
            if (crsIdx === 0) {
              tableBody.push([
                { content: String(candIdx + 1), rowSpan: cCount, styles: { halign: 'center', valign: 'middle' } },
                { content: cand.seatNo, rowSpan: cCount, styles: { halign: 'center', valign: 'middle', fontStyle: 'bold' } },
                { content: cand.studentName, rowSpan: cCount, styles: { halign: 'left', valign: 'middle', fontStyle: 'bold' } },
                { content: crs.display, styles: { halign: 'left', valign: 'middle' } },
                { content: '', rowSpan: cCount, styles: { halign: 'center', valign: 'middle' } }
              ]);
            } else {
              tableBody.push([
                { content: crs.display, styles: { halign: 'left', valign: 'middle' } }
              ]);
            }
          });
        }
      });

      autoTable(doc, {
        startY: 34 + boxHeight + 4,
        head: [[currentConfig.headers.slNo, currentConfig.headers.regNo, currentConfig.headers.name, currentConfig.headers.courses, currentConfig.headers.remarks]],
        body: tableBody,
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: 8, valign: 'middle', cellPadding: { top: 1.5, bottom: 1.5, left: 2.5, right: 2.5 }, minCellHeight: 5, textColor: [0, 0, 0], lineColor: [100, 100, 100], lineWidth: 0.18 },
        headStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontSize: 8.5, fontStyle: 'bold', halign: 'center', valign: 'middle', cellPadding: { top: 2.5, bottom: 2.5, left: 2, right: 2 }, lineColor: [100, 100, 100], lineWidth: 0.18 },
        columnStyles: {
          0: { cellWidth: currentConfig.columnStyles.slNo, halign: 'center', valign: 'middle' },
          1: { cellWidth: currentConfig.columnStyles.regNo, fontStyle: 'bold', halign: 'center', valign: 'middle' },
          2: { cellWidth: currentConfig.columnStyles.name, fontStyle: 'bold', valign: 'middle' },
          3: { cellWidth: currentConfig.columnStyles.courses, valign: 'middle' },
          4: { cellWidth: currentConfig.columnStyles.remarks, valign: 'middle' }
        }
      });
      return doc;
    }

    if (activeArchetype === 'QP_STATEMENT') {
      const g = groupData;
      if (!g) return doc;

      doc.setFontSize(14);
      doc.setFont('helvetica', 'bold');
      doc.text(currentConfig.universityName, 105, 14, { align: 'center' });

      doc.setFontSize(10.5);
      doc.setFont('helvetica', 'bold');
      doc.text(currentConfig.branchName, 105, 19.5, { align: 'center' });

      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'bold');
      doc.text(currentConfig.examTitle, 105, 25, { align: 'center' });

      if (currentConfig.sessionName) {
        doc.setFontSize(9);
        doc.setFont('helvetica', 'bold');
        doc.text(currentConfig.sessionName, 105, 30, { align: 'center' });
      }

      doc.setFontSize(10);
      doc.setFont('helvetica', 'bold');
      doc.text(`${currentConfig.centerPrefix} ${g.venueLabel}`, 105, 36, { align: 'center' });

      const tableBody = g.items.map((it, idx) => [
        idx + 1,
        it.date,
        it.courseDisplay,
        it.studentCount,
        '',
        ''
      ]);

      autoTable(doc, {
        startY: 42,
        head: [[currentConfig.headers.slNo, currentConfig.headers.date, currentConfig.headers.course, currentConfig.headers.count, currentConfig.headers.qp, currentConfig.headers.lp]],
        body: tableBody,
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: 8.5, valign: 'middle', cellPadding: 3, textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.2 },
        headStyles: { fillColor: [240, 240, 240], textColor: [0, 0, 0], fontStyle: 'bold', halign: 'center', valign: 'middle', fontSize: 9, lineColor: [0, 0, 0], lineWidth: 0.3 },
        columnStyles: {
          0: { halign: 'center', cellWidth: currentConfig.columnStyles.slNo },
          1: { halign: 'left', cellWidth: currentConfig.columnStyles.date },
          2: { halign: 'left', cellWidth: currentConfig.columnStyles.course },
          3: { halign: 'center', cellWidth: currentConfig.columnStyles.count, fontStyle: 'bold' },
          4: { halign: 'center', cellWidth: currentConfig.columnStyles.qp },
          5: { halign: 'center', cellWidth: currentConfig.columnStyles.lp }
        }
      });
      return doc;
    }

    if (activeArchetype === 'QP_COVER_LABEL') {
      const lbl = groupData;
      if (!lbl) return doc;

      // University Header with Crest
      if (currentConfig.showLogo && logoBase64) {
        try {
          doc.addImage(logoBase64, 'PNG', 32, 12, 18, 18);
        } catch (e) {
          console.warn('Could not add logo:', e);
        }
      }

      doc.setFont('helvetica', 'bold');
      doc.setFontSize(14.5);
      doc.setTextColor(200, 16, 46); // Red header
      doc.text(currentConfig.universityName, 105, 17, { align: 'center' });

      doc.setFontSize(11);
      doc.setTextColor(200, 16, 46);
      doc.text(currentConfig.malayalamTitle, 105, 23, { align: 'center' });

      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(70, 80, 95);
      doc.text(currentConfig.addressLine, 105, 28, { align: 'center' });
      doc.text(currentConfig.naacRating, 105, 32.5, { align: 'center' });

      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'bold');
      doc.setTextColor(0, 0, 0);
      doc.text(currentConfig.branchName, 105, 38.5, { align: 'center' });
      doc.text(currentConfig.examTitle, 105, 43.5, { align: 'center' });

      // Table Box with borders matching Image 3
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
        startY: 48,
        body: boxBody,
        theme: 'grid',
        styles: { font: 'helvetica', fontSize: 8.5, cellPadding: 3, textColor: [0, 0, 0], lineColor: [0, 0, 0], lineWidth: 0.35, valign: 'middle' },
        columnStyles: {
          0: { fontStyle: 'bold' }
        }
      });

      // Certificate Section
      const finalY = doc.lastAutoTable.finalY || 100;
      doc.setLineDashPattern([2, 2], 0);
      doc.setDrawColor(120, 120, 120);
      doc.line(14, finalY + 5, 196, finalY + 5);
      doc.line(14, finalY + 7, 196, finalY + 7);
      doc.setLineDashPattern([], 0);

      doc.setFontSize(12);
      doc.setFont('helvetica', 'bold');
      doc.text('CERTIFICATE', 105, finalY + 16, { align: 'center' });

      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      const certLines = doc.splitTextToSize(currentConfig.certificateText, 175);
      doc.text(certLines, 16, finalY + 24);

      // Signatures
      let sigY = finalY + 44;
      doc.setFontSize(8.5);
      doc.setFont('helvetica', 'bold');
      doc.text('INVIGILATOR', 16, sigY);
      doc.text('ADDL.CHIEF SUPERINTENDENT', 140, sigY);

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
      doc.text('CHIEF SUPERINTENDENT', 150, sigY);

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
    setStatusMsg('Generating consolidated PDF...');
    setIsProcessing(true);

    try {
      const doc = new jsPDF('p', 'mm', 'a4');

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
      setStatusMsg('Consolidated Master PDF downloaded successfully!');
    } catch (err) {
      console.error(err);
      setStatusMsg(`Error: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownloadZip = async () => {
    if (!processedData.groupKeys.length) return;
    setStatusMsg('Generating ZIP archive with individual PDFs...');
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

      setStatusMsg(`Exported ZIP with ${processedData.totalCount} individual PDF files!`);
    } catch (err) {
      console.error(err);
      setStatusMsg(`Error creating ZIP: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  // Save Template Configuration to Supabase Cloud / Local Library
  const handleSaveTemplate = async () => {
    if (!templateName.trim()) return alert('Please enter a template name.');
    
    const templateData = {
      name: templateName,
      archetype: activeArchetype,
      config: currentConfig,
      columnMappings,
      createdAt: new Date().toISOString()
    };

    if (user?.id) {
      const { error } = await supabase
        .from('templates')
        .insert([{ name: templateName, layout_data: templateData, user_id: user.id }]);
      if (error) {
        alert('Saved locally. Cloud save error: ' + error.message);
      } else {
        alert('Template saved to Cloud Library successfully!');
      }
    } else {
      localStorage.setItem(`template_${Date.now()}`, JSON.stringify(templateData));
      alert('Template saved to your local workspace!');
    }
  };

  return (
    <div style={{ padding: '24px 32px 80px', maxWidth: '1600px', margin: '0 auto', fontFamily: 'var(--font-family)' }}>
      
      {/* Top Title & Archetype Switcher Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '16px' }}>
        <div>
          <h1 style={{ margin: 0, fontSize: '24px', fontWeight: 800, display: 'flex', alignItems: 'center', gap: '10px' }}>
            <Sparkles size={24} color="var(--accent)" /> Report Template Studio
          </h1>
          <p style={{ margin: '4px 0 0 0', color: 'var(--muted)', fontSize: '13.5px' }}>
            Design and generate customizable reports (Nominal Roll, QP Statement, Exam Envelope Cover Labels) dynamically from your uploaded Excel data.
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
      <div style={{ display: 'grid', gridTemplateColumns: isSidebarOpen ? '380px 1fr' : '0px 1fr', gap: isSidebarOpen ? '24px' : '0px', transition: 'all 0.3s ease' }}>
        
        {/* LEFT COLUMN: Template Customization & Column Binding Form */}
        <div style={{ display: isSidebarOpen ? 'block' : 'none' }}>
          
          <div className="card" style={{ padding: '24px', marginBottom: '20px' }}>
            <h3 style={{ margin: '0 0 16px 0', fontSize: '16px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--accent)' }}>
              <Sliders size={18} /> 1. Header & Branding
            </h3>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div>
                <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>University Name (Header 1):</label>
                <input 
                  type="text" 
                  value={currentConfig.universityName || ''} 
                  onChange={(e) => updateCurrentConfig({ universityName: e.target.value })}
                  style={{ width: '100%', fontSize: '13px' }} 
                />
              </div>

              {activeArchetype === 'QP_COVER_LABEL' && (
                <>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Malayalam / Regional Title:</label>
                    <input 
                      type="text" 
                      value={currentConfig.malayalamTitle || ''} 
                      onChange={(e) => updateCurrentConfig({ malayalamTitle: e.target.value })}
                      style={{ width: '100%', fontSize: '13px' }} 
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Address & Location Line:</label>
                    <input 
                      type="text" 
                      value={currentConfig.addressLine || ''} 
                      onChange={(e) => updateCurrentConfig({ addressLine: e.target.value })}
                      style={{ width: '100%', fontSize: '13px' }} 
                    />
                  </div>
                  <div>
                    <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>NAAC Accreditation Grade:</label>
                    <input 
                      type="text" 
                      value={currentConfig.naacRating || ''} 
                      onChange={(e) => updateCurrentConfig({ naacRating: e.target.value })}
                      style={{ width: '100%', fontSize: '13px' }} 
                    />
                  </div>
                </>
              )}

              <div>
                <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Branch / Sub-title:</label>
                <input 
                  type="text" 
                  value={currentConfig.branchName || ''} 
                  onChange={(e) => updateCurrentConfig({ branchName: e.target.value })}
                  style={{ width: '100%', fontSize: '13px' }} 
                />
              </div>

              <div>
                <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Exam Title / Event:</label>
                <input 
                  type="text" 
                  value={currentConfig.examTitle || ''} 
                  onChange={(e) => updateCurrentConfig({ examTitle: e.target.value })}
                  style={{ width: '100%', fontSize: '13px' }} 
                />
              </div>

              {activeArchetype !== 'QP_COVER_LABEL' && (
                <div>
                  <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Session / Month Year:</label>
                  <input 
                    type="text" 
                    value={currentConfig.sessionName || ''} 
                    onChange={(e) => updateCurrentConfig({ sessionName: e.target.value })}
                    style={{ width: '100%', fontSize: '13px' }} 
                  />
                </div>
              )}
            </div>
          </div>

          {/* SECTION 2: Dynamic Excel Column Binding */}
          <div className="card" style={{ padding: '24px', marginBottom: '20px' }}>
            <h3 style={{ margin: '0 0 16px 0', fontSize: '16px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--accent)' }}>
              <Layers size={18} /> 2. Excel Column Data Binding
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

          {/* SECTION 3: Formatting & Toggles */}
          <div className="card" style={{ padding: '24px' }}>
            <h3 style={{ margin: '0 0 16px 0', fontSize: '16px', fontWeight: 700, display: 'flex', alignItems: 'center', gap: '8px', color: 'var(--accent)' }}>
              <Edit3 size={18} /> 3. Layout & Style Controls
            </h3>

            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '13px', fontWeight: 600 }}>
                <input 
                  type="checkbox"
                  checked={!!currentConfig.formatCodeDotNameDot}
                  onChange={(e) => updateCurrentConfig({ formatCodeDotNameDot: e.target.checked })}
                  style={{ width: '16px', height: '16px', accentColor: 'var(--accent)' }}
                />
                Format: CODE. - NAME. (with trailing dots)
              </label>

              {activeArchetype === 'QP_COVER_LABEL' && (
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', fontSize: '13px', fontWeight: 600 }}>
                  <input 
                    type="checkbox"
                    checked={!!currentConfig.showLogo}
                    onChange={(e) => updateCurrentConfig({ showLogo: e.target.checked })}
                    style={{ width: '16px', height: '16px', accentColor: 'var(--accent)' }}
                  />
                  Include University Crest Logo
                </label>
              )}
            </div>
          </div>

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

          {/* Rendered Live Paper Display (1:1 with Output PDF) */}
          <div className="card print-container" style={{ padding: '40px 48px', background: 'white', border: '1px solid #cbd5e1', borderRadius: '8px', boxShadow: '0 8px 30px rgba(0,0,0,0.08)', minHeight: '800px' }}>
            
            {/* ARCHETYPE 1: NOMINAL ROLL LIVE VIEW */}
            {activeArchetype === 'NOMINAL_ROLL' && (
              <div>
                <div style={{ textAlign: 'center', marginBottom: '20px', borderBottom: '1px solid #eee', paddingBottom: '16px' }}>
                  <div style={{ fontSize: '20px', fontWeight: 800, color: '#000', marginBottom: '4px' }}>
                    {currentConfig.universityName}
                  </div>
                  <div style={{ fontSize: '15px', fontWeight: 700, color: '#111', marginBottom: '6px' }}>
                    {currentConfig.branchName}
                  </div>
                  <div style={{ fontSize: '14.5px', fontWeight: 700, color: '#111', marginBottom: '4px' }}>
                    {currentConfig.examTitle}
                  </div>
                  {currentConfig.sessionName && (
                    <div style={{ fontSize: '14px', fontWeight: 600, color: '#222', marginBottom: '6px' }}>
                      {currentConfig.sessionName}
                    </div>
                  )}
                </div>

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

                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                  <thead>
                    <tr style={{ background: '#f1f5f9', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000' }}>
                      <th style={{ border: '1px solid #000', padding: '8px 4px', width: '45px', textAlign: 'center', fontWeight: 700 }}>{currentConfig.headers.slNo}</th>
                      <th style={{ border: '1px solid #000', padding: '8px 6px', width: '140px', textAlign: 'center', fontWeight: 700 }}>{currentConfig.headers.regNo}</th>
                      <th style={{ border: '1px solid #000', padding: '8px 10px', width: '180px', textAlign: 'left', fontWeight: 700 }}>{currentConfig.headers.name}</th>
                      <th style={{ border: '1px solid #000', padding: '8px 10px', textAlign: 'left', fontWeight: 700 }}>{currentConfig.headers.courses}</th>
                      <th style={{ border: '1px solid #000', padding: '8px 6px', width: '90px', textAlign: 'center', fontWeight: 700 }}>{currentConfig.headers.remarks}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {currentPreviewGroup?.candidates?.map((cand, candIdx) => {
                      const courses = cand.courses || [];
                      return courses.map((crs, crsIdx) => (
                        <tr key={`${candIdx}_${crsIdx}`} style={{ borderBottom: '1px solid #000' }}>
                          {crsIdx === 0 && (
                            <>
                              <td rowSpan={courses.length} style={{ border: '1px solid #000', padding: '6px 4px', textAlign: 'center', verticalAlign: 'middle' }}>{candIdx + 1}</td>
                              <td rowSpan={courses.length} style={{ border: '1px solid #000', padding: '6px 8px', textAlign: 'center', verticalAlign: 'middle', fontWeight: 700 }}>{cand.seatNo}</td>
                              <td rowSpan={courses.length} style={{ border: '1px solid #000', padding: '6px 10px', verticalAlign: 'middle', fontWeight: 700 }}>{cand.studentName}</td>
                            </>
                          )}
                          <td style={{ border: '1px solid #000', padding: '5px 8px', verticalAlign: 'middle' }}>{crs.display}</td>
                          {crsIdx === 0 && (
                            <td rowSpan={courses.length} style={{ border: '1px solid #000', padding: '6px', textAlign: 'center', verticalAlign: 'middle' }}></td>
                          )}
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
                <div style={{ textAlign: 'center', marginBottom: '20px' }}>
                  <div style={{ fontSize: '20px', fontWeight: 800, color: '#000', marginBottom: '4px' }}>
                    {currentConfig.universityName}
                  </div>
                  <div style={{ fontSize: '15px', fontWeight: 700, color: '#111', marginBottom: '6px' }}>
                    {currentConfig.branchName}
                  </div>
                  <div style={{ fontSize: '14.5px', fontWeight: 700, color: '#111', marginBottom: '4px' }}>
                    {currentConfig.examTitle}
                  </div>
                  {currentConfig.sessionName && (
                    <div style={{ fontSize: '14px', fontWeight: 600, color: '#222', marginBottom: '6px' }}>
                      {currentConfig.sessionName}
                    </div>
                  )}
                  <div style={{ fontSize: '14.5px', fontWeight: 700, color: '#000', marginTop: '6px' }}>
                    {currentConfig.centerPrefix} {effectiveGroupKey}
                  </div>
                </div>

                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                  <thead>
                    <tr style={{ background: '#f1f5f9', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000' }}>
                      <th style={{ border: '1px solid #000', padding: '8px 4px', width: '50px', textAlign: 'center', fontWeight: 700 }}>{currentConfig.headers.slNo}</th>
                      <th style={{ border: '1px solid #000', padding: '8px 8px', width: '160px', textAlign: 'left', fontWeight: 700 }}>{currentConfig.headers.date}</th>
                      <th style={{ border: '1px solid #000', padding: '8px 12px', textAlign: 'left', fontWeight: 700 }}>{currentConfig.headers.course}</th>
                      <th style={{ border: '1px solid #000', padding: '8px 6px', width: '65px', textAlign: 'center', fontWeight: 700 }}>{currentConfig.headers.count}</th>
                      <th style={{ border: '1px solid #000', padding: '8px 6px', width: '70px', textAlign: 'center', fontWeight: 700 }}>{currentConfig.headers.qp}</th>
                      <th style={{ border: '1px solid #000', padding: '8px 6px', width: '70px', textAlign: 'center', fontWeight: 700 }}>{currentConfig.headers.lp}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {currentPreviewGroup?.items?.map((it, idx) => (
                      <tr key={idx} style={{ borderBottom: '1px solid #000' }}>
                        <td style={{ border: '1px solid #000', padding: '6px 4px', textAlign: 'center' }}>{idx + 1}</td>
                        <td style={{ border: '1px solid #000', padding: '6px 8px' }}>{it.date}</td>
                        <td style={{ border: '1px solid #000', padding: '6px 12px' }}>{it.courseDisplay}</td>
                        <td style={{ border: '1px solid #000', padding: '6px 6px', textAlign: 'center', fontWeight: 700 }}>{it.studentCount}</td>
                        <td style={{ border: '1px solid #000', padding: '6px 6px', textAlign: 'center' }}></td>
                        <td style={{ border: '1px solid #000', padding: '6px 6px', textAlign: 'center' }}></td>
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
                  {/* Header */}
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '16px', marginBottom: '14px' }}>
                    {currentConfig.showLogo && logoBase64 && (
                      <img src={logoBase64} alt="Crest" style={{ height: '60px' }} />
                    )}
                    <div style={{ textAlign: 'center' }}>
                      <div style={{ fontSize: '22px', fontWeight: 800, color: '#c8102e', letterSpacing: '0.3px' }}>
                        {currentConfig.universityName}
                      </div>
                      <div style={{ fontSize: '16px', fontWeight: 700, color: '#c8102e' }}>
                        {currentConfig.malayalamTitle}
                      </div>
                      <div style={{ fontSize: '13px', color: '#475569', marginTop: '2px' }}>
                        {currentConfig.addressLine}
                      </div>
                      <div style={{ fontSize: '12.5px', color: '#475569' }}>
                        {currentConfig.naacRating}
                      </div>
                    </div>
                  </div>

                  <div style={{ textAlign: 'center', marginBottom: '16px' }}>
                    <div style={{ fontSize: '14.5px', fontWeight: 800, color: '#000' }}>
                      {currentConfig.branchName}
                    </div>
                    <div style={{ fontSize: '14.5px', fontWeight: 800, color: '#000' }}>
                      {currentConfig.examTitle}
                    </div>
                  </div>

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
                <div style={{ textAlign: 'center', marginBottom: '20px' }}>
                  <div style={{ fontSize: '20px', fontWeight: 800, color: '#000', marginBottom: '4px' }}>
                    {currentConfig.universityName}
                  </div>
                  <div style={{ fontSize: '15px', fontWeight: 700, color: '#111', marginBottom: '6px' }}>
                    {currentConfig.branchName}
                  </div>
                  <div style={{ fontSize: '14.5px', fontWeight: 700, color: '#111', marginBottom: '4px' }}>
                    {currentConfig.examTitle}
                  </div>
                </div>

                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
                  <thead>
                    <tr style={{ background: '#f1f5f9', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000' }}>
                      <th style={{ border: '1px solid #000', padding: '8px 4px', width: '45px', textAlign: 'center' }}>SL No</th>
                      {dataset.columns.slice(0, 6).map(c => (
                        <th key={c} style={{ border: '1px solid #000', padding: '8px 6px', textAlign: 'left' }}>{c}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {dataset.rows.slice(0, 25).map((row, idx) => (
                      <tr key={idx} style={{ borderBottom: '1px solid #000' }}>
                        <td style={{ border: '1px solid #000', padding: '6px 4px', textAlign: 'center' }}>{idx + 1}</td>
                        {dataset.columns.slice(0, 6).map(c => (
                          <td key={c} style={{ border: '1px solid #000', padding: '6px' }}>{row[c]}</td>
                        ))}
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
