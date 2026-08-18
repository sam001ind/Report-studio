import { useState, useMemo, useEffect } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import JSZip from 'jszip';
import { 
  Download, 
  UploadCloud, 
  FileSpreadsheet, 
  Archive, 
  FileText,
  Printer,
  Search, 
  BookOpen, 
  Users,
  Sliders,
  MapPin,
  RotateCcw,
  Sparkles,
  GraduationCap,
  Building2,
  RefreshCw
} from 'lucide-react';
import { readSpreadsheetFile } from '../utils/excelParser';

const SAMPLE_NOMINAL_DATA = [
  { "Programme": "B.A. English (Private Registration)", "Venue_Code": "101", "Venue_Name": "Government College Kasaragod", "Seat_No": "1001", "Student_Name": "Muhammed Rashid K", "Course_Code": "ENG1B01", "Course_Title": "Reading Poetry", "Session": "FN" },
  { "Programme": "B.A. English (Private Registration)", "Venue_Code": "101", "Venue_Name": "Government College Kasaragod", "Seat_No": "1001", "Student_Name": "Muhammed Rashid K", "Course_Code": "ENG1A01", "Course_Title": "Communication Skills in English", "Session": "AN" },
  { "Programme": "B.A. English (Private Registration)", "Venue_Code": "101", "Venue_Name": "Government College Kasaragod", "Seat_No": "1002", "Student_Name": "Ananya S Nair", "Course_Code": "ENG1B01", "Course_Title": "Reading Poetry", "Session": "FN" },
  { "Programme": "B.A. English (Private Registration)", "Venue_Code": "101", "Venue_Name": "Government College Kasaragod", "Seat_No": "1002", "Student_Name": "Ananya S Nair", "Course_Code": "MAL1A07", "Course_Title": "Malayala Bhashayum Sahithyavum", "Session": "AN" },
  { "Programme": "B.A. English (Private Registration)", "Venue_Code": "101", "Venue_Name": "Government College Kasaragod", "Seat_No": "1003", "Student_Name": "Fathima Hameed", "Course_Code": "ENG1B01", "Course_Title": "Reading Poetry", "Session": "FN" },
  { "Programme": "B.Sc. Computer Science", "Venue_Code": "102", "Venue_Name": "Payyanur College", "Seat_No": "2001", "Student_Name": "Abhijith T", "Course_Code": "BCS1B01", "Course_Title": "Computer Fundamentals & HTML", "Session": "FN" },
  { "Programme": "B.Sc. Computer Science", "Venue_Code": "102", "Venue_Name": "Payyanur College", "Seat_No": "2001", "Student_Name": "Abhijith T", "Course_Code": "MAT1C01", "Course_Title": "Mathematics I", "Session": "AN" },
  { "Programme": "B.Sc. Computer Science", "Venue_Code": "102", "Venue_Name": "Payyanur College", "Seat_No": "2002", "Student_Name": "Devika Menon", "Course_Code": "BCS1B01", "Course_Title": "Computer Fundamentals & HTML", "Session": "FN" },
  { "Programme": "B.Com Finance", "Venue_Code": "103", "Venue_Name": "Sir Syed College Taliparamba", "Seat_No": "3001", "Student_Name": "Rahul K V", "Course_Code": "BCM1B01", "Course_Title": "Management Concepts & Business Ethics", "Session": "FN" },
  { "Programme": "B.Com Finance", "Venue_Code": "103", "Venue_Name": "Sir Syed College Taliparamba", "Seat_No": "3002", "Student_Name": "Sneha Prakash", "Course_Code": "BCM1B01", "Course_Title": "Management Concepts & Business Ethics", "Session": "FN" }
];

const SllNominalPage = () => {

  // Upload and file states
  const [headers, setHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready to upload Excel or ZIP files.');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Header & Title Configuration
  const [universityName, setUniversityName] = useState('Kannur University');
  const [branchName, setBranchName] = useState('(Examination Branch)');
  const [examTitle, setExamTitle] = useState('Venue-Wise Candidate Nominal Roll & Attendance Record');
  const [sessionName, setSessionName] = useState('I Semester Degree Examination - November 2025');
  const [groupByOption, setGroupByOption] = useState('programme_venue'); // 'programme_venue' | 'venue'

  // Column mapping (0-indexed indices)
  const [columnMapping, setColumnMapping] = useState({
    programme: -1,
    venueCode: -1,
    venueName: -1,
    seatNo: -1,
    name: -1,
    courseCode: -1,
    courseTitle: -1,
    session: -1
  });

  // UI Navigation & Filters
  const [activeGroupKey, setActiveGroupKey] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [showConfig, setShowConfig] = useState(false);

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  // Auto-detect columns on headers update
  useEffect(() => {
    if (headers.length > 0) {
      const autoMap = {
        programme: -1,
        venueCode: -1,
        venueName: -1,
        seatNo: -1,
        name: -1,
        courseCode: -1,
        courseTitle: -1,
        session: -1
      };

      headers.forEach((headerName, idx) => {
        const lower = headerName.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (lower.includes('programme') || lower.includes('program') || lower.includes('course') && !lower.includes('code') && !lower.includes('title')) {
          if (autoMap.programme === -1) autoMap.programme = idx;
        }
        if (lower.includes('venuecode') || lower.includes('centercode') || lower.includes('collegecode') || (lower.includes('venue') && lower.includes('code'))) {
          if (autoMap.venueCode === -1) autoMap.venueCode = idx;
        }
        if (lower.includes('venuename') || lower.includes('centername') || lower.includes('collegename') || (lower.includes('venue') && !lower.includes('code')) || lower.includes('college')) {
          if (autoMap.venueName === -1) autoMap.venueName = idx;
        }
        if (lower.includes('seat') || lower.includes('reg') || lower.includes('prn') || lower.includes('roll') || lower.includes('candidateno')) {
          if (autoMap.seatNo === -1) autoMap.seatNo = idx;
        }
        if (lower.includes('studentname') || lower.includes('candidatename') || lower.includes('name') && !lower.includes('venue') && !lower.includes('college')) {
          if (autoMap.name === -1) autoMap.name = idx;
        }
        if (lower.includes('coursecode') || lower.includes('subjectcode') || lower.includes('papercode') || lower.includes('qpcode')) {
          if (autoMap.courseCode === -1) autoMap.courseCode = idx;
        }
        if (lower.includes('coursetitle') || lower.includes('subjectname') || lower.includes('coursename') || lower.includes('title')) {
          if (autoMap.courseTitle === -1) autoMap.courseTitle = idx;
        }
        if (lower.includes('session') || lower.includes('date') || lower.includes('time')) {
          if (autoMap.session === -1) autoMap.session = idx;
        }
      });

      setColumnMapping(autoMap);
    }
  }, [headers]);

  // Load sample dataset
  const loadSampleData = () => {
    setIsProcessing(true);
    setStatus('Loading sample dataset...', 'normal');
    setTimeout(() => {
      const cols = Object.keys(SAMPLE_NOMINAL_DATA[0]);
      setHeaders(cols);
      setRawRows(SAMPLE_NOMINAL_DATA);
      setIsProcessing(false);
      setStatus(`Loaded ${SAMPLE_NOMINAL_DATA.length} sample nominal records.`, 'success');
    }, 150);
  };

  // Universal File Upload Handler (Auto-extracts .zip, .xlsx, .xls, .csv)
  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsProcessing(true);
    setStatus(`Reading file "${file.name}"...`, 'normal');

    try {
      const { rows, columns } = await readSpreadsheetFile(file);
      setHeaders(columns);
      setRawRows(rows);
      setActiveGroupKey('');
      setStatus(`Successfully loaded ${rows.length} rows from ${file.name}!`, 'success');
    } catch (err) {
      setStatus(`Error reading file: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Reset tool state
  const handleReset = () => {
    setHeaders([]);
    setRawRows([]);
    setActiveGroupKey('');
    setSearchQuery('');
    setStatus('Ready to upload Excel or ZIP files.', 'normal');
  };

  // Process & Group Data by [Programme + Venue] or [Venue] with Multi-Subject Merging per Seat No
  const processedGroups = useMemo(() => {
    if (!rawRows.length || !headers.length) return {};

    const groups = {};

    rawRows.forEach((row) => {
      // Extract mapped values
      const getVal = (colIdx, fallback = '') => {
        if (colIdx === -1) return fallback;
        const key = headers[colIdx];
        if (!key) return fallback;
        const val = row[key];
        return val !== undefined && val !== null ? String(val).trim() : fallback;
      };

      const prog = getVal(columnMapping.programme, 'General Programme');
      const vCode = getVal(columnMapping.venueCode, '');
      const vName = getVal(columnMapping.venueName, 'Unassigned Venue');
      const seatNo = getVal(columnMapping.seatNo, 'N/A');
      const studentName = getVal(columnMapping.name, 'Candidate');
      const cCode = getVal(columnMapping.courseCode, '');
      const cTitle = getVal(columnMapping.courseTitle, '');
      const sess = getVal(columnMapping.session, '');

      // Venue Display Label
      const venueLabel = vCode ? `${vCode} - ${vName}` : vName;

      // Group Key Strategy
      const groupKey = groupByOption === 'programme_venue'
        ? `${prog} • ${venueLabel}`
        : venueLabel;

      if (!groups[groupKey]) {
        groups[groupKey] = {
          programme: prog,
          venueCode: vCode,
          venueName: vName,
          venueLabel,
          candidatesMap: {}
        };
      }

      // Group courses by Seat Number for this Candidate
      if (!groups[groupKey].candidatesMap[seatNo]) {
        groups[groupKey].candidatesMap[seatNo] = {
          seatNo,
          studentName,
          courses: []
        };
      }

      // Append course if not duplicate
      if (cCode || cTitle) {
        const cleanCode = String(cCode || '').trim().replace(/^[.\s]+|[.\s]+$/g, '');
        const cleanTitle = String(cTitle || '').trim().replace(/^[-–—:.\s]+|[.\s]+$/g, '');
        
        const displayStr = cleanCode && cleanTitle ? `${cleanCode} - ${cleanTitle}` : (cleanCode || cleanTitle);

        const courseItem = {
          code: cleanCode,
          title: cleanTitle,
          session: sess,
          display: displayStr
        };
        const exists = groups[groupKey].candidatesMap[seatNo].courses.some(c => c.code === cleanCode && c.title === cleanTitle);
        if (!exists) {
          groups[groupKey].candidatesMap[seatNo].courses.push(courseItem);
        }
      }
    });

    // Convert candidatesMap to sorted candidate array
    const finalized = {};
    Object.entries(groups).forEach(([gKey, gData]) => {
      const candidateList = Object.values(gData.candidatesMap).sort((a, b) => {
        // Natural alphanumeric sort on Seat No
        return a.seatNo.localeCompare(b.seatNo, undefined, { numeric: true, sensitivity: 'base' });
      });

      finalized[gKey] = {
        ...gData,
        candidates: candidateList,
        totalCandidates: candidateList.length
      };
    });

    return finalized;
  }, [rawRows, headers, columnMapping, groupByOption]);

  const groupKeys = useMemo(() => Object.keys(processedGroups).sort(), [processedGroups]);
  const effectiveGroupKey = activeGroupKey && processedGroups[activeGroupKey] ? activeGroupKey : (groupKeys[0] || '');

  // Unique list of programmes for cascading selector
  const programmesList = useMemo(() => {
    const set = new Set();
    groupKeys.forEach(k => {
      if (processedGroups[k]?.programme) set.add(processedGroups[k].programme);
    });
    return Array.from(set).sort();
  }, [groupKeys, processedGroups]);

  const currentProgramme = processedGroups[effectiveGroupKey]?.programme || programmesList[0] || '';

  const venuesForCurrentProg = useMemo(() => {
    if (groupByOption !== 'programme_venue') return groupKeys;
    return groupKeys.filter(k => processedGroups[k]?.programme === currentProgramme);
  }, [groupByOption, groupKeys, processedGroups, currentProgramme]);

  // Statistics
  const stats = useMemo(() => {
    if (!groupKeys.length) return { totalGroups: 0, totalCandidates: 0, totalUniqueCourses: 0, totalVenues: 0 };
    
    let totalCands = 0;
    const courseSet = new Set();
    const venueSet = new Set();

    groupKeys.forEach(k => {
      const g = processedGroups[k];
      totalCands += g.totalCandidates;
      venueSet.add(g.venueLabel);
      g.candidates.forEach(c => {
        c.courses.forEach(crs => courseSet.add(crs.code || crs.title));
      });
    });

    return {
      totalGroups: groupKeys.length,
      totalCandidates: totalCands,
      totalUniqueCourses: courseSet.size,
      totalVenues: venueSet.size
    };
  }, [groupKeys, processedGroups]);

  // Filtered candidate list for active tab & search
  const currentTabCandidates = useMemo(() => {
    if (!effectiveGroupKey || !processedGroups[effectiveGroupKey]) return [];
    let list = processedGroups[effectiveGroupKey].candidates;

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      list = list.filter(c => 
        c.seatNo.toLowerCase().includes(q) ||
        c.studentName.toLowerCase().includes(q) ||
        c.courses.some(crs => crs.display.toLowerCase().includes(q))
      );
    }
    return list;
  }, [effectiveGroupKey, processedGroups, searchQuery]);

  // EXPORT: Single Venue / Group PDF
  const exportSingleGroupPdf = (gKey) => {
    const targetKey = gKey || effectiveGroupKey;
    if (!targetKey || !processedGroups[targetKey]) return;

    const gData = processedGroups[targetKey];
    const doc = new jsPDF('p', 'mm', 'a4');

    // Header
    doc.setFontSize(14);
    doc.setFont('helvetica', 'bold');
    doc.text(universityName, 105, 14, { align: 'center' });

    doc.setFontSize(10.5);
    doc.setFont('helvetica', 'bold');
    doc.text(branchName, 105, 19.5, { align: 'center' });

    doc.setFontSize(10);
    doc.setFont('helvetica', 'bold');
    doc.text(examTitle, 105, 25, { align: 'center' });

    if (sessionName) {
      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'normal');
      doc.text(sessionName, 105, 30, { align: 'center' });
    }

    // Venue & Programme Info Box (Stacked with auto-wrapping for long names)
    const progText = `Programme: ${gData.programme || 'N/A'}`;
    const venueText = `Venue: ${gData.venueLabel || 'N/A'}`;
    const candCountText = `Total Candidates: ${gData.totalCandidates}`;

    doc.setFontSize(9);
    doc.setFont('helvetica', 'bold');
    
    // Auto-wrap long lines so they never truncate or overflow
    const progLines = doc.splitTextToSize(progText, 136);
    const venueLines = doc.splitTextToSize(venueText, 136);
    
    const boxHeight = Math.max(16, (progLines.length + venueLines.length) * 4.5 + 5);

    doc.setDrawColor(180, 180, 180);
    doc.setFillColor(248, 249, 250);
    doc.roundedRect(14, 34, 182, boxHeight, 2, 2, 'FD');

    let textY = 38.5;
    doc.setTextColor(20, 20, 20);
    doc.text(progLines, 18, textY);
    textY += progLines.length * 4.5;
    
    doc.setTextColor(23, 107, 135);
    doc.text(venueLines, 18, textY);
    
    doc.setTextColor(20, 20, 20);
    doc.text(candCountText, 150, 39);

    // Table Content
    const tableData = gData.candidates.map((c, idx) => {
      const coursesStr = c.courses.map((crs, cIdx) => `${cIdx + 1}. ${crs.display}`).join('\n');
      return [
        idx + 1,
        c.seatNo,
        c.studentName,
        coursesStr || '—',
        '' // Blank Remarks
      ];
    });

    const startTableY = 34 + boxHeight + 4;

    autoTable(doc, {
      startY: startTableY,
      head: [['Sl No', 'Register Number', 'Candidate Name', 'Courses', 'Remarks']],
      body: tableData,
      theme: 'grid',
      styles: { 
        font: 'helvetica',
        fontSize: 8.5, 
        lineHeightFactor: 1.4,
        valign: 'middle',
        cellPadding: { top: 4, bottom: 4, left: 3, right: 3 }, 
        textColor: [15, 23, 42],
        lineColor: [200, 205, 215],
        lineWidth: 0.2,
        overflow: 'linebreak'
      },
      headStyles: { 
        fillColor: [23, 107, 135], 
        textColor: 255, 
        fontSize: 9, 
        fontStyle: 'bold',
        halign: 'center',
        valign: 'middle',
        cellPadding: { top: 4, bottom: 4, left: 2, right: 2 }
      },
      columnStyles: {
        0: { cellWidth: 10, halign: 'center', valign: 'middle' },
        1: { cellWidth: 30, fontStyle: 'bold', halign: 'center', valign: 'middle' },
        2: { cellWidth: 40, fontStyle: 'bold', valign: 'middle' },
        3: { cellWidth: 78, valign: 'middle', cellPadding: { top: 3.5, bottom: 3.5, left: 3, right: 2 } },
        4: { cellWidth: 24, valign: 'middle' }
      },
      didDrawPage: (data) => {
        doc.setFontSize(8);
        doc.setTextColor(100);
        doc.text(`Page ${data.pageNumber} of ${doc.internal.getNumberOfPages()}`, 105, 290, { align: 'center' });
      }
    });

    const safeName = targetKey.replace(/[/\\?%*:|"<>•]/g, '_').slice(0, 40);
    doc.save(`Nominal_Roll_${safeName}.pdf`);
  };

  // EXPORT: Consolidated Master PDF (All Venues Combined)
  const exportConsolidatedPdf = () => {
    if (!groupKeys.length) return;
    setStatus('Generating consolidated PDF for all venues...', 'normal');
    setIsProcessing(true);

    try {
      const doc = new jsPDF('p', 'mm', 'a4');

      groupKeys.forEach((gKey, gIdx) => {
        if (gIdx > 0) doc.addPage();

        const gData = processedGroups[gKey];

        doc.setFontSize(14);
        doc.setFont('helvetica', 'bold');
        doc.text(universityName, 105, 14, { align: 'center' });

        doc.setFontSize(10.5);
        doc.setFont('helvetica', 'bold');
        doc.text(branchName, 105, 19.5, { align: 'center' });

        doc.setFontSize(10);
        doc.setFont('helvetica', 'bold');
        doc.text(examTitle, 105, 25, { align: 'center' });

        if (sessionName) {
          doc.setFontSize(9.5);
          doc.setFont('helvetica', 'normal');
          doc.text(sessionName, 105, 30, { align: 'center' });
        }

        const progText = `Programme: ${gData.programme || 'N/A'}`;
        const venueText = `Venue: ${gData.venueLabel || 'N/A'}`;
        const candCountText = `Candidates: ${gData.totalCandidates}`;

        doc.setFontSize(9);
        doc.setFont('helvetica', 'bold');
        
        const progLines = doc.splitTextToSize(progText, 136);
        const venueLines = doc.splitTextToSize(venueText, 136);
        
        const boxHeight = Math.max(16, (progLines.length + venueLines.length) * 4.5 + 5);

        doc.setDrawColor(180, 180, 180);
        doc.setFillColor(248, 249, 250);
        doc.roundedRect(14, 34, 182, boxHeight, 2, 2, 'FD');

        let textY = 38.5;
        doc.setTextColor(20, 20, 20);
        doc.text(progLines, 18, textY);
        textY += progLines.length * 4.5;
        
        doc.setTextColor(23, 107, 135);
        doc.text(venueLines, 18, textY);
        
        doc.setTextColor(20, 20, 20);
        doc.text(candCountText, 150, 39);

        const tableData = gData.candidates.map((c, idx) => {
          const coursesStr = c.courses.map((crs, cIdx) => `${cIdx + 1}. ${crs.display}`).join('\n');
          return [idx + 1, c.seatNo, c.studentName, coursesStr || '—', ''];
        });

        const startTableY = 34 + boxHeight + 4;

        autoTable(doc, {
          startY: startTableY,
          head: [['Sl No', 'Register Number', 'Candidate Name', 'Courses', 'Remarks']],
          body: tableData,
          theme: 'grid',
          styles: { 
            font: 'helvetica',
            fontSize: 8.5, 
            lineHeightFactor: 1.4,
            valign: 'middle',
            cellPadding: { top: 4, bottom: 4, left: 3, right: 3 }, 
            textColor: [15, 23, 42],
            lineColor: [200, 205, 215],
            lineWidth: 0.2,
            overflow: 'linebreak'
          },
          headStyles: { 
            fillColor: [23, 107, 135], 
            textColor: 255, 
            fontSize: 9, 
            fontStyle: 'bold',
            halign: 'center',
            valign: 'middle',
            cellPadding: { top: 4, bottom: 4, left: 2, right: 2 }
          },
          columnStyles: {
            0: { cellWidth: 10, halign: 'center', valign: 'middle' },
            1: { cellWidth: 30, fontStyle: 'bold', halign: 'center', valign: 'middle' },
            2: { cellWidth: 40, fontStyle: 'bold', valign: 'middle' },
            3: { cellWidth: 78, valign: 'middle', cellPadding: { top: 3.5, bottom: 3.5, left: 3, right: 2 } },
            4: { cellWidth: 24, valign: 'middle' }
          }
        });
      });

      doc.save('Master_Consolidated_Nominal_Roll.pdf');
      setStatus('Consolidated Master PDF downloaded successfully!', 'success');
    } catch (err) {
      setStatus(`Error generating consolidated PDF: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // EXPORT: All Groups & Complete Report in ZIP Archive
  const exportAllGroupsZip = async () => {
    if (!groupKeys.length) return;
    setStatus('Generating complete ZIP package with all venue nominal rolls...', 'normal');
    setIsProcessing(true);

    try {
      const zip = new JSZip();
      const consolidatedDoc = new jsPDF('p', 'mm', 'a4');

      groupKeys.forEach((gKey, gIdx) => {
        const gData = processedGroups[gKey];
        const doc = new jsPDF('p', 'mm', 'a4');

        if (gIdx > 0) consolidatedDoc.addPage();

        [doc, consolidatedDoc].forEach((targetDoc) => {
          targetDoc.setFontSize(14);
          targetDoc.setFont('helvetica', 'bold');
          targetDoc.text(universityName, 105, 14, { align: 'center' });

          targetDoc.setFontSize(10.5);
          targetDoc.setFont('helvetica', 'bold');
          targetDoc.text(branchName, 105, 19.5, { align: 'center' });

          targetDoc.setFontSize(10);
          targetDoc.setFont('helvetica', 'bold');
          targetDoc.text(examTitle, 105, 25, { align: 'center' });

          if (sessionName) {
            targetDoc.setFontSize(9.5);
            targetDoc.setFont('helvetica', 'normal');
            targetDoc.text(sessionName, 105, 30, { align: 'center' });
          }

          const progText = `Programme: ${gData.programme || 'N/A'}`;
          const venueText = `Venue: ${gData.venueLabel || 'N/A'}`;
          const candCountText = `Candidates: ${gData.totalCandidates}`;

          targetDoc.setFontSize(9);
          targetDoc.setFont('helvetica', 'bold');
          
          const progLines = targetDoc.splitTextToSize(progText, 136);
          const venueLines = targetDoc.splitTextToSize(venueText, 136);
          
          const boxHeight = Math.max(16, (progLines.length + venueLines.length) * 4.5 + 5);

          targetDoc.setDrawColor(180, 180, 180);
          targetDoc.setFillColor(248, 249, 250);
          targetDoc.roundedRect(14, 34, 182, boxHeight, 2, 2, 'FD');

          let textY = 38.5;
          targetDoc.setTextColor(20, 20, 20);
          targetDoc.text(progLines, 18, textY);
          textY += progLines.length * 4.5;
          
          targetDoc.setTextColor(23, 107, 135);
          targetDoc.text(venueLines, 18, textY);
          
          targetDoc.setTextColor(20, 20, 20);
          targetDoc.text(candCountText, 150, 39);

          const tableData = gData.candidates.map((c, idx) => {
            const coursesStr = c.courses.map((crs, cIdx) => `${cIdx + 1}. ${crs.display}`).join('\n');
            return [idx + 1, c.seatNo, c.studentName, coursesStr || '—', ''];
          });

          const startTableY = 34 + boxHeight + 4;

          autoTable(targetDoc, {
            startY: startTableY,
            head: [['Sl No', 'Register Number', 'Candidate Name', 'Courses', 'Remarks']],
            body: tableData,
            theme: 'grid',
            styles: { 
              font: 'helvetica',
              fontSize: 8.5, 
              lineHeightFactor: 1.4,
              valign: 'middle',
              cellPadding: { top: 4, bottom: 4, left: 3, right: 3 }, 
              textColor: [15, 23, 42],
              lineColor: [200, 205, 215],
              lineWidth: 0.2,
              overflow: 'linebreak'
            },
            headStyles: { 
              fillColor: [23, 107, 135], 
              textColor: 255, 
              fontSize: 9, 
              fontStyle: 'bold',
              halign: 'center',
              valign: 'middle',
              cellPadding: { top: 4, bottom: 4, left: 2, right: 2 }
            },
            columnStyles: {
              0: { cellWidth: 10, halign: 'center', valign: 'middle' },
              1: { cellWidth: 30, fontStyle: 'bold', halign: 'center', valign: 'middle' },
              2: { cellWidth: 40, fontStyle: 'bold', valign: 'middle' },
              3: { cellWidth: 78, valign: 'middle', cellPadding: { top: 3.5, bottom: 3.5, left: 3, right: 2 } },
              4: { cellWidth: 24, valign: 'middle' }
            }
          });
        });

        const pdfBlob = doc.output('blob');
        const safeFileName = `Nominal_Roll_${gKey.replace(/[/\\?%*:|"<>•]/g, '_').slice(0, 40)}.pdf`;
        zip.file(safeFileName, pdfBlob);
      });

      // Add Master Consolidated PDF to ZIP
      const consolidatedPdfBlob = consolidatedDoc.output('blob');
      zip.file('Master_Consolidated_Nominal_Roll.pdf', consolidatedPdfBlob);

      // Add Master Excel to ZIP
      const wb = XLSX.utils.book_new();
      const summaryRows = [
        ['VENUE-WISE NOMINAL ROLL MASTER SUMMARY'],
        ['Generated At', new Date().toLocaleString()],
        ['Total Venues', stats.totalVenues],
        ['Total Groups', stats.totalGroups],
        ['Total Candidates', stats.totalCandidates],
        ['Total Unique Courses', stats.totalUniqueCourses],
        [''],
        ['SL NO', 'PROGRAMME', 'VENUE', 'CANDIDATE COUNT']
      ];
      groupKeys.forEach((gKey, idx) => {
        const g = processedGroups[gKey];
        summaryRows.push([idx + 1, g.programme, g.venueLabel, g.totalCandidates]);
      });
      const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
      XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

      groupKeys.forEach((gKey, gIdx) => {
        const g = processedGroups[gKey];
        const rowsData = g.candidates.map((c, cIdx) => ({
          'Sl_No': cIdx + 1,
          'Register_Number': c.seatNo,
          'Candidate_Name': c.studentName,
          'Courses': c.courses.map(crs => crs.display).join('; '),
          'Remarks': '',
          'Programme': g.programme,
          'Venue': g.venueLabel
        }));
        const ws = XLSX.utils.json_to_sheet(rowsData);
        const sheetName = `Group_${gIdx + 1}_${g.venueCode || 'V'}`.slice(0, 30);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      });
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      zip.file('Master_Venue_Nominal_Roll.xlsx', excelBuffer);

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const url = URL.createObjectURL(zipBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Venue_Nominal_Roll_Complete_Report.zip';
      a.click();
      URL.revokeObjectURL(url);

      setStatus(`Exported Complete Report ZIP with ${groupKeys.length} individual venue PDFs, Consolidated PDF, and Master Excel!`, 'success');
    } catch (err) {
      setStatus(`Error generating ZIP: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // EXPORT: Master Multi-Sheet Excel Workbook
  const exportMasterExcel = () => {
    if (!groupKeys.length) return;
    const wb = XLSX.utils.book_new();

    // Summary Sheet
    const summaryRows = [
      ['VENUE-WISE NOMINAL ROLL MASTER SUMMARY'],
      ['Generated At', new Date().toLocaleString()],
      ['Total Venues', stats.totalVenues],
      ['Total Groups', stats.totalGroups],
      ['Total Candidates', stats.totalCandidates],
      ['Total Unique Courses', stats.totalUniqueCourses],
      [''],
      ['SL NO', 'PROGRAMME', 'VENUE', 'CANDIDATE COUNT']
    ];

    groupKeys.forEach((gKey, idx) => {
      const g = processedGroups[gKey];
      summaryRows.push([idx + 1, g.programme, g.venueLabel, g.totalCandidates]);
    });

    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

    // Tab per Venue / Group
    groupKeys.forEach((gKey, gIdx) => {
      const g = processedGroups[gKey];
      const rowsData = g.candidates.map((c, cIdx) => ({
        'Sl_No': cIdx + 1,
        'Register_Number': c.seatNo,
        'Candidate_Name': c.studentName,
        'Courses': c.courses.map(crs => crs.display).join('; '),
        'Remarks': '',
        'Programme': g.programme,
        'Venue': g.venueLabel
      }));

      const ws = XLSX.utils.json_to_sheet(rowsData);
      const sheetName = `Group_${gIdx + 1}_${g.venueCode || 'V'}`.slice(0, 30);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    XLSX.writeFile(wb, 'Master_Venue_Nominal_Roll.xlsx');
    setStatus('Master Excel downloaded successfully!', 'success');
  };

  return (
    <div style={{ maxWidth: '1440px', margin: '0 auto', padding: '24px 20px 80px', fontFamily: 'var(--font-family)' }}>
      
      {/* Top Navbar & Actions Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '16px' }}>
        <div>
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13.5px', marginBottom: '6px' }}>
            ← Back to Portal
          </Link>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div style={{ background: 'var(--accent)', color: 'white', padding: '10px', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <Building2 size={24} />
            </div>
            <div>
              <h1 style={{ margin: 0, fontSize: '24px', fontWeight: 800 }}>Venue-Wise Nominal Roll</h1>
              <p style={{ margin: '4px 0 0 0', color: 'var(--muted)', fontSize: '13.5px' }}>
                Compile nominal roll lists grouped by unique combinations of Programme + Venue with multi-subject merging.
              </p>
            </div>
          </div>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
          {groupKeys.length > 0 && (
            <>
              <button
                className="button"
                onClick={exportAllGroupsZip}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '9px 16px' }}
              >
                <Archive size={15} /> Download Complete Report (.zip)
              </button>

              <button
                className="button secondary"
                onClick={exportConsolidatedPdf}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '9px 16px' }}
              >
                <FileText size={15} /> Consolidated Master PDF
              </button>

              <button
                className="button secondary"
                onClick={exportMasterExcel}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '9px 16px' }}
              >
                <FileSpreadsheet size={15} /> Master Excel (.xlsx)
              </button>

              <button
                className="button secondary"
                onClick={() => setShowConfig(!showConfig)}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '9px 14px' }}
              >
                <Sliders size={15} /> {showConfig ? 'Hide Settings' : 'Settings'}
              </button>

              <button
                className="button secondary"
                onClick={handleReset}
                title="Clear and Upload New File"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '9px 14px' }}
              >
                <RotateCcw size={15} /> Reset
              </button>
            </>
          )}

          <span style={{ 
            fontSize: '12px', 
            fontWeight: 600, 
            padding: '5px 12px', 
            borderRadius: '20px', 
            background: statusType === 'success' ? 'rgba(16,185,129,0.12)' : statusType === 'error' ? 'rgba(239,68,68,0.12)' : 'var(--panel)',
            color: statusType === 'success' ? '#059669' : statusType === 'error' ? 'var(--danger)' : 'var(--muted)',
            border: '1px solid var(--line)'
          }}>
            {statusMsg}
          </span>
        </div>
      </div>

      {/* Upload Zone (Shown When No Data Loaded or in Settings) */}
      {groupKeys.length === 0 && (
        <div className="card" style={{ padding: '36px', marginBottom: '24px' }}>
          <div style={{ border: '2px dashed var(--line)', borderRadius: '12px', padding: '40px 20px', textAlign: 'center', background: 'var(--bg)' }}>
            <UploadCloud size={44} color="var(--accent)" style={{ margin: '0 auto 12px', opacity: 0.8 }} />
            <div style={{ fontWeight: 800, fontSize: '17px', marginBottom: '6px' }}>
              Upload Nominal Roll Spreadsheet or ZIP
            </div>
            <p style={{ color: 'var(--muted)', fontSize: '13.5px', maxWidth: '640px', margin: '0 auto 20px', lineHeight: '1.5' }}>
              Upload one or more <strong>Pre-exam / SLL nominal roll</strong> files (Excel <code>.xlsx, .xls, .xlsm, .csv</code>) or a <code>.zip</code> containing multiple spreadsheets. All files will be merged and mapped before processing.
            </p>
            
            <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', gap: '14px', flexWrap: 'wrap' }}>
              <label className="button" style={{ display: 'inline-flex', alignItems: 'center', gap: '8px', cursor: isProcessing ? 'wait' : 'pointer', padding: '11px 24px', fontSize: '14px', fontWeight: 700 }}>
                {isProcessing ? <RefreshCw size={16} className="spin" /> : <UploadCloud size={16} />}
                {isProcessing ? 'Reading & Extracting...' : 'Select Excel / ZIP File'}
                <input type="file" accept=".xlsx, .xls, .xlsm, .csv, .zip" onChange={handleFileUpload} disabled={isProcessing} style={{ display: 'none' }} />
              </label>

              <button 
                type="button"
                onClick={loadSampleData}
                disabled={isProcessing}
                className="button secondary"
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '11px 20px', fontSize: '14px' }}
              >
                <Sparkles size={16} /> Load Sample Dataset
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Collapsible Settings & Column Header Mappings Card */}
      {(showConfig || (headers.length > 0 && groupKeys.length === 0)) && (
        <div className="card" style={{ padding: '28px', marginBottom: '24px', background: 'var(--panel)', border: '1.5px solid var(--accent)' }}>
          <h3 style={{ margin: '0 0 16px 0', fontSize: '17px', color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '8px' }}>
            <Sliders size={18} /> Header Configurations & Column Header Mappings
          </h3>

          {/* Section 1: Title Customization */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))', gap: '14px', marginBottom: '22px', paddingBottom: '20px', borderBottom: '1px solid var(--line)' }}>
            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>University Name:</label>
              <input type="text" value={universityName} onChange={(e) => setUniversityName(e.target.value)} style={{ width: '100%', fontSize: '13px' }} />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Branch / Department:</label>
              <input type="text" value={branchName} onChange={(e) => setBranchName(e.target.value)} style={{ width: '100%', fontSize: '13px' }} />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Examination Title:</label>
              <input type="text" value={examTitle} onChange={(e) => setExamTitle(e.target.value)} style={{ width: '100%', fontSize: '13px' }} />
            </div>
            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Session / Month Year:</label>
              <input type="text" value={sessionName} onChange={(e) => setSessionName(e.target.value)} style={{ width: '100%', fontSize: '13px' }} />
            </div>
          </div>

          {/* Section 2: Column Dropdown Mappings */}
          <div style={{ fontSize: '13.5px', fontWeight: 700, marginBottom: '12px', color: 'var(--ink)' }}>
            Map Excel Columns to Nominal Roll Fields:
          </div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '14px' }}>
            
            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Programme Column (A)</label>
              <select value={columnMapping.programme} onChange={(e) => setColumnMapping({ ...columnMapping, programme: parseInt(e.target.value, 10) })} style={{ width: '100%', fontSize: '13px' }}>
                <option value="-1">-- Auto / Unspecified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Venue Code Column</label>
              <select value={columnMapping.venueCode} onChange={(e) => setColumnMapping({ ...columnMapping, venueCode: parseInt(e.target.value, 10) })} style={{ width: '100%', fontSize: '13px' }}>
                <option value="-1">-- Auto / Unspecified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Venue Name Column</label>
              <select value={columnMapping.venueName} onChange={(e) => setColumnMapping({ ...columnMapping, venueName: parseInt(e.target.value, 10) })} style={{ width: '100%', fontSize: '13px' }}>
                <option value="-1">-- Auto / Unspecified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Register Number Column</label>
              <select value={columnMapping.seatNo} onChange={(e) => setColumnMapping({ ...columnMapping, seatNo: parseInt(e.target.value, 10) })} style={{ width: '100%', fontSize: '13px' }}>
                <option value="-1">-- Auto / Unspecified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Candidate Name Column</label>
              <select value={columnMapping.name} onChange={(e) => setColumnMapping({ ...columnMapping, name: parseInt(e.target.value, 10) })} style={{ width: '100%', fontSize: '13px' }}>
                <option value="-1">-- Auto / Unspecified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Course Code Column</label>
              <select value={columnMapping.courseCode} onChange={(e) => setColumnMapping({ ...columnMapping, courseCode: parseInt(e.target.value, 10) })} style={{ width: '100%', fontSize: '13px' }}>
                <option value="-1">-- Auto / Unspecified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Course Title Column</label>
              <select value={columnMapping.courseTitle} onChange={(e) => setColumnMapping({ ...columnMapping, courseTitle: parseInt(e.target.value, 10) })} style={{ width: '100%', fontSize: '13px' }}>
                <option value="-1">-- Auto / Unspecified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, marginBottom: '4px' }}>Group By Strategy</label>
              <select value={groupByOption} onChange={(e) => setGroupByOption(e.target.value)} style={{ width: '100%', fontSize: '13px' }}>
                <option value="programme_venue">Programme + Venue (Default)</option>
                <option value="venue">Venue Code & Name Only</option>
              </select>
            </div>

          </div>
        </div>
      )}

      {/* Main Content Area (When Data Loaded) */}
      {groupKeys.length > 0 && (
        <div>
          
          {/* Executive KPI Summary Cards */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px', marginBottom: '24px' }}>
            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'var(--accent-soft)', padding: '12px', borderRadius: '10px', color: 'var(--accent)' }}>
                <MapPin size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>Total Venues</div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{stats.totalVenues}</div>
              </div>
            </div>

            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'rgba(23, 107, 135, 0.15)', padding: '12px', borderRadius: '10px', color: 'var(--accent)' }}>
                <GraduationCap size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>Total Groups</div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{stats.totalGroups}</div>
              </div>
            </div>

            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'rgba(16, 185, 129, 0.15)', padding: '12px', borderRadius: '10px', color: '#10b981' }}>
                <Users size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>Total Candidates</div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{stats.totalCandidates}</div>
              </div>
            </div>

            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'rgba(139, 92, 246, 0.15)', padding: '12px', borderRadius: '10px', color: '#8b5cf6' }}>
                <BookOpen size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>Unique Courses</div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{stats.totalUniqueCourses}</div>
              </div>
            </div>
          </div>

          {/* Venue & Programme Navigation Toolbar */}
          <div className="card" style={{ padding: '18px 24px', marginBottom: '20px', background: 'var(--panel)', border: '1px solid var(--line)', borderRadius: '10px' }}>
            
            {/* Top Row: Programme & Venue Selectors */}
            <div style={{ display: 'grid', gridTemplateColumns: groupByOption === 'programme_venue' && programmesList.length > 1 ? 'repeat(auto-fit, minmax(320px, 1fr))' : '1fr', gap: '16px', marginBottom: '14px' }}>
              
              {groupByOption === 'programme_venue' && programmesList.length > 1 && (
                <div>
                  <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', fontWeight: 700, color: 'var(--ink)', marginBottom: '6px' }}>
                    <GraduationCap size={15} color="var(--accent)" /> 1. Select Programme ({programmesList.length}):
                  </label>
                  <select
                    value={currentProgramme}
                    onChange={(e) => {
                      const newProg = e.target.value;
                      const firstMatchingKey = groupKeys.find(k => processedGroups[k]?.programme === newProg);
                      if (firstMatchingKey) setActiveGroupKey(firstMatchingKey);
                    }}
                    style={{ width: '100%', padding: '9px 12px', fontSize: '13px', fontWeight: 600, borderRadius: '8px', border: '1.5px solid var(--line)', background: 'white' }}
                  >
                    {programmesList.map((prog, pIdx) => {
                      const vCount = groupKeys.filter(k => processedGroups[k]?.programme === prog).length;
                      return (
                        <option key={pIdx} value={prog}>
                          {prog} ({vCount} Venues)
                        </option>
                      );
                    })}
                  </select>
                </div>
              )}

              <div>
                <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', fontWeight: 700, color: 'var(--ink)', marginBottom: '6px' }}>
                  <MapPin size={15} color="var(--accent)" /> {groupByOption === 'programme_venue' && programmesList.length > 1 ? '2. Select Examination Venue' : 'Select Examination Venue'} ({venuesForCurrentProg.length}):
                </label>
                <select
                  value={effectiveGroupKey}
                  onChange={(e) => setActiveGroupKey(e.target.value)}
                  style={{ width: '100%', padding: '9px 12px', fontSize: '13px', fontWeight: 600, borderRadius: '8px', border: '1.5px solid var(--accent)', background: 'white' }}
                >
                  {venuesForCurrentProg.map((gKey, vIdx) => {
                    const g = processedGroups[gKey];
                    return (
                      <option key={gKey} value={gKey}>
                        {vIdx + 1}. {g.venueLabel} ({g.totalCandidates} Candidates)
                      </option>
                    );
                  })}
                </select>
              </div>

            </div>

            {/* Bottom Row: Quick Actions, Search, Prev/Next & Download */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px', paddingTop: '12px', borderTop: '1px solid var(--line)' }}>
              <div style={{ position: 'relative', width: '100%', maxWidth: '300px' }}>
                <Search size={15} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
                <input 
                  type="text"
                  placeholder="Search candidates, register no..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  style={{ paddingLeft: '32px', fontSize: '13px', width: '100%' }}
                />
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <span style={{ fontSize: '12.5px', color: 'var(--muted)', fontWeight: 600 }}>
                  Venue {groupKeys.indexOf(effectiveGroupKey) + 1} of {groupKeys.length}
                </span>

                <button
                  className="button secondary"
                  disabled={groupKeys.indexOf(effectiveGroupKey) <= 0}
                  onClick={() => {
                    const currIdx = groupKeys.indexOf(effectiveGroupKey);
                    if (currIdx > 0) setActiveGroupKey(groupKeys[currIdx - 1]);
                  }}
                  style={{ padding: '7px 12px', fontSize: '12.5px' }}
                >
                  ← Prev
                </button>
                <button
                  className="button secondary"
                  disabled={groupKeys.indexOf(effectiveGroupKey) >= groupKeys.length - 1}
                  onClick={() => {
                    const currIdx = groupKeys.indexOf(effectiveGroupKey);
                    if (currIdx < groupKeys.length - 1) setActiveGroupKey(groupKeys[currIdx + 1]);
                  }}
                  style={{ padding: '7px 12px', fontSize: '12.5px' }}
                >
                  Next →
                </button>

                <button
                  className="button secondary"
                  onClick={() => exportSingleGroupPdf(effectiveGroupKey)}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '7px 14px' }}
                >
                  <Download size={15} /> Download Venue PDF
                </button>

                <button
                  className="button secondary"
                  onClick={() => window.print()}
                  title="Direct Print A4 from browser"
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '7px 14px' }}
                >
                  <Printer size={15} /> Print A4
                </button>

                <button
                  className="button"
                  onClick={exportAllGroupsZip}
                  disabled={isProcessing}
                  style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', padding: '7px 16px' }}
                >
                  <Archive size={15} /> Download Complete Report (.zip)
                </button>
              </div>
            </div>

          </div>

          {/* Live Preview Paper Display (Matches PDF Layout 1:1) */}
          <div className="card print-container" style={{ padding: '36px 40px', background: 'white', border: '1px solid var(--line)', borderRadius: '8px', boxShadow: '0 4px 20px rgba(0,0,0,0.06)' }}>
            
            {/* Rendered Header inside Paper */}
            <div style={{ textAlign: 'center', marginBottom: '24px', borderBottom: '1px solid #eee', paddingBottom: '16px' }}>
              <div style={{ fontSize: '20px', fontWeight: 800, color: '#000', letterSpacing: '0.2px', marginBottom: '4px' }}>
                {universityName}
              </div>
              <div style={{ fontSize: '15px', fontWeight: 700, color: '#111', marginBottom: '6px' }}>
                {branchName}
              </div>
              <div style={{ fontSize: '14.5px', fontWeight: 700, color: '#111', marginBottom: '4px' }}>
                {examTitle}
              </div>
              {sessionName && (
                <div style={{ fontSize: '14px', fontWeight: 600, color: '#222', marginBottom: '6px' }}>
                  {sessionName}
                </div>
              )}
              
              {/* Separate and Stacked Programme & Venue Box */}
              <div style={{ margin: '14px auto 0', maxWidth: '850px', background: '#f8fafc', padding: '12px 18px', borderRadius: '8px', border: '1px solid #e2e8f0', textAlign: 'center' }}>
                <div style={{ fontSize: '13.5px', color: '#1e293b', marginBottom: '6px', lineHeight: '1.4' }}>
                  <span style={{ fontWeight: 800, color: '#0f172a', textTransform: 'uppercase', fontSize: '11px', letterSpacing: '0.5px', background: '#e2e8f0', padding: '2px 8px', borderRadius: '4px', marginRight: '8px' }}>
                    Programme
                  </span>
                  <strong>{processedGroups[effectiveGroupKey]?.programme || 'N/A'}</strong>
                </div>
                <div style={{ fontSize: '14px', color: '#0f172a', lineHeight: '1.4' }}>
                  <span style={{ fontWeight: 800, color: '#0f172a', textTransform: 'uppercase', fontSize: '11px', letterSpacing: '0.5px', background: '#e2e8f0', padding: '2px 8px', borderRadius: '4px', marginRight: '8px' }}>
                    Venue
                  </span>
                  <strong style={{ color: 'var(--accent)' }}>{processedGroups[effectiveGroupKey]?.venueLabel || 'N/A'}</strong>
                </div>
              </div>
            </div>

            {/* Rendered Nominal Roll Table */}
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', fontFamily: 'inherit' }}>
                <thead>
                  <tr style={{ background: '#176b87', color: '#ffffff' }}>
                    <th style={{ border: '1px solid #0f4c5c', padding: '10px 6px', width: '55px', textAlign: 'center', fontWeight: 700, verticalAlign: 'middle' }}>SL<br/>No</th>
                    <th style={{ border: '1px solid #0f4c5c', padding: '10px 10px', width: '140px', textAlign: 'center', fontWeight: 700, verticalAlign: 'middle' }}>Register Number</th>
                    <th style={{ border: '1px solid #0f4c5c', padding: '10px 12px', width: '200px', textAlign: 'left', fontWeight: 700, verticalAlign: 'middle' }}>Candidate Name</th>
                    <th style={{ border: '1px solid #0f4c5c', padding: '10px 14px', textAlign: 'left', fontWeight: 700, verticalAlign: 'middle' }}>Courses</th>
                    <th style={{ border: '1px solid #0f4c5c', padding: '10px 8px', width: '120px', textAlign: 'center', fontWeight: 700, verticalAlign: 'middle' }}>Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  {currentTabCandidates.length > 0 ? (
                    currentTabCandidates.map((cand, idx) => (
                      <tr key={idx} style={{ borderBottom: '1px solid #cbd5e1', background: idx % 2 === 0 ? '#fff' : '#f8fafc' }}>
                        <td style={{ border: '1px solid #cbd5e1', padding: '10px 6px', textAlign: 'center', fontWeight: 700, verticalAlign: 'middle', fontSize: '13px', color: '#334155' }}>
                          {idx + 1}
                        </td>
                        <td style={{ border: '1px solid #cbd5e1', padding: '10px 10px', textAlign: 'center', fontWeight: 800, color: 'var(--accent)', verticalAlign: 'middle', fontFamily: 'monospace, monospace', fontSize: '13.5px', letterSpacing: '0.5px' }}>
                          {cand.seatNo}
                        </td>
                        <td style={{ border: '1px solid #cbd5e1', padding: '10px 12px', fontWeight: 700, verticalAlign: 'middle', textTransform: 'uppercase', color: '#0f172a', fontSize: '13px' }}>
                          {cand.studentName}
                        </td>
                        <td style={{ border: '1px solid #cbd5e1', padding: '10px 14px', verticalAlign: 'middle' }}>
                          {cand.courses.length > 0 ? (
                            <div style={{ display: 'flex', flexDirection: 'column', gap: '5px' }}>
                              {cand.courses.map((crs, cIdx) => (
                                <div key={cIdx} style={{ fontSize: '12.5px', lineHeight: '1.45', color: '#1e293b' }}>
                                  <span style={{ fontWeight: 700, color: 'var(--accent)', marginRight: '6px' }}>
                                    {cIdx + 1}. {crs.code}
                                  </span>
                                  {crs.title && <span style={{ color: '#334155' }}>— {crs.title}</span>}
                                </div>
                              ))}
                            </div>
                          ) : (
                            <span style={{ color: '#94a3b8' }}>—</span>
                          )}
                        </td>
                        <td style={{ border: '1px solid #cbd5e1', padding: '10px 8px', verticalAlign: 'middle', background: 'rgba(248, 250, 252, 0.5)' }}>
                          {/* Blank for Remarks */}
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={5} style={{ textAlign: 'center', padding: '30px', color: 'var(--muted)' }}>
                        No records matching query "{searchQuery}"
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Table Footer / Summary */}
            <div style={{ marginTop: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', fontSize: '13px', color: 'var(--muted)' }}>
              <div>
                Showing <strong>{currentTabCandidates.length}</strong> candidate(s) for <strong>{effectiveGroupKey}</strong>
              </div>
              <div style={{ fontWeight: 700, color: 'var(--ink)' }}>
                Venue Total: {processedGroups[effectiveGroupKey]?.totalCandidates} Candidates
              </div>
            </div>

          </div>

        </div>
      )}

    </div>
  );
};

export default SllNominalPage;
