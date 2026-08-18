import { useState, useMemo } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import JSZip from 'jszip';
import { 
  Settings, 
  Download, 
  FileText, 
  HelpCircle, 
  UploadCloud, 
  FileSpreadsheet, 
  Archive, 
  Printer, 
  CheckCircle2, 
  Search, 
  Calendar, 
  BookOpen, 
  Users,
  Sliders,
  MapPin
} from 'lucide-react';
import { readSpreadsheetAsAoa } from '../utils/excelParser';

const QpStatementPage = () => {

  // Upload and File states
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready to upload Excel files.');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Header and Configuration settings (Venue-Wise by Default)
  const [universityName, setUniversityName] = useState('Kannur University');
  const [branchName, setBranchName] = useState('(Examination Branch)');
  const [examTitle, setExamTitle] = useState('QP Statement for 1st Semester Degree Private Registration Regular/Supplementary Examination');
  const [sessionName, setSessionName] = useState('November 2025');
  const [centerPrefix, setCenterPrefix] = useState('Venue :'); // Default to 'Venue :'
  const [groupByOption, setGroupByOption] = useState('venue'); // Default to 'venue'
  const [courseCodeDot, setCourseCodeDot] = useState(true); // "KU3MDCARS202. - Kerala Culture in Arabic Narratives."
  const [aggregateSameCourse, setAggregateSameCourse] = useState(true); // Sum Student Count if same course & date

  // Column mapping (indices)
  const [columnMapping, setColumnMapping] = useState({
    eventName: -1,
    venueCode: -1,
    venueName: -1,
    centerCode: -1,
    centerName: -1,
    programName: -1,
    examDate: -1,
    courseStartTime: -1,
    courseEndTime: -1,
    courseCode: -1,
    courseName: -1,
    studentCount: -1,
    venueId: -1
  });

  // UI Navigation & Filters
  const [activeGroupKey, setActiveGroupKey] = useState('');
  const [searchQuery, setSearchQuery] = useState('');
  const [showConfig, setShowConfig] = useState(false);

  const setStatus = (msg, type = 'normal') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  // Helper to format date with Day of Week
  const formatDateWithDay = (dateVal) => {
    if (!dateVal) return '';
    let dateObj;

    if (dateVal instanceof Date && !isNaN(dateVal)) {
      dateObj = dateVal;
    } else if (typeof dateVal === 'number') {
      // Excel serial date format
      dateObj = new Date(Math.round((dateVal - 25569) * 86400 * 1000));
    } else {
      const strVal = String(dateVal).trim();
      // Try YYYY-MM-DD
      if (/^\d{4}-\d{2}-\d{2}/.test(strVal)) {
        const parts = strVal.split(/[-T\s]/);
        dateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      } else if (/^\d{1,2}[/-]\d{1,2}[/-]\d{4}/.test(strVal)) {
        // DD/MM/YYYY or DD-MM-YYYY
        const parts = strVal.split(/[/-]/);
        dateObj = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
      } else {
        const parsed = new Date(strVal);
        if (!isNaN(parsed.getTime())) {
          dateObj = parsed;
        } else {
          return strVal;
        }
      }
    }

    if (!dateObj || isNaN(dateObj.getTime())) {
      return String(dateVal).trim();
    }

    const yyyy = dateObj.getFullYear();
    const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    const dd = String(dateObj.getDate()).padStart(2, '0');
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const dayName = days[dateObj.getDay()];

    return `${yyyy}-${mm}-${dd} ${dayName}`;
  };

  // Raw date key for sorting (YYYY-MM-DD)
  const getSortableDate = (dateVal) => {
    if (!dateVal) return '9999-99-99';
    let dateObj;
    if (dateVal instanceof Date && !isNaN(dateVal)) {
      dateObj = dateVal;
    } else if (typeof dateVal === 'number') {
      dateObj = new Date(Math.round((dateVal - 25569) * 86400 * 1000));
    } else {
      const strVal = String(dateVal).trim();
      if (/^\d{4}-\d{2}-\d{2}/.test(strVal)) {
        const parts = strVal.split(/[-T\s]/);
        dateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      } else if (/^\d{1,2}[/-]\d{1,2}[/-]\d{4}/.test(strVal)) {
        const parts = strVal.split(/[/-]/);
        dateObj = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
      } else {
        const parsed = new Date(strVal);
        if (!isNaN(parsed.getTime())) {
          dateObj = parsed;
        }
      }
    }
    if (!dateObj || isNaN(dateObj.getTime())) return String(dateVal);
    const yyyy = dateObj.getFullYear();
    const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
    const dd = String(dateObj.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  };

  // Auto-detect column headers based on fuzzy keywords
  const autoDetectColumns = (headerList) => {
    const mapping = {
      eventName: -1,
      venueCode: -1,
      venueName: -1,
      centerCode: -1,
      centerName: -1,
      programName: -1,
      examDate: -1,
      courseStartTime: -1,
      courseEndTime: -1,
      courseCode: -1,
      courseName: -1,
      studentCount: -1,
      venueId: -1
    };

    headerList.forEach((colName, idx) => {
      const lower = String(colName).toLowerCase().replace(/[^a-z0-9]/g, '');

      if (lower.includes('eventname') || (lower.includes('event') && !lower.includes('id'))) mapping.eventName = idx;
      if (lower === 'venuecode') mapping.venueCode = idx;
      if (lower === 'venuename') mapping.venueName = idx;
      if (lower === 'centercode' || lower === 'centrecode') mapping.centerCode = idx;
      if (lower === 'centername' || lower === 'centrename') mapping.centerName = idx;
      if (lower.includes('programname') || lower.includes('programme')) mapping.programName = idx;
      if (lower.includes('examdate') || lower === 'date') mapping.examDate = idx;
      if (lower.includes('coursestarttime') || lower.includes('starttime')) mapping.courseStartTime = idx;
      if (lower.includes('courseendtime') || lower.includes('endtime')) mapping.courseEndTime = idx;
      if (lower === 'coursecode' || lower === 'qpcoursename' || lower.includes('subjectcode')) mapping.courseCode = idx;
      if (lower === 'coursename' || lower.includes('subjectname') || lower.includes('coursetitle')) mapping.courseName = idx;
      if (lower.includes('studentcount') || lower === 'count' || lower === 'nc' || lower.includes('candidatecount')) mapping.studentCount = idx;
      if (lower === 'venueid') mapping.venueId = idx;
    });

    // Fallbacks if Center vs Venue is singular
    if (mapping.venueCode === -1 && mapping.centerCode !== -1) mapping.venueCode = mapping.centerCode;
    if (mapping.venueName === -1 && mapping.centerName !== -1) mapping.venueName = mapping.centerName;
    if (mapping.centerCode === -1 && mapping.venueCode !== -1) mapping.centerCode = mapping.venueCode;
    if (mapping.centerName === -1 && mapping.venueName !== -1) mapping.centerName = mapping.venueName;

    return mapping;
  };

  // Handle uploaded files
  const handleFileChange = async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    setSelectedFiles(files);
    setIsProcessing(true);
    setStatus(`Reading ${files.length} file(s)...`, 'normal');

    try {
      const allFilesRows = await Promise.all(files.map(file => readSpreadsheetAsAoa(file)));
      let combinedData = [];
      let detectedHeaders = [];

      allFilesRows.forEach((fileRows) => {
        if (fileRows && fileRows.length > 0) {
          if (detectedHeaders.length === 0) {
            detectedHeaders = fileRows[0].map((cell, idx) => cell ? String(cell).trim() : `Column ${idx + 1}`);
            // Exclude header row from data
            combinedData = combinedData.concat(fileRows.slice(1));
          } else {
            // Append data rows
            combinedData = combinedData.concat(fileRows.slice(1));
          }
        }
      });

      if (detectedHeaders.length > 0) {
        setHeaders(detectedHeaders);
        setRawRows(combinedData);

        const newMapping = autoDetectColumns(detectedHeaders);
        setColumnMapping(newMapping);

        // Auto-populate Event Name from data if available
        if (newMapping.eventName !== -1 && combinedData.length > 0) {
          const sampleEvent = combinedData.find(r => r[newMapping.eventName])?.[newMapping.eventName];
          if (sampleEvent && String(sampleEvent).trim()) {
            const cleanEvent = String(sampleEvent).trim();
            if (!cleanEvent.toLowerCase().startsWith('qp statement for')) {
              setExamTitle(`QP Statement for ${cleanEvent}`);
            } else {
              setExamTitle(cleanEvent);
            }
          }
        }

        setStatus(`Successfully loaded ${combinedData.length} records from ${files.length} file(s).`, 'success');
      } else {
        setStatus('No data found in uploaded files.', 'error');
      }
    } catch (err) {
      console.error(err);
      setStatus(`Error reading files: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Group and Process data into structured Venue / Center sets
  const processedGroups = useMemo(() => {
    if (rawRows.length === 0) return {};

    const groups = {};

    rawRows.forEach(row => {
      // Determine Venue / Center identifiers (Venue Priority)
      let code;
      let name;

      if (groupByOption === 'venue') {
        const vCode = columnMapping.venueCode !== -1 && row[columnMapping.venueCode] ? String(row[columnMapping.venueCode]).trim() : '';
        const vName = columnMapping.venueName !== -1 && row[columnMapping.venueName] ? String(row[columnMapping.venueName]).trim() : '';
        code = vCode || (columnMapping.centerCode !== -1 ? String(row[columnMapping.centerCode] || '').trim() : '');
        name = vName || (columnMapping.centerName !== -1 ? String(row[columnMapping.centerName] || '').trim() : '');
      } else {
        const cCode = columnMapping.centerCode !== -1 && row[columnMapping.centerCode] ? String(row[columnMapping.centerCode]).trim() : '';
        const cName = columnMapping.centerName !== -1 && row[columnMapping.centerName] ? String(row[columnMapping.centerName]).trim() : '';
        code = cCode || (columnMapping.venueCode !== -1 ? String(row[columnMapping.venueCode] || '').trim() : '');
        name = cName || (columnMapping.venueName !== -1 ? String(row[columnMapping.venueName] || '').trim() : '');
      }

      if (!code && !name) {
        code = 'DEFAULT';
        name = 'General Examination Venue';
      }

      const groupKey = code && name ? `${code} - ${name}` : (code || name);

      if (!groups[groupKey]) {
        groups[groupKey] = {
          code: code,
          name: name,
          fullLabel: groupKey,
          items: []
        };
      }

      const dateRaw = columnMapping.examDate !== -1 ? row[columnMapping.examDate] : '';
      const formattedDate = formatDateWithDay(dateRaw);
      const sortDate = getSortableDate(dateRaw);

      let cCode = columnMapping.courseCode !== -1 && row[columnMapping.courseCode] ? String(row[columnMapping.courseCode]).trim() : '';
      let cName = columnMapping.courseName !== -1 && row[columnMapping.courseName] ? String(row[columnMapping.courseName]).trim() : '';
      
      const countVal = columnMapping.studentCount !== -1 ? parseInt(row[columnMapping.studentCount], 10) || 0 : 0;
      const startTime = columnMapping.courseStartTime !== -1 && row[columnMapping.courseStartTime] ? String(row[columnMapping.courseStartTime]).trim() : '';
      const endTime = columnMapping.courseEndTime !== -1 && row[columnMapping.courseEndTime] ? String(row[columnMapping.courseEndTime]).trim() : '';

      groups[groupKey].items.push({
        rawDate: dateRaw,
        formattedDate,
        sortDate,
        courseCode: cCode,
        courseName: cName,
        studentCount: countVal,
        startTime,
        endTime
      });
    });

    // Aggregate / Sort within each group
    Object.keys(groups).forEach(key => {
      let items = groups[key].items;

      if (aggregateSameCourse) {
        const mergedMap = new Map();
        items.forEach(item => {
          // Normalize course code and name by removing trailing dots, trimming whitespace, and ignoring case
          const cleanCode = item.courseCode.replace(/\.+$/, '').replace(/\s+/g, ' ').trim().toLowerCase();
          const cleanName = item.courseName.replace(/\.+$/, '').replace(/\s+/g, ' ').trim().toLowerCase();
          const mKey = `${item.sortDate}__${cleanCode}__${cleanName}`;

          if (mergedMap.has(mKey)) {
            const existing = mergedMap.get(mKey);
            existing.studentCount += item.studentCount;
          } else {
            // Keep clean normalized base representation
            mergedMap.set(mKey, {
              ...item,
              courseCode: item.courseCode.replace(/\.+$/, '').trim(),
              courseName: item.courseName.replace(/\.+$/, '').trim()
            });
          }
        });
        items = Array.from(mergedMap.values());
      }

      // Sort chronologically by date, then course code
      items.sort((a, b) => {
        if (a.sortDate !== b.sortDate) {
          return a.sortDate.localeCompare(b.sortDate);
        }
        if (a.startTime && b.startTime && a.startTime !== b.startTime) {
          return a.startTime.localeCompare(b.startTime);
        }
        return a.courseCode.localeCompare(b.courseCode);
      });

      groups[key].items = items;
    });

    return groups;
  }, [rawRows, columnMapping, groupByOption, aggregateSameCourse]);

  // Derived effective group key
  const groupKeys = useMemo(() => Object.keys(processedGroups), [processedGroups]);
  const effectiveGroupKey = (activeGroupKey && processedGroups[activeGroupKey]) ? activeGroupKey : (groupKeys[0] || '');

  // Overall Statistics
  const overallStats = useMemo(() => {
    let totalExams = 0;
    let totalCandidates = 0;
    const allDates = new Set();

    groupKeys.forEach(k => {
      const items = processedGroups[k].items;
      totalExams += items.length;
      items.forEach(it => {
        totalCandidates += it.studentCount;
        if (it.sortDate) allDates.add(it.sortDate);
      });
    });

    const sortedDates = Array.from(allDates).sort();
    const dateRange = sortedDates.length > 0 
      ? `${sortedDates[0]} to ${sortedDates[sortedDates.length - 1]}`
      : 'N/A';

    return {
      totalGroups: groupKeys.length,
      totalExams,
      totalCandidates,
      dateRange,
      uniqueDays: sortedDates.length
    };
  }, [processedGroups, groupKeys]);

  // Filtered rows for current active tab
  const currentTabItems = useMemo(() => {
    if (!effectiveGroupKey || !processedGroups[effectiveGroupKey]) return [];
    const items = processedGroups[effectiveGroupKey].items;
    if (!searchQuery.trim()) return items;

    const q = searchQuery.toLowerCase();
    return items.filter(it => 
      it.formattedDate.toLowerCase().includes(q) ||
      it.courseCode.toLowerCase().includes(q) ||
      it.courseName.toLowerCase().includes(q) ||
      String(it.studentCount).includes(q)
    );
  }, [effectiveGroupKey, processedGroups, searchQuery]);

  // Helper to format course display text e.g. "KU3MDCARS202. - Kerala Culture in Arabic Narratives."
  const formatCourseDisplay = (courseCode, courseName) => {
    let codeStr = String(courseCode || '').trim();
    let nameStr = String(courseName || '').trim();

    // Strip existing trailing dots first
    codeStr = codeStr.replace(/\.+$/, '').trim();
    nameStr = nameStr.replace(/\.+$/, '').trim();

    if (courseCodeDot) {
      if (codeStr) codeStr += '.';
      if (nameStr) nameStr += '.';
    }

    if (codeStr && nameStr) {
      return `${codeStr} - ${nameStr}`;
    }
    return codeStr || nameStr || '-';
  };

  // Draw Page Header on PDF document
  const drawPdfHeader = (doc, groupLabel) => {
    const pageWidth = doc.internal.pageSize.getWidth();
    let currentY = 32;

    // 1. University Name
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(15);
    doc.setTextColor(0, 0, 0);
    doc.text(universityName, pageWidth / 2, currentY, { align: 'center' });

    // 2. Branch Name
    currentY += 18;
    doc.setFontSize(11);
    doc.setFont('helvetica', 'bold');
    doc.text(branchName, pageWidth / 2, currentY, { align: 'center' });

    // 3. Exam Title / Event
    if (examTitle) {
      currentY += 16;
      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'bold');
      doc.text(examTitle, pageWidth / 2, currentY, { align: 'center' });
    }

    // 4. Session / Month Year
    if (sessionName) {
      currentY += 14;
      doc.setFontSize(9.5);
      doc.setFont('helvetica', 'bold');
      doc.text(sessionName, pageWidth / 2, currentY, { align: 'center' });
    }

    // 5. Venue / Center Name (e.g. Venue : KI - Concord Arts and Science College, Muttannur)
    currentY += 17;
    doc.setFontSize(10.5);
    doc.setFont('helvetica', 'bold');
    const headerPrefixText = `${centerPrefix} ${groupLabel}`;
    doc.text(headerPrefixText, pageWidth / 2, currentY, { align: 'center' });

    return currentY + 12; // Return bottom Y for autotable start
  };

  // Generate Single Venue PDF
  const generateVenuePdf = (groupKey, shouldSave = true) => {
    const group = processedGroups[groupKey];
    if (!group) return null;

    const doc = new jsPDF({
      orientation: 'portrait',
      unit: 'pt',
      format: 'a4'
    });

    const startTableY = drawPdfHeader(doc, group.fullLabel);

    const tableBody = group.items.map((item, idx) => {
      const courseText = formatCourseDisplay(item.courseCode, item.courseName);
      return [
        idx + 1,
        item.formattedDate,
        courseText,
        item.studentCount,
        '', // QP column empty for manual entry / verify
        ''  // LP column empty for manual entry / verify
      ];
    });

    autoTable(doc, {
      startY: startTableY,
      head: [['SL\nNo', 'Date', 'Course', 'NC', 'QP', 'LP']],
      body: tableBody,
      theme: 'grid',
      margin: { top: 30, left: 36, right: 36, bottom: 36 },
      styles: {
        fontSize: 8.5,
        textColor: [0, 0, 0],
        lineColor: [0, 0, 0],
        lineWidth: 0.65,
        cellPadding: { top: 5, right: 5, bottom: 5, left: 5 },
        valign: 'middle'
      },
      headStyles: {
        fillColor: [240, 240, 240],
        textColor: [0, 0, 0],
        fontStyle: 'bold',
        halign: 'center',
        valign: 'middle',
        fontSize: 9,
        lineColor: [0, 0, 0],
        lineWidth: 0.8
      },
      columnStyles: {
        0: { halign: 'center', cellWidth: 32 },
        1: { halign: 'left', cellWidth: 105 },
        2: { halign: 'left', cellWidth: 245 },
        3: { halign: 'center', cellWidth: 42 },
        4: { halign: 'center', cellWidth: 48 },
        5: { halign: 'center', cellWidth: 48 }
      }
    });

    const cleanFilename = `QP_Statement_Venue_${group.fullLabel.replace(/[^a-zA-Z0-9_-]/g, '_')}.pdf`;
    if (shouldSave) {
      doc.save(cleanFilename);
    }
    return { doc, filename: cleanFilename };
  };

  // Download All Venues as ZIP
  const downloadAllZip = async () => {
    const groupKeys = Object.keys(processedGroups);
    if (groupKeys.length === 0) return;

    setStatus('Generating ZIP archive with all Venue Statements...', 'normal');
    setIsProcessing(true);

    try {
      const zip = new JSZip();

      groupKeys.forEach(groupKey => {
        const result = generateVenuePdf(groupKey, false);
        if (result) {
          const pdfBlob = result.doc.output('blob');
          zip.file(result.filename, pdfBlob);
        }
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(zipBlob);
      link.download = `QP_Statements_All_Venues_${new Date().toISOString().slice(0, 10)}.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      setStatus(`Successfully exported ZIP containing ${groupKeys.length} venue statement(s)!`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to generate ZIP: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Download All Venues Combined into a single Master PDF
  const downloadCombinedMasterPdf = () => {
    const groupKeys = Object.keys(processedGroups);
    if (groupKeys.length === 0) return;

    setStatus('Compiling All Venues into Master PDF...', 'normal');
    setIsProcessing(true);

    try {
      const masterDoc = new jsPDF({
        orientation: 'portrait',
        unit: 'pt',
        format: 'a4'
      });

      groupKeys.forEach((groupKey, groupIdx) => {
        if (groupIdx > 0) {
          masterDoc.addPage();
        }

        const group = processedGroups[groupKey];
        const startTableY = drawPdfHeader(masterDoc, group.fullLabel);

        const tableBody = group.items.map((item, idx) => [
          idx + 1,
          item.formattedDate,
          formatCourseDisplay(item.courseCode, item.courseName),
          item.studentCount,
          '',
          ''
        ]);

        autoTable(masterDoc, {
          startY: startTableY,
          head: [['SL\nNo', 'Date', 'Course', 'NC', 'QP', 'LP']],
          body: tableBody,
          theme: 'grid',
          margin: { top: 30, left: 36, right: 36, bottom: 36 },
          styles: {
            fontSize: 8.5,
            textColor: [0, 0, 0],
            lineColor: [0, 0, 0],
            lineWidth: 0.65,
            cellPadding: { top: 5, right: 5, bottom: 5, left: 5 },
            valign: 'middle'
          },
          headStyles: {
            fillColor: [240, 240, 240],
            textColor: [0, 0, 0],
            fontStyle: 'bold',
            halign: 'center',
            valign: 'middle',
            fontSize: 9,
            lineColor: [0, 0, 0],
            lineWidth: 0.8
          },
          columnStyles: {
            0: { halign: 'center', cellWidth: 32 },
            1: { halign: 'left', cellWidth: 105 },
            2: { halign: 'left', cellWidth: 245 },
            3: { halign: 'center', cellWidth: 42 },
            4: { halign: 'center', cellWidth: 48 },
            5: { halign: 'center', cellWidth: 48 }
          }
        });
      });

      masterDoc.save(`QP_Statement_Consolidated_Master_Venues_${new Date().toISOString().slice(0, 10)}.pdf`);
      setStatus('Master Venue PDF successfully downloaded!', 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to generate Master PDF: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Export Consolidated Excel
  const exportConsolidatedExcel = () => {
    const groupKeys = Object.keys(processedGroups);
    if (groupKeys.length === 0) return;

    try {
      const rowsForExcel = [];

      groupKeys.forEach(groupKey => {
        const group = processedGroups[groupKey];
        group.items.forEach((item, idx) => {
          rowsForExcel.push({
            'Venue': group.fullLabel,
            'SL No': idx + 1,
            'Date': item.formattedDate,
            'Course Code': item.courseCode,
            'Course Name': item.courseName,
            'Full Course Details': formatCourseDisplay(item.courseCode, item.courseName),
            'Student Count (NC)': item.studentCount,
            'Start Time': item.startTime,
            'End Time': item.endTime
          });
        });
      });

      const ws = XLSX.utils.json_to_sheet(rowsForExcel);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'QP_Statement_Venue_Data');
      XLSX.writeFile(wb, `QP_Statement_Venues_Consolidated_${new Date().toISOString().slice(0, 10)}.xlsx`);

      setStatus('Consolidated Venue Excel report downloaded successfully!', 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to export Excel: ${err.message}`, 'error');
    }
  };

  // Open Direct Print / Preview
  const handlePrintPreview = () => {
    if (!activeGroupKey) return;
    const result = generateVenuePdf(activeGroupKey, false);
    if (result) {
      const blobUrl = result.doc.output('bloburl');
      window.open(blobUrl, '_blank');
    }
  };

  return (
    <div style={{ maxWidth: '1440px', margin: '0 auto', padding: '24px 20px 80px' }}>
      
      {/* Top Breadcrumb & Status Bar */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', flexWrap: 'wrap', gap: '12px' }}>
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
          border: '1px solid var(--line)',
          display: 'flex',
          alignItems: 'center',
          gap: '8px'
        }}>
          {statusType === 'success' && <CheckCircle2 size={16} />}
          {statusMsg}
        </div>
      </div>

      {/* Header Banner */}
      <div style={{ marginBottom: '28px' }}>
        <h2 style={{ fontSize: '28px', fontWeight: 700, margin: '0 0 6px 0', color: 'var(--ink)' }}>
          QP Statement Report (Venue-Wise)
        </h2>
        <p style={{ color: 'var(--muted)', margin: 0, fontSize: '14.5px' }}>
          Generate official University Question Paper (QP) statements, candidate counts (NC), and venue distribution packing lists grouped by <strong>Venue Code</strong>.
        </p>
      </div>

      {/* Upload & Setup Section */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(360px, 1fr))', gap: '24px', marginBottom: '28px' }}>
        
        {/* Upload Card */}
        <div className="card" style={{ padding: '24px', margin: 0, display: 'flex', flexDirection: 'column', gap: '16px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, fontSize: '17px', display: 'flex', alignItems: 'center', gap: '8px' }}>
              <UploadCloud size={20} color="var(--accent)" /> 1. Upload Source Excel / Report
            </h3>
            {selectedFiles.length > 0 && (
              <span style={{ fontSize: '12px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '12px', fontWeight: 600 }}>
                {selectedFiles.length} file(s) loaded
              </span>
            )}
          </div>

          <div style={{
            border: '2px dashed var(--line)',
            borderRadius: '10px',
            padding: '32px 20px',
            textAlign: 'center',
            background: 'var(--bg)',
            position: 'relative',
            cursor: 'pointer',
            transition: 'border-color 0.2s'
          }}>
            <input
              type="file"
              multiple
              accept=".xlsx,.xls,.xlsm,.csv"
              onChange={handleFileChange}
              style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, opacity: 0, cursor: 'pointer', width: '100%', height: '100%' }}
            />
            <FileSpreadsheet size={36} color="var(--accent)" style={{ margin: '0 auto 10px', display: 'block' }} />
            <p style={{ margin: '0 0 4px 0', fontWeight: 600, fontSize: '14px' }}>
              Click or Drag & Drop Excel Report Sheet(s)
            </p>
            <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
              Supports .xlsx, .xls, .csv files (Multiple files auto-merge)
            </span>
          </div>

          {/* Quick instructions */}
          <div style={{ fontSize: '12.5px', color: 'var(--muted)', display: 'flex', alignItems: 'flex-start', gap: '6px', background: 'var(--panel)', padding: '10px 12px', borderRadius: '6px', border: '1px solid var(--line)' }}>
            <HelpCircle size={15} style={{ flexShrink: 0, marginTop: '2px', color: 'var(--accent)' }} />
            <span>
              Expected headers: <strong>Venue Code, Venue Name, Event Name, Exam Date, Course Code, Course Name, Student Count</strong>, etc.
            </span>
          </div>
        </div>

        {/* Report Header Customizer Card */}
        <div className="card" style={{ padding: '24px', margin: 0, display: 'flex', flexDirection: 'column', gap: '16px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, fontSize: '17px', display: 'flex', alignItems: 'center', gap: '8px' }}>
              <Settings size={20} color="var(--accent)" /> 2. PDF Document Header Settings
            </h3>
            <button 
              className="secondary" 
              onClick={() => setShowConfig(!showConfig)}
              style={{ padding: '4px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}
            >
              <Sliders size={14} /> {showConfig ? 'Hide Column Mapping' : 'Column Mapping'}
            </button>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
            <div className="form-group">
              <label>University / Institution Name (Line 1)</label>
              <input 
                type="text" 
                value={universityName} 
                onChange={(e) => setUniversityName(e.target.value)} 
                placeholder="e.g. Kannur University"
              />
            </div>
            <div className="form-group">
              <label>Branch / Section (Line 2)</label>
              <input 
                type="text" 
                value={branchName} 
                onChange={(e) => setBranchName(e.target.value)} 
                placeholder="e.g. (Examination Branch)"
              />
            </div>
          </div>

          <div className="form-group">
            <label>QP Statement Event Title (Line 3)</label>
            <input 
              type="text" 
              value={examTitle} 
              onChange={(e) => setExamTitle(e.target.value)} 
              placeholder="e.g. QP Statement for 1st Semester Degree Private Registration..."
            />
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr', gap: '12px' }}>
            <div className="form-group">
              <label>Exam Session / Month-Year (Line 4)</label>
              <input 
                type="text" 
                value={sessionName} 
                onChange={(e) => setSessionName(e.target.value)} 
                placeholder="e.g. November 2025"
              />
            </div>
            <div className="form-group">
              <label>Venue Line Prefix (Line 5)</label>
              <select 
                value={centerPrefix} 
                onChange={(e) => setCenterPrefix(e.target.value)}
              >
                <option value="Venue :">Venue :</option>
                <option value="Center Name :">Center Name :</option>
                <option value="College :">College :</option>
                <option value="Venue & Center :">Venue & Center :</option>
              </select>
            </div>
          </div>

          {/* Quick formatting toggles */}
          <div style={{ display: 'flex', gap: '16px', flexWrap: 'wrap', paddingTop: '4px' }}>
            <label style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', cursor: 'pointer' }}>
              <input 
                type="checkbox" 
                checked={aggregateSameCourse} 
                onChange={(e) => setAggregateSameCourse(e.target.checked)} 
              />
              Sum NC for duplicate course on same date
            </label>

            <label style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13px', cursor: 'pointer' }}>
              <input 
                type="checkbox" 
                checked={courseCodeDot} 
                onChange={(e) => setCourseCodeDot(e.target.checked)} 
              />
              Format: CODE. - NAME.
            </label>
          </div>
        </div>
      </div>

      {/* Advanced Column Mapping Drawer / Accordion */}
      {showConfig && (
        <div className="card" style={{ padding: '20px 24px', marginBottom: '28px', background: 'var(--panel)', border: '1px solid var(--accent)' }}>
          <h4 style={{ margin: '0 0 12px 0', fontSize: '15px', color: 'var(--accent)', display: 'flex', alignItems: 'center', gap: '6px' }}>
            <Sliders size={16} /> Column Header Mappings
          </h4>
          <p style={{ fontSize: '12.5px', color: 'var(--muted)', marginTop: 0, marginBottom: '16px' }}>
            Verify or re-link which column in your uploaded Excel sheet maps to each statement field.
          </p>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '14px' }}>
            
            <div className="form-group">
              <label>Event Name Column</label>
              <select 
                value={columnMapping.eventName}
                onChange={(e) => setColumnMapping({ ...columnMapping, eventName: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Venue Code Column</label>
              <select 
                value={columnMapping.venueCode}
                onChange={(e) => setColumnMapping({ ...columnMapping, venueCode: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Venue Name Column</label>
              <select 
                value={columnMapping.venueName}
                onChange={(e) => setColumnMapping({ ...columnMapping, venueName: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Center Code Column</label>
              <select 
                value={columnMapping.centerCode}
                onChange={(e) => setColumnMapping({ ...columnMapping, centerCode: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Center Name Column</label>
              <select 
                value={columnMapping.centerName}
                onChange={(e) => setColumnMapping({ ...columnMapping, centerName: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Exam Date Column</label>
              <select 
                value={columnMapping.examDate}
                onChange={(e) => setColumnMapping({ ...columnMapping, examDate: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Course Code Column</label>
              <select 
                value={columnMapping.courseCode}
                onChange={(e) => setColumnMapping({ ...columnMapping, courseCode: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Course Name Column</label>
              <select 
                value={columnMapping.courseName}
                onChange={(e) => setColumnMapping({ ...columnMapping, courseName: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Student Count (NC) Column</label>
              <select 
                value={columnMapping.studentCount}
                onChange={(e) => setColumnMapping({ ...columnMapping, studentCount: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Group By Strategy</label>
              <select 
                value={groupByOption}
                onChange={(e) => setGroupByOption(e.target.value)}
              >
                <option value="venue">Venue Code & Venue Name (Default)</option>
                <option value="center">Center Code & Center Name</option>
              </select>
            </div>

          </div>
        </div>
      )}

      {/* Main Content Area (When Data Loaded) */}
      {Object.keys(processedGroups).length > 0 ? (
        <div>
          
          {/* Summary Statistics Cards */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px', marginBottom: '24px' }}>
            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'var(--accent-soft)', padding: '12px', borderRadius: '10px', color: 'var(--accent)' }}>
                <MapPin size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>
                  {groupByOption === 'venue' ? 'Total Venues' : 'Total Centers'}
                </div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{overallStats.totalGroups}</div>
              </div>
            </div>

            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'rgba(23, 107, 135, 0.15)', padding: '12px', borderRadius: '10px', color: 'var(--accent)' }}>
                <BookOpen size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>Total Exam Papers</div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{overallStats.totalExams}</div>
              </div>
            </div>

            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'rgba(40, 167, 69, 0.15)', padding: '12px', borderRadius: '10px', color: '#28a745' }}>
                <Users size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>Total Candidates (NC)</div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{overallStats.totalCandidates}</div>
              </div>
            </div>

            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'rgba(255, 193, 7, 0.15)', padding: '12px', borderRadius: '10px', color: '#d39e00' }}>
                <Calendar size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>Exam Days</div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{overallStats.uniqueDays} Days</div>
              </div>
            </div>
          </div>

          {/* Action Buttons Toolbar */}
          <div className="card" style={{ padding: '16px 20px', marginBottom: '24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
              <button 
                onClick={() => generateVenuePdf(effectiveGroupKey)}
                disabled={!effectiveGroupKey || isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <Download size={16} /> Download Selected Venue PDF
              </button>

              <button 
                className="secondary" 
                onClick={handlePrintPreview}
                disabled={!effectiveGroupKey || isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <Printer size={16} /> Print / PDF Preview
              </button>

              <button 
                className="secondary" 
                onClick={downloadCombinedMasterPdf}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <FileText size={16} /> All Venues (Combined PDF)
              </button>

              <button 
                className="secondary" 
                onClick={downloadAllZip}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <Archive size={16} /> All Venues (ZIP Bundle)
              </button>

              <button 
                className="secondary" 
                onClick={exportConsolidatedExcel}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <FileSpreadsheet size={16} /> Export Excel
              </button>
            </div>

            <div style={{ position: 'relative', minWidth: '240px' }}>
              <Search size={16} style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
              <input 
                type="text" 
                placeholder="Search course, date..." 
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                style={{ paddingLeft: '32px', width: '100%', fontSize: '13px' }}
              />
            </div>
          </div>

          {/* Venue Dropdown Selector Bar */}
          <div className="card" style={{ padding: '16px 20px', marginBottom: '20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '16px', flexWrap: 'wrap', background: 'var(--panel)', border: '1px solid var(--line)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flex: 1, minWidth: '320px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700, fontSize: '14px', color: 'var(--accent)', whiteSpace: 'nowrap' }}>
                <MapPin size={18} /> Select Venue:
              </div>
              <select
                value={effectiveGroupKey}
                onChange={(e) => setActiveGroupKey(e.target.value)}
                style={{
                  flex: 1,
                  padding: '10px 14px',
                  fontSize: '14px',
                  fontWeight: 600,
                  borderRadius: '8px',
                  border: '1.5px solid var(--accent)',
                  background: 'var(--bg)',
                  color: 'var(--ink)',
                  cursor: 'pointer'
                }}
              >
                {groupKeys.map((groupKey, idx) => {
                  const items = processedGroups[groupKey].items;
                  const totalNC = items.reduce((sum, it) => sum + it.studentCount, 0);
                  return (
                    <option key={groupKey} value={groupKey}>
                      {idx + 1}. {groupKey}  ({items.length} Papers • {totalNC} NC)
                    </option>
                  );
                })}
              </select>
            </div>

            {/* Quick Prev / Next Buttons and Counter */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span style={{ fontSize: '13px', color: 'var(--muted)', fontWeight: 600 }}>
                Venue {groupKeys.indexOf(effectiveGroupKey) + 1} of {groupKeys.length}
              </span>
              <button
                className="secondary"
                disabled={groupKeys.indexOf(effectiveGroupKey) <= 0}
                onClick={() => {
                  const currIdx = groupKeys.indexOf(effectiveGroupKey);
                  if (currIdx > 0) setActiveGroupKey(groupKeys[currIdx - 1]);
                }}
                style={{ padding: '7px 12px', fontSize: '12px' }}
              >
                ← Prev
              </button>
              <button
                className="secondary"
                disabled={groupKeys.indexOf(effectiveGroupKey) >= groupKeys.length - 1}
                onClick={() => {
                  const currIdx = groupKeys.indexOf(effectiveGroupKey);
                  if (currIdx < groupKeys.length - 1) setActiveGroupKey(groupKeys[currIdx + 1]);
                }}
                style={{ padding: '7px 12px', fontSize: '12px' }}
              >
                Next →
              </button>
            </div>
          </div>

          {/* Live Preview Paper Display (Matches PDF Layout 1:1) */}
          <div className="card" style={{ padding: '36px 40px', background: 'white', border: '1px solid var(--line)', borderRadius: '8px', boxShadow: '0 4px 20px rgba(0,0,0,0.06)' }}>
            
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
                <div style={{ fontSize: '14px', fontWeight: 600, color: '#222', marginBottom: '4px' }}>
                  {sessionName}
                </div>
              )}
              <div style={{ fontSize: '14.5px', fontWeight: 700, color: '#000', marginTop: '6px' }}>
                {centerPrefix} {effectiveGroupKey}
              </div>
            </div>

            {/* Rendered Statement Table */}
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12.5px', fontFamily: 'inherit' }}>
                <thead>
                  <tr style={{ background: '#f2f2f2', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000' }}>
                    <th style={{ border: '1px solid #000', padding: '8px 6px', width: '50px', textAlign: 'center', fontWeight: 700 }}>SL<br/>No</th>
                    <th style={{ border: '1px solid #000', padding: '8px 10px', width: '180px', textAlign: 'left', fontWeight: 700 }}>Date</th>
                    <th style={{ border: '1px solid #000', padding: '8px 12px', textAlign: 'left', fontWeight: 700 }}>Course</th>
                    <th style={{ border: '1px solid #000', padding: '8px 6px', width: '70px', textAlign: 'center', fontWeight: 700 }}>NC</th>
                    <th style={{ border: '1px solid #000', padding: '8px 6px', width: '75px', textAlign: 'center', fontWeight: 700 }}>QP</th>
                    <th style={{ border: '1px solid #000', padding: '8px 6px', width: '75px', textAlign: 'center', fontWeight: 700 }}>LP</th>
                  </tr>
                </thead>
                <tbody>
                  {currentTabItems.length > 0 ? (
                    currentTabItems.map((row, idx) => (
                      <tr key={idx} style={{ borderBottom: '1px solid #000', background: idx % 2 === 0 ? '#fff' : '#fafafa' }}>
                        <td style={{ border: '1px solid #000', padding: '8px 6px', textAlign: 'center', fontWeight: 600 }}>{idx + 1}</td>
                        <td style={{ border: '1px solid #000', padding: '8px 10px', verticalAlign: 'middle', whiteSpace: 'nowrap' }}>{row.formattedDate}</td>
                        <td style={{ border: '1px solid #000', padding: '8px 12px', verticalAlign: 'middle', lineHeight: '1.4' }}>
                          {formatCourseDisplay(row.courseCode, row.courseName)}
                        </td>
                        <td style={{ border: '1px solid #000', padding: '8px 6px', textAlign: 'center', fontWeight: 700 }}>{row.studentCount}</td>
                        <td style={{ border: '1px solid #000', padding: '8px 6px', textAlign: 'center' }}></td>
                        <td style={{ border: '1px solid #000', padding: '8px 6px', textAlign: 'center' }}></td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={6} style={{ textAlign: 'center', padding: '30px', color: 'var(--muted)' }}>
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
                Showing <strong>{currentTabItems.length}</strong> exam paper(s) for <strong>{effectiveGroupKey}</strong>
              </div>
              <div style={{ fontWeight: 700, color: 'var(--ink)' }}>
                Total NC Candidates: {currentTabItems.reduce((acc, curr) => acc + curr.studentCount, 0)}
              </div>
            </div>

          </div>

        </div>
      ) : (
        /* Empty State */
        <div className="card" style={{ padding: '60px 20px', textAlign: 'center', background: 'var(--panel)', border: '1px dashed var(--line)' }}>
          <BookOpen size={48} color="var(--accent)" style={{ margin: '0 auto 16px', opacity: 0.8 }} />
          <h3 style={{ margin: '0 0 8px 0', fontSize: '18px', fontWeight: 700 }}>No Report Loaded Yet</h3>
          <p style={{ color: 'var(--muted)', maxWidth: '520px', margin: '0 auto 20px', fontSize: '14px', lineHeight: '1.6' }}>
            Upload your <strong>Venue-wise / Course-wise / Date-wise</strong> examination Excel file above. The system will group reports by <strong>Venue Code</strong>, format candidate counts (NC), and generate clean university statements.
          </p>
        </div>
      )}

    </div>
  );
};

export default QpStatementPage;
