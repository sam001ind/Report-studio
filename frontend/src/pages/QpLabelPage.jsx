import { useState, useMemo } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
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
  MapPin,
  Tag,
  ChevronLeft,
  ChevronRight
} from 'lucide-react';
import { logoBase64 } from '../assets/logoBase64';

const QpLabelPage = () => {
  
  // File and Upload states
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [rawRows, setRawRows] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Ready to upload Excel files.');
  const [statusType, setStatusType] = useState('normal'); // 'normal' | 'error' | 'success'

  // Settings & Metadata
  const [examName, setExamName] = useState('Second Semester Degree (Private Registration) Regular Examinations April 2025');
  const [branchName, setBranchName] = useState('(Examination Branch)');
  const [venuePrefix, setVenuePrefix] = useState('CENTRE CODE AND NAME');
  const [courseCodeDot, setCourseCodeDot] = useState(true);
  const [showConfig, setShowConfig] = useState(false);

  // Column Mapping states
  const [columnMapping, setColumnMapping] = useState({
    eventName: -1,
    venueCode: -1,
    venueName: -1,
    centerCode: -1,
    centerName: -1,
    courseCode: -1,
    courseName: -1,
    examDate: -1,
    courseStartTime: -1,
    courseEndTime: -1,
    studentCount: -1
  });

  // UI Navigation & Filters
  const [activeVenueKey, setActiveVenueKey] = useState('');
  const [activeLabelIndexInVenue, setActiveLabelIndexInVenue] = useState(0);
  const [searchQuery, setSearchQuery] = useState('');

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

  // Auto-detect columns
  const autoDetectColumns = (headerList) => {
    const mapping = {
      eventName: -1,
      venueCode: -1,
      venueName: -1,
      centerCode: -1,
      centerName: -1,
      courseCode: -1,
      courseName: -1,
      examDate: -1,
      courseStartTime: -1,
      courseEndTime: -1,
      studentCount: -1
    };

    headerList.forEach((colName, idx) => {
      const lower = String(colName).toLowerCase().replace(/[^a-z0-9]/g, '');

      if (lower.includes('eventname') || (lower.includes('event') && !lower.includes('id'))) mapping.eventName = idx;
      if (lower === 'venuecode') mapping.venueCode = idx;
      if (lower === 'venuename') mapping.venueName = idx;
      if (lower === 'centercode' || lower === 'centrecode') mapping.centerCode = idx;
      if (lower === 'centername' || lower === 'centrename') mapping.centerName = idx;
      if (lower === 'coursecode' || lower === 'qpcoursename' || lower.includes('subjectcode')) mapping.courseCode = idx;
      if (lower === 'coursename' || lower.includes('subjectname') || lower.includes('coursetitle')) mapping.courseName = idx;
      if (lower.includes('examdate') || lower === 'date') mapping.examDate = idx;
      if (lower.includes('coursestarttime') || lower.includes('starttime')) mapping.courseStartTime = idx;
      if (lower.includes('courseendtime') || lower.includes('endtime')) mapping.courseEndTime = idx;
      if (lower.includes('studentcount') || lower === 'count' || lower === 'nc' || lower.includes('candidatecount')) mapping.studentCount = idx;
    });

    if (mapping.venueCode === -1 && mapping.centerCode !== -1) mapping.venueCode = mapping.centerCode;
    if (mapping.venueName === -1 && mapping.centerName !== -1) mapping.venueName = mapping.centerName;
    if (mapping.centerCode === -1 && mapping.venueCode !== -1) mapping.centerCode = mapping.venueCode;
    if (mapping.centerName === -1 && mapping.venueName !== -1) mapping.centerName = mapping.venueName;

    return mapping;
  };

  const handleFileChange = async (e) => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    setSelectedFiles(files);
    setIsProcessing(true);
    setStatus(`Reading ${files.length} file(s)...`, 'normal');
    setActiveLabelIndexInVenue(0);

    try {
      const allFilesRows = await Promise.all(files.map(file => readExcelFile(file)));
      let combinedData = [];
      let detectedHeaders = [];

      allFilesRows.forEach((fileRows) => {
        if (fileRows.length > 0) {
          if (detectedHeaders.length === 0) {
            detectedHeaders = fileRows[0].map((cell, idx) => cell ? String(cell).trim() : `Column ${idx + 1}`);
            combinedData = combinedData.concat(fileRows.slice(1));
          } else {
            combinedData = combinedData.concat(fileRows.slice(1));
          }
        }
      });

      if (detectedHeaders.length > 0) {
        setHeaders(detectedHeaders);
        setRawRows(combinedData);

        const newMapping = autoDetectColumns(detectedHeaders);
        setColumnMapping(newMapping);

        // Auto-detect exam title from event column or first cell
        if (newMapping.eventName !== -1 && combinedData.length > 0) {
          const sampleEvent = combinedData.find(r => r[newMapping.eventName])?.[newMapping.eventName];
          if (sampleEvent && String(sampleEvent).trim()) {
            setExamName(String(sampleEvent).trim());
          }
        } else if (combinedData.length > 0 && combinedData[0][0] && String(combinedData[0][0]).length > 15) {
          setExamName(String(combinedData[0][0]).trim());
        }

        setStatus(`Successfully loaded ${combinedData.length} records from ${files.length} file(s).`, 'success');
      } else {
        setStatus('No data rows found in uploaded files.', 'error');
      }
    } catch (err) {
      console.error(err);
      setStatus(`Error reading files: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Helper to format course display text
  const formatCourseDisplay = (courseCode, courseName) => {
    let codeStr = String(courseCode || '').trim();
    let nameStr = String(courseName || '').trim();

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

  // Group data by Venue -> Array of Labels per Venue
  const venueGroups = useMemo(() => {
    if (rawRows.length === 0) return {};

    const vMap = {};

    rawRows.forEach((row) => {
      const isBlank = row.every(cell => cell === null || cell === undefined || cell === '');
      if (isBlank) return;

      let vCode = columnMapping.venueCode !== -1 && row[columnMapping.venueCode] ? String(row[columnMapping.venueCode]).trim() : '';
      let vName = columnMapping.venueName !== -1 && row[columnMapping.venueName] ? String(row[columnMapping.venueName]).trim() : '';

      if (!vCode && !vName && columnMapping.centerCode !== -1) {
        vCode = String(row[columnMapping.centerCode] || '').trim();
        vName = String(row[columnMapping.centerName] || '').trim();
      }

      if (!vCode && !vName) {
        vCode = 'DEFAULT';
        vName = 'General Examination Venue';
      }

      const venueKey = vCode && vName ? `${vCode} - ${vName}` : (vCode || vName);

      if (!vMap[venueKey]) {
        vMap[venueKey] = {
          venueCode: vCode,
          venueName: vName,
          fullVenue: venueKey,
          labelMap: new Map()
        };
      }

      const rawCourseCode = columnMapping.courseCode !== -1 && row[columnMapping.courseCode] ? String(row[columnMapping.courseCode]).trim() : '';
      const rawCourseName = columnMapping.courseName !== -1 && row[columnMapping.courseName] ? String(row[columnMapping.courseName]).trim() : '';

      if (!rawCourseCode && !rawCourseName) return;

      // Extract date and day
      let rowDate = '';
      let rowDay = '';
      let sortDate = '9999-99-99';

      if (columnMapping.examDate !== -1) {
        const rawDate = row[columnMapping.examDate];
        let dateObj = null;

        if (rawDate instanceof Date && !isNaN(rawDate)) {
          dateObj = rawDate;
        } else if (typeof rawDate === 'number') {
          dateObj = new Date(Math.round((rawDate - 25569) * 86400 * 1000));
        } else if (rawDate) {
          const strVal = String(rawDate).trim();
          if (/^\d{4}-\d{2}-\d{2}/.test(strVal)) {
            const parts = strVal.split(/[-T\s]/);
            dateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
          } else if (/^\d{1,2}[/-]\d{1,2}[/-]\d{4}/.test(strVal)) {
            const parts = strVal.split(/[/-]/);
            dateObj = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          } else {
            const parsed = new Date(strVal);
            if (!isNaN(parsed.getTime())) dateObj = parsed;
          }
        }

        if (dateObj && !isNaN(dateObj.getTime())) {
          const yyyy = dateObj.getFullYear();
          const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
          const dd = String(dateObj.getDate()).padStart(2, '0');
          const daysOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
          rowDay = daysOfWeek[dateObj.getDay()];
          rowDate = `${yyyy}-${mm}-${dd}`;
          sortDate = rowDate;
        } else if (rawDate) {
          rowDate = String(rawDate).trim();
        }
      }

      // Combine Start Time & End Time
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

      // Candidate count
      let countVal = 1;
      if (columnMapping.studentCount !== -1 && row[columnMapping.studentCount]) {
        const parsed = parseInt(row[columnMapping.studentCount], 10);
        if (!isNaN(parsed) && parsed > 0) countVal = parsed;
      }

      // Clean normalized comparison key
      const normCode = rawCourseCode.replace(/\.+$/, '').replace(/\s+/g, ' ').trim().toLowerCase();
      const normName = rawCourseName.replace(/\.+$/, '').replace(/\s+/g, ' ').trim().toLowerCase();
      const labelKey = `${sortDate}__${normCode}__${normName}`;

      if (vMap[venueKey].labelMap.has(labelKey)) {
        const existing = vMap[venueKey].labelMap.get(labelKey);
        existing.studentCount += countVal;
      } else {
        vMap[venueKey].labelMap.set(labelKey, {
          venueCode: vCode,
          venueName: vName,
          fullVenue: venueKey,
          courseCode: rawCourseCode.replace(/\.+$/, '').trim(),
          courseName: rawCourseName.replace(/\.+$/, '').trim(),
          day: rowDay,
          date: rowDate,
          sortDate,
          time: rowTime,
          studentCount: countVal
        });
      }
    });

    // Flatten labelMap to sorted array for each venue
    const result = {};
    Object.keys(vMap).forEach(vKey => {
      const labels = Array.from(vMap[vKey].labelMap.values());
      labels.sort((a, b) => {
        if (a.sortDate !== b.sortDate) return a.sortDate.localeCompare(b.sortDate);
        return a.courseCode.localeCompare(b.courseCode);
      });

      const totalNC = labels.reduce((sum, item) => sum + item.studentCount, 0);

      result[vKey] = {
        venueCode: vMap[vKey].venueCode,
        venueName: vMap[vKey].venueName,
        fullVenue: vKey,
        labels,
        totalNC
      };
    });

    return result;
  }, [rawRows, columnMapping]);

  // Derived effective venue key
  const venueKeys = useMemo(() => Object.keys(venueGroups), [venueGroups]);
  const effectiveVenueKey = (activeVenueKey && venueGroups[activeVenueKey]) ? activeVenueKey : (venueKeys[0] || '');

  // Overall Statistics
  const overallStats = useMemo(() => {
    let totalLabels = 0;
    let totalNC = 0;
    const uniqueDates = new Set();

    venueKeys.forEach(vKey => {
      const v = venueGroups[vKey];
      totalLabels += v.labels.length;
      totalNC += v.totalNC;
      v.labels.forEach(l => {
        if (l.date) uniqueDates.add(l.date);
      });
    });

    const sortedDates = Array.from(uniqueDates).sort();
    const dateRange = sortedDates.length > 0 
      ? `${sortedDates[0]} to ${sortedDates[sortedDates.length - 1]}`
      : 'N/A';

    return {
      totalVenues: venueKeys.length,
      totalLabels,
      totalCandidates: totalNC,
      uniqueDays: sortedDates.length,
      dateRange
    };
  }, [venueGroups, venueKeys]);

  // Current Venue data & labels
  const currentVenueData = useMemo(() => venueGroups[effectiveVenueKey] || null, [venueGroups, effectiveVenueKey]);
  const currentVenueLabels = useMemo(() => currentVenueData ? currentVenueData.labels : [], [currentVenueData]);

  // Filtered labels in active venue based on search
  const displayVenueLabels = useMemo(() => {
    if (!currentVenueLabels.length) return [];
    if (!searchQuery.trim()) return currentVenueLabels;
    const q = searchQuery.toLowerCase();
    return currentVenueLabels.filter(item => 
      item.courseCode.toLowerCase().includes(q) ||
      item.courseName.toLowerCase().includes(q) ||
      item.date.toLowerCase().includes(q) ||
      item.day.toLowerCase().includes(q) ||
      String(item.studentCount).includes(q)
    );
  }, [currentVenueLabels, searchQuery]);

  // Active Label data
  const activeLabelData = displayVenueLabels[activeLabelIndexInVenue] || currentVenueLabels[0] || null;

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
    doc.text(branchName, 421, 130, { align: 'center' });

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
    doc.text(venuePrefix, startX + 8, startY + 17);
    doc.text(data.venueCode, startX + 160, startY + 17);
    doc.text(data.venueName, startX + 230, startY + 17);

    // Grid Verticals for Row 1
    doc.line(startX + 150, startY, startX + 150, startY + rowHeight);
    doc.line(startX + 220, startY, startX + 220, startY + rowHeight);

    // Grid Row 2: DAY, DATE, TIME
    const row2Y = startY + rowHeight;
    doc.rect(startX, row2Y, tableWidth, rowHeight);
    doc.text('DAY', startX + 8, row2Y + 17);
    doc.setFont('Helvetica', 'normal');
    doc.text(data.day || '', startX + 68, row2Y + 17);

    doc.setFont('Helvetica', 'bold');
    doc.text('DATE', startX + 188, row2Y + 17);
    doc.setFont('Helvetica', 'normal');
    doc.text(data.date || '', startX + 238, row2Y + 17);

    doc.setFont('Helvetica', 'bold');
    doc.text('TIME', startX + 368, row2Y + 17);
    doc.setFont('Helvetica', 'normal');
    doc.text(data.time || '', startX + 418, row2Y + 17);

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
    const subjectFormatted = formatCourseDisplay(data.courseCode, data.courseName);
    doc.text(subjectFormatted, startX + 160, row3Y + 17);

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

  // Download All Labels for the Selected Venue
  const downloadSelectedVenuePDF = () => {
    if (!currentVenueData || currentVenueLabels.length === 0) return;
    setStatus(`Generating PDF for ${effectiveVenueKey}...`, 'normal');
    setIsProcessing(true);

    try {
      const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
      currentVenueLabels.forEach((data, index) => {
        drawLabelPage(doc, data, index === 0);
      });

      const safeVenue = effectiveVenueKey.replace(/[^a-zA-Z0-9_-]/g, "_");
      doc.save(`QP_Labels_${safeVenue}_(${currentVenueLabels.length}_Labels).pdf`);
      setStatus(`Downloaded ${currentVenueLabels.length} labels for ${effectiveVenueKey}!`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to generate Venue PDF: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Download Single Current Active Label
  const downloadCurrentSingleLabelPDF = () => {
    if (!activeLabelData) return;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    drawLabelPage(doc, activeLabelData, true);

    const safeProg = activeLabelData.courseCode.replace(/[^a-zA-Z0-9_-]/g, "_");
    const safeVenue = activeLabelData.venueCode.replace(/[^a-zA-Z0-9_-]/g, "_");
    doc.save(`QP_Label_${safeVenue}_${safeProg}.pdf`);
  };

  // Download All Labels Combined into a Single Master PDF (All Venues)
  const downloadAllVenuesCombinedPDF = () => {
    const venueKeys = Object.keys(venueGroups);
    if (venueKeys.length === 0) return;
    setStatus('Generating Combined Master Landscape PDF for All Venues...', 'normal');
    setIsProcessing(true);

    try {
      const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
      let pageCount = 0;

      venueKeys.forEach(vKey => {
        venueGroups[vKey].labels.forEach(labelData => {
          drawLabelPage(doc, labelData, pageCount === 0);
          pageCount++;
        });
      });

      doc.save(`QP_Labels_All_Venues_Master_${new Date().toISOString().slice(0, 10)}.pdf`);
      setStatus(`Master PDF with ${pageCount} labels across ${venueKeys.length} venues downloaded!`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to generate Master PDF: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Batch ZIP downloads (One PDF per Venue)
  const downloadAllVenuesAsZip = async () => {
    const venueKeys = Object.keys(venueGroups);
    if (venueKeys.length === 0) return;
    setStatus('Generating ZIP package with Venue-wise QP Labels...', 'normal');
    setIsProcessing(true);
    const zip = new JSZip();

    try {
      venueKeys.forEach((vKey, idx) => {
        const vData = venueGroups[vKey];
        const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
        
        vData.labels.forEach((labelData, lIdx) => {
          drawLabelPage(doc, labelData, lIdx === 0);
        });

        const pdfBlob = doc.output('blob');
        const safeVenue = vKey.replace(/[^a-zA-Z0-9_-]/g, "_");
        zip.file(`${String(idx + 1).padStart(3, '0')}_QP_Labels_${safeVenue}.pdf`, pdfBlob);
      });

      const zipBlob = await zip.generateAsync({ type: 'blob' });
      const downloadUrl = URL.createObjectURL(zipBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `QP_Labels_All_Venues_ZIP_${new Date().toISOString().slice(0, 10)}.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(downloadUrl);

      setStatus(`ZIP archive with ${venueKeys.length} venue PDF bundles downloaded!`, 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to generate ZIP: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // Export Consolidated Excel
  const exportConsolidatedExcel = () => {
    const venueKeys = Object.keys(venueGroups);
    if (venueKeys.length === 0) return;

    try {
      const rowsForExcel = [];
      let slNo = 1;

      venueKeys.forEach(vKey => {
        venueGroups[vKey].labels.forEach(data => {
          rowsForExcel.push({
            'SL No': slNo++,
            'Venue Code': data.venueCode,
            'Venue Name': data.venueName,
            'Full Venue': data.fullVenue,
            'Day': data.day,
            'Date': data.date,
            'Time': data.time,
            'Course Code': data.courseCode,
            'Course Name': data.courseName,
            'Formatted Subject': formatCourseDisplay(data.courseCode, data.courseName),
            'Candidates (NC)': data.studentCount
          });
        });
      });

      const ws = XLSX.utils.json_to_sheet(rowsForExcel);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'QP_Labels_Data');
      XLSX.writeFile(wb, `QP_Labels_Venues_Consolidated_${new Date().toISOString().slice(0, 10)}.xlsx`);

      setStatus('Consolidated Labels Excel report downloaded successfully!', 'success');
    } catch (err) {
      console.error(err);
      setStatus(`Failed to export Excel: ${err.message}`, 'error');
    }
  };

  // Open Direct Print / Preview for the active label
  const handlePrintPreview = () => {
    if (!activeLabelData) return;
    const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
    drawLabelPage(doc, activeLabelData, true);
    const blobUrl = doc.output('bloburl');
    window.open(blobUrl, '_blank');
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
          QP Label Generator (Landscape)
        </h2>
        <p style={{ color: 'var(--muted)', margin: 0, fontSize: '14.5px' }}>
          Generate print-ready envelope covers, Question Paper packet labels, and certificate packing slips grouped by <strong>Venue Code</strong>.
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
              Expected headers: <strong>Venue Code, Venue Name, Course Code, Course Name, Exam Date, Start Time, End Time</strong>, etc.
            </span>
          </div>
        </div>

        {/* Label Header Settings Card */}
        <div className="card" style={{ padding: '24px', margin: 0, display: 'flex', flexDirection: 'column', gap: '16px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <h3 style={{ margin: 0, fontSize: '17px', display: 'flex', alignItems: 'center', gap: '8px' }}>
              <Settings size={20} color="var(--accent)" /> 2. Label Header & Template Settings
            </h3>
            <button 
              className="secondary" 
              onClick={() => setShowConfig(!showConfig)}
              style={{ padding: '4px 10px', fontSize: '12px', display: 'flex', alignItems: 'center', gap: '4px' }}
            >
              <Sliders size={14} /> {showConfig ? 'Hide Column Mapping' : 'Column Mapping'}
            </button>
          </div>

          <div className="form-group">
            <label>Examination Title / Event Subheading</label>
            <input 
              type="text" 
              value={examName} 
              onChange={(e) => setExamName(e.target.value)} 
              placeholder="e.g. Second Semester Degree (Private Registration)..."
            />
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
            <div className="form-group">
              <label>Branch Name</label>
              <input 
                type="text" 
                value={branchName} 
                onChange={(e) => setBranchName(e.target.value)} 
                placeholder="e.g. (Examination Branch)"
              />
            </div>
            <div className="form-group">
              <label>Venue Grid Header Label</label>
              <input 
                type="text" 
                value={venuePrefix} 
                onChange={(e) => setVenuePrefix(e.target.value)} 
                placeholder="e.g. CENTRE CODE AND NAME"
              />
            </div>
          </div>

          {/* Quick formatting toggles */}
          <div style={{ display: 'flex', gap: '16px', flexWrap: 'wrap', paddingTop: '4px' }}>
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
            Verify or re-link which column in your uploaded Excel sheet maps to each label field.
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
              <label>Course Start Time</label>
              <select 
                value={columnMapping.courseStartTime}
                onChange={(e) => setColumnMapping({ ...columnMapping, courseStartTime: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

            <div className="form-group">
              <label>Course End Time</label>
              <select 
                value={columnMapping.courseEndTime}
                onChange={(e) => setColumnMapping({ ...columnMapping, courseEndTime: parseInt(e.target.value, 10) })}
              >
                <option value="-1">-- Not Specified --</option>
                {headers.map((h, i) => <option key={i} value={i}>{h} (Col {i + 1})</option>)}
              </select>
            </div>

          </div>
        </div>
      )}

      {/* Main Content Section */}
      {Object.keys(venueGroups).length > 0 ? (
        <div>
          
          {/* Summary Statistics Cards */}
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: '16px', marginBottom: '24px' }}>
            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'var(--accent-soft)', padding: '12px', borderRadius: '10px', color: 'var(--accent)' }}>
                <MapPin size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>Total Venues</div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{overallStats.totalVenues}</div>
              </div>
            </div>

            <div className="card" style={{ padding: '16px 20px', margin: 0, display: 'flex', alignItems: 'center', gap: '14px' }}>
              <div style={{ background: 'rgba(23, 107, 135, 0.15)', padding: '12px', borderRadius: '10px', color: 'var(--accent)' }}>
                <Tag size={24} />
              </div>
              <div>
                <div style={{ fontSize: '12px', color: 'var(--muted)', fontWeight: 600, textTransform: 'uppercase' }}>Total QP Labels</div>
                <div style={{ fontSize: '22px', fontWeight: 700, color: 'var(--ink)' }}>{overallStats.totalLabels}</div>
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

          {/* Action Toolbar */}
          <div className="card" style={{ padding: '16px 20px', marginBottom: '24px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
              <button 
                onClick={downloadSelectedVenuePDF}
                disabled={!currentVenueData || isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <Download size={16} /> Download Selected Venue Labels ({currentVenueLabels.length})
              </button>

              <button 
                className="secondary" 
                onClick={downloadCurrentSingleLabelPDF}
                disabled={!activeLabelData || isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <FileText size={16} /> Single Label PDF
              </button>

              <button 
                className="secondary" 
                onClick={handlePrintPreview}
                disabled={!activeLabelData || isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <Printer size={16} /> Print / Preview
              </button>

              <button 
                className="secondary" 
                onClick={downloadAllVenuesCombinedPDF}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <Archive size={16} /> All Venues (Combined PDF)
              </button>

              <button 
                className="secondary" 
                onClick={downloadAllVenuesAsZip}
                disabled={isProcessing}
                style={{ display: 'inline-flex', alignItems: 'center', gap: '6px' }}
              >
                <Archive size={16} /> All Venues (ZIP Bundles)
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
                onChange={(e) => {
                  setSearchQuery(e.target.value);
                  setActiveLabelIndexInVenue(0);
                }}
                style={{ paddingLeft: '32px', width: '100%', fontSize: '13px' }}
              />
            </div>
          </div>

          {/* Venue Dropdown Selector Bar (Matches QP Statement Venue List) */}
          <div className="card" style={{ padding: '16px 20px', marginBottom: '20px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: '16px', flexWrap: 'wrap', background: 'var(--panel)', border: '1px solid var(--line)' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flex: 1, minWidth: '320px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px', fontWeight: 700, fontSize: '14px', color: 'var(--accent)', whiteSpace: 'nowrap' }}>
                <MapPin size={18} /> Select Venue:
              </div>
              <select
                value={effectiveVenueKey}
                onChange={(e) => {
                  setActiveVenueKey(e.target.value);
                  setActiveLabelIndexInVenue(0);
                }}
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
                {venueKeys.map((vKey, idx) => {
                  const v = venueGroups[vKey];
                  return (
                    <option key={vKey} value={vKey}>
                      {idx + 1}. {vKey}  ({v.labels.length} Labels • {v.totalNC} NC)
                    </option>
                  );
                })}
              </select>
            </div>

            {/* Quick Prev / Next Buttons for Venues */}
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <span style={{ fontSize: '13px', color: 'var(--muted)', fontWeight: 600 }}>
                Venue {venueKeys.indexOf(effectiveVenueKey) + 1} of {venueKeys.length}
              </span>
              <button
                className="secondary"
                disabled={venueKeys.indexOf(effectiveVenueKey) <= 0}
                onClick={() => {
                  const currIdx = venueKeys.indexOf(effectiveVenueKey);
                  if (currIdx > 0) {
                    setActiveVenueKey(venueKeys[currIdx - 1]);
                    setActiveLabelIndexInVenue(0);
                  }
                }}
                style={{ padding: '7px 12px', fontSize: '12px' }}
              >
                ← Prev Venue
              </button>
              <button
                className="secondary"
                disabled={venueKeys.indexOf(effectiveVenueKey) >= venueKeys.length - 1}
                onClick={() => {
                  const currIdx = venueKeys.indexOf(effectiveVenueKey);
                  if (currIdx < venueKeys.length - 1) {
                    setActiveVenueKey(venueKeys[currIdx + 1]);
                    setActiveLabelIndexInVenue(0);
                  }
                }}
                style={{ padding: '7px 12px', fontSize: '12px' }}
              >
                Next Venue →
              </button>
            </div>
          </div>

          {/* Visual Simulated Label Preview (Matches Landscape PDF 1:1) */}
          {activeLabelData && (
            <div className="card" style={{ padding: '32px 36px', background: 'white', border: '1px solid var(--line)', borderRadius: '8px', boxShadow: '0 4px 20px rgba(0,0,0,0.06)' }}>
              
              {/* Secondary Stepper: Label within active venue */}
              <div style={{ 
                display: 'flex', 
                justifyContent: 'space-between', 
                alignItems: 'center', 
                marginBottom: '20px', 
                paddingBottom: '14px', 
                borderBottom: '1px solid var(--line)', 
                flexWrap: 'wrap', 
                gap: '10px' 
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <span style={{ fontSize: '13.5px', fontWeight: 700, color: 'var(--ink)' }}>
                    Label {activeLabelIndexInVenue + 1} of {displayVenueLabels.length} for <em>{effectiveVenueKey}</em>
                  </span>
                  <span style={{ fontSize: '12px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '10px', fontWeight: 600 }}>
                    {activeLabelData.studentCount} Candidates (NC)
                  </span>
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                  <button
                    className="secondary"
                    disabled={activeLabelIndexInVenue <= 0}
                    onClick={() => setActiveLabelIndexInVenue(Math.max(0, activeLabelIndexInVenue - 1))}
                    style={{ padding: '6px 12px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '4px' }}
                  >
                    <ChevronLeft size={14} /> Prev Label
                  </button>
                  <button
                    className="secondary"
                    disabled={activeLabelIndexInVenue >= displayVenueLabels.length - 1}
                    onClick={() => setActiveLabelIndexInVenue(Math.min(displayVenueLabels.length - 1, activeLabelIndexInVenue + 1))}
                    style={{ padding: '6px 12px', fontSize: '12px', display: 'inline-flex', alignItems: 'center', gap: '4px' }}
                  >
                    Next Label <ChevronRight size={14} />
                  </button>
                </div>
              </div>

              {/* Landscape Paper Canvas */}
              <div style={{ 
                maxWidth: '900px', 
                margin: '0 auto', 
                background: 'white', 
                color: 'black', 
                padding: '24px 30px', 
                border: '1px solid #ddd', 
                borderRadius: '6px',
                fontFamily: 'Helvetica, Arial, sans-serif'
              }}>
                {/* University Logo Centered */}
                <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '6px' }}>
                  <img src={logoBase64} alt="Kannur University Logo" style={{ maxWidth: '320px', height: 'auto' }} />
                </div>

                <div style={{ textAlign: 'center', display: 'flex', flexDirection: 'column', gap: '4px', marginBottom: '16px' }}>
                  <strong style={{ fontSize: '13px', letterSpacing: '0.2px' }}>{branchName}</strong>
                  <strong style={{ fontSize: '12px' }}>{examName}</strong>
                </div>

                {/* Simulated Landscape Table Grid */}
                <div style={{ border: '1.5px solid black', display: 'flex', flexDirection: 'column', marginBottom: '20px' }}>
                  
                  {/* Row 1: Venue Code & Name */}
                  <div style={{ display: 'flex', borderBottom: '1.5px solid black', minHeight: '32px', alignItems: 'center' }}>
                    <div style={{ width: '220px', padding: '6px 10px', borderRight: '1.5px solid black', fontWeight: 'bold', fontSize: '12px' }}>
                      {venuePrefix}
                    </div>
                    <div style={{ width: '90px', padding: '6px 10px', borderRight: '1.5px solid black', fontWeight: 'bold', fontSize: '12px', textAlign: 'center' }}>
                      {activeLabelData.venueCode}
                    </div>
                    <div style={{ flex: 1, padding: '6px 12px', fontWeight: 'bold', fontSize: '12px' }}>
                      {activeLabelData.venueName}
                    </div>
                  </div>
                  
                  {/* Row 2: Day, Date, Time */}
                  <div style={{ display: 'flex', borderBottom: '1.5px solid black', minHeight: '32px', alignItems: 'center' }}>
                    <div style={{ width: '80px', padding: '6px 10px', borderRight: '1.5px solid black', fontWeight: 'bold', fontSize: '12px' }}>
                      DAY
                    </div>
                    <div style={{ width: '150px', padding: '6px 10px', borderRight: '1.5px solid black', fontSize: '12px' }}>
                      {activeLabelData.day || '—'}
                    </div>
                    <div style={{ width: '70px', padding: '6px 10px', borderRight: '1.5px solid black', fontWeight: 'bold', fontSize: '12px' }}>
                      DATE
                    </div>
                    <div style={{ width: '160px', padding: '6px 10px', borderRight: '1.5px solid black', fontSize: '12px' }}>
                      {activeLabelData.date || '—'}
                    </div>
                    <div style={{ width: '70px', padding: '6px 10px', borderRight: '1.5px solid black', fontWeight: 'bold', fontSize: '12px' }}>
                      TIME
                    </div>
                    <div style={{ flex: 1, padding: '6px 10px', fontSize: '12px' }}>
                      {activeLabelData.time || '—'}
                    </div>
                  </div>

                  {/* Row 3: Subject */}
                  <div style={{ display: 'flex', borderBottom: '1.5px solid black', minHeight: '32px', alignItems: 'center' }}>
                    <div style={{ width: '220px', padding: '6px 10px', borderRight: '1.5px solid black', fontWeight: 'bold', fontSize: '12px' }}>
                      SUBJECT
                    </div>
                    <div style={{ flex: 1, padding: '6px 12px', fontWeight: 'bold', fontSize: '12px' }}>
                      {formatCourseDisplay(activeLabelData.courseCode, activeLabelData.courseName)}
                    </div>
                  </div>

                  {/* Row 4: No. of Copies & Cover Number (Kept blank for manual checking) */}
                  <div style={{ display: 'flex', minHeight: '32px', alignItems: 'center' }}>
                    <div style={{ width: '220px', padding: '6px 10px', borderRight: '1.5px solid black', fontWeight: 'bold', fontSize: '12px' }}>
                      NO. OF COPIES
                    </div>
                    <div style={{ width: '90px', padding: '6px 10px', borderRight: '1.5px solid black', fontSize: '12px', textAlign: 'center' }}>
                      {/* Blank for manual checking */}
                    </div>
                    <div style={{ width: '230px', padding: '6px 10px', borderRight: '1.5px solid black', fontWeight: 'bold', fontSize: '12px' }}>
                      COVER NUMBER
                    </div>
                    <div style={{ flex: 1, padding: '6px 10px', fontSize: '12px' }}>
                      {/* Blank for manual checking */}
                    </div>
                  </div>
                </div>

                {/* Certificate Section */}
                <div style={{ borderTop: '2px dashed #888', paddingTop: '16px', marginTop: '16px' }}>
                  <h4 style={{ textAlign: 'center', margin: '0 0 10px 0', fontSize: '15px', fontWeight: 800, letterSpacing: '0.8px' }}>
                    CERTIFICATE
                  </h4>
                  <p style={{ margin: 0, lineHeight: '1.8', fontSize: '12px' }}>
                    We hereby certify that we have examined this cover and satisfied ourselves that the seals are intact and that it was opened at ______________________________________ A.M/P.M in our presence.
                  </p>
                </div>

                {/* Signatures */}
                <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '24px', fontWeight: 'bold', fontSize: '12px' }}>
                  <div>
                    <span>INVIGILATOR</span>
                    <div style={{ marginTop: '16px', fontWeight: 'normal' }}>(1) ________________________________</div>
                    <div style={{ marginTop: '10px', fontWeight: 'normal' }}>(2) ________________________________</div>
                  </div>
                  <div>
                    <span>ADDL.CHIEF SUPERINTENDENT</span>
                  </div>
                </div>

                <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '28px', fontWeight: 'bold', fontSize: '12px' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                    <span>PLACE:</span>
                    <span>DATE:</span>
                  </div>
                  <div style={{ alignSelf: 'flex-end' }}>
                    <span>CHIEF SUPERINTENDENT</span>
                  </div>
                </div>

              </div>

            </div>
          )}

        </div>
      ) : (
        /* Empty State */
        <div className="card" style={{ padding: '60px 20px', textAlign: 'center', background: 'var(--panel)', border: '1px dashed var(--line)' }}>
          <BookOpen size={48} color="var(--accent)" style={{ margin: '0 auto 16px', opacity: 0.8 }} />
          <h3 style={{ margin: '0 0 8px 0', fontSize: '18px', fontWeight: 700 }}>No Labels Compiled Yet</h3>
          <p style={{ color: 'var(--muted)', maxWidth: '520px', margin: '0 auto 20px', fontSize: '14px', lineHeight: '1.6' }}>
            Upload your examination nominal roll or course schedule Excel file above. The generator will compile envelope packing labels with dates, times, candidate copies, and certificates grouped by <strong>Venue Code</strong>.
          </p>
        </div>
      )}

    </div>
  );
};

export default QpLabelPage;
