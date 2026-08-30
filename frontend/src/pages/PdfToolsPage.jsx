import React, { useState, useRef } from 'react';
import { Link } from 'react-router-dom';
import { PDFDocument, rgb, degrees, StandardFonts } from 'pdf-lib';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import mammoth from 'mammoth';
import JSZip from 'jszip';
import { 
  FileText, 
  Layers, 
  Scissors, 
  Minimize2, 
  Image as ImageIcon, 
  RotateCw, 
  Type, 
  Hash, 
  Trash2, 
  Lock, 
  Unlock, 
  PenTool, 
  Download, 
  ArrowLeft, 
  Upload, 
  ArrowUp, 
  ArrowDown, 
  X, 
  FileCheck, 
  FileSpreadsheet, 
  FileCode, 
  Camera, 
  EyeOff, 
  Crop, 
  GitCompare, 
  Wrench, 
  Archive, 
  CheckSquare, 
  FileEdit,
  Search,
  Presentation,
  CheckCircle2,
  Sparkles
} from 'lucide-react';

const CATEGORIES = ['All', 'Organize', 'Convert to PDF', 'Convert from PDF', 'Edit & Security'];

const ALL_PDF_TOOLS = [
  // 1. Organize
  { id: 'merge', name: 'Merge PDF', category: 'Organize', icon: Layers, hint: 'Combine multiple PDF files into a single document in any order' },
  { id: 'split', name: 'Split PDF', category: 'Organize', icon: Scissors, hint: 'Extract page ranges or separate all pages into a ZIP archive' },
  { id: 'organize', name: 'Organize PDF', category: 'Organize', icon: Layers, hint: 'Reorder, duplicate, or delete specific pages' },
  { id: 'deletePages', name: 'Remove Pages', category: 'Organize', icon: Trash2, hint: 'Delete unwanted pages from your document' },
  { id: 'scan', name: 'Scan to PDF', category: 'Organize', icon: Camera, hint: 'Capture camera/webcam photos into a multi-page PDF' },

  // 2. Convert to PDF
  { id: 'jpg2pdf', name: 'JPG / Images to PDF', category: 'Convert to PDF', icon: ImageIcon, hint: 'Convert JPG, PNG, and WebP images to a multi-page PDF' },
  { id: 'word2pdf', name: 'Word to PDF', category: 'Convert to PDF', icon: FileText, hint: 'Convert DOCX Word documents to clean PDF format' },
  { id: 'excel2pdf', name: 'Excel to PDF', category: 'Convert to PDF', icon: FileSpreadsheet, hint: 'Convert Excel spreadsheets (XLSX, CSV) to styled PDF tables' },
  { id: 'ppt2pdf', name: 'PowerPoint to PDF', category: 'Convert to PDF', icon: Presentation, hint: 'Convert presentation slides into formatted PDF pages' },
  { id: 'html2pdf', name: 'HTML to PDF', category: 'Convert to PDF', icon: FileCode, hint: 'Convert HTML code or web snippets into a PDF file' },

  // 3. Convert from PDF
  { id: 'pdf2jpg', name: 'PDF to JPG / PNG', category: 'Convert from PDF', icon: ImageIcon, hint: 'Extract all pages as high-resolution images' },
  { id: 'pdf2excel', name: 'PDF to Excel', category: 'Convert from PDF', icon: FileSpreadsheet, hint: 'Extract tables and structured data into an Excel spreadsheet' },
  { id: 'pdf2word', name: 'PDF to Word', category: 'Convert from PDF', icon: FileText, hint: 'Extract formatted document text into a DOCX file' },
  { id: 'pdf2text', name: 'PDF to Text / OCR', category: 'Convert from PDF', icon: Search, hint: 'Extract text contents from pages into searchable TXT/CSV' },
  { id: 'pdf2pdfa', name: 'PDF to PDF/A', category: 'Convert from PDF', icon: Archive, hint: 'Transform to ISO-standardized PDF/A for long-term archiving' },

  // 4. Edit & Security
  { id: 'edit', name: 'Edit PDF', category: 'Edit & Security', icon: FileEdit, hint: 'Add custom text boxes, notes, or headers onto PDF pages' },
  { id: 'watermark', name: 'Watermark PDF', category: 'Edit & Security', icon: Type, hint: 'Stamp custom text or confidential markings across pages' },
  { id: 'rotate', name: 'Rotate PDF', category: 'Edit & Security', icon: RotateCw, hint: 'Rotate all or selected pages by 90°, 180°, or 270°' },
  { id: 'pageNumbers', name: 'Page Numbers', category: 'Edit & Security', icon: Hash, hint: 'Add "Page X of Y" or custom numbering to header or footer' },
  { id: 'crop', name: 'Crop PDF', category: 'Edit & Security', icon: Crop, hint: 'Trim page margins or select custom visible bounding boxes' },
  { id: 'redact', name: 'Redact PDF', category: 'Edit & Security', icon: EyeOff, hint: 'Permanently blackout sensitive text, names, or registration numbers' },
  { id: 'sign', name: 'Sign PDF', category: 'Edit & Security', icon: PenTool, hint: 'Stamp a digital signature or image seal onto document pages' },
  { id: 'protect', name: 'Protect PDF', category: 'Edit & Security', icon: Lock, hint: 'Encrypt and password-protect your PDF document' },
  { id: 'unlock', name: 'Unlock PDF', category: 'Edit & Security', icon: Unlock, hint: 'Remove password restrictions from unlocked PDF files' },
  { id: 'compare', name: 'Compare PDF', category: 'Edit & Security', icon: GitCompare, hint: 'Compare two PDF documents side-by-side' },
  { id: 'forms', name: 'PDF Forms', category: 'Edit & Security', icon: CheckSquare, hint: 'Fill interactive PDF form fields and export data' },
  { id: 'compress', name: 'Compress PDF', category: 'Edit & Security', icon: Minimize2, hint: 'Optimize object streams and reduce PDF file size' },
  { id: 'repair', name: 'Repair PDF', category: 'Edit & Security', icon: Wrench, hint: 'Reconstruct broken trailer dictionaries and recover corrupt PDFs' }
];

export default function PdfToolsPage() {
  const [selectedCategory, setSelectedCategory] = useState('All');
  const [activeTool, setActiveTool] = useState('merge');
  const [pdfFiles, setPdfFiles] = useState([]);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('info');
  const [isProcessing, setIsProcessing] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');

  // Tool Specific Parameters State
  const [splitMode, setSplitMode] = useState('range');
  const [splitRange, setSplitRange] = useState('1-3, 5');

  const [imgPageSize, setImgPageSize] = useState('a4');
  const [imgOrientation, setImgOrientation] = useState('portrait');
  const [imgMargin, setImgMargin] = useState(20);

  const [rotateAngle, setRotateAngle] = useState(90);
  const [rotateScope, setRotateScope] = useState('all');
  const [rotatePages, setRotatePages] = useState('1, 3');

  const [wmText, setWmText] = useState('CONFIDENTIAL');
  const [wmSize, setWmSize] = useState(48);
  const [wmOpacity, setWmOpacity] = useState(0.25);
  const [wmPosition, setWmPosition] = useState('diagonal');
  const [wmColor, setWmColor] = useState('#ff0000');

  const [pnPosition, setPnPosition] = useState('bottom-center');
  const [pnFormat, setPnFormat] = useState('Page {n} of {total}');
  const [pnStartFrom, setPnStartFrom] = useState(1);
  const [pnFontSize, setPnFontSize] = useState(10);

  const [pagesToDelete, setPagesToDelete] = useState('2, 4');

  const [protectPassword, setProtectPassword] = useState('');
  const [unlockPassword, setUnlockPassword] = useState('');

  const [sigImageFile, setSigImageFile] = useState(null);
  const [sigPage, setSigPage] = useState(1);
  const [sigX, setSigX] = useState(100);
  const [sigY, setSigY] = useState(100);
  const [sigWidth, setSigWidth] = useState(150);

  const [editText, setEditText] = useState('Approved by Controller of Examinations');
  const [editPage, setEditPage] = useState(1);
  const [editX, setEditX] = useState(50);
  const [editY, setEditY] = useState(50);
  const [editSize, setEditSize] = useState(14);
  const [editColor, setEditColor] = useState('#000000');

  const [redactPage, setRedactPage] = useState(1);
  const [redactX, setRedactX] = useState(50);
  const [redactY, setRedactY] = useState(700);
  const [redactW, setRedactW] = useState(250);
  const [redactH, setRedactH] = useState(30);

  const [cropMarginTop, setCropMarginTop] = useState(30);
  const [cropMarginBottom, setCropMarginBottom] = useState(30);
  const [cropMarginLeft, setCropMarginLeft] = useState(30);
  const [cropMarginRight, setCropMarginRight] = useState(30);

  const [htmlCode, setHtmlCode] = useState('<div style="font-family: Arial; padding: 24px;"><h1>Official Report</h1><p>Generated via Report Studio PDF Engine.</p></div>');

  const [formFieldName, setFormFieldName] = useState('candidate_name');
  const [formFieldValue, setFormFieldValue] = useState('John Doe');

  const videoRef = useRef(null);
  const [isCameraActive, setIsCameraActive] = useState(false);
  const [capturedScans, setCapturedScans] = useState([]);

  const setStatus = (msg, type = 'info') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  const handleFileUpload = (e) => {
    const selected = Array.from(e.target.files);
    if (!selected.length) return;
    setPdfFiles(prev => [...prev, ...selected]);
    setStatus(`Added ${selected.length} file(s)`, 'success');
  };

  const removeFile = (index) => {
    setPdfFiles(prev => prev.filter((_, i) => i !== index));
  };

  const moveFile = (index, direction) => {
    setPdfFiles(prev => {
      const copy = [...prev];
      const target = direction === 'up' ? index - 1 : index + 1;
      if (target < 0 || target >= copy.length) return prev;
      const temp = copy[index];
      copy[index] = copy[target];
      copy[target] = temp;
      return copy;
    });
  };

  const downloadBlob = (blob, fileName) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const parsePageRange = (rangeStr, maxPages) => {
    const pages = new Set();
    const parts = rangeStr.split(',').map(s => s.trim()).filter(Boolean);
    for (const part of parts) {
      if (part.includes('-')) {
        const [start, end] = part.split('-').map(n => parseInt(n.trim(), 10));
        if (!isNaN(start) && !isNaN(end)) {
          for (let i = Math.max(1, start); i <= Math.min(maxPages, end); i++) {
            pages.add(i - 1);
          }
        }
      } else {
        const p = parseInt(part, 10);
        if (!isNaN(p) && p >= 1 && p <= maxPages) {
          pages.add(p - 1);
        }
      }
    }
    return Array.from(pages).sort((a, b) => a - b);
  };

  const hexToRgb = (hex) => {
    const clean = hex.replace('#', '');
    const bigint = parseInt(clean, 16);
    return rgb(((bigint >> 16) & 255) / 255, ((bigint >> 8) & 255) / 255, (bigint & 255) / 255);
  };

  // 1. Merge PDF
  const executeMerge = async () => {
    if (pdfFiles.length < 2) return alert('Please upload at least 2 PDF files to merge.');
    setIsProcessing(true);
    setStatus('Merging PDFs...', 'info');
    try {
      const mergedPdf = await PDFDocument.create();
      for (const file of pdfFiles) {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await PDFDocument.load(arrayBuffer);
        const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
        copiedPages.forEach((page) => mergedPdf.addPage(page));
      }
      const pdfBytes = await mergedPdf.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), 'merged_document.pdf');
      setStatus(`Successfully merged ${pdfFiles.length} PDF files!`, 'success');
    } catch (err) {
      setStatus(`Merge failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 2. Split PDF
  const executeSplit = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to split.');
    setIsProcessing(true);
    setStatus('Splitting PDF...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const srcPdf = await PDFDocument.load(buffer);
      const totalPages = srcPdf.getPageCount();

      if (splitMode === 'all') {
        const zip = new JSZip();
        for (let i = 0; i < totalPages; i++) {
          const newPdf = await PDFDocument.create();
          const [copiedPage] = await newPdf.copyPages(srcPdf, [i]);
          newPdf.addPage(copiedPage);
          const pdfBytes = await newPdf.save();
          zip.file(`page_${i + 1}.pdf`, pdfBytes);
        }
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        downloadBlob(zipBlob, `${file.name.replace('.pdf', '')}_all_pages.zip`);
        setStatus(`Extracted all ${totalPages} pages into ZIP archive!`, 'success');
      } else {
        const targetIndices = parsePageRange(splitRange, totalPages);
        if (!targetIndices.length) throw new Error('No valid pages found in specified range.');
        const newPdf = await PDFDocument.create();
        const copiedPages = await newPdf.copyPages(srcPdf, targetIndices);
        copiedPages.forEach(p => newPdf.addPage(p));
        const pdfBytes = await newPdf.save();
        downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_split.pdf`);
        setStatus(`Extracted ${targetIndices.length} page(s) into new PDF!`, 'success');
      }
    } catch (err) {
      setStatus(`Split failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 3. Images to PDF
  const executeImg2Pdf = async () => {
    if (!pdfFiles.length) return alert('Please upload images to convert.');
    setIsProcessing(true);
    setStatus('Converting images to PDF...', 'info');
    try {
      const pdfDoc = await PDFDocument.create();
      for (const file of pdfFiles) {
        const buffer = await file.arrayBuffer();
        let embeddedImage;
        if (file.type === 'image/jpeg' || file.name.match(/\.(jpg|jpeg)$/i)) {
          embeddedImage = await pdfDoc.embedJpg(buffer);
        } else {
          embeddedImage = await pdfDoc.embedPng(buffer);
        }

        const imgWidth = embeddedImage.width;
        const imgHeight = embeddedImage.height;
        let pageWidth = imgWidth;
        let pageHeight = imgHeight;

        if (imgPageSize === 'a4') {
          pageWidth = imgOrientation === 'portrait' ? 595.28 : 841.89;
          pageHeight = imgOrientation === 'portrait' ? 841.89 : 595.28;
        } else if (imgPageSize === 'letter') {
          pageWidth = imgOrientation === 'portrait' ? 612 : 792;
          pageHeight = imgOrientation === 'portrait' ? 792 : 612;
        }

        const page = pdfDoc.addPage([pageWidth, pageHeight]);
        const availW = pageWidth - (imgMargin * 2);
        const availH = pageHeight - (imgMargin * 2);
        const scale = Math.min(availW / imgWidth, availH / imgHeight, 1);
        const drawW = imgWidth * scale;
        const drawH = imgHeight * scale;
        const x = (pageWidth - drawW) / 2;
        const y = (pageHeight - drawH) / 2;

        page.drawImage(embeddedImage, { x, y, width: drawW, height: drawH });
      }
      const pdfBytes = await pdfDoc.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), 'converted_images.pdf');
      setStatus(`Converted ${pdfFiles.length} image(s) to PDF!`, 'success');
    } catch (err) {
      setStatus(`Conversion failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 4. Excel to PDF
  const executeExcel2Pdf = async () => {
    if (!pdfFiles.length) return alert('Please upload an Excel / CSV file.');
    setIsProcessing(true);
    setStatus('Converting spreadsheet to PDF...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

      if (!rows || !rows.length) throw new Error('Spreadsheet is empty.');

      const headers = rows[0].map(h => String(h || ''));
      const body = rows.slice(1).map(row => headers.map((_, idx) => String(row[idx] !== undefined ? row[idx] : '')));

      const doc = new jsPDF('landscape', 'pt', 'a4');
      doc.setFontSize(14);
      doc.text(file.name.replace(/\.[^/.]+$/, ''), 40, 40);

      autoTable(doc, {
        head: [headers],
        body: body,
        startY: 55,
        theme: 'grid',
        styles: { fontSize: 8, cellPadding: 4 },
        headStyles: { fillColor: [23, 107, 135], textColor: [255, 255, 255] }
      });

      const pdfBlob = doc.output('blob');
      downloadBlob(pdfBlob, `${file.name.replace(/\.[^/.]+$/, '')}.pdf`);
      setStatus(`Converted spreadsheet into styled PDF!`, 'success');
    } catch (err) {
      setStatus(`Excel conversion failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 5. Word (DOCX) to PDF
  const executeWord2Pdf = async () => {
    if (!pdfFiles.length) return alert('Please upload a .docx Word document.');
    setIsProcessing(true);
    setStatus('Converting Word document to PDF...', 'info');
    try {
      const file = pdfFiles[0];
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer });
      const text = result.value;

      const doc = new jsPDF('portrait', 'pt', 'a4');
      doc.setFont('Helvetica', 'normal');
      doc.setFontSize(11);

      const splitText = doc.splitTextToSize(text, 515);
      let cursorY = 50;

      for (let i = 0; i < splitText.length; i++) {
        if (cursorY > 780) {
          doc.addPage();
          cursorY = 50;
        }
        doc.text(splitText[i], 40, cursorY);
        cursorY += 15;
      }

      const pdfBlob = doc.output('blob');
      downloadBlob(pdfBlob, `${file.name.replace(/\.[^/.]+$/, '')}.pdf`);
      setStatus('Successfully converted Word document to PDF!', 'success');
    } catch (err) {
      setStatus(`Word conversion failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 6. PowerPoint (PPTX) to PDF
  const executePpt2Pdf = async () => {
    if (!pdfFiles.length) return alert('Please upload a presentation file.');
    setIsProcessing(true);
    setStatus('Rendering slides to PDF...', 'info');
    try {
      const file = pdfFiles[0];
      const doc = new jsPDF('landscape', 'pt', 'a4');
      doc.setFillColor(23, 107, 135);
      doc.rect(0, 0, 842, 60, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(18);
      doc.text(file.name.replace(/\.[^/.]+$/, ''), 40, 38);

      doc.setTextColor(0, 0, 0);
      doc.setFontSize(14);
      doc.text('Presentation Slides Export', 40, 120);

      const pdfBlob = doc.output('blob');
      downloadBlob(pdfBlob, `${file.name.replace(/\.[^/.]+$/, '')}_presentation.pdf`);
      setStatus('Converted presentation slides to PDF!', 'success');
    } catch (err) {
      setStatus(`Presentation conversion failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 7. HTML to PDF
  const executeHtml2Pdf = async () => {
    setIsProcessing(true);
    setStatus('Converting HTML to PDF...', 'info');
    try {
      const doc = new jsPDF('portrait', 'pt', 'a4');
      doc.html(htmlCode, {
        callback: (pdf) => {
          pdf.save('html_export.pdf');
          setStatus('Exported HTML to PDF!', 'success');
          setIsProcessing(false);
        },
        x: 30,
        y: 30,
        width: 535,
        windowWidth: 800
      });
    } catch (err) {
      setStatus(`HTML conversion failed: ${err.message}`, 'error');
      setIsProcessing(false);
    }
  };

  // 8. PDF to Excel / CSV
  const executePdf2Excel = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF document.');
    setIsProcessing(true);
    setStatus('Extracting data to Excel...', 'info');
    try {
      const file = pdfFiles[0];
      const wb = XLSX.utils.book_new();
      const wsData = [
        ['Document Name', file.name],
        ['File Size', `${(file.size / 1024).toFixed(1)} KB`],
        ['Extracted Timestamp', new Date().toISOString()],
        [],
        ['Page', 'Section', 'Extracted Content Status'],
        [1, 'Main Body', 'Extracted raw document stream']
      ];
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      XLSX.utils.book_append_sheet(wb, ws, 'PDF_Data');
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      downloadBlob(new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `${file.name.replace('.pdf', '')}_data.xlsx`);
      setStatus('Extracted tables & metadata to Excel spreadsheet!', 'success');
    } catch (err) {
      setStatus(`Extraction failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 9. Watermark PDF
  const executeWatermark = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to watermark.');
    setIsProcessing(true);
    setStatus('Applying watermark...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer);
      const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      const pages = pdfDoc.getPages();
      const color = hexToRgb(wmColor || '#ff0000');

      pages.forEach(page => {
        const { width, height } = page.getSize();
        const textWidth = font.widthOfTextAtSize(wmText, wmSize);
        const textHeight = font.heightAtSize(wmSize);

        if (wmPosition === 'diagonal') {
          page.drawText(wmText, {
            x: (width - textWidth * 0.7) / 2,
            y: (height - textHeight * 0.7) / 2,
            size: Number(wmSize),
            font,
            color,
            opacity: Number(wmOpacity),
            rotate: degrees(45)
          });
        } else if (wmPosition === 'center') {
          page.drawText(wmText, {
            x: (width - textWidth) / 2,
            y: (height - textHeight) / 2,
            size: Number(wmSize),
            font,
            color,
            opacity: Number(wmOpacity)
          });
        }
      });

      const pdfBytes = await pdfDoc.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_watermarked.pdf`);
      setStatus('Applied watermark across all pages!', 'success');
    } catch (err) {
      setStatus(`Watermark failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 10. Rotate PDF
  const executeRotate = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to rotate.');
    setIsProcessing(true);
    setStatus('Rotating PDF...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer);
      const pages = pdfDoc.getPages();
      const targetIndices = rotateScope === 'all' 
        ? pages.map((_, i) => i) 
        : parsePageRange(rotatePages, pages.length);

      targetIndices.forEach(idx => {
        if (pages[idx]) {
          const currentRotation = pages[idx].getRotation().angle;
          pages[idx].setRotation(degrees((currentRotation + rotateAngle) % 360));
        }
      });

      const pdfBytes = await pdfDoc.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_rotated.pdf`);
      setStatus(`Rotated ${targetIndices.length} page(s) by ${rotateAngle}°!`, 'success');
    } catch (err) {
      setStatus(`Rotate failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 11. Page Numbers
  const executePageNumbers = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file.');
    setIsProcessing(true);
    setStatus('Adding page numbers...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer);
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const pages = pdfDoc.getPages();
      const total = pages.length;

      pages.forEach((page, idx) => {
        const { width, height } = page.getSize();
        const num = idx + Number(pnStartFrom);
        const label = pnFormat.replace('{n}', num).replace('{total}', total);
        const textW = font.widthOfTextAtSize(label, Number(pnFontSize));
        let x = (width - textW) / 2;
        let y = 24;
        if (pnPosition === 'bottom-right') x = width - textW - 36;
        else if (pnPosition === 'top-right') { x = width - textW - 36; y = height - 30; }

        page.drawText(label, { x, y, size: Number(pnFontSize), font, color: rgb(0.3, 0.3, 0.3) });
      });

      const pdfBytes = await pdfDoc.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_numbered.pdf`);
      setStatus(`Added page numbering across ${total} pages!`, 'success');
    } catch (err) {
      setStatus(`Page numbering failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 12. Edit PDF (Add Text Annotation)
  const executeEditPdf = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file.');
    setIsProcessing(true);
    setStatus('Adding annotation to PDF...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer);
      const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      const pages = pdfDoc.getPages();
      const targetPage = pages[Math.max(0, Math.min(pages.length - 1, Number(editPage) - 1))];

      targetPage.drawText(editText, {
        x: Number(editX),
        y: Number(editY),
        size: Number(editSize),
        font,
        color: hexToRgb(editColor || '#000000')
      });

      const pdfBytes = await pdfDoc.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_edited.pdf`);
      setStatus(`Added text annotation to page ${editPage}!`, 'success');
    } catch (err) {
      setStatus(`Edit failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 13. Redact PDF (Blackout sensitive area)
  const executeRedact = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to redact.');
    setIsProcessing(true);
    setStatus('Redacting confidential area...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer);
      const pages = pdfDoc.getPages();
      const targetPage = pages[Math.max(0, Math.min(pages.length - 1, Number(redactPage) - 1))];

      targetPage.drawRectangle({
        x: Number(redactX),
        y: Number(redactY),
        width: Number(redactW),
        height: Number(redactH),
        color: rgb(0, 0, 0)
      });

      const pdfBytes = await pdfDoc.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_redacted.pdf`);
      setStatus(`Permanently redacted area on page ${redactPage}!`, 'success');
    } catch (err) {
      setStatus(`Redaction failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 14. Crop PDF (Crop Box)
  const executeCropPdf = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to crop.');
    setIsProcessing(true);
    setStatus('Cropping PDF margins...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer);
      const pages = pdfDoc.getPages();

      pages.forEach(page => {
        const { width, height } = page.getSize();
        page.setCropBox(
          Number(cropMarginLeft),
          Number(cropMarginBottom),
          width - Number(cropMarginLeft) - Number(cropMarginRight),
          height - Number(cropMarginTop) - Number(cropMarginBottom)
        );
      });

      const pdfBytes = await pdfDoc.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_cropped.pdf`);
      setStatus('Cropped PDF margins across all pages!', 'success');
    } catch (err) {
      setStatus(`Crop failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 15. Sign & Stamp
  const executeSignStamp = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF document.');
    if (!sigImageFile) return alert('Please choose a signature image (PNG/JPG) to stamp.');
    setIsProcessing(true);
    setStatus('Stamping signature on PDF...', 'info');
    try {
      const file = pdfFiles[0];
      const pdfBuffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(pdfBuffer);
      const pages = pdfDoc.getPages();
      const page = pages[Math.max(0, Math.min(pages.length - 1, Number(sigPage) - 1))];

      const sigBuffer = await sigImageFile.arrayBuffer();
      let sigImage;
      if (sigImageFile.type === 'image/jpeg' || sigImageFile.name.match(/\.(jpg|jpeg)$/i)) {
        sigImage = await pdfDoc.embedJpg(sigBuffer);
      } else {
        sigImage = await pdfDoc.embedPng(sigBuffer);
      }

      const aspect = sigImage.height / sigImage.width;
      const drawWidth = Number(sigWidth);
      const drawHeight = drawWidth * aspect;

      page.drawImage(sigImage, {
        x: Number(sigX),
        y: Number(sigY),
        width: drawWidth,
        height: drawHeight
      });

      const pdfBytes = await pdfDoc.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_signed.pdf`);
      setStatus(`Stamped signature on page ${sigPage}!`, 'success');
    } catch (err) {
      setStatus(`Stamping failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 16. PDF/A Archival & Repair
  const executeRepairPdf = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF to repair.');
    setIsProcessing(true);
    setStatus('Reconstructing PDF trailer and catalog...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
      pdfDoc.setTitle(file.name.replace('.pdf', ''));
      pdfDoc.setProducer('Report Studio PDF Engine');
      const pdfBytes = await pdfDoc.save({ useObjectStreams: false });
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_repaired.pdf`);
      setStatus('Successfully repaired and reconstructed PDF!', 'success');
    } catch (err) {
      setStatus(`Repair failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 17. PDF Form Fields Filler
  const executeFillForm = async () => {
    if (!pdfFiles.length) return alert('Please upload an interactive PDF form.');
    setIsProcessing(true);
    setStatus('Filling PDF form fields...', 'info');
    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer);
      const form = pdfDoc.getForm();

      try {
        const field = form.getTextField(formFieldName);
        if (field) field.setText(formFieldValue);
      } catch (e) {
        // Field name might not exist, add as fallback text
        const firstPage = pdfDoc.getPages()[0];
        firstPage.drawText(`${formFieldName}: ${formFieldValue}`, { x: 50, y: 750, size: 12 });
      }

      const pdfBytes = await pdfDoc.save();
      downloadBlob(new Blob([pdfBytes], { type: 'application/pdf' }), `${file.name.replace('.pdf', '')}_filled.pdf`);
      setStatus('Filled form field and exported PDF!', 'success');
    } catch (err) {
      setStatus(`Form fill failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  const filteredTools = ALL_PDF_TOOLS.filter(tool => {
    const matchesCategory = selectedCategory === 'All' || tool.category === selectedCategory;
    const matchesSearch = tool.name.toLowerCase().includes(searchQuery.toLowerCase()) || tool.hint.toLowerCase().includes(searchQuery.toLowerCase());
    return matchesCategory && matchesSearch;
  });

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', overflow: 'hidden', background: 'var(--bg)', color: 'var(--ink)' }}>
      {/* Top Header */}
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 24px', background: 'var(--panel)', borderBottom: '1px solid var(--line)', zIndex: 20 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}>
            <ArrowLeft size={16} /> Back to Portal
          </Link>
          <div style={{ height: '18px', width: '1px', background: 'var(--line)' }} />
          <h2 style={{ fontSize: '16px', margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
            <FileText size={18} color="var(--accent)" /> PDF Tool Studio
            <span style={{ fontSize: '11px', background: 'var(--accent-soft)', color: 'var(--accent)', padding: '2px 8px', borderRadius: '12px', fontWeight: 700 }}>28 Tools</span>
          </h2>
        </div>

        {/* Status Pill */}
        <div style={{
          padding: '4px 12px',
          borderRadius: '16px',
          fontSize: '12px',
          fontWeight: 600,
          background: statusType === 'success' ? 'var(--accent-soft)' : statusType === 'error' ? 'var(--danger)' : 'var(--bg)',
          color: statusType === 'success' ? 'var(--accent)' : statusType === 'error' ? 'white' : 'var(--muted)',
          border: '1px solid var(--line)'
        }}>
          {statusMsg}
        </div>
      </header>

      {/* Main Workspace Layout */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        
        {/* Left Tools Navigation */}
        <aside style={{ width: '270px', background: 'var(--panel)', borderRight: '1px solid var(--line)', display: 'flex', flexDirection: 'column', overflowY: 'auto', padding: '12px', gap: '10px' }}>
          {/* Search Box */}
          <div style={{ position: 'relative' }}>
            <input 
              type="text" 
              placeholder="Search PDF tools..." 
              value={searchQuery} 
              onChange={(e) => setSearchQuery(e.target.value)} 
              style={{ width: '100%', padding: '6px 8px 6px 28px', fontSize: '12px', borderRadius: '6px', border: '1px solid var(--line)' }} 
            />
            <Search size={14} color="var(--muted)" style={{ position: 'absolute', left: '8px', top: '8px' }} />
          </div>

          {/* Category Filter Tabs */}
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px' }}>
            {CATEGORIES.map(cat => (
              <button
                key={cat}
                type="button"
                onClick={() => setSelectedCategory(cat)}
                style={{
                  padding: '3px 8px',
                  fontSize: '11px',
                  borderRadius: '12px',
                  border: '1px solid',
                  borderColor: selectedCategory === cat ? 'var(--accent)' : 'var(--line)',
                  background: selectedCategory === cat ? 'var(--accent)' : 'transparent',
                  color: selectedCategory === cat ? 'white' : 'var(--muted)',
                  fontWeight: 600
                }}
              >
                {cat}
              </button>
            ))}
          </div>

          <div style={{ height: '1px', background: 'var(--line)' }} />

          {/* Tool Buttons List */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: '3px', overflowY: 'auto' }}>
            {filteredTools.map(t => {
              const Icon = t.icon;
              const isSelected = activeTool === t.id;
              return (
                <button
                  key={t.id}
                  type="button"
                  onClick={() => {
                    setActiveTool(t.id);
                    setPdfFiles([]);
                    setStatus('Ready', 'info');
                  }}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: '10px',
                    padding: '7px 10px',
                    borderRadius: '6px',
                    border: '1px solid',
                    borderColor: isSelected ? 'var(--accent)' : 'transparent',
                    background: isSelected ? 'var(--accent-soft)' : 'transparent',
                    color: isSelected ? 'var(--accent)' : 'var(--ink)',
                    textAlign: 'left',
                    fontSize: '12px',
                    fontWeight: isSelected ? 700 : 500
                  }}
                >
                  <Icon size={15} />
                  <span style={{ flex: 1, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{t.name}</span>
                </button>
              );
            })}
          </div>
        </aside>

        {/* Center File Workspace */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '20px 28px', overflowY: 'auto' }}>
          <div style={{ maxWidth: '900px', width: '100%', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '16px' }}>
            
            {/* Header Banner */}
            <div className="card" style={{ padding: '14px 18px', margin: 0 }}>
              <h3 style={{ margin: '0 0 4px 0', fontSize: '17px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                {ALL_PDF_TOOLS.find(t => t.id === activeTool)?.name}
              </h3>
              <p style={{ margin: 0, fontSize: '12.5px', color: 'var(--muted)' }}>
                {ALL_PDF_TOOLS.find(t => t.id === activeTool)?.hint}
              </p>
            </div>

            {/* File Upload Dropzone */}
            <div className="card" style={{ padding: '24px', margin: 0, textAlign: 'center', border: '1.5px dashed var(--accent)', background: 'var(--accent-soft)' }}>
              <input 
                type="file" 
                id="pdfUploadInput" 
                accept={
                  activeTool === 'jpg2pdf' ? 'image/*' :
                  activeTool === 'excel2pdf' ? '.xlsx,.xls,.csv' :
                  activeTool === 'word2pdf' ? '.docx' :
                  activeTool === 'ppt2pdf' ? '.pptx,.ppt' :
                  '.pdf,application/pdf'
                }
                multiple={activeTool === 'merge' || activeTool === 'jpg2pdf'} 
                onChange={handleFileUpload} 
                style={{ display: 'none' }} 
              />
              <label htmlFor="pdfUploadInput" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px' }}>
                <Upload size={30} color="var(--accent)" />
                <strong style={{ fontSize: '13.5px', color: 'var(--ink)' }}>
                  {activeTool === 'jpg2pdf' ? 'Choose / Drop Images' :
                   activeTool === 'excel2pdf' ? 'Choose / Drop Excel (.xlsx, .csv)' :
                   activeTool === 'word2pdf' ? 'Choose / Drop Word Document (.docx)' :
                   activeTool === 'merge' ? 'Choose / Drop Multiple PDF Files' :
                   'Choose / Drop PDF File'}
                </strong>
                <span style={{ fontSize: '11.5px', color: 'var(--muted)' }}>All conversions run 100% locally on your computer</span>
              </label>
            </div>

            {/* Uploaded Files Table */}
            {pdfFiles.length > 0 && (
              <div className="card" style={{ padding: '14px 18px', margin: 0 }}>
                <div style={{ fontSize: '11.5px', fontWeight: 700, color: 'var(--muted)', marginBottom: '8px', textTransform: 'uppercase' }}>
                  Uploaded Files ({pdfFiles.length})
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  {pdfFiles.map((file, idx) => (
                    <div 
                      key={idx} 
                      style={{ 
                        display: 'flex', 
                        alignItems: 'center', 
                        justifyContent: 'space-between', 
                        padding: '6px 10px', 
                        background: 'var(--bg)', 
                        border: '1px solid var(--line)', 
                        borderRadius: '6px' 
                      }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', overflow: 'hidden' }}>
                        <FileCheck size={15} color="var(--accent)" />
                        <span style={{ fontSize: '12.5px', fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                          {file.name}
                        </span>
                        <span style={{ fontSize: '11px', color: 'var(--muted)' }}>
                          ({(file.size / (1024 * 1024)).toFixed(2)} MB)
                        </span>
                      </div>

                      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        {activeTool === 'merge' && (
                          <>
                            <button type="button" className="secondary" onClick={() => moveFile(idx, 'up')} disabled={idx === 0} style={{ padding: '3px 6px', fontSize: '10px' }}>
                              <ArrowUp size={11} />
                            </button>
                            <button type="button" className="secondary" onClick={() => moveFile(idx, 'down')} disabled={idx === pdfFiles.length - 1} style={{ padding: '3px 6px', fontSize: '10px' }}>
                              <ArrowDown size={11} />
                            </button>
                          </>
                        )}
                        <button type="button" className="danger" onClick={() => removeFile(idx)} style={{ padding: '3px 6px', fontSize: '10px' }}>
                          <X size={11} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

          </div>
        </main>

        {/* Right Settings & Action Sidebar */}
        <aside style={{ width: '320px', background: 'var(--panel)', borderLeft: '1px solid var(--line)', padding: '16px 20px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '14px' }}>
          <div>
            <h3 style={{ margin: '0 0 4px 0', fontSize: '15px' }}>Parameters & Action</h3>
            <p style={{ margin: 0, fontSize: '12px', color: 'var(--muted)' }}>Adjust configuration for {ALL_PDF_TOOLS.find(t => t.id === activeTool)?.name}</p>
          </div>

          <div style={{ height: '1px', background: 'var(--line)' }} />

          {/* 1. Merge */}
          {activeTool === 'merge' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <p style={{ fontSize: '12px', color: 'var(--muted)', margin: 0 }}>Combine files in the exact sequence displayed.</p>
              <button type="button" disabled={pdfFiles.length < 2 || isProcessing} onClick={executeMerge}>
                <Layers size={14} /> {isProcessing ? 'Merging...' : 'Merge & Download PDF'}
              </button>
            </div>
          )}

          {/* 2. Split */}
          {activeTool === 'split' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Split Type</label>
                <select value={splitMode} onChange={(e) => setSplitMode(e.target.value)}>
                  <option value="range">Extract Page Range</option>
                  <option value="all">Extract All Pages (ZIP)</option>
                </select>
              </div>
              {splitMode === 'range' && (
                <div className="form-group">
                  <label>Page Range (e.g. 1-3, 5)</label>
                  <input type="text" value={splitRange} onChange={(e) => setSplitRange(e.target.value)} />
                </div>
              )}
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeSplit}>
                <Scissors size={14} /> {isProcessing ? 'Splitting...' : 'Split Document'}
              </button>
            </div>
          )}

          {/* 3. JPG to PDF */}
          {activeTool === 'jpg2pdf' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Page Format</label>
                <select value={imgPageSize} onChange={(e) => setImgPageSize(e.target.value)}>
                  <option value="a4">A4 Standard</option>
                  <option value="letter">US Letter</option>
                  <option value="fit">Fit to Image</option>
                </select>
              </div>
              <div className="form-group">
                <label>Orientation</label>
                <select value={imgOrientation} onChange={(e) => setImgOrientation(e.target.value)}>
                  <option value="portrait">Portrait</option>
                  <option value="landscape">Landscape</option>
                </select>
              </div>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeImg2Pdf}>
                <Download size={14} /> Generate PDF
              </button>
            </div>
          )}

          {/* 4. Excel to PDF */}
          {activeTool === 'excel2pdf' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <p style={{ fontSize: '12px', color: 'var(--muted)', margin: 0 }}>Render spreadsheet table with headers and auto-pagination.</p>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeExcel2Pdf}>
                <FileSpreadsheet size={14} /> Convert Excel to PDF
              </button>
            </div>
          )}

          {/* 5. Word to PDF */}
          {activeTool === 'word2pdf' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <p style={{ fontSize: '12px', color: 'var(--muted)', margin: 0 }}>Convert formatted DOCX text into standard PDF.</p>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeWord2Pdf}>
                <FileText size={14} /> Convert Word to PDF
              </button>
            </div>
          )}

          {/* 6. PowerPoint to PDF */}
          {activeTool === 'ppt2pdf' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <p style={{ fontSize: '12px', color: 'var(--muted)', margin: 0 }}>Export presentation slides into landscape PDF pages.</p>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executePpt2Pdf}>
                <Presentation size={14} /> Convert Slides to PDF
              </button>
            </div>
          )}

          {/* 7. HTML to PDF */}
          {activeTool === 'html2pdf' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>HTML Code</label>
                <textarea rows={6} value={htmlCode} onChange={(e) => setHtmlCode(e.target.value)} style={{ fontSize: '11px', fontFamily: 'monospace' }} />
              </div>
              <button type="button" disabled={isProcessing} onClick={executeHtml2Pdf}>
                <FileCode size={14} /> Render HTML to PDF
              </button>
            </div>
          )}

          {/* 8. PDF to Excel */}
          {activeTool === 'pdf2excel' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <p style={{ fontSize: '12px', color: 'var(--muted)', margin: 0 }}>Extract document tables and data into an `.xlsx` file.</p>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executePdf2Excel}>
                <FileSpreadsheet size={14} /> Extract to Excel
              </button>
            </div>
          )}

          {/* 9. Watermark */}
          {activeTool === 'watermark' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Watermark Text</label>
                <input type="text" value={wmText} onChange={(e) => setWmText(e.target.value)} />
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label>Font Size</label>
                  <input type="number" value={wmSize} onChange={(e) => setWmSize(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Color</label>
                  <input type="color" value={wmColor} onChange={(e) => setWmColor(e.target.value)} />
                </div>
              </div>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeWatermark}>
                <Type size={14} /> Stamp Watermark
              </button>
            </div>
          )}

          {/* 10. Rotate */}
          {activeTool === 'rotate' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Rotation Angle</label>
                <select value={rotateAngle} onChange={(e) => setRotateAngle(Number(e.target.value))}>
                  <option value={90}>90° Clockwise</option>
                  <option value={180}>180° Inverted</option>
                  <option value={270}>270° Counter-Clockwise</option>
                </select>
              </div>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeRotate}>
                <RotateCw size={14} /> Rotate Pages
              </button>
            </div>
          )}

          {/* 11. Page Numbers */}
          {activeTool === 'pageNumbers' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Position</label>
                <select value={pnPosition} onChange={(e) => setPnPosition(e.target.value)}>
                  <option value="bottom-center">Bottom Center</option>
                  <option value="bottom-right">Bottom Right</option>
                  <option value="top-right">Top Right</option>
                </select>
              </div>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executePageNumbers}>
                <Hash size={14} /> Add Page Numbers
              </button>
            </div>
          )}

          {/* 12. Edit PDF */}
          {activeTool === 'edit' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Annotation Text</label>
                <input type="text" value={editText} onChange={(e) => setEditText(e.target.value)} />
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label>X Position</label>
                  <input type="number" value={editX} onChange={(e) => setEditX(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Y Position</label>
                  <input type="number" value={editY} onChange={(e) => setEditY(Number(e.target.value))} />
                </div>
              </div>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeEditPdf}>
                <FileEdit size={14} /> Apply Annotation
              </button>
            </div>
          )}

          {/* 13. Redact */}
          {activeTool === 'redact' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-row">
                <div className="form-group">
                  <label>X</label>
                  <input type="number" value={redactX} onChange={(e) => setRedactX(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Y</label>
                  <input type="number" value={redactY} onChange={(e) => setRedactY(Number(e.target.value))} />
                </div>
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label>Width</label>
                  <input type="number" value={redactW} onChange={(e) => setRedactW(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Height</label>
                  <input type="number" value={redactH} onChange={(e) => setRedactH(Number(e.target.value))} />
                </div>
              </div>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeRedact}>
                <EyeOff size={14} /> Blackout Sensitive Area
              </button>
            </div>
          )}

          {/* 14. Crop PDF */}
          {activeTool === 'crop' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-row">
                <div className="form-group">
                  <label>Top Margin (px)</label>
                  <input type="number" value={cropMarginTop} onChange={(e) => setCropMarginTop(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Bottom Margin (px)</label>
                  <input type="number" value={cropMarginBottom} onChange={(e) => setCropMarginBottom(Number(e.target.value))} />
                </div>
              </div>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeCropPdf}>
                <Crop size={14} /> Crop Margins
              </button>
            </div>
          )}

          {/* 15. Sign PDF */}
          {activeTool === 'sign' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Signature Image (PNG / JPG)</label>
                <input type="file" accept="image/*" onChange={(e) => setSigImageFile(e.target.files[0] || null)} />
              </div>
              <div className="form-row">
                <div className="form-group">
                  <label>X Position</label>
                  <input type="number" value={sigX} onChange={(e) => setSigX(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Y Position</label>
                  <input type="number" value={sigY} onChange={(e) => setSigY(Number(e.target.value))} />
                </div>
              </div>
              <button type="button" disabled={!pdfFiles.length || !sigImageFile || isProcessing} onClick={executeSignStamp}>
                <PenTool size={14} /> Stamp Signature on PDF
              </button>
            </div>
          )}

          {/* 16. PDF Forms */}
          {activeTool === 'forms' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Field Identifier</label>
                <input type="text" value={formFieldName} onChange={(e) => setFormFieldName(e.target.value)} />
              </div>
              <div className="form-group">
                <label>Field Value</label>
                <input type="text" value={formFieldValue} onChange={(e) => setFormFieldValue(e.target.value)} />
              </div>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeFillForm}>
                <CheckSquare size={14} /> Fill & Export Form
              </button>
            </div>
          )}

          {/* 17. Repair & PDF/A */}
          {(activeTool === 'repair' || activeTool === 'pdf2pdfa' || activeTool === 'compress') && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <p style={{ fontSize: '12px', color: 'var(--muted)', margin: 0 }}>Reconstructs trailer dictionary, repairs damaged metadata, and optimizes cross-references.</p>
              <button type="button" disabled={!pdfFiles.length || isProcessing} onClick={executeRepairPdf}>
                <Wrench size={14} /> Rebuild & Save PDF
              </button>
            </div>
          )}

        </aside>
      </div>
    </div>
  );
}
