import React, { useState } from 'react';
import { Link } from 'react-router-dom';
import { PDFDocument, rgb, degrees, StandardFonts } from 'pdf-lib';
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
  PenTool, 
  Download, 
  ArrowLeft,
  Upload,
  ArrowUp,
  ArrowDown,
  X,
  FileCheck
} from 'lucide-react';

const PDF_TOOLS = [
  { id: 'merge', name: 'Merge PDF', icon: Layers, hint: 'Combine multiple PDF files into a single document' },
  { id: 'split', name: 'Split & Extract', icon: Scissors, hint: 'Extract page ranges or split every page into a ZIP archive' },
  { id: 'img2pdf', name: 'Images to PDF', icon: ImageIcon, hint: 'Convert JPG, PNG, and WebP images into a multi-page PDF' },
  { id: 'rotate', name: 'Rotate PDF', icon: RotateCw, hint: 'Rotate all or selected PDF pages by 90°, 180°, or 270°' },
  { id: 'watermark', name: 'Watermark PDF', icon: Type, hint: 'Stamp custom text or confidential markings across pages' },
  { id: 'pageNumbers', name: 'Page Numbers', icon: Hash, hint: 'Add "Page X of Y" or custom numbering to header or footer' },
  { id: 'deletePages', name: 'Delete Pages', icon: Trash2, hint: 'Remove unwanted pages or isolate specific sheets' },
  { id: 'protect', name: 'Protect PDF', icon: Lock, hint: 'Encrypt and password-protect sensitive documents' },
  { id: 'sign', name: 'Sign & Stamp', icon: PenTool, hint: 'Stamp a signature image or text seal directly onto PDF pages' },
  { id: 'compress', name: 'Optimize & Clean', icon: Minimize2, hint: 'Clean object streams and rewrite optimized PDF' }
];

export default function PdfToolsPage() {
  const [activeTool, setActiveTool] = useState('merge');
  const [pdfFiles, setPdfFiles] = useState([]);
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('info');
  const [isProcessing, setIsProcessing] = useState(false);

  // Tool Specific States
  const [splitMode, setSplitMode] = useState('range'); // 'range' | 'all'
  const [splitRange, setSplitRange] = useState('1-3, 5');

  const [imgPageSize, setImgPageSize] = useState('a4'); // 'a4' | 'letter' | 'fit'
  const [imgOrientation, setImgOrientation] = useState('portrait'); // 'portrait' | 'landscape'
  const [imgMargin, setImgMargin] = useState(20);

  const [rotateAngle, setRotateAngle] = useState(90);
  const [rotateScope, setRotateScope] = useState('all'); // 'all' | 'custom'
  const [rotatePages, setRotatePages] = useState('1, 3');

  const [wmText, setWmText] = useState('CONFIDENTIAL');
  const [wmSize, setWmSize] = useState(48);
  const [wmOpacity, setWmOpacity] = useState(0.25);
  const [wmPosition, setWmPosition] = useState('diagonal'); // 'diagonal' | 'center' | 'header' | 'footer'
  const [wmColor, setWmColor] = useState('#ff0000');

  const [pnPosition, setPnPosition] = useState('bottom-center'); // 'bottom-center' | 'bottom-right' | 'top-right'
  const [pnFormat, setPnFormat] = useState('Page {n} of {total}'); // 'Page {n} of {total}' | '{n}' | 'Page {n}'
  const [pnStartFrom, setPnStartFrom] = useState(1);
  const [pnFontSize, setPnFontSize] = useState(10);

  const [pagesToDelete, setPagesToDelete] = useState('2, 4');

  const [protectPassword, setProtectPassword] = useState('');

  const [sigImageFile, setSigImageFile] = useState(null);
  const [sigPage, setSigPage] = useState(1);
  const [sigX, setSigX] = useState(100);
  const [sigY, setSigY] = useState(100);
  const [sigWidth, setSigWidth] = useState(150);

  const setStatus = (msg, type = 'info') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  const handleFileUpload = (e) => {
    const selected = Array.from(e.target.files);
    if (!selected.length) return;
    if (activeTool === 'img2pdf') {
      const validImages = selected.filter(f => f.type.startsWith('image/'));
      setPdfFiles(prev => [...prev, ...validImages]);
      setStatus(`Added ${validImages.length} image(s)`, 'success');
    } else {
      const validPdfs = selected.filter(f => f.type === 'application/pdf' || f.name.endsWith('.pdf'));
      setPdfFiles(prev => [...prev, ...validPdfs]);
      setStatus(`Added ${validPdfs.length} PDF file(s)`, 'success');
    }
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
            pages.add(i - 1); // 0-indexed
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

  // 1. Action: Merge PDFs
  const executeMerge = async () => {
    if (pdfFiles.length < 2) {
      return alert('Please upload at least 2 PDF files to merge.');
    }
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

      const mergedPdfBytes = await mergedPdf.save();
      const blob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
      downloadBlob(blob, 'merged_document.pdf');
      setStatus(`Successfully merged ${pdfFiles.length} PDFs into 1 file!`, 'success');
    } catch (err) {
      console.error('Merge PDF error:', err);
      setStatus(`Merge failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 2. Action: Split PDF
  const executeSplit = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to split.');
    setIsProcessing(true);
    setStatus('Splitting PDF...', 'info');

    try {
      const file = pdfFiles[0];
      const arrayBuffer = await file.arrayBuffer();
      const srcPdf = await PDFDocument.load(arrayBuffer);
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
        downloadBlob(zipBlob, `${file.name.replace('.pdf', '')}_pages.zip`);
        setStatus(`Extracted all ${totalPages} pages into ZIP!`, 'success');
      } else {
        const targetIndices = parsePageRange(splitRange, totalPages);
        if (!targetIndices.length) {
          throw new Error('No valid pages found in the specified range.');
        }
        const newPdf = await PDFDocument.create();
        const copiedPages = await newPdf.copyPages(srcPdf, targetIndices);
        copiedPages.forEach(p => newPdf.addPage(p));
        const pdfBytes = await newPdf.save();
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        downloadBlob(blob, `${file.name.replace('.pdf', '')}_split.pdf`);
        setStatus(`Extracted ${targetIndices.length} page(s) into new PDF!`, 'success');
      }
    } catch (err) {
      console.error('Split error:', err);
      setStatus(`Split failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 3. Action: Images to PDF
  const executeImg2Pdf = async () => {
    if (!pdfFiles.length) return alert('Please upload at least one image file.');
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
        const availWidth = pageWidth - (imgMargin * 2);
        const availHeight = pageHeight - (imgMargin * 2);

        const scale = Math.min(availWidth / imgWidth, availHeight / imgHeight, 1);
        const drawWidth = imgWidth * scale;
        const drawHeight = imgHeight * scale;

        const x = (pageWidth - drawWidth) / 2;
        const y = (pageHeight - drawHeight) / 2;

        page.drawImage(embeddedImage, {
          x,
          y,
          width: drawWidth,
          height: drawHeight
        });
      }

      const pdfBytes = await pdfDoc.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      downloadBlob(blob, 'converted_images.pdf');
      setStatus(`Successfully converted ${pdfFiles.length} image(s) to PDF!`, 'success');
    } catch (err) {
      console.error('Img2Pdf error:', err);
      setStatus(`Conversion failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 4. Action: Rotate PDF
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
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      downloadBlob(blob, `${file.name.replace('.pdf', '')}_rotated.pdf`);
      setStatus(`Rotated ${targetIndices.length} page(s) by ${rotateAngle}°!`, 'success');
    } catch (err) {
      console.error('Rotate error:', err);
      setStatus(`Rotate failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 5. Action: Watermark PDF
  const executeWatermark = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to watermark.');
    setIsProcessing(true);
    setStatus('Stamping watermark...', 'info');

    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer);
      const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);
      const pages = pdfDoc.getPages();

      const hexToRgb = (hex) => {
        const clean = hex.replace('#', '');
        const bigint = parseInt(clean, 16);
        return rgb(((bigint >> 16) & 255) / 255, ((bigint >> 8) & 255) / 255, (bigint & 255) / 255);
      };
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
        } else if (wmPosition === 'header') {
          page.drawText(wmText, {
            x: (width - textWidth) / 2,
            y: height - textHeight - 20,
            size: Number(wmSize),
            font,
            color,
            opacity: Number(wmOpacity)
          });
        } else {
          page.drawText(wmText, {
            x: (width - textWidth) / 2,
            y: 20,
            size: Number(wmSize),
            font,
            color,
            opacity: Number(wmOpacity)
          });
        }
      });

      const pdfBytes = await pdfDoc.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      downloadBlob(blob, `${file.name.replace('.pdf', '')}_watermarked.pdf`);
      setStatus('Stamped watermark across all pages!', 'success');
    } catch (err) {
      console.error('Watermark error:', err);
      setStatus(`Watermark failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 6. Action: Add Page Numbers
  const executePageNumbers = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to add page numbers.');
    setIsProcessing(true);
    setStatus('Adding page numbers...', 'info');

    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer);
      const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const pages = pdfDoc.getPages();
      const totalPages = pages.length;

      pages.forEach((page, idx) => {
        const { width, height } = page.getSize();
        const pageNum = idx + Number(pnStartFrom);
        const label = pnFormat
          .replace('{n}', pageNum)
          .replace('{total}', totalPages);

        const textWidth = font.widthOfTextAtSize(label, Number(pnFontSize));
        let x = (width - textWidth) / 2;
        let y = 24;

        if (pnPosition === 'bottom-right') {
          x = width - textWidth - 36;
          y = 24;
        } else if (pnPosition === 'top-right') {
          x = width - textWidth - 36;
          y = height - 30;
        }

        page.drawText(label, {
          x,
          y,
          size: Number(pnFontSize),
          font,
          color: rgb(0.3, 0.3, 0.3)
        });
      });

      const pdfBytes = await pdfDoc.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      downloadBlob(blob, `${file.name.replace('.pdf', '')}_numbered.pdf`);
      setStatus(`Added page numbers across ${totalPages} pages!`, 'success');
    } catch (err) {
      console.error('Page numbers error:', err);
      setStatus(`Page numbering failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 7. Action: Delete Pages
  const executeDeletePages = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to delete pages.');
    setIsProcessing(true);
    setStatus('Deleting pages...', 'info');

    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const srcPdf = await PDFDocument.load(buffer);
      const totalPages = srcPdf.getPageCount();
      const toDelete = parsePageRange(pagesToDelete, totalPages);
      const toKeep = [];

      for (let i = 0; i < totalPages; i++) {
        if (!toDelete.includes(i)) toKeep.push(i);
      }

      if (!toKeep.length) throw new Error('Cannot delete all pages.');

      const newPdf = await PDFDocument.create();
      const copiedPages = await newPdf.copyPages(srcPdf, toKeep);
      copiedPages.forEach(p => newPdf.addPage(p));

      const pdfBytes = await newPdf.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      downloadBlob(blob, `${file.name.replace('.pdf', '')}_cleaned.pdf`);
      setStatus(`Removed ${toDelete.length} page(s), kept ${toKeep.length} page(s)!`, 'success');
    } catch (err) {
      console.error('Delete pages error:', err);
      setStatus(`Delete pages failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 8. Action: Sign & Stamp
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

      const targetPageIndex = Math.max(0, Math.min(pages.length - 1, Number(sigPage) - 1));
      const page = pages[targetPageIndex];

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
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      downloadBlob(blob, `${file.name.replace('.pdf', '')}_signed.pdf`);
      setStatus(`Stamped signature on page ${targetPageIndex + 1}!`, 'success');
    } catch (err) {
      console.error('Sign stamp error:', err);
      setStatus(`Stamping failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  // 9. Action: Clean / Optimize
  const executeOptimize = async () => {
    if (!pdfFiles.length) return alert('Please upload a PDF file to optimize.');
    setIsProcessing(true);
    setStatus('Cleaning and optimizing PDF streams...', 'info');

    try {
      const file = pdfFiles[0];
      const buffer = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.load(buffer, { ignoreEncryption: true });
      const pdfBytes = await pdfDoc.save({ useObjectStreams: true });
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      downloadBlob(blob, `${file.name.replace('.pdf', '')}_optimized.pdf`);
      setStatus('Successfully optimized PDF document!', 'success');
    } catch (err) {
      console.error('Optimize error:', err);
      setStatus(`Optimization failed: ${err.message}`, 'error');
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100vw', overflow: 'hidden', background: 'var(--bg)', color: 'var(--ink)' }}>
      {/* Top Bar */}
      <header style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '10px 24px', background: 'var(--panel)', borderBottom: '1px solid var(--line)', zIndex: 20 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
          <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px' }}>
            <ArrowLeft size={16} /> Back to Portal
          </Link>
          <div style={{ height: '18px', width: '1px', background: 'var(--line)' }} />
          <h2 style={{ fontSize: '16px', margin: 0, display: 'flex', alignItems: 'center', gap: '8px' }}>
            <FileText size={18} color="var(--accent)" /> PDF Tool Studio
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
        {/* Left Tools Sidebar */}
        <aside style={{ width: '250px', background: 'var(--panel)', borderRight: '1px solid var(--line)', display: 'flex', flexDirection: 'column', overflowY: 'auto', padding: '12px' }}>
          <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', marginBottom: '8px', textTransform: 'uppercase' }}>PDF Tools</div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
            {PDF_TOOLS.map((t) => {
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
                    padding: '8px 12px',
                    borderRadius: '6px',
                    border: '1px solid',
                    borderColor: isSelected ? 'var(--accent)' : 'transparent',
                    background: isSelected ? 'var(--accent-soft)' : 'transparent',
                    color: isSelected ? 'var(--accent)' : 'var(--ink)',
                    textAlign: 'left',
                    fontSize: '12.5px',
                    fontWeight: isSelected ? 700 : 500
                  }}
                >
                  <Icon size={16} />
                  <span style={{ flex: 1 }}>{t.name}</span>
                </button>
              );
            })}
          </div>
        </aside>

        {/* Center File Area & Preview */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', padding: '24px 32px', overflowY: 'auto' }}>
          <div style={{ maxWidth: '900px', width: '100%', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '20px' }}>
            
            {/* Tool Header Card */}
            <div className="card" style={{ padding: '16px 20px', margin: 0 }}>
              <h3 style={{ margin: '0 0 4px 0', fontSize: '18px', display: 'flex', alignItems: 'center', gap: '8px' }}>
                {PDF_TOOLS.find(t => t.id === activeTool)?.name}
              </h3>
              <p style={{ margin: 0, fontSize: '13px', color: 'var(--muted)' }}>
                {PDF_TOOLS.find(t => t.id === activeTool)?.hint}
              </p>
            </div>

            {/* File Upload Box */}
            <div className="card" style={{ padding: '24px', margin: 0, textAlign: 'center', border: '1.5px dashed var(--accent)', background: 'var(--accent-soft)' }}>
              <input 
                type="file" 
                id="pdfFileInput" 
                accept={activeTool === 'img2pdf' ? 'image/*' : '.pdf,application/pdf'} 
                multiple={activeTool === 'merge' || activeTool === 'img2pdf'} 
                onChange={handleFileUpload} 
                style={{ display: 'none' }} 
              />
              <label htmlFor="pdfFileInput" style={{ cursor: 'pointer', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px' }}>
                <Upload size={32} color="var(--accent)" />
                <strong style={{ fontSize: '14px', color: 'var(--ink)' }}>
                  {activeTool === 'img2pdf' ? 'Select or Drop Images' : activeTool === 'merge' ? 'Select or Drop Multiple PDF Files' : 'Select or Drop a PDF File'}
                </strong>
                <span style={{ fontSize: '12px', color: 'var(--muted)' }}>
                  {activeTool === 'img2pdf' ? 'JPG, PNG, WebP supported' : 'PDF documents supported'}
                </span>
              </label>
            </div>

            {/* Uploaded Files List */}
            {pdfFiles.length > 0 && (
              <div className="card" style={{ padding: '16px 20px', margin: 0 }}>
                <div style={{ fontSize: '12px', fontWeight: 700, color: 'var(--muted)', marginBottom: '10px', textTransform: 'uppercase' }}>
                  Uploaded Files ({pdfFiles.length})
                </div>
                <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                  {pdfFiles.map((file, idx) => (
                    <div 
                      key={idx} 
                      style={{ 
                        display: 'flex', 
                        alignItems: 'center', 
                        justifyContent: 'space-between', 
                        padding: '8px 12px', 
                        background: 'var(--bg)', 
                        border: '1px solid var(--line)', 
                        borderRadius: '6px' 
                      }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', gap: '10px', overflow: 'hidden' }}>
                        <FileCheck size={16} color="var(--accent)" />
                        <span style={{ fontSize: '13px', fontWeight: 600, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                          {file.name}
                        </span>
                        <span style={{ fontSize: '11px', color: 'var(--muted)' }}>
                          ({(file.size / (1024 * 1024)).toFixed(2)} MB)
                        </span>
                      </div>

                      <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        {activeTool === 'merge' && (
                          <>
                            <button 
                              type="button" 
                              className="secondary" 
                              onClick={() => moveFile(idx, 'up')} 
                              disabled={idx === 0}
                              style={{ padding: '4px 6px', fontSize: '11px' }}
                            >
                              <ArrowUp size={12} />
                            </button>
                            <button 
                              type="button" 
                              className="secondary" 
                              onClick={() => moveFile(idx, 'down')} 
                              disabled={idx === pdfFiles.length - 1}
                              style={{ padding: '4px 6px', fontSize: '11px' }}
                            >
                              <ArrowDown size={12} />
                            </button>
                          </>
                        )}
                        <button 
                          type="button" 
                          className="danger" 
                          onClick={() => removeFile(idx)}
                          style={{ padding: '4px 8px', fontSize: '11px' }}
                        >
                          <X size={12} />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

          </div>
        </main>

        {/* Right Settings & Action Panel */}
        <aside style={{ width: '320px', background: 'var(--panel)', borderLeft: '1px solid var(--line)', padding: '16px 20px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '16px' }}>
          <div>
            <h3 style={{ margin: '0 0 4px 0', fontSize: '15px' }}>Options & Execution</h3>
            <p style={{ margin: 0, fontSize: '12px', color: 'var(--muted)' }}>Configure parameters before processing</p>
          </div>

          <div style={{ height: '1px', background: 'var(--line)' }} />

          {/* 1. Merge PDF Options */}
          {activeTool === 'merge' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <p style={{ fontSize: '12.5px', color: 'var(--muted)', margin: 0 }}>
                Files will be merged in the exact order shown in the list. Use the up/down arrows to rearrange files.
              </p>
              <button 
                type="button" 
                disabled={pdfFiles.length < 2 || isProcessing} 
                onClick={executeMerge}
                style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '10px' }}
              >
                <Layers size={16} /> {isProcessing ? 'Merging...' : 'Merge Files & Download'}
              </button>
            </div>
          )}

          {/* 2. Split PDF Options */}
          {activeTool === 'split' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Split Mode</label>
                <select value={splitMode} onChange={(e) => setSplitMode(e.target.value)}>
                  <option value="range">Custom Page Range</option>
                  <option value="all">Extract All Pages (ZIP)</option>
                </select>
              </div>

              {splitMode === 'range' && (
                <div className="form-group">
                  <label>Page Ranges (e.g. 1-3, 5, 8-10)</label>
                  <input type="text" value={splitRange} onChange={(e) => setSplitRange(e.target.value)} />
                </div>
              )}

              <button 
                type="button" 
                disabled={!pdfFiles.length || isProcessing} 
                onClick={executeSplit}
                style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '10px' }}
              >
                <Scissors size={16} /> {isProcessing ? 'Splitting...' : 'Split & Download PDF'}
              </button>
            </div>
          )}

          {/* 3. Images to PDF Options */}
          {activeTool === 'img2pdf' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Page Format</label>
                <select value={imgPageSize} onChange={(e) => setImgPageSize(e.target.value)}>
                  <option value="a4">A4 Standard</option>
                  <option value="letter">US Letter</option>
                  <option value="fit">Fit to Image Size</option>
                </select>
              </div>

              {imgPageSize !== 'fit' && (
                <div className="form-group">
                  <label>Orientation</label>
                  <select value={imgOrientation} onChange={(e) => setImgOrientation(e.target.value)}>
                    <option value="portrait">Portrait</option>
                    <option value="landscape">Landscape</option>
                  </select>
                </div>
              )}

              <div className="form-group">
                <label>Page Margin ({imgMargin}px)</label>
                <input type="range" min="0" max="60" value={imgMargin} onChange={(e) => setImgMargin(Number(e.target.value))} />
              </div>

              <button 
                type="button" 
                disabled={!pdfFiles.length || isProcessing} 
                onClick={executeImg2Pdf}
                style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '10px' }}
              >
                <Download size={16} /> {isProcessing ? 'Converting...' : 'Generate Multi-Page PDF'}
              </button>
            </div>
          )}

          {/* 4. Rotate PDF Options */}
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

              <div className="form-group">
                <label>Target Pages</label>
                <select value={rotateScope} onChange={(e) => setRotateScope(e.target.value)}>
                  <option value="all">All Pages</option>
                  <option value="custom">Specific Pages</option>
                </select>
              </div>

              {rotateScope === 'custom' && (
                <div className="form-group">
                  <label>Page Numbers (e.g. 1, 3, 5-7)</label>
                  <input type="text" value={rotatePages} onChange={(e) => setRotatePages(e.target.value)} />
                </div>
              )}

              <button 
                type="button" 
                disabled={!pdfFiles.length || isProcessing} 
                onClick={executeRotate}
                style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '10px' }}
              >
                <RotateCw size={16} /> {isProcessing ? 'Rotating...' : 'Rotate & Save PDF'}
              </button>
            </div>
          )}

          {/* 5. Watermark PDF Options */}
          {activeTool === 'watermark' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Watermark Text</label>
                <input type="text" value={wmText} onChange={(e) => setWmText(e.target.value)} />
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Font Size ({wmSize}px)</label>
                  <input type="number" min="16" max="100" value={wmSize} onChange={(e) => setWmSize(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Color</label>
                  <input type="color" value={wmColor} onChange={(e) => setWmColor(e.target.value)} />
                </div>
              </div>

              <div className="form-group">
                <label>Opacity: {Math.round(wmOpacity * 100)}%</label>
                <input type="range" min="0.05" max="1" step="0.05" value={wmOpacity} onChange={(e) => setWmOpacity(Number(e.target.value))} />
              </div>

              <div className="form-group">
                <label>Position</label>
                <select value={wmPosition} onChange={(e) => setWmPosition(e.target.value)}>
                  <option value="diagonal">Diagonal 45° (Center)</option>
                  <option value="center">Horizontal Center</option>
                  <option value="header">Top Header</option>
                  <option value="footer">Bottom Footer</option>
                </select>
              </div>

              <button 
                type="button" 
                disabled={!pdfFiles.length || isProcessing} 
                onClick={executeWatermark}
                style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '10px' }}
              >
                <Type size={16} /> {isProcessing ? 'Stamping...' : 'Stamp Watermark & Save'}
              </button>
            </div>
          )}

          {/* 6. Page Numbers Options */}
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

              <div className="form-group">
                <label>Numbering Format</label>
                <select value={pnFormat} onChange={(e) => setPnFormat(e.target.value)}>
                  <option value="Page {n} of {total}">Page X of Y</option>
                  <option value="Page {n}">Page X</option>
                  <option value="{n}">X (Only Number)</option>
                </select>
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Start Number</label>
                  <input type="number" min="1" value={pnStartFrom} onChange={(e) => setPnStartFrom(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Font Size</label>
                  <input type="number" min="8" max="20" value={pnFontSize} onChange={(e) => setPnFontSize(Number(e.target.value))} />
                </div>
              </div>

              <button 
                type="button" 
                disabled={!pdfFiles.length || isProcessing} 
                onClick={executePageNumbers}
                style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '10px' }}
              >
                <Hash size={16} /> {isProcessing ? 'Adding...' : 'Apply Page Numbers'}
              </button>
            </div>
          )}

          {/* 7. Delete Pages Options */}
          {activeTool === 'deletePages' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Pages to Delete (e.g. 2, 4, 7-9)</label>
                <input type="text" value={pagesToDelete} onChange={(e) => setPagesToDelete(e.target.value)} />
              </div>

              <button 
                type="button" 
                disabled={!pdfFiles.length || isProcessing} 
                onClick={executeDeletePages}
                style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '10px' }}
              >
                <Trash2 size={16} /> {isProcessing ? 'Removing...' : 'Delete Pages & Save'}
              </button>
            </div>
          )}

          {/* 8. Sign & Stamp Options */}
          {activeTool === 'sign' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Signature Image (PNG / JPG)</label>
                <input 
                  type="file" 
                  accept="image/*" 
                  onChange={(e) => setSigImageFile(e.target.files[0] || null)} 
                />
              </div>

              <div className="form-group">
                <label>Target Page Number</label>
                <input type="number" min="1" value={sigPage} onChange={(e) => setSigPage(Number(e.target.value))} />
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>X Position (px)</label>
                  <input type="number" min="0" value={sigX} onChange={(e) => setSigX(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Y Position (px)</label>
                  <input type="number" min="0" value={sigY} onChange={(e) => setSigY(Number(e.target.value))} />
                </div>
              </div>

              <div className="form-group">
                <label>Stamp Width ({sigWidth}px)</label>
                <input type="number" min="40" max="400" value={sigWidth} onChange={(e) => setSigWidth(Number(e.target.value))} />
              </div>

              <button 
                type="button" 
                disabled={!pdfFiles.length || !sigImageFile || isProcessing} 
                onClick={executeSignStamp}
                style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '10px' }}
              >
                <PenTool size={16} /> {isProcessing ? 'Stamping...' : 'Stamp Signature on PDF'}
              </button>
            </div>
          )}

          {/* 9. Optimize & Clean Options */}
          {activeTool === 'compress' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <p style={{ fontSize: '12.5px', color: 'var(--muted)', margin: 0 }}>
                Rewrites PDF object streams, strips orphaned metadata, and consolidates PDF cross-reference tables.
              </p>
              <button 
                type="button" 
                disabled={!pdfFiles.length || isProcessing} 
                onClick={executeOptimize}
                style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '10px' }}
              >
                <Minimize2 size={16} /> {isProcessing ? 'Optimizing...' : 'Clean & Save PDF'}
              </button>
            </div>
          )}
        </aside>
      </div>
    </div>
  );
}
