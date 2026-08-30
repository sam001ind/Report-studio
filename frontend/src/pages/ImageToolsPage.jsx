import React, { useState, useRef, useEffect } from 'react';
import { Link } from 'react-router-dom';
import { 
  Sparkles, 
  Image as ImageIcon, 
  Crop, 
  Maximize2, 
  Minimize2, 
  RotateCw, 
  Sliders, 
  Download, 
  Undo2, 
  RotateCcw, 
  Type, 
  EyeOff, 
  Code2, 
  Palette,
  FileImage,
  Layers,
  ArrowLeft,
  CheckCircle2,
  Trash2
} from 'lucide-react';

const TOOLS = [
  { id: 'removebg', name: 'Signature / Remove BG', icon: Sparkles, hint: 'Clean scanned signatures or transparentize backgrounds' },
  { id: 'compress', name: 'Compress Image', icon: Minimize2, hint: 'Reduce file size with fine-tuned quality control' },
  { id: 'resize', name: 'Resize', icon: Maximize2, hint: 'Resize by exact pixel dimensions or percentage' },
  { id: 'crop', name: 'Crop', icon: Crop, hint: 'Crop canvas to custom bounding box' },
  { id: 'convert', name: 'Convert Format', icon: FileImage, hint: 'Convert between PNG, JPG, and WebP' },
  { id: 'editor', name: 'Photo Editor', icon: Sliders, hint: 'Adjust brightness, contrast, saturation, and filters' },
  { id: 'upscale', name: 'Upscale', icon: Layers, hint: 'Super-sample 2x, 3x, or 4x with high-quality scaling' },
  { id: 'watermark', name: 'Watermark', icon: Type, hint: 'Stamp custom text or copyright markings' },
  { id: 'meme', name: 'Meme Generator', icon: Palette, hint: 'Add classic top & bottom caption text' },
  { id: 'rotate', name: 'Rotate & Flip', icon: RotateCw, hint: 'Rotate 90°/180°/270° or flip horizontally/vertically' },
  { id: 'blur', name: 'Privacy Blur', icon: EyeOff, hint: 'Redact sensitive numbers, signatures, or photos' },
  { id: 'html', name: 'HTML to Image', icon: Code2, hint: 'Render pasted HTML code directly into an image' }
];

export default function ImageToolsPage() {
  const [activeTool, setActiveTool] = useState('removebg');
  const [files, setFiles] = useState([]);
  const [activeFileIndex, setActiveFileIndex] = useState(0);
  const [hasImage, setHasImage] = useState(false);
  const [dimensions, setDimensions] = useState({ width: 0, height: 0 });
  const [statusMsg, setStatusMsg] = useState('Ready');
  const [statusType, setStatusType] = useState('info');
  const [historyLength, setHistoryLength] = useState(0);

  // Tool Specific Parameters State
  const [bgMode, setBgMode] = useState('signature'); // 'signature' | 'corner' | 'green'
  const [bgThreshold, setBgThreshold] = useState(52);
  const [bgSoftness, setBgSoftness] = useState(24);
  const [bgDarken, setBgDarken] = useState(true);

  const [compressFormat, setCompressFormat] = useState('image/jpeg');
  const [compressQuality, setCompressQuality] = useState(0.85);

  const [resizeWidth, setResizeWidth] = useState(800);
  const [resizeHeight, setResizeHeight] = useState(600);
  const [resizePercent, setResizePercent] = useState(100);
  const [resizeLock, setResizeLock] = useState(true);

  const [cropX, setCropX] = useState(0);
  const [cropY, setCropY] = useState(0);
  const [cropW, setCropW] = useState(400);
  const [cropH, setCropH] = useState(300);

  const [convertFormat, setConvertFormat] = useState('image/png');
  const [convertBg, setConvertBg] = useState('#ffffff');
  const [convertQuality, setConvertQuality] = useState(0.92);

  const [brightness, setBrightness] = useState(1.0);
  const [contrast, setContrast] = useState(1.0);
  const [saturate, setSaturate] = useState(1.0);
  const [softBlur, setSoftBlur] = useState(0);
  const [grayscale, setGrayscale] = useState(false);
  const [sepia, setSepia] = useState(false);

  const [upscaleFactor, setUpscaleFactor] = useState(2);

  const [wmText, setWmText] = useState('CONFIDENTIAL');
  const [wmColor, setWmColor] = useState('#ff0000');
  const [wmSize, setWmSize] = useState(36);
  const [wmOpacity, setWmOpacity] = useState(0.5);
  const [wmPosition, setWmPosition] = useState('center');

  const [memeTop, setMemeTop] = useState('TOP TEXT');
  const [memeBottom, setMemeBottom] = useState('BOTTOM TEXT');
  const [memeSize, setMemeSize] = useState(48);

  const [blurX, setBlurX] = useState(50);
  const [blurY, setBlurY] = useState(50);
  const [blurW, setBlurW] = useState(200);
  const [blurH, setBlurH] = useState(100);
  const [blurAmount, setBlurAmount] = useState(16);

  const [htmlCode, setHtmlCode] = useState('<div style="width:600px;height:350px;background:#176b87;color:#fff;display:flex;align-items:center;justify-content:center;font-family:sans-serif;border-radius:12px;"><h2>Report Studio Image</h2></div>');
  const [htmlW, setHtmlW] = useState(600);
  const [htmlH, setHtmlH] = useState(350);

  const canvasRef = useRef(null);
  const historyRef = useRef([]);
  const originalRef = useRef(null);

  const setStatus = (msg, type = 'info') => {
    setStatusMsg(msg);
    setStatusType(type);
  };

  const remember = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    const snap = ctx.getImageData(0, 0, canvas.width, canvas.height);
    historyRef.current.push({
      width: canvas.width,
      height: canvas.height,
      imageData: snap
    });
    setHistoryLength(historyRef.current.length);
  };

  const handleUndo = () => {
    if (!historyRef.current.length) return;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    const last = historyRef.current.pop();
    setHistoryLength(historyRef.current.length);
    canvas.width = last.width;
    canvas.height = last.height;
    ctx.putImageData(last.imageData, 0, 0);
    setDimensions({ width: canvas.width, height: canvas.height });
    setStatus('Undid last modification', 'info');
  };

  const handleReset = () => {
    if (!originalRef.current) return;
    remember();
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    canvas.width = originalRef.current.width;
    canvas.height = originalRef.current.height;
    ctx.putImageData(originalRef.current.imageData, 0, 0);
    setDimensions({ width: canvas.width, height: canvas.height });
    setStatus('Reset canvas to original image', 'info');
  };

  const loadFileToCanvas = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = canvasRef.current;
        if (!canvas) return;
        const ctx = canvas.getContext('2d', { willReadFrequently: true });
        canvas.width = img.width;
        canvas.height = img.height;
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0);

        const snap = ctx.getImageData(0, 0, img.width, img.height);
        originalRef.current = { width: img.width, height: img.height, imageData: snap };
        historyRef.current = [];
        setHistoryLength(0);

        setDimensions({ width: img.width, height: img.height });
        setResizeWidth(img.width);
        setResizeHeight(img.height);
        setCropW(Math.min(img.width, 400));
        setCropH(Math.min(img.height, 300));
        setHasImage(true);
        setStatus(`Loaded "${file.name}" (${img.width} × ${img.height})`, 'success');
      };
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  };

  const handleFileUpload = (e) => {
    const selected = Array.from(e.target.files).filter(f => f.type.startsWith('image/') || f.name.endsWith('.svg'));
    if (!selected.length) return;
    setFiles(selected);
    setActiveFileIndex(0);
    loadFileToCanvas(selected[0]);
  };

  const selectFileFromQueue = (index) => {
    if (!files[index]) return;
    setActiveFileIndex(index);
    loadFileToCanvas(files[index]);
  };

  // Helper Canvas
  const makeTempCanvas = (w, h) => {
    const c = document.createElement('canvas');
    c.width = w;
    c.height = h;
    return { canvas: c, ctx: c.getContext('2d', { willReadFrequently: true }) };
  };

  // Action: Remove Background / Signature Extractor
  const executeRemoveBg = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const d = imgData.data;
    const sample = bgMode === 'green' ? [40, 180, 80] : [d[0], d[1], d[2]];

    for (let i = 0; i < d.length; i += 4) {
      const r = d[i], g = d[i + 1], b = d[i + 2];
      const dist = bgMode === 'signature'
        ? Math.max(0, 255 - Math.min(r, g, b)) + Math.abs(r - g) * 0.5 + Math.abs(g - b) * 0.5
        : Math.hypot(r - sample[0], g - sample[1], b - sample[2]);

      let alpha = 0;
      if (bgSoftness <= 0) {
        alpha = dist > bgThreshold ? 255 : 0;
      } else {
        alpha = Math.max(0, Math.min(255, ((dist - bgThreshold) / bgSoftness) * 255));
      }
      d[i + 3] = alpha;

      if (bgDarken && bgMode === 'signature' && d[i + 3] > 0) {
        const ink = Math.min(r, g, b) * 0.35;
        d[i] = ink;
        d[i + 1] = ink;
        d[i + 2] = ink;
      }
    }
    ctx.putImageData(imgData, 0, 0);
    setStatus('Extracted signature / removed background', 'success');
  };

  // Action: Resize
  const executeResize = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    let targetW = resizeWidth;
    let targetH = resizeHeight;
    if (resizePercent !== 100) {
      targetW = Math.max(1, Math.round((canvas.width * resizePercent) / 100));
      targetH = Math.max(1, Math.round((canvas.height * resizePercent) / 100));
    }
    const temp = makeTempCanvas(canvas.width, canvas.height);
    temp.ctx.drawImage(canvas, 0, 0);
    canvas.width = targetW;
    canvas.height = targetH;
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(temp.canvas, 0, 0, targetW, targetH);
    setDimensions({ width: targetW, height: targetH });
    setStatus(`Resized to ${targetW} × ${targetH}px`, 'success');
  };

  // Action: Crop
  const executeCrop = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    const x = Math.max(0, Math.min(canvas.width - 1, Number(cropX)));
    const y = Math.max(0, Math.min(canvas.height - 1, Number(cropY)));
    const w = Math.max(1, Math.min(canvas.width - x, Number(cropW)));
    const h = Math.max(1, Math.min(canvas.height - y, Number(cropH)));
    const cropped = ctx.getImageData(x, y, w, h);
    canvas.width = w;
    canvas.height = h;
    ctx.putImageData(cropped, 0, 0);
    setDimensions({ width: w, height: h });
    setStatus(`Cropped to ${w} × ${h}px`, 'success');
  };

  // Action: Photo Filters
  const executeFilters = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    const ctx = canvas.getContext('2d');
    const filter = [
      `brightness(${brightness})`,
      `contrast(${contrast})`,
      `saturate(${saturate})`,
      grayscale ? 'grayscale(1)' : '',
      sepia ? 'sepia(1)' : '',
      softBlur > 0 ? `blur(${softBlur}px)` : ''
    ].filter(Boolean).join(' ');

    const temp = makeTempCanvas(canvas.width, canvas.height);
    temp.ctx.drawImage(canvas, 0, 0);
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.filter = filter || 'none';
    ctx.drawImage(temp.canvas, 0, 0);
    ctx.filter = 'none';
    setStatus('Applied photo adjustments & filters', 'success');
  };

  // Action: Upscale
  const executeUpscale = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    const scale = Number(upscaleFactor || 2);
    const targetW = canvas.width * scale;
    const targetH = canvas.height * scale;
    const temp = makeTempCanvas(canvas.width, canvas.height);
    temp.ctx.drawImage(canvas, 0, 0);
    canvas.width = targetW;
    canvas.height = targetH;
    const ctx = canvas.getContext('2d');
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(temp.canvas, 0, 0, targetW, targetH);
    setDimensions({ width: targetW, height: targetH });
    setStatus(`Upscaled ${scale}x to ${targetW} × ${targetH}px`, 'success');
  };

  // Action: Watermark
  const executeWatermark = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    const ctx = canvas.getContext('2d');
    ctx.save();
    ctx.globalAlpha = Number(wmOpacity);
    ctx.fillStyle = wmColor;
    ctx.strokeStyle = 'rgba(0,0,0,0.35)';
    ctx.lineWidth = Math.max(2, Number(wmSize) * 0.06);
    ctx.font = `700 ${wmSize}px Inter, sans-serif`;
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';

    const draw = (x, y) => {
      ctx.strokeText(wmText, x, y);
      ctx.fillText(wmText, x, y);
    };

    if (wmPosition === 'tile') {
      ctx.rotate(-Math.PI / 8);
      for (let y = -canvas.height; y < canvas.height * 2; y += wmSize * 3) {
        for (let x = -canvas.width; x < canvas.width * 2; x += wmSize * 6) draw(x, y);
      }
    } else {
      const pad = wmSize * 1.5;
      const positions = {
        center: [canvas.width / 2, canvas.height / 2],
        'top-left': [pad, pad],
        'top-right': [canvas.width - pad, pad],
        'bottom-left': [pad, canvas.height - pad],
        'bottom-right': [canvas.width - pad, canvas.height - pad]
      };
      const [px, py] = positions[wmPosition] || positions.center;
      draw(px, py);
    }
    ctx.restore();
    setStatus('Applied watermark', 'success');
  };

  // Action: Meme
  const executeMeme = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    const ctx = canvas.getContext('2d');
    ctx.save();
    ctx.font = `900 ${memeSize}px Impact, Arial Black, sans-serif`;
    ctx.fillStyle = '#ffffff';
    ctx.strokeStyle = '#000000';
    ctx.lineWidth = Math.max(4, Number(memeSize) * 0.09);
    ctx.textAlign = 'center';

    if (memeTop) {
      ctx.textBaseline = 'top';
      ctx.strokeText(memeTop.toUpperCase(), canvas.width / 2, memeSize * 0.35, canvas.width * 0.9);
      ctx.fillText(memeTop.toUpperCase(), canvas.width / 2, memeSize * 0.35, canvas.width * 0.9);
    }
    if (memeBottom) {
      ctx.textBaseline = 'bottom';
      ctx.strokeText(memeBottom.toUpperCase(), canvas.width / 2, canvas.height - memeSize * 0.35, canvas.width * 0.9);
      ctx.fillText(memeBottom.toUpperCase(), canvas.width / 2, canvas.height - memeSize * 0.35, canvas.width * 0.9);
    }
    ctx.restore();
    setStatus('Created meme', 'success');
  };

  // Action: Rotate
  const executeRotate = (deg) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    const rad = (deg * Math.PI) / 180;
    const swap = deg === 90 || deg === 270;
    const temp = makeTempCanvas(canvas.width, canvas.height);
    temp.ctx.drawImage(canvas, 0, 0);

    canvas.width = swap ? temp.canvas.height : temp.canvas.width;
    canvas.height = swap ? temp.canvas.width : temp.canvas.height;
    const ctx = canvas.getContext('2d');
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate(rad);
    ctx.drawImage(temp.canvas, -temp.canvas.width / 2, -temp.canvas.height / 2);
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    setDimensions({ width: canvas.width, height: canvas.height });
    setStatus(`Rotated ${deg}°`, 'success');
  };

  // Action: Flip
  const executeFlip = (dir) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    const temp = makeTempCanvas(canvas.width, canvas.height);
    temp.ctx.drawImage(canvas, 0, 0);
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.save();
    if (dir === 'h') {
      ctx.translate(canvas.width, 0);
      ctx.scale(-1, 1);
    } else {
      ctx.translate(0, canvas.height);
      ctx.scale(1, -1);
    }
    ctx.drawImage(temp.canvas, 0, 0);
    ctx.restore();
    setStatus(`Flipped ${dir === 'h' ? 'horizontally' : 'vertically'}`, 'success');
  };

  // Action: Blur Area
  const executeBlurBox = () => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    remember();
    const x = Math.max(0, Math.min(canvas.width - 1, Number(blurX)));
    const y = Math.max(0, Math.min(canvas.height - 1, Number(blurY)));
    const w = Math.max(1, Math.min(canvas.width - x, Number(blurW)));
    const h = Math.max(1, Math.min(canvas.height - y, Number(blurH)));

    const temp = makeTempCanvas(w, h);
    temp.ctx.drawImage(canvas, x, y, w, h, 0, 0, w, h);
    const ctx = canvas.getContext('2d');
    ctx.save();
    ctx.filter = `blur(${blurAmount}px)`;
    ctx.drawImage(temp.canvas, x, y);
    ctx.restore();
    setStatus('Applied privacy blur box', 'success');
  };

  // Action: Render HTML
  const executeRenderHtml = () => {
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${htmlW}" height="${htmlH}"><foreignObject width="100%" height="100%">${htmlCode}</foreignObject></svg>`;
    const img = new Image();
    img.onload = () => {
      remember();
      const canvas = canvasRef.current;
      if (!canvas) return;
      canvas.width = htmlW;
      canvas.height = htmlH;
      const ctx = canvas.getContext('2d');
      ctx.clearRect(0, 0, htmlW, htmlH);
      ctx.drawImage(img, 0, 0);
      const snap = ctx.getImageData(0, 0, htmlW, htmlH);
      originalRef.current = { width: htmlW, height: htmlH, imageData: snap };
      setDimensions({ width: htmlW, height: htmlH });
      setHasImage(true);
      setStatus('Rendered HTML to canvas image', 'success');
    };
    img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  };

  // Action: Download
  const executeDownload = (fmt = 'image/png', q = 0.92) => {
    const canvas = canvasRef.current;
    if (!canvas || !hasImage) return;

    let targetCanvas = canvas;
    if (fmt === 'image/jpeg') {
      const temp = makeTempCanvas(canvas.width, canvas.height);
      temp.ctx.fillStyle = convertBg || '#ffffff';
      temp.ctx.fillRect(0, 0, canvas.width, canvas.height);
      temp.ctx.drawImage(canvas, 0, 0);
      targetCanvas = temp.canvas;
    }

    const ext = fmt === 'image/jpeg' ? 'jpg' : fmt === 'image/webp' ? 'webp' : 'png';
    const baseName = (files[activeFileIndex]?.name || 'image').replace(/\.[^/.]+$/, '');
    
    targetCanvas.toBlob((blob) => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${baseName}_edited.${ext}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setStatus(`Downloaded ${baseName}_edited.${ext}`, 'success');
    }, fmt, Number(q));
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
            <Sparkles size={18} color="var(--accent)" /> Image Tools & Signature Studio
          </h2>
        </div>

        {/* Global Controls & Status */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <div style={{
            padding: '4px 10px',
            borderRadius: '16px',
            fontSize: '12px',
            fontWeight: 600,
            background: statusType === 'success' ? 'var(--accent-soft)' : 'var(--bg)',
            color: statusType === 'success' ? 'var(--accent)' : 'var(--muted)',
            border: '1px solid var(--line)'
          }}>
            {statusMsg}
          </div>

          <button 
            type="button" 
            className="secondary"
            disabled={!historyLength}
            onClick={handleUndo}
            title="Undo"
            style={{ display: 'flex', alignItems: 'center', gap: '4px', padding: '6px 12px', fontSize: '12px' }}
          >
            <Undo2 size={14} /> Undo ({historyLength})
          </button>

          <button 
            type="button" 
            className="secondary"
            disabled={!hasImage}
            onClick={handleReset}
            title="Reset"
            style={{ display: 'flex', alignItems: 'center', gap: '4px', padding: '6px 12px', fontSize: '12px' }}
          >
            <RotateCcw size={14} /> Reset
          </button>

          <button 
            type="button" 
            disabled={!hasImage}
            onClick={() => executeDownload(convertFormat, convertQuality)}
            style={{ display: 'flex', alignItems: 'center', gap: '6px', padding: '6px 16px', fontSize: '12px' }}
          >
            <Download size={14} /> Download Image
          </button>
        </div>
      </header>

      {/* Main Workspace Layout */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        {/* Left Tools Navigation */}
        <aside style={{ width: '250px', background: 'var(--panel)', borderRight: '1px solid var(--line)', display: 'flex', flexDirection: 'column', overflowY: 'auto', padding: '12px' }}>
          {/* File Dropzone */}
          <label style={{
            border: '1.5px dashed var(--accent)',
            borderRadius: '8px',
            padding: '14px 10px',
            textAlign: 'center',
            cursor: 'pointer',
            background: 'var(--accent-soft)',
            marginBottom: '14px',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            gap: '4px'
          }}>
            <input type="file" accept="image/*,.svg" multiple onChange={handleFileUpload} style={{ display: 'none' }} />
            <ImageIcon size={22} color="var(--accent)" />
            <strong style={{ fontSize: '12px', color: 'var(--ink)' }}>Choose / Drop Images</strong>
            <span style={{ fontSize: '10.5px', color: 'var(--muted)' }}>PNG, JPG, WebP, SVG</span>
          </label>

          {/* Queue List (if multiple files) */}
          {files.length > 1 && (
            <div style={{ marginBottom: '12px' }}>
              <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', marginBottom: '6px', textTransform: 'uppercase' }}>Image Queue ({files.length})</div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '4px', maxHeight: '100px', overflowY: 'auto' }}>
                {files.map((file, idx) => (
                  <button
                    key={idx}
                    type="button"
                    onClick={() => selectFileFromQueue(idx)}
                    style={{
                      padding: '5px 8px',
                      fontSize: '11.5px',
                      textAlign: 'left',
                      background: activeFileIndex === idx ? 'var(--accent)' : 'transparent',
                      color: activeFileIndex === idx ? 'white' : 'var(--ink)',
                      border: '1px solid var(--line)',
                      borderRadius: '4px',
                      whiteSpace: 'nowrap',
                      overflow: 'hidden',
                      textOverflow: 'ellipsis'
                    }}
                  >
                    {file.name}
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Tools Menu */}
          <div style={{ fontSize: '11px', fontWeight: 700, color: 'var(--muted)', marginBottom: '6px', textTransform: 'uppercase' }}>Modules</div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
            {TOOLS.map((t) => {
              const Icon = t.icon;
              const isSelected = activeTool === t.id;
              return (
                <button
                  key={t.id}
                  type="button"
                  onClick={() => setActiveTool(t.id)}
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

        {/* Center Canvas Stage */}
        <main style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', background: 'var(--bg)', padding: '16px', position: 'relative', overflow: 'auto' }}>
          <div style={{
            position: 'relative',
            background: 'repeating-conic-gradient(#e2e8f0 0% 25%, #ffffff 0% 50%) 50% / 16px 16px',
            boxShadow: '0 4px 20px rgba(0,0,0,0.1)',
            borderRadius: '6px',
            overflow: 'hidden',
            maxWidth: '100%',
            maxHeight: '100%',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
          }}>
            <canvas 
              ref={canvasRef} 
              style={{
                display: hasImage ? 'block' : 'none',
                maxWidth: 'calc(100vw - 640px)',
                maxHeight: 'calc(100vh - 160px)',
                objectFit: 'contain'
              }} 
            />

            {!hasImage && (
              <div style={{ padding: '60px 40px', textAlign: 'center', color: 'var(--muted)' }}>
                <ImageIcon size={48} style={{ opacity: 0.3, marginBottom: '12px' }} />
                <h3 style={{ margin: '0 0 6px 0', fontSize: '16px', color: 'var(--ink)' }}>No Image Loaded</h3>
                <p style={{ margin: 0, fontSize: '13px' }}>Upload an image from the sidebar or click "Render HTML" to begin.</p>
              </div>
            )}
          </div>

          {hasImage && (
            <div style={{ position: 'absolute', bottom: '12px', left: '16px', fontSize: '11px', color: 'var(--muted)', background: 'rgba(255,255,255,0.85)', backdropFilter: 'blur(4px)', padding: '4px 10px', borderRadius: '12px', border: '1px solid var(--line)' }}>
              Canvas: <strong>{dimensions.width} × {dimensions.height} px</strong>
            </div>
          )}
        </main>

        {/* Right Settings & Action Panel */}
        <aside style={{ width: '320px', background: 'var(--panel)', borderLeft: '1px solid var(--line)', padding: '16px 20px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: '14px' }}>
          <div>
            <h3 style={{ margin: '0 0 4px 0', fontSize: '15px' }}>{TOOLS.find(t => t.id === activeTool)?.name}</h3>
            <p style={{ margin: 0, fontSize: '12px', color: 'var(--muted)' }}>{TOOLS.find(t => t.id === activeTool)?.hint}</p>
          </div>

          <div style={{ height: '1px', background: 'var(--line)' }} />

          {/* 1. Signature / Background Remover */}
          {activeTool === 'removebg' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Extraction Mode</label>
                <select value={bgMode} onChange={(e) => setBgMode(e.target.value)}>
                  <option value="signature">Signature & Paper Cleaner</option>
                  <option value="corner">Sample Top-Left Color</option>
                  <option value="green">Green Screen Removal</option>
                </select>
              </div>

              <div className="form-group">
                <label>Paper / Color Tolerance: {bgThreshold}</label>
                <input type="range" min="0" max="160" value={bgThreshold} onChange={(e) => setBgThreshold(Number(e.target.value))} />
              </div>

              <div className="form-group">
                <label>Edge Softness: {bgSoftness}</label>
                <input type="range" min="0" max="100" value={bgSoftness} onChange={(e) => setBgSoftness(Number(e.target.value))} />
              </div>

              {bgMode === 'signature' && (
                <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px', fontWeight: 600, cursor: 'pointer' }}>
                  <input type="checkbox" checked={bgDarken} onChange={(e) => setBgDarken(e.target.checked)} />
                  Enhance & Darken Ink
                </label>
              )}

              <button type="button" onClick={executeRemoveBg} disabled={!hasImage}>
                ✨ Clean Background
              </button>
            </div>
          )}

          {/* 2. Compress */}
          {activeTool === 'compress' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Export Format</label>
                <select value={compressFormat} onChange={(e) => setCompressFormat(e.target.value)}>
                  <option value="image/jpeg">JPG (Standard Web)</option>
                  <option value="image/webp">WebP (High Compression)</option>
                  <option value="image/png">PNG (Lossless)</option>
                </select>
              </div>

              <div className="form-group">
                <label>Quality: {Math.round(compressQuality * 100)}%</label>
                <input type="range" min="0.1" max="1.0" step="0.01" value={compressQuality} onChange={(e) => setCompressQuality(Number(e.target.value))} />
              </div>

              <button type="button" onClick={() => executeDownload(compressFormat, compressQuality)} disabled={!hasImage}>
                🗜️ Download Compressed Image
              </button>
            </div>
          )}

          {/* 3. Resize */}
          {activeTool === 'resize' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-row">
                <div className="form-group">
                  <label>Width (px)</label>
                  <input type="number" min="1" value={resizeWidth} onChange={(e) => {
                    const w = Number(e.target.value);
                    setResizeWidth(w);
                    if (resizeLock && dimensions.width) setResizeHeight(Math.round((w / dimensions.width) * dimensions.height));
                  }} />
                </div>
                <div className="form-group">
                  <label>Height (px)</label>
                  <input type="number" min="1" value={resizeHeight} onChange={(e) => {
                    const h = Number(e.target.value);
                    setResizeHeight(h);
                    if (resizeLock && dimensions.height) setResizeWidth(Math.round((h / dimensions.height) * dimensions.width));
                  }} />
                </div>
              </div>

              <div className="form-group">
                <label>Scaling Percentage: {resizePercent}%</label>
                <input type="range" min="1" max="200" value={resizePercent} onChange={(e) => setResizePercent(Number(e.target.value))} />
              </div>

              <label style={{ display: 'flex', alignItems: 'center', gap: '8px', fontSize: '12px', fontWeight: 600, cursor: 'pointer' }}>
                <input type="checkbox" checked={resizeLock} onChange={(e) => setResizeLock(e.target.checked)} />
                Lock Aspect Ratio
              </label>

              <button type="button" onClick={executeResize} disabled={!hasImage}>
                Apply Dimensions
              </button>
            </div>
          )}

          {/* 4. Crop */}
          {activeTool === 'crop' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-row">
                <div className="form-group">
                  <label>X Offset</label>
                  <input type="number" min="0" value={cropX} onChange={(e) => setCropX(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Y Offset</label>
                  <input type="number" min="0" value={cropY} onChange={(e) => setCropY(Number(e.target.value))} />
                </div>
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Crop Width</label>
                  <input type="number" min="1" value={cropW} onChange={(e) => setCropW(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Crop Height</label>
                  <input type="number" min="1" value={cropH} onChange={(e) => setCropH(Number(e.target.value))} />
                </div>
              </div>

              <button type="button" onClick={executeCrop} disabled={!hasImage}>
                ✂️ Execute Crop
              </button>
            </div>
          )}

          {/* 5. Convert */}
          {activeTool === 'convert' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Target Format</label>
                <select value={convertFormat} onChange={(e) => setConvertFormat(e.target.value)}>
                  <option value="image/png">PNG (Preserves Transparency)</option>
                  <option value="image/jpeg">JPG (Solid Background)</option>
                  <option value="image/webp">WebP (Optimized)</option>
                </select>
              </div>

              {convertFormat === 'image/jpeg' && (
                <div className="form-group">
                  <label>Background Fill Color</label>
                  <input type="color" value={convertBg} onChange={(e) => setConvertBg(e.target.value)} />
                </div>
              )}

              <button type="button" onClick={() => executeDownload(convertFormat, convertQuality)} disabled={!hasImage}>
                Download As {convertFormat.split('/')[1].toUpperCase()}
              </button>
            </div>
          )}

          {/* 6. Photo Editor */}
          {activeTool === 'editor' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Brightness: {Math.round(brightness * 100)}%</label>
                <input type="range" min="0" max="2" step="0.05" value={brightness} onChange={(e) => setBrightness(Number(e.target.value))} />
              </div>

              <div className="form-group">
                <label>Contrast: {Math.round(contrast * 100)}%</label>
                <input type="range" min="0" max="2" step="0.05" value={contrast} onChange={(e) => setContrast(Number(e.target.value))} />
              </div>

              <div className="form-group">
                <label>Saturation: {Math.round(saturate * 100)}%</label>
                <input type="range" min="0" max="2" step="0.05" value={saturate} onChange={(e) => setSaturate(Number(e.target.value))} />
              </div>

              <div className="form-group">
                <label>Soft Blur: {softBlur}px</label>
                <input type="range" min="0" max="8" step="0.5" value={softBlur} onChange={(e) => setSoftBlur(Number(e.target.value))} />
              </div>

              <div style={{ display: 'flex', gap: '16px' }}>
                <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', cursor: 'pointer' }}>
                  <input type="checkbox" checked={grayscale} onChange={(e) => setGrayscale(e.target.checked)} /> Grayscale
                </label>
                <label style={{ display: 'flex', alignItems: 'center', gap: '6px', fontSize: '12px', cursor: 'pointer' }}>
                  <input type="checkbox" checked={sepia} onChange={(e) => setSepia(e.target.checked)} /> Sepia
                </label>
              </div>

              <button type="button" onClick={executeFilters} disabled={!hasImage}>
                Apply Adjustments
              </button>
            </div>
          )}

          {/* 7. Upscale */}
          {activeTool === 'upscale' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Multiplier</label>
                <select value={upscaleFactor} onChange={(e) => setUpscaleFactor(Number(e.target.value))}>
                  <option value={2}>2x Super Scale</option>
                  <option value={3}>3x High Detail</option>
                  <option value={4}>4x Ultra High Definition</option>
                </select>
              </div>

              <button type="button" onClick={executeUpscale} disabled={!hasImage}>
                ⚡ Upscale Canvas
              </button>
            </div>
          )}

          {/* 8. Watermark */}
          {activeTool === 'watermark' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Watermark Text</label>
                <input type="text" value={wmText} onChange={(e) => setWmText(e.target.value)} />
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Font Size ({wmSize}px)</label>
                  <input type="number" min="12" max="160" value={wmSize} onChange={(e) => setWmSize(Number(e.target.value))} />
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
                <label>Placement</label>
                <select value={wmPosition} onChange={(e) => setWmPosition(e.target.value)}>
                  <option value="center">Center</option>
                  <option value="top-left">Top Left</option>
                  <option value="top-right">Top Right</option>
                  <option value="bottom-left">Bottom Left</option>
                  <option value="bottom-right">Bottom Right</option>
                  <option value="tile">Tile Across Entire Page</option>
                </select>
              </div>

              <button type="button" onClick={executeWatermark} disabled={!hasImage}>
                Stamp Watermark
              </button>
            </div>
          )}

          {/* 9. Meme Generator */}
          {activeTool === 'meme' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Top Caption</label>
                <input type="text" value={memeTop} onChange={(e) => setMemeTop(e.target.value)} />
              </div>

              <div className="form-group">
                <label>Bottom Caption</label>
                <input type="text" value={memeBottom} onChange={(e) => setMemeBottom(e.target.value)} />
              </div>

              <div className="form-group">
                <label>Font Size ({memeSize}px)</label>
                <input type="number" min="20" max="120" value={memeSize} onChange={(e) => setMemeSize(Number(e.target.value))} />
              </div>

              <button type="button" onClick={executeMeme} disabled={!hasImage}>
                Render Caption
              </button>
            </div>
          )}

          {/* 10. Rotate & Flip */}
          {activeTool === 'rotate' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
              <button type="button" onClick={() => executeRotate(90)} disabled={!hasImage}>
                🔄 Rotate 90° Clockwise
              </button>
              <button type="button" onClick={() => executeRotate(180)} disabled={!hasImage}>
                🔄 Rotate 180°
              </button>
              <button type="button" onClick={() => executeRotate(270)} disabled={!hasImage}>
                🔄 Rotate 270°
              </button>
              <button type="button" className="secondary" onClick={() => executeFlip('h')} disabled={!hasImage}>
                ↔️ Flip Horizontally
              </button>
              <button type="button" className="secondary" onClick={() => executeFlip('v')} disabled={!hasImage}>
                ↕️ Flip Vertically
              </button>
            </div>
          )}

          {/* 11. Privacy Blur Box */}
          {activeTool === 'blur' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-row">
                <div className="form-group">
                  <label>X Position</label>
                  <input type="number" min="0" value={blurX} onChange={(e) => setBlurX(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Y Position</label>
                  <input type="number" min="0" value={blurY} onChange={(e) => setBlurY(Number(e.target.value))} />
                </div>
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Box Width</label>
                  <input type="number" min="1" value={blurW} onChange={(e) => setBlurW(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Box Height</label>
                  <input type="number" min="1" value={blurH} onChange={(e) => setBlurH(Number(e.target.value))} />
                </div>
              </div>

              <div className="form-group">
                <label>Blur Intensity ({blurAmount}px)</label>
                <input type="range" min="2" max="40" value={blurAmount} onChange={(e) => setBlurAmount(Number(e.target.value))} />
              </div>

              <button type="button" onClick={executeBlurBox} disabled={!hasImage}>
                🙈 Censor & Blur Box
              </button>
            </div>
          )}

          {/* 12. HTML to Image */}
          {activeTool === 'html' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
              <div className="form-group">
                <label>Custom HTML Code</label>
                <textarea 
                  rows={6}
                  value={htmlCode} 
                  onChange={(e) => setHtmlCode(e.target.value)} 
                  style={{ width: '100%', padding: '8px', border: '1px solid var(--line)', borderRadius: '6px', fontFamily: 'monospace', fontSize: '11.5px' }} 
                />
              </div>

              <div className="form-row">
                <div className="form-group">
                  <label>Width</label>
                  <input type="number" min="100" value={htmlW} onChange={(e) => setHtmlW(Number(e.target.value))} />
                </div>
                <div className="form-group">
                  <label>Height</label>
                  <input type="number" min="100" value={htmlH} onChange={(e) => setHtmlH(Number(e.target.value))} />
                </div>
              </div>

              <button type="button" onClick={executeRenderHtml}>
                💻 Render HTML to Canvas
              </button>
            </div>
          )}
        </aside>
      </div>
    </div>
  );
}
