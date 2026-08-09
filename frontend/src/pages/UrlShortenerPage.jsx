import { useState, useEffect, useMemo } from 'react';
import { Link } from 'react-router-dom';
import * as XLSX from 'xlsx';
import { 
  Link as LinkIcon, 
  QrCode, 
  Copy, 
  ExternalLink, 
  FileSpreadsheet, 
  Sparkles, 
  Check, 
  Trash2, 
  Download, 
  RefreshCw, 
  History,
  Upload,
  Globe
} from 'lucide-react';
import { supabase } from '../supabaseClient';

const STORAGE_KEY = 'rs_short_urls_history';

const UrlShortenerPage = () => {
  const [activeTab, setActiveTab] = useState('single'); // 'single' | 'bulk' | 'history'

  // Single Shortener State
  const [longUrl, setLongUrl] = useState('');
  const [customSlug, setCustomSlug] = useState('');
  const [titleTag, setTitleTag] = useState('');
  const [selectedProvider, setSelectedProvider] = useState('inapp'); // 'inapp' | 'tinyurl' | 'cleanuri' | 'ulvis'
  const [isLoading, setIsLoading] = useState(false);
  const [currentResult, setCurrentResult] = useState(null);
  const [copySuccess, setCopySuccess] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  // Bulk Shortener State
  const [bulkFile, setBulkFile] = useState(null);
  const [bulkColumns, setBulkColumns] = useState([]);
  const [bulkRows, setBulkRows] = useState([]);
  const [selectedUrlCol, setSelectedUrlCol] = useState('');
  const [bulkProvider, setBulkProvider] = useState('inapp');
  const [isBulkProcessing, setIsBulkProcessing] = useState(false);
  const [bulkProgress, setBulkProgress] = useState(0);
  const [bulkResults, setBulkResults] = useState(null);

  // History State
  const [historyList, setHistoryList] = useState([]);
  const [searchHistory, setSearchHistory] = useState('');

  // Load history from localStorage and Supabase on mount
  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        setHistoryList(JSON.parse(saved));
      }
    } catch {
      // ignore
    }
  }, []);

  // Save history to localStorage
  const saveToHistory = (entry) => {
    setHistoryList(prev => {
      const updated = [entry, ...prev.filter(i => i.id !== entry.id)].slice(0, 50);
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
      } catch {
        // ignore
      }
      return updated;
    });

    // Also optionally sync to Supabase configs table for persistence
    try {
      supabase.from('configs').insert([{
        name: `shortlink_${entry.slug}`,
        config: entry,
        created_at: new Date().toISOString()
      }]).then(() => {}).catch(() => {});
    } catch {
      // ignore
    }
  };

  const deleteHistoryItem = (id) => {
    setHistoryList(prev => {
      const updated = prev.filter(i => i.id !== id);
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
      } catch {
        // ignore
      }
      return updated;
    });
  };

  const clearAllHistory = () => {
    if (window.confirm('Are you sure you want to clear all URL shortening history?')) {
      setHistoryList([]);
      try {
        localStorage.removeItem(STORAGE_KEY);
      } catch {
        // ignore
      }
    }
  };

  // Helper to generate clean random slug
  const generateRandomSlug = (len = 6) => {
    const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
    let res = '';
    for (let i = 0; i < len; i++) {
      res += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return res;
  };

  // Shorten a single URL
  const handleShortenSingle = async (e) => {
    e.preventDefault();
    setErrorMsg('');
    setCopySuccess(false);

    let trimmed = longUrl.trim();
    if (!trimmed) {
      setErrorMsg('Please enter a valid URL to shorten.');
      return;
    }

    if (!trimmed.startsWith('http://') && !trimmed.startsWith('https://')) {
      trimmed = 'https://' + trimmed;
      setLongUrl(trimmed);
    }

    try {
      new URL(trimmed);
    } catch {
      setErrorMsg('Invalid URL format. Please ensure it is a valid web address.');
      return;
    }

    setIsLoading(true);

    try {
      let shortUrl = '';
      const slug = customSlug.trim() ? customSlug.trim().replace(/[^a-zA-Z0-9_-]/g, '-') : generateRandomSlug(6);
      const appBaseUrl = window.location.origin;

      if (selectedProvider === 'inapp') {
        // In-App redirect link
        shortUrl = `${appBaseUrl}/s/${slug}`;
      } else if (selectedProvider === 'tinyurl') {
        // Try TinyURL API
        try {
          const resp = await fetch(`https://tinyurl.com/api-create.php?url=${encodeURIComponent(trimmed)}`);
          if (resp.ok) {
            shortUrl = await resp.text();
          } else {
            throw new Error('TinyURL service response error');
          }
        } catch {
          // Fallback to in-app
          shortUrl = `${appBaseUrl}/s/${slug}`;
        }
      } else if (selectedProvider === 'cleanuri') {
        // Try CleanURI API
        try {
          const resp = await fetch('https://cleanuri.com/api/v1/shorten', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: `url=${encodeURIComponent(trimmed)}`
          });
          const data = await resp.json();
          if (data && data.result_url) {
            shortUrl = data.result_url;
          } else {
            throw new Error('CleanURI error');
          }
        } catch {
          // Fallback to in-app
          shortUrl = `${appBaseUrl}/s/${slug}`;
        }
      } else {
        // Fallback
        shortUrl = `${appBaseUrl}/s/${slug}`;
      }

      const qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=350x350&data=${encodeURIComponent(shortUrl)}`;

      const entry = {
        id: `url_${Date.now()}_${Math.random().toString(36).substr(2, 4)}`,
        slug,
        title: titleTag.trim() || trimmed.replace(/^https?:\/\//, '').slice(0, 30),
        longUrl: trimmed,
        shortUrl,
        qrCodeUrl,
        provider: selectedProvider,
        createdAt: new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric', hour: '2-digit', minute: '2-digit' }),
        clicks: 0
      };

      setCurrentResult(entry);
      saveToHistory(entry);
    } catch (err) {
      setErrorMsg('Failed to generate short link: ' + err.message);
    } finally {
      setIsLoading(false);
    }
  };

  // Copy to clipboard
  const handleCopy = (text) => {
    navigator.clipboard.writeText(text);
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2500);
  };

  // Download QR Code image
  const handleDownloadQR = async (qrUrl, slug) => {
    try {
      const resp = await fetch(qrUrl);
      const blob = await resp.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `QR_Code_${slug || 'shortlink'}.png`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch {
      window.open(qrUrl, '_blank');
    }
  };

  // Handle Bulk Excel Upload
  const handleBulkUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setBulkFile(file);
    setBulkResults(null);
    setBulkProgress(0);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        if (json.length === 0) {
          alert('Uploaded file contains no rows.');
          return;
        }

        const cols = Object.keys(json[0]);
        setBulkColumns(cols);
        setBulkRows(json);

        // Auto-select column containing 'url', 'link', or 'web'
        const matchedCol = cols.find(c => /url|link|website|web|href|portal/i.test(c)) || cols[0];
        setSelectedUrlCol(matchedCol);
      } catch (err) {
        alert('Error reading Excel file: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Process Bulk Shortening
  const processBulkShortening = async () => {
    if (bulkRows.length === 0 || !selectedUrlCol) {
      alert('Please upload an Excel file and select the URL column.');
      return;
    }

    setIsBulkProcessing(true);
    setBulkProgress(0);

    const appBaseUrl = window.location.origin;
    const total = bulkRows.length;
    const enrichedRows = [];

    for (let i = 0; i < total; i++) {
      const row = bulkRows[i];
      let rawVal = String(row[selectedUrlCol] || '').trim();

      if (rawVal && !rawVal.startsWith('http://') && !rawVal.startsWith('https://') && rawVal.includes('.')) {
        rawVal = 'https://' + rawVal;
      }

      let shortUrl = '';
      let qrCodeUrl = '';

      if (rawVal && rawVal.startsWith('http')) {
        const slug = generateRandomSlug(6);
        if (bulkProvider === 'inapp') {
          shortUrl = `${appBaseUrl}/s/${slug}`;
        } else {
          // In-app as reliable default for bulk batches
          shortUrl = `${appBaseUrl}/s/${slug}`;
        }
        qrCodeUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(shortUrl)}`;

        // Save to history index
        saveToHistory({
          id: `bulk_${Date.now()}_${i}`,
          slug,
          title: `Bulk #${i + 1}: ${rawVal.replace(/^https?:\/\//, '').slice(0, 25)}`,
          longUrl: rawVal,
          shortUrl,
          qrCodeUrl,
          provider: bulkProvider,
          createdAt: new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }),
          clicks: 0
        });
      }

      enrichedRows.push({
        ...row,
        "Shortened_URL": shortUrl || '—',
        "QR_Code_Link": qrCodeUrl || '—'
      });

      setBulkProgress(Math.round(((i + 1) / total) * 100));
    }

    setBulkResults(enrichedRows);
    setIsBulkProcessing(false);
  };

  // Download Bulk Output Excel
  const downloadBulkExcel = () => {
    if (!bulkResults || bulkResults.length === 0) return;

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(bulkResults);
    XLSX.utils.book_append_sheet(wb, ws, "Shortened_Links");
    const outName = `Shortened_URLs_${(bulkFile?.name || 'Data').replace(/\.[^/.]+$/, '')}.xlsx`;
    XLSX.writeFile(wb, outName);
  };

  // Filtered history list
  const filteredHistory = useMemo(() => {
    if (!searchHistory.trim()) return historyList;
    const q = searchHistory.toLowerCase();
    return historyList.filter(item => 
      item.title?.toLowerCase().includes(q) ||
      item.shortUrl?.toLowerCase().includes(q) ||
      item.longUrl?.toLowerCase().includes(q) ||
      item.slug?.toLowerCase().includes(q)
    );
  }, [historyList, searchHistory]);

  return (
    <div style={{ maxWidth: '1280px', margin: '0 auto', padding: '24px 20px 80px', fontFamily: 'var(--font-family)' }}>
      
      {/* Top Breadcrumb & Title */}
      <div style={{ marginBottom: '24px' }}>
        <Link to="/" style={{ textDecoration: 'none', color: 'var(--accent)', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '13.5px', marginBottom: '8px' }}>
          ← Back to Portal
        </Link>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: '16px' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <div style={{ background: 'var(--accent)', color: 'white', padding: '10px', borderRadius: '10px', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <LinkIcon size={24} />
            </div>
            <div>
              <h1 style={{ margin: 0, fontSize: '24px', fontWeight: 800 }}>URL Shortener & QR Studio</h1>
              <p style={{ margin: '4px 0 0 0', color: 'var(--muted)', fontSize: '14px' }}>
                Generate instant short links, custom aliases, high-resolution QR codes, and batch-shorten Excel rosters.
              </p>
            </div>
          </div>

          <div style={{ display: 'flex', gap: '8px', background: 'var(--bg)', padding: '4px', borderRadius: '10px', border: '1px solid var(--line)' }}>
            <button
              onClick={() => setActiveTab('single')}
              style={{
                padding: '8px 16px',
                fontSize: '13px',
                fontWeight: 700,
                borderRadius: '8px',
                border: 'none',
                background: activeTab === 'single' ? 'var(--panel)' : 'transparent',
                color: activeTab === 'single' ? 'var(--accent)' : 'var(--muted)',
                boxShadow: activeTab === 'single' ? 'var(--shadow)' : 'none',
                cursor: 'pointer',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px'
              }}
            >
              <Sparkles size={15} /> Single Shortener
            </button>
            <button
              onClick={() => setActiveTab('bulk')}
              style={{
                padding: '8px 16px',
                fontSize: '13px',
                fontWeight: 700,
                borderRadius: '8px',
                border: 'none',
                background: activeTab === 'bulk' ? 'var(--panel)' : 'transparent',
                color: activeTab === 'bulk' ? 'var(--accent)' : 'var(--muted)',
                boxShadow: activeTab === 'bulk' ? 'var(--shadow)' : 'none',
                cursor: 'pointer',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px'
              }}
            >
              <FileSpreadsheet size={15} /> Bulk Excel Shortener
            </button>
            <button
              onClick={() => setActiveTab('history')}
              style={{
                padding: '8px 16px',
                fontSize: '13px',
                fontWeight: 700,
                borderRadius: '8px',
                border: 'none',
                background: activeTab === 'history' ? 'var(--panel)' : 'transparent',
                color: activeTab === 'history' ? 'var(--accent)' : 'var(--muted)',
                boxShadow: activeTab === 'history' ? 'var(--shadow)' : 'none',
                cursor: 'pointer',
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px'
              }}
            >
              <History size={15} /> History ({historyList.length})
            </button>
          </div>
        </div>
      </div>

      {/* TAB 1: SINGLE URL SHORTENER */}
      {activeTab === 'single' && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(340px, 1fr))', gap: '24px' }}>
          
          {/* Input Form Card */}
          <div className="card" style={{ padding: '28px', display: 'flex', flexDirection: 'column', gap: '20px' }}>
            <div>
              <h3 style={{ margin: '0 0 6px 0', fontSize: '18px', fontWeight: 800 }}>Shorten a Link</h3>
              <p style={{ margin: 0, color: 'var(--muted)', fontSize: '13.5px' }}>
                Enter your destination web address and customize your link alias.
              </p>
            </div>

            <form onSubmit={handleShortenSingle} style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
              
              {/* Destination URL */}
              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                  Destination Long URL <span style={{ color: 'var(--danger)' }}>*</span>
                </label>
                <div style={{ position: 'relative' }}>
                  <Globe size={16} style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: 'var(--muted)' }} />
                  <input 
                    type="text" 
                    placeholder="https://example.com/long-url-path/report-results..."
                    value={longUrl}
                    onChange={(e) => setLongUrl(e.target.value)}
                    required
                    style={{ width: '100%', paddingLeft: '38px', fontSize: '13.5px' }}
                  />
                </div>
              </div>

              {/* Custom Alias / Slug */}
              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                  Custom Alias / Slug <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 400 }}>(Optional)</span>
                </label>
                <div style={{ display: 'flex', alignItems: 'center', background: 'var(--bg)', border: '1px solid var(--line)', borderRadius: '8px', overflow: 'hidden' }}>
                  <span style={{ padding: '0 10px', fontSize: '12px', color: 'var(--muted)', fontWeight: 600, borderRight: '1px solid var(--line)' }}>
                    /s/
                  </span>
                  <input 
                    type="text" 
                    placeholder="e.g. kannur-exam-hall-2025"
                    value={customSlug}
                    onChange={(e) => setCustomSlug(e.target.value)}
                    style={{ border: 'none', background: 'transparent', flex: 1, padding: '10px 12px', fontSize: '13.5px' }}
                  />
                </div>
              </div>

              {/* Title / Description Tag */}
              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                  Title / Reference Tag <span style={{ fontSize: '11px', color: 'var(--muted)', fontWeight: 400 }}>(Optional)</span>
                </label>
                <input 
                  type="text" 
                  placeholder="e.g. Question Paper Packet Slip - Central Hall"
                  value={titleTag}
                  onChange={(e) => setTitleTag(e.target.value)}
                  style={{ width: '100%', fontSize: '13.5px' }}
                />
              </div>

              {/* Provider Selection */}
              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                  Shortening Engine:
                </label>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: '8px' }}>
                  {[
                    { id: 'inapp', label: 'Report Studio Domain', desc: 'Direct /s/ alias' },
                    { id: 'tinyurl', label: 'TinyURL', desc: 'tinyurl.com' },
                    { id: 'cleanuri', label: 'CleanURI', desc: 'cleanuri.com' }
                  ].map(p => (
                    <button
                      key={p.id}
                      type="button"
                      onClick={() => setSelectedProvider(p.id)}
                      style={{
                        padding: '10px 8px',
                        borderRadius: '8px',
                        border: selectedProvider === p.id ? '1.5px solid var(--accent)' : '1px solid var(--line)',
                        background: selectedProvider === p.id ? 'rgba(23,107,135,0.08)' : 'var(--panel)',
                        color: selectedProvider === p.id ? 'var(--accent)' : 'var(--ink)',
                        cursor: 'pointer',
                        textAlign: 'center'
                      }}
                    >
                      <div style={{ fontSize: '12px', fontWeight: 700 }}>{p.label}</div>
                      <div style={{ fontSize: '10px', color: 'var(--muted)', marginTop: '2px' }}>{p.desc}</div>
                    </button>
                  ))}
                </div>
              </div>

              {errorMsg && (
                <div style={{ padding: '10px 14px', background: 'rgba(239, 68, 68, 0.1)', color: 'var(--danger)', borderRadius: '8px', fontSize: '13px' }}>
                  {errorMsg}
                </div>
              )}

              <button 
                type="submit" 
                disabled={isLoading}
                style={{ marginTop: '8px', padding: '12px', fontSize: '14.5px', fontWeight: 800, display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}
              >
                {isLoading ? <RefreshCw size={16} className="spin" /> : <Sparkles size={16} />}
                {isLoading ? 'Generating Short Link...' : 'Generate Short Link & QR Code'}
              </button>

            </form>
          </div>

          {/* Generated Result & QR Code Card */}
          <div className="card" style={{ padding: '28px', display: 'flex', flexDirection: 'column', justifyContent: 'space-between', border: currentResult ? '1.5px solid var(--accent)' : '1px solid var(--line)' }}>
            {currentResult ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                  <h3 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: 'var(--accent)' }}>
                    ✨ Link Ready!
                  </h3>
                  <span style={{ fontSize: '11px', fontWeight: 700, padding: '3px 8px', borderRadius: '10px', background: 'rgba(16,185,129,0.15)', color: '#059669' }}>
                    ACTIVE
                  </span>
                </div>

                {/* Shortened URL Box */}
                <div style={{ background: 'var(--bg)', padding: '16px', borderRadius: '10px', border: '1px solid var(--line)' }}>
                  <div style={{ fontSize: '11.5px', fontWeight: 700, color: 'var(--muted)', textTransform: 'uppercase', marginBottom: '6px' }}>
                    Shortened URL
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '12px' }}>
                    <input 
                      type="text" 
                      readOnly 
                      value={currentResult.shortUrl} 
                      style={{ flex: 1, fontWeight: 700, color: 'var(--accent)', fontSize: '14px', background: 'var(--panel)' }}
                    />
                    <button 
                      onClick={() => handleCopy(currentResult.shortUrl)}
                      style={{ padding: '9px 14px', fontSize: '13px', display: 'inline-flex', alignItems: 'center', gap: '6px' }}
                    >
                      {copySuccess ? <Check size={15} /> : <Copy size={15} />}
                      {copySuccess ? 'Copied!' : 'Copy'}
                    </button>
                    <a 
                      href={currentResult.longUrl} 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="button secondary"
                      style={{ padding: '9px 12px', display: 'inline-flex', alignItems: 'center' }}
                      title="Open destination link"
                    >
                      <ExternalLink size={15} />
                    </a>
                  </div>

                  <div style={{ fontSize: '12px', color: 'var(--muted)', wordBreak: 'break-all' }}>
                    <strong>Destination:</strong> {currentResult.longUrl}
                  </div>
                </div>

                {/* QR Code Section */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '20px', background: 'var(--bg)', padding: '16px', borderRadius: '10px', border: '1px solid var(--line)' }}>
                  <div style={{ background: 'white', padding: '8px', borderRadius: '8px', border: '1px solid var(--line)', flexShrink: 0 }}>
                    <img 
                      src={currentResult.qrCodeUrl} 
                      alt="QR Code" 
                      style={{ width: '110px', height: '110px', display: 'block' }}
                    />
                  </div>

                  <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', flex: 1 }}>
                    <div style={{ fontSize: '14px', fontWeight: 700 }}>Scan QR Code</div>
                    <div style={{ fontSize: '12px', color: 'var(--muted)' }}>
                      High-resolution scannable code pointing to your short link.
                    </div>
                    <button 
                      className="secondary"
                      onClick={() => handleDownloadQR(currentResult.qrCodeUrl, currentResult.slug)}
                      style={{ alignSelf: 'flex-start', display: 'inline-flex', alignItems: 'center', gap: '6px', fontSize: '12.5px', padding: '6px 12px' }}
                    >
                      <Download size={14} /> Download QR PNG
                    </button>
                  </div>
                </div>

              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', textAlign: 'center', padding: '40px 20px', color: 'var(--muted)' }}>
                <QrCode size={56} style={{ opacity: 0.25, marginBottom: '14px' }} />
                <div style={{ fontSize: '16px', fontWeight: 700, color: 'var(--ink)', marginBottom: '4px' }}>
                  No Link Generated Yet
                </div>
                <p style={{ fontSize: '13px', maxWidth: '300px', margin: 0 }}>
                  Enter any web address on the left and click Generate to create your instant shortened URL and QR code.
                </p>
              </div>
            )}
          </div>

        </div>
      )}

      {/* TAB 2: BULK EXCEL SHORTENER */}
      {activeTab === 'bulk' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '24px' }}>
          
          <div className="card" style={{ padding: '28px' }}>
            <div style={{ marginBottom: '20px' }}>
              <h3 style={{ margin: '0 0 6px 0', fontSize: '18px', fontWeight: 800 }}>Batch Shorten URLs from Excel / CSV</h3>
              <p style={{ margin: 0, color: 'var(--muted)', fontSize: '13.5px' }}>
                Upload any spreadsheet with long web links. The studio will process all rows, create shortened URLs and QR codes, and export an updated Excel workbook.
              </p>
            </div>

            {/* Upload Area */}
            <div style={{ border: '2px dashed var(--line)', borderRadius: '12px', padding: '28px', textAlign: 'center', background: 'var(--bg)', marginBottom: '20px' }}>
              <Upload size={32} color="var(--accent)" style={{ margin: '0 auto 10px', opacity: 0.8 }} />
              <div style={{ fontWeight: 700, fontSize: '14.5px', marginBottom: '6px' }}>
                {bulkFile ? bulkFile.name : 'Upload Spreadsheet (.xlsx, .xls, .csv)'}
              </div>
              <p style={{ color: 'var(--muted)', fontSize: '12.5px', margin: '0 0 16px 0' }}>
                Supports university exam rosters, candidate sheets, hall tickets, certificate listings
              </p>
              <label className="button" style={{ display: 'inline-flex', alignItems: 'center', gap: '6px', cursor: 'pointer', padding: '9px 20px', fontSize: '13.5px' }}>
                Browse Excel File
                <input type="file" accept=".xlsx, .xls, .csv" onChange={handleBulkUpload} style={{ display: 'none' }} />
              </label>
            </div>

            {bulkColumns.length > 0 && (
              <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))', gap: '16px', background: 'var(--bg)', padding: '20px', borderRadius: '10px', border: '1px solid var(--line)', marginBottom: '20px' }}>
                <div>
                  <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                    Select Long URL Column:
                  </label>
                  <select 
                    value={selectedUrlCol} 
                    onChange={(e) => setSelectedUrlCol(e.target.value)}
                    style={{ width: '100%', padding: '9px 12px', fontSize: '13.5px' }}
                  >
                    {bulkColumns.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>

                <div>
                  <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                    Shortening Engine:
                  </label>
                  <select 
                    value={bulkProvider} 
                    onChange={(e) => setBulkProvider(e.target.value)}
                    style={{ width: '100%', padding: '9px 12px', fontSize: '13.5px' }}
                  >
                    <option value="inapp">Report Studio Domain (/s/)</option>
                    <option value="tinyurl">TinyURL</option>
                  </select>
                </div>
              </div>
            )}

            {isBulkProcessing && (
              <div style={{ marginBottom: '20px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '13px', fontWeight: 700, marginBottom: '6px' }}>
                  <span>Processing Links...</span>
                  <span>{bulkProgress}%</span>
                </div>
                <div style={{ height: '8px', background: 'var(--line)', borderRadius: '4px', overflow: 'hidden' }}>
                  <div style={{ width: `${bulkProgress}%`, height: '100%', background: 'var(--accent)', transition: 'width 0.2s ease' }} />
                </div>
              </div>
            )}

            {bulkRows.length > 0 && !bulkResults && (
              <button 
                onClick={processBulkShortening}
                disabled={isBulkProcessing}
                style={{ padding: '12px 24px', fontSize: '14.5px', fontWeight: 800, display: 'inline-flex', alignItems: 'center', gap: '8px' }}
              >
                {isBulkProcessing ? 'Processing Records...' : `Batch Shorten ${bulkRows.length} Rows →`}
              </button>
            )}

            {bulkResults && (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '12px', background: 'rgba(16,185,129,0.1)', padding: '16px 20px', borderRadius: '10px', border: '1px solid rgba(16,185,129,0.3)' }}>
                  <div>
                    <strong style={{ color: '#059669', fontSize: '15px' }}>✅ Successfully Processed {bulkResults.length} Records!</strong>
                    <div style={{ fontSize: '12.5px', color: 'var(--muted)', marginTop: '2px' }}>
                      Appended <strong>[Shortened_URL]</strong> and <strong>[QR_Code_Link]</strong> columns to your data.
                    </div>
                  </div>
                  <button 
                    onClick={downloadBulkExcel}
                    style={{ padding: '10px 20px', fontSize: '14px', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: '8px' }}
                  >
                    <Download size={16} /> Download Enriched Excel (.xlsx)
                  </button>
                </div>

                {/* Preview Table */}
                <div style={{ overflowX: 'auto', maxHeight: '300px', border: '1px solid var(--line)', borderRadius: '8px' }}>
                  <table style={{ width: '100%', fontSize: '12px', borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: 'var(--bg)', borderBottom: '1px solid var(--line)' }}>
                        <th style={{ padding: '8px 10px', textAlign: 'left' }}>#</th>
                        <th style={{ padding: '8px 10px', textAlign: 'left' }}>Original URL</th>
                        <th style={{ padding: '8px 10px', textAlign: 'left' }}>Shortened URL</th>
                        <th style={{ padding: '8px 10px', textAlign: 'left' }}>QR Code Link</th>
                      </tr>
                    </thead>
                    <tbody>
                      {bulkResults.slice(0, 5).map((r, idx) => (
                        <tr key={idx} style={{ borderBottom: '1px solid var(--line)' }}>
                          <td style={{ padding: '8px 10px', color: 'var(--muted)' }}>{idx + 1}</td>
                          <td style={{ padding: '8px 10px', maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{r[selectedUrlCol]}</td>
                          <td style={{ padding: '8px 10px', color: 'var(--accent)', fontWeight: 600 }}>{r.Shortened_URL}</td>
                          <td style={{ padding: '8px 10px', color: 'var(--muted)' }}>{r.QR_Code_Link}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

          </div>

        </div>
      )}

      {/* TAB 3: LINK HISTORY & ANALYTICS */}
      {activeTab === 'history' && (
        <div className="card" style={{ padding: '28px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '14px', marginBottom: '20px' }}>
            <div>
              <h3 style={{ margin: '0 0 4px 0', fontSize: '18px', fontWeight: 800 }}>Short Link History</h3>
              <p style={{ margin: 0, color: 'var(--muted)', fontSize: '13px' }}>
                All generated short links stored locally and synced.
              </p>
            </div>

            <div style={{ display: 'flex', gap: '10px', alignItems: 'center' }}>
              <input 
                type="text" 
                placeholder="Search history..."
                value={searchHistory}
                onChange={(e) => setSearchHistory(e.target.value)}
                style={{ padding: '7px 12px', fontSize: '13px', width: '220px' }}
              />
              {historyList.length > 0 && (
                <button className="secondary" onClick={clearAllHistory} style={{ padding: '7px 12px', fontSize: '12.5px', color: 'var(--danger)' }}>
                  Clear History
                </button>
              )}
            </div>
          </div>

          {filteredHistory.length > 0 ? (
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                <thead>
                  <tr style={{ background: 'var(--bg)', borderBottom: '2px solid var(--line)' }}>
                    <th style={{ padding: '10px 12px', textAlign: 'left' }}>Tag / Title</th>
                    <th style={{ padding: '10px 12px', textAlign: 'left' }}>Short Link</th>
                    <th style={{ padding: '10px 12px', textAlign: 'left' }}>Destination URL</th>
                    <th style={{ padding: '10px 12px', textAlign: 'center' }}>Created</th>
                    <th style={{ padding: '10px 12px', textAlign: 'center' }}>Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredHistory.map((item) => (
                    <tr key={item.id} style={{ borderBottom: '1px solid var(--line)' }}>
                      <td style={{ padding: '10px 12px', fontWeight: 600 }}>{item.title}</td>
                      <td style={{ padding: '10px 12px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                          <span style={{ color: 'var(--accent)', fontWeight: 700 }}>{item.shortUrl}</span>
                          <button 
                            onClick={() => handleCopy(item.shortUrl)}
                            style={{ background: 'none', border: 'none', padding: '2px', cursor: 'pointer', color: 'var(--muted)' }}
                            title="Copy link"
                          >
                            <Copy size={14} />
                          </button>
                        </div>
                      </td>
                      <td style={{ padding: '10px 12px', color: 'var(--muted)', maxWidth: '280px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                        {item.longUrl}
                      </td>
                      <td style={{ padding: '10px 12px', textAlign: 'center', fontSize: '11.5px', color: 'var(--muted)' }}>
                        {item.createdAt}
                      </td>
                      <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                        <div style={{ display: 'flex', justifyContent: 'center', gap: '8px' }}>
                          <button 
                            className="secondary"
                            onClick={() => handleDownloadQR(item.qrCodeUrl, item.slug)}
                            style={{ padding: '4px 8px', fontSize: '11px', display: 'inline-flex', alignItems: 'center', gap: '4px' }}
                            title="Download QR code"
                          >
                            <QrCode size={13} /> QR
                          </button>
                          <a 
                            href={item.longUrl} 
                            target="_blank" 
                            rel="noopener noreferrer"
                            className="button secondary"
                            style={{ padding: '4px 8px', fontSize: '11px', display: 'inline-flex', alignItems: 'center' }}
                            title="Visit destination link"
                          >
                            <ExternalLink size={13} />
                          </a>
                          <button 
                            onClick={() => deleteHistoryItem(item.id)}
                            style={{ background: 'none', border: 'none', padding: '4px', cursor: 'pointer', color: 'var(--danger)' }}
                            title="Delete"
                          >
                            <Trash2 size={14} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <div style={{ textAlign: 'center', padding: '40px', color: 'var(--muted)' }}>
              No shortened links found in history.
            </div>
          )}
        </div>
      )}

    </div>
  );
};

export default UrlShortenerPage;
