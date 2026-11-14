// Date Parsing and Formatting Utilities for Report Studio & Admission Engines

export const cleanEscapeArtifacts = (val) => {
  if (val === null || val === undefined) return '';
  const str = String(val).trim();
  if (str === '\\N' || str.toLowerCase() === 'null' || str.toLowerCase() === 'undefined' || str.toLowerCase() === 'nan') {
    return '';
  }
  return str;
};

// Formats date strictly to YYYY-MM-DD (removes any timestamp, preserves exact local day)
export const parseDobYMD = (val) => {
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  const clean = cleanEscapeArtifacts(val);
  if (!clean) return '';

  // 1. If it starts with YYYY-MM-DD (e.g. "2002-01-17", "2008-08-19", "2008-08-19 00:00:00", "2008-08-19T00:00:00.000Z")
  const ymdPrefix = clean.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (ymdPrefix) {
    const y = ymdPrefix[1];
    const m = String(parseInt(ymdPrefix[2], 10)).padStart(2, '0');
    const d = String(parseInt(ymdPrefix[3], 10)).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // 2. Excel serial number (e.g. 37273, 39679)
  if (/^\d{4,5}$/.test(clean)) {
    const serial = parseInt(clean, 10);
    if (serial > 10000 && serial < 60000) {
      const epoch = Date.UTC(1899, 11, 30);
      const d = new Date(epoch + serial * 86400 * 1000);
      if (!isNaN(d.getTime())) {
        const y = d.getUTCFullYear();
        const m = String(d.getUTCMonth() + 1).padStart(2, '0');
        const day = String(d.getUTCDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
      }
    }
  }

  // 3. DD-MM-YYYY or DD/MM/YYYY (e.g. "17-01-2002", "19/08/2008", "19-08-2008 12:00:00")
  const dmyPrefix = clean.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dmyPrefix) {
    const d = String(parseInt(dmyPrefix[1], 10)).padStart(2, '0');
    const m = String(parseInt(dmyPrefix[2], 10)).padStart(2, '0');
    const y = dmyPrefix[3];
    return `${y}-${m}-${d}`;
  }

  // 4. YY-MM-DD (e.g. "08-12-22" from dump)
  const yyMatch = clean.match(/^(\d{2})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (yyMatch) {
    let year = parseInt(yyMatch[1], 10);
    year = year > 50 ? 1900 + year : 2000 + year;
    const m = String(parseInt(yyMatch[2], 10)).padStart(2, '0');
    const d = String(parseInt(yyMatch[3], 10)).padStart(2, '0');
    return `${year}-${m}-${d}`;
  }

  // 5. Standard JS Date fallback (Use local getters to avoid UTC timezone back-shifting)
  const d = new Date(clean);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  return clean.split('T')[0].split(' ')[0];
};

// Converts YYYY-MM-DD date strictly to Excel integer day serial (no time fraction)
export const dateToExcelSerial = (dateStr) => {
  if (!dateStr) return null;
  const parts = String(dateStr).split('T')[0].split(' ')[0].split(/[-/.]/);
  if (parts.length !== 3) return null;
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const day = parseInt(parts[2], 10);
  if (isNaN(year) || isNaN(month) || isNaN(day)) return null;
  const utcDate = Date.UTC(year, month, day);
  const epoch = Date.UTC(1899, 11, 30);
  return Math.round((utcDate - epoch) / (24 * 60 * 60 * 1000));
};

// Formats date to DD/MM/YYYY
export const parseDobDMY = (val) => {
  const ymd = parseDobYMD(val);
  if (!ymd) return '';
  const parts = ymd.split('-');
  if (parts.length === 3) {
    return `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  return ymd;
};
