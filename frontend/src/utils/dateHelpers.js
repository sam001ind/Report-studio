// Date Parsing and Formatting Utilities for Report Studio & Admission Engines

export const cleanEscapeArtifacts = (val) => {
  if (val === null || val === undefined) return '';
  const str = String(val).trim();
  if (str === '\\N' || str.toLowerCase() === 'null' || str.toLowerCase() === 'undefined' || str.toLowerCase() === 'nan') {
    return '';
  }
  return str;
};

// Formats date strictly to YYYY-MM-DD (removes any timestamp, preserves exact calendar day across all timezones)
export const parseDobYMD = (val) => {
  if (val === null || val === undefined) return '';

  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    const t = val.getTime();
    const utcHours = val.getUTCHours();
    let mid;
    if (utcHours >= 18) {
      // In positive timezones like IST (+5:30), evening UTC (18:00 - 23:59) is midnight of target day
      mid = new Date(t + 7 * 3600 * 1000);
    } else if (val.getHours() === 23 && val.getMinutes() >= 50) {
      // Historical V8 8m 50s offset artifact where midnight became 23:59:50 local
      mid = new Date(t + 20 * 60 * 1000);
    } else {
      mid = val;
    }
    const y = mid.getFullYear();
    const m = String(mid.getMonth() + 1).padStart(2, '0');
    const d = String(mid.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  const clean = cleanEscapeArtifacts(val);
  if (!clean) return '';

  // 1. If it starts with YYYY-MM-DD (e.g. "2002-01-17", "2009-03-17", "2008-08-19 00:00:00", "2008-08-19T00:00:00.000Z")
  const ymdPrefix = clean.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (ymdPrefix) {
    const y = ymdPrefix[1];
    const m = String(parseInt(ymdPrefix[2], 10)).padStart(2, '0');
    const d = String(parseInt(ymdPrefix[3], 10)).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // 2. Excel serial number (e.g. 37273 for 2002-01-17, 39889 for 2009-03-17)
  if (/^\d{4,5}(\.\d+)?$/.test(clean)) {
    const serial = Math.floor(parseFloat(clean));
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

  // 3. Named month formats (e.g. "17-Jan-2002", "17 Jan 2002", "17/January/2002", "Jan 17, 2002", "17-March-2009")
  const MONTHS = {
    jan: '01', january: '01',
    feb: '02', february: '02',
    mar: '03', march: '03',
    apr: '04', april: '04',
    may: '05',
    jun: '06', june: '06',
    jul: '07', july: '07',
    aug: '08', august: '08',
    sep: '09', sept: '09', september: '09',
    oct: '10', october: '10',
    nov: '11', november: '11',
    dec: '12', december: '12'
  };
  const namedDmy = clean.match(/^(\d{1,2})[-/.\s]+([A-Za-z]+)[-/.\s]+(\d{2,4})/);
  if (namedDmy) {
    const d = String(parseInt(namedDmy[1], 10)).padStart(2, '0');
    const monKey = namedDmy[2].toLowerCase();
    const m = MONTHS[monKey] || MONTHS[monKey.slice(0, 3)];
    let y = parseInt(namedDmy[3], 10);
    if (y < 100) y = y > 50 ? 1900 + y : 2000 + y;
    if (m) return `${y}-${m}-${d}`;
  }
  const namedMdy = clean.match(/^([A-Za-z]+)[-/.\s]+(\d{1,2})[,]?[-/.\s]+(\d{2,4})/);
  if (namedMdy) {
    const monKey = namedMdy[1].toLowerCase();
    const m = MONTHS[monKey] || MONTHS[monKey.slice(0, 3)];
    const d = String(parseInt(namedMdy[2], 10)).padStart(2, '0');
    let y = parseInt(namedMdy[3], 10);
    if (y < 100) y = y > 50 ? 1900 + y : 2000 + y;
    if (m) return `${y}-${m}-${d}`;
  }

  // 4. DD-MM-YYYY or MM/DD/YYYY (4-digit year, handles both Indian dd/mm/yyyy and m/d/yyyy)
  const dmyPrefix = clean.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dmyPrefix) {
    const first = parseInt(dmyPrefix[1], 10);
    const second = parseInt(dmyPrefix[2], 10);
    const y = dmyPrefix[3];
    if (second > 12) {
      // Second number cannot be month (e.g. 1/17/2002 -> Month 1, Day 17, Year 2002)
      const m = String(first).padStart(2, '0');
      const d = String(second).padStart(2, '0');
      return `${y}-${m}-${d}`;
    } else {
      // Default to standard Indian dd/mm/yyyy format
      const d = String(first).padStart(2, '0');
      const m = String(second).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
  }

  // 5. YY-MM-DD with hyphens (e.g. "02-01-17", "09-03-17", "08-12-22", or DD-MM-YY "17-01-02", "16-01-02")
  const yyHyphen = clean.match(/^(\d{2})-(\d{1,2})-(\d{1,2})$/);
  if (yyHyphen) {
    const first = parseInt(yyHyphen[1], 10);
    const second = parseInt(yyHyphen[2], 10);
    const third = parseInt(yyHyphen[3], 10);
    if (second <= 12) {
      if (first > 12) {
        // First is definitely Day (e.g. 17-01-02 -> Day 17, Month 01, Year 2002)
        const year = third > 50 ? 1900 + third : 2000 + third;
        return `${year}-${String(second).padStart(2, '0')}-${String(first).padStart(2, '0')}`;
      } else {
        // Default to YY-MM-DD (e.g. 02-01-17 -> Year 2002, Month 01, Day 17; 09-03-17 -> Year 2009, Month 03, Day 17; 08-12-22 -> Year 2008, Month 12, Day 22)
        const year = first > 50 ? 1900 + first : 2000 + first;
        return `${year}-${String(second).padStart(2, '0')}-${String(third).padStart(2, '0')}`;
      }
    }
  }

  // 6. DD/MM/YY or MM/DD/YY (2-digit year at end, handles both Indian dd/mm/yy and Excel m/d/yy like 1/17/02)
  const dmyShort = clean.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2})$/);
  if (dmyShort) {
    const first = parseInt(dmyShort[1], 10);
    const second = parseInt(dmyShort[2], 10);
    let third = parseInt(dmyShort[3], 10);
    third = third > 50 ? 1900 + third : 2000 + third;
    if (second > 12) {
      // Second number cannot be month (e.g. 1/17/02 -> Month 1, Day 17, Year 2002)
      const m = String(first).padStart(2, '0');
      const d = String(second).padStart(2, '0');
      return `${third}-${m}-${d}`;
    } else if (first > 12) {
      // First number cannot be month (e.g. 17/01/02 -> Day 17, Month 1, Year 2002)
      const d = String(first).padStart(2, '0');
      const m = String(second).padStart(2, '0');
      return `${third}-${m}-${d}`;
    } else {
      // Default to DD-MM-YY standard Indian format (e.g. 05/01/02 -> Day 5, Month 1, Year 2002)
      const d = String(first).padStart(2, '0');
      const m = String(second).padStart(2, '0');
      return `${third}-${m}-${d}`;
    }
  }

  // 7. Standard JS Date fallback (Use shifted mid-day UTC/local to avoid any timezone edge shift)
  const d = new Date(clean);
  if (!isNaN(d.getTime())) {
    const t = d.getTime();
    const utcHours = d.getUTCHours();
    let mid;
    if (utcHours >= 18) {
      mid = new Date(t + 7 * 3600 * 1000);
    } else if (d.getHours() === 23 && d.getMinutes() >= 50) {
      mid = new Date(t + 20 * 60 * 1000);
    } else {
      mid = d;
    }
    const y = mid.getFullYear();
    const m = String(mid.getMonth() + 1).padStart(2, '0');
    const day = String(mid.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  return clean.split('T')[0].split(' ')[0];
};

// Converts YYYY-MM-DD or YY-MM-DD date strictly to Excel integer day serial (no time fraction)
export const dateToExcelSerial = (dateStr) => {
  if (!dateStr) return null;
  const parts = String(dateStr).split('T')[0].split(' ')[0].split(/[-/.]/);
  if (parts.length !== 3) return null;
  let year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const day = parseInt(parts[2], 10);
  if (isNaN(year) || isNaN(month) || isNaN(day)) return null;
  if (year < 100) year = year > 50 ? 1900 + year : 2000 + year;
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

// Formats date to YY-MM-DD (e.g. 2002-01-17 -> 02-01-17)
export const formatDobYYMMDD = (val) => {
  if (!val && val !== 0) return '';
  const clean = String(val).trim();
  if (/^\d{2}-\d{2}-\d{2}$/.test(clean)) return clean;
  const ymd = parseDobYMD(clean);
  if (ymd) {
    const parts = ymd.split('-');
    if (parts.length === 3) {
      return `${parts[0].slice(-2)}-${parts[1]}-${parts[2]}`;
    }
  }
  return clean;
};
