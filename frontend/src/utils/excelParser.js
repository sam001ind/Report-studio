import * as XLSX from "xlsx";
import JSZip from "jszip";

/**
 * Checks if a zip entry path is a macOS/System artifact or directory
 */
const isIgnoredZipEntry = (name, zipEntry) => {
  if (!name || (zipEntry && zipEntry.dir)) return true;
  if (name.includes("__MACOSX") || name.includes(".DS_Store")) return true;
  const basename = name.split("/").pop();
  if (basename.startsWith("._") || basename.startsWith(".")) return true;
  return false;
};

/**
 * Checks if a file name matches known spreadsheet/data file extensions
 */
const hasSpreadsheetExtension = (name) => {
  return /\.(xlsx|xls|csv|xlsm|tsv|txt|html|htm|xml|xlsb|ods)$/i.test(name);
};

/**
 * Strips 1899 / 1900 date and GMT timezone artifacts from time strings.
 * Excel stores time-only values with base epoch Dec 30, 1899, which SheetJS
 * converts to e.g. "Sat Dec 30 1899 13:30:00 GMT+0521 (India Standard Time)".
 * This function preserves clean time text (e.g. "13:30", "13:30:00", "13:30 - 15:00", "01:30 PM").
 */
export const cleanTimeText = (val) => {
  if (val === null || val === undefined) return '';
  if (val instanceof Date) {
    if (val.getFullYear() <= 1900) {
      const hh = String(val.getHours()).padStart(2, '0');
      const mm = String(val.getMinutes()).padStart(2, '0');
      const ss = val.getSeconds() ? ':' + String(val.getSeconds()).padStart(2, '0') : '';
      return `${hh}:${mm}${ss}`;
    }
    return val.toLocaleTimeString();
  }
  let str = String(val).trim();
  // Match full JS date strings for 1899 / 1900 (e.g. "Sat Dec 30 1899 13:30:00 GMT+0521 (India Standard Time)")
  const regex = /(?:[A-Za-z]{3}\s+[A-Za-z]{3}\s+\d+\s+)?(?:1899|1900)\s+(\d{1,2}:\d{2}(?::\d{2})?)\s+GMT[^\s-]+(?:\s*\([^)]*\))?/gi;
  str = str.replace(regex, '$1');
  // Match ISO strings for 1899 / 1900
  str = str.replace(/(?:1899|1900)-12-30[T\s](\d{1,2}:\d{2}(?::\d{2})?)(?:\.\d+)?Z?/gi, '$1');
  return str.trim();
};

/**
 * Normalizes text for comparison / composite deduplication keys:
 * Strips non-breaking spaces, zero-width characters, leading/trailing dots and spaces,
 * collapses internal whitespace sequences, and converts to lowercase.
 */
export const normalizeKeyText = (str) => {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/[\u00A0\u200B\uFEFF]/g, ' ')
    .trim()
    .replace(/^[\s.]+|[\s.]+$/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
};

/**
 * Cleans user-facing text:
 * Strips non-breaking spaces, zero-width characters, and leading/trailing dots and whitespace,
 * while preserving original letter casing and internal structure.
 */
export const cleanDisplayField = (str) => {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/[\u00A0\u200B\uFEFF]/g, ' ')
    .trim()
    .replace(/^[\s.]+|[\s.]+$/g, '')
    .trim();
};

/**
 * Sanitizes worksheet cells: if a cell is an Excel time-only value (epoch year <= 1900),
 * replace it with the formatted text string (cell.w) or clean time string, keeping it as text.
 */
export const sanitizeWorksheetTimes = (sheet) => {
  if (!sheet || typeof sheet !== 'object') return sheet;

  const fixCell = (cell) => {
    if (!cell || typeof cell !== 'object') return;
    if (cell.t === 'd' || cell.v instanceof Date) {
      const d = cell.v;
      if (d instanceof Date && d.getFullYear() <= 1900) {
        let timeStr = cell.w;
        if (!timeStr) {
          const hh = String(d.getHours()).padStart(2, '0');
          const mm = String(d.getMinutes()).padStart(2, '0');
          const ss = d.getSeconds() ? ':' + String(d.getSeconds()).padStart(2, '0') : '';
          timeStr = `${hh}:${mm}${ss}`;
        }
        cell.t = 's';
        cell.v = timeStr;
        cell.w = timeStr;
      }
    }
  };

  for (const key of Object.keys(sheet)) {
    if (key.startsWith('!')) continue;
    const val = sheet[key];
    if (Array.isArray(val)) {
      for (let c = 0; c < val.length; c++) fixCell(val[c]);
    } else {
      fixCell(val);
    }
  }
  return sheet;
};

/**
 * Parses raw ArrayBuffer / Uint8Array with multi-stage fallback and memory-efficient dense mode:
 * 1. Binary (True OpenXML .xlsx & BIFF8 .xls) with dense: true
 * 2. UTF-8 String (HTML-tables, CSVs, TSVs exported as .xls/.xlsx)
 * 3. Windows-1252 String
 */
export const parseWorkbookFromBuffer = (buffer) => {
  const uint8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  let wb = null;
  try {
    wb = XLSX.read(uint8, { type: "array", cellDates: true, dense: true });
  } catch {
    try {
      const text = new TextDecoder("utf-8").decode(uint8);
      wb = XLSX.read(text, { type: "string", raw: true, dense: true });
    } catch {
      try {
        const text = new TextDecoder("windows-1252").decode(uint8);
        wb = XLSX.read(text, { type: "string", raw: true, dense: true });
      } catch {
        return null;
      }
    }
  }

  if (wb && wb.Sheets) {
    Object.keys(wb.Sheets).forEach(sheetName => {
      sanitizeWorksheetTimes(wb.Sheets[sheetName]);
    });
  }
  return wb;
};

/**
 * Universal Streamlined Spreadsheet & ZIP Archive Reader (Returns JSON Objects)
 * Auto-extracts any spreadsheets from ZIPs (including nested folders & HTML/ERP exports).
 */
export const readSpreadsheetFile = async (file) => {
  let allRows = [];
  let detectedCols = [];

  const isZip = file.name.toLowerCase().endsWith(".zip") || file.type === "application/zip" || file.type === "application/x-zip-compressed";

  if (isZip) {
    try {
      const zip = await JSZip.loadAsync(file);
      const allEntries = Object.keys(zip.files).filter(name => !isIgnoredZipEntry(name, zip.files[name]));

      // 1. Try files with spreadsheet extensions first
      let targetFiles = allEntries.filter(hasSpreadsheetExtension);

      // 2. If no extensions match, try ALL non-system files in the ZIP
      if (targetFiles.length === 0) {
        targetFiles = allEntries;
      }

      for (const name of targetFiles) {
        try {
          const uint8 = await zip.files[name].async("uint8array");
          const wb = parseWorkbookFromBuffer(uint8);

          if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
            wb.SheetNames.forEach(sheetName => {
              const sheet = wb.Sheets[sheetName];
              if (sheet) {
                const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
                if (json && json.length > 0) {
                  if (detectedCols.length === 0) {
                    detectedCols = Object.keys(json[0]);
                  }
                  allRows = allRows.concat(json);
                }
              }
            });
          }
        } catch (fileErr) {
          console.warn(`Could not parse entry ${name} inside ZIP:`, fileErr);
        }
      }
    } catch (zipErr) {
      console.warn("JSZip parse failed, falling back to direct spreadsheet parse:", zipErr);
      // Fallback: try parsing file directly if it was not a valid zip
      const buf = await file.arrayBuffer();
      const wb = parseWorkbookFromBuffer(buf);
      if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
        if (json && json.length > 0) {
          detectedCols = Object.keys(json[0]);
          allRows = json;
        }
      }
    }
  } else {
    const buf = await file.arrayBuffer();
    const wb = parseWorkbookFromBuffer(buf);

    if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
      wb.SheetNames.forEach(sheetName => {
        const sheet = wb.Sheets[sheetName];
        if (sheet) {
          const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          if (json && json.length > 0) {
            if (detectedCols.length === 0) {
              detectedCols = Object.keys(json[0]);
            }
            allRows = allRows.concat(json);
          }
        }
      });
    }
  }

  if (allRows.length === 0) {
    throw new Error("No data rows found in the uploaded file(s). Please verify that the spreadsheet or ZIP contains valid table rows.");
  }

  return { rows: allRows, columns: detectedCols };
};

/**
 * Universal Streamlined Spreadsheet Reader returning Array-of-Arrays (AOA)
 */
export const readSpreadsheetAsAoa = async (file) => {
  let combinedAoa = [];
  let headerRow = null;

  const isZip = file.name.toLowerCase().endsWith(".zip") || file.type === "application/zip" || file.type === "application/x-zip-compressed";

  if (isZip) {
    try {
      const zip = await JSZip.loadAsync(file);
      const allEntries = Object.keys(zip.files).filter(name => !isIgnoredZipEntry(name, zip.files[name]));

      let targetFiles = allEntries.filter(hasSpreadsheetExtension);
      if (targetFiles.length === 0) {
        targetFiles = allEntries;
      }

      for (const name of targetFiles) {
        try {
          const uint8 = await zip.files[name].async("uint8array");
          const wb = parseWorkbookFromBuffer(uint8);
          if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
            wb.SheetNames.forEach(sheetName => {
              const sheet = wb.Sheets[sheetName];
              if (sheet) {
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
                if (rows && rows.length > 0) {
                  if (!headerRow) {
                    headerRow = rows[0];
                    combinedAoa.push(headerRow);
                    combinedAoa = combinedAoa.concat(rows.slice(1));
                  } else {
                    combinedAoa = combinedAoa.concat(rows.slice(1));
                  }
                }
              }
            });
          }
        } catch (fileErr) {
          console.warn(`Could not parse entry ${name} inside ZIP as AOA:`, fileErr);
        }
      }
    } catch (zipErr) {
      console.warn("JSZip AOA parse failed, falling back to direct parse:", zipErr);
      const buf = await file.arrayBuffer();
      const wb = parseWorkbookFromBuffer(buf);
      if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
        if (rows && rows.length > 0) {
          combinedAoa = rows;
        }
      }
    }
  } else {
    const buf = await file.arrayBuffer();
    const wb = parseWorkbookFromBuffer(buf);
    if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      if (rows && rows.length > 0) {
        combinedAoa = rows;
      }
    }
  }

  return combinedAoa;
};

/**
 * Universal Workbook Reader (Returns Workbook instance + SheetNames)
 */
export const readSpreadsheetWorkbook = async (file) => {
  const isZip = file.name.toLowerCase().endsWith(".zip") || file.type === "application/zip" || file.type === "application/x-zip-compressed";

  if (isZip) {
    const zip = await JSZip.loadAsync(file);
    const allEntries = Object.keys(zip.files).filter(name => !isIgnoredZipEntry(name, zip.files[name]));

    let targetFiles = allEntries.filter(hasSpreadsheetExtension);
    if (targetFiles.length === 0) {
      targetFiles = allEntries;
    }

    if (targetFiles.length === 0) {
      throw new Error("No readable files found inside the uploaded ZIP archive.");
    }

    for (const name of targetFiles) {
      try {
        const uint8 = await zip.files[name].async("uint8array");
        const wb = parseWorkbookFromBuffer(uint8);
        if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
          return { workbook: wb, sheetNames: wb.SheetNames };
        }
      } catch (err) {
        console.warn(`Failed reading sheet from ${name}:`, err);
      }
    }
  }

  const buf = await file.arrayBuffer();
  const wb = parseWorkbookFromBuffer(buf);
  return { workbook: wb, sheetNames: wb.SheetNames || [] };
};
