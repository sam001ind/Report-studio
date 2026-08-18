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
 * Parses raw ArrayBuffer / Uint8Array with multi-stage fallback and memory-efficient dense mode:
 * 1. Binary (True OpenXML .xlsx & BIFF8 .xls) with dense: true
 * 2. UTF-8 String (HTML-tables, CSVs, TSVs exported as .xls/.xlsx)
 * 3. Windows-1252 String
 */
export const parseWorkbookFromBuffer = (buffer) => {
  const uint8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  try {
    return XLSX.read(uint8, { type: "array", cellDates: true, dense: true });
  } catch {
    try {
      const text = new TextDecoder("utf-8").decode(uint8);
      return XLSX.read(text, { type: "string", raw: true, dense: true });
    } catch {
      try {
        const text = new TextDecoder("windows-1252").decode(uint8);
        return XLSX.read(text, { type: "string", raw: true, dense: true });
      } catch {
        return null;
      }
    }
  }
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
