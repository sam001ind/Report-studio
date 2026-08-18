import * as XLSX from "xlsx";
import JSZip from "jszip";

/**
 * Parses raw ArrayBuffer / Uint8Array with multi-stage fallback and memory-efficient dense mode:
 * 1. Binary (True OpenXML .xlsx & BIFF8 .xls) with dense: true
 * 2. UTF-8 String (HTML-tables and CSVs exported as .xls/.xlsx)
 * 3. Windows-1252 String
 */
export const parseWorkbookFromBuffer = (buffer) => {
  const uint8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  let wb;
  try {
    wb = XLSX.read(uint8, { type: "array", cellDates: true, dense: true });
  } catch {
    try {
      const text = new TextDecoder("utf-8").decode(uint8);
      wb = XLSX.read(text, { type: "string", raw: true, dense: true });
    } catch {
      const text = new TextDecoder("windows-1252").decode(uint8);
      wb = XLSX.read(text, { type: "string", raw: true, dense: true });
    }
  }
  return wb;
};

/**
 * Universal Streamlined Spreadsheet & ZIP Archive Reader (Returns JSON Objects)
 * Memory-efficient: extracts and parses files one-by-one to prevent ArrayBuffer allocation limits.
 */
export const readSpreadsheetFile = async (file) => {
  let allRows = [];
  let detectedCols = [];

  if (file.name.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file);
    const validFiles = Object.keys(zip.files).filter(name =>
      !name.startsWith("__MACOSX/") &&
      !name.startsWith(".") &&
      !zip.files[name].dir &&
      (name.toLowerCase().endsWith(".xlsx") || 
       name.toLowerCase().endsWith(".xls") || 
       name.toLowerCase().endsWith(".csv") || 
       name.toLowerCase().endsWith(".xlsm"))
    );

    if (validFiles.length === 0) {
      throw new Error("No spreadsheet files (.xlsx, .xls, .csv, .xlsm) found inside the uploaded ZIP archive.");
    }

    // Process files one-by-one sequentially to release memory immediately
    for (const name of validFiles) {
      const uint8 = await zip.files[name].async("uint8array");
      const wb = parseWorkbookFromBuffer(uint8);

      if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
        const firstSheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
        if (json && json.length > 0) {
          if (detectedCols.length === 0) {
            detectedCols = Object.keys(json[0]);
          }
          allRows = allRows.concat(json);
        }
      }
    }
  } else {
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

  if (allRows.length === 0) {
    throw new Error("No data rows found in the uploaded file(s).");
  }

  return { rows: allRows, columns: detectedCols };
};

/**
 * Universal Streamlined Spreadsheet Reader returning Array-of-Arrays (AOA)
 */
export const readSpreadsheetAsAoa = async (file) => {
  let combinedAoa = [];
  let headerRow = null;

  if (file.name.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file);
    const validFiles = Object.keys(zip.files).filter(name =>
      !name.startsWith("__MACOSX/") &&
      !name.startsWith(".") &&
      !zip.files[name].dir &&
      (name.toLowerCase().endsWith(".xlsx") || 
       name.toLowerCase().endsWith(".xls") || 
       name.toLowerCase().endsWith(".csv") || 
       name.toLowerCase().endsWith(".xlsm"))
    );

    if (validFiles.length === 0) {
      throw new Error("No spreadsheet files found inside the uploaded ZIP archive.");
    }

    for (const name of validFiles) {
      const uint8 = await zip.files[name].async("uint8array");
      const wb = parseWorkbookFromBuffer(uint8);
      if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
        if (rows.length > 0) {
          if (!headerRow) {
            headerRow = rows[0];
            combinedAoa.push(headerRow);
            combinedAoa = combinedAoa.concat(rows.slice(1));
          } else {
            combinedAoa = combinedAoa.concat(rows.slice(1));
          }
        }
      }
    }
  } else {
    const buf = await file.arrayBuffer();
    const wb = parseWorkbookFromBuffer(buf);
    if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      if (rows.length > 0) {
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
  if (file.name.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file);
    const validFiles = Object.keys(zip.files).filter(name =>
      !name.startsWith("__MACOSX/") &&
      !name.startsWith(".") &&
      !zip.files[name].dir &&
      (name.toLowerCase().endsWith(".xlsx") || 
       name.toLowerCase().endsWith(".xls") || 
       name.toLowerCase().endsWith(".csv") || 
       name.toLowerCase().endsWith(".xlsm"))
    );

    if (validFiles.length === 0) {
      throw new Error("No spreadsheet files found inside the uploaded ZIP archive.");
    }

    const uint8 = await zip.files[validFiles[0]].async("uint8array");
    const wb = parseWorkbookFromBuffer(uint8);
    return { workbook: wb, sheetNames: wb.SheetNames || [] };
  }

  const buf = await file.arrayBuffer();
  const wb = parseWorkbookFromBuffer(buf);
  return { workbook: wb, sheetNames: wb.SheetNames || [] };
};
