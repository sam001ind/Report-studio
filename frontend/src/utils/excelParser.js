import * as XLSX from "xlsx";
import JSZip from "jszip";

/**
 * Parses raw ArrayBuffer with multi-stage fallback:
 * 1. Binary ArrayBuffer (True OpenXML .xlsx & BIFF8 .xls)
 * 2. UTF-8 String (HTML-tables and CSVs exported as .xls/.xlsx)
 * 3. Windows-1252 String
 */
export const parseWorkbookFromBuffer = (buffer) => {
  let wb;
  try {
    wb = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: true, cellStyles: true });
  } catch {
    try {
      const text = new TextDecoder("utf-8").decode(buffer);
      wb = XLSX.read(text, { type: "string", raw: true });
    } catch {
      const text = new TextDecoder("windows-1252").decode(buffer);
      wb = XLSX.read(text, { type: "string", raw: true });
    }
  }
  return wb;
};

/**
 * Universal Spreadsheet & ZIP Archive Reader (Returns JSON Objects)
 */
export const readSpreadsheetFile = async (file) => {
  let buffers = [];

  if (file.name.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file);
    const validFiles = Object.keys(zip.files).filter(name =>
      !name.startsWith("__MACOSX/") &&
      !name.startsWith(".") &&
      !zip.files[name].dir &&
      (name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".csv") || name.toLowerCase().endsWith(".xlsm"))
    );

    if (validFiles.length === 0) {
      throw new Error("No spreadsheet files (.xlsx, .xls, .csv, .xlsm) found inside the uploaded ZIP archive.");
    }

    for (const name of validFiles) {
      const buf = await zip.files[name].async("arraybuffer");
      buffers.push({ name, buffer: buf });
    }
  } else {
    const buf = await file.arrayBuffer();
    buffers.push({ name: file.name, buffer: buf });
  }

  let allRows = [];
  let detectedCols = [];

  for (const item of buffers) {
    const wb = parseWorkbookFromBuffer(item.buffer);

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

  if (allRows.length === 0) {
    throw new Error("No data rows found in the uploaded file(s).");
  }

  return { rows: allRows, columns: detectedCols };
};

/**
 * Universal Spreadsheet Reader returning Array-of-Arrays (AOA)
 */
export const readSpreadsheetAsAoa = async (file) => {
  let buffers = [];

  if (file.name.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file);
    const validFiles = Object.keys(zip.files).filter(name =>
      !name.startsWith("__MACOSX/") &&
      !name.startsWith(".") &&
      !zip.files[name].dir &&
      (name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".csv") || name.toLowerCase().endsWith(".xlsm"))
    );

    if (validFiles.length === 0) {
      throw new Error("No spreadsheet files found inside the uploaded ZIP archive.");
    }

    for (const name of validFiles) {
      const buf = await zip.files[name].async("arraybuffer");
      buffers.push({ name, buffer: buf });
    }
  } else {
    const buf = await file.arrayBuffer();
    buffers.push({ name: file.name, buffer: buf });
  }

  let combinedAoa = [];
  let headerRow = null;

  for (const item of buffers) {
    const wb = parseWorkbookFromBuffer(item.buffer);
    if (wb && wb.SheetNames && wb.SheetNames.length > 0) {
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      if (rows.length > 0) {
        if (!headerRow) {
          headerRow = rows[0];
          combinedAoa.push(headerRow);
          combinedAoa = combinedAoa.concat(rows.slice(1));
        } else {
          // Append subsequent files' data rows
          combinedAoa = combinedAoa.concat(rows.slice(1));
        }
      }
    }
  }

  return combinedAoa;
};

/**
 * Universal Workbook Reader (Returns Workbook instance + SheetNames)
 */
export const readSpreadsheetWorkbook = async (file) => {
  const buf = await file.arrayBuffer();
  const wb = parseWorkbookFromBuffer(buf);
  return { workbook: wb, sheetNames: wb.SheetNames || [] };
};
