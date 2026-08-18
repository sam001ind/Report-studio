import * as XLSX from "xlsx";
import JSZip from "jszip";

/**
 * Universal Spreadsheet & ZIP Archive Reader
 * Supports: True .xlsx (OpenXML), .xls (BIFF8), HTML-tables (.xls/.xlsx), CSV/TSV, and .zip archives
 */
export const readSpreadsheetFile = async (file) => {
  let buffers = [];

  // Handle .zip archives
  if (file.name.toLowerCase().endsWith(".zip")) {
    const zip = await JSZip.loadAsync(file);
    const validFiles = Object.keys(zip.files).filter(name =>
      !name.startsWith("__MACOSX/") &&
      !name.startsWith(".") &&
      !zip.files[name].dir &&
      (name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xls") || name.toLowerCase().endsWith(".csv"))
    );

    if (validFiles.length === 0) {
      throw new Error("No .xlsx, .xls, or .csv spreadsheets found inside the uploaded ZIP archive.");
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
    let wb;
    try {
      // 1. Binary ArrayBuffer (True OpenXML .xlsx & BIFF8 .xls)
      wb = XLSX.read(new Uint8Array(item.buffer), { type: "array", cellDates: true });
    } catch {
      try {
        // 2. UTF-8 Text Fallback (HTML tables / CSVs exported with .xls extension)
        const text = new TextDecoder("utf-8").decode(item.buffer);
        wb = XLSX.read(text, { type: "string", raw: true });
      } catch {
        // 3. Windows-1252 Fallback
        const text = new TextDecoder("windows-1252").decode(item.buffer);
        wb = XLSX.read(text, { type: "string", raw: true });
      }
    }

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
