import * as XLSX from 'xlsx';

const testData = [
  { "First Name": "John", "Last Name": "Doe", Age: 30 },
  { "First Name": "Jane", "Last Name": "Smith", Age: 25 },
  {} // Empty row
];

const worksheet = XLSX.utils.json_to_sheet(testData);
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

const data = XLSX.utils.sheet_to_json(worksheet, {defval: ""});

console.log("data", data);

let columns = Object.keys(data[0] || {}).map(c => c.trim());
console.log("columns", columns);

const records = data.map(row => {
   const newRow = {};
   for (let col of columns) {
      // Find the original key that matches the trimmed col
      const origKey = Object.keys(row).find(k => k.trim() === col) || col;
      let val = row[origKey];
      newRow[col] = (val === null || val === undefined) ? "" : String(val);
   }
   return newRow;
});

console.log("records", records);
