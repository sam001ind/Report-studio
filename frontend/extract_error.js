import fs from 'fs';

const source = fs.readFileSync('dist/assets/index-03htjlVQ.js', 'utf8');
const lines = source.split('\n');
const line80 = lines[79]; // 0-indexed

// Extract surrounding code (col 4134 - 1 is index 4133)
const start = Math.max(0, 4133 - 100);
const end = Math.min(line80.length, 4133 + 100);

console.log("SURROUNDING CODE:");
console.log(line80.substring(start, end));
console.log("\nAT EXACT COLUMN:");
console.log("--> " + line80.substring(4133, 4133 + 50));
