import fs from 'fs';

const source = fs.readFileSync('dist/assets/index-03htjlVQ.js', 'utf8');
const lines = source.split('\n');
const line80 = lines[79];

// Extract code from 4134 up to 500 characters
console.log(line80.substring(4134, 4134 + 500));
