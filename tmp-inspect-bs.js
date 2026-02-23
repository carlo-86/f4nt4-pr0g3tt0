const XLSX = require('xlsx');
const dbPath = 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/Fanta Tosti 2026/DB Excel/Fanta Tosti 2026 - DB completo (06.02.2026).xlsx';
const wb = XLSX.readFile(dbPath, { password: "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99", cellFormula: true });
const wsDB = wb.Sheets['DB'];

function getF(sheet, col, row) {
  const addr = col + row;
  const cell = sheet[addr];
  if (!cell) return '(empty)';
  if (cell.f) return '=' + cell.f;
  return 'VALUE: ' + cell.v;
}

// Find Portiere and non-Portiere rows (Ruolo is column A)
let portiereRows = [];
let attaccanteRows = [];
let centrocampistaRows = [];
let difensoreRows = [];
for (let r = 3; r <= 700; r++) {
  const roleCell = wsDB['A' + r];
  if (!roleCell) continue;
  const role = String(roleCell.v).trim();
  if (role === 'P' && portiereRows.length < 3) portiereRows.push(r);
  if (role === 'A' && attaccanteRows.length < 2) attaccanteRows.push(r);
  if (role === 'C' && centrocampistaRows.length < 2) centrocampistaRows.push(r);
  if (role === 'D' && difensoreRows.length < 2) difensoreRows.push(r);
}

console.log('Portiere rows:', portiereRows);
console.log('Attaccante rows:', attaccanteRows);
console.log('Centrocampista rows:', centrocampistaRows);
console.log('Difensore rows:', difensoreRows);
console.log('');

console.log('=== BS (NEGATIVA 1° anno) ===');
console.log('--- Portieri ---');
for (const r of portiereRows) {
  const name = wsDB['C' + r] ? wsDB['C' + r].v : '?';
  console.log(`  BS${r} (P, ${name}):`);
  console.log(`    ${getF(wsDB, 'BS', r)}`);
}
console.log('--- Attaccanti ---');
for (const r of attaccanteRows) {
  const name = wsDB['C' + r] ? wsDB['C' + r].v : '?';
  console.log(`  BS${r} (A, ${name}):`);
  console.log(`    ${getF(wsDB, 'BS', r)}`);
}
console.log('--- Centrocampisti ---');
for (const r of centrocampistaRows) {
  const name = wsDB['C' + r] ? wsDB['C' + r].v : '?';
  console.log(`  BS${r} (C, ${name}):`);
  console.log(`    ${getF(wsDB, 'BS', r)}`);
}
console.log('--- Difensori ---');
for (const r of difensoreRows) {
  const name = wsDB['C' + r] ? wsDB['C' + r].v : '?';
  console.log(`  BS${r} (D, ${name}):`);
  console.log(`    ${getF(wsDB, 'BS', r)}`);
}

console.log('');
console.log('=== BT (NEGATIVA 2° anno) for comparison ===');
for (const r of portiereRows.slice(0, 2).concat(attaccanteRows.slice(0, 1)).concat(difensoreRows.slice(0, 1))) {
  const name = wsDB['C' + r] ? wsDB['C' + r].v : '?';
  const role = wsDB['A' + r] ? wsDB['A' + r].v : '?';
  console.log(`  BT${r} (${role}, ${name}):`);
  console.log(`    ${getF(wsDB, 'BT', r)}`);
}

// Also get the full BY formula for row 3 (correct) and row 4 (wrong)
console.log('');
console.log('=== BY full formulas (row 3=correct, row 4=wrong) ===');
console.log(`  BY3: ${getF(wsDB, 'BY', 3)}`);
console.log('');
console.log(`  BY4: ${getF(wsDB, 'BY', 4)}`);

// Count how many rows have data
let lastRow = 0;
for (let r = 3; r <= 1000; r++) {
  if (wsDB['A' + r] && wsDB['A' + r].v) lastRow = r;
}
console.log('');
console.log('Last data row in DB:', lastRow);
