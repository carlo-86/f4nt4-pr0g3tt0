const XLSX = require('xlsx');

// Read the FT DB file (both DBs have identical formula issues)
const dbPath = 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/Fanta Tosti 2026/DB Excel/Fanta Tosti 2026 - DB completo (06.02.2026).xlsx';

const wb = XLSX.readFile(dbPath, { password: "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99", cellFormula: true });
const wsDB = wb.Sheets['DB'];

// Helper: get formula or value from cell
function getF(sheet, col, row) {
  const addr = col + row;
  const cell = sheet[addr];
  if (!cell) return '(empty)';
  if (cell.f) return '=' + cell.f;
  return 'VALUE: ' + cell.v;
}

console.log('=== BUG 1: BV/BW/BX (INVARIATA) — rows 3-12 vs 13-15 ===');
console.log('');
for (const col of ['BV', 'BW', 'BX']) {
  console.log(`--- ${col} ---`);
  // Sample old (rows 3-5) and correct (rows 13-15)
  for (const r of [3, 4, 5, 12, 13, 14, 15]) {
    console.log(`  ${col}${r}: ${getF(wsDB, col, r)}`);
  }
  console.log('');
}

console.log('=== BUG 2: BY/BZ/CA (POSITIVA) — check formula variants ===');
console.log('');
for (const col of ['BY', 'BZ', 'CA']) {
  console.log(`--- ${col} ---`);
  // Check a few rows to find both variants
  for (const r of [3, 4, 5, 13, 14, 50, 100, 200, 300, 400, 500, 600, 697]) {
    const f = getF(wsDB, col, r);
    // Show only if different patterns
    const hasBF = f.includes('BF') || f.includes('BG') || f.includes('BH') || f.includes('BI');
    const hasBB = f.includes('BB') || f.includes('BC') || f.includes('BD') || f.includes('BE');
    const tag = hasBF ? '[CORRECT-BF]' : hasBB ? '[OLD-BB]' : '[OTHER]';
    console.log(`  ${col}${r} ${tag}: ${f.substring(0, 120)}${f.length > 120 ? '...' : ''}`);
  }
  console.log('');
}

console.log('=== BUG 3: BS (NEGATIVA 1° anno) — Portiere vs others ===');
console.log('');
// Need to find a Portiere row. Check column C (Ruolo)
console.log('--- Finding Portiere rows ---');
let portiereRows = [];
let nonPortiereRows = [];
for (let r = 3; r <= 700; r++) {
  const roleCell = wsDB['C' + r];
  if (!roleCell) continue;
  const role = String(roleCell.v).trim();
  if (role === 'P' && portiereRows.length < 3) portiereRows.push(r);
  if (role === 'A' && nonPortiereRows.length < 3) nonPortiereRows.push(r);
}
console.log('Portiere rows found:', portiereRows);
console.log('Attaccante rows found:', nonPortiereRows);
console.log('');

console.log('--- BS formulas for Portieri ---');
for (const r of portiereRows) {
  const name = wsDB['B' + r] ? wsDB['B' + r].v : '?';
  console.log(`  BS${r} (${name}): ${getF(wsDB, 'BS', r)}`);
}
console.log('');
console.log('--- BS formulas for Attaccanti ---');
for (const r of nonPortiereRows) {
  const name = wsDB['B' + r] ? wsDB['B' + r].v : '?';
  console.log(`  BS${r} (${name}): ${getF(wsDB, 'BS', r)}`);
}

console.log('');
console.log('--- BT (2° anno) for comparison ---');
for (const r of portiereRows.concat(nonPortiereRows.slice(0, 1))) {
  const name = wsDB['B' + r] ? wsDB['B' + r].v : '?';
  console.log(`  BT${r} (${name}): ${getF(wsDB, 'BT', r)}`);
}

console.log('');
console.log('=== BN column — sample values ===');
for (const r of [3, 4, 5, 100]) {
  console.log(`  BN${r}: ${getF(wsDB, 'BN', r)}`);
  console.log(`  AO${r}: ${getF(wsDB, 'AO', r)}`);
}

console.log('');
console.log('=== BP header ===');
console.log(`  BP1: ${getF(wsDB, 'BP', 1)}`);
console.log(`  BP2: ${getF(wsDB, 'BP', 2)}`);

// Also check: which column is "Ruolo" to identify Portieri in VBA
console.log('');
console.log('=== Column headers (row 2) ===');
for (const col of ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']) {
  console.log(`  ${col}2: ${getF(wsDB, col, 2)}`);
}
