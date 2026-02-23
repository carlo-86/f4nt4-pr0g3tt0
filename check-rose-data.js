const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function checkRose(filePath, label) {
  console.log(`\n=== ${label} ===`);
  const wb = XLSX.readFile(filePath, { password: PW, cellFormula: true });
  const ws = wb.Sheets['TutteLeRose'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  console.log(`Rows: ${data.length}`);

  // Print headers (row 0 and 1)
  for (let r = 0; r < 4; r++) {
    const row = data[r];
    if (row) {
      const nonEmpty = [];
      for (let c = 0; c < Math.min(row.length, 30); c++) {
        if (row[c] !== '') nonEmpty.push(`${c}="${row[c]}"`);
      }
      console.log(`Row ${r}: ${nonEmpty.join(', ')}`);
    }
  }

  // Print a sample of data rows to understand structure
  console.log('\nSample data rows:');
  for (let r = 4; r < Math.min(data.length, 12); r++) {
    const row = data[r];
    if (row) {
      const nonEmpty = [];
      for (let c = 0; c < Math.min(row.length, 20); c++) {
        if (row[c] !== '') nonEmpty.push(`${c}="${row[c]}"`);
      }
      console.log(`Row ${r}: ${nonEmpty.join(', ')}`);
    }
  }

  // Also check for specific missing players
  const searchNames = ['CIRCATI', 'BERISHA', 'BELGHALI', 'STREFEZZA', 'CELIK', 'RASPADORI', 'RATKOV'];
  console.log('\nSearching for missing players:');
  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    if (!row) continue;
    for (let c = 0; c < row.length; c++) {
      const val = String(row[c] || '').toUpperCase();
      for (const name of searchNames) {
        if (val.includes(name)) {
          // Print surrounding columns
          const ctx = [];
          for (let cc = Math.max(0, c-2); cc < Math.min(row.length, c+10); cc++) {
            if (row[cc] !== '') ctx.push(`${cc}="${row[cc]}"`);
          }
          console.log(`  Found ${name} at row ${r} col ${c}: ${ctx.join(', ')}`);
        }
      }
    }
  }
}

checkRose(
  path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
  'FT Rose'
);

checkRose(
  path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
  'FM Rose'
);
