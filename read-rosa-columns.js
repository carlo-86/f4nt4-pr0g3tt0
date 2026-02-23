const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function readRosaColumns(filePath, league) {
  const wb = XLSX.readFile(filePath, { password: PW, cellFormula: true });
  const ws = wb.Sheets['ROSA'];

  console.log(`\n=== ${league} - ROSA Columns K-O ===\n`);

  // Print headers in row 3 (index 3) for columns J-T
  console.log('Headers (row 4):');
  for (let c = 9; c <= 19; c++) {
    const addr = XLSX.utils.encode_cell({ r: 3, c });
    const cell = ws[addr];
    const colLetter = XLSX.utils.encode_col(c);
    console.log(`  ${colLetter}4: ${cell ? JSON.stringify(cell.v) : '(empty)'}${cell?.f ? ' f=' + cell.f : ''}`);
  }

  // Print first 6 player rows (P section: rows 5-10) showing K, L, M, N, O values and formulas
  console.log('\nPlayer data (P section, rows 5-10):');
  for (let r = 4; r <= 9; r++) {
    const parts = [];
    for (let c = 9; c <= 19; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const colLetter = XLSX.utils.encode_col(c);
      if (cell) {
        let info = `${colLetter}=${cell.v !== undefined ? cell.v : ''}`;
        if (cell.f) info += ` [f:${cell.f.substring(0, 60)}]`;
        parts.push(info);
      }
    }
    console.log(`  Row ${r + 1}: ${parts.join(' | ')}`);
  }

  // Print D section (rows 12-16)
  console.log('\nPlayer data (D section, rows 12-16):');
  for (let r = 11; r <= 15; r++) {
    const parts = [];
    for (let c = 9; c <= 19; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const colLetter = XLSX.utils.encode_col(c);
      if (cell) {
        let info = `${colLetter}=${cell.v !== undefined ? cell.v : ''}`;
        if (cell.f) info += ` [f:${cell.f.substring(0, 60)}]`;
        parts.push(info);
      }
    }
    console.log(`  Row ${r + 1}: ${parts.join(' | ')}`);
  }

  // Print what column K contains (player name)
  console.log('\nColumn K (all non-empty):');
  for (let r = 4; r <= 49; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: 10 });
    const cell = ws[addr];
    if (cell && cell.v) {
      const eAddr = XLSX.utils.encode_cell({ r, c: 4 });
      const eCell = ws[eAddr];
      const mAddr = XLSX.utils.encode_cell({ r, c: 12 });
      const mCell = ws[mAddr];
      const nAddr = XLSX.utils.encode_cell({ r, c: 13 });
      const nCell = ws[nAddr];
      console.log(`  Row ${r + 1}: K=${cell.v} | M=${mCell?.v || ''} | N=${nCell?.v || ''} | E(cost)=${eCell?.v || ''}`);
    }
  }
}

readRosaColumns(
  path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'),
  'FT'
);
readRosaColumns(
  path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
  'FM'
);
