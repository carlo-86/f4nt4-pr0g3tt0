const XLSX = require('xlsx');
const path = require('path');

const files = [
  { label: 'FT', path: 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/Fanta Tosti 2026/DB Excel/Fanta Tosti 2026 - DB completo (06.02.2026).xlsx' },
  { label: 'FM', path: 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/FantaMantra Manageriale/DB Excel/FantaMantra Manageriale - DB completo (06.02.2026).xlsx' }
];

function cellVal(ws, r, c) {
  const addr = XLSX.utils.encode_cell({r, c});
  const cell = ws[addr];
  if (!cell) return '';
  return cell.v !== undefined ? cell.v : '';
}

function cellFormula(ws, r, c) {
  const addr = XLSX.utils.encode_cell({r, c});
  const cell = ws[addr];
  if (!cell) return '';
  if (cell.f) return '=' + cell.f;
  return String(cell.v !== undefined ? cell.v : '');
}

function cellRaw(ws, r, c) {
  const addr = XLSX.utils.encode_cell({r, c});
  const cell = ws[addr];
  if (!cell) return null;
  return { v: cell.v, t: cell.t, f: cell.f, w: cell.w };
}

for (const file of files) {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`${file.label}: ${path.basename(file.path)}`);
  console.log('='.repeat(60));

  const wb = XLSX.read(require('fs').readFileSync(file.path), { cellFormula: true, cellDates: true });

  // 1. List all sheets
  console.log('\n--- SHEET NAMES ---');
  wb.SheetNames.forEach((name, i) => console.log(`  ${i}: ${name}`));

  // 2. ROSA sheet
  const wsRosa = wb.Sheets['ROSA'];
  if (wsRosa) {
    const range = XLSX.utils.decode_range(wsRosa['!ref']);
    console.log(`\n--- ROSA sheet (range: ${wsRosa['!ref']}) ---`);
    console.log(`Rows: ${range.s.r}-${range.e.r}, Cols: ${range.s.c}-${range.e.c}`);

    // Headers row 0 and row 1
    console.log('\nROSA Headers (row 1 = index 0):');
    for (let c = 0; c <= Math.min(range.e.c, 35); c++) {
      const v = cellVal(wsRosa, 0, c);
      if (v !== '') console.log(`  Col ${c} (${XLSX.utils.encode_col(c)}): ${v}`);
    }
    console.log('\nROSA Headers (row 2 = index 1):');
    for (let c = 0; c <= Math.min(range.e.c, 35); c++) {
      const v = cellVal(wsRosa, 1, c);
      if (v !== '') console.log(`  Col ${c} (${XLSX.utils.encode_col(c)}): ${v}`);
    }

    // Show ALL columns from col 10+ with headers and formulas
    console.log('\nROSA Late columns (col 10+) - headers + first data row formulas:');
    for (let c = 10; c <= range.e.c; c++) {
      const h0 = cellVal(wsRosa, 0, c);
      const h1 = cellVal(wsRosa, 1, c);
      const f2 = cellFormula(wsRosa, 2, c);
      const f3 = cellFormula(wsRosa, 3, c);
      if (h0 !== '' || h1 !== '' || f2 !== '' || f3 !== '') {
        console.log(`  Col ${c} (${XLSX.utils.encode_col(c)}): h0="${h0}" h1="${h1}" r3="${f2}" r4="${f3}"`);
      }
    }

    // Print a few data rows (rows 3-7 = index 2-6)
    console.log('\nROSA Data rows 3-7 (first 20 cols):');
    for (let r = 2; r <= 6; r++) {
      const row = [];
      for (let c = 0; c < 20; c++) {
        row.push(String(cellVal(wsRosa, r, c)).substring(0, 15));
      }
      console.log(`  Row ${r+1}: ${row.join(' | ')}`);
    }
  } else {
    console.log('\nROSA sheet NOT FOUND');
  }

  // 3. QUOTE / MONTEPREMI sheets
  for (const sheetName of wb.SheetNames) {
    if (sheetName.toUpperCase().includes('QUOT') || sheetName.toUpperCase().includes('MONTE') || sheetName.toUpperCase().includes('PREMI')) {
      const ws = wb.Sheets[sheetName];
      const range = XLSX.utils.decode_range(ws['!ref']);
      console.log(`\n--- Sheet "${sheetName}" (range: ${ws['!ref']}) ---`);

      const maxR = Math.min(range.e.r, 60);
      const maxC = Math.min(range.e.c, 30);

      for (let r = 0; r <= maxR; r++) {
        const cells = [];
        let hasContent = false;
        for (let c = 0; c <= maxC; c++) {
          const raw = cellRaw(ws, r, c);
          if (raw) {
            hasContent = true;
            const fStr = raw.f ? `[F:=${raw.f}]` : '';
            const vStr = raw.w || String(raw.v || '');
            cells.push(`${XLSX.utils.encode_col(c)}:${vStr.substring(0,25)}${fStr ? ' '+fStr.substring(0,60) : ''}`);
          }
        }
        if (hasContent) {
          console.log(`  Row ${r+1}: ${cells.join(' | ')}`);
        }
      }
    }
  }
}
