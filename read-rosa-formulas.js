const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function analyzeRosa(filePath, league) {
  // Read with cellFormula option to get formulas
  const wb = XLSX.readFile(filePath, { password: PW, cellFormula: true, cellStyles: true });
  const ws = wb.Sheets['ROSA'];

  console.log(`\n=== ${league} - ROSA Sheet Analysis ===\n`);

  // Get the range
  const range = XLSX.utils.decode_range(ws['!ref']);
  console.log(`Range: ${ws['!ref']}`);
  console.log(`Rows: ${range.s.r}-${range.e.r}, Cols: ${range.s.c}-${range.e.c}`);

  // Print first 10 rows, columns A-H (0-7) to understand structure
  console.log('\n--- Rows 0-15, Cols A-H ---');
  for (let r = 0; r <= 15; r++) {
    const cells = [];
    for (let c = 0; c <= 7; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell) {
        let info = `[${addr}]`;
        if (cell.f) info += ` f="${cell.f}"`;
        if (cell.v !== undefined) info += ` v=${JSON.stringify(cell.v).substring(0, 40)}`;
        if (cell.t) info += ` t=${cell.t}`;
        cells.push(info);
      }
    }
    if (cells.length > 0) console.log(`  Row ${r}: ${cells.join(' | ')}`);
  }

  // Now specifically look at column E (index 4) for formulas
  console.log('\n--- Column E (index 4) formulas, rows 0-50 ---');
  for (let r = 0; r <= 50; r++) {
    const addr = XLSX.utils.encode_cell({ r, c: 4 });
    const cell = ws[addr];
    if (cell) {
      let info = `Row ${r} [${addr}]:`;
      if (cell.f) info += ` FORMULA="${cell.f}"`;
      if (cell.v !== undefined) info += ` VALUE=${cell.v}`;
      if (cell.t) info += ` TYPE=${cell.t}`;
      console.log(`  ${info}`);
    }
  }

  // Also check for data validation (dropdown)
  console.log('\n--- Data Validations ---');
  if (ws['!dataValidation']) {
    console.log(JSON.stringify(ws['!dataValidation'], null, 2));
  } else {
    console.log('No data validations found in sheet metadata');
  }

  // Check defined names in workbook
  console.log('\n--- Defined Names ---');
  if (wb.Workbook && wb.Workbook.Names) {
    for (const name of wb.Workbook.Names) {
      console.log(`  ${name.Name} = ${name.Ref}`);
    }
  }

  // Check what's in column A (likely team dropdown reference) and column B
  console.log('\n--- Column A & B, rows 0-10 ---');
  for (let r = 0; r <= 10; r++) {
    const a = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    const b = ws[XLSX.utils.encode_cell({ r, c: 1 })];
    console.log(`  Row ${r}: A=${a ? (a.f ? 'f:'+a.f : a.v) : ''} | B=${b ? (b.f ? 'f:'+b.f : b.v) : ''}`);
  }
}

analyzeRosa(
  path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'),
  'Fanta Tosti'
);

analyzeRosa(
  path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
  'FantaMantra'
);
