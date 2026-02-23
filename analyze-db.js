const XLSX = require('xlsx');
const path = require('path');

const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function analyzeDB(filePath, leagueName) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`DB COMPLETO - ${leagueName}`);
  console.log('='.repeat(80));

  let wb;
  try {
    wb = XLSX.readFile(filePath, { password: PW });
  } catch (e) {
    console.log(`Errore con password: ${e.message}`);
    try {
      wb = XLSX.readFile(filePath);
    } catch (e2) {
      console.log(`Errore senza password: ${e2.message}`);
      return;
    }
  }

  console.log('\nFogli disponibili:', wb.SheetNames);

  // Analyze each sheet briefly
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
    if (range) {
      console.log(`  "${sheetName}": ${range.e.r + 1} righe x ${range.e.c + 1} colonne`);
    }
  }

  // SQUADRE sheet - contains team data with player details
  const squadreSheet = wb.SheetNames.find(s => s.toUpperCase() === 'SQUADRE');
  if (squadreSheet) {
    console.log(`\n--- Foglio SQUADRE ---`);
    const ws = wb.Sheets[squadreSheet];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Print header rows
    console.log('Header (row 0):', JSON.stringify(data[0]?.slice(0, 30)));
    console.log('Header (row 1):', JSON.stringify(data[1]?.slice(0, 30)));

    // Find all team sections
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const firstCell = String(row[0] || '').trim();
      if (firstCell && firstCell.length > 3 && !firstCell.match(/^\d/) && i > 0) {
        // Might be a team header
        const prevRow = data[i-1];
        const prevEmpty = !prevRow || prevRow.every(c => !c || String(c).trim() === '');
        if (prevEmpty || i === 2) {
          console.log(`\nTeam at row ${i}: "${firstCell}"`);
          // Print next 5 rows
          for (let j = i; j < Math.min(i + 5, data.length); j++) {
            console.log(`  Row ${j}: ${JSON.stringify(data[j].slice(0, 20))}`);
          }
        }
      }
    }
  }

  // ROSA sheet - contains detailed roster with costs
  const rosaSheet = wb.SheetNames.find(s => s.toUpperCase() === 'ROSA');
  if (rosaSheet) {
    console.log(`\n--- Foglio ROSA ---`);
    const ws = wb.Sheets[rosaSheet];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Print ALL rows (it shouldn't be too many)
    console.log(`Totale righe: ${data.length}`);

    // Print header and first rows
    for (let i = 0; i < Math.min(40, data.length); i++) {
      const row = data[i];
      if (row.some(c => c !== '' && c !== null && c !== undefined)) {
        // Print cols A-J + cols around AC-AF (28-31)
        const mainCols = row.slice(0, 10).map(c => c === '' ? '' : c);
        const dropdownCols = row.length > 28 ? row.slice(28, 33) : [];
        console.log(`Row ${i}: A-J=${JSON.stringify(mainCols)} | AC-AG=${JSON.stringify(dropdownCols)}`);
      }
    }
  }

  // FVM/Quotazioni sheet
  const fvmSheets = wb.SheetNames.filter(s =>
    s.toUpperCase().includes('FVM') ||
    s.toUpperCase().includes('QUOTAZ') ||
    s.toUpperCase().includes('LISTONE')
  );
  for (const sheetName of fvmSheets) {
    console.log(`\n--- Foglio ${sheetName} ---`);
    const ws = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    console.log(`Righe: ${data.length}`);
    for (let i = 0; i < Math.min(5, data.length); i++) {
      console.log(`Row ${i}: ${JSON.stringify(data[i].slice(0, 15))}`);
    }
  }
}

const league = process.argv[2] || 'ft';
if (league === 'ft') {
  analyzeDB(
    path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'),
    'Fanta Tosti'
  );
} else {
  analyzeDB(
    path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
    'FantaMantra Manageriale'
  );
}
