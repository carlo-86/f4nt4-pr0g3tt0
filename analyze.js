const XLSX = require('xlsx');
const path = require('path');

const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';

// ===== 1. CONTEGGI SVINCOLI =====
function parseConteggiSvincoli(filePath, leagueName) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`CONTEGGI SVINCOLI - ${leagueName}`);
  console.log(`File: ${path.basename(filePath)}`);
  console.log('='.repeat(80));

  const wb = XLSX.readFile(filePath);

  console.log('\nFogli disponibili:', wb.SheetNames);

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    console.log(`\n--- Foglio: "${sheetName}" (righe: ${range.e.r + 1}, colonne: ${range.e.c + 1}) ---`);

    // Print all data as JSON
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    for (let i = 0; i < Math.min(data.length, 100); i++) {
      const row = data[i];
      // Only print non-empty rows
      if (row.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
        console.log(`Row ${i}: ${JSON.stringify(row)}`);
      }
    }
    if (data.length > 100) {
      console.log(`... (${data.length - 100} more rows)`);
    }
  }
}

// ===== 2. ROSE =====
function parseRose(filePath, leagueName) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`ROSE - ${leagueName}`);
  console.log(`File: ${path.basename(filePath)}`);
  console.log('='.repeat(80));

  const wb = XLSX.readFile(filePath);
  console.log('\nFogli disponibili:', wb.SheetNames);

  const sheetName = wb.SheetNames.find(s => s.toLowerCase().includes('tutterose') || s.toLowerCase().includes('tutte')) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  console.log(`\nFoglio: "${sheetName}" - ${data.length} righe`);

  // Find team headers and credits
  let currentTeam = '';
  let teamCredits = {};
  let teamPlayers = {};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const firstCell = String(row[0] || '').trim();

    // Look for team name pattern (usually uppercase or specific pattern)
    if (firstCell && !firstCell.match(/^\d+$/) && row.length > 0) {
      // Check if this might be a team header
      const secondCell = String(row[1] || '').trim();
      const thirdCell = String(row[2] || '').trim();

      // Print first 5 rows to understand structure
      if (i < 10) {
        console.log(`Row ${i}: ${JSON.stringify(row.slice(0, 10))}`);
      }
    }
  }

  // Print all rows to understand the structure
  console.log('\n--- Prime 30 righe ---');
  for (let i = 0; i < Math.min(30, data.length); i++) {
    const row = data[i];
    if (row.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
      console.log(`Row ${i}: ${JSON.stringify(row.slice(0, 8))}`);
    }
  }

  // Search for "Crediti" or credit patterns
  console.log('\n--- Ricerca crediti squadre ---');
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    for (let j = 0; j < row.length; j++) {
      const cell = String(row[j] || '');
      if (cell.toLowerCase().includes('credit') || cell.toLowerCase().includes('cred')) {
        console.log(`Row ${i}, Col ${j}: "${cell}" | Full row: ${JSON.stringify(row.slice(0, 8))}`);
      }
    }
  }

  // Also look for team names
  console.log('\n--- Ricerca nomi squadra ---');
  const teamNames = ['Hellas', 'Partizan', 'Kung Fu', 'CKC', 'Muttley', 'Deportivo', 'Millwall',
                     'Papaie', 'Legenda', 'HQA', 'FICA', 'Lino Banfield', 'Minnesota'];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    for (let j = 0; j < row.length; j++) {
      const cell = String(row[j] || '');
      for (const tn of teamNames) {
        if (cell.toLowerCase().includes(tn.toLowerCase())) {
          console.log(`Row ${i}, Col ${j}: "${cell}" | Row: ${JSON.stringify(row.slice(0, 8))}`);
          break;
        }
      }
    }
  }
}

// ===== 3. DB COMPLETO - ROSA SHEET =====
function parseDBCompleto(filePath, leagueName) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`DB COMPLETO - ${leagueName}`);
  console.log(`File: ${path.basename(filePath)}`);
  console.log('='.repeat(80));

  let wb;
  try {
    wb = XLSX.readFile(filePath);
  } catch(e) {
    // Try with password
    const pwFile = filePath.replace(/\.xlsx$/, '').replace(/ \([^)]+\)/, '') + '_pw.txt';
    console.log(`File protetto, provo a leggere password da: ${pwFile}`);
    try {
      const fs = require('fs');
      const pw = fs.readFileSync(path.join(path.dirname(filePath), path.basename(path.dirname(filePath)).replace('DB Excel', leagueName) + ' - DB completo_pw.txt'), 'utf8').trim();
      console.log(`Password trovata: ${pw}`);
      wb = XLSX.readFile(filePath, { password: pw });
    } catch(e2) {
      console.log(`Errore lettura: ${e2.message}`);
      // Try without password protection
      wb = XLSX.readFile(filePath, { password: '' });
    }
  }

  console.log('\nFogli disponibili:', wb.SheetNames);

  // Look for ROSA, SQUADRE, and FVM sheets
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null;
    if (range) {
      console.log(`  "${sheetName}": ${range.e.r + 1} righe x ${range.e.c + 1} colonne`);
    }
  }

  // Parse SQUADRE sheet
  const squadreSheet = wb.SheetNames.find(s => s.toUpperCase() === 'SQUADRE' || s.toUpperCase().includes('SQUADRE'));
  if (squadreSheet) {
    console.log(`\n--- Foglio SQUADRE ---`);
    const ws = wb.Sheets[squadreSheet];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    console.log(`Righe: ${data.length}`);
    // Print first 5 rows for structure
    for (let i = 0; i < Math.min(5, data.length); i++) {
      console.log(`Row ${i}: ${JSON.stringify(data[i].slice(0, 15))}`);
    }
  }

  // Parse ROSA sheet
  const rosaSheet = wb.SheetNames.find(s => s.toUpperCase() === 'ROSA' || s.toUpperCase().includes('ROSA'));
  if (rosaSheet) {
    console.log(`\n--- Foglio ROSA ---`);
    const ws = wb.Sheets[rosaSheet];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    console.log(`Righe: ${data.length}`);
    // Print first 10 rows for structure
    for (let i = 0; i < Math.min(10, data.length); i++) {
      console.log(`Row ${i}: ${JSON.stringify(data[i].slice(0, 10))}`);
    }

    // Find column E (index 4) headers and check for cost data
    console.log('\n--- Colonna E (costo) primi 30 righe ---');
    for (let i = 0; i < Math.min(30, data.length); i++) {
      if (data[i][4] !== '' && data[i][4] !== null && data[i][4] !== undefined) {
        console.log(`Row ${i}: Col A="${data[i][0]}", Col B="${data[i][1]}", Col C="${data[i][2]}", Col D="${data[i][3]}", Col E="${data[i][4]}"`);
      }
    }

    // Print column headers (rows AC-AF area, around col 28-31)
    console.log('\n--- Colonne AC-AF (menu a tendina squadre) ---');
    for (let i = 0; i < Math.min(5, data.length); i++) {
      const row = data[i];
      if (row.length > 28) {
        console.log(`Row ${i}: Col AC(28)="${row[28]}", Col AD(29)="${row[29]}", Col AE(30)="${row[30]}", Col AF(31)="${row[31]}"`);
      }
    }

    // Print ALL data in ROSA sheet (all rows, all relevant columns)
    console.log('\n--- ROSA: tutte le righe con dati ---');
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
        // Print cols A-J (0-9)
        console.log(`Row ${i}: ${JSON.stringify(row.slice(0, 10))}`);
      }
    }
  }

  // Parse FVM-related sheets
  const fvmSheet = wb.SheetNames.find(s => s.toUpperCase().includes('FVM') || s.toUpperCase().includes('QUOTAZ'));
  if (fvmSheet) {
    console.log(`\n--- Foglio ${fvmSheet} ---`);
    const ws = wb.Sheets[fvmSheet];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    console.log(`Righe: ${data.length}`);
    for (let i = 0; i < Math.min(5, data.length); i++) {
      console.log(`Row ${i}: ${JSON.stringify(data[i].slice(0, 15))}`);
    }
  }
}

// Run analysis
const args = process.argv.slice(2);
const mode = args[0] || 'all';

if (mode === 'svincoli-ft' || mode === 'all') {
  parseConteggiSvincoli(
    path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Conteggi svincoli (07.02.2026).xlsx'),
    'Fanta Tosti'
  );
}

if (mode === 'svincoli-fm' || mode === 'all') {
  parseConteggiSvincoli(
    path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Conteggi svincoli (07.02.2026).xlsx'),
    'FantaMantra Manageriale'
  );
}

if (mode === 'rose-ft' || mode === 'all') {
  parseRose(
    path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
    'Fanta Tosti'
  );
}

if (mode === 'rose-fm' || mode === 'all') {
  parseRose(
    path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
    'FantaMantra Manageriale'
  );
}

if (mode === 'db-ft' || mode === 'all') {
  parseDBCompleto(
    path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'),
    'Fanta Tosti'
  );
}

if (mode === 'db-fm' || mode === 'all') {
  parseDBCompleto(
    path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
    'FantaMantra Manageriale'
  );
}
