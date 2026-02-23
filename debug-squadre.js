const XLSX = require('xlsx');
const path = require('path');

const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function debug(filePath, league) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['SQUADRE'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  console.log(`\n=== ${league} SQUADRE ===`);
  console.log(`Righe: ${data.length}, Colonne max: ${Math.max(...data.map(r => r.length))}`);

  // Print rows 0-6 for first 30 columns
  for (let r = 0; r <= 6; r++) {
    const row = data[r] || [];
    const cells = [];
    for (let c = 0; c < Math.min(30, row.length); c++) {
      const v = String(row[c] || '').trim();
      if (v) cells.push(`[${c}]=${v.substring(0, 25)}`);
    }
    console.log(`Row ${r}: ${cells.join(' | ')}`);
  }

  // Find all "Calciatore" header positions
  for (let r = 0; r < Math.min(10, data.length); r++) {
    const row = data[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (String(row[c]).trim() === 'Calciatore') {
        // Check what's in the row above (team name row)
        const teamRow = data[r - 1] || [];
        const aboveCell = String(teamRow[c] || '').trim();
        console.log(`  "Calciatore" at row ${r}, col ${c} => above="${aboveCell}"`);
      }
    }
  }

  // Also check for team names in rows 2-3
  console.log('\nTeam names scan:');
  for (let r = 2; r <= 4; r++) {
    const row = data[r] || [];
    for (let c = 0; c < row.length; c++) {
      const v = String(row[c] || '').trim();
      if (v && v.length > 3 && !v.match(/^(Calciatore|Ruolo|Squadra|Ass|Data|FVM|Spesa|Numero|Crediti|Q\.|Reparto)/)) {
        console.log(`  Row ${r}, Col ${c}: "${v}"`);
      }
    }
  }
}

debug(
  path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'),
  'FT'
);
debug(
  path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
  'FM'
);
