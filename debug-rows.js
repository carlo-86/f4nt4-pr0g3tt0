const XLSX = require('xlsx');
const path = require('path');

const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

// Look at FT SQUADRE, first team block (col 2 = FCK deportivo)
const wb = XLSX.readFile(path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'), { password: PW });
const ws = wb.Sheets['SQUADRE'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

// Print col 2 (Calciatore) for rows 4-50 for FCK deportivo
console.log('=== FCK deportivo (col 2) rows 4-50 ===');
for (let r = 4; r <= 50; r++) {
  const row = data[r] || [];
  const c2 = String(row[2] || '').trim();
  const c3 = String(row[3] || '').trim();
  const c4 = String(row[4] || '').trim();
  const c5 = String(row[5] || '').trim();
  console.log(`Row ${r}: col2="${c2}" col3="${c3}" col4="${c4}" col5="${c5}"`);
}

// Print col 14 (Calciatore for Hellas) for rows 4-50
console.log('\n=== Hellas Madonna (col 14) rows 4-50 ===');
for (let r = 4; r <= 50; r++) {
  const row = data[r] || [];
  const c14 = String(row[14] || '').trim();
  const c15 = String(row[15] || '').trim();
  const c16 = String(row[16] || '').trim();
  const c17 = String(row[17] || '').trim();
  console.log(`Row ${r}: col14="${c14}" col15="${c15}" col16="${c16}" col17="${c17}"`);
}

// Also check FM first team
const wb2 = XLSX.readFile(path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'), { password: PW });
const ws2 = wb2.Sheets['SQUADRE'];
const data2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: '' });

console.log('\n=== FM Papaie Top Team (col 3) rows 4-50 ===');
for (let r = 4; r <= 50; r++) {
  const row = data2[r] || [];
  const c3 = String(row[3] || '').trim();
  const c4 = String(row[4] || '').trim();
  const c5 = String(row[5] || '').trim();
  const c6 = String(row[6] || '').trim();
  console.log(`Row ${r}: col3="${c3}" col4="${c4}" col5="${c5}" col6="${c6}"`);
}
