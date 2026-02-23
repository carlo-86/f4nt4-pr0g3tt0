const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

// Check FT Hellas for specific missing players (Circati, Moreo, Durosinmi)
const wb = XLSX.readFile(path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'), { password: PW });
const ws = wb.Sheets['SQUADRE'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

// Hellas is at col 14
console.log('=== FT Hellas Madonna (col 14) all rows 5-60 ===');
for (let r = 5; r < 60; r++) {
  const v = String(data[r]?.[14] || '').trim();
  if (v && v !== 'Calciatore' && v !== 'Ass. = A') {
    const ins = String(data[r]?.[17] || '').trim();
    const buyDate = data[r]?.[18];
    let bd = '';
    if (typeof buyDate === 'number') {
      const d = XLSX.SSF.parse_date_code(buyDate);
      bd = `${d.d}/${d.m}/${d.y}`;
    }
    console.log(`  Row ${r}: ${v.padEnd(22)} ins=${ins.padEnd(3)} buyDate=${bd} sp=${data[r]?.[24]}`);
  }
}

// Check what names KFP has at col 62
console.log('\n=== FT KFP (col 62) all rows 5-60 ===');
for (let r = 5; r < 60; r++) {
  const v = String(data[r]?.[62] || '').trim();
  if (v && v !== 'Calciatore' && v !== 'Ass. = A') {
    const ins = String(data[r]?.[65] || '').trim();
    console.log(`  Row ${r}: ${v.padEnd(22)} ins=${ins}`);
  }
}

// Check Millwall at col 74
console.log('\n=== FT Millwall (col 74) all rows 5-60 ===');
for (let r = 5; r < 60; r++) {
  const v = String(data[r]?.[74] || '').trim();
  if (v && v !== 'Calciatore' && v !== 'Ass. = A') {
    const ins = String(data[r]?.[77] || '').trim();
    console.log(`  Row ${r}: ${v.padEnd(22)} ins=${ins}`);
  }
}

// Now check FM - missing ones
const wb2 = XLSX.readFile(path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'), { password: PW });
const ws2 = wb2.Sheets['SQUADRE'];
const data2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: '' });

// FC CKC 26 FM: check all Calciatore columns to find it
// From debug: row3 col93 = "FC CKC 26", Calciatore at col94
console.log('\n=== FM FC CKC 26 (col 94) all rows 5-60 ===');
for (let r = 5; r < 60; r++) {
  const v = String(data2[r]?.[94] || '').trim();
  if (v && v !== 'Calciatore' && v !== 'Ass. = A') {
    const ins = String(data2[r]?.[97] || '').trim();
    console.log(`  Row ${r}: ${v.padEnd(22)} ins=${ins}`);
  }
}

// Hellas Madonna FM at col 68 (from debug: row3 col67="Hellas Madonna")
console.log('\n=== FM Hellas Madonna (col 68) all rows 5-60 ===');
for (let r = 5; r < 60; r++) {
  const v = String(data2[r]?.[68] || '').trim();
  if (v && v !== 'Calciatore' && v !== 'Ass. = A') {
    const ins = String(data2[r]?.[71] || '').trim();
    console.log(`  Row ${r}: ${v.padEnd(22)} ins=${ins}`);
  }
}

// Legenda Aurea FM at col 16 (from debug: row3 col15="Legenda Aurea")
console.log('\n=== FM Legenda Aurea (col 16) all rows 5-60 ===');
for (let r = 5; r < 60; r++) {
  const v = String(data2[r]?.[16] || '').trim();
  if (v && v !== 'Calciatore' && v !== 'Ass. = A') {
    const ins = String(data2[r]?.[19] || '').trim();
    console.log(`  Row ${r}: ${v.padEnd(22)} ins=${ins}`);
  }
}

// Minnesota FM at col 81 (from debug: row3 col80="MINNESOTA AL MAX")
console.log('\n=== FM Minnesota (col 81) all rows 5-60 ===');
for (let r = 5; r < 60; r++) {
  const v = String(data2[r]?.[81] || '').trim();
  if (v && v !== 'Calciatore' && v !== 'Ass. = A') {
    const ins = String(data2[r]?.[84] || '').trim();
    console.log(`  Row ${r}: ${v.padEnd(22)} ins=${ins}`);
  }
}

// Lino Banfield FM at col 29 (from debug: row3 col28="Lino Banfield FC")
console.log('\n=== FM Lino Banfield (col 29) all rows 5-60 ===');
for (let r = 5; r < 60; r++) {
  const v = String(data2[r]?.[29] || '').trim();
  if (v && v !== 'Calciatore' && v !== 'Ass. = A') {
    const ins = String(data2[r]?.[32] || '').trim();
    console.log(`  Row ${r}: ${v.padEnd(22)} ins=${ins}`);
  }
}
