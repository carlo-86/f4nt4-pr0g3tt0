const XLSX = require('xlsx');
const fs = require('fs');

const files = [
  { label: 'FT', path: 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/Fanta Tosti 2026/DB Excel/Fanta Tosti 2026 - DB completo (06.02.2026).xlsx',
    teams: [
      {name:'FCK',col:3},{name:'Hellas',col:15},{name:'muttley',col:27},{name:'PARTIZAN',col:39},
      {name:'Legenda',col:51},{name:'KFP',col:63},{name:'Millwall',col:75},{name:'CKC',col:87},
      {name:'Papaie',col:99},{name:'Tronzano',col:111}
    ]
  },
  { label: 'FM', path: 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/FantaMantra Manageriale/DB Excel/FantaMantra Manageriale - DB completo (06.02.2026).xlsx',
    teams: [
      {name:'Papaie',col:4},{name:'Legenda',col:17},{name:'Lino',col:30},{name:'KFP',col:43},
      {name:'FICA',col:56},{name:'Hellas',col:69},{name:'Minnesota',col:82},{name:'CKC',col:95},
      {name:'HQA',col:108},{name:'Mastri',col:121}
    ]
  }
];

function cellVal(ws, r, c) {
  const addr = XLSX.utils.encode_cell({r, c});
  const cell = ws[addr];
  if (!cell) return '';
  return cell.v !== undefined ? cell.v : '';
}

function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return `${String(v.getDate()).padStart(2,'0')}/${String(v.getMonth()+1).padStart(2,'0')}/${v.getFullYear()}`;
  }
  return String(v);
}

// Players we're insuring in this macro run (names to search for)
const ftInsuredNames = ['Sportiello','Circati','Berisha','Moreo','Durosinmi','Belghali','Strefezza','Przyborek',
  'Malen','Vergara','Beukema','Tiago Gabriel','Vaz','Muharemovic','Baldanzi','Santos','Bijlow','Bernasconi','Kon',
  'Ostigard','Luis','Solomon','Muric','Celik','Ratkov','Zaragoza','Perrone','Paleari','Boga','Holm',
  'Hien','Di Gregorio','Sommer','Martinez','Kalulu','Bartesaghi','Lovric','Taylor','Fagioli','Ekkelenkamp',
  'Miretti','Bonazzoli','Raspadori','Vitinha'];

const fmInsuredNames = ['Kon','Raspadori','Ferguson','Durosinmi','Vergara','Zaniolo','Holm','Ndicka','Gallo',
  'Vasquez','Gudmundsson','Frendrup','Britschgi','Sulemana','Taylor','Malen','Sommer','David','Cheddira',
  'Zaragoza','Ekkelenkamp','Brescianini','Belghali','Scamacca','Luis','Fullkrug','Celik','Obert','Marcandalli',
  'Bernasconi','Bowie','Caprile','Cambiaghi','Vaz','Baldanzi','Koopmeiners','Mazzitelli','Montip','Marianucci',
  'Cataldi','Fagioli','Miller','Bakola','Adzic','Ratkov','Bellanova','Kolasinac','Hien','Pasalic',
  'Nicolussi','Solomon','Vlahovic','Nelsson','Dossena','Bartesaghi','Gandelman','Barbieri','Leao','Zappa'];

for (const file of files) {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`${file.label}`);
  console.log('='.repeat(60));

  const wb = XLSX.read(fs.readFileSync(file.path), { cellDates: true });
  const ws = wb.Sheets['SQUADRE'];

  // First, show the column headers for the first team to understand layout
  const firstTeam = file.teams[0];
  console.log(`\n--- SQUADRE column layout (${firstTeam.name}, col ${firstTeam.col}) ---`);
  const baseCol = firstTeam.col - 1; // 0-based
  // Show headers in rows 2-5 (0-based 1-4)
  for (let r = 0; r <= 4; r++) {
    const vals = [];
    for (let c = baseCol; c < baseCol + 12; c++) {
      const v = cellVal(ws, r, c);
      vals.push(`+${c-baseCol}:${String(v).substring(0,20)}`);
    }
    console.log(`  Row ${r+1}: ${vals.join(' | ')}`);
  }

  // Now show first 3 data rows with all columns
  console.log(`\n--- First 3 players of ${firstTeam.name} (full data) ---`);
  for (let r = 5; r <= 7; r++) {
    const vals = [];
    for (let c = baseCol; c < baseCol + 12; c++) {
      const v = cellVal(ws, r, c);
      const fv = (v instanceof Date) ? fmtDate(v) : String(v).substring(0,20);
      vals.push(`+${c-baseCol}:${fv}`);
    }
    console.log(`  Row ${r+1}: ${vals.join(' | ')}`);
  }

  // Now for each team, find players that will be insured by our macro and show their current state
  const insuredNames = file.label === 'FT' ? ftInsuredNames : fmInsuredNames;

  console.log(`\n--- Players being insured that already have flag "A" ---`);
  let renewalCount = 0;
  let newCount = 0;

  for (const team of file.teams) {
    const bc = team.col - 1; // 0-based

    for (let r = 5; r <= 51; r++) {
      const name = String(cellVal(ws, r, bc)).trim();
      if (!name || name === 'Calciatore') continue;

      // Check if this player is in our insurance list
      const nameUpper = name.toUpperCase().replace(/'/g, '').replace(/\./g, '');
      let matched = false;
      for (const search of insuredNames) {
        const searchUpper = search.toUpperCase().replace(/'/g, '').replace(/\./g, '');
        if (nameUpper.includes(searchUpper)) {
          matched = true;
          break;
        }
      }
      if (!matched) continue;

      const flag = String(cellVal(ws, r, bc + 3)).trim();
      const insDate = cellVal(ws, r, bc + 7);
      const insDateStr = fmtDate(insDate);

      if (flag === 'A' && insDate) {
        // This is a RENEWAL - player already insured
        // Calculate triennium expiry
        let expiry = '';
        if (insDate instanceof Date) {
          const exp = new Date(insDate);
          exp.setFullYear(exp.getFullYear() + 3);
          expiry = fmtDate(exp);
          const cutoff = new Date(2026, 1, 14); // 14/02/2026
          const isPreventive = exp > cutoff;
          console.log(`  ${team.name.padEnd(12)} ${name.padEnd(22)} Flag=${flag} Date=${insDateStr} Expiry=${expiry} ${isPreventive ? '*** PREVENTIVO ***' : '(scaduto)'}`);
        } else {
          console.log(`  ${team.name.padEnd(12)} ${name.padEnd(22)} Flag=${flag} Date=${insDateStr} (non-date value)`);
        }
        renewalCount++;
      } else {
        // NEW insurance
        console.log(`  ${team.name.padEnd(12)} ${name.padEnd(22)} Flag=${flag || '(vuoto)'} Date=${insDateStr || '(vuoto)'} -> NUOVA ASSICURAZIONE`);
        newCount++;
      }
    }
  }
  console.log(`\nTotale: ${renewalCount} rinnovi, ${newCount} nuove assicurazioni`);
}
