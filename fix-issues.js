const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function norm(s) {
  return String(s||'').trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase()
    .replace(/[''`\u2019\u2018]/g,'').replace(/\./g,'').replace(/\s+/g,' ').trim();
}

// ===== 1. Search for Santos A. in DB and Rose =====
console.log('=== 1. SANTOS A. (= Allison S.) ===\n');

function searchInDB(filePath, label, searchTerms) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['DB'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  for (const term of searchTerms) {
    const nTerm = norm(term);
    for (let r = 2; r < Math.min(data.length, 1000); r++) {
      const val = norm(String(data[r]?.[62] || ''));
      if (val && val.includes(nTerm)) {
        console.log(`  ${label} DB: "${data[r][62]}" FVM=${data[r][79]} (row ${r})`);
      }
    }
  }
}

function searchInRose(filePath, label, searchTerms) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['TutteLeRose'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  for (const term of searchTerms) {
    const nTerm = norm(term);
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < (data[r]?.length || 0); c++) {
        const val = norm(String(data[r]?.[c] || ''));
        if (val && val.includes(nTerm) && val.length < 30) {
          const ctx = [];
          for (let cc = Math.max(0,c-1); cc < Math.min(data[r].length, c+4); cc++) {
            if (data[r][cc] !== '') ctx.push(`col${cc}="${data[r][cc]}"`);
          }
          console.log(`  ${label} Rose row ${r}: ${ctx.join(', ')}`);
        }
      }
    }
  }
}

searchInDB(path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'), 'FT', ['SANTOS']);
searchInRose(path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'), 'FT', ['SANTOS']);

// ===== 2. Kone I. in FT DB =====
console.log('\n=== 2. KONE I. vs KONE M. nel DB FT ===\n');
searchInDB(path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'), 'FT', ['KONE']);

// Also check FT Rose for CKC's Kone
console.log('\nCerco Kone nella Rose FT (per capire quale Kone ha CKC):');
searchInRose(path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'), 'FT', ['KONE']);

// ===== 3. David in FM Rose - detailed check in Hellas column =====
console.log('\n=== 3. DAVID nella Rose FM - verifica colonna Hellas ===\n');

{
  const wb = XLSX.readFile(path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'), { password: PW });
  const ws = wb.Sheets['TutteLeRose'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Find Hellas Madonna team position
  const roles = new Set(['P','D','C','A','Por','Dc','Dd','Ds','E','B','M','T','W','Pc']);
  function isRole(v) { return v && String(v).trim().split(';').every(p => roles.has(p.trim())); }

  // Parse all teams and search for David
  for (const baseCol of [0, 5]) {
    let currentTeam = null;
    for (let r = 4; r < data.length; r++) {
      const c0 = String(data[r]?.[baseCol] || '').trim();
      const c1 = String(data[r]?.[baseCol + 1] || '').trim();
      if (c0 && !isRole(c0) && !c0.includes(';') && c0.length > 1 && !c0.includes('Crediti') && c0 !== 'Ruolo') {
        currentTeam = c0;
        continue;
      }
      if (currentTeam && c1 && norm(c1).includes('DAVID') && !norm(c1).includes('DAVIDSON')) {
        console.log(`  Team: ${currentTeam}, Row ${r}: col${baseCol}="${c0}" col${baseCol+1}="${c1}" col${baseCol+3}="${data[r]?.[baseCol+3]}"`);
      }
    }
  }
}

// ===== 4. Check for unlisted players (asterisk marked) =====
console.log('\n=== 4. CALCIATORI NON PIU LISTATI (con asterisco *) ===\n');

function checkUnlisted(filePath, label, insuredPlayers) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['TutteLeRose'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Find ALL players with asterisk in their name
  const asteriskPlayers = new Set();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < (data[r]?.length || 0); c++) {
      const val = String(data[r]?.[c] || '');
      if (val.includes('*') && val.length > 1 && val.length < 40) {
        asteriskPlayers.add(norm(val.replace('*', '')));
        // Check if this is one of our insured players
        const nv = norm(val.replace('*', ''));
        for (const ip of insuredPlayers) {
          if (nv.includes(norm(ip)) || norm(ip).includes(nv)) {
            console.log(`  ${label} ATTENZIONE: "${val}" (row ${r}, col ${c}) -> non piu listato!`);
          }
        }
      }
    }
  }
  console.log(`  ${label}: ${asteriskPlayers.size} giocatori con asterisco totali`);

  // Also check: players with very low FVM that might be unlisted
  const dbWb = XLSX.readFile(
    filePath.includes('Fanta Tosti')
      ? path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx')
      : path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
    { password: PW }
  );
  const dbWs = dbWb.Sheets['DB'];
  const dbData = XLSX.utils.sheet_to_json(dbWs, { header: 1, defval: '' });

  // Check each insured player's FVM and if they're asterisked
  console.log(`\n  ${label} - Insured players with FVM=1 or asterisk:`);
  for (const ip of insuredPlayers) {
    const nip = norm(ip);
    // Check FVM
    for (let r = 2; r < Math.min(dbData.length, 1000); r++) {
      const name = norm(String(dbData[r]?.[62] || ''));
      if (name && (name.includes(nip) || nip.includes(name)) && Math.min(name.length, nip.length) >= 4) {
        const fvm = Number(dbData[r]?.[79]) || 0;
        const isAsterisk = asteriskPlayers.has(name);
        if (fvm <= 1 || isAsterisk) {
          console.log(`    ${ip.padEnd(24)} FVM=${fvm} ${isAsterisk ? '*** NON LISTATO ***' : '(FVM basso, verificare)'}`);
        }
        break;
      }
    }
  }
}

const FT_INSURED = [
  'SPORTIELLO', 'CIRCATI', 'BERISHA', 'MOREO', 'DUROSINMI',
  'BELGHALI', 'STREFEZZA', 'PRZYBOREK',
  'MALEN', 'VERGARA', 'BEUKEMA', 'KOUAME',
  'SANTOS', 'VAZ', 'MUHAREMOVIC', 'BALDANZI', 'BIJLOW', 'BERNASCONI', 'KONE',
  'OSTIGARD', 'LUIS HENRIQUE', 'SOLOMON',
  'MURIC', 'CELIK', 'RATKOV', 'ZARAGOZA', 'PERRONE', 'PALEARI', 'BOGA', 'HOLM',
  'HIEN',
  'DI GREGORIO', 'SOMMER', 'MARTINEZ', 'KALULU', 'BARTESAGHI', 'LOVRIC',
  'TAYLOR', 'FAGIOLI', 'EKKELENKAMP', 'MIRETTI', 'BONAZZOLI', 'RASPADORI', 'VITINHA'
];

const FM_INSURED = [
  'KONE', 'RASPADORI', 'FERGUSON', 'KOUAME',
  'DUROSINMI', 'VERGARA', 'ZANIOLO',
  'HOLM', 'NDICKA', 'GALLO', 'VASQUEZ', 'GUDMUNDSSON', 'FRENDRUP', 'BRITSCHGI', 'SULEMANA', 'TAYLOR', 'MALEN', 'SOMMER',
  'DAVID', 'CHEDDIRA', 'ZARAGOZA', 'EKKELENKAMP', 'BRESCIANINI', 'BELGHALI', 'SCAMACCA',
  'LUIS HENRIQUE', 'FULLKRUG',
  'CELIK', 'OBERT', 'MARCANDALLI', 'BERNASCONI', 'BOWIE', 'CAPRILE', 'CAMBIAGHI', 'VAZ', 'BALDANZI', 'KOOPMEINERS', 'TAVARES', 'MAZZITELLI',
  'MONTIPO', 'MARIANUCCI', 'CATALDI', 'FAGIOLI', 'MILLER', 'BAKOLA', 'ADZIC', 'RATKOV', 'BELLANOVA',
  'KOLASINAC', 'HIEN', 'PASALIC', 'NICOLUSSI CAVIGLIA', 'SOLOMON', 'VLAHOVIC',
  'NELSSON', 'DOSSENA', 'BARTESAGHI', 'GANDELMAN', 'BARBIERI', 'LEAO', 'ZAPPA'
];

checkUnlisted(
  path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
  'FT', FT_INSURED
);

checkUnlisted(
  path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
  'FM', FM_INSURED
);
