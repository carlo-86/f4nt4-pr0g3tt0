const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

// ========== NORMALIZATION ==========
function norm(name) {
  return String(name || '').trim()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')  // strip ALL diacritics
    .toUpperCase()
    .replace(/[''`\u2019\u2018]/g, '')
    .replace(/\./g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// ========== ROSE PARSER (17/02 - post-market, has all players with Spesa) ==========
function parseRose(filePath) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['TutteLeRose'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const roles = new Set(['P','D','C','A','Por','Dc','Dd','Ds','E','B','M','T','W','Pc']);
  function isRole(v) {
    if (!v) return false;
    return String(v).trim().split(';').every(p => roles.has(p.trim()));
  }

  function parseGroup(baseCol) {
    const result = [];
    let currentTeam = null;
    let currentPlayers = [];
    for (let r = 4; r < data.length; r++) {
      const c0 = String(data[r]?.[baseCol] || '').trim();
      const c1 = String(data[r]?.[baseCol + 1] || '').trim();
      const c3 = data[r]?.[baseCol + 3];
      if (c0 === 'Ruolo' || c1 === 'Calciatore') continue;
      if (c0.includes('Crediti') || c1.includes('Crediti')) continue;
      if (c0 && !isRole(c0) && !c0.includes(';') && c0.length > 1) {
        if (currentTeam) result.push({ team: currentTeam, players: currentPlayers });
        currentTeam = c0;
        currentPlayers = [];
        continue;
      }
      if (currentTeam && c1 && isRole(c0)) {
        currentPlayers.push({ name: c1, spesa: Number(c3) || 0 });
      }
    }
    if (currentTeam) result.push({ team: currentTeam, players: currentPlayers });
    return result;
  }

  const allGroups = [...parseGroup(0), ...parseGroup(5)];
  const map = new Map();
  for (const g of allGroups) {
    for (const p of g.players) {
      const key = norm(p.name);
      map.set(`${key}|${g.team}`, { name: p.name, team: g.team, spesa: p.spesa });
      if (!map.has(key)) map.set(key, { name: p.name, team: g.team, spesa: p.spesa });
    }
  }
  return { groups: allGroups, map };
}

// ========== DB PARSER (06/02 - for FVM Prop.) ==========
function parseFVM(filePath) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['DB'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const map = new Map();
  for (let r = 2; r < Math.min(data.length, 1000); r++) {
    const name = String(data[r]?.[62] || '').trim();
    const fvm = Number(data[r]?.[79]) || 0;
    if (name && name.length > 1) {
      const key = norm(name);
      if (!map.has(key)) map.set(key, { name, fvm });
    }
  }
  return map;
}

// ========== SQUADRE PARSER (06/02 - for insurance status) ==========
function parseSquadre(filePath, teamCols) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['SQUADRE'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const map = new Map();
  const skip = /^(Calciatore|Ass|Totali|PORTIERI|DIFENSORI|CENTROCAMPISTI|ATTACCANTI|P|D|C|A)$/i;
  for (const [teamName, col] of teamCols) {
    for (let r = 5; r < 52; r++) {
      const name = String(data[r]?.[col] || '').trim();
      if (!name || skip.test(name) || name.length < 2) continue;
      const key = norm(name);
      const ins = String(data[r]?.[col + 3] || '').trim();
      if (key) map.set(key, { name, team: teamName, ins });
    }
  }
  return map;
}

// ========== NAME ALIASES (known mismatches) ==========
const NAME_FIXES = {
  'LUIS ENRIQUE': 'LUIS HENRIQUE',
  'SCAMMACCA': 'SCAMACCA',
};

function levenshtein(a, b) {
  const m=a.length, n=b.length;
  const dp = Array.from({length:m+1},()=>Array(n+1).fill(0));
  for(let i=0;i<=m;i++) dp[i][0]=i;
  for(let j=0;j<=n;j++) dp[0][j]=j;
  for(let i=1;i<=m;i++) for(let j=1;j<=n;j++)
    dp[i][j]=Math.min(dp[i-1][j]+1,dp[i][j-1]+1,dp[i-1][j-1]+(a[i-1]!==b[j-1]?1:0));
  return dp[m][n];
}

// ========== PLAYER FINDER ==========
function findInMap(map, searchName, preferTeam) {
  let key = norm(searchName);
  // Apply name fixes
  if (NAME_FIXES[key]) key = norm(NAME_FIXES[key]);

  // 1. Exact match with preferred team
  if (preferTeam) {
    const tk = `${key}|${preferTeam}`;
    if (map.has(tk)) return map.get(tk);
    for (const [k, v] of map) {
      if (!k.includes('|')) continue;
      const [nk, t] = k.split('|');
      if (t === preferTeam && nk.length >= 4 && key.length >= 4) {
        if (nk.startsWith(key) || key.startsWith(nk)) return v;
      }
    }
  }

  // 2. Global exact
  if (map.has(key)) return map.get(key);

  // 3. Global starts-with (min 4 chars)
  if (key.length >= 4) {
    for (const [k, v] of map) {
      if (k.includes('|')) continue;
      const nk = norm(k);
      if (nk.length >= 4 && (nk.startsWith(key) || key.startsWith(nk))) return v;
    }
  }

  // 4. Global contains (min 5 chars)
  if (key.length >= 5) {
    for (const [k, v] of map) {
      if (k.includes('|')) continue;
      const nk = norm(k);
      if (nk.length >= 5 && (nk.includes(key) || key.includes(nk))) return v;
    }
  }

  // 5. Levenshtein distance <=1 (min 5 chars)
  if (key.length >= 5) {
    for (const [k, v] of map) {
      if (k.includes('|')) continue;
      const nk = norm(k);
      if (nk.length >= 5 && levenshtein(nk, key) <= 1) return v;
    }
  }

  return null;
}

// ========== COST FORMULA ==========
function calcCost(spesa, fvm) {
  if (spesa <= fvm) return Math.max(fvm / 10, 1);
  return Math.max((spesa + fvm) / 2 / 10, 1);
}

// ========== TEAM COLUMNS FOR SQUADRE ==========
const FT_COLS = [
  ['FCK Deportivo', 2], ['Hellas Madonna', 14], ['muttley superstar', 26],
  ['PARTIZAN', 38], ['Legenda Aurea', 50], ['Kung Fu Pandev', 62],
  ['Millwall', 74], ['FC CKC 26', 86], ['Papaie Top Team', 98], ['Tronzano', 110]
];
const FM_COLS = [
  ['Papaie Top Team', 3], ['Legenda Aurea', 16], ['Lino Banfield FC', 29],
  ['Kung Fu Pandev', 42], ['FICA', 55], ['Hellas Madonna', 68],
  ['MINNESOTA AL MAX', 81], ['FC CKC 26', 94], ['H-Q-A Barcelona', 107], ['Mastri Birrai', 120]
];

// ========== INSURANCE REQUESTS ==========
const FT_REQ = {
  'Hellas Madonna': [
    { n: 'SPORTIELLO' }, { n: 'CIRCATI' }, { n: 'BERISHA' },
    { n: 'MOREO' }, { n: 'DUROSINMI', note: 'scritto "Duronisimi"' }
  ],
  'PARTIZAN': [
    { n: 'BELGHALI' }, { n: 'STREFEZZA' }, { n: 'PRZYBOREK' }
  ],
  'Kung Fu Pandev': [
    { n: 'MALEN' }, { n: 'VERGARA' },
    { n: 'BEUKEMA', note: 'rinnovo preventivo triennale' },
    { n: 'KOUAME' }
  ],
  'FC CKC 26': [
    { n: 'TIAGO GABRIEL' }, { n: 'VAZ' }, { n: 'MUHAREMOVIC' },
    { n: 'BALDANZI' }, { n: 'ALLISON', note: 'scritto "Allison S."' },
    { n: 'BIJLOW' }, { n: 'BERNASCONI' }, { n: 'KONE', note: 'scritto "Kone I."' }
  ],
  'muttley superstar': [
    { n: 'OSTIGARD' }, { n: 'LUIS ENRIQUE' }, { n: 'SOLOMON' }
  ],
  'Millwall': [
    { n: 'MURIC' }, { n: 'CELIK' }, { n: 'RATKOV' }, { n: 'ZARAGOZA' },
    { n: 'PERRONE' }, { n: 'PALEARI' }, { n: 'BOGA' }, { n: 'HOLM' }
  ],
  'Papaie Top Team': [
    { n: 'HIEN' }
  ],
  'Legenda Aurea': [
    { n: 'DI GREGORIO' }, { n: 'SOMMER' }, { n: 'MARTINEZ' },
    { n: 'KALULU' }, { n: 'BARTESAGHI' }, { n: 'LOVRIC' },
    { n: 'TAYLOR' }, { n: 'FAGIOLI' }, { n: 'EKKELENKAMP' },
    { n: 'MIRETTI' }, { n: 'BONAZZOLI' }, { n: 'RASPADORI' }, { n: 'VITINHA' }
  ]
};

const FM_REQ = {
  'Kung Fu Pandev': [
    { n: 'KONE', note: 'scritto "Kone"' }, { n: 'RASPADORI' },
    { n: 'POSCH', note: 'RESPINTO - svincolato, non assicurabile' },
    { n: 'FERGUSON' }, { n: 'KOUAME' }
  ],
  'FC CKC 26': [
    { n: 'DUROSINMI' }, { n: 'VERGARA' }, { n: 'ZANIOLO' }
  ],
  'H-Q-A Barcelona': [
    { n: 'HOLM' }, { n: 'NDICKA' }, { n: 'GALLO' }, { n: 'VASQUEZ' },
    { n: 'GUDMUNDSSON', note: 'scritto "Gudmusson"' },
    { n: 'FRENDRUP', note: 'scritto "Frendup"' },
    { n: 'BRITSCHGI' }, { n: 'SULEMANA' }, { n: 'TAYLOR' },
    { n: 'MALEN' }, { n: 'SOMMER' }
  ],
  'Hellas Madonna': [
    { n: 'DAVID', note: 'scritto "Davids", corretto' },
    { n: 'CHEDDIRA' }, { n: 'ZARAGOZA' },
    { n: 'EKKELENKAMP', note: 'scritto "Ekkelekamp"' },
    { n: 'BRESCIANINI' }, { n: 'BELGHALI' }, { n: 'SCAMMACCA' }
  ],
  'FICA': [
    { n: 'LUIS HENRIQUE' }, { n: 'FULLKRUG' }
  ],
  'Lino Banfield FC': [
    { n: 'CELIK' }, { n: 'OBERT' }, { n: 'MARCANDALLI' },
    { n: 'BERNASCONI' }, { n: 'BOWIE' }, { n: 'CAPRILE' },
    { n: 'CAMBIAGHI' }, { n: 'VAZ' }, { n: 'BALDANZI' },
    { n: 'KOOPMEINERS', note: 'da Minnesota (scambio 13/02)' },
    { n: 'TAVARES', note: 'da Minnesota (scambio 13/02)' },
    { n: 'MAZZITELLI', note: 'da Minnesota (scambio 13/02)' }
  ],
  'MINNESOTA AL MAX': [
    { n: 'MONTIPO' }, { n: 'MARIANUCCI', note: 'potrebbe essere Marinucci' },
    { n: 'CATALDI' },
    { n: 'FAGIOLI', note: 'da Lino (scambio 13/02)' },
    { n: 'MILLER', note: 'da Lino (scambio 13/02)' },
    { n: 'BAKOLA' }, { n: 'ADZIC' }, { n: 'RATKOV' },
    { n: 'BELLANOVA', note: 'da Lino (scambio 13/02)' }
  ],
  'Papaie Top Team': [
    { n: 'KOLASINAC' }, { n: 'HIEN', note: 'da Minnesota (acquisto 11/02)' },
    { n: 'PASALIC' }, { n: 'NICOLUSSI CAVIGLIA' },
    { n: 'SOLOMON' }, { n: 'VLAHOVIC' }
  ],
  'Legenda Aurea': [
    { n: 'NELSSON' }, { n: 'DOSSENA' }, { n: 'BARTESAGHI' },
    { n: 'GANDELMAN' }, { n: 'BARBIERI' }, { n: 'LEAO' }, { n: 'ZAPPA' }
  ]
};

// Also map FICA full name for Rose lookup
const TEAM_ALIASES = {
  'FICA': 'Federazione Italiana Calcio Amatoriale'
};

// ========== MAIN ==========
function processLeague(label, roseFile, dbFile, teamCols, requests) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`  ${label}`);
  console.log(`  Formula: Se Sp<=FVM: MAX(FVM/10, 1) | Se Sp>FVM: MAX(AVG(Sp,FVM)/10, 1)`);
  console.log(`  Fonti: Rose 17/02 (Spesa) + DB 06/02 (FVM Prop.)`);
  console.log(`${'='.repeat(80)}`);

  const rose = parseRose(roseFile);
  const fvmMap = parseFVM(dbFile);
  const sqMap = parseSquadre(dbFile, teamCols);

  console.log(`  Rose: ${rose.groups.length} squadre, ${rose.map.size} entries`);
  console.log(`  DB FVM: ${fvmMap.size} giocatori`);
  console.log(`  SQUADRE (06/02): ${sqMap.size} giocatori\n`);

  let grandTotal = 0;
  let grandCount = 0;

  for (const [team, playerList] of Object.entries(requests)) {
    console.log(`\n${'─'.repeat(70)}`);
    console.log(`  ${team.toUpperCase()}`);
    console.log(`${'─'.repeat(70)}`);

    let teamTotal = 0;
    let teamCount = 0;

    for (const req of playerList) {
      const pName = req.n;
      const noteStr = req.note ? ` (${req.note})` : '';

      // REJECTED
      if (req.note && req.note.includes('RESPINTO')) {
        console.log(`  RESPINTO  ${pName.padEnd(24)} ${req.note}`);
        continue;
      }

      // Look up Spesa from Rose (try team-specific first)
      const roseTeam = TEAM_ALIASES[team] || team;
      const roseData = findInMap(rose.map, pName, roseTeam);

      // Look up FVM from DB
      const fvmData = findInMap(fvmMap, pName, null);

      // Look up insurance status from SQUADRE (06/02)
      const sqData = findInMap(sqMap, pName, null);

      if (!roseData && !fvmData) {
        console.log(`  MANCANTE  ${pName.padEnd(24)} Non trovato ne' in Rose ne' nel DB${noteStr}`);
        continue;
      }

      const spesa = roseData ? roseData.spesa : null;
      const fvm = fvmData ? fvmData.fvm : null;

      if (spesa === null && fvm !== null && fvm <= 10) {
        // With FVM<=10, cost is always 1 (minimum) regardless of Spesa
        const forcedCost = 1;
        teamTotal += forcedCost;
        teamCount++;
        console.log(`     1.0 cr  ${pName.padEnd(24)} Sp=?     FVM=${String(fvm).padEnd(5)} Costo=1.0 (minimo, FVM<=10 -> MAX(FVM/10,1)=1)${noteStr}`);
        continue;
      }
      if (spesa === null) {
        console.log(`  NO_SPESA  ${pName.padEnd(24)} FVM=${fvm}  Spesa=? (non in Rose 17/02)${noteStr}`);
        continue;
      }
      if (fvm === null) {
        console.log(`  NO_FVM    ${pName.padEnd(24)} Spesa=${spesa}  FVM=? (non nel DB)${noteStr}`);
        continue;
      }

      const cost = calcCost(spesa, fvm);
      teamTotal += cost;
      teamCount++;

      const formulaType = spesa <= fvm ? 'FVM/10' : 'AVG/10';
      let detail;
      if (spesa <= fvm) {
        detail = `${fvm}/10=${(fvm/10).toFixed(1)}`;
        if (fvm/10 < 1) detail += '->1';
      } else {
        detail = `(${spesa}+${fvm})/2/10=${((spesa+fvm)/2/10).toFixed(1)}`;
        if ((spesa+fvm)/2/10 < 1) detail += '->1';
      }

      const roseTeamStr = roseData.team !== team && roseData.team !== roseTeam
        ? ` [rose: ${roseData.team}]` : '';
      const insStr = sqData && sqData.ins === 'A' ? ' [ASS]' : '';
      const nameMatch = norm(pName) !== norm(roseData.name) ? ` ~${roseData.name}` : '';

      console.log(`  ${cost.toFixed(1).padStart(6)} cr  ${pName.padEnd(24)} Sp=${String(spesa).padEnd(5)} FVM=${String(fvm).padEnd(5)} ${formulaType}: ${detail}${roseTeamStr}${insStr}${nameMatch}${noteStr}`);
    }

    console.log(`  ${'- '.repeat(35)}`);
    console.log(`  SUBTOTALE ${team}: ${teamTotal.toFixed(1)} crediti (${teamCount} giocatori)\n`);
    grandTotal += teamTotal;
    grandCount += teamCount;
  }

  console.log(`${'═'.repeat(70)}`);
  console.log(`  TOTALE ${label}: ${grandTotal.toFixed(1)} crediti per ${grandCount} giocatori`);
  console.log(`${'═'.repeat(70)}\n`);
}

console.log('╔══════════════════════════════════════════════════════════════════════════════╗');
console.log('║   CALCOLO DEFINITIVO COSTI ASSICURATIVI - MERCATO INVERNALE 2026           ║');
console.log('╚══════════════════════════════════════════════════════════════════════════════╝');

processLeague(
  'FANTA TOSTI 2026',
  path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
  path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'),
  FT_COLS, FT_REQ
);

processLeague(
  'FANTAMANTRA MANAGERIALE',
  path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
  path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
  FM_COLS, FM_REQ
);
