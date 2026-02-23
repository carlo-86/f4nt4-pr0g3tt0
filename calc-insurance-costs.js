const XLSX = require('xlsx');
const path = require('path');

const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

// ========== TEAM CONFIGURATION ==========
// [teamName, calciatoreColIndex]
const FT_TEAMS = [
  ['FCK Deportivo', 2], ['Hellas Madonna', 14], ['muttley superstar', 26],
  ['PARTIZAN', 38], ['Legenda Aurea', 50], ['Kung Fu Pandev', 62],
  ['Millwall', 74], ['FC CKC 26', 86], ['Papaie Top Team', 98], ['Tronzano', 110]
];

const FM_TEAMS = [
  ['Papaie Top Team', 3], ['Legenda Aurea', 16], ['Lino Banfield FC', 29],
  ['Kung Fu Pandev', 42], ['FICA', 55], ['Hellas Madonna', 68],
  ['MINNESOTA AL MAX', 81], ['FC CKC 26', 94], ['H-Q-A Barcelona', 107], ['Mastri Birrai', 120]
];

// ========== INSURANCE REQUESTS FROM COMMUNICATIONS ==========
// FT - Comunicazione 13-14/02/2026
const FT_REQUESTS = {
  'Hellas Madonna': [
    { name: 'SPORTIELLO' }, { name: 'CIRCATI' }, { name: 'BERISHA' },
    { name: 'MOREO' }, { name: 'DUROSINMI', note: 'scritto "Duronisimi"' }
  ],
  'PARTIZAN': [
    { name: 'BELGHALI' }, { name: 'STREFEZZA' }, { name: 'PRZYBOREK' }
  ],
  'Kung Fu Pandev': [
    { name: 'MALEN' }, { name: 'VERGARA' },
    { name: 'BEUKEMA', note: 'assicurabile preventivamente (triennale)' },
    { name: 'KOUAME' }
  ],
  'FC CKC 26': [
    { name: 'TIAGO GABRIEL' }, { name: 'VAZ' }, { name: 'MUHAREMOVIC' },
    { name: 'BALDANZI' }, { name: 'ALLISON', note: 'scritto "Allison S."' },
    { name: 'BIJLOW' }, { name: 'BERNASCONI' }, { name: 'KONE', note: 'scritto "Kone I."' }
  ],
  'muttley superstar': [
    { name: 'OSTIGARD' }, { name: 'LUIS ENRIQUE' }, { name: 'SOLOMON' }
  ],
  'Millwall': [
    { name: 'MURIC' }, { name: 'CELIK' }, { name: 'RATKOV' }, { name: 'ZARAGOZA' },
    { name: 'PERRONE' }, { name: 'PALEARI' }, { name: 'BOGA' }, { name: 'HOLM' }
  ],
  'Papaie Top Team': [
    { name: 'HIEN' }
  ],
  'Legenda Aurea': [
    { name: 'DI GREGORIO' }, { name: 'SOMMER' }, { name: 'MARTINEZ' },
    { name: 'KALULU' }, { name: 'BARTESAGHI' }, { name: 'LOVRIC' },
    { name: 'TAYLOR' }, { name: 'FAGIOLI' }, { name: 'EKKELENKAMP' },
    { name: 'MIRETTI' }, { name: 'BONAZZOLI' }, { name: 'RASPADORI' }, { name: 'VITINHA' }
  ]
};

// FM - Comunicazione 13-14/02/2026
const FM_REQUESTS = {
  'Kung Fu Pandev': [
    { name: 'KONE', note: 'scritto "Kone"' },
    { name: 'RASPADORI' },
    { name: 'POSCH', note: 'RESPINTO - svincolato da KFP, non assicurabile' },
    { name: 'FERGUSON' }, { name: 'KOUAME' }
  ],
  'FC CKC 26': [
    { name: 'DUROSINMI' }, { name: 'VERGARA' }, { name: 'ZANIOLO' }
  ],
  'H-Q-A Barcelona': [
    { name: 'HOLM' }, { name: 'NDICKA' }, { name: 'GALLO' }, { name: 'VASQUEZ' },
    { name: 'GUDMUNDSSON', note: 'scritto "Gudmusson"' },
    { name: 'FRENDRUP', note: 'scritto "Frendup"' },
    { name: 'BRITSCHGI' }, { name: 'SULEMANA' }, { name: 'TAYLOR' },
    { name: 'MALEN' }, { name: 'SOMMER' }
  ],
  'Hellas Madonna': [
    { name: 'DAVID', note: 'scritto "Davids", corretto in David - trasf. da Mastri a Minnesota (04/02)' },
    { name: 'CHEDDIRA' }, { name: 'ZARAGOZA' },
    { name: 'EKKELENKAMP', note: 'scritto "Ekkelekamp"' },
    { name: 'BRESCIANINI' }, { name: 'BELGHALI' }, { name: 'SCAMMACCA' }
  ],
  'FICA': [
    { name: 'LUIS HENRIQUE' }, { name: 'FULLKRUG' }
  ],
  'Lino Banfield FC': [
    { name: 'CELIK' }, { name: 'OBERT' }, { name: 'MARCANDALLI' },
    { name: 'BERNASCONI' }, { name: 'BOWIE' }, { name: 'CAPRILE' },
    { name: 'CAMBIAGHI' }, { name: 'VAZ' }, { name: 'BALDANZI' },
    { name: 'KOOPMEINERS', note: 'da Minnesota (scambio 13/02)' },
    { name: 'TAVARES', note: 'da Minnesota (scambio 13/02)' },
    { name: 'MAZZITELLI', note: 'da Minnesota (scambio 13/02)' }
  ],
  'MINNESOTA AL MAX': [
    { name: 'MONTIPO' },
    { name: 'MARIANUCCI', note: 'potrebbe essere MARINUCCI' },
    { name: 'CATALDI' },
    { name: 'FAGIOLI', note: 'da Lino Banfield (scambio 13/02)' },
    { name: 'MILLER', note: 'da Lino Banfield (scambio 13/02)' },
    { name: 'BAKOLA' }, { name: 'ADZIC' }, { name: 'RATKOV' },
    { name: 'BELLANOVA', note: 'da Lino Banfield (scambio 13/02)' }
  ],
  'Papaie Top Team': [
    { name: 'KOLASINAC' },
    { name: 'HIEN', note: 'da Minnesota (acquisto 11/02 per 22cr)' },
    { name: 'PASALIC' }, { name: 'NICOLUSSI CAVIGLIA' },
    { name: 'SOLOMON' }, { name: 'VLAHOVIC' }
  ],
  'Legenda Aurea': [
    { name: 'NELSSON' }, { name: 'DOSSENA' }, { name: 'BARTESAGHI' },
    { name: 'GANDELMAN' }, { name: 'BARBIERI' }, { name: 'LEAO' }, { name: 'ZAPPA' }
  ]
};

// ========== NAME MATCHING UTILITIES ==========

function normalize(name) {
  return String(name || '').trim().toUpperCase()
    .replace(/[''`\u2019\u2018\u0300\u0301]/g, '')
    .replace(/\s+/g, ' ')
    .replace(/\./g, '')
    .trim();
}

function findInMap(map, searchName) {
  const key = normalize(searchName);

  // 1. Exact match
  for (const [k, v] of map) {
    if (normalize(k) === key) return { ...v, matchedAs: k };
  }

  // 2. Key starts with search or vice versa (handles suffixes like "I.", "S.", "N.", "L.")
  for (const [k, v] of map) {
    const nk = normalize(k);
    if (nk.startsWith(key + ' ') || key.startsWith(nk + ' ')) return { ...v, matchedAs: k };
  }

  // 3. One contains the other (min 4 chars for safety)
  if (key.length >= 4) {
    for (const [k, v] of map) {
      const nk = normalize(k);
      if (nk.length >= 4 && (nk.includes(key) || key.includes(nk))) return { ...v, matchedAs: k };
    }
  }

  // 4. First word match (for compound names, min 5 chars)
  const firstWord = key.split(' ')[0];
  if (firstWord.length >= 5) {
    for (const [k, v] of map) {
      const fk = normalize(k).split(' ')[0];
      if (fk === firstWord) return { ...v, matchedAs: k };
    }
  }

  // 5. Levenshtein distance 1-2 for names >= 5 chars
  if (key.length >= 5) {
    for (const [k, v] of map) {
      const nk = normalize(k);
      if (nk.length >= 5 && levenshtein(nk, key) <= 2) return { ...v, matchedAs: k };
    }
  }

  return null;
}

function levenshtein(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,
        dp[i][j - 1] + 1,
        dp[i - 1][j - 1] + (a[i - 1] !== b[j - 1] ? 1 : 0)
      );
    }
  }
  return dp[m][n];
}

// ========== DATA EXTRACTION ==========

function parseSquadre(filePath, teams) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['SQUADRE'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Map: normalizedName -> { name, team, spesa, ins }
  const players = new Map();
  const skipLabels = /^(Calciatore|Ass\.?\s*=|Totali|PORTIERI|DIFENSORI|CENTROCAMPISTI|ATTACCANTI|P|D|C|A)$/i;

  for (const [teamName, col] of teams) {
    for (let r = 5; r < 52; r++) {
      const raw = data[r]?.[col];
      const name = String(raw || '').trim();
      if (!name || skipLabels.test(name) || name.length < 2) continue;

      const spesa = Number(data[r]?.[col + 10]) || 0;
      const ins = String(data[r]?.[col + 3] || '').trim();

      const key = normalize(name);
      if (key && !players.has(key)) {
        players.set(key, { name, team: teamName, spesa, ins });
      }
    }
  }

  return players;
}

function parseFVM(filePath) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['DB'];
  if (!ws) {
    console.log('  ATTENZIONE: Foglio DB non trovato!');
    return new Map();
  }
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Verify columns: BK=62 should have player names, CB=79 should have FVM Prop.
  // Row 2 (index 2) is likely the header row
  const hdrRow = 2; // 0-indexed, so this is row 3 in Excel
  console.log(`  DB headers (row ${hdrRow + 1}): col62="${data[hdrRow]?.[62]}", col79="${data[hdrRow]?.[79]}"`);

  // Also check a few sample data rows
  console.log(`  DB sample row 3: col62="${data[3]?.[62]}", col79="${data[3]?.[79]}"`);
  console.log(`  DB sample row 4: col62="${data[4]?.[62]}", col79="${data[4]?.[79]}"`);
  console.log(`  DB sample row 5: col62="${data[5]?.[62]}", col79="${data[5]?.[79]}"`);

  const fvmMap = new Map();
  for (let r = 2; r < Math.min(data.length, 1000); r++) {
    const name = String(data[r]?.[62] || '').trim();
    const fvm = data[r]?.[79];

    if (name && name.length > 1) {
      const key = normalize(name);
      const val = Number(fvm) || 0;
      if (!fvmMap.has(key)) {
        fvmMap.set(key, { name, fvm: val });
      }
    }
  }

  return fvmMap;
}

// ========== COST FORMULA ==========
// From ROSA sheet column E:
// FT: IF($K=0,"",IF($M<=$N,MAX($N/10,1),MAX(AVERAGE($M,$N)/10,1)))
// FM: IF($K=0,"",IF($N<=$O,MAX($O/10,1),MAX(AVERAGE($N,$O)/10,1)))
// Where: M/N(FT) or N/O(FM) = Spesa/FVM Prop.
//
// Spesa = asta price (from SQUADRE col+10)
// FVM Prop. = current FVM (from DB!BK:CB col 18 = col CB)
//
// If Spesa <= FVM: cost = MAX(FVM/10, 1)
// If Spesa >  FVM: cost = MAX(AVG(Spesa,FVM)/10, 1)

function calcCost(spesa, fvm) {
  if (spesa <= fvm) {
    return Math.max(fvm / 10, 1);
  } else {
    return Math.max((spesa + fvm) / 2 / 10, 1);
  }
}

// ========== MAIN PROCESSOR ==========

function processLeague(label, dbFilePath, teams, requests) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`  ${label}`);
  console.log(`${'='.repeat(80)}\n`);

  console.log('Caricamento SQUADRE...');
  const squadre = parseSquadre(dbFilePath, teams);
  console.log(`  ${squadre.size} giocatori trovati in SQUADRE\n`);

  console.log('Caricamento FVM da foglio DB...');
  const fvmMap = parseFVM(dbFilePath);
  console.log(`  ${fvmMap.size} giocatori con FVM nel DB\n`);

  let grandTotal = 0;
  let grandCount = 0;
  const allResults = {};

  for (const [team, playerList] of Object.entries(requests)) {
    console.log(`\n${'─'.repeat(70)}`);
    console.log(`  ${team.toUpperCase()}`);
    console.log(`${'─'.repeat(70)}`);

    let teamTotal = 0;
    let teamCount = 0;

    for (const req of playerList) {
      const pName = req.name;
      const noteStr = req.note ? ` (${req.note})` : '';

      // Special: rejected requests
      if (req.note && req.note.includes('RESPINTO')) {
        console.log(`  RESPINTO   ${pName.padEnd(24)} ${req.note}`);
        continue;
      }

      // Find in SQUADRE (global search across all teams in the league)
      const sqData = findInMap(squadre, pName);
      // Find FVM in DB sheet
      const fvmData = findInMap(fvmMap, pName);

      if (!sqData && !fvmData) {
        console.log(`  MANCANTE   ${pName.padEnd(24)} Non trovato in SQUADRE ne' nel DB${noteStr}`);
        continue;
      }

      if (!sqData) {
        console.log(`  NO_SPESA   ${pName.padEnd(24)} FVM=${fvmData.fvm}  Spesa=? (non in SQUADRE al 06/02)${noteStr}`);
        continue;
      }

      if (!fvmData) {
        console.log(`  NO_FVM     ${pName.padEnd(24)} Spesa=${sqData.spesa}  FVM=? (non nel DB)${noteStr}`);
        continue;
      }

      const spesa = sqData.spesa;
      const fvm = fvmData.fvm;
      const cost = calcCost(spesa, fvm);
      const matchedSQ = sqData.matchedAs || sqData.name;
      const matchedFVM = fvmData.matchedAs || fvmData.name;

      teamTotal += cost;
      teamCount++;

      const formulaType = spesa <= fvm ? 'FVM/10' : 'AVG/10';
      let detail;
      if (spesa <= fvm) {
        detail = `${fvm}/10 = ${(fvm / 10).toFixed(1)}`;
        if (fvm / 10 < 1) detail += ' -> MAX 1';
      } else {
        detail = `(${spesa}+${fvm})/2/10 = ${((spesa + fvm) / 2 / 10).toFixed(1)}`;
        if ((spesa + fvm) / 2 / 10 < 1) detail += ' -> MAX 1';
      }

      const fromTeamStr = sqData.team !== team ? ` [rosa ${sqData.team}]` : '';
      const insStr = sqData.ins === 'A' ? ' [GIA ASS.]' : '';
      const nameMatch = normalize(pName) !== normalize(matchedSQ) ? ` ~${matchedSQ}` : '';

      console.log(`  ${cost.toFixed(1).padStart(6)} cr  ${pName.padEnd(24)} Sp=${String(spesa).padEnd(5)} FVM=${String(fvm).padEnd(6)} ${formulaType}: ${detail}${fromTeamStr}${insStr}${nameMatch}${noteStr}`);
    }

    console.log(`  ${'- '.repeat(35)}`);
    console.log(`  SUBTOTALE ${team}: ${teamTotal.toFixed(1)} crediti (${teamCount} giocatori calcolati)\n`);

    grandTotal += teamTotal;
    grandCount += teamCount;
    allResults[team] = { total: teamTotal, count: teamCount };
  }

  console.log(`\n${'='.repeat(70)}`);
  console.log(`  RIEPILOGO ${label}`);
  console.log(`${'='.repeat(70)}`);
  for (const [team, res] of Object.entries(allResults)) {
    console.log(`  ${team.padEnd(30)} ${res.total.toFixed(1).padStart(7)} cr  (${res.count} giocatori)`);
  }
  console.log(`  ${'─'.repeat(60)}`);
  console.log(`  TOTALE: ${grandTotal.toFixed(1)} crediti per ${grandCount} giocatori`);
  console.log(`${'='.repeat(70)}\n`);

  return allResults;
}

// ========== EXECUTE ==========

console.log('╔══════════════════════════════════════════════════════════════════════════════╗');
console.log('║    CALCOLO COSTI ASSICURATIVI - SESSIONE MERCATO INVERNALE 2026             ║');
console.log('║    Formula: Se Sp<=FVM: MAX(FVM/10, 1)                                     ║');
console.log('║             Se Sp> FVM: MAX(AVG(Sp,FVM)/10, 1)                              ║');
console.log('║    Fonte dati: DB Excel 06/02/2026 (fogli SQUADRE + DB)                     ║');
console.log('╚══════════════════════════════════════════════════════════════════════════════╝');

const ftRes = processLeague(
  'FANTA TOSTI 2026',
  path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'),
  FT_TEAMS,
  FT_REQUESTS
);

const fmRes = processLeague(
  'FANTAMANTRA MANAGERIALE',
  path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
  FM_TEAMS,
  FM_REQUESTS
);
