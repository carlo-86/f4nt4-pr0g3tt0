const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function parseRoseFile(filePath, label) {
  console.log(`\n${'='.repeat(70)}`);
  console.log(`  ${label}`);
  console.log(`${'='.repeat(70)}`);

  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['TutteLeRose'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  console.log(`Total rows: ${data.length}`);

  // The Rose file has 2 column groups: cols 0-3 and cols 5-8
  // Each column group contains 5 teams stacked vertically
  // Team names appear in the "Ruolo" column position (0 or 5) as non-role values
  // Roles: P, D, C, A (Classic) or Por, Dc, Dd, Ds, E, B, M, C, T, W, Pc, A (Mantra)
  const roles = new Set(['P', 'D', 'C', 'A', 'Por', 'Dc', 'Dd', 'Ds', 'E', 'B', 'M', 'T', 'W', 'Pc']);

  function isRole(val) {
    if (!val) return false;
    const v = String(val).trim();
    // Could be compound roles like "B;Dd;E"
    const parts = v.split(';');
    return parts.every(p => roles.has(p.trim()));
  }

  // Parse a column group (starting at baseCol)
  function parseColumnGroup(baseCol) {
    const result = []; // { team, players: [{ name, spesa, ruolo }] }
    let currentTeam = null;
    let currentPlayers = [];

    for (let r = 4; r < data.length; r++) {
      const col0 = String(data[r]?.[baseCol] || '').trim();
      const col1 = String(data[r]?.[baseCol + 1] || '').trim();
      const col3 = data[r]?.[baseCol + 3];

      // Skip row 5 (header row with "Ruolo", "Calciatore", etc.)
      if (col0 === 'Ruolo' || col1 === 'Calciatore') continue;

      // Check if this is a team name row
      // Team name appears in col0 when it's NOT a role and NOT empty
      // and col1 is either empty or also looks like a label
      if (col0 && !isRole(col0) && col0 !== 'Ruolo' && col0.length > 1 && !col0.includes('Crediti')) {
        // Could be a team name
        // Verify: col1 should be empty or the row should not have a valid player pattern
        // Actually, team names also have col1 empty or same-row as "Crediti Residui"
        const looksLikeTeam = !col1 || col1 === '' || col0.includes('Crediti');

        if (looksLikeTeam || (col1 && !isRole(col0))) {
          // If col0 is not a role and has a non-empty substantial string, it's likely a team name
          // But we need to be careful: some compound values like "B;Dd;E" are roles
          if (!col0.includes(';') && !isRole(col0)) {
            // Save previous team
            if (currentTeam) {
              result.push({ team: currentTeam, players: currentPlayers });
            }
            currentTeam = col0;
            currentPlayers = [];
            continue;
          }
        }
      }

      // Check for "Crediti Residui" marker
      if (col0.includes('Crediti') || col1.includes('Crediti')) continue;

      // If we have a current team and this looks like a player row
      if (currentTeam && col1 && col1 !== 'Calciatore' && isRole(col0)) {
        currentPlayers.push({
          name: col1,
          spesa: Number(col3) || 0,
          ruolo: col0
        });
      }
    }

    // Don't forget the last team
    if (currentTeam) {
      result.push({ team: currentTeam, players: currentPlayers });
    }

    return result;
  }

  const group1 = parseColumnGroup(0);
  const group2 = parseColumnGroup(5);
  const allGroups = [...group1, ...group2];

  console.log(`\nTeam trovati: ${allGroups.length}`);
  for (const g of allGroups) {
    console.log(`  ${g.team.padEnd(30)} ${g.players.length} giocatori`);
  }

  // Build player lookup: Map<normalizedName + "|" + teamName, { name, team, spesa }>
  // Also Map<normalizedName, { name, team, spesa }> for global fallback
  const playerMap = new Map();

  for (const g of allGroups) {
    for (const p of g.players) {
      const key = p.name.toUpperCase().replace(/[''`\u2019\u2018]/g, '').replace(/\s+/g, ' ').trim();
      const teamKey = `${key}|${g.team}`;
      playerMap.set(teamKey, { name: p.name, team: g.team, spesa: p.spesa, ruolo: p.ruolo });
      if (!playerMap.has(key)) {
        playerMap.set(key, { name: p.name, team: g.team, spesa: p.spesa, ruolo: p.ruolo });
      }
    }
  }

  return { groups: allGroups, playerMap };
}

function normalize(name) {
  return String(name || '').trim().toUpperCase()
    .replace(/[''`\u2019\u2018\u0300\u0301]/g, '')
    .replace(/\./g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function findPlayer(playerMap, pName, team) {
  const key = normalize(pName);

  // 1. Exact match with team
  for (const [k, v] of playerMap) {
    if (!k.includes('|')) continue;
    const [nameKey, teamKey] = k.split('|');
    if (normalize(nameKey) === key && teamKey === team) return v;
  }

  // 2. Starts-with match with team (for names like "MARTINEZ JO" matching "MARTINEZ")
  for (const [k, v] of playerMap) {
    if (!k.includes('|')) continue;
    const [nameKey, teamKey] = k.split('|');
    const nk = normalize(nameKey);
    if (teamKey === team && (nk.startsWith(key) || key.startsWith(nk)) && Math.min(nk.length, key.length) >= 4) return v;
  }

  // 3. Global exact match (any team - for traded players)
  for (const [k, v] of playerMap) {
    if (k.includes('|')) continue;
    if (normalize(k) === key) return v;
  }

  // 4. Global starts-with
  for (const [k, v] of playerMap) {
    if (k.includes('|')) continue;
    const nk = normalize(k);
    if ((nk.startsWith(key) || key.startsWith(nk)) && Math.min(nk.length, key.length) >= 4) return v;
  }

  // 5. Contains match
  if (key.length >= 5) {
    for (const [k, v] of playerMap) {
      if (k.includes('|')) continue;
      const nk = normalize(k);
      if (nk.length >= 5 && (nk.includes(key) || key.includes(nk))) return v;
    }
  }

  return null;
}

// ========== Parse both Rose files ==========
const ftRose = parseRoseFile(
  path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
  'FT Rose 17/02/2026'
);

const fmRose = parseRoseFile(
  path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
  'FM Rose 17/02/2026'
);

// ========== Search for all insurance-requested players ==========
console.log('\n\n╔═══════════════════════════════════════════════════════════════╗');
console.log('║   FT: SPESA DA ROSE (17/02) PER GIOCATORI ASSICURAZIONI    ║');
console.log('╚═══════════════════════════════════════════════════════════════╝');

const ftSearch = [
  ['Hellas Madonna', ['SPORTIELLO', 'CIRCATI', 'BERISHA', 'MOREO', 'DUROSINMI']],
  ['PARTIZAN', ['BELGHALI', 'STREFEZZA', 'PRZYBOREK']],
  ['Kung Fu Pandev', ['MALEN', 'VERGARA', 'BEUKEMA', 'KOUAME']],
  ['FC CKC 26', ['TIAGO GABRIEL', 'VAZ', 'MUHAREMOVIC', 'BALDANZI', 'ALLISON', 'BIJLOW', 'BERNASCONI', 'KONE']],
  ['muttley superstar', ['OSTIGARD', 'LUIS ENRIQUE', 'SOLOMON']],
  ['Millwall', ['MURIC', 'CELIK', 'RATKOV', 'ZARAGOZA', 'PERRONE', 'PALEARI', 'BOGA', 'HOLM']],
  ['Papaie Top Team', ['HIEN']],
  ['Legenda Aurea', ['DI GREGORIO', 'SOMMER', 'MARTINEZ', 'KALULU', 'BARTESAGHI', 'LOVRIC', 'TAYLOR', 'FAGIOLI', 'EKKELENKAMP', 'MIRETTI', 'BONAZZOLI', 'RASPADORI', 'VITINHA']]
];

for (const [team, players] of ftSearch) {
  console.log(`\n--- ${team} ---`);
  for (const pName of players) {
    const found = findPlayer(ftRose.playerMap, pName, team);
    if (found) {
      const matchFlag = found.team !== team ? ` [trovato in ${found.team}]` : '';
      console.log(`  ${pName.padEnd(24)} Sp=${String(found.spesa).padEnd(4)} (${found.team})${matchFlag}`);
    } else {
      console.log(`  ${pName.padEnd(24)} NON TROVATO`);
    }
  }
}

console.log('\n\n╔═══════════════════════════════════════════════════════════════╗');
console.log('║   FM: SPESA DA ROSE (17/02) PER GIOCATORI ASSICURAZIONI    ║');
console.log('╚═══════════════════════════════════════════════════════════════╝');

const fmSearch = [
  ['Kung Fu Pandev', ['KONE', 'RASPADORI', 'FERGUSON', 'KOUAME']],
  ['FC CKC 26', ['DUROSINMI', 'VERGARA', 'ZANIOLO']],
  ['H-Q-A Barcelona', ['HOLM', 'NDICKA', 'GALLO', 'VASQUEZ', 'GUDMUNDSSON', 'FRENDRUP', 'BRITSCHGI', 'SULEMANA', 'TAYLOR', 'MALEN', 'SOMMER']],
  ['Hellas Madonna', ['DAVID', 'CHEDDIRA', 'ZARAGOZA', 'EKKELENKAMP', 'BRESCIANINI', 'BELGHALI', 'SCAMMACCA']],
  ['FICA', ['LUIS HENRIQUE', 'FULLKRUG']],
  ['Lino Banfield FC', ['CELIK', 'OBERT', 'MARCANDALLI', 'BERNASCONI', 'BOWIE', 'CAPRILE', 'CAMBIAGHI', 'VAZ', 'BALDANZI', 'KOOPMEINERS', 'TAVARES', 'MAZZITELLI']],
  ['MINNESOTA AL MAX', ['MONTIPO', 'MARIANUCCI', 'CATALDI', 'FAGIOLI', 'MILLER', 'BAKOLA', 'ADZIC', 'RATKOV', 'BELLANOVA']],
  ['Papaie Top Team', ['KOLASINAC', 'HIEN', 'PASALIC', 'NICOLUSSI CAVIGLIA', 'SOLOMON', 'VLAHOVIC']],
  ['Legenda Aurea', ['NELSSON', 'DOSSENA', 'BARTESAGHI', 'GANDELMAN', 'BARBIERI', 'LEAO', 'ZAPPA']]
];

for (const [team, players] of fmSearch) {
  console.log(`\n--- ${team} ---`);
  for (const pName of players) {
    const found = findPlayer(fmRose.playerMap, pName, team);
    if (found) {
      const matchFlag = found.team !== team ? ` [trovato in ${found.team}]` : '';
      console.log(`  ${pName.padEnd(24)} Sp=${String(found.spesa).padEnd(4)} (${found.team})${matchFlag}`);
    } else {
      console.log(`  ${pName.padEnd(24)} NON TROVATO`);
    }
  }
}
