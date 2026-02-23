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

  // Find team names in row 4 (index 4)
  const teamRow = data[4] || [];
  const teams = [];
  for (let c = 0; c < teamRow.length; c++) {
    const v = String(teamRow[c] || '').trim();
    if (v && v !== 'Crediti Residui:' && !v.match(/^\d+$/)) {
      teams.push({ name: v, col: c });
    }
  }

  console.log(`\nTeam trovati (${teams.length}):`);
  teams.forEach(t => console.log(`  Col ${t.col}: ${t.name}`));

  // Parse each team's players
  // Each team block: col+0=Ruolo, col+1=Calciatore, col+2=Squadra, col+3=Costo
  const allPlayers = new Map(); // normalized name -> { name, team, spesa }

  for (const team of teams) {
    const calcCol = team.col + 1;
    const costoCol = team.col + 3;
    let count = 0;

    for (let r = 6; r < data.length; r++) {
      const name = String(data[r]?.[calcCol] || '').trim();
      const costo = data[r]?.[costoCol];
      const ruolo = String(data[r]?.[team.col] || '').trim();

      if (!name || name === 'Calciatore') continue;
      // Skip if it looks like a credit residue note
      if (name.includes('Crediti') || name.includes('residui')) continue;

      const key = name.toUpperCase().replace(/[''`\u2019\u2018\u0300\u0301]/g, '').replace(/\s+/g, ' ').trim();
      const spesa = Number(costo) || 0;

      if (key && key.length > 1) {
        // Store with team context
        allPlayers.set(`${key}|${team.name}`, { name, team: team.name, spesa, ruolo });
        // Also store without team for global lookup (first occurrence wins)
        if (!allPlayers.has(key)) {
          allPlayers.set(key, { name, team: team.name, spesa, ruolo });
        }
        count++;
      }
    }
    console.log(`  ${team.name}: ${count} giocatori`);
  }

  return { teams, allPlayers, data };
}

// Parse both Rose files
const ftRose = parseRoseFile(
  path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
  'FT Rose 17/02/2026'
);

const fmRose = parseRoseFile(
  path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
  'FM Rose 17/02/2026'
);

// Now search for specific missing players from our insurance requests
console.log('\n\n=== FT: RICERCA GIOCATORI PER ASSICURAZIONI ===');
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
    // First try team-specific lookup
    const teamKey = `${pName}|${team}`;
    let found = ftRose.allPlayers.get(teamKey);

    if (!found) {
      // Try normalized team name variations
      for (const [k, v] of ftRose.allPlayers) {
        if (k.includes('|') && k.startsWith(pName)) {
          found = v;
          break;
        }
      }
    }

    if (!found) {
      // Try global lookup
      found = ftRose.allPlayers.get(pName);
    }

    if (!found) {
      // Fuzzy: first 4+ chars
      for (const [k, v] of ftRose.allPlayers) {
        if (k.includes('|')) continue;
        if (k.length >= 4 && pName.length >= 4) {
          if (k.startsWith(pName) || pName.startsWith(k)) {
            found = v;
            break;
          }
        }
      }
    }

    if (found) {
      console.log(`  ${pName.padEnd(24)} Team: ${found.team.padEnd(20)} Spesa: ${found.spesa}`);
    } else {
      console.log(`  ${pName.padEnd(24)} NON TROVATO`);
    }
  }
}

console.log('\n\n=== FM: RICERCA GIOCATORI PER ASSICURAZIONI ===');
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
    const teamKey = `${pName}|${team}`;
    let found = fmRose.allPlayers.get(teamKey);

    if (!found) {
      for (const [k, v] of fmRose.allPlayers) {
        if (k.includes('|') && k.startsWith(pName)) {
          found = v;
          break;
        }
      }
    }

    if (!found) {
      found = fmRose.allPlayers.get(pName);
    }

    if (!found) {
      for (const [k, v] of fmRose.allPlayers) {
        if (k.includes('|')) continue;
        if (k.length >= 4 && pName.length >= 4) {
          if (k.startsWith(pName) || pName.startsWith(k)) {
            found = v;
            break;
          }
        }
      }
    }

    if (found) {
      console.log(`  ${pName.padEnd(24)} Team: ${found.team.padEnd(25)} Spesa: ${found.spesa}`);
    } else {
      console.log(`  ${pName.padEnd(24)} NON TROVATO`);
    }
  }
}
