const XLSX = require('xlsx');
const path = require('path');

const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function parseSquadreAll(filePath) {
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['SQUADRE'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  const row3 = data[3] || [];
  const row4 = data[4] || [];

  // Find team blocks by "Calciatore" in row4
  const teamBlocks = [];
  for (let col = 0; col < row4.length; col++) {
    if (String(row4[col]).trim() === 'Calciatore') {
      let teamName = String(row3[col] || '').trim();
      if (!teamName || teamName.match(/^\d+$/) || teamName.includes('Numero') || teamName.includes('Crediti')) {
        teamName = String(row3[col - 1] || '').trim();
      }
      teamBlocks.push({ col, teamName });
    }
  }

  const allPlayers = [];

  for (const block of teamBlocks) {
    const { col: startCol, teamName } = block;
    if (!teamName || teamName === 'Valori medi rose') continue;

    // Data is in 4 sections (P/D/C/A) separated by empty rows + "Calciatore" headers
    // Each section gap is 3-5 empty rows. After all 4 sections, there are 10+ empty rows.
    // We stop after 8 consecutive non-player rows.
    let emptyCount = 0;
    for (let r = 5; r < Math.min(data.length, 100); r++) {
      const row = data[r];
      const playerName = String(row[startCol] || '').trim();

      // Skip headers and empty rows
      if (!playerName || playerName === 'Calciatore' || playerName === 'Ass. = A' || playerName === 'Totali') {
        emptyCount++;
        if (emptyCount >= 8) break;
        continue;
      }
      if (playerName.startsWith('Numero') || playerName.startsWith('Crediti')) {
        emptyCount++;
        if (emptyCount >= 8) break;
        continue;
      }
      if (playerName.length <= 1) { emptyCount++; continue; }

      emptyCount = 0;

      let buyDateStr = '';
      const buyDate = row[startCol + 4];
      if (typeof buyDate === 'number') {
        const d = XLSX.SSF.parse_date_code(buyDate);
        buyDateStr = `${String(d.d).padStart(2,'0')}/${String(d.m).padStart(2,'0')}/${d.y}`;
      }

      let insDateStr = '';
      const insDate = row[startCol + 7];
      if (typeof insDate === 'number') {
        const d = XLSX.SSF.parse_date_code(insDate);
        insDateStr = `${String(d.d).padStart(2,'0')}/${String(d.m).padStart(2,'0')}/${d.y}`;
      }

      allPlayers.push({
        fantaTeam: teamName,
        name: playerName,
        role: String(row[startCol + 1] || '').trim(),
        realTeam: String(row[startCol + 2] || '').trim(),
        insured: String(row[startCol + 3] || '').trim(),
        buyDate: buyDateStr,
        quoteBuy: row[startCol + 5],
        fvmPropBuy: row[startCol + 6],
        insDate: insDateStr,
        quoteRenew: row[startCol + 8],
        fvmPropRenew: row[startCol + 9],
        spesa: row[startCol + 10]
      });
    }
  }

  return allPlayers;
}

function findPlayer(allPlayers, teamSearch, playerSearch) {
  const norm = s => s.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z ]/g, '').trim();
  const pNorm = norm(playerSearch);
  const tNorm = norm(teamSearch);

  // Filter candidates by team
  let candidates = allPlayers.filter(p => {
    const tn = norm(p.fantaTeam);
    if (tn === tNorm) return true;
    if (tn.includes(tNorm) || tNorm.includes(tn)) return true;
    const tWords = tNorm.split(' ').filter(w => w.length > 2);
    const pWords = tn.split(' ').filter(w => w.length > 2);
    return tWords.some(tw => pWords.some(pw => pw === tw));
  });

  if (candidates.length === 0) candidates = allPlayers;

  // 1. Exact normalized match
  let found = candidates.find(p => norm(p.name) === pNorm);
  if (found) return found;

  // 2. Contains
  found = candidates.find(p => norm(p.name).includes(pNorm) || pNorm.includes(norm(p.name)));
  if (found) return found;

  // 3. Word matching
  const searchWords = pNorm.split(' ').filter(w => w.length > 1);
  for (const sw of searchWords) {
    if (sw.length >= 4) {
      found = candidates.find(p => norm(p.name).includes(sw));
      if (found) return found;
    }
  }

  // 4. First 5 chars
  if (pNorm.length >= 5) {
    found = candidates.find(p => norm(p.name).startsWith(pNorm.substring(0, 5)));
    if (found) return found;
  }

  // 5. Global fallback
  if (candidates !== allPlayers) {
    found = allPlayers.find(p => norm(p.name) === pNorm);
    if (found) return found;
    found = allPlayers.find(p => norm(p.name).includes(pNorm) || pNorm.includes(norm(p.name)));
    if (found) return found;
    for (const sw of searchWords) {
      if (sw.length >= 5) {
        found = allPlayers.find(p => norm(p.name).includes(sw));
        if (found) return found;
      }
    }
  }

  return null;
}

function formatPlayer(p) {
  const status = p.insured === 'A' ? 'GIA ASS.' : 'da assic.';
  const fvmAcq = p.fvmPropBuy || '-';
  const fvmRin = p.fvmPropRenew || '-';
  return `  [${status}] ${p.name.padEnd(22)} ${p.role.padEnd(8)} ${p.realTeam.padEnd(12)} Acq:${p.buyDate.padEnd(11)} Q.Acq:${String(p.quoteBuy).padEnd(5)} FVMp:${String(fvmAcq).padEnd(5)} Ins:${(p.insDate || '-').padEnd(11)} Q.Rin:${String(p.quoteRenew || '-').padEnd(5)} FVMpR:${String(fvmRin).padEnd(5)} Sp:${p.spesa}`;
}

// FT Insurance Requests
const ftRequests = {
  'Hellas Madonna': ['Sportiello', 'Circati', 'Berisha', 'Moreo', 'Durosinmi'],
  'PARTIZAN': ['Belghali', 'Strefezza', 'Przyborek'],
  'Kung Fu Pandev': ['Malen', 'Vergara', 'Beukema', 'Kouame'],
  'FC CKC 26': ['Tiago Gabriel', 'Vaz', 'Muharemovic', 'Baldanzi', 'Allison', 'Bijlow', 'Bernasconi', 'Kone I'],
  'muttley superstar': ['Ostigard', 'Luis Henrique', 'Solomon'],
  'Millwall': ['Muric', 'Celik', 'Ratkov', 'Zaragoza', 'Perrone', 'Paleari', 'Boga', 'Holm'],
  'Papaie Top Team': ['Hien'],
  'Legenda Aurea': ['Di Gregorio', 'Sommer', 'Martinez Jo', 'Kalulu', 'Bartesaghi', 'Lovric', 'Taylor', 'Fagioli', 'Ekkelenkamp', 'Miretti', 'Bonazzoli', 'Raspadori', 'Vitinha']
};

// FM Insurance Requests
const fmRequests = {
  'Kung Fu Pandev': ['Kone I', 'Raspadori', 'Posch', 'Ferguson', 'Kouame'],
  'FC CKC 26': ['Durosinmi', 'Vergara', 'Zaniolo'],
  'H-Q-A Barcelona': ['Holm', 'Ndicka', 'Gallo', 'Vasquez', 'Gudmundsson', 'Frendrup', 'Britschgi', 'Sulemana', 'Taylor', 'Malen', 'Sommer'],
  'Hellas Madonna': ['David', 'Cheddira', 'Zaragoza', 'Ekkelenkamp', 'Brescianini', 'Belghali', 'Scamacca'],
  'FICA': ['Luis Henrique', 'Fullkrug'],
  'Lino Banfield FC': ['Celik', 'Obert', 'Marcandalli', 'Bernasconi', 'Bowie', 'Caprile', 'Cambiaghi', 'Vaz', 'Baldanzi', 'Koopmeiners', 'Tavares', 'Mazzitelli'],
  'MINNESOTA AL MAX': ['Montipo', 'Marianucci', 'Cataldi', 'Fagioli', 'Miller', 'Bakola', 'Adzic', 'Ratkov', 'Bellanova'],
  'Papaie Top Team': ['Kolasinac', 'Hien', 'Pasalic', 'Nicolussi Caviglia', 'Solomon', 'Vlahovic'],
  'Legenda Aurea': ['Nelsson', 'Dossena', 'Bartesaghi', 'Gandelman', 'Barbieri', 'Leao', 'Zappa']
};

console.log('=== FANTA TOSTI - DATI ASSICURAZIONE PER GIOCATORE ===\n');
const ftPlayers = parseSquadreAll(path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'));
const ftTeamNames = [...new Set(ftPlayers.map(p => p.fantaTeam))];
console.log('Squadre FT:', ftTeamNames.join(', '));
console.log(`Totale giocatori: ${ftPlayers.length}`);
for (const tn of ftTeamNames) {
  console.log(`  ${tn}: ${ftPlayers.filter(p => p.fantaTeam === tn).length}`);
}

for (const [team, players] of Object.entries(ftRequests)) {
  console.log(`\n--- ${team} ---`);
  for (const pName of players) {
    const p = findPlayer(ftPlayers, team, pName);
    if (p) {
      console.log(formatPlayer(p));
    } else {
      console.log(`  *** NON TROVATO: ${pName}`);
    }
  }
}

console.log('\n\n=== FANTAMANTRA MANAGERIALE - DATI ASSICURAZIONE PER GIOCATORE ===\n');
const fmPlayers = parseSquadreAll(path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'));
const fmTeamNames = [...new Set(fmPlayers.map(p => p.fantaTeam))];
console.log('Squadre FM:', fmTeamNames.join(', '));
console.log(`Totale giocatori: ${fmPlayers.length}`);
for (const tn of fmTeamNames) {
  console.log(`  ${tn}: ${fmPlayers.filter(p => p.fantaTeam === tn).length}`);
}

for (const [team, players] of Object.entries(fmRequests)) {
  console.log(`\n--- ${team} ---`);
  for (const pName of players) {
    const p = findPlayer(fmPlayers, team, pName);
    if (p) {
      console.log(formatPlayer(p));
    } else {
      console.log(`  *** NON TROVATO: ${pName}`);
    }
  }
}
