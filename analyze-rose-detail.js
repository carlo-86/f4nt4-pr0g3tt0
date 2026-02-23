const XLSX = require('xlsx');
const path = require('path');

const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';

function parseRoseDetailed(filePath, leagueName) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`ROSE DETTAGLIATE - ${leagueName}`);
  console.log('='.repeat(80));

  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets['TutteLeRose'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Parse teams: structure is pairs of teams side by side (cols 0-3 left, cols 5-8 right)
  const teams = {};
  let leftTeam = null;
  let rightTeam = null;
  let leftPlayers = [];
  let rightPlayers = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const leftCell = String(row[0] || '').trim();
    const rightCell = String(row[5] || '').trim();

    // Check for team header row (the row before "Ruolo" header)
    if (i + 1 < data.length) {
      const nextRow = data[i + 1];
      const nextLeft = String(nextRow[0] || '').trim();
      if (nextLeft === 'Ruolo' || nextLeft === 'ruolo') {
        // Save previous teams
        if (leftTeam) {
          teams[leftTeam] = { players: leftPlayers, credits: null };
        }
        if (rightTeam) {
          teams[rightTeam] = { players: rightPlayers, credits: null };
        }
        leftTeam = leftCell;
        rightTeam = rightCell;
        leftPlayers = [];
        rightPlayers = [];
        continue;
      }
    }

    // Check for credit row
    if (leftCell.startsWith('Crediti Residui:')) {
      const credits = parseInt(leftCell.replace('Crediti Residui:', '').trim());
      if (leftTeam && !teams[leftTeam]) {
        teams[leftTeam] = { players: leftPlayers, credits };
        leftPlayers = [];
      } else if (leftTeam && teams[leftTeam]) {
        teams[leftTeam].credits = credits;
      }
    }
    if (rightCell.startsWith('Crediti Residui:')) {
      const credits = parseInt(rightCell.replace('Crediti Residui:', '').trim());
      if (rightTeam && !teams[rightTeam]) {
        teams[rightTeam] = { players: rightPlayers, credits };
        rightPlayers = [];
      } else if (rightTeam && teams[rightTeam]) {
        teams[rightTeam].credits = credits;
      }
    }

    // Check for player rows (role in col 0 or col 5)
    const roles = ['P', 'D', 'C', 'A', 'Por', 'Dc', 'Dd', 'Ds', 'E', 'M', 'C', 'T', 'W', 'Pc', 'A', 'B'];
    const rolePattern = /^[A-Za-z;]+$/;

    if (leftCell && rolePattern.test(leftCell) && row[1]) {
      leftPlayers.push({
        role: leftCell,
        name: String(row[1]).trim(),
        team: String(row[2]).trim(),
        cost: row[3] !== '' ? Number(row[3]) : 0
      });
    }
    if (rightCell && rolePattern.test(rightCell) && row[6]) {
      rightPlayers.push({
        role: rightCell,
        name: String(row[6]).trim(),
        team: String(row[7]).trim(),
        cost: row[8] !== '' ? Number(row[8]) : 0
      });
    }
  }

  // Save last teams
  if (leftTeam && !teams[leftTeam]) {
    teams[leftTeam] = { players: leftPlayers, credits: null };
  }
  if (rightTeam && !teams[rightTeam]) {
    teams[rightTeam] = { players: rightPlayers, credits: null };
  }

  // Output results
  console.log('\n--- RIEPILOGO CREDITI PER SQUADRA ---');
  for (const [teamName, teamData] of Object.entries(teams)) {
    console.log(`\n${teamName}: ${teamData.credits} crediti residui, ${teamData.players.length} giocatori`);
    // Print all players with costs
    for (const p of teamData.players) {
      console.log(`  ${p.role.padEnd(8)} ${p.name.padEnd(25)} ${p.team.padEnd(5)} Costo: ${p.cost}`);
    }
  }

  return teams;
}

const league = process.argv[2] || 'ft';

if (league === 'ft') {
  parseRoseDetailed(
    path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
    'Fanta Tosti'
  );
} else {
  parseRoseDetailed(
    path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
    'FantaMantra Manageriale'
  );
}
