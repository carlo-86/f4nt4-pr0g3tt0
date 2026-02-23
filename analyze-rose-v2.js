const XLSX = require('xlsx');
const path = require('path');

const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';

function parseRose(filePath, leagueName) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`ROSE - ${leagueName}`);
  console.log('='.repeat(80));

  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets['TutteLeRose'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // Step 1: Find all team header rows (row where next row starts with "Ruolo" or "ruolo")
  const pairStarts = [];
  for (let i = 0; i < data.length - 1; i++) {
    const nextLeft = String(data[i + 1][0] || '').trim().toLowerCase();
    if (nextLeft === 'ruolo') {
      pairStarts.push(i);
    }
  }
  console.log(`Trovate ${pairStarts.length} coppie di squadre a righe: ${pairStarts.join(', ')}`);

  // Step 2: Parse each pair
  const teams = {};

  for (let p = 0; p < pairStarts.length; p++) {
    const startRow = pairStarts[p];
    const endRow = p + 1 < pairStarts.length ? pairStarts[p + 1] : data.length;

    const leftTeam = String(data[startRow][0] || '').trim();
    const rightTeam = String(data[startRow][5] || '').trim();

    const leftData = { players: [], credits: null };
    const rightData = { players: [], credits: null };

    // Parse rows from startRow+2 (skip header "Ruolo" row) to endRow
    for (let i = startRow + 2; i < endRow; i++) {
      const row = data[i];
      const leftCell = String(row[0] || '').trim();
      const rightCell = String(row[5] || '').trim();

      // Left side: credits
      if (leftCell.startsWith('Crediti Residui:')) {
        leftData.credits = parseInt(leftCell.replace('Crediti Residui:', '').trim());
      }
      // Right side: credits
      if (rightCell.startsWith('Crediti Residui:')) {
        rightData.credits = parseInt(rightCell.replace('Crediti Residui:', '').trim());
      }

      // Left side: player (role is short, max ~10 chars, no spaces)
      if (leftCell && !leftCell.startsWith('Crediti') && leftCell.length < 15 && row[1]) {
        const name = String(row[1]).trim();
        if (name && name !== '' && name.length > 1) {
          leftData.players.push({
            role: leftCell,
            name: name,
            team: String(row[2] || '').trim(),
            cost: row[3] !== '' ? Number(row[3]) : 0
          });
        }
      }
      // Right side: player
      if (rightCell && !rightCell.startsWith('Crediti') && rightCell.length < 15 && row[6]) {
        const name = String(row[6]).trim();
        if (name && name !== '' && name.length > 1) {
          rightData.players.push({
            role: rightCell,
            name: name,
            team: String(row[7] || '').trim(),
            cost: row[8] !== '' ? Number(row[8]) : 0
          });
        }
      }
    }

    teams[leftTeam] = leftData;
    teams[rightTeam] = rightData;
  }

  // Output
  console.log('\n--- RIEPILOGO CREDITI ---');
  const sortedTeams = Object.entries(teams).sort((a, b) => a[0].localeCompare(b[0]));
  for (const [name, data] of sortedTeams) {
    console.log(`${name.padEnd(35)} Crediti: ${String(data.credits).padStart(4)}  Giocatori: ${data.players.length}`);
  }

  // Detailed rosters
  for (const [name, teamData] of sortedTeams) {
    console.log(`\n--- ${name} (${teamData.credits} crediti, ${teamData.players.length} giocatori) ---`);
    for (const p of teamData.players) {
      console.log(`  ${p.role.padEnd(10)} ${p.name.padEnd(25)} ${p.team.padEnd(5)} Costo: ${p.cost}`);
    }
  }

  return teams;
}

const league = process.argv[2] || 'ft';
if (league === 'ft') {
  parseRose(
    path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
    'Fanta Tosti'
  );
} else {
  parseRose(
    path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
    'FantaMantra Manageriale'
  );
}
