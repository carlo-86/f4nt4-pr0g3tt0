const XLSX = require('xlsx');
const path = require('path');

const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function parseSquadre(filePath, leagueName) {
  console.log(`\n${'='.repeat(80)}`);
  console.log(`SQUADRE - ${leagueName}`);
  console.log('='.repeat(80));

  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['SQUADRE'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // First, print the full header rows to understand column layout
  console.log('\n--- Struttura colonne ---');
  // Row 3 has team info, Row 4 has column headers
  const row3 = data[3] || [];
  const row4 = data[4] || [];

  // Find team block boundaries by looking at row 3 (team names + metadata)
  // and row 4 (column headers like "Calciatore", "Ruolo", etc.)
  const teamBlocks = [];
  for (let col = 0; col < row4.length; col++) {
    if (String(row4[col]).trim() === 'Calciatore') {
      teamBlocks.push(col);
    }
  }

  console.log(`Team blocks start at columns: ${teamBlocks.join(', ')}`);

  // Get team names from row 3
  for (const startCol of teamBlocks) {
    // Team name should be at row 3, startCol
    const teamName = String(row3[startCol] || '').trim();
    console.log(`\nBlock col ${startCol}: Team = "${teamName}"`);

    // Get metadata: number of players, credits
    // Look at cells around the team name in row 3
    let numPlayers = '';
    let credits = '';
    for (let c = startCol; c < startCol + 14; c++) {
      const val = String(row3[c] || '').trim();
      if (val.includes('Numero calciatori')) {
        numPlayers = row3[c + 3] || row3[c + 2] || row3[c + 1] || '';
      }
      if (val.includes('Crediti disponibili')) {
        credits = row3[c + 3] || row3[c + 2] || row3[c + 1] || '';
      }
    }

    // Column mapping for this block
    // Calciatore, Ruolo, Squadra, Ass. = A, Data acquisto, Q. all'acquisto, FVM Prop. all'acquisto, Data assicuraz., Q. rinn. ass., FVM Prop. rinn. ass., Spesa
    const headers = [];
    for (let c = startCol; c < Math.min(startCol + 14, row4.length); c++) {
      headers.push(String(row4[c] || '').trim());
    }
    console.log(`  Headers: ${headers.join(' | ')}`);

    // Parse player data from row 5 onwards
    const players = [];
    for (let r = 5; r < data.length; r++) {
      const row = data[r];
      const playerName = String(row[startCol] || '').trim();

      if (!playerName || playerName === 'Totali' || playerName === '') continue;
      // Stop if we hit empty data or a different section marker
      if (playerName.startsWith('Numero') || playerName.startsWith('Crediti')) break;

      const player = {
        name: playerName,
        role: String(row[startCol + 1] || '').trim(),
        team: String(row[startCol + 2] || '').trim(),
        insured: String(row[startCol + 3] || '').trim(),
        buyDate: row[startCol + 4],
        quoteBuy: row[startCol + 5],
        fvmPropBuy: row[startCol + 6],
        insDate: row[startCol + 7],
        quoteRenew: row[startCol + 8],
        fvmPropRenew: row[startCol + 9],
        spesa: row[startCol + 10]
      };

      if (player.name.length > 1) {
        players.push(player);
      }
    }

    console.log(`  Giocatori: ${players.length}, Crediti: ${credits}`);
    for (const p of players) {
      // Format buy date
      let buyDateStr = '';
      if (typeof p.buyDate === 'number') {
        // Excel serial date
        const d = XLSX.SSF.parse_date_code(p.buyDate);
        buyDateStr = `${String(d.d).padStart(2,'0')}/${String(d.m).padStart(2,'0')}/${d.y}`;
      }
      let insDateStr = '';
      if (typeof p.insDate === 'number') {
        const d = XLSX.SSF.parse_date_code(p.insDate);
        insDateStr = `${String(d.d).padStart(2,'0')}/${String(d.m).padStart(2,'0')}/${d.y}`;
      } else {
        insDateStr = String(p.insDate || '');
      }

      console.log(`  ${p.insured.padEnd(3)} ${p.name.padEnd(25)} ${p.role.padEnd(6)} ${p.team.padEnd(12)} Acq:${buyDateStr.padEnd(12)} Q.Acq:${String(p.quoteBuy).padEnd(6)} FVM.Acq:${String(p.fvmPropBuy).padEnd(8)} Ins:${insDateStr.padEnd(12)} Q.Rin:${String(p.quoteRenew).padEnd(6)} FVM.Rin:${String(p.fvmPropRenew).padEnd(8)} Sp:${p.spesa}`);
    }
  }
}

const league = process.argv[2] || 'ft';
if (league === 'ft') {
  parseSquadre(
    path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'),
    'Fanta Tosti'
  );
} else {
  parseSquadre(
    path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
    'FantaMantra Manageriale'
  );
}
