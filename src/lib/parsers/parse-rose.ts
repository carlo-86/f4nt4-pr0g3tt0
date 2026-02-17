import * as XLSX from 'xlsx';

export interface ParsedRosterPlayer {
  role: string;
  name: string;
  teamAbbr: string; // abbreviated team name from Leghe FC
  cost: number; // purchase price in credits
}

export interface ParsedTeamRoster {
  teamName: string;
  creditsRemaining: number;
  players: ParsedRosterPlayer[];
}

/**
 * Parses the Leghe FC "Rose" Excel export.
 * 
 * Format (from file analysis):
 * Single sheet "TutteLeRose"
 * Row 1: League title
 * Row 2: URL
 * Row 3: Notes about released players + download date
 * 
 * Then blocks of 2 teams side by side:
 * - Team name row: Col A = Team1 name, Col F = Team2 name
 * - Header row: "Ruolo | Calciatore | Squadra | Costo" (x2)
 * - Player rows
 * - Credits row: "Crediti Residui: X"
 * - Empty row(s)
 * - Next block
 * 
 * Teams in columns A-D (left) and F-I (right)
 */
export function parseRose(buffer: Buffer): ParsedTeamRoster[] {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheet = workbook.Sheets['TutteLeRose'];
  if (!sheet) throw new Error('Sheet "TutteLeRose" not found');

  // Convert to array of arrays for easier positional parsing
  const data = XLSX.utils.sheet_to_json<(string | number | null)[]>(sheet, {
    header: 1,
    defval: null,
  });

  const teams: ParsedTeamRoster[] = [];
  let i = 0;

  while (i < data.length) {
    const row = data[i];

    // Skip empty rows and header rows
    if (!row || !row[0]) {
      i++;
      continue;
    }

    const cellA = String(row[0] || '').trim();

    // Skip title, URL, notes rows
    if (
      cellA.startsWith('Rose lega') ||
      cellA.startsWith('https://') ||
      cellA.startsWith('*')
    ) {
      i++;
      continue;
    }

    // Skip "Ruolo" header rows
    if (cellA === 'Ruolo') {
      i++;
      continue;
    }

    // Check if this is a credits row
    if (cellA.startsWith('Crediti Residui:')) {
      i++;
      continue;
    }

    // Check if this is a team name row (not a role like P, D, C, A, Por, Dc, etc.)
    const knownRoles = ['P', 'D', 'C', 'A', 'Por', 'Dc', 'Dd', 'Ds', 'E', 'M', 'T', 'W', 'Pc'];
    if (!knownRoles.includes(cellA) && !cellA.startsWith('Crediti')) {
      // This is a team name row — start parsing a block
      const teamNameLeft = cellA;
      const teamNameRight = row[5] ? String(row[5]).trim() : null;

      i++; // skip to header row (Ruolo | Calciatore | ...)
      if (i < data.length && String(data[i]?.[0] || '').trim() === 'Ruolo') {
        i++; // skip header row
      }

      // Parse players for both teams
      const leftPlayers: ParsedRosterPlayer[] = [];
      const rightPlayers: ParsedRosterPlayer[] = [];
      let leftCredits = 0;
      let rightCredits = 0;

      while (i < data.length) {
        const pRow = data[i];
        if (!pRow) { i++; continue; }

        const colA = String(pRow[0] || '').trim();
        const colF = String(pRow[5] || '').trim();

        // Check for credits row (left side)
        if (colA.startsWith('Crediti Residui:')) {
          const match = colA.match(/Crediti Residui:\s*(\d+)/);
          if (match) leftCredits = parseInt(match[1]);

          // Right side might also have credits or a player
          if (colF.startsWith('Crediti Residui:')) {
            const matchR = colF.match(/Crediti Residui:\s*(\d+)/);
            if (matchR) rightCredits = parseInt(matchR[1]);
          }
          i++;
          break; // end of this block
        }

        // Parse left side player
        if (colA && knownRoles.includes(colA)) {
          leftPlayers.push({
            role: colA,
            name: String(pRow[1] || '').trim(),
            teamAbbr: String(pRow[2] || '').trim(),
            cost: Number(pRow[3]) || 0,
          });
        }

        // Parse right side player
        if (colF && knownRoles.includes(colF)) {
          rightPlayers.push({
            role: colF,
            name: String(pRow[6] || '').trim(),
            teamAbbr: String(pRow[7] || '').trim(),
            cost: Number(pRow[8]) || 0,
          });
        }

        // Check if right side has credits
        if (colF.startsWith('Crediti Residui:')) {
          const matchR = colF.match(/Crediti Residui:\s*(\d+)/);
          if (matchR) rightCredits = parseInt(matchR[1]);
        }

        i++;
      }

      // Look for right-side credits if we haven't found them yet
      // (right side credits might be on the row after left side credits)
      if (rightCredits === 0 && i < data.length) {
        const nextRow = data[i];
        if (nextRow) {
          const colA = String(nextRow[0] || '').trim();
          const colF = String(nextRow[5] || '').trim();
          if (colA.startsWith('Crediti Residui:')) {
            const match = colA.match(/Crediti Residui:\s*(\d+)/);
            if (match) rightCredits = parseInt(match[1]);
            i++;
          } else if (colF.startsWith('Crediti Residui:')) {
            const match = colF.match(/Crediti Residui:\s*(\d+)/);
            if (match) rightCredits = parseInt(match[1]);
            i++;
          }
        }
      }

      teams.push({
        teamName: teamNameLeft,
        creditsRemaining: leftCredits,
        players: leftPlayers,
      });

      if (teamNameRight) {
        teams.push({
          teamName: teamNameRight,
          creditsRemaining: rightCredits,
          players: rightPlayers,
        });
      }

      continue;
    }

    i++;
  }

  return teams;
}

/**
 * Maps abbreviated team names from Leghe FC to full names.
 * This handles the 3-letter codes used in the rose export.
 */
export const TEAM_ABBR_MAP: Record<string, string> = {
  'Ata': 'Atalanta',
  'Bol': 'Bologna',
  'Cag': 'Cagliari',
  'Com': 'Como',
  'Cre': 'Cremonese',
  'Emp': 'Empoli',
  'Fio': 'Fiorentina',
  'Gen': 'Genoa',
  'Int': 'Inter',
  'Juv': 'Juventus',
  'Laz': 'Lazio',
  'Lec': 'Lecce',
  'Mil': 'Milan',
  'Mon': 'Monza',
  'Nap': 'Napoli',
  'Par': 'Parma',
  'Pis': 'Pisa',
  'Rom': 'Roma',
  'Sal': 'Salernitana',
  'Sam': 'Sampdoria',
  'Sas': 'Sassuolo',
  'Spe': 'Spezia',
  'Tor': 'Torino',
  'Udi': 'Udinese',
  'Ven': 'Venezia',
  'Ver': 'Verona',
};
