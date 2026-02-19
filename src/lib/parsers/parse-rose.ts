import * as XLSX from 'xlsx';

export interface ParsedRosterPlayer {
  role: string;
  name: string;
  teamAbbr: string;
  cost: number;
}

export interface ParsedTeamRoster {
  teamName: string;
  creditsRemaining: number;
  players: ParsedRosterPlayer[];
}

const KNOWN_ROLES = new Set(['P', 'D', 'C', 'A', 'Por', 'Dc', 'Dd', 'Ds', 'E', 'M', 'T', 'W', 'Pc']);

function isRole(val: string): boolean {
  return KNOWN_ROLES.has(val);
}

function parseCredits(val: string): number | null {
  const match = val.match(/Crediti Residui:\s*(\d+)/);
  return match ? parseInt(match[1]) : null;
}

/**
 * Parses the Leghe FC "Rose" Excel export.
 * 
 * Handles asymmetric rosters where left and right teams have different
 * numbers of players, so "Crediti Residui" rows appear on different lines.
 */
export function parseRose(buffer: Buffer): ParsedTeamRoster[] {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheet = workbook.Sheets['TutteLeRose'];
  if (!sheet) throw new Error('Sheet "TutteLeRose" not found');

  const data = XLSX.utils.sheet_to_json<(string | number | null)[]>(sheet, {
    header: 1,
    defval: null,
  });

  const teams: ParsedTeamRoster[] = [];

  // First pass: find all team name rows (blocks)
  const blocks: { row: number; leftName: string; rightName: string | null }[] = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0]) continue;

    const cellA = String(row[0]).trim();

    // Skip known non-team rows
    if (
      cellA.startsWith('Rose lega') ||
      cellA.startsWith('https://') ||
      cellA.startsWith('*') ||
      cellA === 'Ruolo' ||
      cellA.startsWith('Crediti') ||
      isRole(cellA)
    ) continue;

    // This should be a team name — verify by checking next row is "Ruolo" header
    const nextRow = data[i + 1];
    if (nextRow && String(nextRow[0] || '').trim() === 'Ruolo') {
      blocks.push({
        row: i,
        leftName: cellA,
        rightName: row[5] ? String(row[5]).trim() : null,
      });
    }
  }

  // Second pass: parse each block
  for (let b = 0; b < blocks.length; b++) {
    const block = blocks[b];
    const startRow = block.row + 2; // skip team name + header row
    const endRow = b + 1 < blocks.length ? blocks[b + 1].row : data.length;

    const leftPlayers: ParsedRosterPlayer[] = [];
    const rightPlayers: ParsedRosterPlayer[] = [];
    let leftCredits = 0;
    let rightCredits = 0;

    for (let i = startRow; i < endRow; i++) {
      const row = data[i];
      if (!row) continue;

      const colA = String(row[0] || '').trim();
      const colF = String(row[5] || '').trim();

      // Left side
      if (isRole(colA)) {
        const name = String(row[1] || '').trim();
        if (name) {
          leftPlayers.push({
            role: colA,
            name: name.replace(/^\*\s*/, ''), // Remove asterisk from ceduti
            teamAbbr: String(row[2] || '').trim(),
            cost: Number(row[3]) || 0,
          });
        }
      } else {
        const credits = parseCredits(colA);
        if (credits !== null) leftCredits = credits;
      }

      // Right side
      if (isRole(colF)) {
        const name = String(row[6] || '').trim();
        if (name) {
          rightPlayers.push({
            role: colF,
            name: name.replace(/^\*\s*/, ''),
            teamAbbr: String(row[7] || '').trim(),
            cost: Number(row[8]) || 0,
          });
        }
      } else {
        const credits = parseCredits(colF);
        if (credits !== null) rightCredits = credits;
      }
    }

    teams.push({
      teamName: block.leftName,
      creditsRemaining: leftCredits,
      players: leftPlayers,
    });

    if (block.rightName) {
      teams.push({
        teamName: block.rightName,
        creditsRemaining: rightCredits,
        players: rightPlayers,
      });
    }
  }

  return teams;
}

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
