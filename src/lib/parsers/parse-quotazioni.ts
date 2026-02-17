import * as XLSX from 'xlsx';

export interface ParsedPlayer {
  id: number;
  name: string;
  roleClassic: string;
  roleMantra: string;
  team: string;
  quoteClassic: number;
  quoteInitClassic: number;
  quoteMantra: number;
  quoteInitMantra: number;
  fvm: number;
  fvmMantra: number;
  isActive: boolean; // false if from "Ceduti" sheet
}

/**
 * Parses the Leghe FC "Quotazioni Fantacalcio" Excel file.
 * 
 * Expected format (from file analysis):
 * Sheet "Tutti" contains all active players
 * Sheet "Ceduti" contains released players
 * 
 * Columns (starting from row 2 as header):
 * A: Id | B: R (role classic) | C: RM (role mantra) | D: Nome
 * E: Squadra | F: Qt.A | G: Qt.I | H: Diff.
 * I: Qt.A M | J: Qt.I M | K: Diff.M | L: FVM | M: FVM M
 */
export function parseQuotazioni(buffer: Buffer): ParsedPlayer[] {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const players: ParsedPlayer[] = [];

  // Parse "Tutti" sheet (active players)
  const tuttiSheet = workbook.Sheets['Tutti'];
  if (tuttiSheet) {
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(tuttiSheet, {
      range: 1, // skip title row, use row 2 as header
      defval: null,
    });

    for (const row of rows) {
      const id = Number(row['Id']);
      if (!id || isNaN(id)) continue;

      players.push({
        id,
        name: String(row['Nome'] || ''),
        roleClassic: String(row['R'] || ''),
        roleMantra: String(row['RM'] || ''),
        team: String(row['Squadra'] || ''),
        quoteClassic: Number(row['Qt.A']) || 0,
        quoteInitClassic: Number(row['Qt.I']) || 0,
        quoteMantra: Number(row['Qt.A M']) || 0,
        quoteInitMantra: Number(row['Qt.I M']) || 0,
        fvm: Number(row['FVM']) || 0,
        fvmMantra: Number(row['FVM M']) || 0,
        isActive: true,
      });
    }
  }

  // Parse "Ceduti" sheet (released players)
  const cedutiSheet = workbook.Sheets['Ceduti'];
  if (cedutiSheet) {
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(cedutiSheet, {
      range: 1,
      defval: null,
    });

    for (const row of rows) {
      const id = Number(row['Id']);
      if (!id || isNaN(id)) continue;

      // Don't duplicate if already in "Tutti"
      if (players.some(p => p.id === id)) continue;

      players.push({
        id,
        name: String(row['Nome'] || ''),
        roleClassic: String(row['R'] || ''),
        roleMantra: String(row['RM'] || ''),
        team: String(row['Squadra'] || ''),
        quoteClassic: Number(row['Qt.A']) || 0,
        quoteInitClassic: Number(row['Qt.I']) || 0,
        quoteMantra: Number(row['Qt.A M']) || 0,
        quoteInitMantra: Number(row['Qt.I M']) || 0,
        fvm: Number(row['FVM']) || 0,
        fvmMantra: Number(row['FVM M']) || 0,
        isActive: false,
      });
    }
  }

  return players;
}
