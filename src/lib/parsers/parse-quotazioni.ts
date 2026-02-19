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
  isActive: boolean;
}

// Known column header variations
const COLUMN_MATCHERS: Record<string, (h: string) => boolean> = {
  id:               h => h === 'id',
  name:             h => h === 'nome' || h === 'calciatore',
  roleClassic:      h => h === 'r',
  roleMantra:       h => h === 'rm',
  team:             h => h === 'squadra',
  quoteClassic:     h => h === 'qt.a' || h === 'qta',
  quoteInitClassic: h => h === 'qt.i' || h === 'qti',
  quoteMantra:      h => h === 'qt.a m' || h === 'qtam',
  quoteInitMantra:  h => h === 'qt.i m' || h === 'qtim',
  fvm:              h => h === 'fvm',
  fvmMantra:        h => h === 'fvm m' || h === 'fvmm',
};

/**
 * Finds the header row and maps column indices dynamically.
 * Works regardless of column order or whether there's a title row.
 */
function findColumns(data: (string | number | null)[][]): { headerIdx: number; cols: Record<string, number> } | null {
  for (let i = 0; i < Math.min(5, data.length); i++) {
    const row = data[i];
    if (!row) continue;

    // Check if this row looks like a header (has "Id" and "Nome" or "R")
    const normalized = row.map(v => String(v || '').trim().toLowerCase());
    
    const hasId = normalized.some(h => h === 'id');
    const hasNome = normalized.some(h => h === 'nome' || h === 'calciatore');
    
    if (hasId && hasNome) {
      // Map each known column to its index
      const cols: Record<string, number> = {};
      for (const [field, matcher] of Object.entries(COLUMN_MATCHERS)) {
        const idx = normalized.findIndex(matcher);
        if (idx >= 0) cols[field] = idx;
      }
      return { headerIdx: i, cols };
    }
  }
  return null;
}

function parseSheet(sheet: XLSX.WorkSheet, isActive: boolean): ParsedPlayer[] {
  const players: ParsedPlayer[] = [];

  const data = XLSX.utils.sheet_to_json<(string | number | null)[]>(sheet, {
    header: 1,
    defval: null,
    blankrows: false,
  });

  const result = findColumns(data);
  if (!result) return players;

  const { headerIdx, cols } = result;

  for (let i = headerIdx + 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.length < 5) continue;

    const id = cols.id !== undefined ? Number(row[cols.id]) : 0;
    if (!id || isNaN(id)) continue;

    players.push({
      id,
      name:             cols.name !== undefined ? String(row[cols.name] || '').trim() : '',
      roleClassic:      cols.roleClassic !== undefined ? String(row[cols.roleClassic] || '').trim() : '',
      roleMantra:       cols.roleMantra !== undefined ? String(row[cols.roleMantra] || '').trim() : '',
      team:             cols.team !== undefined ? String(row[cols.team] || '').trim() : '',
      quoteClassic:     cols.quoteClassic !== undefined ? (Number(row[cols.quoteClassic]) || 0) : 0,
      quoteInitClassic: cols.quoteInitClassic !== undefined ? (Number(row[cols.quoteInitClassic]) || 0) : 0,
      quoteMantra:      cols.quoteMantra !== undefined ? (Number(row[cols.quoteMantra]) || 0) : 0,
      quoteInitMantra:  cols.quoteInitMantra !== undefined ? (Number(row[cols.quoteInitMantra]) || 0) : 0,
      fvm:              cols.fvm !== undefined ? (Number(row[cols.fvm]) || 0) : 0,
      fvmMantra:        cols.fvmMantra !== undefined ? (Number(row[cols.fvmMantra]) || 0) : 0,
      isActive,
    });
  }

  return players;
}

export function parseQuotazioni(buffer: Buffer): ParsedPlayer[] {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const players: ParsedPlayer[] = [];
  const seenIds = new Set<number>();

  // Parse "Tutti" sheet (active players)
  const tuttiSheet = workbook.Sheets['Tutti'];
  if (tuttiSheet) {
    for (const p of parseSheet(tuttiSheet, true)) {
      if (!seenIds.has(p.id)) {
        players.push(p);
        seenIds.add(p.id);
      }
    }
  }

  // Parse "Ceduti" sheet (released players)
  const cedutiSheet = workbook.Sheets['Ceduti'];
  if (cedutiSheet) {
    for (const p of parseSheet(cedutiSheet, false)) {
      if (!seenIds.has(p.id)) {
        players.push(p);
        seenIds.add(p.id);
      }
    }
  }

  return players;
}
