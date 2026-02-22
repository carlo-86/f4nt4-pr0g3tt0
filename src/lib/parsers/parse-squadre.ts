import * as XLSX from 'xlsx';

export interface SquadrePlayerData {
  name: string;
  role: string;           // P, D, C, A (Classic) or Por, Dc, E, M;C, etc. (Mantra)
  squadraSA: string | null;  // Serie A team
  insured: boolean;
  purchaseDate: string | null;    // ISO date string
  quoteAtPurchase: number | null;
  fvmPropAtPurchase: number | null;
  insuranceDate: string | null;   // ISO date string
  quoteRenewal: number | null;
  fvmPropRenewal: number | null;
  purchasePrice: number;
}

export interface SquadreTeamData {
  teamName: string;
  credits: number | null;
  players: SquadrePlayerData[];
}

/**
 * Parse the SQUADRE sheet from a DB Excel file.
 * Works for both Classic (Fanta Tosti) and Mantra (FantaMantra) formats.
 * Returns only active players (excludes historical/released section).
 */
export function parseSquadre(buffer: Buffer): SquadreTeamData[] {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  
  const sheetName = workbook.SheetNames.find(n => n.toUpperCase().includes('SQUADRE'));
  if (!sheetName) {
    throw new Error('Foglio SQUADRE non trovato nel file');
  }
  
  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  
  // Helper to read cell value
  const getCell = (row: number, col: number): any => {
    const addr = XLSX.utils.encode_cell({ r: row, c: col }); // 0-indexed
    const cell = ws[addr];
    return cell ? cell.v : null;
  };
  
  // Helper to read cell as date ISO string
  const getCellDate = (row: number, col: number): string | null => {
    const addr = XLSX.utils.encode_cell({ r: row, c: col });
    const cell = ws[addr];
    if (!cell) return null;
    
    // XLSX with cellDates: true returns Date objects for date cells
    if (cell.t === 'd' && cell.v instanceof Date) {
      return cell.v.toISOString();
    }
    // Also handle if it comes as a number (Excel serial date)
    if (cell.t === 'n' && cell.v > 30000 && cell.v < 60000) {
      const date = XLSX.SSF.parse_date_code(cell.v);
      return new Date(date.y, date.m - 1, date.d).toISOString();
    }
    // Handle string dates
    if (typeof cell.v === 'string') {
      const d = new Date(cell.v);
      if (!isNaN(d.getTime())) return d.toISOString();
    }
    return null;
  };
  
  // 1. Find 'Calciatore' columns in row 5 (0-indexed row 4)
  const headerRow = 4; // 0-indexed
  const calciatoreColumns: number[] = [];
  
  for (let col = 0; col <= range.e.c; col++) {
    const val = getCell(headerRow, col);
    if (val === 'Calciatore') {
      calciatoreColumns.push(col);
    }
  }
  
  if (calciatoreColumns.length === 0) {
    throw new Error('Impossibile trovare le colonne "Calciatore" nella riga 5');
  }
  
  // 2. Map each calciatore column to a team name from row 4 (0-indexed row 3)
  const teamRow = 3; // 0-indexed
  
  interface TeamBlock {
    name: string;
    calCol: number;
    credits: number | null;
  }
  
  const teamBlocks: TeamBlock[] = [];
  
  for (const calCol of calciatoreColumns) {
    let teamName: string | null = null;
    
    // Search for team name at calCol-1 and calCol in row 4
    for (const checkCol of [calCol - 1, calCol]) {
      if (checkCol < 0) continue;
      const val = getCell(teamRow, checkCol);
      if (val && typeof val === 'string') {
        const trimmed = val.trim();
        if (trimmed && 
            !trimmed.startsWith('Numero') && 
            !trimmed.startsWith('Crediti') &&
            !trimmed.startsWith('Valori medi')) {
          teamName = trimmed;
          break;
        }
      }
    }
    
    if (!teamName) continue;
    
    // Get credits from Spesa column position in row 4
    const spesaCol = calCol + 10;
    const creditsVal = getCell(teamRow, spesaCol);
    const credits = typeof creditsVal === 'number' ? Math.round(creditsVal) : null;
    
    teamBlocks.push({ name: teamName, calCol, credits });
  }
  
  // Filter out summary columns
  const filteredBlocks = teamBlocks.filter(
    b => !b.name.startsWith('Valori medi')
  );
  
  // 3. Find historical section boundary ('Elenco storico')
  let historicalRow = range.e.r + 2; // beyond the sheet if not found
  for (let row = 0; row <= range.e.r; row++) {
    for (let col = 0; col < Math.min(10, range.e.c); col++) {
      const val = getCell(row, col);
      if (val && typeof val === 'string' && val.toLowerCase().includes('elenco storico')) {
        historicalRow = row;
        break;
      }
    }
    if (historicalRow < range.e.r + 2) break;
  }
  
  // 4. Find section header rows (repeating 'Calciatore' in first team's column)
  const firstCalCol = calciatoreColumns[0];
  const sectionHeaders: number[] = [];
  
  for (let row = 0; row < historicalRow; row++) {
    if (getCell(row, firstCalCol) === 'Calciatore') {
      sectionHeaders.push(row);
    }
  }
  
  // Build sections: each runs from header+1 to next_header-1 (or historicalRow-1)
  const sections: Array<{ startRow: number; endRow: number }> = [];
  for (let i = 0; i < sectionHeaders.length; i++) {
    const startRow = sectionHeaders[i] + 1;
    const endRow = i < sectionHeaders.length - 1 
      ? sectionHeaders[i + 1] - 1 
      : historicalRow - 1;
    sections.push({ startRow, endRow });
}
  
  // 5. Parse players for each team
  const result: SquadreTeamData[] = [];
  
  for (const block of filteredBlocks) {
    const teamData: SquadreTeamData = {
      teamName: block.name,
      credits: block.credits,
      players: [],
    };
    
    for (const section of sections) {
      for (let row = section.startRow; row <= section.endRow; row++) {
        const playerName = getCell(row, block.calCol);
        if (!playerName || (typeof playerName === 'string' && playerName.trim() === '')) {
          continue;
        }
        
        // Read fields (offset from calciatore column)
        const role = getCell(row, block.calCol + 1);
        const squadra = getCell(row, block.calCol + 2);
        const ass = getCell(row, block.calCol + 3);
        const qAcq = getCell(row, block.calCol + 5);
        const fvmPropAcq = getCell(row, block.calCol + 6);
        const qRinn = getCell(row, block.calCol + 8);
        const fvmPropRinn = getCell(row, block.calCol + 9);
        const spesa = getCell(row, block.calCol + 10);
        
        // Validate: must have role and purchase price
        if (!role || typeof spesa !== 'number') continue;
        
        // Clean "/" values
        const cleanNum = (v: any): number | null => {
          if (v == null || v === '' || v === '/') return null;
          if (typeof v === 'number') return v === 0 ? null : v;
          const parsed = parseFloat(String(v));
          return isNaN(parsed) ? null : parsed;
        };
        
        const player: SquadrePlayerData = {
          name: String(playerName).trim(),
          role: String(role).trim(),
          squadraSA: squadra ? String(squadra).trim() : null,
          insured: ass ? String(ass).trim().toUpperCase() === 'A' : false,
          purchaseDate: getCellDate(row, block.calCol + 4),
          quoteAtPurchase: typeof qAcq === 'number' ? Math.round(qAcq) : null,
          fvmPropAtPurchase: cleanNum(fvmPropAcq),
          insuranceDate: getCellDate(row, block.calCol + 7),
          quoteRenewal: typeof qRinn === 'number' ? Math.round(qRinn) : (
            cleanNum(qRinn) !== null ? Math.round(cleanNum(qRinn)!) : null
          ),
          fvmPropRenewal: cleanNum(fvmPropRinn),
          purchasePrice: Math.round(spesa),
        };
        
        teamData.players.push(player);
      }
    }
    
    result.push(teamData);
  }
  
  return result;
}
