const XLSX = require('xlsx');
const path = require('path');

const FT_FILE = 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/Fanta Tosti 2026/DB Excel/Fanta Tosti 2026 - DB completo (06.02.2026).xlsx';
const FM_FILE = 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/FantaMantra Manageriale/DB Excel/FantaMantra Manageriale - DB completo (06.02.2026).xlsx';

const FT_TEAMS = {
  'FCK': 2, 'Hellas': 14, 'muttley': 26, 'PARTIZAN': 38, 'Legenda': 50,
  'KFP': 62, 'Millwall': 74, 'CKC': 86, 'Papaie': 98, 'Tronzano': 110
};

const FM_TEAMS = {
  'Papaie': 3, 'Legenda': 16, 'Lino': 29, 'KFP': 42, 'FICA': 55,
  'Hellas': 68, 'Minnesota': 81, 'CKC': 94, 'HQA': 107, 'Mastri': 120
};
const ROW_START = 5;
const ROW_END = 51;

function getCellValue(ws, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = ws[addr];
  if (!cell) return null;
  return cell.v;
}

function formatDate(val) {
  if (val == null) return "(empty)";
  if (val instanceof Date) {
    const d = val.getDate().toString().padStart(2, "0");
    const m = (val.getMonth() + 1).toString().padStart(2, "0");
    const y = val.getFullYear();
    return d + "/" + m + "/" + y;
  }
  if (typeof val === "number") {
    const date = XLSX.SSF.parse_date_code(val);
    if (date) {
      const d = date.d.toString().padStart(2, "0");
      const m = date.m.toString().padStart(2, "0");
      const y = date.y;
      return d + "/" + m + "/" + y;
    }
  }
  return String(val);
}

function processLeague(filePath, leagueName, teams) {
  console.log("");
  console.log("=".repeat(80));
  console.log("Reading " + leagueName + ": " + path.basename(filePath));
  console.log("=".repeat(80));

  const wb = XLSX.readFile(filePath, { cellDates: true });
  const ws = wb.Sheets["SQUADRE"];

  if (!ws) {
    console.log("ERROR: SQUADRE sheet not found!");
    return [];
  }

  const results = [];

  for (const [teamName, baseCol] of Object.entries(teams)) {
    const colName = baseCol + 0;
    const colInsFlag = baseCol + 3;
    const colInsDate = baseCol + 7;
    const colSpesa = baseCol + 10;

    for (let r = ROW_START; r <= ROW_END; r++) {
      const playerName = getCellValue(ws, r, colName);
      const insFlag = getCellValue(ws, r, colInsFlag);

      if (insFlag === "A" && playerName) {
        const insDate = getCellValue(ws, r, colInsDate);
        const spesa = getCellValue(ws, r, colSpesa);
        const dateStr = formatDate(insDate);
        const spesaStr = spesa != null ? spesa : "(empty)";

        results.push({
          league: leagueName,
          team: teamName,
          player: playerName,
          date: dateStr,
          spesa: spesaStr
        });
      }
    }
  }

  return results;
}

// Main
const allResults = [];
const ftResults = processLeague(FT_FILE, "FT", FT_TEAMS);
allResults.push(...ftResults);
const fmResults = processLeague(FM_FILE, "FM", FM_TEAMS);
allResults.push(...fmResults);

console.log("");
console.log("=".repeat(80));
console.log("INSURED PLAYERS (flag = A)");
console.log("=".repeat(80));
console.log("LEAGUE".padEnd(6)+" | "+"TEAM".padEnd(12)+" | "+"Player Name".padEnd(25)+" | "+"Insurance Date".padEnd(15)+" | Spesa");
console.log("-".repeat(6)+" | "+"-".repeat(12)+" | "+"-".repeat(25)+" | "+"-".repeat(15)+" | "+"-".repeat(10));

for (const r of allResults) {
  console.log(r.league.padEnd(6)+" | "+r.team.padEnd(12)+" | "+r.player.padEnd(25)+" | "+r.date.padEnd(15)+" | "+r.spesa);
}

console.log("");
console.log("Total insured players found: "+allResults.length);
console.log("  FT: "+ftResults.length);
console.log("  FM: "+fmResults.length);

// Date summary
const dateCounts = {};
for (const r of allResults) {
  dateCounts[r.date] = (dateCounts[r.date] || 0) + 1;
}
console.log("");
console.log("UNIQUE INSURANCE DATES:");
const sorted = Object.entries(dateCounts).sort((a,b) => b[1] - a[1]);
for (const [date, count] of sorted) {
  console.log("  " + date.padEnd(15) + " : " + count + " players");
}
