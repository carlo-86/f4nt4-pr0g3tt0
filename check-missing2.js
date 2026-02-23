const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

// Check remaining FT teams
const wb = XLSX.readFile(path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'), { password: PW });
const ws = wb.Sheets['SQUADRE'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

function dumpTeam(label, col, data) {
  console.log(`\n=== ${label} (col ${col}) ===`);
  for (let r = 5; r < 52; r++) {
    const v = String(data[r]?.[col] || '').trim();
    if (v && v !== 'Calciatore' && v !== 'Ass. = A' && v !== 'Totali') {
      const ins = String(data[r]?.[col + 3] || '').trim();
      const buyDate = data[r]?.[col + 4];
      let bd = '';
      if (typeof buyDate === 'number') {
        const d = XLSX.SSF.parse_date_code(buyDate);
        bd = `${d.d}/${d.m}/${d.y}`;
      }
      const qAcq = data[r]?.[col + 5] || '';
      const fvmAcq = data[r]?.[col + 6] || '';
      const insDate = data[r]?.[col + 7];
      let id = '';
      if (typeof insDate === 'number') {
        const d = XLSX.SSF.parse_date_code(insDate);
        id = `${d.d}/${d.m}/${d.y}`;
      }
      const qRin = data[r]?.[col + 8] || '';
      const fvmRin = data[r]?.[col + 9] || '';
      const sp = data[r]?.[col + 10] || '';
      console.log(`  ${ins === 'A' ? 'ASS' : '---'} ${v.padEnd(22)} Acq:${bd.padEnd(11)} Q:${String(qAcq).padEnd(4)} FVMp:${String(fvmAcq).padEnd(5)} Ins:${id.padEnd(11)} QR:${String(qRin).padEnd(4)} FVMpR:${String(fvmRin).padEnd(5)} Sp:${sp}`);
    }
  }
}

// FT: PARTIZAN at col 38
dumpTeam('FT PARTIZAN', 38, data);
// FT: muttley at col 26
dumpTeam('FT muttley superstar', 26, data);
// FT: FC CKC 26 at col 86
dumpTeam('FT FC CKC 26', 86, data);
// FT: Legenda Aurea at col 50
dumpTeam('FT Legenda Aurea', 50, data);

// FM remaining teams
const wb2 = XLSX.readFile(path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'), { password: PW });
const ws2 = wb2.Sheets['SQUADRE'];
const data2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: '' });

// FM: KFP at col 42 (row3 col41="Kung Fu Pandev"), Calciatore at col 42
dumpTeam('FM Kung Fu Pandev', 42, data2);
// FM: FICA at col 55 (row3 col54="FICA"), Calciatore at col 55
dumpTeam('FM FICA', 55, data2);
// FM: HQA at col 107 (row3 col106="H-Q-A Barcelona"), Calciatore at col 107
dumpTeam('FM H-Q-A Barcelona', 107, data2);
// FM: Papaie at col 3 (already have this data)
