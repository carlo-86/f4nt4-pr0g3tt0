const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

// Check what sheets exist in Rose files and if they have SQUADRE data
function checkFile(filePath, label) {
  console.log(`\n=== ${label} ===`);
  const wb = XLSX.readFile(filePath, { password: PW });
  console.log(`Sheets: ${wb.SheetNames.join(', ')}`);

  // Check for SQUADRE
  if (wb.Sheets['SQUADRE']) {
    const ws = wb.Sheets['SQUADRE'];
    const range = ws['!ref'];
    console.log(`SQUADRE range: ${range}`);
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // Check first team columns for both FT and FM
    console.log(`SQUADRE row count: ${data.length}`);
    console.log(`Row 3 sample: col2="${data[3]?.[2]}" col3="${data[3]?.[3]}" col14="${data[3]?.[14]}"`);
    console.log(`Row 4 sample: col2="${data[4]?.[2]}" col3="${data[4]?.[3]}" col14="${data[4]?.[14]}"`);
    console.log(`Row 5 sample: col2="${data[5]?.[2]}" col3="${data[5]?.[3]}" col14="${data[5]?.[14]}"`);
    console.log(`Row 6 sample: col2="${data[6]?.[2]}" col3="${data[6]?.[3]}" col14="${data[6]?.[14]}"`);
  }

  // Check ROSA sheet too
  if (wb.Sheets['ROSA']) {
    const ws = wb.Sheets['ROSA'];
    const range = ws['!ref'];
    console.log(`ROSA range: ${range}`);
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    // Print column headers
    console.log(`ROSA headers (row 3): `, data[3]?.slice(0, 20)?.map((v,i) => `${i}="${v}"`).join(', '));
    // Print first few data rows
    for (let r = 4; r < 8; r++) {
      const row = data[r];
      if (row) {
        console.log(`ROSA row ${r+1}: K="${row[10]}" L="${row[11]}" M="${row[12]}" N="${row[13]}" O="${row[14]}" E="${row[4]}"`);
      }
    }
  }
}

checkFile(
  path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
  'FT Rose 17/02'
);

checkFile(
  path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
  'FM Rose 17/02'
);
