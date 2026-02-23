const XLSX = require('xlsx');
const path = require('path');
const BASE = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26';
const PW = "89y3R8HF'(()h7t87gH)(/0?9U38Qyp99";

function norm(s) {
  return String(s||'').trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[''`]/g,'').replace(/\./g,'').replace(/\s+/g,' ');
}

// Search in Rose files
function searchRose(filePath, label, searchTerms) {
  console.log(`\n=== ${label} ===`);
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['TutteLeRose'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  for (const term of searchTerms) {
    const nTerm = norm(term);
    console.log(`\nSearching: "${term}" (norm: "${nTerm}")`);
    let found = false;
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      if (!row) continue;
      for (let c = 0; c < row.length; c++) {
        const val = String(row[c] || '');
        const nVal = norm(val);
        if (nVal && nVal.length >= 3 && (
          nVal.includes(nTerm) || nTerm.includes(nVal) ||
          (nTerm.length >= 4 && nVal.length >= 4 && levenshtein(nVal, nTerm) <= 2)
        )) {
          // Print context
          const ctx = [];
          for (let cc = Math.max(0, c-2); cc < Math.min(row.length, c+5); cc++) {
            if (row[cc] !== '') ctx.push(`col${cc}="${row[cc]}"`);
          }
          console.log(`  Row ${r}: ${ctx.join(', ')}`);
          found = true;
        }
      }
    }
    if (!found) console.log(`  NOT FOUND in Rose`);
  }
}

// Search in DB sheet
function searchDB(filePath, label, searchTerms) {
  console.log(`\n=== ${label} - DB Sheet ===`);
  const wb = XLSX.readFile(filePath, { password: PW });
  const ws = wb.Sheets['DB'];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  for (const term of searchTerms) {
    const nTerm = norm(term);
    console.log(`\nSearching DB for: "${term}" (norm: "${nTerm}")`);
    let found = false;
    // Search in col 62 (BK = player name)
    for (let r = 2; r < Math.min(data.length, 1000); r++) {
      const val = String(data[r]?.[62] || '');
      const nVal = norm(val);
      if (nVal && (
        nVal.includes(nTerm) || nTerm.includes(nVal) ||
        (nTerm.length >= 4 && nVal.length >= 4 && levenshtein(nVal, nTerm) <= 2)
      )) {
        console.log(`  Row ${r}: name="${val}" FVM=${data[r]?.[79]}`);
        found = true;
      }
    }
    if (!found) console.log(`  NOT FOUND in DB`);
  }
}

function levenshtein(a, b) {
  const m=a.length, n=b.length;
  const dp = Array.from({length:m+1},()=>Array(n+1).fill(0));
  for(let i=0;i<=m;i++) dp[i][0]=i;
  for(let j=0;j<=n;j++) dp[0][j]=j;
  for(let i=1;i<=m;i++) for(let j=1;j<=n;j++)
    dp[i][j]=Math.min(dp[i-1][j]+1,dp[i][j-1]+1,dp[i-1][j-1]+(a[i-1]!==b[j-1]?1:0));
  return dp[m][n];
}

const MISSING = ['KOUAME', 'ALLISON', 'LUIS ENRIQUE', 'SCAMMACCA', 'SCAMACCA'];

searchRose(
  path.join(BASE, 'Fanta Tosti 2026', 'Fanta Tosti 2026 - Rose (17.02.2026).xlsx'),
  'FT Rose', MISSING
);

searchRose(
  path.join(BASE, 'FantaMantra Manageriale', 'FantaMantra Manageriale - Rose (17.02.2026).xlsx'),
  'FM Rose', MISSING
);

searchDB(
  path.join(BASE, 'Fanta Tosti 2026', 'DB Excel', 'Fanta Tosti 2026 - DB completo (06.02.2026).xlsx'),
  'FT', MISSING
);

searchDB(
  path.join(BASE, 'FantaMantra Manageriale', 'DB Excel', 'FantaMantra Manageriale - DB completo (06.02.2026).xlsx'),
  'FM', MISSING
);
