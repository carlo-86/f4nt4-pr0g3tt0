const XLSX = require('xlsx');

// Check if DB sheet values match LISTA values (to understand if DB references LISTA)
const ftPath = 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/Fanta Tosti 2026/DB Excel/Fanta Tosti 2026 - DB completo (06.02.2026).xlsx';
const wb = XLSX.readFile(ftPath);

const wsLista = wb.Sheets['LISTA'];
const lista = XLSX.utils.sheet_to_json(wsLista, {header:1, defval:''});

const wsDB = wb.Sheets['DB'];
const db = XLSX.utils.sheet_to_json(wsDB, {header:1, defval:''});

// DB col 62 (0-based) = player names, col 79 = FVM
// LISTA col 1 = Calciatore, col 7 = FVM

// Build LISTA map by name
const listaMap = {};
for (let i = 1; i < lista.length; i++) {
    const name = String(lista[i][1] || '').toUpperCase().trim();
    if (name) listaMap[name] = { fvm: lista[i][7], qtA: lista[i][5], row: i };
}

// Check some DB players
console.log('=== DB vs LISTA FVM comparison ===');
console.log('DB col 62=name, col 79=FVM | LISTA col 1=name, col 7=FVM\n');
let matches = 0, mismatches = 0, notFound = 0;
for (let i = 2; i < Math.min(30, db.length); i++) {
    const dbName = String(db[i][62] || '').toUpperCase().trim();
    const dbFVM = db[i][79];
    if (!dbName) continue;
    
    const listaEntry = listaMap[dbName];
    if (listaEntry) {
        const match = (Number(dbFVM) === Number(listaEntry.fvm));
        if (match) matches++;
        else mismatches++;
        console.log(`  ${dbName}: DB_FVM=${dbFVM} | LISTA_FVM=${listaEntry.fvm} | ${match ? 'OK' : 'DIVERSO!'}`);
    } else {
        notFound++;
        console.log(`  ${dbName}: DB_FVM=${dbFVM} | LISTA: non trovato`);
    }
}

console.log(`\nPrimi 30: ${matches} match, ${mismatches} diversi, ${notFound} non in LISTA`);

// Also check DB sheet row 0-2 headers for cols 60-82
console.log('\n=== DB sheet headers (cols 60-82) ===');
for (let c = 60; c <= 82; c++) {
    const vals = [];
    for (let r = 0; r < 3; r++) {
        const v = db[r]?.[c];
        if (v !== undefined && String(v) !== '') vals.push(`R${r}="${v}"`);
    }
    if (vals.length > 0) console.log(`  Col ${c}: ${vals.join(' | ')}`);
}

// Check total rows with data in DB
let dbRows = 0;
for (let i = 2; i < db.length; i++) {
    if (String(db[i][62] || '').trim()) dbRows++;
}
console.log(`\nDB sheet: ${dbRows} giocatori con nome in col 62`);

// Check ROSA sheet to understand how it references DB
const wsRosa = wb.Sheets['ROSA'];
const rosa = XLSX.utils.sheet_to_json(wsRosa, {header:1, defval:''});
console.log('\n=== ROSA sheet structure (first 5 rows, cols 0-20) ===');
for (let i = 0; i < Math.min(5, rosa.length); i++) {
    const r = rosa[i] || [];
    const cols = [];
    for (let c = 0; c < Math.min(20, r.length); c++) {
        if (String(r[c]) !== '') cols.push(`[${c}]=${JSON.stringify(r[c])}`);
    }
    console.log(`  R${i}: ${cols.join(' | ')}`);
}
