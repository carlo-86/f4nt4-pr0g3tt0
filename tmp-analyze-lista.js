const XLSX = require('xlsx');
const ftPath = 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/Fanta Tosti 2026/DB Excel/Fanta Tosti 2026 - DB completo (06.02.2026).xlsx';
const fmPath = 'C:/Users/Carlo/Il mio Drive/Personale/Fantacalcio/2025-26/FantaMantra Manageriale/DB Excel/FantaMantra Manageriale - DB completo (06.02.2026).xlsx';

function analyze(path, label) {
    const wb = XLSX.readFile(path);
    const ws = wb.Sheets['LISTA'];
    const d = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
    const maxC = d.reduce((m,r) => Math.max(m,r.length), 0);
    console.log(`\n${'='.repeat(60)}`);
    console.log(`${label} - LISTA: ${d.length} righe x ${maxC} colonne`);
    console.log('='.repeat(60));
    
    console.log('\nPrime 6 righe:');
    for (let i = 0; i < Math.min(6, d.length); i++) {
        const r = d[i] || [];
        const cols = [];
        for (let c = 0; c < r.length; c++) {
            if (String(r[c]) !== '') cols.push(`[${c}]=${JSON.stringify(r[c])}`);
        }
        console.log(`  R${i}: ${cols.join(' | ')}`);
    }
    
    console.log('\nUltime righe con dati:');
    let cnt2 = 0;
    for (let i = d.length - 1; i >= 0 && cnt2 < 5; i--) {
        const r = d[i] || [];
        if (r.some(v => String(v) !== '')) {
            const cols = [];
            for (let c = 0; c < r.length; c++) {
                if (String(r[c]) !== '') cols.push(`[${c}]=${JSON.stringify(r[c])}`);
            }
            console.log(`  R${i}: ${cols.join(' | ')}`);
            cnt2++;
        }
    }
    
    console.log('\nRicerca giocatori specifici:');
    const targets = ['Carnesecchi','Dimarco','Sportiello','Kouam','Santos A','David','Adzic'];
    for (const t of targets) {
        let found = false;
        for (let i = 0; i < d.length && !found; i++) {
            for (let c = 0; c < (d[i]||[]).length; c++) {
                if (String(d[i][c]).toUpperCase().includes(t.toUpperCase())) {
                    const cols = [];
                    for (let cc = 0; cc < d[i].length; cc++) {
                        if (String(d[i][cc]) !== '') cols.push(`[${cc}]=${JSON.stringify(d[i][cc])}`);
                    }
                    console.log(`  ${t} -> R${i}: ${cols.join(' | ')}`);
                    found = true; break;
                }
            }
        }
        if (!found) console.log(`  ${t} -> NON TROVATO`);
    }
    
    let dotCnt = 0, dotEx = [];
    let nameCol = -1;
    for (let c = 0; c < 5; c++) {
        for (let r = 0; r < 5; r++) {
            const v = String(d[r]?.[c] || '').toUpperCase();
            if (v === 'NOME' || v === 'CALCIATORE') { nameCol = c; break; }
        }
        if (nameCol >= 0) break;
    }
    if (nameCol < 0) nameCol = 1;
    for (let i = 2; i < d.length; i++) {
        const n = String(d[i]?.[nameCol] || '');
        if (n.includes('.')) { dotCnt++; if (dotEx.length < 8) dotEx.push(n); }
    }
    console.log(`\nColonna nomi: ${nameCol} | Nomi con ".": ${dotCnt}`);
    if (dotEx.length > 0) console.log(`  Esempi: ${dotEx.join(', ')}`);
    
    console.log('\nConteggio celle non-vuote per colonna:');
    for (let c = 0; c < maxC; c++) {
        let cnt = 0, hdr = '';
        for (let r = 0; r < d.length; r++) {
            if (String(d[r]?.[c] || '') !== '') cnt++;
        }
        for (let r = 0; r < 3; r++) {
            const v = String(d[r]?.[c] || '');
            if (v !== '') { hdr = v; break; }
        }
        console.log(`  Col ${c}: "${hdr}" -> ${cnt} celle non-vuote`);
    }
}

analyze(ftPath, 'FANTA TOSTI');
analyze(fmPath, 'FANTAMANTRA');
