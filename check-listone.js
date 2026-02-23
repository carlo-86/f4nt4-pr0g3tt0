const XLSX = require('xlsx');

// Listone 22/02/2026 (today)
const path22 = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\f4nt4-pr0g3tt0_varie\\Quotazioni_Fantacalcio_Stagione_2025_26_22.02.2026.xlsx';
// Listone 17/02/2026
const path17 = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26\\Quotazioni ufficiali Leghe FC\\Quotazioni_Fantacalcio_Stagione_2025_26_17.02.2026.xlsx';

function analyzeListone(path, label) {
    console.log(`\n${'='.repeat(70)}`);
    console.log(`LISTONE: ${label}`);
    console.log(`File: ${path}`);
    console.log('='.repeat(70));

    const wb = XLSX.readFile(path);
    console.log(`\nFogli: ${wb.SheetNames.join(', ')}`);

    for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        console.log(`\n--- Foglio: "${sheetName}" (${data.length} righe) ---`);

        // Show first 5 rows to understand structure
        console.log('\nPrime 5 righe (struttura):');
        for (let i = 0; i < Math.min(5, data.length); i++) {
            const row = data[i];
            const cols = [];
            for (let c = 0; c < Math.min(row.length, 15); c++) {
                const v = row[c];
                if (v !== '' && v !== undefined && v !== null) {
                    cols.push(`[${c}]=${v}`);
                }
            }
            console.log(`  Riga ${i}: ${cols.join(' | ')}`);
        }

        // Show total columns
        const maxCols = data.reduce((m, r) => Math.max(m, r.length), 0);
        console.log(`\nColonne totali: ${maxCols}`);

        // Show header row (likely row 0 or row 1)
        if (data.length > 0) {
            console.log(`\nHeader completo:`);
            for (let c = 0; c < maxCols; c++) {
                const h0 = data[0] ? data[0][c] : '';
                const h1 = data[1] ? data[1][c] : '';
                if (h0 || h1) {
                    console.log(`  Col ${c}: R0="${h0}" | R1="${h1}"`);
                }
            }
        }

        // Search for specific players: Sportiello, Adzic, Kouame
        console.log('\n--- Ricerca giocatori specifici ---');
        const targets = ['SPORTIELLO', 'ADZIC', 'KOUAME', 'KOUAM'];
        let nameCol = -1;

        // Find which column has player names
        for (let c = 0; c < maxCols; c++) {
            for (let r = 0; r < Math.min(10, data.length); r++) {
                const v = String(data[r][c] || '').toUpperCase();
                if (v === 'NOME' || v === 'CALCIATORE' || v.includes('NOME')) {
                    nameCol = c;
                    break;
                }
            }
            if (nameCol >= 0) break;
        }

        if (nameCol < 0) {
            // Try to find by looking for known player names
            for (let c = 0; c < Math.min(maxCols, 10); c++) {
                let found = 0;
                for (let r = 0; r < Math.min(50, data.length); r++) {
                    const v = String(data[r][c] || '').toUpperCase();
                    if (['DONNARUMMA', 'MERET', 'SOMMER', 'BARELLA'].some(n => v.includes(n))) found++;
                }
                if (found > 0) { nameCol = c; break; }
            }
        }

        console.log(`  Colonna nomi trovata: ${nameCol}`);

        // Find all columns for context
        for (let r = 0; r < data.length; r++) {
            const name = String(data[r][nameCol] || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
            for (const target of targets) {
                if (name.includes(target)) {
                    const rowData = [];
                    for (let c = 0; c < Math.min(maxCols, 15); c++) {
                        rowData.push(`[${c}]=${data[r][c]}`);
                    }
                    console.log(`  TROVATO: ${name} (riga ${r}): ${rowData.join(' | ')}`);
                }
            }
        }

        // Count total players (non-empty name rows)
        if (nameCol >= 0) {
            let playerCount = 0;
            for (let r = 1; r < data.length; r++) {
                const v = String(data[r][nameCol] || '').trim();
                if (v && v !== 'Nome' && v !== 'Calciatore' && v.length > 1) playerCount++;
            }
            console.log(`\n  Totale giocatori nel listone: ${playerCount}`);
        }

        // Check if any names have dots
        if (nameCol >= 0) {
            let dotCount = 0;
            let dotExamples = [];
            for (let r = 1; r < data.length; r++) {
                const v = String(data[r][nameCol] || '').trim();
                if (v.includes('.')) {
                    dotCount++;
                    if (dotExamples.length < 5) dotExamples.push(v);
                }
            }
            console.log(`\n  Nomi con "." nel listone: ${dotCount}`);
            if (dotExamples.length > 0) console.log(`  Esempi: ${dotExamples.join(', ')}`);
        }
    }
}

analyzeListone(path22, '22/02/2026 (oggi)');
analyzeListone(path17, '17/02/2026');
