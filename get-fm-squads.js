const XLSX = require('xlsx');

// FM Rose file - has squad abbreviations in TutteLeRose
const fmRosePath = 'C:\\Users\\Carlo\\Il mio Drive\\Personale\\Fantacalcio\\2025-26\\FantaMantra Manageriale\\FantaMantra Manageriale - Rose (17.02.2026).xlsx';

const wb = XLSX.readFile(fmRosePath);
const ws = wb.Sheets['TutteLeRose'];
const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

// Players I need squad abbreviations for (asta riparazione + some traded players)
const targets = [
    'David', 'Cheddira', 'Brescianini', 'Fullkrug', 'Obert', 'Marcandalli',
    'Bowie', 'Marianucci', 'Marinucci', 'Bakola', 'Adzic', 'Nelsson', 'Dossena',
    'Gandelman', 'Barbieri', 'Britschgi', 'Kouame', 'Raspadori',
    'Durosinmi', 'Vergara', 'Taylor', 'Malen', 'Zaragoza', 'Ekkelenkamp',
    'Belghali', 'Luis Henrique', 'Celik', 'Bernasconi', 'Vaz', 'Ratkov',
    'Solomon', 'Bartesaghi', 'Hien', 'Diego Carlos', 'Mazzitelli', 'Tavares',
    'Koopmeiners', 'Fagioli', 'Bellanova', 'Gimenez', 'Miller', 'Bernab',
    'Cancellieri', 'Scamacca', 'Luis', 'Caprile', 'Cambiaghi', 'Baldanzi',
    'Montip', 'Cataldi', 'Kolasinac', 'Pasalic', 'Nicolussi', 'Vlahovic',
    'Leao', 'Zappa', 'Kon', 'Ferguson', 'Zaniolo', 'Holm', 'Ndicka',
    'Gallo', 'Vasquez', 'Gudmundsson', 'Frendrup', 'Sulemana', 'Sommer',
    'Obert'
];

const roles = new Set(['P','D','C','A','Por','Dc','Dd','Ds','E','B','M','T','W','Pc']);

function isRole(val) {
    if (!val) return false;
    const s = String(val).trim();
    return s.split(';').every(part => roles.has(part.trim()));
}

function norm(s) {
    return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().trim();
}

// Parse teams and players
let currentTeam = '';
const playersByTeam = {};

console.log('=== FM Rose TutteLeRose - Player Squad Lookup ===\n');

for (let i = 0; i < data.length; i++) {
    const row = data[i];

    for (const [roleCol, nameCol, squadCol, spesaCol] of [[0, 1, 2, 3], [5, 6, 7, 8]]) {
        const roleVal = String(row[roleCol] || '').trim();
        const nameVal = String(row[nameCol] || '').trim();
        const squadVal = String(row[squadCol] || '').trim();
        const spesaVal = row[spesaCol];

        if (!roleVal && !nameVal) continue;

        // Check if this row is a team name (non-role value in role column)
        if (roleVal && !isRole(roleVal) && roleVal !== 'Ruolo') {
            // Could be a team name
            if (nameVal === 'Calciatore' || !nameVal) {
                currentTeam = roleVal;
                continue;
            }
        }

        if (!nameVal || nameVal === 'Calciatore') continue;

        // Check if this player is one of our targets
        const nameNorm = norm(nameVal);
        for (const target of targets) {
            const targetNorm = norm(target);
            if (nameNorm.includes(targetNorm) || targetNorm.includes(nameNorm.substring(0, Math.min(5, nameNorm.length)))) {
                console.log(`[${currentTeam}] ${nameVal} -> Sq: "${squadVal}", Sp: ${spesaVal}, Role: ${roleVal} (row ${i})`);
            }
        }
    }
}

// Also specifically look for David to verify team
console.log('\n=== SPECIFIC: David location ===');
for (let i = 0; i < data.length; i++) {
    const row = data[i];
    for (const [roleCol, nameCol, squadCol, spesaCol] of [[0, 1, 2, 3], [5, 6, 7, 8]]) {
        const nameVal = String(row[nameCol] || '').trim();
        if (norm(nameVal).includes('DAVID') && !norm(nameVal).includes('DAVIDE')) {
            const roleVal = String(row[roleCol] || '').trim();
            const squadVal = String(row[squadCol] || '').trim();
            const spesaVal = row[spesaCol];

            // Find which team block this belongs to by searching backwards for team name
            let team = '?';
            for (let j = i; j >= 0; j--) {
                const r = data[j];
                const rv = String(r[roleCol] || '').trim();
                const nv = String(r[nameCol] || '').trim();
                if (rv && !isRole(rv) && rv !== 'Ruolo' && (nv === 'Calciatore' || !nv)) {
                    team = rv;
                    break;
                }
            }
            console.log(`  David found: "${nameVal}" in team "${team}" (row ${i}, col group ${roleCol}, sq="${squadVal}", sp=${spesaVal})`);
        }
    }
}
