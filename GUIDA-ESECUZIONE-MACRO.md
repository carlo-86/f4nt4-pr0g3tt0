# Guida Passo-Passo: Esecuzione Macro DEFINITIVA

## Prerequisiti
- Microsoft Excel (con macro abilitate)
- I file DB Excel nella cartella `2025-26`:
  - `Fanta Tosti 2026/DB Excel/Fanta Tosti 2026 - DB completo (06.02.2026).xlsx`
  - `FantaMantra Manageriale/DB Excel/FantaMantra Manageriale - DB completo (06.02.2026).xlsx`
- Il file listone aggiornato (scaricato da fantacalcio.it):
  - `Quotazioni_Fantacalcio_Stagione_2025_26_22.02.2026.xlsx`
- I file macro VBA **DEFINITIVA**:
  - `VBA-MACRO-FT-DEFINITIVA.bas` (per Fanta Tosti)
  - `VBA-MACRO-FM-DEFINITIVA.bas` (per FantaMantra Manageriale)

## IMPORTANTE: Backup prima di procedere
**Crea una copia di backup di entrambi i DB Excel PRIMA di eseguire le macro!**

## Protezione fogli di lavoro
Le macro gestiscono **automaticamente** la protezione dei fogli: rimuovono la protezione all'inizio dell'esecuzione e la ripristinano al termine. Solo i fogli che erano effettivamente protetti vengono ri-protetti. Non è necessario rimuovere manualmente la protezione prima di eseguire le macro.

---

## Cosa fanno le macro DEFINITIVA

Le macro DEFINITIVA automatizzano **tutte** le operazioni in un'unica esecuzione, incluso l'aggiornamento del listone:

### FT (`ESEGUI_TUTTO_FT`):
1. **FASE 0 - Listone**: Importa il listone aggiornato nel foglio LISTA (aggiorna quotazioni, FVM, ruoli, squadre; aggiunge nuovi calciatori; rimuove i punti dai nomi; riordina alfabeticamente)
2. **FASE 1 - Asta Riparazione**: Inserisce i giocatori acquisiti nell'asta di riparazione (post 06/02) nelle colonne SQUADRE
3. **FASE 2 - Allineamento Date**: Allinea retroattivamente tutte le date di assicurazione al ciclo triennale (regola triennio rigido: decorrenza = scatto triennale dalla data acquisto)
4. **FASE 3 - Assicurazioni**: Registra il flag "A" e la data per tutti i giocatori assicurati (con gestione rinnovi preventivi: se il triennio non e' scaduto, la nuova decorrenza parte dalla scadenza del triennio corrente)
5. **FASE 4 - Fix Formule DB**: Corregge 3 bug nelle formule del foglio DB + rinomina header BP2
6. **FASE 5 - Contratti Invernali**: Calcola la quota contratti per gli acquisti del mercato di riparazione (Qt.Attuale x 0.05 EUR per giocatore) e la scrive nel foglio QUOTE+MONTEPREMI 2026 colonna I

### FM (`ESEGUI_TUTTO_FM`):
1. **FASE 0 - Listone**: Importa il listone aggiornato nel foglio LISTA (con quotazioni Mantra: Qt.A M, Qt.I M, FVM M)
2. **FASE 1 - Scambi**: Sposta i giocatori coinvolti negli scambi post-06/02 (Minnesota-Lino, Minnesota-Papaie, David Minnesota-Hellas, Bernabe/Cancellieri)
3. **FASE 2 - Asta Riparazione**: Inserisce i nuovi giocatori acquisiti nell'asta di riparazione
4. **FASE 3 - Allineamento Date**: Allinea retroattivamente tutte le date di assicurazione al ciclo triennale (regola triennio rigido)
5. **FASE 4 - Assicurazioni**: Registra il flag "A" e la data per tutti i giocatori assicurati (con gestione rinnovi preventivi)
6. **FASE 5 - Fix Formule DB**: Corregge 3 bug nelle formule del foglio DB + rinomina header BP2
7. **FASE 6 - Contratti Invernali**: Calcola la quota contratti per gli acquisti del mercato di riparazione e la scrive nel foglio QUOTE+MONTEPREMI 2026 colonna I

### Note importanti:
- **Listone**: La macro chiede di selezionare il file listone tramite dialog. Scegliere il file piu' recente (es. `Quotazioni_Fantacalcio_Stagione_2025_26_22.02.2026.xlsx`). I punti vengono rimossi automaticamente dai nomi (es. "Martinez L." diventa "Martinez L").
- **Kouame** (FT+FM KFP): NON assicurabile (non piu' listato su Leghe Fantacalcio) - la macro lo salta con log
- **Posch** (FM KFP): RESPINTO (svincolato) - la macro lo salta con log
- **Tavares N.** (FM Lino): RESPINTO (triennale non decorso, scade ago 2027) - la macro lo salta con log
- **Santos A.** = "Allison S." nella comunicazione FT CKC
- **Kone I.** (non Kone M.) per FT CKC e FM KFP

### Regola triennio rigido (date assicurazione):
La decorrenza della copertura assicurativa e' SEMPRE allineata al ciclo triennale dalla Data acquisto del calciatore, mai alla data di comunicazione. In pratica:
- **Allineamento retroattivo**: La macro corregge le date di tutti i giocatori gia' assicurati, allineandole al boundary triennale piu' vicino alla catena Dacq, Dacq+3, Dacq+6, ...
- **Rinnovi preventivi**: Se un fantallenatore rinnova un'assicurazione prima della scadenza del triennio, la nuova copertura decorre dalla scadenza del triennio corrente (non dalla data di comunicazione)
- **Assicurazioni tardive**: Se un fantallenatore assicura un calciatore dopo lo scatto di triennio, la copertura decorre dal precedente scatto triennale

### Quote contratti invernali:
La quota contratti del mercato di riparazione viene calcolata come:
- **Formula**: Qt.Attuale (dal listone) x 0.05 EUR per ogni giocatore acquisito
- **Destinazione**: Foglio QUOTE+MONTEPREMI 2026, colonna I ("Quota contratti mercato di riparazione")
- Le squadre senza acquisti invernali non vengono toccate (restano a 0.00)

### Correzioni formule DB (audit 22/02/2026):
Le macro correggono automaticamente 3 bug individuati nel foglio DB:
1. **BV/BW/BX righe 3-12**: Usavano `$AO` invece di `MAX($J,$AO)` nella formula INVARIATA (10 righe su 983)
2. **BY/BZ/CA**: La maggior parte delle righe referenziava le vecchie colonne `$BB/$BC/$BD/$BE` invece delle corrette `$BF/$BG/$BH/$BI` (formula POSITIVA, ~580 righe)
3. **BS ramo Portieri**: Parentesizzazione errata + coefficiente 0.75 al posto di 0.65 — la formula calcolava `MAX(J,AO) - (AL*VLOOKUP)*0.75` invece di `(MAX(J,AO) - AL*VLOOKUP)*0.65`
- **BP2**: Header rinominato da "Valore se A e % positiva" a "Valore tabellare per ruolo"
- **BN**: Segnalata come ridondante (duplicato di AO, nessun riferimento nell'intera cartella)

> **Nota**: Le macro `CorreggiFormuleDB_FT` e `CorreggiFormuleDB_FM` possono essere eseguite anche **separatamente** (senza rieseguire l'intero ESEGUI_TUTTO) se si vuole applicare solo la correzione formule.

---

## Fanta Tosti 2026

### Passo 1: Apri il file DB
1. Apri `Fanta Tosti 2026 - DB completo (06.02.2026).xlsx`
2. Quando richiesta la password, inserisci: `89y3R8HF'(()h7t87gH)(/0?9U38Qyp99`
3. Se appare un avviso sulle macro, clicca "Abilita contenuto"

### Passo 2: Apri l'editor VBA
1. Premi **Alt+F11** (si apre l'editor Visual Basic)
2. Nel menu, vai su **Inserisci > Modulo** (si apre un foglio bianco)

### Passo 3: Incolla il codice macro
1. Apri il file `VBA-MACRO-FT-DEFINITIVA.bas` con un editor di testo (Notepad, VS Code, ecc.)
2. **Seleziona tutto** (Ctrl+A) e **copia** (Ctrl+C)
3. Torna nell'editor VBA e **incolla** (Ctrl+V) nel modulo vuoto

### Passo 4: Esegui la macro principale
1. Premi **F5** (o menu Esegui > Esegui Sub/UserForm)
2. Nella finestra di dialogo, seleziona **ESEGUI_TUTTO_FT**
3. Clicca **Esegui**
4. **Appare una finestra di selezione file**: scegli il file listone (`Quotazioni_Fantacalcio_Stagione_2025_26_22.02.2026.xlsx`)
5. La macro esegue automaticamente tutte le 6 fasi (listone + asta riparazione + allineamento date + assicurazioni + fix formule DB + contratti invernali)
6. Al termine appare un messaggio di conferma

### Passo 5: Controlla il log
1. Vai al foglio **LOG_MACRO** (creato automaticamente)
2. Verifica che tutte le operazioni siano andate a buon fine:
   - **FASE 0 (Listone)**: `X aggiornati, Y aggiunti, Z skippati`
   - **FASE 1 (Asta rip.)**: `INSERITO` / `GIA' PRESENTE`
   - **FASE 2 (Allineamento)**: `ALLINEATO: Squadra / Calciatore - da DD/MM/YYYY a DD/MM/YYYY` + contatori finali
   - **FASE 3 (Assicurazioni)**: `ASSICURATO` / `RINNOVO` / `RINNOVO PREVENTIVO` / `SKIP`
   - **FASE 4 (Fix Formule)**: `Corretto: BV/BW/BX`, `Corretto: BY/BZ/CA`, `Corretto: BS` + `Rinominato: BP2`
   - **FASE 5 (Contratti)**: Per ogni squadra: quota EUR con dettaglio per giocatore + `TOTALE`

### Passo 6: Verifica
1. Nell'editor VBA, premi **F5** di nuovo
2. Seleziona **VerificaAssicuratiFT**
3. Appare un riepilogo di tutti i giocatori assicurati per squadra
4. Verifica che corrisponda alle comunicazioni

### Passo 7: Salva
1. Chiudi l'editor VBA (Alt+Q o X)
2. **Salva il file** (Ctrl+S)
3. Se Excel chiede il formato, scegli `.xlsx` (le macro non servono piu', i dati sono gia' salvati)

---

## FantaMantra Manageriale

### Passo 1: Apri il file DB
1. Apri `FantaMantra Manageriale - DB completo (06.02.2026).xlsx`
2. Password: `89y3R8HF'(()h7t87gH)(/0?9U38Qyp99`
3. Abilita contenuto se richiesto

### Passo 2: Apri l'editor VBA
1. Premi **Alt+F11**
2. **Inserisci > Modulo**

### Passo 3: Incolla il codice macro
1. Apri `VBA-MACRO-FM-DEFINITIVA.bas` con un editor di testo
2. **Seleziona tutto** (Ctrl+A) e **copia** (Ctrl+C)
3. Incolla (Ctrl+V) nel modulo vuoto

### Passo 4: Esegui la macro principale
1. Premi **F5**
2. Seleziona **ESEGUI_TUTTO_FM**
3. Clicca **Esegui**
4. **Seleziona il file listone** quando appare la finestra di selezione file
5. La macro esegue in sequenza le 7 fasi:
   - Aggiornamento LISTA dal listone (quotazioni Mantra)
   - Scambi post-06/02 (spostamento giocatori tra colonne)
   - Inserimento asta riparazione
   - Allineamento retroattivo date assicurazione
   - Registrazione assicurazioni (con rinnovi preventivi)
   - Correzione formule DB
   - Calcolo contratti invernali
6. Al termine appare un messaggio di conferma

### Passo 5: Controlla il log
1. Vai al foglio **LOG_MACRO**
2. Verifica le operazioni:
   - **FASE 0**: N. aggiornati/aggiunti nella LISTA
   - **FASE 1**: `SPOSTATO` / `GIA' IN DESTINAZIONE` / `NON TROVATO nella colonna X (potrebbe essere gia' spostato)`
   - **FASE 2**: `INSERITO` / `GIA' PRESENTE`
   - **FASE 3**: `ALLINEATO: Squadra / Calciatore - da ... a ...` + contatori finali
   - **FASE 4**: `ASSICURATO` / `RINNOVO` / `RINNOVO PREVENTIVO` / `SKIP`
   - **FASE 5**: `Corretto: BV/BW/BX`, `Corretto: BY/BZ/CA`, `Corretto: BS` + `Rinominato: BP2`
   - **FASE 6**: Per ogni squadra: quota EUR con dettaglio per giocatore + `TOTALE`

### Passo 6: Verifica
1. Premi **F5** > seleziona **VerificaAssicuratiFM**
2. Confronta con le comunicazioni

### Passo 7: Salva
1. Chiudi VBA (Alt+Q)
2. Salva come `.xlsx`

---

## Verifica finale

Dopo aver completato entrambe le macro:

1. Vai al foglio **LISTA** di ciascun DB:
   - Verifica che i nomi non abbiano punti (es. "Martinez L" e non "Martinez L.")
   - Controlla che le quotazioni e FVM siano aggiornati
   - Verifica che la lista sia ordinata alfabeticamente
2. Vai al foglio **ROSA** di ciascun DB:
   - Seleziona ogni squadra dal menu a tendina
   - Verifica che:
     - I giocatori assicurati mostrino "A" nella colonna appropriata
     - La colonna E (Costo) mostri il valore calcolato
3. Confronta i valori con il file `REPORT-ASSICURAZIONI-2026.md`
4. Verifica le formule nel foglio **DB**:
   - **BV3**: deve contenere `MAX($J3,$AO3)` (non solo `$AO3`)
   - **BY4**: deve contenere `$BF` (non `$BB`)
   - **BS** (qualsiasi riga): il ramo Portiere deve avere `(MAX(...)-...)*0.65` (non `...-(...)*0.75`)
   - **BP2**: deve mostrare "Valore tabellare per ruolo"
5. Verifica le **date di assicurazione** nel foglio SQUADRE:
   - Per ogni giocatore assicurato, la data (+7) deve corrispondere a uno scatto triennale dalla Data acquisto (+4)
   - Per i rinnovi preventivi, la data sara' futura rispetto al 14/02/2026 (es. se Dacq=12/09/2020, la data potrebbe essere 12/09/2026 o 12/09/2029)
   - Controlla il LOG_MACRO per la lista degli allineamenti effettuati
6. Verifica il foglio **QUOTE+MONTEPREMI 2026**:
   - La colonna I ("Quota contratti mercato di riparazione") deve mostrare i costi per le squadre con acquisti invernali
   - Le squadre senza acquisti (FCK, Tronzano per FT; Mastri per FM) devono avere 0.00
   - La riga TOTALI deve mostrare la somma corretta

---

## Riepilogo costi attesi

### FT (114.4 cr totali)
| Squadra | Costo |
|---------|-------|
| Hellas Madonna | 11.0 |
| PARTIZAN | 5.6 |
| Kung Fu Pandev | 15.2 |
| FC CKC 26 | 15.2 |
| muttley superstar | 5.6 |
| Millwall | 15.4 |
| Papaie Top Team | 1.9 |
| Legenda Aurea | 44.5 |

### FM (208.1 cr totali)
| Squadra | Costo |
|---------|-------|
| Kung Fu Pandev | 14.7 |
| FC CKC 26 | 16.3 |
| H-Q-A Barcelona | 54.5 |
| Hellas Madonna | 40.9 |
| FICA | 4.1 |
| Lino Banfield FC | 20.9 |
| Minnesota al Max | 14.3 |
| Papaie Top Team | 17.1 |
| Legenda Aurea | 25.4 |
