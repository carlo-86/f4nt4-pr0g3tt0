# Lessons Learned — Fantacalcio Web App

> Questo file viene aggiornato ogni volta che Carlo corregge un errore di Claude Code.
> All'inizio di ogni sessione, Claude Code legge questo file per non ripetere gli stessi errori.

---

## Come aggiungere una lezione

Dopo ogni correzione, Claude Code aggiunge una voce nel formato:

```
## [Data] — [Breve titolo dell'errore]
**Cosa è successo**: descrizione breve dell'errore commesso
**Regola per il futuro**: la regola concreta da seguire per non ripeterlo
```

---

## Lezioni

### 24/02/2026 — Password apertura file vs protezione fogli
**Cosa è successo**: La guida diceva "inserisci la password" all'apertura del DB Excel, ma i file si aprono senza password. La password serve solo per la protezione dei fogli (Unprotect/Protect), gestita internamente dalla macro.
**Regola per il futuro**: Distinguere sempre tra password di apertura file e password di protezione fogli. I DB fantacalcio non hanno password di apertura.

### 24/02/2026 — F5 nell'editor VBA non mostra la lista macro
**Cosa è successo**: La guida diceva "premi F5 e seleziona la macro", ma F5 nell'editor VBA esegue direttamente la Sub corrente. Per scegliere la macro serve Alt+F8 da Excel.
**Regola per il futuro**: Per eseguire una macro specifica: tornare in Excel, premere Alt+F8 (Visualizza macro), selezionare la macro, cliccare Esegui. Mai dire "premi F5" come modo per scegliere tra macro.

### 24/02/2026 — VBA non ha cortocircuito nell'Or
**Cosa è successo**: `If IsEmpty(x) Or x = "" Then` causa Type Mismatch se la cella contiene un errore (#N/A ecc.), perché VBA valuta ENTRAMBE le condizioni.
**Regola per il futuro**: In VBA, usare sempre If/ElseIf annidati quando si controlla il tipo di un valore prima di confrontarlo. Mai usare `Or` con condizioni che dipendono dal tipo del valore.

### 24/02/2026 — Celle unite nel foglio SQUADRE
**Cosa è successo**: ClearContents fallisce su celle che fanno parte di un merge. UnMerge da solo non basta se le unioni si estendono oltre l'area prevista.
**Regola per il futuro**: Nei fogli SQUADRE dei DB fantacalcio ci sono molte celle unite. Ogni operazione di scrittura/cancellazione deve essere protetta con On Error Resume Next per l'intero blocco, non solo per l'UnMerge.

### 24/02/2026 — Limite caratteri per cella Excel (~32K)
**Cosa è successo**: Scrivere tutto il logText in una singola cella causa "Out of memory" (errore 7) quando il log è molto lungo (centinaia di lookup date nascita).
**Regola per il futuro**: Scrivere il log su più righe con Split(logText, vbCrLf), una riga per cella. Mai scrivere testi lunghi in una singola cella Excel.

### 24/02/2026 — Inserimento calciatori: rispettare struttura DB esistente
**Cosa è successo**: La macro inseriva i nuovi calciatori nella prima riga vuota della squadra senza rispettare il reparto (P/D/C/A), sovrascriveva le colonne Ruolo e Squadra (che contengono formule CERCA.VERT), e non compilava Data acquisto, Q. all'acquisto e FVM Prop. all'acquisto.
**Regola per il futuro**:
1. Prima di scrivere in un foglio Excel esistente, VERIFICARE SEMPRE quali colonne contengono formule — non sovrascriverle mai
2. I calciatori vanno inseriti nella sezione del reparto corretto (P/D/C/A), non nella prima riga vuota
3. Per i dati che Carlo richiede (date, quotazioni, FVM), chiedere sempre da dove prenderli se non è esplicitato
4. La Data acquisto dell'asta riparazione NON è una data fissa — ogni giocatore ha la sua data specifica dal "Mercato ASTA CLASSICA"
5. Q. all'acquisto = quotazione attuale dal listone del 06/02/2026 (dal DB originale pre-modifica)
6. FVM Prop. all'acquisto = FVM dal DB del 06/02/2026

### 24/02/2026 — Le formule CERCA.VERT non esistono nelle righe vuote
**Cosa è successo**: Pensavo che le colonne Ruolo (+1) e Squadra (+2) nel foglio SQUADRE avessero formule CERCA.VERT pre-compilate in tutte le righe, e che bastasse scrivere il nome del calciatore per farle "auto-popolare". In realtà le formule esistono solo nelle righe dove c'è già un calciatore; le righe vuote non hanno nessuna formula.
**Regola per il futuro**: Quando si inserisce un nuovo calciatore in una riga vuota di SQUADRE, bisogna ANCHE scrivere le formule CERCA.VERT per Ruolo (+1) e Squadra (+2). FT: Ruolo=VLOOKUP da LISTA col C (Classic); FM: Ruolo=VLOOKUP da LISTA col D (Mantra). Squadra=VLOOKUP da LISTA col E per entrambi.
