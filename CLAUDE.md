# CLAUDE.md — Fantacalcio Web App

## Contesto del Progetto

Applicazione web per la gestione di due leghe di fantacalcio:
- **Fanta Tosti Classic** — lega classica
- **FantaMantra Manageriale** — lega in formato continuativo (mantra)

Il progetto nasce dalla migrazione di un database Excel esistente verso una web app.

L'utente (Carlo) non ha esperienza di programmazione web: ogni soluzione deve essere spiegata in modo chiaro e non dare nulla per scontato.

---

## Stack Tecnologico

- **Backend**: Node.js
- **Lingua del codice**: commenti e variabili in italiano o inglese (scegli coerenza e mantienila)
- **Lingua delle risposte**: SEMPRE italiano

---

## Regole di Comportamento

### Comunicazione
- Rispondi SEMPRE in italiano
- Spiega ogni scelta tecnica con linguaggio semplice
- Se ci sono più opzioni, elencale con pro e contro prima di procedere
- Non assumere mai che Carlo conosca un concetto: spiegalo brevemente

### Pianificazione
- Per qualsiasi task con 3+ step: scrivi prima un piano in `tasks/todo.md` e aspetta conferma
- Se qualcosa va storto: FERMATI, spiega il problema, ri-pianifica
- Non procedere mai su più fronti contemporaneamente senza conferma

### Qualità del Codice
- Semplicità prima di tutto: la soluzione più semplice che funziona
- Nessuna fix temporanea: trova la causa radice
- Ogni modifica deve toccare solo ciò che è strettamente necessario
- Prima di presentare il lavoro, chiediti: "Questa soluzione è la più chiara possibile?"

### Verifica
- Non dichiarare mai un task completato senza dimostrare che funziona
- Mostra sempre l'output o il risultato atteso dopo ogni modifica
- Se scrivi codice che tocca i dati (leggi/scrivi file, database): conferma prima con Carlo

### Gestione File
- Struttura del progetto da mantenere ordinata e documentata
- Ogni file nuovo va spiegato: cosa fa, perché esiste
- Non eliminare mai file senza conferma esplicita

---

## Gestione Task

1. **Pianifica**: scrivi il piano in `tasks/todo.md` con item spuntabili
2. **Conferma**: aspetta ok prima di iniziare
3. **Avanza**: spunta i task man mano che li completi
4. **Spiega**: un riassunto breve ad ogni step
5. **Documenta**: aggiungi una sezione "Risultati" in fondo a `tasks/todo.md`
6. **Impara**: dopo ogni correzione di Carlo, aggiorna `tasks/lessons.md`

---

## Auto-Miglioramento

> **Regola fondamentale**: dopo OGNI correzione da parte di Carlo, aggiungi una nuova regola
> in `tasks/lessons.md` nel formato:
>
> ```
> ## [Data] — [Breve titolo dell'errore]
> **Cosa è successo**: ...
> **Regola per il futuro**: ...
> ```

All'inizio di ogni nuova sessione, leggi `tasks/lessons.md` per applicare le lezioni apprese.

---

## Funzionalità Previste (backlog iniziale)

- [ ] Importazione dati da Excel esistente
- [ ] Gestione rose dei partecipanti
- [ ] Inserimento e visualizzazione voti/punteggi
- [ ] Classifica aggiornata delle due leghe
- [ ] Interfaccia web semplice e usabile da mobile

---

## Note Importanti

- I dati delle leghe sono sensibili (sono di Carlo e dei suoi amici): nessun dato va mai esposto pubblicamente
- Il progetto è personale, non commerciale
- La priorità è la funzionalità, ma senza trascurare l'estetica
