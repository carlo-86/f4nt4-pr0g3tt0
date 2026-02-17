# Guida Setup — Fantacalcio Manager

Segui questi passi nell'ordine esatto. Ogni passo ti dice cosa fare e dove.

---

## PREREQUISITO: Installa Node.js

Node.js è il "motore" che fa funzionare l'app. Serve installarlo sul tuo computer.

1. Vai su **https://nodejs.org**
2. Scarica la versione **LTS** (il bottone verde grande a sinistra)
3. Installa seguendo le istruzioni (clicca "Avanti" su tutto)
4. Per verificare: apri il **Terminale** (Mac) o **Prompt dei comandi** (Windows) e scrivi:
   ```
   node --version
   ```
   Deve mostrarti un numero tipo `v20.x.x`. Se sì, tutto ok.

---

## STEP 1: Crea il database gratuito su Supabase

1. Vai su **https://supabase.com** e clicca **Start your project**
2. Accedi con il tuo account **GitHub** (quello appena creato)
3. Clicca **New Project**
4. Scegli:
   - **Name**: `fantacalcio`
   - **Database Password**: scegli una password e **salvala da qualche parte** (ti servirà)
   - **Region**: scegli `Central EU (Frankfurt)` (il più vicino all'Italia)
5. Clicca **Create new project** e aspetta ~2 minuti
6. Una volta pronto, vai su **Settings** (icona ingranaggio a sinistra) → **Database**
7. Nella sezione **Connection string**, clicca **URI** e copia il testo
   - Sarà tipo: `postgresql://postgres.xxxx:[YOUR-PASSWORD]@aws-0-eu-central-1.pooler.supabase.com:6543/postgres`
   - Sostituisci `[YOUR-PASSWORD]` con la password che hai scelto al punto 4

---

## STEP 2: Scarica il progetto e configuralo

1. Scarica il file ZIP del progetto (te lo fornirò io)
2. Estralo in una cartella, ad esempio `Documenti/fantacalcio-app`
3. Nella cartella del progetto, crea un file chiamato `.env` (nota: inizia con il punto)
4. Dentro `.env` scrivi queste due righe, incollando la connection string di Supabase:

```
DATABASE_URL="postgresql://postgres.xxxx:TUA_PASSWORD@aws-0-eu-central-1.pooler.supabase.com:6543/postgres"
DIRECT_URL="postgresql://postgres.xxxx:TUA_PASSWORD@aws-0-eu-central-1.pooler.supabase.com:5432/postgres"
```

⚠️ **ATTENZIONE**: la prima riga (DATABASE_URL) usa la porta **6543** (pooler), la seconda (DIRECT_URL) usa la porta **5432** (connessione diretta). Entrambe sono nella pagina Connection string di Supabase.

---

## STEP 3: Installa le dipendenze e crea il database

Apri il terminale/prompt nella cartella del progetto e esegui questi comandi uno alla volta:

```bash
npm install
```
(aspetta che finisca — scarica tutte le librerie necessarie)

```bash
npx prisma generate
```
(genera il client per comunicare col database)

```bash
npx prisma db push
```
(crea tutte le tabelle nel database Supabase)

Se tutto va bene, vedrai un messaggio tipo "Your database is now in sync with your Prisma schema."

---

## STEP 4: Prova in locale

```bash
npm run dev
```

Apri il browser e vai su **http://localhost:3000**

Dovresti vedere la homepage "Fantacalcio Manager" con un avviso di primo avvio.

---

## STEP 5: Carica il progetto su GitHub

1. Vai su **https://github.com** → clicca il bottone **+** in alto a destra → **New repository**
2. Nome: `fantacalcio-app`
3. Lascia tutto il resto come sta (NON spuntare "Add README")
4. Clicca **Create repository**
5. Nel terminale, dalla cartella del progetto:

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/TUO_USERNAME/fantacalcio-app.git
git push -u origin main
```

(sostituisci `TUO_USERNAME` con il tuo username GitHub)

---

## STEP 6: Deploy su Vercel (messa online)

1. Vai su **https://vercel.com** e accedi con **GitHub**
2. Clicca **Add New** → **Project**
3. Trova `fantacalcio-app` nella lista dei tuoi repository e clicca **Import**
4. Nella sezione **Environment Variables** aggiungi:
   - `DATABASE_URL` → incolla la connection string con porta 6543
   - `DIRECT_URL` → incolla la connection string con porta 5432
5. Clicca **Deploy**
6. Aspetta ~2 minuti. Quando finisce, Vercel ti darà un URL tipo `fantacalcio-app.vercel.app`

**La tua app è online!** 🎉

---

## STEP 7: Primo utilizzo

1. Vai sull'URL di Vercel → clicca **Admin** nel menu
2. Clicca **Crea le leghe** (Step 0)
3. Carica il **listone quotazioni** (Step 1)
4. Carica le **rose** di ciascuna lega (Step 2)

Fatto! Le rose e le quotazioni sono nel database.

---

## Aggiornamenti futuri

Ogni volta che vuoi aggiornare le quotazioni:
1. Scarica il nuovo listone da Leghe FC
2. Vai su Admin → Step 1 → carica il file
3. I dati si aggiornano automaticamente

Ogni volta che cambia una rosa (post mercato):
1. Scarica le rose da Leghe FC
2. Vai su Admin → Step 2 → carica il file per la lega interessata

---

## Problemi comuni

**"npm: command not found"** → Node.js non è installato. Torna al prerequisito.

**"Error: P1001 Can't reach database"** → La connection string in `.env` è sbagliata. Controlla password e formato.

**La pagina è bianca** → Apri la console del browser (F12 → Console) e dimmi cosa dice.

**Git chiede username/password** → GitHub ora richiede un "Personal Access Token". Vai su GitHub → Settings → Developer Settings → Personal Access Tokens → Tokens (classic) → Generate new token. Usa quello al posto della password.
