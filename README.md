# Fantacalcio Manager

Piattaforma web per la gestione di leghe fantacalcio manageriale continuativo (Classic e Mantra).

Sostituisce i database Excel manuali con un'app web che:
- Importa automaticamente quotazioni e FVM dal listone Leghe FC
- Importa e sincronizza le rose dall'export Leghe FC
- Calcola valorizzazioni, assicurazioni e contratti
- Gestisce crediti, mercato e transazioni

## Stack

- **Next.js 14** (App Router)
- **TypeScript**
- **Tailwind CSS**
- **Prisma** + **PostgreSQL** (Supabase)
- **xlsx** (parsing file Excel)

## Setup

Vedi [SETUP_GUIDE.md](./SETUP_GUIDE.md) per la guida completa passo-passo.

## Sviluppo locale

```bash
npm install
npx prisma generate
npx prisma db push
npm run dev
```

Apri http://localhost:3000
