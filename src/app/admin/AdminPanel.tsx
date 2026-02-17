'use client';

import { useState } from 'react';
import FileUpload from '@/components/FileUpload';

interface LeagueInfo {
  id: string;
  name: string;
  type: string;
  teamCount: number;
}

export default function AdminPanel({ leagues }: { leagues: LeagueInfo[] }) {
  const [setupDone, setSetupDone] = useState(leagues.length > 0);
  const [setupStatus, setSetupStatus] = useState('');

  const handleSetup = async () => {
    setSetupStatus('Creazione leghe...');
    try {
      const res = await fetch('/api/setup', { method: 'POST' });
      const data = await res.json();
      if (res.ok) {
        setSetupStatus('Leghe create! Ricarica la pagina.');
        setSetupDone(true);
        window.location.reload();
      } else {
        setSetupStatus('Errore: ' + data.error);
      }
    } catch {
      setSetupStatus('Errore di rete');
    }
  };

  return (
    <div className="space-y-8">
      {/* Step 0: Initial setup */}
      {!setupDone && (
        <section className="bg-white rounded-lg border border-gray-200 p-6">
          <h2 className="text-lg font-semibold text-gray-900 mb-3">
            Step 0 — Configurazione iniziale
          </h2>
          <p className="text-sm text-gray-600 mb-4">
            Crea le due leghe (Fanta Tosti Classic e FantaMantra Manageriale).
            Questo va fatto solo la prima volta.
          </p>
          <button
            onClick={handleSetup}
            className="px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700"
          >
            Crea le leghe
          </button>
          {setupStatus && (
            <p className="mt-3 text-sm text-gray-700">{setupStatus}</p>
          )}
        </section>
      )}

      {/* Step 1: Import Quotazioni */}
      <section className="bg-white rounded-lg border border-gray-200 p-6">
        <h2 className="text-lg font-semibold text-gray-900 mb-1">
          Step 1 — Importa Listone Quotazioni
        </h2>
        <p className="text-sm text-gray-500 mb-4">
          Carica il file &quot;Quotazioni Fantacalcio&quot; scaricato da Leghe FC.
          Aggiorna quotazioni e FVM di tutti i calciatori (per entrambe le leghe).
        </p>
        <FileUpload
          endpoint="/api/import-quotazioni"
          label="File quotazioni (.xlsx)"
          accept=".xlsx"
        />
      </section>

      {/* Step 2: Import Rose */}
      {leagues.map((league) => (
        <section key={league.id} className="bg-white rounded-lg border border-gray-200 p-6">
          <h2 className="text-lg font-semibold text-gray-900 mb-1">
            Step 2 — Importa Rose: {league.name}
            <span className={`ml-2 px-2 py-0.5 rounded text-xs font-medium ${
              league.type === 'CLASSIC'
                ? 'bg-blue-100 text-blue-700'
                : 'bg-orange-100 text-orange-700'
            }`}>
              {league.type}
            </span>
          </h2>
          <p className="text-sm text-gray-500 mb-4">
            Carica il file &quot;Rose&quot; esportato da Leghe FC per {league.name}.
            Sincronizza le rose di tutte le {league.teamCount || 10} squadre.
          </p>
          <FileUpload
            endpoint="/api/import-rose"
            label={`File rose ${league.name} (.xlsx)`}
            accept=".xlsx"
            extraFields={{ leagueId: league.id }}
          />
        </section>
      ))}

      {/* Status */}
      <section className="bg-gray-50 rounded-lg border border-gray-200 p-6">
        <h2 className="text-lg font-semibold text-gray-900 mb-3">Stato attuale</h2>
        <div className="grid grid-cols-2 gap-4 text-sm">
          <div>
            <p className="text-gray-500">Leghe configurate</p>
            <p className="text-2xl font-bold text-gray-900">{leagues.length}</p>
          </div>
          {leagues.map((l) => (
            <div key={l.id}>
              <p className="text-gray-500">{l.name}</p>
              <p className="text-2xl font-bold text-gray-900">{l.teamCount} squadre</p>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}
