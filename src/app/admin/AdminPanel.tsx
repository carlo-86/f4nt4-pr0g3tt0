'use client';

import { useState } from 'react';
import FileUpload from '@/components/FileUpload';

interface LeagueInfo {
  id: string;
  name: string;
  type: string;
  teamCount: number;
}

interface TeamResult {
  team: string;
  index: number;
  updated: number;
  historicalCreated: number;
  insuranceCreated: number;
  insuranceUpdated: number;
  notFound: string[];
}

function SquadreImport({ league }: { league: LeagueInfo }) {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<'idle' | 'running' | 'done' | 'error'>('idle');
  const [progress, setProgress] = useState('');
  const [results, setResults] = useState<TeamResult[]>([]);
  const [error, setError] = useState('');

  const handleImport = async () => {
    if (!file) return;

    setStatus('running');
    setResults([]);
    setError('');

    try {
      // Step 1: Get team list
      setProgress('Analisi file in corso...');
      const listForm = new FormData();
      listForm.append('file', file);
      listForm.append('leagueId', league.id);

      const listRes = await fetch('/api/import-squadre', { method: 'POST', body: listForm });
      const listData = await listRes.json();

      if (!listRes.ok || !listData.success) {
        throw new Error(listData.error || 'Failed to parse file');
      }

      const teams = listData.teams as { index: number; name: string; players: number }[];

      // Step 2: Process each team sequentially
      const teamResults: TeamResult[] = [];

      for (const team of teams) {
        setProgress(`Importo ${team.name} (${team.index + 1}/${teams.length})...`);

        const teamForm = new FormData();
        teamForm.append('file', file);
        teamForm.append('leagueId', league.id);
        teamForm.append('teamIndex', String(team.index));

        const teamRes = await fetch('/api/import-squadre', { method: 'POST', body: teamForm });
        const teamData = await teamRes.json();

        if (!teamRes.ok || !teamData.success) {
          teamResults.push({
            team: team.name,
            index: team.index,
            updated: 0,
            historicalCreated: 0,
            insuranceCreated: 0,
            insuranceUpdated: 0,
            notFound: [teamData.error || 'Unknown error'],
          });
        } else {
          teamResults.push(teamData as TeamResult);
        }

        setResults([...teamResults]);
      }

      setStatus('done');
      setProgress('Import completato!');
    } catch (err) {
      setStatus('error');
      setError(err instanceof Error ? err.message : 'Unknown error');
    }
  };

  const totalUpdated = results.reduce((s, r) => s + r.updated, 0);
  const totalHistorical = results.reduce((s, r) => s + r.historicalCreated, 0);
  const totalInsCreated = results.reduce((s, r) => s + r.insuranceCreated, 0);
  const totalInsUpdated = results.reduce((s, r) => s + r.insuranceUpdated, 0);

  return (
    <div>
      <div className="flex items-center gap-3 mb-3">
        <label className="px-3 py-1.5 bg-white border border-gray-300 rounded-lg text-sm text-blue-600 hover:bg-gray-50 cursor-pointer">
          Scegli file
          <input
            type="file"
            accept=".xlsx"
            className="hidden"
            onChange={(e) => {
              setFile(e.target.files?.[0] ?? null);
              setStatus('idle');
              setResults([]);
              setError('');
            }}
          />
        </label>
        <span className="text-sm text-gray-600">
          {file ? file.name : 'Nessun file selezionato'}
        </span>
        <button
          onClick={handleImport}
          disabled={!file || status === 'running'}
          className="px-4 py-1.5 bg-blue-600 text-white text-sm rounded-lg hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed"
        >
          {status === 'running' ? 'Importo...' : 'Importa'}
        </button>
      </div>

      {/* Progress */}
      {progress && (
        <p className={`text-sm mb-2 ${status === 'done' ? 'text-green-700 font-semibold' : 'text-gray-600'}`}>
          {progress}
        </p>
      )}

      {/* Error */}
      {error && (
        <div className="bg-red-50 border border-red-200 rounded-lg p-3 mb-3">
          <p className="text-sm text-red-700 font-semibold">Errore</p>
          <p className="text-sm text-red-600">{error}</p>
        </div>
      )}

      {/* Summary */}
      {results.length > 0 && (
        <div className="bg-gray-50 border border-gray-200 rounded-lg p-4 mt-3">
          <div className="grid grid-cols-4 gap-3 text-center mb-3">
            <div>
              <p className="text-xs text-gray-500">Aggiornati</p>
              <p className="text-lg font-bold text-gray-900">{totalUpdated}</p>
            </div>
            <div>
              <p className="text-xs text-gray-500">Storici creati</p>
              <p className="text-lg font-bold text-blue-700">{totalHistorical}</p>
            </div>
            <div>
              <p className="text-xs text-gray-500">Assicurazioni</p>
              <p className="text-lg font-bold text-green-700">{totalInsCreated + totalInsUpdated}</p>
            </div>
            <div>
              <p className="text-xs text-gray-500">Squadre</p>
              <p className="text-lg font-bold text-gray-900">{results.length}</p>
            </div>
          </div>

          {/* Per-team details (collapsible) */}
          <details className="text-sm">
            <summary className="cursor-pointer text-gray-500 hover:text-gray-700">
              Dettagli per squadra
            </summary>
            <div className="mt-2 space-y-2">
              {results.map((r, i) => (
                <div key={i} className="bg-white rounded p-2 border border-gray-100">
                  <p className="font-medium text-gray-800">{r.team}</p>
                  <p className="text-xs text-gray-500">
                    Aggiornati: {r.updated} | Storici: {r.historicalCreated} |
                    Ass. create: {r.insuranceCreated} | Ass. aggiornate: {r.insuranceUpdated}
                  </p>
                  {r.notFound.length > 0 && (
                    <p className="text-xs text-orange-600 mt-1">
                      Non trovati: {r.notFound.join(', ')}
                    </p>
                  )}
                </div>
              ))}
            </div>
          </details>
        </div>
      )}
    </div>
  );
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
          Carica il file &quot;Quotazioni Fantacalcio&quot; scaricato da Leghe
          FC. Aggiorna quotazioni e FVM di tutti i calciatori (per entrambe le
          leghe).
        </p>
        <FileUpload
          endpoint="/api/import-quotazioni"
          label="File quotazioni (.xlsx)"
          accept=".xlsx"
        />
      </section>

      {/* Step 2: Import Rose */}
      {leagues.map((league) => (
        <section
          key={league.id}
          className="bg-white rounded-lg border border-gray-200 p-6"
        >
          <h2 className="text-lg font-semibold text-gray-900 mb-1">
            Step 2 — Importa Rose: {league.name}
            <span
              className={`ml-2 px-2 py-0.5 rounded text-xs font-medium ${
                league.type === 'CLASSIC'
                  ? 'bg-blue-100 text-blue-700'
                  : 'bg-orange-100 text-orange-700'
              }`}
            >
              {league.type}
            </span>
          </h2>
          <p className="text-sm text-gray-500 mb-4">
            Carica il file &quot;Rose&quot; esportato da Leghe FC per{' '}
            {league.name}. Sincronizza le rose di tutte le{' '}
            {league.teamCount || 10} squadre.
          </p>
          <FileUpload
            endpoint="/api/import-rose"
            label={`File rose ${league.name} (.xlsx)`}
            accept=".xlsx"
            extraFields={{ leagueId: league.id }}
          />
        </section>
      ))}

      {/* Step 3: Import Squadre (historical data enrichment) */}
      {leagues.map((league) => (
        <section
          key={`squadre-${league.id}`}
          className="bg-white rounded-lg border border-gray-200 p-6"
        >
          <h2 className="text-lg font-semibold text-gray-900 mb-1">
            Step 3 — Importa Dati Storici: {league.name}
            <span
              className={`ml-2 px-2 py-0.5 rounded text-xs font-medium ${
                league.type === 'CLASSIC'
                  ? 'bg-blue-100 text-blue-700'
                  : 'bg-orange-100 text-orange-700'
              }`}
            >
              {league.type}
            </span>
          </h2>
          <p className="text-sm text-gray-500 mb-4">
            Carica il file &quot;DB completo&quot; Excel di {league.name}.
            Arricchisce le rose già importate con date di acquisto, quotazioni
            storiche e assicurazioni dal foglio SQUADRE. Importa una squadra alla
            volta per rispettare i limiti del server.
          </p>
          <SquadreImport league={league} />
        </section>
      ))}

      {/* Status */}
      <section className="bg-gray-50 rounded-lg border border-gray-200 p-6">
        <h2 className="text-lg font-semibold text-gray-900 mb-3">
          Stato attuale
        </h2>
        <div className="grid grid-cols-2 gap-4 text-sm">
          <div>
            <p className="text-gray-500">Leghe configurate</p>
            <p className="text-2xl font-bold text-gray-900">{leagues.length}</p>
          </div>
          {leagues.map((l) => (
            <div key={l.id}>
              <p className="text-gray-500">{l.name}</p>
              <p className="text-2xl font-bold text-gray-900">
                {l.teamCount} squadre
              </p>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}
