export const dynamic = 'force-dynamic';

import { prisma } from '@/lib/prisma';

export default async function Home() {
  const leagues = await prisma.league.findMany({
    include: {
      teams: true,
      seasons: { where: { isActive: true } },
    },
  });

  const playerCount = await prisma.player.count();

  return (
    <div>
      <h1 className="text-3xl font-bold text-gray-900 mb-2">Fantacalcio Manager</h1>
      <p className="text-gray-500 mb-8">
        Gestione leghe fantacalcio manageriale continuativo
      </p>

      {leagues.length === 0 ? (
        <div className="bg-amber-50 border border-amber-200 rounded-lg p-6">
          <h2 className="text-lg font-semibold text-amber-800 mb-2">Primo avvio</h2>
          <p className="text-amber-700 mb-4">
            Nessuna lega configurata. Vai al pannello Admin per importare i dati.
          </p>
          <a
            href="/admin"
            className="inline-block bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700"
          >
            Vai al pannello Admin →
          </a>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {leagues.map((league) => (
            <div
              key={league.id}
              className="bg-white rounded-lg border border-gray-200 p-6 shadow-sm"
            >
              <div className="flex items-center gap-3 mb-4">
                <span className={`px-2 py-1 rounded text-xs font-medium ${
                  league.type === 'CLASSIC'
                    ? 'bg-blue-100 text-blue-700'
                    : 'bg-orange-100 text-orange-700'
                }`}>
                  {league.type}
                </span>
                <h2 className="text-xl font-bold text-gray-900">{league.name}</h2>
              </div>

              <div className="space-y-2 text-sm text-gray-600">
                <p>Squadre: <span className="font-medium text-gray-900">{league.teams.length}</span></p>
                <p>Stagione attiva: <span className="font-medium text-gray-900">
                  {league.seasons[0]?.label || 'Nessuna'}
                </span></p>
              </div>

              <div className="mt-4 flex gap-2">
                <a
                  href={`/squadre?league=${league.id}`}
                  className="text-sm text-blue-600 hover:text-blue-800"
                >
                  Vedi squadre →
                </a>
              </div>
            </div>
          ))}
        </div>
      )}

      {playerCount > 0 && (
        <p className="mt-6 text-sm text-gray-400">
          {playerCount} calciatori nel listone
        </p>
      )}
    </div>
  );
}
