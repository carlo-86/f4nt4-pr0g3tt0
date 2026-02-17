import { prisma } from '@/lib/prisma';
import AdminPanel from './AdminPanel';

export default async function AdminPage() {
  const leagues = await prisma.league.findMany({
    include: { teams: true },
  });

  return (
    <div>
      <h1 className="text-2xl font-bold text-gray-900 mb-6">Pannello Admin</h1>
      <AdminPanel leagues={leagues.map(l => ({ id: l.id, name: l.name, type: l.type, teamCount: l.teams.length }))} />
    </div>
  );
}
