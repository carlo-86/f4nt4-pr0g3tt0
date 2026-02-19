import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/prisma';
import { parseRose, TEAM_ABBR_MAP } from '@/lib/parsers';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    const leagueId = formData.get('leagueId') as string;

    if (!file || !leagueId) {
      return NextResponse.json(
        { error: 'File and leagueId are required' },
        { status: 400 }
      );
    }

    const league = await prisma.league.findUnique({ where: { id: leagueId } });
    if (!league) {
      return NextResponse.json({ error: 'League not found' }, { status: 404 });
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    const teamRosters = parseRose(buffer);

    if (teamRosters.length === 0) {
      return NextResponse.json({ error: 'No teams found in file' }, { status: 400 });
    }

    // Get or create active season
    let activeSeason = await prisma.season.findFirst({
      where: { leagueId, isActive: true },
    });

    if (!activeSeason) {
      activeSeason = await prisma.season.create({
        data: {
          leagueId,
          label: '2025/2026',
          startDate: new Date('2025-08-23'),
          isActive: true,
        },
      });
    }

    // Pre-load ALL data we need in bulk (minimize DB round trips)
    const allPlayers = await prisma.player.findMany({
      select: { id: true, name: true },
    });
    const playerByName = new Map<string, number>();
    for (const p of allPlayers) {
      playerByName.set(p.name.toLowerCase(), p.id);
    }

    const allTeams = await prisma.team.findMany({
      where: { leagueId },
    });
    const teamByName = new Map(allTeams.map(t => [t.name, t.id]));

    const allSeasonTeams = await prisma.seasonTeam.findMany({
      where: { seasonId: activeSeason.id },
      include: { rosterEntries: { select: { id: true, playerId: true } } },
    });
    const seasonTeamByTeamId = new Map(allSeasonTeams.map(st => [st.teamId, st]));

    let nextAutoId = allPlayers.reduce((max, p) => Math.max(max, p.id), 0) + 1000;

    const results: { team: string; players: number; matched: number; autoCreated: string[]; credits: number }[] = [];

    for (const roster of teamRosters) {
      // Get or create team
      let teamId = teamByName.get(roster.teamName);
      if (!teamId) {
        const newTeam = await prisma.team.create({
          data: { name: roster.teamName, leagueId },
        });
        teamId = newTeam.id;
        teamByName.set(roster.teamName, teamId);
      }

      // Get or create season-team
      let seasonTeamData = seasonTeamByTeamId.get(teamId);
      if (!seasonTeamData) {
        const newST = await prisma.seasonTeam.create({
          data: {
            seasonId: activeSeason.id,
            teamId: teamId,
            creditsAvailable: roster.creditsRemaining,
          },
        });
        seasonTeamData = { ...newST, rosterEntries: [] };
        seasonTeamByTeamId.set(teamId, seasonTeamData);
      } else {
        await prisma.seasonTeam.update({
          where: { id: seasonTeamData.id },
          data: { creditsAvailable: roster.creditsRemaining },
        });
      }

      const seasonTeamId = seasonTeamData.id;
      const existingEntries = new Map(
        seasonTeamData.rosterEntries.map(e => [e.playerId, e.id])
      );

      // Deactivate all current entries in one query
      await prisma.rosterEntry.updateMany({
        where: { seasonTeamId },
        data: { isActive: false },
      });

      // Prepare batch operations
      let matched = 0;
      const autoCreated: string[] = [];
      const entriesToCreate: { seasonTeamId: string; playerId: number; purchasePrice: number; purchaseDate: Date; purchaseType: 'AUCTION'; isActive: boolean }[] = [];
      const entriesToReactivate: { id: string; cost: number }[] = [];

      for (const p of roster.players) {
        let playerId = playerByName.get(p.name.toLowerCase());

        if (!playerId) {
          // Auto-create missing player
          const fullTeam = TEAM_ABBR_MAP[p.teamAbbr] || p.teamAbbr;
          const newPlayer = await prisma.player.create({
            data: {
              id: nextAutoId++,
              name: p.name,
              currentTeam: fullTeam,
              roleClassic: p.role.length === 1 ? p.role : null,
              roleMantra: p.role.length > 1 ? p.role : null,
              isActive: false,
            },
          });
          playerId = newPlayer.id;
          playerByName.set(p.name.toLowerCase(), playerId);
          autoCreated.push(p.name);
        } else {
          matched++;
        }

        const existingEntryId = existingEntries.get(playerId);
        if (existingEntryId) {
          entriesToReactivate.push({ id: existingEntryId, cost: p.cost });
        } else {
          entriesToCreate.push({
            seasonTeamId,
            playerId,
            purchasePrice: p.cost,
            purchaseDate: new Date(),
            purchaseType: 'AUCTION',
            isActive: true,
          });
        }
      }

      // Batch create new roster entries
      if (entriesToCreate.length > 0) {
        await prisma.rosterEntry.createMany({
          data: entriesToCreate,
        });
      }

      // Batch reactivate existing entries (in one transaction)
      if (entriesToReactivate.length > 0) {
        await prisma.$transaction(
          entriesToReactivate.map(e =>
            prisma.rosterEntry.update({
              where: { id: e.id },
              data: { isActive: true, purchasePrice: e.cost },
            })
          )
        );
      }

      results.push({
        team: roster.teamName,
        players: roster.players.length,
        matched,
        autoCreated,
        credits: roster.creditsRemaining,
      });
    }

    return NextResponse.json({
      success: true,
      league: league.name,
      teams: results,
    });
  } catch (error) {
    console.error('Import rose error:', error);
    return NextResponse.json(
      { error: 'Import failed: ' + (error instanceof Error ? error.message : 'Unknown error') },
      { status: 500 }
    );
  }
}
