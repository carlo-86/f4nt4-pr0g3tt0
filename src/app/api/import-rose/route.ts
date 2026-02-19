import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/prisma';
import { parseRose, TEAM_ABBR_MAP } from '@/lib/parsers';

export const maxDuration = 60;

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

    // Pre-load all players for fast lookup by name
    const allPlayers = await prisma.player.findMany({
      select: { id: true, name: true },
    });
    const playerByName = new Map<string, number>();
    for (const p of allPlayers) {
      playerByName.set(p.name.toLowerCase(), p.id);
    }

    // Track the next available ID for auto-created players
    const maxId = allPlayers.reduce((max, p) => Math.max(max, p.id), 0);
    let nextAutoId = maxId + 1000; // Start well above existing IDs

    const results: { team: string; players: number; matched: number; autoCreated: string[]; credits: number }[] = [];

    for (const roster of teamRosters) {
      // Get or create team
      let team = await prisma.team.findFirst({
        where: { leagueId, name: roster.teamName },
      });

      if (!team) {
        team = await prisma.team.create({
          data: { name: roster.teamName, leagueId },
        });
      }

      // Get or create season-team link
      let seasonTeam = await prisma.seasonTeam.findFirst({
        where: { seasonId: activeSeason.id, teamId: team.id },
      });

      if (!seasonTeam) {
        seasonTeam = await prisma.seasonTeam.create({
          data: {
            seasonId: activeSeason.id,
            teamId: team.id,
            creditsAvailable: roster.creditsRemaining,
          },
        });
      } else {
        await prisma.seasonTeam.update({
          where: { id: seasonTeam.id },
          data: { creditsAvailable: roster.creditsRemaining },
        });
      }

      // Deactivate all current roster entries
      await prisma.rosterEntry.updateMany({
        where: { seasonTeamId: seasonTeam.id },
        data: { isActive: false },
      });

      let matched = 0;
      const autoCreated: string[] = [];

      for (const p of roster.players) {
        let playerId = playerByName.get(p.name.toLowerCase());

        if (!playerId) {
          // Auto-create player as "ceduto" (not in official listone)
          const fullTeam = TEAM_ABBR_MAP[p.teamAbbr] || p.teamAbbr;
          const newPlayer = await prisma.player.create({
            data: {
              id: nextAutoId++,
              name: p.name,
              currentTeam: fullTeam,
              roleClassic: p.role.length === 1 ? p.role : null, // P, D, C, A
              roleMantra: p.role.length > 1 ? p.role : null,    // Por, Dc, etc.
              isActive: false, // Mark as not in official listone
            },
          });
          playerId = newPlayer.id;
          playerByName.set(p.name.toLowerCase(), playerId);
          autoCreated.push(p.name);
        } else {
          matched++;
        }

        // Check if roster entry already exists
        const existingEntry = await prisma.rosterEntry.findFirst({
          where: {
            seasonTeamId: seasonTeam.id,
            playerId: playerId,
          },
        });

        if (existingEntry) {
          await prisma.rosterEntry.update({
            where: { id: existingEntry.id },
            data: {
              isActive: true,
              purchasePrice: p.cost,
            },
          });
        } else {
          await prisma.rosterEntry.create({
            data: {
              seasonTeamId: seasonTeam.id,
              playerId: playerId,
              purchasePrice: p.cost,
              purchaseDate: new Date(),
              purchaseType: 'AUCTION',
              isActive: true,
            },
          });
        }
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
