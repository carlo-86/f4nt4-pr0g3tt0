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

    // Verify league exists
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
      // Create a default season
      activeSeason = await prisma.season.create({
        data: {
          leagueId,
          label: '2025/2026',
          startDate: new Date('2025-08-23'),
          isActive: true,
        },
      });
    }

    const results: { team: string; players: number; credits: number }[] = [];

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

      // Sync roster: deactivate all current entries, then create/reactivate
      await prisma.rosterEntry.updateMany({
        where: { seasonTeamId: seasonTeam.id },
        data: { isActive: false },
      });

      for (const p of roster.players) {
        // Find player by name in our DB
        // The export uses abbreviated team names, so we match by player name
        const fullTeamName = TEAM_ABBR_MAP[p.teamAbbr] || p.teamAbbr;

        const player = await prisma.player.findFirst({
          where: {
            name: { equals: p.name, mode: 'insensitive' },
          },
        });

        if (!player) {
          // Player not in listone yet — skip or log
          console.warn(`Player not found in listone: ${p.name} (${p.teamAbbr})`);
          continue;
        }

        // Check if roster entry already exists
        const existingEntry = await prisma.rosterEntry.findFirst({
          where: {
            seasonTeamId: seasonTeam.id,
            playerId: player.id,
          },
        });

        if (existingEntry) {
          // Reactivate and update
          await prisma.rosterEntry.update({
            where: { id: existingEntry.id },
            data: {
              isActive: true,
              purchasePrice: p.cost,
            },
          });
        } else {
          // Create new entry
          await prisma.rosterEntry.create({
            data: {
              seasonTeamId: seasonTeam.id,
              playerId: player.id,
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
