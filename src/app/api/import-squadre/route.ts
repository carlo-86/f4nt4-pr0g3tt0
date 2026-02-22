import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/prisma';
import { parseSquadre } from '@/lib/parsers';

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
    const parsedTeams = parseSquadre(buffer);

    if (parsedTeams.length === 0) {
      return NextResponse.json({ error: 'No teams found in file' }, { status: 400 });
    }

    // Get active season (must exist — import rose first)
    const activeSeason = await prisma.season.findFirst({
      where: { leagueId, isActive: true },
    });

    if (!activeSeason) {
      return NextResponse.json(
        { error: 'No active season found. Import rose first.' },
        { status: 400 }
      );
    }

    // Pre-load all data in bulk
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
      include: {
        rosterEntries: {
          where: { isActive: true },
          include: { insurance: { select: { id: true } } },
        },
      },
    });
    const seasonTeamByTeamId = new Map(allSeasonTeams.map(st => [st.teamId, st]));

    const results: {
      team: string;
      playersInFile: number;
      updated: number;
      insuranceCreated: number;
      insuranceUpdated: number;
      notMatched: string[];
    }[] = [];

    for (const parsedTeam of parsedTeams) {
      const teamId = teamByName.get(parsedTeam.teamName);
      if (!teamId) {
        results.push({
          team: parsedTeam.teamName,
          playersInFile: parsedTeam.players.length,
          updated: 0,
          insuranceCreated: 0,
          insuranceUpdated: 0,
          notMatched: [`Team "${parsedTeam.teamName}" not found in league`],
        });
        continue;
      }

      const seasonTeamData = seasonTeamByTeamId.get(teamId);
      if (!seasonTeamData) {
        results.push({
          team: parsedTeam.teamName,
          playersInFile: parsedTeam.players.length,
          updated: 0,
          insuranceCreated: 0,
          insuranceUpdated: 0,
          notMatched: [`SeasonTeam not found for "${parsedTeam.teamName}"`],
        });
        continue;
      }

      // Map playerId -> rosterEntry for this team
      const rosterByPlayerId = new Map(
        seasonTeamData.rosterEntries.map(re => [re.playerId, re])
      );

      let updated = 0;
      let insuranceCreated = 0;
      let insuranceUpdated = 0;
      const notMatched: string[] = [];

      for (const sp of parsedTeam.players) {
        // Match player by name (case-insensitive)
        const playerId = playerByName.get(sp.name.toLowerCase());
        if (!playerId) {
          notMatched.push(sp.name);
          continue;
        }

        // Find existing roster entry
        const rosterEntry = rosterByPlayerId.get(playerId);
        if (!rosterEntry) {
          notMatched.push(`${sp.name} (no roster entry)`);
          continue;
        }

        // Enrich roster entry with historical data from SQUADRE
        const purchaseDate = sp.purchaseDate ? new Date(sp.purchaseDate) : new Date();

        await prisma.rosterEntry.update({
          where: { id: rosterEntry.id },
          data: {
            purchasePrice: sp.purchasePrice,
            purchaseDate: purchaseDate,
            quoteAtPurchase: sp.quoteAtPurchase,
            fvmPropAtPurchase: sp.fvmPropAtPurchase,
          },
        });
        updated++;

        // Handle insurance records
        if (sp.insured) {
          const insuranceDate = sp.insuranceDate
            ? new Date(sp.insuranceDate)
            : purchaseDate;

          // Expiry = activation date + 3 years
          const expiryDate = new Date(insuranceDate);
          expiryDate.setFullYear(expiryDate.getFullYear() + 3);

          // Cost = 50% of purchase price (standard rule)
          const insuranceCost = Math.round(sp.purchasePrice * 0.5);

          const existingInsurance = rosterEntry.insurance;

          if (existingInsurance) {
            await prisma.insurance.update({
              where: { id: existingInsurance.id },
              data: {
                activationDate: insuranceDate,
                expiryDate: expiryDate,
                cost: insuranceCost,
                isActive: true,
                quoteAtActivation: sp.quoteAtPurchase,
                fvmPropAtActivation: sp.fvmPropAtPurchase,
                quoteAtRenewal: sp.quoteRenewal,
                fvmPropAtRenewal: sp.fvmPropRenewal,
              },
            });
            insuranceUpdated++;
          } else {
            await prisma.insurance.create({
              data: {
                rosterEntryId: rosterEntry.id,
                activationDate: insuranceDate,
                expiryDate: expiryDate,
                cost: insuranceCost,
                isActive: true,
                quoteAtActivation: sp.quoteAtPurchase,
                fvmPropAtActivation: sp.fvmPropAtPurchase,
                quoteAtRenewal: sp.quoteRenewal,
                fvmPropAtRenewal: sp.fvmPropRenewal,
              },
            });
            insuranceCreated++;
          }
        }
      }

      // Update team credits if present in the file
      if (parsedTeam.credits !== null) {
        await prisma.seasonTeam.update({
          where: { id: seasonTeamData.id },
          data: { creditsAvailable: parsedTeam.credits },
        });
      }

      results.push({
        team: parsedTeam.teamName,
        playersInFile: parsedTeam.players.length,
        updated,
        insuranceCreated,
        insuranceUpdated,
        notMatched,
      });
    }

    const totalUpdated = results.reduce((s, r) => s + r.updated, 0);
    const totalInsurance = results.reduce(
      (s, r) => s + r.insuranceCreated + r.insuranceUpdated, 0
    );

    return NextResponse.json({
      success: true,
      league: league.name,
      summary: {
        teams: results.length,
        rosterEntriesUpdated: totalUpdated,
        insuranceRecords: totalInsurance,
      },
      teams: results,
    });
  } catch (error) {
    console.error('Import squadre error:', error);
    return NextResponse.json(
      {
        error:
          'Import failed: ' +
          (error instanceof Error ? error.message : 'Unknown error'),
      },
      { status: 500 }
    );
  }
}
