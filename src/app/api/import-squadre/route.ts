import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/prisma';
import { parseSquadre } from '@/lib/parsers';

export const maxDuration = 60;

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;
    const leagueId = formData.get('leagueId') as string;
    const teamIndex = formData.get('teamIndex');

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
    const allParsedTeams = parseSquadre(buffer);

    if (allParsedTeams.length === 0) {
      return NextResponse.json({ error: 'No teams found in file' }, { status: 400 });
    }

    // If no teamIndex, return team list for the frontend to iterate
    if (teamIndex === null || teamIndex === undefined) {
      return NextResponse.json({
        success: true,
        mode: 'list',
        league: league.name,
        teams: allParsedTeams.map((t, i) => ({
          index: i,
          name: t.teamName,
          players: t.players.length,
        })),
      });
    }

    // Process a single team
    const idx = parseInt(String(teamIndex), 10);
    if (isNaN(idx) || idx < 0 || idx >= allParsedTeams.length) {
      return NextResponse.json({ error: 'Invalid teamIndex' }, { status: 400 });
    }

    const parsedTeam = allParsedTeams[idx];

    const activeSeason = await prisma.season.findFirst({
      where: { leagueId, isActive: true },
    });

    if (!activeSeason) {
      return NextResponse.json(
        { error: 'No active season found. Import rose first.' },
        { status: 400 }
      );
    }

    // Load only what we need
    const allPlayers = await prisma.player.findMany({
      select: { id: true, name: true },
    });
    const playerByName = new Map<string, number>();
    for (const p of allPlayers) {
      playerByName.set(p.name.toLowerCase(), p.id);
    }

    const team = await prisma.team.findFirst({
      where: { leagueId, name: parsedTeam.teamName },
    });

    if (!team) {
      return NextResponse.json({
        success: true,
        mode: 'team',
        team: parsedTeam.teamName,
        index: idx,
        updated: 0,
        historicalCreated: 0,
        insuranceCreated: 0,
        insuranceUpdated: 0,
        notFound: [`Team "${parsedTeam.teamName}" not found in league`],
      });
    }

    const seasonTeam = await prisma.seasonTeam.findFirst({
      where: { seasonId: activeSeason.id, teamId: team.id },
      include: {
        rosterEntries: {
          select: {
            id: true,
            playerId: true,
            isActive: true,
            insurance: { select: { id: true } },
          },
        },
      },
    });

    if (!seasonTeam) {
      return NextResponse.json({
        success: true,
        mode: 'team',
        team: parsedTeam.teamName,
        index: idx,
        updated: 0,
        historicalCreated: 0,
        insuranceCreated: 0,
        insuranceUpdated: 0,
        notFound: [`SeasonTeam not found`],
      });
    }

    const rosterByPlayerId = new Map(
      seasonTeam.rosterEntries.map(re => [re.playerId, re])
    );

    let updated = 0;
    let historicalCreated = 0;
    let insuranceCreated = 0;
    let insuranceUpdated = 0;
    const notFound: string[] = [];

    // Collect batch data
    const rosterUpdates: { id: string; price: number; date: Date; quote: number | null; fvm: number | null }[] = [];
    const rosterCreateData: {
      seasonTeamId: string; playerId: number; purchasePrice: number;
      purchaseDate: Date; purchaseType: 'AUCTION'; quoteAtPurchase: number | null;
      fvmPropAtPurchase: number | null; isActive: boolean;
    }[] = [];
    const insCreateData: {
      playerId: number; insured: boolean;
      date: Date; expiry: Date; cost: number;
      qAct: number | null; fAct: number | null;
      qRen: number | null; fRen: number | null;
    }[] = [];
    const insUpdateIds: {
      id: string; date: Date; expiry: Date; cost: number; active: boolean;
      qAct: number | null; fAct: number | null;
      qRen: number | null; fRen: number | null;
    }[] = [];
    const insCreateForExisting: {
      rosterEntryId: string; date: Date; expiry: Date; cost: number; active: boolean;
      qAct: number | null; fAct: number | null;
      qRen: number | null; fRen: number | null;
    }[] = [];

    for (const sp of parsedTeam.players) {
      const playerId = playerByName.get(sp.name.toLowerCase());
      if (!playerId) {
        notFound.push(sp.name);
        continue;
      }

      const pDate = sp.purchaseDate ? new Date(sp.purchaseDate) : new Date();
      const iDate = sp.insuranceDate ? new Date(sp.insuranceDate) : pDate;
      const eDate = new Date(iDate);
      eDate.setFullYear(eDate.getFullYear() + 3);
      const iCost = Math.round(sp.purchasePrice * 0.5);

      const existing = rosterByPlayerId.get(playerId);

      if (existing) {
        rosterUpdates.push({
          id: existing.id, price: sp.purchasePrice, date: pDate,
          quote: sp.quoteAtPurchase, fvm: sp.fvmPropAtPurchase,
        });
        updated++;

        if (sp.insured) {
          if (existing.insurance?.id) {
            insUpdateIds.push({
              id: existing.insurance.id, date: iDate, expiry: eDate,
              cost: iCost, active: existing.isActive,
              qAct: sp.quoteAtPurchase, fAct: sp.fvmPropAtPurchase,
              qRen: sp.quoteRenewal, fRen: sp.fvmPropRenewal,
            });
            insuranceUpdated++;
          } else {
            insCreateForExisting.push({
              rosterEntryId: existing.id, date: iDate, expiry: eDate,
              cost: iCost, active: existing.isActive,
              qAct: sp.quoteAtPurchase, fAct: sp.fvmPropAtPurchase,
              qRen: sp.quoteRenewal, fRen: sp.fvmPropRenewal,
            });
            insuranceCreated++;
          }
        }
      } else {
        rosterCreateData.push({
          seasonTeamId: seasonTeam.id, playerId,
          purchasePrice: sp.purchasePrice, purchaseDate: pDate,
          purchaseType: 'AUCTION', quoteAtPurchase: sp.quoteAtPurchase,
          fvmPropAtPurchase: sp.fvmPropAtPurchase, isActive: false,
        });
        historicalCreated++;

        if (sp.insured) {
          insCreateData.push({
            playerId, insured: true, date: iDate, expiry: eDate, cost: iCost,
            qAct: sp.quoteAtPurchase, fAct: sp.fvmPropAtPurchase,
            qRen: sp.quoteRenewal, fRen: sp.fvmPropRenewal,
          });
          insuranceCreated++;
        }
      }
    }

    // Execute batch operations

    // 1. Update existing roster entries
    if (rosterUpdates.length > 0) {
      await prisma.$transaction(
        rosterUpdates.map(u =>
          prisma.rosterEntry.update({
            where: { id: u.id },
            data: {
              purchasePrice: u.price,
              purchaseDate: u.date,
              quoteAtPurchase: u.quote,
              fvmPropAtPurchase: u.fvm,
            },
          })
        )
      );
    }

    // 2. Create historical roster entries
    if (rosterCreateData.length > 0) {
      await prisma.rosterEntry.createMany({
        data: rosterCreateData,
        skipDuplicates: true,
      });
    }

    // 3. Update existing insurance
    if (insUpdateIds.length > 0) {
      await prisma.$transaction(
        insUpdateIds.map(i =>
          prisma.insurance.update({
            where: { id: i.id },
            data: {
              activationDate: i.date, expiryDate: i.expiry,
              cost: i.cost, isActive: i.active,
              quoteAtActivation: i.qAct, fvmPropAtActivation: i.fAct,
              quoteAtRenewal: i.qRen, fvmPropAtRenewal: i.fRen,
            },
          })
        )
      );
    }

    // 4. Create insurance for existing roster entries
    if (insCreateForExisting.length > 0) {
      await prisma.insurance.createMany({
        data: insCreateForExisting.map(i => ({
          rosterEntryId: i.rosterEntryId,
          activationDate: i.date, expiryDate: i.expiry,
          cost: i.cost, isActive: i.active,
          quoteAtActivation: i.qAct, fvmPropAtActivation: i.fAct,
          quoteAtRenewal: i.qRen, fvmPropAtRenewal: i.fRen,
        })),
      });
    }

    // 5. Create insurance for newly created historical entries
    if (insCreateData.length > 0) {
      const playerIds = insCreateData.map(i => i.playerId);
      const newEntries = await prisma.rosterEntry.findMany({
        where: {
          seasonTeamId: seasonTeam.id,
          playerId: { in: playerIds },
          isActive: false,
        },
        select: { id: true, playerId: true },
      });

      const entryMap = new Map(newEntries.map(e => [e.playerId, e.id]));

      const insRecords = insCreateData
        .map(i => {
          const entryId = entryMap.get(i.playerId);
          if (!entryId) return null;
          return {
            rosterEntryId: entryId,
            activationDate: i.date, expiryDate: i.expiry,
            cost: i.cost, isActive: false,
            quoteAtActivation: i.qAct, fvmPropAtActivation: i.fAct,
            quoteAtRenewal: i.qRen, fvmPropAtRenewal: i.fRen,
          };
        })
        .filter((d): d is NonNullable<typeof d> => d !== null);

      if (insRecords.length > 0) {
        await prisma.insurance.createMany({ data: insRecords });
      }
    }

    // Update credits
    if (parsedTeam.credits !== null) {
      await prisma.seasonTeam.update({
        where: { id: seasonTeam.id },
        data: { creditsAvailable: parsedTeam.credits },
      });
    }

    return NextResponse.json({
      success: true,
      mode: 'team',
      team: parsedTeam.teamName,
      index: idx,
      updated,
      historicalCreated,
      insuranceCreated,
      insuranceUpdated,
      notFound,
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
