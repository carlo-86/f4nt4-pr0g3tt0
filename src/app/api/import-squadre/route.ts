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

    const allTeams = await prisma.team.findMany({ where: { leagueId } });
    const teamByName = new Map(allTeams.map(t => [t.name, t.id]));

    const allSeasonTeams = await prisma.seasonTeam.findMany({
      where: { seasonId: activeSeason.id },
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
    const seasonTeamByTeamId = new Map(allSeasonTeams.map(st => [st.teamId, st]));

    // ---- PHASE 1: Collect all operations per team ----

    interface UpdateOp {
      rosterEntryId: string;
      insuranceId: string | null;
      isActive: boolean;
      purchasePrice: number;
      purchaseDate: Date;
      quoteAtPurchase: number | null;
      fvmPropAtPurchase: number | null;
      insured: boolean;
      insuranceDate: Date;
      expiryDate: Date;
      insuranceCost: number;
      quoteRenewal: number | null;
      fvmPropRenewal: number | null;
    }

    interface CreateOp {
      seasonTeamId: string;
      playerId: number;
      purchasePrice: number;
      purchaseDate: Date;
      quoteAtPurchase: number | null;
      fvmPropAtPurchase: number | null;
      insured: boolean;
      insuranceDate: Date;
      expiryDate: Date;
      insuranceCost: number;
      quoteRenewal: number | null;
      fvmPropRenewal: number | null;
    }

    const results: {
      team: string;
      playersInFile: number;
      updated: number;
      historicalCreated: number;
      insuranceCreated: number;
      insuranceUpdated: number;
      notFound: string[];
    }[] = [];

    const allUpdates: UpdateOp[] = [];
    const allCreates: CreateOp[] = [];
    const teamStats = new Map<string, {
      playersInFile: number;
      updated: number;
      historicalCreated: number;
      insuranceCreated: number;
      insuranceUpdated: number;
      notFound: string[];
    }>();

    for (const parsedTeam of parsedTeams) {
      const stats = {
        playersInFile: parsedTeam.players.length,
        updated: 0,
        historicalCreated: 0,
        insuranceCreated: 0,
        insuranceUpdated: 0,
        notFound: [] as string[],
      };

      const teamId = teamByName.get(parsedTeam.teamName);
      if (!teamId) {
        stats.notFound.push(`Team "${parsedTeam.teamName}" not found`);
        results.push({ team: parsedTeam.teamName, ...stats });
        continue;
      }

      const seasonTeamData = seasonTeamByTeamId.get(teamId);
      if (!seasonTeamData) {
        stats.notFound.push(`SeasonTeam not found for "${parsedTeam.teamName}"`);
        results.push({ team: parsedTeam.teamName, ...stats });
        continue;
      }

      const rosterByPlayerId = new Map(
        seasonTeamData.rosterEntries.map(re => [re.playerId, re])
      );

      for (const sp of parsedTeam.players) {
        const playerId = playerByName.get(sp.name.toLowerCase());
        if (!playerId) {
          stats.notFound.push(sp.name);
          continue;
        }

        const purchaseDate = sp.purchaseDate ? new Date(sp.purchaseDate) : new Date();
        const insuranceDate = sp.insuranceDate
          ? new Date(sp.insuranceDate)
          : purchaseDate;
        const expiryDate = new Date(insuranceDate);
        expiryDate.setFullYear(expiryDate.getFullYear() + 3);
        const insuranceCost = Math.round(sp.purchasePrice * 0.5);

        const existing = rosterByPlayerId.get(playerId);

        if (existing) {
          allUpdates.push({
            rosterEntryId: existing.id,
            insuranceId: existing.insurance?.id ?? null,
            isActive: existing.isActive,
            purchasePrice: sp.purchasePrice,
            purchaseDate,
            quoteAtPurchase: sp.quoteAtPurchase,
            fvmPropAtPurchase: sp.fvmPropAtPurchase,
            insured: sp.insured,
            insuranceDate,
            expiryDate,
            insuranceCost,
            quoteRenewal: sp.quoteRenewal,
            fvmPropRenewal: sp.fvmPropRenewal,
          });
          stats.updated++;
          if (sp.insured) {
            if (existing.insurance?.id) stats.insuranceUpdated++;
            else stats.insuranceCreated++;
          }
        } else {
          allCreates.push({
            seasonTeamId: seasonTeamData.id,
            playerId,
            purchasePrice: sp.purchasePrice,
            purchaseDate,
            quoteAtPurchase: sp.quoteAtPurchase,
            fvmPropAtPurchase: sp.fvmPropAtPurchase,
            insured: sp.insured,
            insuranceDate,
            expiryDate,
            insuranceCost,
            quoteRenewal: sp.quoteRenewal,
            fvmPropRenewal: sp.fvmPropRenewal,
          });
          stats.historicalCreated++;
          if (sp.insured) stats.insuranceCreated++;
        }
      }

      // Update credits
      if (parsedTeam.credits !== null) {
        await prisma.seasonTeam.update({
          where: { id: seasonTeamData.id },
          data: { creditsAvailable: parsedTeam.credits },
        });
      }

      teamStats.set(parsedTeam.teamName, stats);
      results.push({ team: parsedTeam.teamName, ...stats });
    }

    // ---- PHASE 2: Execute batch updates in a single transaction ----

    const BATCH_SIZE = 25;

    // Batch update existing roster entries
    for (let i = 0; i < allUpdates.length; i += BATCH_SIZE) {
      const batch = allUpdates.slice(i, i + BATCH_SIZE);
      await prisma.$transaction(
        batch.map(op =>
          prisma.rosterEntry.update({
            where: { id: op.rosterEntryId },
            data: {
              purchasePrice: op.purchasePrice,
              purchaseDate: op.purchaseDate,
              quoteAtPurchase: op.quoteAtPurchase,
              fvmPropAtPurchase: op.fvmPropAtPurchase,
            },
          })
        )
      );
    }

    // Batch update/create insurance for existing entries
    for (let i = 0; i < allUpdates.length; i += BATCH_SIZE) {
      const batch = allUpdates.slice(i, i + BATCH_SIZE).filter(op => op.insured);
      if (batch.length === 0) continue;

      const toUpdate = batch.filter(op => op.insuranceId);
      const toCreate = batch.filter(op => !op.insuranceId);

      if (toUpdate.length > 0) {
        await prisma.$transaction(
          toUpdate.map(op =>
            prisma.insurance.update({
              where: { id: op.insuranceId! },
              data: {
                activationDate: op.insuranceDate,
                expiryDate: op.expiryDate,
                cost: op.insuranceCost,
                isActive: op.isActive,
                quoteAtActivation: op.quoteAtPurchase,
                fvmPropAtActivation: op.fvmPropAtPurchase,
                quoteAtRenewal: op.quoteRenewal,
                fvmPropAtRenewal: op.fvmPropRenewal,
              },
            })
          )
        );
      }

      if (toCreate.length > 0) {
        await prisma.insurance.createMany({
          data: toCreate.map(op => ({
            rosterEntryId: op.rosterEntryId,
            activationDate: op.insuranceDate,
            expiryDate: op.expiryDate,
            cost: op.insuranceCost,
            isActive: op.isActive,
            quoteAtActivation: op.quoteAtPurchase,
            fvmPropAtActivation: op.fvmPropAtPurchase,
            quoteAtRenewal: op.quoteRenewal,
            fvmPropAtRenewal: op.fvmPropRenewal,
          })),
        });
      }
    }

    // ---- PHASE 3: Batch create historical entries ----

    // Create roster entries in batches
    for (let i = 0; i < allCreates.length; i += BATCH_SIZE) {
      const batch = allCreates.slice(i, i + BATCH_SIZE);
      await prisma.rosterEntry.createMany({
        data: batch.map(op => ({
          seasonTeamId: op.seasonTeamId,
          playerId: op.playerId,
          purchasePrice: op.purchasePrice,
          purchaseDate: op.purchaseDate,
          purchaseType: 'AUCTION' as const,
          quoteAtPurchase: op.quoteAtPurchase,
          fvmPropAtPurchase: op.fvmPropAtPurchase,
          isActive: false,
        })),
      });
    }

    // Now fetch back the created entries to get their IDs for insurance
    const insuredCreates = allCreates.filter(op => op.insured);
    if (insuredCreates.length > 0) {
      // Query the entries we just created (inactive ones matching our playerIds)
      const playerIds = insuredCreates.map(op => op.playerId);
      const seasonTeamIds = [...new Set(insuredCreates.map(op => op.seasonTeamId))];

      const createdEntries = await prisma.rosterEntry.findMany({
        where: {
          seasonTeamId: { in: seasonTeamIds },
          playerId: { in: playerIds },
          isActive: false,
        },
        select: { id: true, playerId: true, seasonTeamId: true },
      });

      // Map (seasonTeamId, playerId) -> rosterEntryId
      const entryMap = new Map<string, string>();
      for (const e of createdEntries) {
        entryMap.set(`${e.seasonTeamId}-${e.playerId}`, e.id);
      }

      // Batch create insurance records
      const insuranceData = insuredCreates
        .map(op => {
          const entryId = entryMap.get(`${op.seasonTeamId}-${op.playerId}`);
          if (!entryId) return null;
          return {
            rosterEntryId: entryId,
            activationDate: op.insuranceDate,
            expiryDate: op.expiryDate,
            cost: op.insuranceCost,
            isActive: false,
            quoteAtActivation: op.quoteAtPurchase,
            fvmPropAtActivation: op.fvmPropAtPurchase,
            quoteAtRenewal: op.quoteRenewal,
            fvmPropAtRenewal: op.fvmPropRenewal,
          };
        })
        .filter((d): d is NonNullable<typeof d> => d !== null);

      if (insuranceData.length > 0) {
        for (let i = 0; i < insuranceData.length; i += BATCH_SIZE) {
          await prisma.insurance.createMany({
            data: insuranceData.slice(i, i + BATCH_SIZE),
          });
        }
      }
    }

    const totalUpdated = results.reduce((s, r) => s + r.updated, 0);
    const totalHistorical = results.reduce((s, r) => s + r.historicalCreated, 0);
    const totalInsurance = results.reduce(
      (s, r) => s + r.insuranceCreated + r.insuranceUpdated, 0
    );

    return NextResponse.json({
      success: true,
      league: league.name,
      summary: {
        teams: results.length,
        rosterEntriesUpdated: totalUpdated,
        historicalEntriesCreated: totalHistorical,
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
