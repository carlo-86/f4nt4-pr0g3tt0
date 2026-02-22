import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/prisma';
import { parseSquadre } from '@/lib/parsers';

export const maxDuration = 60;

// Helper: escape a string for SQL (prevent injection)
function esc(val: string): string {
  return val.replace(/'/g, "''");
}

// Helper: format date for SQL
function sqlDate(d: Date): string {
  return `'${d.toISOString()}'::timestamp`;
}

// Helper: format nullable number
function sqlNum(v: number | null): string {
  return v === null ? 'NULL' : String(v);
}

// Helper: format nullable float
function sqlFloat(v: number | null | undefined): string {
  if (v === null || v === undefined) return 'NULL';
  return String(v);
}

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

    // Pre-load all data (3 queries)
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

    // ---- Collect operations ----

    const rosterUpdates: {
      id: string; price: number; date: Date;
      quote: number | null; fvm: number | null;
    }[] = [];

    const rosterCreates: {
      seasonTeamId: string; playerId: number; price: number; date: Date;
      quote: number | null; fvm: number | null;
    }[] = [];

    const insUpdates: {
      id: string; date: Date; expiry: Date; cost: number; active: boolean;
      qAct: number | null; fAct: number | null;
      qRen: number | null; fRen: number | null;
    }[] = [];

    const insCreatesForExisting: {
      rosterEntryId: string; date: Date; expiry: Date; cost: number; active: boolean;
      qAct: number | null; fAct: number | null;
      qRen: number | null; fRen: number | null;
    }[] = [];

    const insCreatesForNew: {
      playerId: number; seasonTeamId: string;
      date: Date; expiry: Date; cost: number;
      qAct: number | null; fAct: number | null;
      qRen: number | null; fRen: number | null;
    }[] = [];

    const results: {
      team: string; playersInFile: number; updated: number;
      historicalCreated: number; insuranceCreated: number;
      insuranceUpdated: number; notFound: string[];
    }[] = [];

    for (const parsedTeam of parsedTeams) {
      const stats = {
        playersInFile: parsedTeam.players.length,
        updated: 0, historicalCreated: 0,
        insuranceCreated: 0, insuranceUpdated: 0,
        notFound: [] as string[],
      };

      const teamId = teamByName.get(parsedTeam.teamName);
      const seasonTeamData = teamId ? seasonTeamByTeamId.get(teamId) : null;

      if (!teamId || !seasonTeamData) {
        stats.notFound.push(`Team/SeasonTeam not found`);
        results.push({ team: parsedTeam.teamName, ...stats });
        continue;
      }

      const rosterByPlayerId = new Map(
        seasonTeamData.rosterEntries.map(re => [re.playerId, re])
      );

      for (const sp of parsedTeam.players) {
        const playerId = playerByName.get(sp.name.toLowerCase());
        if (!playerId) { stats.notFound.push(sp.name); continue; }

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
          stats.updated++;

          if (sp.insured) {
            if (existing.insurance?.id) {
              insUpdates.push({
                id: existing.insurance.id, date: iDate, expiry: eDate,
                cost: iCost, active: existing.isActive,
                qAct: sp.quoteAtPurchase, fAct: sp.fvmPropAtPurchase,
                qRen: sp.quoteRenewal, fRen: sp.fvmPropRenewal,
              });
              stats.insuranceUpdated++;
            } else {
              insCreatesForExisting.push({
                rosterEntryId: existing.id, date: iDate, expiry: eDate,
                cost: iCost, active: existing.isActive,
                qAct: sp.quoteAtPurchase, fAct: sp.fvmPropAtPurchase,
                qRen: sp.quoteRenewal, fRen: sp.fvmPropRenewal,
              });
              stats.insuranceCreated++;
            }
          }
        } else {
          rosterCreates.push({
            seasonTeamId: seasonTeamData.id, playerId,
            price: sp.purchasePrice, date: pDate,
            quote: sp.quoteAtPurchase, fvm: sp.fvmPropAtPurchase,
          });
          stats.historicalCreated++;

          if (sp.insured) {
            insCreatesForNew.push({
              playerId, seasonTeamId: seasonTeamData.id,
              date: iDate, expiry: eDate, cost: iCost,
              qAct: sp.quoteAtPurchase, fAct: sp.fvmPropAtPurchase,
              qRen: sp.quoteRenewal, fRen: sp.fvmPropRenewal,
            });
            stats.insuranceCreated++;
          }
        }
      }

      if (parsedTeam.credits !== null) {
        await prisma.seasonTeam.update({
          where: { id: seasonTeamData.id },
          data: { creditsAvailable: parsedTeam.credits },
        });
      }

      results.push({ team: parsedTeam.teamName, ...stats });
    }

    // ---- EXECUTE: Single SQL statements for bulk operations ----

    // 1. Bulk UPDATE roster entries (1 SQL statement)
    if (rosterUpdates.length > 0) {
      const values = rosterUpdates.map(u =>
        `('${esc(u.id)}', ${u.price}, ${sqlDate(u.date)}, ${sqlNum(u.quote)}, ${sqlFloat(u.fvm)})`
      ).join(',\n');

      await prisma.$executeRawUnsafe(`
        UPDATE "RosterEntry" AS r
        SET "purchasePrice" = v.price::int,
            "purchaseDate" = v.pdate::timestamp,
            "quoteAtPurchase" = v.quote::int,
            "fvmPropAtPurchase" = v.fvm::double precision
        FROM (VALUES ${values})
          AS v(id, price, pdate, quote, fvm)
        WHERE r."id" = v.id
      `);
    }

    // 2. Bulk CREATE historical roster entries (1 query)
    if (rosterCreates.length > 0) {
      await prisma.rosterEntry.createMany({
        data: rosterCreates.map(c => ({
          seasonTeamId: c.seasonTeamId,
          playerId: c.playerId,
          purchasePrice: c.price,
          purchaseDate: c.date,
          purchaseType: 'AUCTION' as const,
          quoteAtPurchase: c.quote,
          fvmPropAtPurchase: c.fvm,
          isActive: false,
        })),
        skipDuplicates: true,
      });
    }

    // 3. Bulk UPDATE existing insurance (1 SQL statement)
    if (insUpdates.length > 0) {
      const values = insUpdates.map(i =>
        `('${esc(i.id)}', ${sqlDate(i.date)}, ${sqlDate(i.expiry)}, ${i.cost}, ${i.active}, ${sqlNum(i.qAct)}, ${sqlFloat(i.fAct)}, ${sqlNum(i.qRen)}, ${sqlFloat(i.fRen)})`
      ).join(',\n');

      await prisma.$executeRawUnsafe(`
        UPDATE "Insurance" AS ins
        SET "activationDate" = v.adate::timestamp,
            "expiryDate" = v.edate::timestamp,
            "cost" = v.cost::double precision,
            "isActive" = v.active::boolean,
            "quoteAtActivation" = v.qact::int,
            "fvmPropAtActivation" = v.fact::double precision,
            "quoteAtRenewal" = v.qren::int,
            "fvmPropAtRenewal" = v.fren::double precision
        FROM (VALUES ${values})
          AS v(id, adate, edate, cost, active, qact, fact, qren, fren)
        WHERE ins."id" = v.id
      `);
    }

    // 4. Bulk CREATE insurance for existing roster entries (1 query)
    if (insCreatesForExisting.length > 0) {
      await prisma.insurance.createMany({
        data: insCreatesForExisting.map(i => ({
          rosterEntryId: i.rosterEntryId,
          activationDate: i.date,
          expiryDate: i.expiry,
          cost: i.cost,
          isActive: i.active,
          quoteAtActivation: i.qAct,
          fvmPropAtActivation: i.fAct,
          quoteAtRenewal: i.qRen,
          fvmPropAtRenewal: i.fRen,
        })),
      });
    }

    // 5. Insurance for newly created historical entries (2 queries: fetch + createMany)
    if (insCreatesForNew.length > 0) {
      const playerIds = insCreatesForNew.map(i => i.playerId);
      const stIds = [...new Set(insCreatesForNew.map(i => i.seasonTeamId))];

      const newEntries = await prisma.rosterEntry.findMany({
        where: {
          seasonTeamId: { in: stIds },
          playerId: { in: playerIds },
          isActive: false,
        },
        select: { id: true, playerId: true, seasonTeamId: true },
      });

      const entryMap = new Map<string, string>();
      for (const e of newEntries) {
        entryMap.set(`${e.seasonTeamId}-${e.playerId}`, e.id);
      }

      const insData = insCreatesForNew
        .map(i => {
          const entryId = entryMap.get(`${i.seasonTeamId}-${i.playerId}`);
          if (!entryId) return null;
          return {
            rosterEntryId: entryId,
            activationDate: i.date,
            expiryDate: i.expiry,
            cost: i.cost,
            isActive: false,
            quoteAtActivation: i.qAct,
            fvmPropAtActivation: i.fAct,
            quoteAtRenewal: i.qRen,
            fvmPropAtRenewal: i.fRen,
          };
        })
        .filter((d): d is NonNullable<typeof d> => d !== null);

      if (insData.length > 0) {
        await prisma.insurance.createMany({ data: insData });
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
