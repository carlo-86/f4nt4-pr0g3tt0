import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/prisma';
import { parseQuotazioni } from '@/lib/parsers';

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const file = formData.get('file') as File;

    if (!file) {
      return NextResponse.json({ error: 'No file provided' }, { status: 400 });
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    const players = parseQuotazioni(buffer);

    if (players.length === 0) {
      return NextResponse.json({ error: 'No players found in file' }, { status: 400 });
    }

    // Upsert all players
    let created = 0;
    let updated = 0;

    for (const p of players) {
      const existing = await prisma.player.findUnique({ where: { id: p.id } });

      if (existing) {
        await prisma.player.update({
          where: { id: p.id },
          data: {
            name: p.name,
            currentTeam: p.team,
            roleClassic: p.roleClassic,
            roleMantra: p.roleMantra,
            quoteClassic: p.quoteClassic,
            quoteInitClassic: p.quoteInitClassic,
            quoteMantra: p.quoteMantra,
            quoteInitMantra: p.quoteInitMantra,
            fvm: p.fvm,
            fvmMantra: p.fvmMantra,
            isActive: p.isActive,
          },
        });
        updated++;
      } else {
        await prisma.player.create({
          data: {
            id: p.id,
            name: p.name,
            currentTeam: p.team,
            roleClassic: p.roleClassic,
            roleMantra: p.roleMantra,
            quoteClassic: p.quoteClassic,
            quoteInitClassic: p.quoteInitClassic,
            quoteMantra: p.quoteMantra,
            quoteInitMantra: p.quoteInitMantra,
            fvm: p.fvm,
            fvmMantra: p.fvmMantra,
            isActive: p.isActive,
          },
        });
        created++;
      }
    }

    return NextResponse.json({
      success: true,
      total: players.length,
      created,
      updated,
    });
  } catch (error) {
    console.error('Import quotazioni error:', error);
    return NextResponse.json(
      { error: 'Import failed: ' + (error instanceof Error ? error.message : 'Unknown error') },
      { status: 500 }
    );
  }
}
