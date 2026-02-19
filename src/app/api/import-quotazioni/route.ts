import { NextRequest, NextResponse } from 'next/server';
import { prisma } from '@/lib/prisma';
import { parseQuotazioni } from '@/lib/parsers';

export const maxDuration = 60; // Allow up to 60s on Vercel

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

    // Get all existing player IDs in one query
    const existingPlayers = await prisma.player.findMany({
      select: { id: true },
    });
    const existingIds = new Set(existingPlayers.map(p => p.id));

    // Split into creates and updates
    const toCreate = players.filter(p => !existingIds.has(p.id));
    const toUpdate = players.filter(p => existingIds.has(p.id));

    // Batch create new players
    if (toCreate.length > 0) {
      await prisma.player.createMany({
        data: toCreate.map(p => ({
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
        })),
        skipDuplicates: true,
      });
    }

    // Batch update existing players in chunks of 50
    const CHUNK_SIZE = 50;
    for (let i = 0; i < toUpdate.length; i += CHUNK_SIZE) {
      const chunk = toUpdate.slice(i, i + CHUNK_SIZE);
      await prisma.$transaction(
        chunk.map(p =>
          prisma.player.update({
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
          })
        )
      );
    }

    return NextResponse.json({
      success: true,
      total: players.length,
      created: toCreate.length,
      updated: toUpdate.length,
    });
  } catch (error) {
    console.error('Import quotazioni error:', error);
    return NextResponse.json(
      { error: 'Import failed: ' + (error instanceof Error ? error.message : 'Unknown error') },
      { status: 500 }
    );
  }
}
