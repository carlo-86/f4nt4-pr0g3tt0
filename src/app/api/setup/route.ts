import { NextResponse } from 'next/server';
import { prisma } from '@/lib/prisma';

export async function POST() {
  try {
    // Check if leagues already exist
    const existing = await prisma.league.count();
    if (existing > 0) {
      return NextResponse.json({ error: 'Leagues already exist' }, { status: 400 });
    }

    // Create both leagues
    const fantaTosti = await prisma.league.create({
      data: {
        name: 'Fanta Tosti',
        type: 'CLASSIC',
      },
    });

    const fantaMantra = await prisma.league.create({
      data: {
        name: 'FantaMantra Manageriale',
        type: 'MANTRA',
      },
    });

    // Create active seasons for both
    await prisma.season.create({
      data: {
        leagueId: fantaTosti.id,
        label: '2025/2026',
        startDate: new Date('2025-08-23'),
        isActive: true,
      },
    });

    await prisma.season.create({
      data: {
        leagueId: fantaMantra.id,
        label: '2025/2026',
        startDate: new Date('2025-08-23'),
        isActive: true,
      },
    });

    // Create release penalty configuration
    const penalties = [
      { role: 'P', costCredits: 3 },
      { role: 'D', costCredits: 5 },
      { role: 'C', costCredits: 10 },
      { role: 'A', costCredits: 15 },
    ];

    for (const p of penalties) {
      await prisma.releasePenalty.create({ data: p });
    }

    return NextResponse.json({
      success: true,
      leagues: [
        { id: fantaTosti.id, name: fantaTosti.name },
        { id: fantaMantra.id, name: fantaMantra.name },
      ],
    });
  } catch (error) {
    console.error('Setup error:', error);
    return NextResponse.json(
      { error: 'Setup failed: ' + (error instanceof Error ? error.message : 'Unknown error') },
      { status: 500 }
    );
  }
}
