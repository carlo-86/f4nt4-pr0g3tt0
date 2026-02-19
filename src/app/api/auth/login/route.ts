import { NextRequest, NextResponse } from 'next/server';

export async function POST(request: NextRequest) {
  const { password } = await request.json();
  const sitePassword = process.env.SITE_PASSWORD;

  if (!sitePassword) {
    return NextResponse.json({ error: 'Password not configured' }, { status: 500 });
  }

  if (password === sitePassword) {
    const response = NextResponse.json({ success: true });
    response.cookies.set('fantacalcio-auth', 'authenticated', {
      httpOnly: true,
      secure: true,
      sameSite: 'lax',
      maxAge: 60 * 60 * 24 * 30, // 30 days
    });
    return response;
  }

  return NextResponse.json({ error: 'Password errata' }, { status: 401 });
}
