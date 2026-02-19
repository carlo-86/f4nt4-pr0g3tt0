import { NextRequest, NextResponse } from 'next/server';

export function middleware(request: NextRequest) {
  // Skip auth check for the login page and API login route
  if (
    request.nextUrl.pathname === '/login' ||
    request.nextUrl.pathname === '/api/auth/login'
  ) {
    return NextResponse.next();
  }

  // Check for auth cookie
  const authCookie = request.cookies.get('fantacalcio-auth');
  if (authCookie?.value === 'authenticated') {
    return NextResponse.next();
  }

  // Redirect to login
  const loginUrl = new URL('/login', request.url);
  return NextResponse.redirect(loginUrl);
}

export const config = {
  matcher: ['/((?!_next/static|_next/image|favicon.ico).*)'],
};
