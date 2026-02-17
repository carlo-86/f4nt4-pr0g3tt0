import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'Fantacalcio Manager',
  description: 'Gestione leghe fantacalcio manageriale continuativo',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="it">
      <body className="min-h-screen bg-gray-50">
        <nav className="bg-white border-b border-gray-200 px-6 py-3">
          <div className="max-w-7xl mx-auto flex items-center justify-between">
            <a href="/" className="text-xl font-bold text-blue-700">
              ⚽ Fantacalcio Manager
            </a>
            <div className="flex gap-4 text-sm">
              <a href="/listone" className="text-gray-600 hover:text-blue-600">Listone</a>
              <a href="/squadre" className="text-gray-600 hover:text-blue-600">Squadre</a>
              <a href="/rosa" className="text-gray-600 hover:text-blue-600">Rose</a>
              <a href="/admin" className="text-gray-600 hover:text-blue-600 font-medium">Admin</a>
            </div>
          </div>
        </nav>
        <main className="max-w-7xl mx-auto px-6 py-8">
          {children}
        </main>
      </body>
    </html>
  );
}
