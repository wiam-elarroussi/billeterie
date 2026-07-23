import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'E-Ticket Pro — Solution Globale de Billetterie & Contrôle d\'Accès (FIFA Ready)',
  description: 'Application de billetterie informatisée, paiement cashless et contrôle d\'accès de classe internationale pour le Grand Stade de Casablanca.',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="fr" class="dark">
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="anonymous" />
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet" />
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" />
      </head>
      <body class="bg-slate-950 text-slate-100 font-sans antialiased selection:bg-indigo-500 selection:text-white min-h-screen flex flex-col">
        {children}
      </body>
    </html>
  );
}
