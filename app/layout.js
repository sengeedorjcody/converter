import './globals.css';

export const metadata = {
  title: 'Инженер Геодези ХХК',
  description: 'Excel файлаас Word баримт үүсгэх',
  icons: { icon: '/logo.png' },
};

export default function RootLayout({ children }) {
  return (
    <html lang="mn">
      <body>
        <NavBar />
        {children}
      </body>
    </html>
  );
}

function NavBar() {
  return (
    <nav>
      <a href="/" className="brand">
        <img src="/logo.png" alt="Инженер Геодези ХХК" className="nav-logo" />
        <span>Инженер Геодези ХХК</span>
      </a>
      <a href="/">Хувийн хэрэг</a>
      <a href="/name-request">Хүсэлтийн маягт</a>
    </nav>
  );
}
