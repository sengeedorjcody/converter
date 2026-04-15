import './globals.css';

export const metadata = {
  title: 'Газар зүйн нэрийн хөрвүүлэгч',
  description: 'Excel файлаас Word баримт үүсгэх',
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
      <a href="/" className="brand">ГЗН Хөрвүүлэгч</a>
      <a href="/">Хувийн хэрэг</a>
      <a href="/name-request">Хүсэлтийн маягт</a>
    </nav>
  );
}
