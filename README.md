# Газар зүйн нэрийн файл хөрвүүлэгч

Excel файлаас Word (.docx) баримт бичиг үүсгэх вэб аппликейшн.

Хоёр хувилбартай:
- **Next.js** — шинэ хувилбар, Vercel-д deploy хийнэ
- **Django (Python)** — хуучин хувилбар, локал орчинд ажиллана

---

## Next.js хувилбар

### Суулгах

```bash
npm install
```

### Локал дээр ажиллуулах

```bash
npm run dev
```

Браузерт нээх: [http://localhost:3000](http://localhost:3000)

### Production build

```bash
npm run build
npm run start
```

### Vercel-д deploy хийх

```bash
git add -A
git commit -m "deploy"
git push
```

Vercel дээр repository холбосон бол push хийхэд автоматаар deploy болно.

### Алдаа гарвал

```bash
rm -rf .next && npm run dev
```

---

## Django (Python) хувилбар

### Virtual environment үүсгэх

```bash
# macOS / Linux
python3 -m venv venv
source venv/bin/activate

# Windows
python -m venv venv
venv\Scripts\activate
```

### Хамаарлуудыг суулгах

```bash
pip install -r requirements.txt
```

### Серверийг ажиллуулах

```bash
python manage.py runserver
```

Браузерт нээх: [http://127.0.0.1:8000](http://127.0.0.1:8000)

---

## Төслийн бүтэц

```
converter/
├── package.json             # Next.js тохиргоо
├── next.config.mjs
├── vercel.json
├── app/
│   ├── layout.js            # Navigation + layout
│   ├── page.js              # Хувийн хэрэг хуудас
│   ├── name-request/
│   │   └── page.js          # Хүсэлтийн маягт хуудас
│   └── api/
│       ├── huviin-hereg/
│       │   └── route.js     # Excel → Хувийн хэрэг API
│       └── name-request/
│           └── route.js     # Excel → Хүсэлтийн маягт API
├── lib/
│   ├── huviin-hereg.js      # Word doc үүсгэгч
│   └── name-request-doc.js  # Word doc үүсгэгч
├── manage.py                # Django (хуучин хувилбар)
├── requirements.txt
└── fileconverter/           # Django app (хуучин хувилбар)
```

---

## Боловсруулалтын урсгал (Next.js)

```
Excel файл (.xlsx)
      ↓
  Browser → POST /api/...
      ↓
  Next.js API route (сервер)
      ↓
xlsx → өгөгдөл унших
      ↓
docx → Word үүсгэх (Buffer)
      ↓
  .docx татаж авах
```

---

## Excel файлын баганын дараалал

| Багана | Индекс | Утга |
|--------|--------|------|
| B | 1 | Дахин давтагдашгүй дугаар |
| C | 2 | Нэрийн зургийн индекс |
| D | 3 | Газар зүйн нэр |
| E | 4 | Төрөл |
| F | 5 | Дэвсгэр нэр / Ангилал |
| J | 9 | Байр зүйн зургийн нэрлэвэр |
| K | 10 | 1:100 000 зурагт |
| N | 13 | Гарал үүсэл |
| O | 14 | Өргөрөг 1 |
| P | 15 | Уртраг 1 |
| Q | 16 | Өргөрөг 2 |
| R | 17 | Уртраг 2 |
| S | 18 | Аймаг / сум / баг |
| T | 19 | Байрлал тайлбар |
| U | 20 | Актын дугаар |
