# 📄 Газар зүйн нэрийн файл хөрвүүлэгч

Excel файлаас Word (.docx) баримт бичиг үүсгэх Django вэб аппликейшн.

---

## ⚙️ Суулгах заавар

### 1. Repository татах

```bash
git clone https://github.com/your-username/converter.git
cd converter
```

### 2. Virtual environment үүсгэх

**macOS / Linux:**

```bash
python3 -m venv venv
source venv/bin/activate
```

**Windows:**

```bash
python -m venv venv
venv\Scripts\activate
```

> ✅ Амжилттай идэвхжсэн бол терминал дээр `(venv)` гарч ирнэ.

### 3. Хамаарлуудыг суулгах

```bash
pip install -r requirements.txt
```

### 4. Django тохиргоо

```bash
python manage.py migrate
python manage.py collectstatic
```

### 5. Серверийг ажиллуулах

```bash
python manage.py runserver
```

Браузерт нээх: [http://127.0.0.1:8000](http://127.0.0.1:8000)

---

## 📦 requirements.txt

```
Django>=4.2
openpyxl>=3.1.0
python-docx>=1.1.0
```

---

## 🗂️ Төслийн бүтэц

```
converter/
├── manage.py
├── requirements.txt
├── README.md
├── converter/               # Django тохиргоо
│   ├── settings.py
│   ├── urls.py
│   └── wsgi.py
└── fileconverter/           # Үндсэн апп
    ├── views.py             # Файл хөрвүүлэх логик
    ├── forms.py             # Upload форм
    ├── urls.py
    └── templates/
        └── fileconverter/
            ├── upload.html      # Хувийн хэрэг үүсгэх хуудас
            └── name_request.html # Газар зүйн нэрийн өргөдөл
```

---

## 🔄 Боловсруулалтын урсгал

```
Excel файл (.xlsx)
      ↓
  Django view
      ↓
openpyxl → өгөгдөл унших
      ↓
python-docx → Word үүсгэх
      ↓
  .docx татаж авах
```

---

## 📋 Функцүүд

| URL              | Тайлбар                                |
| ---------------- | -------------------------------------- |
| `/upload/`       | Нэрийн жагсаалтийг хувийн хэрэг болгох |
| `/name-request/` | Газар зүйн нэр өгөх өргөдлийн маягт    |

---

## 🛠️ Хөгжүүлэлтийн тохиргоо

### Virtual environment дахин идэвхжүүлэх

```bash
# macOS / Linux
source venv/bin/activate

# Windows
venv\Scripts\activate
```

### Шинэ package суулгасны дараа хадгалах

```bash
pip freeze > requirements.txt
```

### Virtual environment идэвхгүй болгох

```bash
deactivate
```

---

## 🖼️ Зураг замын тохиргоо

`views.py` дотор `IMAGE_PATHS` болон `EXPORT_PATHS` хувьсагчдыг өөрийн серверийн замд тохируулна:

```python
IMAGE_PATHS = [
    'E:\\5 sum zurag\\photos\\Gurvantes',  # Windows
    # '/home/user/photos/Gurvantes',       # Linux/macOS
]
```

---

## ❗ Түгээмэл алдаа

| Алдаа                         | Шийдэл                                       |
| ----------------------------- | -------------------------------------------- |
| `ModuleNotFoundError: django` | `pip install -r requirements.txt` ажиллуулна |
| `venv` идэвхгүй               | `source venv/bin/activate` ажиллуулна        |
| Зураг олдохгүй                | `IMAGE_PATHS` замыг шалгана                  |
| Port ашиглагдаж байна         | `python manage.py runserver 8001`            |
