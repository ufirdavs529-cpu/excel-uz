# Microsoft Excel 365 Web Loyihasi (Full-Stack Versiya)

Ushbu loyiha haqiqiy **Microsoft Office Excel** ko'rinishi va funksionalligini brauzerda to'liq taqdim etuvchi professional veb-dasturga aylantirildi.

---

## 🚀 Qilingan asosiy o'zgarishlar va yangiliklar

### 1. `nav` qismidagi "File" tugmasi va `excel.html` sahifasi
- **Oldingi holat:** `excel.html` da faqat statik `image.png` rasmi turgan edi.
- **Yangi holat:** `image.png` dagi haqiqiy Excel menyusi to'liq **HTML5, CSS3 va JavaScript** yordamida interaktiv qilib noldan yaratildi:
  - **Chap menyu (Sidebar):** Orqaga qaytish tugmasi (`←`), *Главная, Создать, Открыть, Сохранить, Сохранить как (CSV/XLSX), Печать, Экспорт*.
  - **Yuqori sarlavha:** Foydalanuvchi vaqti bo'yicha dinamik salomlashish (*"Добрый день / Добрый вечер, Dovud!"*).
  - **Shablonlar (Создать):** *"Пустая книга"*, *"Добро пожаловать в Excel"*, *"Учебник по формулам"*, *"Личный бюджет на месяц"* va boshqalar. Har bir shablonni bosganda u avtomatik `index.html` ga yuklanadi!
  - **Oxirgi fayllar (Последние):** `2-maktab_baholari`, `sirdaryo_11-sinf_lotin`, `shaxsiy_moliyaviy_yordamchi`. Bosilganda tegishli jadvallar jurnali ochiladi.

---

### 2. Takroriy kodlarni optimallashtirish (`index.html`)
- **Oldingi holat:** `index.html` da 37 ta takroriy `<div class="in1">` va har birida 15 tadan qattiq yozilgan `<input>` mavjud edi (694 qator noqulay kod).
- **Yangi holat:** Barcha yacheykalar va ustunlar (`A-Z`, `1-60+`) `script.js` orqali dinamik, toza va tezkor generatsiya qilinadi.

---

### 3. Haqiqiy Microsoft Excel imkoniyatlari
- **Lenta paneli (Ribbon):**
  - **Буфер обмена:** Вставить, Вырезать, Копировать.
  - **Шрифт:** Shrift turi (Calibri, Aptos, Arial va h.k.), shrift o'lchami, **Ж** (Bold), *К* (Italic), <u>Ч</u> (Underline), fon rangi va matn rangini bo'yash (Color pickers).
  - **Выравнивание:** Chapga, o'rtaga, o'ngga tekislash, kataklarni birlashtirish.
  - **Редактирование:** `Σ Автосумма` (ustunni avtomatik yig'indisini hisoblaydi), katakni tozalash.
- **Formula paneli (Formula Bar):**
  - **Name Box (`A1`):** Joriy yacheyka manzilini ko'rsatadi. Istalgan manzilni (masalan `C10`) yozib Enter bossangiz, kursor o'sha joyga o'tadi.
  - **Formula qatori (`fx`):** Haqiqiy formulalarni qo'llab-quvvatlaydi:
    - `=SUM(A1:A5)` — yig'indi
    - `=AVERAGE(B1:B10)` — o'rtacha qiymat
    - `=COUNT(C1:C10)` — sonlar soni
    - `=MAX(...)` va `=MIN(...)`
    - Matematik amallar: `=A1+B1`, `=A1*10`, `=(B4/C4)*100`
- **Klaviatura bilan boshqarish:**
  - `Strelkalar (↑, ↓, ←, →)` — yacheykalar bo'ylab erkin harakatlanish.
  - `Enter` — pastki yacheykaga o'tish (`Shift + Enter` yuqoriga).
  - `Tab` — o'ngdagi yacheykaga o'tish (`Shift + Tab` chapga).
  - `Delete` / `Backspace` — yacheykani tozalash.
  - `Ctrl + S` — saqlash.
  - `Ctrl + B`, `Ctrl + I`, `Ctrl + U` — shrift stillarini o'zgartirish.
- **Pastki panel (Sheets & Status Bar):**
  - Varaqlar boshqaruvi (`Лист1`, yangi varaq qo'shish `+`).
  - Holat satri: tanlangan yacheyka bo'yicha **O'rtacha qiymat (СРЕДНЕЕ)**, **Soni (КОЛИЧЕСТВО)** va **Yig'indisi (СУММА)** avtomatik hisoblanadi.
  - Masshtab slayderi (Zoom: 50% - 150%).

---

### 4. Backend va Ma'lumotlar bazasi (PHP, Python, MySQL)

1. **`database.sql` (MySQL bazasi):**
   - `excel_db` ma'lumotlar bazasi.
   - `workbooks` va `cells` jadvallari.
   - `2-maktab_baholari`, `sirdaryo_11-sinf_lotin` kabi tayyor maktab va reyting jadvallari ma'lumotlari bilan to'ldirilgan.
2. **`api.php` (PHP REST API):**
   - Jadvallarni MySQL bazasiga saqlash (`POST ?action=save`).
   - Bazadan yuklab olish (`GET ?action=load&name=...`).
   - Barcha saqlangan kitoblar ro'yxatini olish (`GET ?action=list`).
3. **`server.py` (Python Web & API Server):**
   - Standart Python kutubxonalarida yozilgan, hech qanday qo'shimcha paketlarsiz ishlaydi.
   - SQLite va MySQL ulanish imkoniyati bilan veb-serverni ishga tushiradi.

---

## 💻 Loyihani ishga tushirish usullari

### 1-usul: Oddiy brauzerda (server talab qilinmaydi)
- Shunchaki `index.html` faylini ikki marta bosib brauzerda oching.
- Yuqoridagi **"Файл"** tugmasini bosib `excel.html` (Backstage) oynasiga o'tishingiz mumkin.
- Barcha kiritilgan ma'lumotlar brauzer xotirasiga (`LocalStorage`) avtomatik saqlanib turadi.

### 2-usul: Python orqali ishga tushirish
Agar kompyuteringizda Python o'rnatilgan bo'lsa, terminalda quyidagi buyruqni bering:
```bash
python server.py
```
Brauzerda quyidagi manzilni oching:
`http://localhost:8000`

### 3-usul: PHP va MySQL orqali (XAMPP, Laragon, OpenServer)
1. Ushbu papkani `htdocs` (yoki `www`) papkasiga joylashtiring.
2. phpMyAdmin orqali `database.sql` faylini import qiling.
3. Brauzerda loyihani ochib, tepadagi **"Сохранить в MySQL"** tugmasini bossangiz, barcha ma'lumotlar MySQL bazasiga yoziladi.
