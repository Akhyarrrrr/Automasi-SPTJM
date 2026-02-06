# 📄 SPTJM Otomatis — Excel → PDF → Email (Opsional)

Aplikasi Streamlit untuk membuat **SPTJM otomatis** dari file Excel menjadi **PDF per orang**, digabung dalam **ZIP**, dan **opsional dikirim via email SMTP**.

Alur aman:
> Upload Excel → Generate PDF → Cek sample → (Opsional) Kirim Email  
Tidak ada auto-email sebelum kamu konfirmasi.

---

# ✅ Fitur Utama

- Generate SPTJM PDF otomatis per NIP
- Data Excel multi proposal (NoProp1, NoProp2, dst)
- Template SPTJM + lampiran dibuat otomatis
- Konversi DOCX → PDF via LibreOffice
- Download ZIP semua PDF
- Sample preview download
- Report generate & report email (CSV)
- Email massal dengan:
  - Delay anti-limit
  - Dry-run (simulasi)
  - Mapping email via file terpisah

---

# 🧰 Syarat Sistem

## 1️⃣ Python

Disarankan:

```
Python 3.10+
```

Cek versi:

```bash
python --version
```

---

## 2️⃣ LibreOffice (WAJIB)

Digunakan untuk convert DOCX → PDF.

Install:
```
LibreOffice Desktop
```

Default path Windows:

```
C:\Program Files\LibreOffice\program\soffice.exe
```

Bisa diisi manual di sidebar aplikasi jika berbeda.

Test:

```bash
soffice --version
```

---

## 3️⃣ Dependency Python

Install:

```bash
pip install -r requirements.txt
```

Isi `requirements.txt`:

```txt
streamlit==1.37.1
pandas==2.2.2
openpyxl==3.1.5
python-docx==1.1.2
python-slugify==8.0.4
python-dotenv==1.0.1
```

---

# 📁 Struktur Project

```
project/
├── app.py
├── generator.py
├── excel_utils.py
├── emailer.py
├── requirements.txt
├── .env
```

---

# 📊 Format Excel Wajib

## Kolom Identitas

| Kolom | Wajib |
|--------|---------|
NIP | ✅
Nama | ✅
Fakultas | ✅
Norek | ✅
Nama Bank | opsional
Email | opsional (bisa mapping)

---

## Kolom Proposal (Pola Wajib)

Minimal:

```
NoProp1
Judul1
Skema1
Jumlah_dana1
```

Boleh banyak:

```
NoProp1 … NoPropN
Judul1 … JudulN
Skema1 … SkemaN
Jumlah_dana1 … Jumlah_danaN
```

Baris tanpa NoProp akan dilewati otomatis.

---

# 📧 Konfigurasi Email (.env) — Opsional

Buat file `.env`:

```
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=email@domain.com
SMTP_PASS=password_atau_app_password
SMTP_FROM_NAME=SPTJM System

SOFFICE_PATH=C:\Program Files\LibreOffice\program\soffice.exe
```

---

## ⚠️ Gmail Khusus

Gunakan **App Password**, bukan password biasa.

Langkah:

```
Google Account → Security → App Password
```

---

# ▶️ Menjalankan Aplikasi

Dari folder project:

```bash
streamlit run app.py
```

Browser akan terbuka otomatis.

---

# 🧭 Cara Pakai di UI

## Step 1 — Upload Excel

- Upload file .xlsx
- Pilih sheet
- Klik **Load & Preview**

---

## Step 2 — Generate PDF

- Tentukan jumlah sample
- Tentukan jumlah generate
- Klik:

```
Generate ZIP PDF
```

Output:
- ZIP semua PDF
- Sample PDF
- Report CSV

---

## Step 3 — Kirim Email (Opsional)

- Centang kirim email
- Isi subject & body template
- Bisa pakai variabel:

```
{nama}
{nip}
```

Contoh subject:

```
SPTJM - {nama} ({nip})
```

- Centang konfirmasi
- Klik kirim

---

# 🧪 Dry Run Mode

Default aktif.

Artinya:

```
Email tidak benar-benar dikirim
Hanya simulasi + report
```

Disarankan test dulu sebelum kirim real.

---

# 📧 Mapping Email Terpisah

Jika Excel utama tidak punya kolom Email:

Upload file mapping dengan kolom:

```
NIP
Email
```

Sistem akan auto-merge.

---

# 📦 Output Sistem

## Generate

- ZIP PDF
- Sample PDF
- Generate report CSV

## Email

- Email report CSV
- Status: OK / FAIL / SKIP / DRY-RUN

---

# ❗ Error Umum

## LibreOffice tidak ditemukan

Isi manual di sidebar:

```
C:\Program Files\LibreOffice\program\soffice.exe
```

---

## PDF gagal generate

Biasanya:
- LibreOffice belum install
- Path salah
- Timeout convert

---

## Email gagal

Cek:

```
.env SMTP config
App password
Firewall kantor
```

---

# 🔒 Keamanan

- Tidak auto-kirim email
- Ada sample check
- Ada konfirmasi kirim
- Ada dry-run mode
- SMTP pakai env file
