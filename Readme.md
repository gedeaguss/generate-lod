# 📄 Generate LoD (Letter of Discharge) Asuransi

Aplikasi web untuk mengotomatisasi pembuatan dokumen **Letter of Discharge (LoD)** klaim asuransi secara massal. Cukup siapkan template Word dan data Excel, aplikasi akan menghasilkan seluruh dokumen secara otomatis dalam hitungan detik.

---

## ✨ Fitur Utama

- Upload template `.docx` dan data `.xlsx` langsung dari browser
- Validasi otomatis kelengkapan kolom dan kompatibilitas placeholder
- Format nilai klaim otomatis sesuai kaidah nominal Rupiah (contoh: `1234567` → `Rp 1.234.567`)
- Generate dokumen massal sekaligus dengan progress real-time
- Download semua hasil dalam satu file `.zip`
- Panduan inline di setiap tahap agar mudah digunakan siapapun

---

## 📋 Persyaratan File

### File Excel (`.xlsx`)
Baris pertama **harus** berisi nama kolom. Kolom berikut wajib ada:

| Nama Kolom | Keterangan |
|---|---|
| `Nomor Polis` | Digunakan sebagai nama file output |
| `Nilai Klaim` | Diformat otomatis sebagai Rupiah |
| `Terbilang` | Nilai klaim dalam bentuk teks |
| `Akibat` | Penyebab/akibat kejadian |
| `Tanggal Kejadian` | Tanggal terjadinya klaim |
| `Lokasi Kejadian` | Lokasi terjadinya kejadian |

Kolom tambahan di luar daftar di atas tetap bisa digunakan selama placeholder-nya ada di template Word.

### File Template Word (`.docx`)
Gunakan placeholder `{{Nama Kolom}}` di bagian yang ingin diisi otomatis. Nama di dalam kurung kurawal harus **persis sama** dengan nama kolom di Excel, termasuk huruf besar/kecilnya.

**Contoh:**
```
Nomor Polis      : {{Nomor Polis}}
Tanggal Kejadian : {{Tanggal Kejadian}}
Lokasi Kejadian  : {{Lokasi Kejadian}}
Nilai Klaim      : {{Nilai Klaim}}
```

> **Tips:** Jika placeholder tidak terbaca, ketik ulang dengan cara copy-paste dari Notepad ke Word untuk menghindari masalah formatting tersembunyi.

---

## 🚀 Cara Menjalankan Aplikasi

### Akses via Web
Aplikasi dapat diakses langsung melalui Streamlit Cloud atau [Google Colaboratory](https://colab.research.google.com/drive/1BGmzq6j0xGnsubjcJAQOQU7d4q8RZPHt?usp=sharing)

### Menjalankan Secara Lokal
Jika ingin menjalankan di komputer sendiri:

**1. Clone repository ini**
```bash
git clone https://github.com/username/nama-repo.git
cd nama-repo
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Jalankan aplikasi**
```bash
streamlit run app.py
```

---

## 📁 Struktur Project

```
├── app.py                  # File utama aplikasi Streamlit
├── requirements.txt        # Daftar library yang dibutuhkan
├── README.md               # Dokumentasi project
└── .streamlit/
    └── config.toml         # Konfigurasi tampilan aplikasi
```

---

## 🔧 Teknologi yang Digunakan

| Library | Kegunaan |
|---|---|
| `streamlit` | Framework aplikasi web |
| `python-docx` | Membaca dan mengisi template Word |
| `openpyxl` | Membaca data dari file Excel |
| `tempfile` | Manajemen file sementara di server |
| `re` | Parsing dan deteksi placeholder |

---

## ⚠️ Catatan Penting

- Output aplikasi berupa file **`.docx`** — untuk mengubah ke PDF, buka file di Microsoft Word lalu gunakan **Save As → PDF**
- Pastikan tidak ada baris kosong di antara data pada file Excel
- Setiap sesi generate menggunakan template baru — sesuaikan tanggal di area tanda tangan template sebelum upload
- File yang diupload hanya tersimpan sementara selama sesi berlangsung dan otomatis terhapus setelahnya

---
