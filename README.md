# Certificate Management Tool

Aplikasi untuk mengelola file sertifikat PDF dengan fitur rename, split, dan merge. Tersedia dalam bentuk GUI (interface grafis) dan CLI (command line).

---

## Daftar Isi

- [Fitur Utama](#fitur-utama)
- [Persyaratan Sistem](#persyaratan-sistem)
- [Library yang Diperlukan](#library-yang-diperlukan)
- [Instalasi](#instalasi)
- [Cara Penggunaan](#cara-penggunaan)
  - [Menggunakan GUI (Recommended)](#menggunakan-gui)
  - [Menggunakan CLI](#menggunakan-cli)
- [Contoh Kasus Nyata](#contoh-kasus-nyata)
- [Troubleshooting](#troubleshooting)

---

## Fitur Utama

### 1. **Rename Certificate Files**

Mengubah nama file PDF berdasarkan data Excel dengan format custom

- Rename otomatis menggunakan placeholder dari file Excel
- Penanganan duplikat otomatis
- Skip file kosong

### 2. **Split PDF**

Memisahkan file PDF multi-halaman menjadi file individual per halaman

- Mengatur format nama file output
- Mengatur nomor awal
- Progress real-time

### 3. **Merge PDF**

Menggabungkan multiple PDF menjadi satu file

- Merge dari folder
- Merge file spesifik
- Urutan otomatis

---

## Persyaratan Sistem

- **OS**: Windows 7 atau lebih baru
- **Python**: 3.8 atau lebih baru (jika menjalankan dari source code)
- **RAM**: Minimum 2GB
- **Disk Space**: Minimal 500MB untuk aplikasi + file temporer

---

## Library yang Diperlukan

### Jika Menggunakan Python Source Code:

```txt
pandas>=1.3.0        # Membaca file Excel
openpyxl>=3.6.0      # Support untuk file .xlsx
PyPDF2>=2.0.0        # Operasi manipulasi PDF
```

### Instalasi Library:

```powershell
pip install pandas openpyxl PyPDF2
```

**Untuk GUI tambahan:**

```powershell
pip install tkinter  # Biasanya sudah include di Python
```

---

## Instalasi

### Option 1: Menggunakan File .EXE (Recommended - Paling Mudah)

1. Download atau akses file `Certificate Tool.exe` di folder:

   ```
   C:\sertifikat eea\SourceCode\dist\Certificate Tool.exe
   ```

2. Double-click untuk menjalankan aplikasi
   - **Tidak perlu instalasi Python atau library**
   - **Langsung bisa digunakan**

### Option 2: Menjalankan dari Python Source Code

**Prasyarat**: Python 3.8+ sudah terinstall

1. Buka Command Prompt/PowerShell

2. Navigate ke folder project:

   ```powershell
   cd "c:\sertifikat eea\SourceCode"
   ```

3. Install library yang diperlukan:

   ```powershell
   pip install pandas openpyxl PyPDF2
   ```

4. Jalankan GUI:

   ```powershell
   python certificate_tool_gui.py
   ```

   Atau jalankan CLI:

   ```powershell
   python certificate_tool.py
   ```

---

## Cara Penggunaan

### Menggunakan GUI (Graphical User Interface)

#### **Tab 1: Rename Files**

1. **Folder Utama**: Pilih folder tempat data Anda berada

   - Default: `C:\sertifikat eea`

2. **Folder Sumber**: Folder yang berisi PDF original

   - Contoh: `rsrc` (berisi file `sertif-1186.pdf`, `sertif-1187.pdf`, dst)

3. **Folder Tujuan**: Folder output untuk hasil rename

   - Contoh: `extracted` (akan dibuat otomatis jika belum ada)

4. **File Excel**: Pilih file Excel yang berisi data nama

   - Format: Setiap baris = 1 sertifikat
   - Kolom yang diperlukan: sesuai placeholder di format tujuan

5. **Format File Sumber**: Template nama file original

   - Default: `sertifikat-{nomor}`
   - Akan dicocokkan dengan file di folder sumber

6. **Nomor Awal**: Nomor mulai

   - Default: `1`

7. **Format File Tujuan**: Template nama file hasil

   - Contoh: `{nama}_Peserta EEA 2025`
   - Placeholder: `{nama}`, `{email}`, `{divisi}`, dst (sesuai kolom Excel)

8. Klik **"Mulai Rename"** dan lihat progress di Log Output

#### **Tab 2: Split PDF**

1. **File PDF**: Pilih file PDF yang akan di-split

   - Contoh: `sertifikat peserta seminar.pdf`

2. **Folder Output**: Folder tempat hasil split disimpan

   - Default: `split_output`

3. **Format Nama File**: Template nama file hasil

   - Default: `sertifikat-{nomor}`
   - Output: `sertifikat-1.pdf`, `sertifikat-2.pdf`, dst

4. **Nomor Awal**: Nomor halaman pertama

   - Default: `1`

5. Klik **"Mulai Split"**

#### **Tab 3: Merge PDF**

1. **Folder PDF**: Pilih folder yang berisi file PDF

   - Contoh: `split_output`

2. **Nama File Output**: Nama file hasil gabungan

   - Contoh: `merged.pdf`

3. Klik **"Mulai Merge"**

---

### Menggunakan CLI (Command Line Interface)

Jalankan:

```powershell
python certificate_tool.py
```

Menu:

```
==============================================================
CERTIFICATE MANAGEMENT TOOL
==============================================================
1. Rename Certificate Files
2. Split PDF (Page per File)
3. Merge PDF Files
4. Exit
==============================================================

Pilih menu (1/2/3/4):
```

Ikuti instruksi di layar untuk setiap fitur.

---

## Contoh Kasus Nyata

### Kasus 1: Rename Sertifikat Peserta Seminar

**Data yang tersedia:**

- File PDF: `rsrc/sertif-1186.pdf`, `sertif-1187.pdf`, dst
- File Excel: `acara.xlsx` dengan kolom: `nama`, `divisi`, `email`

**Langkah:**

1. Buka `Certificate Tool.exe`
2. Di Tab "Rename Files":

   - Folder Utama: `C:\sertifikat eea`
   - Folder Sumber: `rsrc`
   - Folder Tujuan: `extracted`
   - File Excel: `acara.xlsx`
   - Format File Sumber: `sertif-{nomor}`
   - Nomor Awal: `1186`
   - Format File Tujuan: `{nama}_Peserta {divisi}`

3. Klik "Mulai Rename"

**Hasil:**

```
Budi Santoso_Peserta Acara.pdf
Siti Nurhaliza_Peserta Danus.pdf
Ahmad Wijaya_Peserta Humas.pdf
(dst sesuai data Excel)
```

**Output folder:** `C:\sertifikat eea\extracted\`

---

### Kasus 2: Split PDF Sertifikat Multi-Halaman

**Data yang tersedia:**

- File PDF: `sertifikat peserta seminar.pdf` (2000+ halaman)

**Langkah:**

1. Buka `Certificate Tool.exe`
2. Di Tab "Split PDF":

   - File PDF: `sertifikat peserta seminar.pdf`
   - Folder Output: `split_output`
   - Format Nama File: `peserta-{nomor}`
   - Nomor Awal: `1`

3. Klik "Mulai Split"

**Hasil:**

```
split_output/peserta-1.pdf
split_output/peserta-2.pdf
split_output/peserta-3.pdf
...
split_output/peserta-2000.pdf
```

---

### Kasus 3: Merge PDF Hasil Split

**Data yang tersedia:**

- Folder hasil split: `split_output/` (berisi 2000 file PDF)

**Langkah:**

1. Buka `Certificate Tool.exe`
2. Di Tab "Merge PDF":

   - Folder PDF: `split_output`
   - Nama File Output: `merged_final.pdf`

3. Klik "Mulai Merge"

**Hasil:**

```
C:\sertifikat eea\merged_final.pdf
```

---

## Format File Excel yang Benar

### Contoh 1: Rename dengan Nama & Divisi

**File: `acara.xlsx`**

| nama           | divisi | nomor |
| -------------- | ------ | ----- |
| Budi Santoso   | Acara  | 1186  |
| Siti Nurhaliza | Danus  | 1187  |
| Ahmad Wijaya   | Humas  | 1188  |

**Format Output:** `{nama}_Peserta {divisi}`

**Hasil:**

```
Budi Santoso_Peserta Acara.pdf
Siti Nurhaliza_Peserta Danus.pdf
Ahmad Wijaya_Peserta Humas.pdf
```

---

### Contoh 2: Rename dengan Multiple Kolom

**File: `data sertifikat umum 206-1185.xlsx`**

| nama         | keterangan       | nomor |
| ------------ | ---------------- | ----- |
| Ari Prasetyo | Peserta EEA 2025 | 206   |
| Dewi Lestari | Peserta EEA 2025 | 207   |

**Format Output:** `{nama}_{keterangan}`

**Hasil:**

```
Ari Prasetyo_Peserta EEA 2025.pdf
Dewi Lestari_Peserta EEA 2025.pdf
```

---

## Log Output & Summary

Setiap operasi akan menampilkan summary:

```
==================================================
SUMMARY:
Berhasil: 821
Tidak ditemukan: 5
Duplikat: 2
Dilewati: 10
Output: C:\sertifikat eea\extracted
==================================================

DETAIL TIDAK DITEMUKAN:
 • Baris 15: sertif-1200.pdf
 • Baris 42: sertif-1227.pdf

DETAIL DUPLIKAT:
 • Baris 8: Ari_Peserta -> Ari_Peserta (2)

DETAIL DILEWATI:
 • Baris 25: Kolom kosong -> nama
```

---

## Troubleshooting

### Error: "File tidak ada"

**Penyebab**: File PDF tidak ditemukan di folder sumber
**Solusi**:

1. Pastikan folder sumber benar
2. Pastikan format nama file sesuai (nama + nomor)
3. Periksa nomor awal dan jumlah file

### Error: "Kolom tidak ditemukan"

**Penyebab**: Placeholder di format tujuan tidak ada di file Excel
**Solusi**:

1. Periksa nama kolom di Excel (case-sensitive)
2. Pastikan kolom yang digunakan ada di file Excel
3. Contoh: jika format `{nama}_{email}`, pastikan Excel punya kolom "nama" dan "email"

### Error: "Module tidak ditemukan"

**Penyebab**: Library Python belum terinstall
**Solusi**:

```powershell
pip install pandas openpyxl PyPDF2
```

### Aplikasi Hang/Freeze

**Penyebab**: File PDF terlalu besar atau folder dengan banyak file
**Solusi**:

1. Tunggu sampai proses selesai (lihat di taskbar)
2. Untuk file besar, gunakan split terlebih dahulu
3. Restart aplikasi jika tidak responsif

### File Duplikat

**Penyebab**: Nama file hasil sudah ada
**Solusi**:

- Aplikasi otomatis menambah nomor: `nama (2).pdf`, `nama (3).pdf`
- Tidak ada file yang tertimpa

---

## Tips & Tricks

### 1. Test dengan Beberapa File Dulu

Sebelum proses ribuan file, test dengan 5-10 file untuk memastikan format benar.

### 2. Backup Data Original

Selalu backup folder source sebelum proses rename.

### 3. Nama Kolom Excel Harus Cocok

Format tujuan: `{nama}_{divisi}`
Excel harus punya kolom dengan nama: `nama` dan `divisi` (lowercase)

### 4. Gunakan Split untuk File Besar

File PDF 2000+ halaman lebih cepat dengan split dulu daripada rename langsung.

### 5. Log Output Bisa Dicopy

Hasil summary bisa dicopy dari log area untuk dokumentasi atau laporan.

---

## Support & Kontak

Untuk pertanyaan atau masalah, silakan cek:

1. Folder `rsrc/` - Contoh file PDF
2. File `*.xlsx` - Contoh struktur data Excel
3. File `extracted/` - Contoh hasil output

---

**Versi**: 1.0  
**Last Updated**: December 7, 2025  
**Status**: Production Ready
