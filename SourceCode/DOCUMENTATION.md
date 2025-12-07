# Certificate Tool - Documentation

## Files Overview

### Python Scripts

1. **certificate_tool_gui.py** (RECOMMENDED)

   - GUI Interface menggunakan Tkinter
   - User-friendly dengan tombol Browse
   - Tab untuk setiap fitur (Rename, Split, Merge)
   - Real-time logging

   Jalankan:

   ```powershell
   python certificate_tool_gui.py
   ```

2. **certificate_tool.py**

   - CLI Interface (Command Line)
   - Untuk pengguna yang prefer terminal
   - Fitur sama dengan GUI

   Jalankan:

   ```powershell
   python certificate_tool.py
   ```

3. **tempCodeRunnerFile.py**

   - Entry point untuk menjalankan GUI
   - Bisa dijalankan langsung dari IDE

   Jalankan:

   ```powershell
   python tempCodeRunnerFile.py
   ```

### Executable

- **dist/Certificate Tool.exe**
  - Compiled executable (.EXE)
  - Tidak butuh Python terinstall
  - Double-click untuk menjalankan
  - Recommended untuk end-users

### Specification Files

- **Certificate Tool.spec**
  - PyInstaller specification file
  - Gunakan untuk rebuild exe:
  ```powershell
  pyinstaller "Certificate Tool.spec"
  ```

---

## Quick Start

### For GUI Users (Easiest)

```powershell
# Double-click Certificate Tool.exe
# OR
python certificate_tool_gui.py
```

### For CLI Users

```powershell
python certificate_tool.py
```

### To Install Dependencies

```powershell
pip install -r ../requirements.txt
```

---

## File Structure Example

```
c:\sertifikat eea\
├── rsrc/                          # Source PDF files
│   ├── sertif-1186.pdf
│   ├── sertif-1187.pdf
│   └── ... (1186-1585)
│
├── extracted/                     # Output folder (for rename)
│   ├── Budi Santoso_Peserta EEA 2025.pdf
│   ├── Siti Nurhaliza_Peserta EEA 2025.pdf
│   └── ...
│
├── split_output/                  # Output folder (for split)
│   ├── peserta-1.pdf
│   ├── peserta-2.pdf
│   └── ...
│
├── acara.xlsx                     # Data file with names
├── humas.xlsx                     # Data file with names
├── danus.xlsx                     # Data file with names
│
└── SourceCode/
    ├── certificate_tool.py        # CLI application
    ├── certificate_tool_gui.py    # GUI application
    └── dist/
        └── Certificate Tool.exe   # Compiled executable
```

---

## Usage Examples

### Example 1: Basic Rename

```
Input:
- Source: rsrc/sertif-1186.pdf, sertif-1187.pdf, ...
- Data: acara.xlsx (columns: nama, divisi)
- Format: {nama}_Peserta {divisi}

Output:
- extracted/Budi Santoso_Peserta Acara.pdf
- extracted/Siti Nurhaliza_Peserta Danus.pdf
```

### Example 2: Split & Merge

```
Input: sertifikat peserta seminar.pdf (2000 pages)

Split:
- split_output/peserta-1.pdf
- split_output/peserta-2.pdf
- ...
- split_output/peserta-2000.pdf

Merge back:
- merged_final.pdf
```

---

## Troubleshooting

### "ModuleNotFoundError: No module named 'pandas'"

```powershell
pip install pandas openpyxl PyPDF2
```

### ".exe not working"

1. Make sure you have Windows 7 or newer
2. Try running as Administrator
3. Check antivirus settings
4. Rebuild exe: `pyinstaller "Certificate Tool.spec"`

### "File not found" error

- Check folder names are correct
- Ensure PDF files exist in source folder
- Verify file naming pattern matches format

### "Excel column not found" error

- Check placeholder names match Excel columns exactly
- Placeholder is case-sensitive: {nama} != {Nama}
- Make sure Excel file has the required columns

---

## Features

✓ Rename multiple PDFs using Excel data
✓ Auto-handle duplicates
✓ Skip empty rows
✓ Split PDF pages into separate files
✓ Merge multiple PDFs
✓ Real-time progress logging
✓ Export summary reports

---

## Version History

- **v1.0** (Dec 7, 2025): Initial release
  - GUI with Tkinter
  - CLI support
  - All features working
  - Emoji removed for cleaner interface

---

## For Developers

### Modify and Rebuild EXE

```powershell
# Edit certificate_tool_gui.py
# Then rebuild:
python -m PyInstaller --onefile --windowed --name "Certificate Tool" certificate_tool_gui.py
```

### Clean Build Files

```powershell
Remove-Item -Recurse build/
Remove-Item -Recurse dist/
Remove-Item "Certificate Tool.spec"
```

---

**Questions?** Check the main README.md in parent directory.
