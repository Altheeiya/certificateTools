import os
import pandas as pd
import shutil
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import re

# Github: Altheeiya


def safe_filename(name):
    """Membersihkan nama file agar aman dan unik"""
    safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).rstrip()
    return safe_name[:200] if len(safe_name) > 200 else safe_name

def list_files(base_folder, extension):
    """Menampilkan daftar file dengan ekstensi tertentu"""
    if isinstance(extension, tuple):
        files = [f for f in os.listdir(base_folder) if f.endswith(extension)]
    else:
        files = [f for f in os.listdir(base_folder) if f.endswith(extension)]
    
    if files:
        print(f"\nFile yang tersedia:")
        for i, file in enumerate(files, 1):
            print(f"   {i}. {file}")
    return files

def select_file(base_folder, extension, prompt_text):
    """Memilih file berdasarkan nomor"""
    files = list_files(base_folder, extension)
    
    if not files:
        print(f"\nTidak ada file dengan ekstensi {extension} di folder")
        return None
    
    while True:
        choice = input(f"\n{prompt_text} (ketik nomor atau nama file): ").strip()
        
        # Cek apakah input adalah nomor
        if choice.isdigit():
            index = int(choice) - 1
            if 0 <= index < len(files):
                return files[index]
            else:
                print(f"Nomor tidak valid! Pilih antara 1-{len(files)}")
        else:
            # Jika bukan nomor, anggap sebagai nama file
            if choice in files:
                return choice
            else:
                print(f"File '{choice}' tidak ditemukan!")

def get_base_folder():
    """Mendapatkan folder utama dengan deteksi otomatis"""
    current_dir = os.getcwd()
    
    print(f"\nFolder saat ini: {current_dir}")
    choice = input("   Gunakan folder ini? (y/n, default: y): ").strip().lower()
    
    if choice == 'n':
        custom_path = input("   Masukkan path folder utama: ").strip()
        if os.path.exists(custom_path):
            return custom_path
        else:
            print(f"   Folder tidak ditemukan: {custom_path}")
            return None
    else:
        return current_dir

def _extract_placeholders(pattern: str):
    """Ambil placeholder dalam kurung kurawal {placeholder}"""
    return set(m.strip().lower() for m in re.findall(r"{([^{}]+)}", pattern))

def _format_with_row(pattern: str, row: pd.Series, col_map: dict):
    """Gantikan placeholder pada pattern dengan nilai dari row (disanitasi per bagian)"""
    def repl(match):
        key = match.group(1).strip().lower()
        if key in col_map:
            v = row[col_map[key]]
            if pd.isna(v):
                return ""
            return safe_filename(str(v).strip())
        # biarkan placeholder yang tidak dikenal apa adanya (seharusnya sudah divalidasi)
        return match.group(0)
    return re.sub(r"{([^{}]+)}", repl, pattern)


def _read_int(prompt: str, default: int) -> int:
    """Membaca input angka dengan fallback default agar interaksi lebih singkat."""
    raw = input(f"{prompt} (default: {default}): ").strip()
    if raw == "":
        return default
    try:
        return int(raw)
    except ValueError:
        print(f"Input bukan angka, menggunakan default {default}")
        return default


def rename_certificates():
    """Fitur 1: Rename certificate files dengan alur yang lebih ringkas."""
    duplicate_list = []
    not_found_list = []
    skipped_list = []
    success_list = []

    try:
        # 1) Lokasi & file excel
        base_folder = get_base_folder()
        if not base_folder:
            return

        src_folder = input("\nNama folder sumber (default: rsrc): ").strip() or "rsrc"
        dst_folder = input("Nama folder tujuan (default: extracted): ").strip() or "extracted"
        src_path = os.path.join(base_folder, src_folder)
        dst_path = os.path.join(base_folder, dst_folder)
        os.makedirs(dst_path, exist_ok=True)

        excel_file = select_file(base_folder, ('.xlsx', '.xls'), "Pilih file Excel")
        if not excel_file:
            return
        excel_path = os.path.join(base_folder, excel_file)

        print(f"\nFile Excel dipilih: {excel_file}")
        data = pd.read_excel(excel_path)
        col_map = {str(c).strip().lower(): c for c in data.columns}
        print(f"Kolom: {', '.join([str(c) for c in data.columns])}")
        print(f"Total baris: {len(data)}\n")

        if not os.path.exists(src_path):
            print(f"ERROR: Folder sumber tidak ditemukan: {src_path}")
            return

        # 2) Pola nama file
        file_pattern = input("Format nama file sumber (tanpa .pdf, default: sertifikat-{nomor}): ").strip() or "sertifikat-{nomor}"
        if not file_pattern.endswith('.pdf'):
            file_pattern += '.pdf'

        output_pattern = input("Format nama file tujuan (tanpa .pdf, contoh: {nama}_{keterangan} EEA 2025): ").strip() or "{nama}_EEA 2025"
        if not output_pattern.endswith('.pdf'):
            output_pattern += '.pdf'

        used_placeholders = _extract_placeholders(output_pattern)
        missing_cols = [p for p in used_placeholders if p not in col_map]
        if missing_cols:
            print(f"\nERROR: Placeholder tidak ditemukan di kolom Excel: {', '.join(missing_cols)}")
            return

        start_number = _read_int("Nomor awal sertifikat", 1)

        print(f"\n{'='*60}")
        print("MEMULAI PROSES RENAME")
        print(f"{'='*60}\n")

        used_names = {}
        processed = not_found = duplicates = skipped = 0

        for i, row in data.iterrows():
            nomor_sertifikat = start_number + i
            old_name = file_pattern.replace("{nomor}", str(nomor_sertifikat))
            old_path = os.path.join(src_path, old_name)

            missing_values = [
                p for p in used_placeholders
                if pd.isna(row[col_map[p]])
                or str(row[col_map[p]]).strip() == ""
                or str(row[col_map[p]]).strip().lower() == "nan"
            ]
            if missing_values:
                skipped += 1
                skipped_list.append(f"Baris {i+1}: Kolom kosong -> {', '.join(missing_values)}")
                print(f"Baris {i+1}: Kolom kosong ({', '.join(missing_values)}) - dilewati")
                continue

            base_new_name = _format_with_row(output_pattern, row, col_map)

            if base_new_name in used_names:
                used_names[base_new_name] += 1
                name_parts = base_new_name.rsplit('.', 1)
                if len(name_parts) == 2:
                    new_name = f"{name_parts[0]} ({used_names[base_new_name]}).{name_parts[1]}"
                else:
                    new_name = f"{base_new_name} ({used_names[base_new_name]})"
                duplicates += 1
                duplicate_list.append({
                    'baris': i+1,
                    'nama': base_new_name,
                    'file_asli': base_new_name,
                    'file_hasil': new_name
                })
                print(f"Duplikasi: {base_new_name} -> {new_name}")
            else:
                used_names[base_new_name] = 0
                new_name = base_new_name

            new_path = os.path.join(dst_path, new_name)

            if os.path.exists(old_path):
                shutil.copy2(old_path, new_path)
                processed += 1
                success_list.append(f"{old_name} -> {new_name}")
                print(f"[{processed}] {old_name} -> {new_name}")
            else:
                not_found += 1
                ctx = ", ".join([f"{p}={str(row[col_map[p]])}" for p in used_placeholders])
                not_found_list.append(f"Baris {i+1}: {old_name} ({ctx})")
                print(f"File tidak ada: {old_name}")

        print(f"\n{'='*60}")
        print("SUMMARY RENAME:")
        print(f"{'='*60}")
        print(f"Total data Excel        : {len(data)}")
        print(f"Berhasil diproses       : {processed}")
        print(f"File tidak ditemukan    : {not_found}")
        print(f"Nama duplikat           : {duplicates}")
        print(f"Data dilewati (kosong)  : {skipped}")
        print(f"Total files output      : {len(os.listdir(dst_path))}")
        print(f"Lokasi output           : {dst_path}")
        print(f"{'='*60}")

        if not_found_list or duplicate_list or skipped_list:
            print(f"\n{'='*60}")
            print("DETAIL LAPORAN:")
            print(f"{'='*60}")

            if not_found_list:
                print(f"\nFILE TIDAK DITEMUKAN ({len(not_found_list)}):")
                for item in not_found_list:
                    print(f"   • {item}")

            if duplicate_list:
                print(f"\nNAMA DUPLIKAT ({len(duplicate_list)}):")
                for item in duplicate_list:
                    print(f"   • Baris {item['baris']}: {item['nama']}")
                    print(f"     Original : {item['file_asli']}")
                    print(f"     Renamed  : {item['file_hasil']}")

            if skipped_list:
                print(f"\nDATA DILEWATI ({len(skipped_list)}):")
                for item in skipped_list:
                    print(f"   • {item}")

            print(f"{'='*60}")

    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()
def split_pdf():
    """Fitur 2: Split PDF menjadi page per file."""
    try:
        base_folder = get_base_folder()
        if not base_folder:
            return

        pdf_file = select_file(base_folder, '.pdf', "Pilih file PDF yang akan di-split")
        if not pdf_file:
            return

        pdf_path = os.path.join(base_folder, pdf_file)
        output_folder = input("\nNama folder output (default: split_output): ").strip() or "split_output"
        output_path = os.path.join(base_folder, output_folder)
        os.makedirs(output_path, exist_ok=True)

        file_pattern = input("Format nama file hasil (tanpa .pdf, default: sertifikat-{nomor}): ").strip() or "sertifikat-{nomor}"
        if not file_pattern.endswith('.pdf'):
            file_pattern += '.pdf'

        start_number = _read_int("Nomor awal", 1)

        print("\nMembaca file PDF...")
        reader = PdfReader(pdf_path)
        total_pages = len(reader.pages)
        print(f"Total halaman: {total_pages}\n")

        success_list = []
        failed_list = []

        print(f"{'='*60}")
        print("MEMULAI PROSES SPLIT...")
        print(f"{'='*60}\n")

        for i in range(total_pages):
            try:
                writer = PdfWriter()
                writer.add_page(reader.pages[i])
                output_filename = file_pattern.replace("{nomor}", str(start_number + i))
                output_filepath = os.path.join(output_path, output_filename)
                with open(output_filepath, 'wb') as output_file:
                    writer.write(output_file)
                success_list.append(output_filename)
                print(f"[{i+1}/{total_pages}] Halaman {i+1} -> {output_filename}")
            except Exception as err:
                failed_list.append(f"Halaman {i+1}: {str(err)}")
                print(f"[{i+1}/{total_pages}] Halaman {i+1} GAGAL: {str(err)}")

        print(f"\n{'='*60}")
        print("SUMMARY SPLIT PDF:")
        print(f"{'='*60}")
        print(f"Total halaman         : {total_pages}")
        print(f"Berhasil diproses     : {len(success_list)}")
        print(f"Gagal diproses        : {len(failed_list)}")
        print(f"Lokasi output         : {output_path}")
        print(f"{'='*60}")

        if failed_list:
            print(f"\n{'='*60}")
            print("DETAIL HALAMAN GAGAL:")
            print(f"{'='*60}")
            for item in failed_list:
                print(f"   • {item}")
            print(f"{'='*60}")

    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()


def merge_pdf():
    """Fitur 3: Merge multiple PDF files menjadi satu."""
    try:
        base_folder = get_base_folder()
        if not base_folder:
            return

        print("\nPilih metode merge:")
        print("   1. Merge dari folder (semua PDF dalam folder)")
        print("   2. Merge file spesifik (pilih file dengan nomor)")
        method = input("Pilih (1/2): ").strip()

        merger = PdfMerger()
        pdf_files = []
        failed_files = []

        if method == "1":
            src_folder = input("\nNama folder sumber PDF: ").strip()
            src_path = os.path.join(base_folder, src_folder)

            if not os.path.exists(src_path):
                print(f"ERROR: Folder tidak ditemukan: {src_path}")
                return

            pdf_files = sorted([f for f in os.listdir(src_path) if f.endswith('.pdf')])
            if not pdf_files:
                print(f"ERROR: Tidak ada file PDF di folder: {src_path}")
                return

            print(f"\nDitemukan {len(pdf_files)} file PDF")
            if len(pdf_files) > 10:
                print("   Preview file:")
                for i in range(5):
                    print(f"      {i+1}. {pdf_files[i]}")
                print(f"      ... ({len(pdf_files)-10} file lainnya) ...")
                for i in range(len(pdf_files)-5, len(pdf_files)):
                    print(f"      {i+1}. {pdf_files[i]}")
            else:
                for i, f in enumerate(pdf_files, 1):
                    print(f"      {i}. {f}")

            confirm = input("\nMerge semua file? (y/n): ").strip().lower()
            if confirm != 'y':
                print("Merge dibatalkan")
                return

            print(f"\n{'='*60}")
            print("MEMULAI PROSES MERGE...")
            print(f"{'='*60}\n")

            for i, pdf_file in enumerate(pdf_files, 1):
                try:
                    pdf_path = os.path.join(src_path, pdf_file)
                    merger.append(pdf_path)
                    print(f"[{i}/{len(pdf_files)}] Menambahkan: {pdf_file}")
                except Exception as err:
                    failed_files.append(f"{pdf_file}: {str(err)}")
                    print(f"[{i}/{len(pdf_files)}] GAGAL: {pdf_file} - {str(err)}")

        elif method == "2":
            available_files = list_files(base_folder, '.pdf')
            if not available_files:
                return

            pdf_files_paths = []
            print("\nMasukkan nomor file PDF (pisahkan dengan koma, contoh: 1,3,5)")
            print("   Atau ketik 'all' untuk semua file")
            print("   Atau ketik nomor satu per satu (ketik 'done' jika selesai)")

            choice = input("\nPilihan: ").strip()
            if choice.lower() == 'all':
                pdf_files_paths = [os.path.join(base_folder, f) for f in available_files]
                pdf_files = available_files
            elif ',' in choice:
                numbers = [n.strip() for n in choice.split(',')]
                for num in numbers:
                    if num.isdigit():
                        index = int(num) - 1
                        if 0 <= index < len(available_files):
                            pdf_files_paths.append(os.path.join(base_folder, available_files[index]))
                            pdf_files.append(available_files[index])
                            print(f"   Ditambahkan: {available_files[index]}")
                        else:
                            print(f"   Nomor {num} tidak valid")
            else:
                while True:
                    if choice.lower() == 'done':
                        break
                    if choice.isdigit():
                        index = int(choice) - 1
                        if 0 <= index < len(available_files):
                            pdf_path = os.path.join(base_folder, available_files[index])
                            if pdf_path not in pdf_files_paths:
                                pdf_files_paths.append(pdf_path)
                                pdf_files.append(available_files[index])
                                print(f"   Ditambahkan: {available_files[index]}")
                            else:
                                print(f"   File sudah ditambahkan sebelumnya")
                        else:
                            print(f"   Nomor tidak valid")
                    choice = input(f"   File #{len(pdf_files)+1} (nomor/done): ").strip()

            if not pdf_files_paths:
                print("\nTidak ada file untuk di-merge")
                return

            print(f"\n{'='*60}")
            print("MEMULAI PROSES MERGE...")
            print(f"{'='*60}\n")

            for i, pdf_path in enumerate(pdf_files_paths, 1):
                try:
                    merger.append(pdf_path)
                    print(f"[{i}/{len(pdf_files_paths)}] Menambahkan: {os.path.basename(pdf_path)}")
                except Exception as err:
                    failed_files.append(f"{os.path.basename(pdf_path)}: {str(err)}")
                    print(f"[{i}/{len(pdf_files_paths)}] GAGAL: {os.path.basename(pdf_path)} - {str(err)}")
        else:
            print("Pilihan tidak valid!")
            return

        output_file = input("\nNama file hasil merge (tanpa .pdf, contoh: merged): ").strip()
        if not output_file.endswith('.pdf'):
            output_file += '.pdf'
            print(f"   Nama lengkap: {output_file}")

        output_path = os.path.join(base_folder, output_file)

        print(f"\nMenyimpan hasil merge...")
        merger.write(output_path)
        merger.close()

        print(f"\n{'='*60}")
        print("SUMMARY MERGE PDF:")
        print(f"{'='*60}")
        print(f"Total file          : {len(pdf_files)}")
        print(f"Berhasil digabung   : {len(pdf_files) - len(failed_files)}")
        print(f"Gagal digabung      : {len(failed_files)}")
        print(f"File output         : {output_path}")
        print(f"Merge selesai!")
        print(f"{'='*60}")

        if failed_files:
            print(f"\n{'='*60}")
            print("DETAIL FILE GAGAL:")
            print(f"{'='*60}")
            for item in failed_files:
                print(f"   • {item}")
            print(f"{'='*60}")

    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()

def main():
    """Menu utama"""
    while True:
        print(f"\n{'='*60}")
        print("CERTIFICATE MANAGEMENT TOOL")
        print(f"{'='*60}")
        print("1. Rename Certificate Files")
        print("2. Split PDF (Page per File)")
        print("3. Merge PDF Files")
        print("4. Exit")
        print(f"{'='*60}")
        
        choice = input("\nPilih menu (1/2/3/4): ").strip()
        
        if choice == "1":
            rename_certificates()
        elif choice == "2":
            split_pdf()
        elif choice == "3":
            merge_pdf()
        elif choice == "4":
            print("\nTerima kasih telah menggunakan Certificate Management Tool!")
            break
        else:
            print("Pilihan tidak valid! Silakan pilih 1, 2, 3, atau 4.")
            continue
        
        print(f"\n{'-'*60}")
        again = input("Apakah ingin menjalankan proses lain? (y/n): ").strip().lower()
        if again != 'y':
            print("\nTerima kasih telah menggunakan Certificate Management Tool!")
            break

if __name__ == "__main__":
    main()