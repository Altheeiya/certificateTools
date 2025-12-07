import os
import pandas as pd
import shutil
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path

# Github: Altheeiya

class CertificateToolGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Certificate Management Tool")
        self.root.geometry("900x700")
        
        # Variabel
        self.base_folder = tk.StringVar(value=os.getcwd())
        self.src_folder = tk.StringVar(value="rsrc")
        self.dst_folder = tk.StringVar(value="extracted")
        
        self.setup_ui()
    
    def setup_ui(self):
        # Notebook (Tabs)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab 1: Rename
        tab_rename = ttk.Frame(notebook)
        notebook.add(tab_rename, text="Rename Files")
        self.setup_rename_tab(tab_rename)
        
        # Tab 2: Split
        tab_split = ttk.Frame(notebook)
        notebook.add(tab_split, text="Split PDF")
        self.setup_split_tab(tab_split)
        
        # Tab 3: Merge
        tab_merge = ttk.Frame(notebook)
        notebook.add(tab_merge, text="Merge PDF")
        self.setup_merge_tab(tab_merge)
        
        # Log area di bawah
        log_frame = ttk.LabelFrame(self.root, text="Log Output", padding=10)
        log_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(fill='both', expand=True)
    
    def setup_rename_tab(self, parent):
        frame = ttk.Frame(parent, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Base folder
        ttk.Label(frame, text="Folder Utama:").grid(row=0, column=0, sticky='w', pady=5)
        ttk.Entry(frame, textvariable=self.base_folder, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.browse_base_folder).grid(row=0, column=2)
        
        # Source folder
        ttk.Label(frame, text="Folder Sumber:").grid(row=1, column=0, sticky='w', pady=5)
        ttk.Entry(frame, textvariable=self.src_folder, width=50).grid(row=1, column=1, padx=5)
        
        # Destination folder
        ttk.Label(frame, text="Folder Tujuan:").grid(row=2, column=0, sticky='w', pady=5)
        ttk.Entry(frame, textvariable=self.dst_folder, width=50).grid(row=2, column=1, padx=5)
        
        # Excel file
        ttk.Label(frame, text="File Excel:").grid(row=3, column=0, sticky='w', pady=5)
        self.excel_file = tk.StringVar()
        ttk.Entry(frame, textvariable=self.excel_file, width=50).grid(row=3, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.browse_excel).grid(row=3, column=2)
        
        # File pattern source
        ttk.Label(frame, text="Format File Sumber:").grid(row=4, column=0, sticky='w', pady=5)
        self.file_pattern_src = tk.StringVar(value="sertifikat-{nomor}")
        ttk.Entry(frame, textvariable=self.file_pattern_src, width=50).grid(row=4, column=1, padx=5)
        ttk.Label(frame, text="Contoh: sertifikat-{nomor}", font=('', 8, 'italic')).grid(row=5, column=1, sticky='w')
        
        # Start number
        ttk.Label(frame, text="Nomor Awal:").grid(row=6, column=0, sticky='w', pady=5)
        self.start_num = tk.StringVar(value="1")
        ttk.Entry(frame, textvariable=self.start_num, width=50).grid(row=6, column=1, padx=5)
        
        # File pattern destination
        ttk.Label(frame, text="Format File Tujuan:").grid(row=7, column=0, sticky='w', pady=5)
        self.file_pattern_dst = tk.StringVar(value="{nama}_Peserta EEA 2025")
        ttk.Entry(frame, textvariable=self.file_pattern_dst, width=50).grid(row=7, column=1, padx=5)
        ttk.Label(frame, text="Contoh: {nama}_{keterangan} EEA 2025", font=('', 8, 'italic')).grid(row=8, column=1, sticky='w')
        
        # Button
        ttk.Button(frame, text="Mulai Rename", command=self.run_rename, style='Accent.TButton').grid(row=9, column=1, pady=20)
    
    def setup_split_tab(self, parent):
        frame = ttk.Frame(parent, padding=20)
        frame.pack(fill='both', expand=True)
        
        # PDF file
        ttk.Label(frame, text="File PDF:").grid(row=0, column=0, sticky='w', pady=5)
        self.split_pdf_file = tk.StringVar()
        ttk.Entry(frame, textvariable=self.split_pdf_file, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.browse_split_pdf).grid(row=0, column=2)
        
        # Output folder
        ttk.Label(frame, text="Folder Output:").grid(row=1, column=0, sticky='w', pady=5)
        self.split_output = tk.StringVar(value="split_output")
        ttk.Entry(frame, textvariable=self.split_output, width=50).grid(row=1, column=1, padx=5)
        
        # File pattern
        ttk.Label(frame, text="Format Nama File:").grid(row=2, column=0, sticky='w', pady=5)
        self.split_pattern = tk.StringVar(value="sertifikat-{nomor}")
        ttk.Entry(frame, textvariable=self.split_pattern, width=50).grid(row=2, column=1, padx=5)
        
        # Start number
        ttk.Label(frame, text="Nomor Awal:").grid(row=3, column=0, sticky='w', pady=5)
        self.split_start_num = tk.StringVar(value="1")
        ttk.Entry(frame, textvariable=self.split_start_num, width=50).grid(row=3, column=1, padx=5)
        
        # Button
        ttk.Button(frame, text="Mulai Split", command=self.run_split, style='Accent.TButton').grid(row=4, column=1, pady=20)
    
    def setup_merge_tab(self, parent):
        frame = ttk.Frame(parent, padding=20)
        frame.pack(fill='both', expand=True)
        
        # Folder source
        ttk.Label(frame, text="Folder PDF:").grid(row=0, column=0, sticky='w', pady=5)
        self.merge_folder = tk.StringVar()
        ttk.Entry(frame, textvariable=self.merge_folder, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(frame, text="Browse", command=self.browse_merge_folder).grid(row=0, column=2)
        
        # Output file
        ttk.Label(frame, text="Nama File Output:").grid(row=1, column=0, sticky='w', pady=5)
        self.merge_output = tk.StringVar(value="merged.pdf")
        ttk.Entry(frame, textvariable=self.merge_output, width=50).grid(row=1, column=1, padx=5)
        
        # Button
        ttk.Button(frame, text="Mulai Merge", command=self.run_merge, style='Accent.TButton').grid(row=2, column=1, pady=20)
    
    # Helper functions
    def browse_base_folder(self):
        folder = filedialog.askdirectory(initialdir=self.base_folder.get())
        if folder:
            self.base_folder.set(folder)
    
    def browse_excel(self):
        file = filedialog.askopenfilename(
            initialdir=self.base_folder.get(),
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            self.excel_file.set(file)
    
    def browse_split_pdf(self):
        file = filedialog.askopenfilename(
            initialdir=self.base_folder.get(),
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if file:
            self.split_pdf_file.set(file)
    
    def browse_merge_folder(self):
        folder = filedialog.askdirectory(initialdir=self.base_folder.get())
        if folder:
            self.merge_folder.set(folder)
    
    def log(self, message):
        self.log_text.insert(tk.END, message + '\n')
        self.log_text.see(tk.END)
        self.root.update()
    
    def clear_log(self):
        self.log_text.delete('1.0', tk.END)
    
    def safe_filename(self, name):
        safe_name = "".join(c for c in name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        return safe_name[:200] if len(safe_name) > 200 else safe_name
    
    def extract_placeholders(self, pattern):
        return set(m.strip().lower() for m in re.findall(r"{([^{}]+)}", pattern))
    
    def format_with_row(self, pattern, row, col_map):
        def repl(match):
            key = match.group(1).strip().lower()
            if key in col_map:
                v = row[col_map[key]]
                if pd.isna(v):
                    return ""
                return self.safe_filename(str(v).strip())
            return match.group(0)
        return re.sub(r"{([^{}]+)}", repl, pattern)
    
    # Main functions
    def run_rename(self):
        self.clear_log()
        try:
            # Validasi input
            if not self.excel_file.get():
                messagebox.showerror("Error", "Pilih file Excel terlebih dahulu!")
                return
            
            base = self.base_folder.get()
            src_path = os.path.join(base, self.src_folder.get())
            dst_path = os.path.join(base, self.dst_folder.get())
            
            os.makedirs(dst_path, exist_ok=True)
            
            self.log(f"Membaca file Excel: {os.path.basename(self.excel_file.get())}")
            data = pd.read_excel(self.excel_file.get())
            
            col_map = {str(c).strip().lower(): c for c in data.columns}
            self.log(f"Kolom: {', '.join([str(c) for c in data.columns])}")
            self.log(f"Total baris: {len(data)}\n")
            
            # Pattern
            file_pattern = self.file_pattern_src.get()
            if not file_pattern.endswith('.pdf'):
                file_pattern += '.pdf'
            
            output_pattern = self.file_pattern_dst.get()
            if not output_pattern.endswith('.pdf'):
                output_pattern += '.pdf'
            
            # Validasi placeholder
            used_placeholders = self.extract_placeholders(output_pattern)
            missing_cols = [p for p in used_placeholders if p not in col_map]
            if missing_cols:
                messagebox.showerror("Error", f"Kolom tidak ditemukan: {', '.join(missing_cols)}")
                return
            
            try:
                start_num = int(self.start_num.get())
            except ValueError:
                start_num = 1
                self.log("Nomor awal tidak valid, menggunakan 1")
            
            # Process
            processed = not_found = duplicates = skipped = 0
            used_names = {}
            duplicate_list = []
            not_found_list = []
            skipped_list = []
            
            self.log("Memulai proses rename...\n")
            
            for i, row in data.iterrows():
                nomor = start_num + i
                old_name = file_pattern.replace("{nomor}", str(nomor))
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
                    self.log(f"Baris {i+1}: Dilewati (kolom kosong: {', '.join(missing_values)})")
                    continue
                
                base_new_name = self.format_with_row(output_pattern, row, col_map)
                
                # Handle duplicates
                if base_new_name in used_names:
                    used_names[base_new_name] += 1
                    name_parts = base_new_name.rsplit('.', 1)
                    new_name = f"{name_parts[0]} ({used_names[base_new_name]}).{name_parts[1]}"
                    duplicates += 1
                    duplicate_list.append(f"Baris {i+1}: {base_new_name} -> {new_name}")
                    self.log(f"Duplikat: {base_new_name} -> {new_name}")
                else:
                    used_names[base_new_name] = 0
                    new_name = base_new_name
                
                new_path = os.path.join(dst_path, new_name)
                
                if os.path.exists(old_path):
                    shutil.copy2(old_path, new_path)
                    processed += 1
                    self.log(f"[{processed}] {old_name} -> {new_name}")
                else:
                    not_found += 1
                    ctx = ", ".join([f"{p}={str(row[col_map[p]])}" for p in used_placeholders])
                    not_found_list.append(f"Baris {i+1}: {old_name} ({ctx})")
                    self.log(f"File tidak ada: {old_name}")
            
            # Summary
            self.log(f"\n{'='*50}")
            self.log("SUMMARY:")
            self.log(f"Berhasil: {processed}")
            self.log(f"Tidak ditemukan: {not_found}")
            self.log(f"Duplikat: {duplicates}")
            self.log(f"Dilewati: {skipped}")
            self.log(f"Output: {dst_path}")
            self.log(f"{'='*50}")

            # Detail report (log ringkas)
            if not_found_list:
                self.log("DETAIL TIDAK DITEMUKAN:")
                for item in not_found_list:
                    self.log(f" • {item}")
            if duplicate_list:
                self.log("DETAIL DUPLIKAT:")
                for item in duplicate_list:
                    self.log(f" • {item}")
            if skipped_list:
                self.log("DETAIL DILEWATI:")
                for item in skipped_list:
                    self.log(f" • {item}")
            
            messagebox.showinfo("Selesai", f"Rename selesai!\n\nBerhasil: {processed}\nTidak ditemukan: {not_found}\nDilewati: {skipped}")
            
        except Exception as e:
            self.log(f"\nERROR: {str(e)}")
            messagebox.showerror("Error", str(e))
    
    def run_split(self):
        self.clear_log()
        try:
            if not self.split_pdf_file.get():
                messagebox.showerror("Error", "Pilih file PDF terlebih dahulu!")
                return
            
            pdf_path = self.split_pdf_file.get()
            output_folder = os.path.join(self.base_folder.get(), self.split_output.get())
            os.makedirs(output_folder, exist_ok=True)
            
            pattern = self.split_pattern.get()
            if not pattern.endswith('.pdf'):
                pattern += '.pdf'
            
            start_num = int(self.split_start_num.get())
            
            self.log(f"Membaca PDF: {os.path.basename(pdf_path)}")
            reader = PdfReader(pdf_path)
            total = len(reader.pages)
            self.log(f"Total halaman: {total}\n")
            
            self.log("Memulai split...\n")
            
            for i in range(total):
                writer = PdfWriter()
                writer.add_page(reader.pages[i])
                
                filename = pattern.replace("{nomor}", str(start_num + i))
                filepath = os.path.join(output_folder, filename)
                
                with open(filepath, 'wb') as f:
                    writer.write(f)
                
                self.log(f"[{i+1}/{total}] {filename}")
            
            self.log(f"\n{'='*50}")
            self.log(f"Split selesai! Total: {total} file")
            self.log(f"Output: {output_folder}")
            self.log(f"{'='*50}")
            
            messagebox.showinfo("Selesai", f"Split selesai!\n\nTotal: {total} file")
            
        except Exception as e:
            self.log(f"\nERROR: {str(e)}")
            messagebox.showerror("Error", str(e))
    
    def run_merge(self):
        self.clear_log()
        try:
            if not self.merge_folder.get():
                messagebox.showerror("Error", "Pilih folder PDF terlebih dahulu!")
                return
            
            folder = self.merge_folder.get()
            pdf_files = sorted([f for f in os.listdir(folder) if f.endswith('.pdf')])
            
            if not pdf_files:
                messagebox.showerror("Error", "Tidak ada file PDF di folder!")
                return
            
            self.log(f"Ditemukan {len(pdf_files)} file PDF\n")
            
            merger = PdfMerger()
            
            self.log("Memulai merge...\n")
            
            for i, pdf_file in enumerate(pdf_files, 1):
                pdf_path = os.path.join(folder, pdf_file)
                merger.append(pdf_path)
                self.log(f"[{i}/{len(pdf_files)}] {pdf_file}")
            
            output_name = self.merge_output.get()
            if not output_name.endswith('.pdf'):
                output_name += '.pdf'
            
            output_path = os.path.join(self.base_folder.get(), output_name)
            
            self.log(f"\nMenyimpan: {output_name}")
            merger.write(output_path)
            merger.close()
            
            self.log(f"\n{'='*50}")
            self.log(f"Merge selesai! Total: {len(pdf_files)} file")
            self.log(f"Output: {output_path}")
            self.log(f"{'='*50}")
            
            messagebox.showinfo("Selesai", f"Merge selesai!\n\nTotal: {len(pdf_files)} file")
            
        except Exception as e:
            self.log(f"\nERROR: {str(e)}")
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateToolGUI(root)
    root.mainloop()