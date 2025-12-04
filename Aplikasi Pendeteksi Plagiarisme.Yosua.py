import tkinter as tk
from tkinter import messagebox, filedialog
import string
import zipfile
import xml.etree.ElementTree as ET

# ======================================================
# Fungsi untuk membaca file .docx TANPA library eksternal
# ======================================================
def baca_docx_tanpa_library(path):
    try:
        with zipfile.ZipFile(path) as z:
            if "word/document.xml" not in z.namelist():
                return "[ERROR] File DOCX tidak valid atau corrupt."

            # buka file XML docx sebagai stream (bytes)
            with z.open("word/document.xml") as xml_file:
                xml_content = xml_file.read()

            # parsing root dari XML
            root = ET.fromstring(xml_content)

            # ambil namespace secara otomatis dari stream
            namespaces = {}
            with z.open("word/document.xml") as xml_file_for_ns:
                for event, node in ET.iterparse(xml_file_for_ns, events=['start-ns']):
                    prefix, uri = node
                    # normalisasi prefix: jika prefix == '' gunakan 'w'
                    if prefix == "":
                        prefix = "w"
                    namespaces[prefix] = uri

            # fallback ke namespace default jika kosong
            if not namespaces:
                namespaces = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

            paragraphs = []
            # cari paragraf dan teks
            for p in root.findall('.//w:p', namespaces):
                texts = [t.text for t in p.findall('.//w:t', namespaces) if t.text]
                if texts:
                    paragraphs.append("".join(texts))

            if not paragraphs:
                return "[WARNING] File DOCX terbaca tetapi kosong."

            return "\n".join(paragraphs)

    except Exception as e:
        return f"[ERROR BACA DOCX] {e}"

# ======================================================
# MODUL 1 – Variabel, Tipe Data, Array
# ======================================================
judul = "Aplikasi Cek Plagiarisme"
versi = 1.0
daftar_kata = ["plagiarisme", "cek", "teks", "python"]

def info_modul1():
    print("=== MODUL 1: Variabel, Tipe Data, Array ===")
    print("Judul :", judul)
    print("Versi :", versi)
    print("Kata Penting:", daftar_kata)

# ======================================================
# MODUL 2 – Pengkondisian
# ======================================================
def cek_kemiripan(nilai):
    if nilai >= 80:
        return "Plagiat Tinggi"
    elif nilai >= 50:
        return "Plagiat Sedang"
    elif nilai >= 20:
        return "Plagiat Rendah"
    else:
        return "Tidak Plagiat"

# ======================================================
# MODUL 3 – Perulangan
# ======================================================
def hitung_kemiripan(teks1, teks2):
    kata1 = teks1.lower().split()
    kata2 = teks2.lower().split()

    if len(kata1) == 0:
        return 0

    total = 0
    for kata in kata1:
        if kata in kata2:
            total += 1

    persentase = (total / len(kata1)) * 100
    return round(persentase, 2)

# ======================================================
# MODUL 4 – Function
# ======================================================
def bersihkan(teks):
    for tanda in string.punctuation:
        teks = teks.replace(tanda, "")
    # hapus multiple whitespace jadi satu spasi dan strip
    teks = " ".join(teks.split())
    return teks

# ======================================================
# MODUL 5 – OOP
# ======================================================
class PlagiarismChecker:
    def __init__(self, t1, t2):
        self.t1 = t1
        self.t2 = t2

    def proses(self, func_bersih, func_hitung):
        teks1 = func_bersih(self.t1)
        teks2 = func_bersih(self.t2)
        return func_hitung(teks1, teks2)

# ======================================================
# MODUL 8 – GUI (Aplikasi Cek Plagiarisme)
# ======================================================

root = tk.Tk()
root.title("Smart Plagiarism Checker")
root.geometry("820x650")
root.config(bg="#D9EAF7")

info_modul1()

frame = tk.Frame(root, bg="white", padx=20, pady=20)
frame.pack(pady=20, padx=20, fill="both", expand=True)

tk.Label(frame, text="Aplikasi Cek Plagiarisme",
         font=("Arial", 16, "bold"), bg="white", fg="#003366").pack(pady=8)

# =================== FUNGSI UPLOAD FILE ===================
def upload_file_to_textbox(textbox):
    filepath = filedialog.askopenfilename(
        title="Pilih File Teks",
        filetypes=[
            ("Semua Dokumen", "*.txt *.docx"),
            ("File Teks", "*.txt"),
            ("File Word", "*.docx"),
            ("Semua File", "*.*")
        ]
    )

    if not filepath:
        return

    try:
        # File TXT
        if filepath.lower().endswith(".txt"):
            with open(filepath, "r", encoding="utf-8") as file:
                isi = file.read()

        # File DOCX (tanpa python-docx)
        elif filepath.lower().endswith(".docx"):
            isi = baca_docx_tanpa_library(filepath)
            # jika pesan error/warning kembalikan sebagai messagebox
            if isi.startswith("[ERROR") or isi.startswith("[WARNING"):
                messagebox.showerror("Info File DOCX", isi) if isi.startswith("[ERROR") else messagebox.showwarning("Info File DOCX", isi)

        else:
            messagebox.showwarning("Format Tidak Didukung", "Hanya .txt dan .docx yang didukung!")
            return

        textbox.delete("1.0", "end")
        textbox.insert("1.0", isi)

    except Exception as e:
        messagebox.showerror("Error", f"Gagal membuka file!\n{e}")

# =================== INPUT TEKS 1 ===================
tk.Label(frame, text="TEKS 1:", font=("Arial", 12), bg="white").pack(anchor="w", pady=(6,0))

input1 = tk.Text(frame, height=10, width=90, bd=2, relief="groove", wrap="word")
input1.pack(pady=5)

btn_frame1 = tk.Frame(frame, bg="white")
btn_frame1.pack(fill="x", pady=(0,8))
btn_upload1 = tk.Button(btn_frame1, text="Upload File Teks 1", bg="#A1C9F1", fg="black",
                        command=lambda: upload_file_to_textbox(input1))
btn_upload1.pack(side="left", padx=(0,6))

btn_clear1 = tk.Button(btn_frame1, text="Clear", command=lambda: input1.delete("1.0", "end"))
btn_clear1.pack(side="left")

# =================== INPUT TEKS 2 ===================
tk.Label(frame, text="TEKS 2:", font=("Arial", 12), bg="white").pack(anchor="w", pady=(6,0))

input2 = tk.Text(frame, height=10, width=90, bd=2, relief="groove", wrap="word")
input2.pack(pady=5)

btn_frame2 = tk.Frame(frame, bg="white")
btn_frame2.pack(fill="x", pady=(0,8))
btn_upload2 = tk.Button(btn_frame2, text="Upload File Teks 2", bg="#A1C9F1", fg="black",
                        command=lambda: upload_file_to_textbox(input2))
btn_upload2.pack(side="left", padx=(0,6))

btn_clear2 = tk.Button(btn_frame2, text="Clear", command=lambda: input2.delete("1.0", "end"))
btn_clear2.pack(side="left")

# =================== FUNGSI CEK ===================
def cek_plagiarisme():
    t1 = input1.get("1.0", "end").strip()
    t2 = input2.get("1.0", "end").strip()

    if not t1 or not t2:
        messagebox.showwarning("Peringatan", "Kedua teks harus diisi!")
        return

    checker = PlagiarismChecker(t1, t2)
    persentase = checker.proses(bersihkan, hitung_kemiripan)
    kategori = cek_kemiripan(persentase)

    hasil = f"Persentase Kemiripan: {persentase}%\nKategori: {kategori}"
    messagebox.showinfo("Hasil Cek", hasil)

# =================== TOMBOL CEK ===================
tk.Button(frame,
          text="CEK PLAGIARISME",
          font=("Arial", 13, "bold"),
          bg="#4A90E2",
          fg="white",
          padx=10, pady=6,
          activebackground="#357ABD",
          cursor="hand2",
          command=cek_plagiarisme).pack(pady=12)

# =================== FOOTER INFO ===================
footer = tk.Label(root, text=f"{judul} — Versi {versi}", bg="#D9EAF7", fg="#333333")
footer.pack(side="bottom", pady=6)

root.mainloop()
