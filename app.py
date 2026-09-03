"""
app.py
Aplikasi desktop pemeriksa Laporan Rekap Absen + banding Struk Gaji.

Proses berat (parse PDF, audit, banding, ekspor) dijalankan di background
thread agar jendela tidak freeze. Progress 0–100% di-update lewat queue.
"""

import os
import queue
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

import rules
import absensi_parser
import gaji_parser
import pengupahan_parser
import bank_transfer_parser
import compare
import exporter


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pemeriksa Absensi & Struk Gaji - PT. URASE PRIMA")
        self._pusatkan_window(1360, 880)
        self.minsize(1180, 720)

        self.data_absensi = []
        self.anomali_absensi = []
        self.data_gaji = []
        self.hasil_banding = []
        self.data_pengupahan = []
        self.pengupahan_parse_lengkap = None
        self.hasil_banding_pengupahan = []
        self.data_bank = []
        self.hasil_banding_bank = []

        self.mode_var = tk.StringVar(value=rules.MODE_NORMAL)
        self._busy = False
        self._ui_queue = queue.Queue()
        self._tombol = []

        self._setup_style()
        self._build_ui()
        self.after(40, self._proses_antrian_ui)
        self._catat_log("Aplikasi siap. Pilih mode, lalu buka PDF absensi.")

    def _mode_aktif(self):
        return rules.normalisasi_mode(self.mode_var.get())

    def _pusatkan_window(self, lebar, tinggi):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, (sw - lebar) // 2)
        y = max(0, (sh - tinggi) // 2)
        self.geometry(f"{lebar}x{tinggi}+{x}+{y}")

    def _teks_status_mode(self):
        label, aturan = rules.deskripsi_mode(self._mode_aktif())
        return f"Aktif: {label}  |  Aturan: {aturan}"

    # ---------- Tema & Style ----------
    COLORS = {
        "bg": "#f3f5fa",
        "card_bg": "#ffffff",
        "header_bg": "#173b7a",
        "header_bg2": "#2454a6",
        "primary": "#2454a6",
        "primary_dark": "#173b7a",
        "accent": "#0e8f6e",
        "accent_dark": "#0b7259",
        "warn": "#c98a12",
        "danger": "#c0392b",
        "text": "#1f2937",
        "muted": "#64748b",
        "border": "#dde3ee",
        "tab_bg": "#e6ebf5",
        "ok_bg": "#e6f7ef",
        "ok_fg": "#0f5132",
        "bad_bg": "#fdecea",
        "bad_fg": "#7a1f1f",
    }

    def _setup_style(self):
        C = self.COLORS
        self.configure(bg=C["bg"])

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        base_font = ("Segoe UI", 9)
        bold_font = ("Segoe UI", 9, "bold")

        style.configure(".", font=base_font, background=C["bg"], foreground=C["text"])
        style.configure("TFrame", background=C["bg"])
        style.configure("TLabel", background=C["bg"], foreground=C["text"])
        style.configure("Muted.TLabel", background=C["bg"], foreground=C["muted"])
        style.configure("CardTitle.TLabel", background=C["card_bg"], foreground=C["primary_dark"],
                        font=("Segoe UI", 10, "bold"))
        style.configure("CardStatus.TLabel", background=C["card_bg"], foreground=C["muted"],
                        font=("Segoe UI", 8, "italic"))
        style.configure("Mode.TLabel", background=C["bg"], foreground=C["primary_dark"],
                        font=("Segoe UI", 9, "bold"))

        style.configure("Card.TFrame", background=C["card_bg"], relief="solid", borderwidth=1)
        style.configure("Card.TLabel", background=C["card_bg"], foreground=C["text"])

        # Tombol tahap (biru) & tombol aksi utama/hijau (bandingkan)
        style.configure("Step.TButton", font=base_font, padding=(10, 7),
                        background="#eef2f9", foreground=C["primary_dark"], borderwidth=1,
                        focusthickness=0)
        style.map("Step.TButton",
                  background=[("active", "#dfe7f5"), ("disabled", "#f1f3f7")],
                  foreground=[("disabled", "#a7b0c0")])

        style.configure("Accent.TButton", font=bold_font, padding=(10, 8),
                        background=C["accent"], foreground="white", borderwidth=0)
        style.map("Accent.TButton",
                  background=[("active", C["accent_dark"]), ("disabled", "#a7d9c7")],
                  foreground=[("disabled", "#eefaf5")])

        style.configure("Export.TButton", font=bold_font, padding=(14, 9),
                        background=C["primary"], foreground="white", borderwidth=0)
        style.map("Export.TButton",
                  background=[("active", C["primary_dark"]), ("disabled", "#9fb4d6")],
                  foreground=[("disabled", "#eaf0fb")])

        style.configure("TRadiobutton", background=C["bg"], foreground=C["text"], font=base_font)

        style.configure("TLabelframe", background=C["bg"], foreground=C["primary_dark"],
                        borderwidth=1, relief="solid")
        style.configure("TLabelframe.Label", background=C["bg"], foreground=C["primary_dark"],
                        font=bold_font)

        style.configure("Horizontal.TProgressbar", troughcolor="#e4e9f2",
                        background=C["accent"], bordercolor="#e4e9f2",
                        lightcolor=C["accent"], darkcolor=C["accent"], thickness=14)

        style.configure("TNotebook", background=C["bg"], borderwidth=0, tabmargins=(6, 6, 6, 0))
        style.configure("TNotebook.Tab", background=C["tab_bg"], foreground=C["text"],
                        font=base_font, padding=(10, 6))
        style.map("TNotebook.Tab",
                  background=[("selected", C["card_bg"])],
                  foreground=[("selected", C["primary_dark"])],
                  font=[("selected", bold_font)])

        style.configure("Treeview", font=base_font, rowheight=25, background="white",
                        fieldbackground="white", foreground=C["text"], borderwidth=0)
        style.configure("Treeview.Heading", font=bold_font, background=C["primary"],
                        foreground="white", relief="flat", padding=(6, 6))
        style.map("Treeview.Heading", background=[("active", C["primary_dark"])])
        style.map("Treeview", background=[("selected", C["primary"])],
                  foreground=[("selected", "white")])

        style.configure("Status.TLabel", background="#eef2f9", foreground=C["text"],
                        font=base_font, padding=(8, 5))

    # ---------- UI ----------
    def _build_ui(self):
        C = self.COLORS

        # ===== Header banner =====
        header = tk.Frame(self, bg=C["header_bg"])
        header.pack(fill="x")
        header_in = tk.Frame(header, bg=C["header_bg"])
        header_in.pack(fill="x", padx=18, pady=12)
        header_in.columnconfigure(0, weight=1)

        judul_box = tk.Frame(header_in, bg=C["header_bg"])
        judul_box.grid(row=0, column=0, sticky="w")
        tk.Label(judul_box, text="📊 Pemeriksa Absensi & Struk Gaji", bg=C["header_bg"],
                 fg="white", font=("Segoe UI", 15, "bold")).pack(anchor="w")
        tk.Label(judul_box, text="PT. URASE PRIMA — Audit Absensi, Gaji, Pengupahan & Transfer Bank",
                 bg=C["header_bg"], fg="#c7d5ee", font=("Segoe UI", 9)).pack(anchor="w")

        self.btn_tentang = ttk.Button(header_in, text="ℹ Tentang", style="Step.TButton",
                                      command=self._buka_tentang)
        self.btn_tentang.grid(row=0, column=1, sticky="e", padx=(12, 0))

        self.btn_ekspor = ttk.Button(header_in, text="⬇ Ekspor ke Excel", style="Export.TButton",
                                     command=self.ekspor)
        self.btn_ekspor.grid(row=0, column=2, sticky="e", padx=(12, 0))

        # ===== Baris kartu tahapan (responsif: 3 kolom, lebar proporsional) =====
        cards_wrap = ttk.Frame(self, padding=(12, 10, 12, 4))
        cards_wrap.pack(fill="x")
        cards_wrap.columnconfigure(0, weight=1, uniform="cards")
        cards_wrap.columnconfigure(1, weight=1, uniform="cards")
        cards_wrap.columnconfigure(2, weight=1, uniform="cards")

        # -- Kartu 1: Absensi & Struk Gaji --
        card1 = ttk.LabelFrame(cards_wrap, text=" 🗂️  Absensi & Struk Gaji ", padding=10)
        card1.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        card1.columnconfigure(0, weight=1)

        self.btn_absensi = ttk.Button(card1, text="1. Buka PDF Absensi", style="Step.TButton",
                                      command=self.buka_absensi)
        self.btn_absensi.grid(row=0, column=0, sticky="ew", pady=(0, 3))
        self.lbl_absensi = ttk.Label(card1, text="belum ada file", style="Muted.TLabel", wraplength=280)
        self.lbl_absensi.grid(row=1, column=0, sticky="w", pady=(0, 8))

        self.btn_gaji = ttk.Button(card1, text="2. Buka PDF Struk Gaji", style="Step.TButton",
                                   command=self.buka_gaji)
        self.btn_gaji.grid(row=2, column=0, sticky="ew", pady=(0, 3))
        self.lbl_gaji = ttk.Label(card1, text="belum ada file", style="Muted.TLabel", wraplength=280)
        self.lbl_gaji.grid(row=3, column=0, sticky="w", pady=(0, 8))

        self.btn_banding = ttk.Button(card1, text="3. Bandingkan vs Struk Gaji", style="Accent.TButton",
                                      command=self.jalankan_banding)
        self.btn_banding.grid(row=4, column=0, sticky="ew")

        # -- Kartu 2: Laporan Pengupahan --
        card2 = ttk.LabelFrame(cards_wrap, text=" 🧾  Laporan Pengupahan ", padding=10)
        card2.grid(row=0, column=1, sticky="nsew", padx=6)
        card2.columnconfigure(0, weight=1)

        self.btn_pengupahan = ttk.Button(card2, text="4. Buka PDF Laporan Pengupahan", style="Step.TButton",
                                         command=self.buka_pengupahan)
        self.btn_pengupahan.grid(row=0, column=0, sticky="ew", pady=(0, 3))
        self.lbl_pengupahan = ttk.Label(card2, text="belum ada file", style="Muted.TLabel", wraplength=280)
        self.lbl_pengupahan.grid(row=1, column=0, sticky="w", pady=(0, 8))

        # Spacer setinggi persis blok "tombol file ke-2 + label" di kartu 1
        # (63px, diukur langsung dari render), supaya tombol hijau di ketiga
        # kartu benar-benar sejajar horizontal, bukan cuma didekati.
        ttk.Frame(card2, height=63).grid(row=2, column=0, sticky="ew")

        self.btn_banding_pengupahan = ttk.Button(
            card2, text="5. Bandingkan vs Struk Gaji", style="Accent.TButton",
            command=self.jalankan_banding_pengupahan)
        self.btn_banding_pengupahan.grid(row=3, column=0, sticky="ew")

        # -- Kartu 3: Transfer Bank --
        card3 = ttk.LabelFrame(cards_wrap, text=" 🏦  Transfer Bank ", padding=10)
        card3.grid(row=0, column=2, sticky="nsew", padx=(6, 0))
        card3.columnconfigure(0, weight=1)

        self.btn_bank = ttk.Button(card3, text="6. Buka Excel Transfer Bank", style="Step.TButton",
                                   command=self.buka_bank)
        self.btn_bank.grid(row=0, column=0, sticky="ew", pady=(0, 3))
        self.lbl_bank = ttk.Label(card3, text="belum ada file", style="Muted.TLabel", wraplength=280)
        self.lbl_bank.grid(row=1, column=0, sticky="w", pady=(0, 8))

        ttk.Frame(card3, height=63).grid(row=2, column=0, sticky="ew")

        self.btn_banding_bank = ttk.Button(
            card3, text="7. Bandingkan vs Struk Gaji", style="Accent.TButton",
            command=self.jalankan_banding_bank)
        self.btn_banding_bank.grid(row=3, column=0, sticky="ew")

        self._tombol = [self.btn_absensi, self.btn_gaji, self.btn_banding, self.btn_ekspor,
                        self.btn_pengupahan, self.btn_banding_pengupahan,
                        self.btn_bank, self.btn_banding_bank]

        # ===== Mode pemrosesan =====
        mode_bar = ttk.LabelFrame(self, text=" ⚙ Mode Pemrosesan (pilih sebelum audit / banding) ", padding=8)
        mode_bar.pack(fill="x", padx=12, pady=(6, 4))

        self.rb_normal = ttk.Radiobutton(
            mode_bar, text="Mode Normal — HKP dihitung otomatis (batas group − izin)",
            value=rules.MODE_NORMAL, variable=self.mode_var,
            command=self._on_mode_berubah,
        )
        self.rb_normal.pack(side="left", padx=8)
        self.rb_sipil = ttk.Radiobutton(
            mode_bar, text="Mode Sipil — HKP dari dokumen, tidak dihitung otomatis",
            value=rules.MODE_SIPIL, variable=self.mode_var,
            command=self._on_mode_berubah,
        )
        self.rb_sipil.pack(side="left", padx=8)

        self.lbl_mode = ttk.Label(self, text=self._teks_status_mode(), style="Mode.TLabel",
                                  wraplength=1100, justify="left")
        self.lbl_mode.pack(fill="x", padx=16, pady=(0, 2))

        # ===== Progres =====
        prog = ttk.Frame(self, padding=(12, 2, 12, 6))
        prog.pack(fill="x")
        ttk.Label(prog, text="Progres:").pack(side="left")
        self.progress = ttk.Progressbar(prog, orient="horizontal", mode="determinate",
                                        maximum=100, length=420)
        self.progress.pack(side="left", fill="x", expand=True, padx=8)
        self.lbl_persen = ttk.Label(prog, text="0%", width=6)
        self.lbl_persen.pack(side="left")
        self.lbl_progres_teks = ttk.Label(self, text="Siap.", style="Muted.TLabel",
                                          wraplength=1100, justify="left")
        self.lbl_progres_teks.pack(fill="x", padx=16)

        # ===== Notebook (tab hasil) =====
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=10)

        frame1 = ttk.Frame(self.notebook)
        self.notebook.add(frame1, text="⚠ Anomali Absensi")
        self.tree_anomali = self._buat_tabel(
            frame1, ["NRP", "Nama", "Aturan", "Detail", "Tanggal"],
            lebar=[80, 170, 230, 420, 80], stretch_kolom="Detail")

        frame2 = ttk.Frame(self.notebook)
        self.notebook.add(frame2, text="⇄ Banding Absensi")
        self.tree_banding = self._buat_tabel(
            frame2, ["NRP", "Nama (Absensi)", "Nama (Struk Gaji)", "Status",
                     "Mode", "Sumber HKP", "Detail Perbedaan"],
            lebar=[90, 170, 170, 160, 90, 200, 420])

        frame_gaji = ttk.Frame(self.notebook)
        self.notebook.add(frame_gaji, text="💵 Struk Gaji")
        self.tree_gaji = self._buat_tabel(
            frame_gaji,
            ["NRP", "Nama", "Mode", "Sumber HKP", "HKP dokumen", "HKP dipakai",
             "HKP otomatis", "U.Pokok", "Upah Kotor", "U.Bersih", "Catatan HKP"],
            lebar=[90, 170, 90, 200, 100, 100, 100, 110, 110, 110, 400])

        frame_pengupahan = ttk.Frame(self.notebook)
        self.notebook.add(frame_pengupahan, text="🧾 Pengupahan")
        self.tree_pengupahan = self._buat_tabel(
            frame_pengupahan,
            ["NRP", "Nama", "Bagian", "Bagian/Departemen", "Upah Kotor",
             "BPJS Kesehatan", "BPJS Tenaga Kerja", "Pot. Lain-lain", "Upah Bersih"],
            lebar=[90, 180, 150, 170, 120, 130, 140, 120, 120])

        frame_banding_pengupahan = ttk.Frame(self.notebook)
        self.notebook.add(frame_banding_pengupahan, text="⇄ Banding Pengupahan")
        self.tree_banding_pengupahan = self._buat_tabel(
            frame_banding_pengupahan,
            ["NRP", "Nama (Pengupahan)", "Nama (Struk Gaji)", "Status", "Detail Perbedaan"],
            lebar=[90, 190, 190, 220, 400])

        frame_bank = ttk.Frame(self.notebook)
        self.notebook.add(frame_bank, text="🏦 Transfer Bank")
        self.tree_bank = self._buat_tabel(
            frame_bank,
            ["Trx ID", "Tipe Transfer", "No. Rekening", "Nama Penerima", "Jumlah", "Remark"],
            lebar=[120, 130, 150, 220, 130, 250])

        frame_banding_bank = ttk.Frame(self.notebook)
        self.notebook.add(frame_banding_bank, text="⇄ Banding Bank")
        self.tree_banding_bank = self._buat_tabel(
            frame_banding_bank,
            ["Nama", "Jumlah Transfer Bank", "Upah Bersih Struk Gaji", "Status", "Detail"],
            lebar=[190, 150, 160, 300, 350])

        frame3 = ttk.Frame(self.notebook)
        self.notebook.add(frame3, text="📝 Log Pemrosesan")
        self.txt_log = scrolledtext.ScrolledText(frame3, height=12, wrap="word", state="disabled",
                                                 font=("Consolas", 9), bg="white", relief="flat",
                                                 borderwidth=1)
        self.txt_log.pack(fill="both", expand=True, padx=2, pady=2)

        # ===== Status bar =====
        self.status_bar = ttk.Label(self, text="Siap. " + self._teks_status_mode(),
                                    style="Status.TLabel", anchor="w",
                                    wraplength=1100, justify="left")
        self.status_bar.pack(fill="x", side="bottom")

        # Wraplength label² di atas mengikuti lebar jendela supaya teks panjang
        # (mis. deskripsi aturan mode / ringkasan banding) tidak terpotong di
        # tepi layar saat window di-resize — ini bagian "responsive" utamanya.
        self.bind("<Configure>", self._on_resize)

    def _on_resize(self, event):
        if event.widget is not self:
            return
        lebar = max(400, event.width - 40)
        for lbl in (self.lbl_mode, self.lbl_progres_teks, self.status_bar):
            if lbl.cget("wraplength") != lebar:
                lbl.configure(wraplength=lebar)

    def _buat_tabel(self, parent, kolom, lebar=None, stretch_kolom=None):
        wrap = ttk.Frame(parent)
        wrap.pack(fill="both", expand=True)

        tree = ttk.Treeview(wrap, columns=kolom, show="headings")
        idx_stretch = kolom.index(stretch_kolom) if stretch_kolom in (kolom or []) else len(kolom) - 1
        for idx, k in enumerate(kolom):
            tree.heading(k, text=k, anchor="w")
            w = lebar[idx] if lebar and idx < len(lebar) else 130
            # Kolom teks panjang (mis. "Detail"/"Catatan") melar mengisi sisa
            # ruang saat window dilebarkan; kolom lain tetap pada lebar tetap.
            tree.column(k, width=w, minwidth=50, anchor="w", stretch=(idx == idx_stretch))

        vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        tree.tag_configure("bad", background=self.COLORS["bad_bg"], foreground=self.COLORS["bad_fg"])
        tree.tag_configure("ok", background=self.COLORS["ok_bg"], foreground=self.COLORS["ok_fg"])
        tree.tag_configure("warn", background="#fff6e0", foreground="#7a5b0b")
        return tree

    # ---------- Threading / progress ----------
    def _set_busy(self, busy):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for b in self._tombol:
            b.config(state=state)
        self.rb_normal.config(state=state)
        self.rb_sipil.config(state=state)

    def _set_progress(self, persen, pesan=""):
        persen = int(max(0, min(100, persen)))
        self.progress["value"] = persen
        self.lbl_persen.config(text=f"{persen}%")
        if pesan:
            self.lbl_progres_teks.config(text=pesan)

    def _cb_progress(self, persen, pesan=""):
        self._ui_queue.put(("progress", (persen, pesan)))

    @staticmethod
    def _skala(progress, awal, akhir):
        rentang = akhir - awal

        def inner(pct, msg=""):
            progress(awal + int(rentang * max(0, min(100, pct)) / 100), msg)

        return inner

    def _jalankan_latar(self, pekerjaan, selesai, gagal=None, judul="Memproses..."):
        if self._busy:
            messagebox.showinfo("Masih memproses", "Tunggu proses sebelumnya selesai.")
            return
        self._set_busy(True)
        self._set_progress(0, judul)
        self._catat_log(judul)

        def worker():
            try:
                hasil = pekerjaan(self._cb_progress)
                self._ui_queue.put(("done", (selesai, hasil)))
            except Exception as e:
                self._ui_queue.put(("error", (gagal, e)))

        threading.Thread(target=worker, daemon=True).start()

    def _proses_antrian_ui(self):
        try:
            while True:
                jenis, payload = self._ui_queue.get_nowait()
                if jenis == "progress":
                    persen, pesan = payload
                    self._set_progress(persen, pesan)
                elif jenis == "log":
                    self._catat_log(payload)
                elif jenis == "done":
                    cb, hasil = payload
                    self._set_busy(False)
                    self._set_progress(100, "Selesai.")
                    if cb:
                        cb(hasil)
                elif jenis == "error":
                    cb, err = payload
                    self._set_busy(False)
                    self._set_progress(0, "Gagal.")
                    if cb:
                        cb(err)
                    else:
                        messagebox.showerror("Kesalahan", str(err))
        except queue.Empty:
            pass
        self.after(40, self._proses_antrian_ui)

    def _catat_log(self, pesan):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"[{stamp}] {pesan}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _refresh_label_mode(self):
        self.lbl_mode.config(text=self._teks_status_mode())

    def _isi_tabel_anomali(self):
        self.tree_anomali.delete(*self.tree_anomali.get_children())
        for a in self.anomali_absensi:
            self.tree_anomali.insert("", "end", values=(a.nrp, a.nama, a.aturan, a.detail, a.tanggal),
                                     tags=("bad",))

    def _isi_tabel_banding(self):
        self.tree_banding.delete(*self.tree_banding.get_children())
        for h in self.hasil_banding:
            tag = "ok" if h.status == "COCOK" else "bad"
            self.tree_banding.insert("", "end", values=(
                h.nrp, h.nama_absensi, h.nama_gaji, h.status,
                h.mode, h.sumber_hkp, "; ".join(h.detail)), tags=(tag,))

    def _isi_tabel_gaji(self):
        self.tree_gaji.delete(*self.tree_gaji.get_children())
        for g in self.data_gaji:
            self.tree_gaji.insert("", "end", values=(
                g.get("nrp"), g.get("nama"), g.get("mode_label") or g.get("mode"),
                g.get("sumber_hkp"), g.get("hkp_dokumen", g.get("hkp")),
                g.get("hkp_dipakai"), g.get("hkp_hitung_otomatis"),
                g.get("u_pokok"), g.get("upah_kotor"), g.get("u_bersih"),
                g.get("catatan_hkp"),
            ))

    def _isi_tabel_pengupahan(self):
        self.tree_pengupahan.delete(*self.tree_pengupahan.get_children())
        for k in self.data_pengupahan:
            self.tree_pengupahan.insert("", "end", values=(
                k.get("nrp"), k.get("nama"), k.get("bagian"), k.get("bagian_section"),
                k.get("upah_kotor"), k.get("bpjs_kesehatan"), k.get("bpjs_tenaga_kerja"),
                k.get("pot_lain"), k.get("upah_bersih"),
            ))

    def _isi_tabel_banding_pengupahan(self):
        self.tree_banding_pengupahan.delete(*self.tree_banding_pengupahan.get_children())
        for h in self.hasil_banding_pengupahan:
            tag = "ok" if h.status == "COCOK" else "bad"
            self.tree_banding_pengupahan.insert("", "end", values=(
                h.nrp, h.nama_pengupahan, h.nama_gaji, h.status, "; ".join(h.detail)), tags=(tag,))

    def _isi_tabel_bank(self):
        self.tree_bank.delete(*self.tree_bank.get_children())
        for b in self.data_bank:
            self.tree_bank.insert("", "end", values=(
                b.get("trx_id"), b.get("transfer_type"), b.get("credited_account"),
                b.get("nama"), b.get("jumlah"), b.get("remark"),
            ))

    def _isi_tabel_banding_bank(self):
        self.tree_banding_bank.delete(*self.tree_banding_bank.get_children())
        for h in self.hasil_banding_bank:
            if h.status == "COCOK":
                tag = "ok"
            elif h.status in ("TIDAK DITRANSFER (UPAH 0 - WAJAR)", "AMBIGU (NAMA KEMBAR)"):
                tag = "warn"
            else:
                tag = "bad"
            self.tree_banding_bank.insert("", "end", values=(
                h.nama, h.jumlah_bank, h.jumlah_gaji, h.status, "; ".join(h.detail)), tags=(tag,))

    def _buka_tentang(self):
        C = self.COLORS
        top = tk.Toplevel(self)
        top.title("Tentang Aplikasi")
        top.configure(bg=C["card_bg"])
        top.resizable(False, False)
        top.transient(self)
        top.grab_set()

        pad = tk.Frame(top, bg=C["card_bg"], padx=30, pady=26)
        pad.pack()

        tk.Label(pad, text="📊 Pemeriksa Absensi & Struk Gaji", bg=C["card_bg"],
                 fg=C["primary_dark"], font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(pad, text="PT. URASE PRIMA", bg=C["card_bg"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 16))

        ttk.Separator(pad).pack(fill="x", pady=(0, 16))

        tk.Label(pad, text="DIRANCANG OLEH", bg=C["card_bg"], fg=C["muted"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        tk.Label(pad, text="Handrian Rivaldi", bg=C["card_bg"], fg=C["text"],
                 font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(2, 2))
        tk.Label(pad, text="Dibuat tahun 2026", bg=C["card_bg"], fg=C["muted"],
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 12))

        link = tk.Label(pad, text="🔗 github.com/HandrianRivaldi74", bg=C["card_bg"],
                        fg=C["primary"], font=("Segoe UI", 10, "underline"), cursor="hand2")
        link.pack(anchor="w")
        link.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/HandrianRivaldi74"))

        ttk.Button(pad, text="Tutup", style="Step.TButton", command=top.destroy).pack(anchor="e", pady=(24, 0))

        top.update_idletasks()
        x = self.winfo_rootx() + (self.winfo_width() - top.winfo_width()) // 2
        y = self.winfo_rooty() + (self.winfo_height() - top.winfo_height()) // 2
        top.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _on_mode_berubah(self):
        if self._busy:
            return
        self._refresh_label_mode()
        label, aturan = rules.deskripsi_mode(self._mode_aktif())
        self._catat_log(f"Mode diganti ke {label}. Aturan: {aturan}")
        if self.hasil_banding:
            self.hasil_banding = []
            self.tree_banding.delete(*self.tree_banding.get_children())
            self._catat_log("Hasil banding sebelumnya dibersihkan. Jalankan ulang langkah 3.")
        if self.data_absensi:
            self._audit_absensi_latar(f"Mengulang audit absensi dengan {label}")
        elif self.data_gaji:
            self._anotasi_gaji_latar()
        else:
            self.status_bar.config(text=self._teks_status_mode())

    # ---------- Aksi ----------
    def buka_absensi(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        mode = self._mode_aktif()
        gaji_awal = list(self.data_gaji)

        def kerja(progress):
            data = absensi_parser.parse_absensi(path, progress=self._skala(progress, 0, 55))
            anomali = rules.jalankan_semua_aturan_banyak(
                data, mode=mode, progress=self._skala(progress, 55, 85))
            gaji = gaji_awal
            if gaji:
                gaji_parser.terapkan_konteks_mode(
                    gaji, data_absensi=data, mode=mode, progress=self._skala(progress, 85, 100))
            else:
                progress(100, "Selesai")
            return path, data, anomali, gaji

        def selesai(hasil):
            path_, data, anomali, gaji = hasil
            self.data_absensi = data
            self.anomali_absensi = anomali
            if gaji:
                self.data_gaji = gaji
                self._isi_tabel_gaji()
            self.lbl_absensi.config(text=os.path.basename(path_) + f" ({len(data)} karyawan)")
            self._isi_tabel_anomali()
            self._refresh_label_mode()
            label, aturan = rules.deskripsi_mode(mode)
            self.status_bar.config(
                text=f"Absensi diperiksa [{label}]: {len(data)} karyawan, "
                     f"{len(anomali)} anomali. Aturan: {aturan}")
            self._catat_log(
                f"PDF absensi dimuat: {os.path.basename(path_)} ({len(data)} karyawan). "
                f"Mode={label}. Anomali={len(anomali)}.")

        def gagal(err):
            messagebox.showerror("Gagal membaca PDF Absensi", str(err))
            self._catat_log(f"Gagal parse absensi: {err}")

        self._jalankan_latar(kerja, selesai, gagal, judul="Membaca PDF absensi...")

    def buka_gaji(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        mode = self._mode_aktif()
        absen_awal = list(self.data_absensi)

        def kerja(progress):
            data = gaji_parser.parse_struk_gaji(path, mode=mode, progress=self._skala(progress, 0, 75))
            gaji_parser.terapkan_konteks_mode(
                data, data_absensi=absen_awal, mode=mode, progress=self._skala(progress, 75, 100))
            return path, data

        def selesai(hasil):
            path_, data = hasil
            self.data_gaji = data
            self.lbl_gaji.config(text=os.path.basename(path_) + f" ({len(data)} karyawan)")
            self._isi_tabel_gaji()
            label, _ = rules.deskripsi_mode(mode)
            self.status_bar.config(
                text=f"Struk gaji dimuat: {len(data)} karyawan. {self._teks_status_mode()}")
            self._catat_log(
                f"PDF struk gaji dimuat: {os.path.basename(path_)} ({len(data)} karyawan). "
                f"Mode={label}.")
            self.notebook.select(2)

        def gagal(err):
            messagebox.showerror("Gagal membaca PDF Struk Gaji", str(err))
            self._catat_log(f"Gagal parse struk gaji: {err}")

        self._jalankan_latar(kerja, selesai, gagal, judul="Membaca PDF struk gaji...")

    def buka_pengupahan(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not path:
            return

        def kerja(progress):
            hasil = pengupahan_parser.parse_pengupahan(path, progress=self._skala(progress, 0, 90))
            masalah = pengupahan_parser.validasi_internal(hasil)
            progress(100, "Selesai")
            return path, hasil, masalah

        def selesai(payload):
            path_, hasil, masalah = payload
            self.data_pengupahan = hasil["karyawan"]
            self.pengupahan_parse_lengkap = hasil
            self.lbl_pengupahan.config(text=os.path.basename(path_) + f" ({len(hasil['karyawan'])} karyawan)")
            self._isi_tabel_pengupahan()
            n = len(hasil["karyawan"])
            if masalah:
                self.status_bar.config(
                    text=f"Laporan pengupahan dimuat: {n} karyawan. "
                         f"{len(masalah)} peringatan validasi internal (subtotal/grand total) — lihat Log Pemrosesan.")
                self._catat_log(
                    f"PDF laporan pengupahan dimuat: {os.path.basename(path_)} ({n} karyawan). "
                    f"Peringatan validasi internal ({len(masalah)}):")
                for m in masalah:
                    self._catat_log("  - " + m)
            else:
                self.status_bar.config(
                    text=f"Laporan pengupahan dimuat: {n} karyawan. Validasi internal OK (subtotal & grand total cocok).")
                self._catat_log(
                    f"PDF laporan pengupahan dimuat: {os.path.basename(path_)} ({n} karyawan). "
                    f"Validasi internal OK (subtotal & grand total cocok).")
            self.notebook.select(3)

        def gagal(err):
            messagebox.showerror("Gagal membaca PDF Laporan Pengupahan", str(err))
            self._catat_log(f"Gagal parse laporan pengupahan: {err}")

        self._jalankan_latar(kerja, selesai, gagal, judul="Membaca PDF laporan pengupahan...")

    def buka_bank(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not path:
            return

        def kerja(progress):
            return path, bank_transfer_parser.parse_transfer_bank(path, progress=progress)

        def selesai(payload):
            path_, data = payload
            self.data_bank = data
            self.lbl_bank.config(text=os.path.basename(path_) + f" ({len(data)} transaksi)")
            self._isi_tabel_bank()
            total = sum((b.get("jumlah") or 0) for b in data)
            self.status_bar.config(
                text=f"Transfer bank dimuat: {len(data)} transaksi, total = {total:,.0f}".replace(",", "."))
            self._catat_log(
                f"Excel transfer bank dimuat: {os.path.basename(path_)} "
                f"({len(data)} transaksi, total = {total:,.0f})".replace(",", "."))
            self.notebook.select(5)

        def gagal(err):
            messagebox.showerror("Gagal membaca Excel Transfer Bank", str(err))
            self._catat_log(f"Gagal parse transfer bank: {err}")

        self._jalankan_latar(kerja, selesai, gagal, judul="Membaca Excel transfer bank...")

    def _audit_absensi_latar(self, judul):
        mode = self._mode_aktif()
        data = list(self.data_absensi)
        gaji = list(self.data_gaji)

        def kerja(progress):
            anomali = rules.jalankan_semua_aturan_banyak(data, mode=mode, progress=progress)
            if gaji:
                gaji_parser.terapkan_konteks_mode(gaji, data_absensi=data, mode=mode, progress=progress)
            return anomali, gaji

        def selesai(hasil):
            anomali, gaji_baru = hasil
            self.anomali_absensi = anomali
            if gaji_baru:
                self.data_gaji = gaji_baru
                self._isi_tabel_gaji()
            self._isi_tabel_anomali()
            self._refresh_label_mode()
            label, aturan = rules.deskripsi_mode(mode)
            self.status_bar.config(
                text=f"Absensi diperiksa [{label}]: {len(self.data_absensi)} karyawan, "
                     f"{len(anomali)} anomali. Aturan: {aturan}")
            self._catat_log(
                f"Audit absensi selesai. Mode={label}. Aturan: {aturan}. Anomali: {len(anomali)}.")

        self._jalankan_latar(kerja, selesai, judul=judul)

    def _anotasi_gaji_latar(self):
        mode = self._mode_aktif()
        gaji = list(self.data_gaji)
        absen = list(self.data_absensi)

        def kerja(progress):
            gaji_parser.terapkan_konteks_mode(gaji, data_absensi=absen, mode=mode, progress=progress)
            return gaji

        def selesai(data):
            self.data_gaji = data
            self._isi_tabel_gaji()
            self.status_bar.config(text=self._teks_status_mode())
            self._catat_log("Konteks mode struk gaji diperbarui.")

        self._jalankan_latar(kerja, selesai, judul="Memperbarui konteks struk gaji...")

    def jalankan_banding(self):
        if not self.data_absensi:
            messagebox.showwarning("Data belum lengkap", "Buka dulu PDF Absensi (langkah 1).")
            return
        if not self.data_gaji:
            messagebox.showwarning("Data belum lengkap", "Buka dulu PDF Struk Gaji (langkah 2).")
            return
        mode = self._mode_aktif()
        absen = list(self.data_absensi)
        gaji = list(self.data_gaji)

        def kerja(progress):
            gaji_parser.terapkan_konteks_mode(gaji, data_absensi=absen, mode=mode, progress=progress)
            hasil = compare.bandingkan(absen, gaji, mode=mode, progress=progress)
            return hasil, gaji

        def selesai(payload):
            hasil, gaji_baru = payload
            self.hasil_banding = hasil
            self.data_gaji = gaji_baru
            self._isi_tabel_banding()
            self._isi_tabel_gaji()
            ring = compare.ringkasan(hasil)
            label, aturan = rules.deskripsi_mode(mode)
            self.status_bar.config(
                text=f"Banding selesai [{label}]: {ring['cocok']} cocok, {ring['tidak_sinkron']} tidak sinkron, "
                     f"{ring['hanya_di_absensi']} hanya di absensi, {ring['hanya_di_struk_gaji']} hanya di struk gaji. "
                     f"Sumber HKP: {ring.get('sumber_hkp') or '-'}")
            self._catat_log(
                f"Banding selesai. Mode={label}. Sumber HKP: {ring.get('sumber_hkp')}. "
                f"Aturan: {aturan}. Cocok={ring['cocok']}, tidak sinkron={ring['tidak_sinkron']}.")
            self.notebook.select(1)

        self._jalankan_latar(kerja, selesai, judul="Membandingkan absensi vs struk gaji...")

    def jalankan_banding_pengupahan(self):
        if not self.data_pengupahan:
            messagebox.showwarning("Data belum lengkap", "Buka dulu PDF Laporan Pengupahan (langkah 4).")
            return
        if not self.data_gaji:
            messagebox.showwarning("Data belum lengkap", "Buka dulu PDF Struk Gaji (langkah 2).")
            return
        pengupahan = list(self.data_pengupahan)
        gaji = list(self.data_gaji)

        def kerja(progress):
            return compare.bandingkan_pengupahan(pengupahan, gaji, progress=progress)

        def selesai(hasil):
            self.hasil_banding_pengupahan = hasil
            self._isi_tabel_banding_pengupahan()
            ring = compare.ringkasan_pengupahan(hasil)
            self.status_bar.config(
                text=f"Banding pengupahan selesai: {ring['cocok']} cocok, {ring['tidak_sinkron']} tidak sinkron, "
                     f"{ring['hanya_di_pengupahan']} hanya di laporan pengupahan, "
                     f"{ring['hanya_di_struk_gaji']} hanya di struk gaji.")
            self._catat_log(
                f"Banding pengupahan selesai. Cocok={ring['cocok']}, tidak sinkron={ring['tidak_sinkron']}, "
                f"hanya di pengupahan={ring['hanya_di_pengupahan']}, hanya di struk gaji={ring['hanya_di_struk_gaji']}.")
            self.notebook.select(4)

        self._jalankan_latar(kerja, selesai, judul="Membandingkan laporan pengupahan vs struk gaji...")

    def jalankan_banding_bank(self):
        if not self.data_bank:
            messagebox.showwarning("Data belum lengkap", "Buka dulu Excel Transfer Bank (langkah 6).")
            return
        if not self.data_gaji:
            messagebox.showwarning("Data belum lengkap", "Buka dulu PDF Struk Gaji (langkah 2).")
            return
        bank = list(self.data_bank)
        gaji = list(self.data_gaji)

        def kerja(progress):
            return compare.bandingkan_transfer_bank(bank, gaji, progress=progress)

        def selesai(hasil):
            self.hasil_banding_bank = hasil
            self._isi_tabel_banding_bank()
            ring = compare.ringkasan_transfer_bank(hasil)
            self.status_bar.config(
                text=f"Banding transfer bank selesai: {ring['cocok']} cocok, {ring['tidak_sinkron']} tidak sinkron, "
                     f"{ring['tidak_ditransfer_wajar']} upah 0 (wajar), {ring['hanya_di_bank']} hanya di bank, "
                     f"{ring['hanya_di_struk_gaji']} belum ditransfer, {ring['ambigu']} ambigu (nama kembar).")
            self._catat_log(
                f"Banding transfer bank selesai. Cocok={ring['cocok']}, tidak sinkron={ring['tidak_sinkron']}, "
                f"hanya di bank={ring['hanya_di_bank']}, hanya di struk gaji={ring['hanya_di_struk_gaji']}, "
                f"ambigu={ring['ambigu']}.")
            self.notebook.select(6)

        self._jalankan_latar(kerja, selesai, judul="Membandingkan transfer bank vs struk gaji...")

    def ekspor(self):
        if not self.anomali_absensi and not self.hasil_banding and not self.data_gaji \
                and not self.hasil_banding_pengupahan and not self.data_pengupahan \
                and not self.hasil_banding_bank and not self.data_bank:
            messagebox.showwarning("Belum ada hasil", "Jalankan pemeriksaan/banding dulu sebelum ekspor.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                             initialfile="hasil_pemeriksaan.xlsx")
        if not path:
            return
        mode = self._mode_aktif()
        anomali = list(self.anomali_absensi)
        banding = list(self.hasil_banding)
        gaji = list(self.data_gaji)
        ring = compare.ringkasan(banding) if banding else {
            "mode": mode, "sumber_hkp": rules.deskripsi_mode(mode)[1],
        }
        banding_pengupahan = list(self.hasil_banding_pengupahan)
        ring_pengupahan = compare.ringkasan_pengupahan(banding_pengupahan) if banding_pengupahan else None
        banding_bank = list(self.hasil_banding_bank)
        ring_bank = compare.ringkasan_transfer_bank(banding_bank) if banding_bank else None
        pengupahan = list(self.data_pengupahan)
        bank = list(self.data_bank)

        def kerja(progress):
            return exporter.ekspor_hasil(
                path, anomali, banding, ring, data_gaji=gaji, mode=mode, progress=progress,
                data_pengupahan=pengupahan,
                hasil_banding_pengupahan=banding_pengupahan,
                ringkasan_banding_pengupahan=ring_pengupahan,
                data_bank=bank,
                hasil_banding_bank=banding_bank,
                ringkasan_banding_bank=ring_bank)

        def selesai(out):
            self._catat_log(f"Hasil diekspor ke {out} (mode={mode}).")
            messagebox.showinfo("Selesai", f"Hasil diekspor ke:\n{out}")

        def gagal(err):
            messagebox.showerror("Gagal ekspor", str(err))

        self._jalankan_latar(kerja, selesai, gagal, judul="Mengekspor Excel...")


if __name__ == "__main__":
    App().mainloop()
