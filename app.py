"""
app.py
Aplikasi desktop pemeriksa Laporan Rekap Absen + banding Struk Gaji.

Proses berat (parse PDF, audit, banding, ekspor) dijalankan di background
thread agar jendela tidak freeze. Progress 0–100% di-update lewat queue.
"""

import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

import rules
import absensi_parser
import gaji_parser
import pengupahan_parser
import compare
import exporter


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pemeriksa Absensi & Struk Gaji - PT. URASE PRIMA")
        self.geometry("1180x800")

        self.data_absensi = []
        self.anomali_absensi = []
        self.data_gaji = []
        self.hasil_banding = []
        self.data_pengupahan = []
        self.pengupahan_parse_lengkap = None
        self.hasil_banding_pengupahan = []

        self.mode_var = tk.StringVar(value=rules.MODE_NORMAL)
        self._busy = False
        self._ui_queue = queue.Queue()
        self._tombol = []

        self._build_ui()
        self.after(40, self._proses_antrian_ui)
        self._catat_log("Aplikasi siap. Pilih mode, lalu buka PDF absensi.")

    def _mode_aktif(self):
        return rules.normalisasi_mode(self.mode_var.get())

    def _teks_status_mode(self):
        label, aturan = rules.deskripsi_mode(self._mode_aktif())
        return f"Aktif: {label}  |  Aturan: {aturan}"

    # ---------- UI ----------
    def _build_ui(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        self.btn_absensi = ttk.Button(top, text="1. Buka PDF Absensi", command=self.buka_absensi)
        self.btn_absensi.pack(side="left", padx=4)
        self.lbl_absensi = ttk.Label(top, text="belum ada file")
        self.lbl_absensi.pack(side="left", padx=4)

        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=10)

        self.btn_gaji = ttk.Button(top, text="2. Buka PDF Struk Gaji", command=self.buka_gaji)
        self.btn_gaji.pack(side="left", padx=4)
        self.lbl_gaji = ttk.Label(top, text="belum ada file")
        self.lbl_gaji.pack(side="left", padx=4)

        ttk.Separator(top, orient="vertical").pack(side="left", fill="y", padx=10)

        self.btn_banding = ttk.Button(top, text="3. Bandingkan", command=self.jalankan_banding)
        self.btn_banding.pack(side="left", padx=4)
        self.btn_ekspor = ttk.Button(top, text="Ekspor ke Excel", command=self.ekspor)
        self.btn_ekspor.pack(side="left", padx=4)

        top2 = ttk.Frame(self, padding=(8, 0, 8, 8))
        top2.pack(fill="x")

        self.btn_pengupahan = ttk.Button(top2, text="4. Buka PDF Laporan Pengupahan", command=self.buka_pengupahan)
        self.btn_pengupahan.pack(side="left", padx=4)
        self.lbl_pengupahan = ttk.Label(top2, text="belum ada file")
        self.lbl_pengupahan.pack(side="left", padx=4)

        ttk.Separator(top2, orient="vertical").pack(side="left", fill="y", padx=10)

        self.btn_banding_pengupahan = ttk.Button(
            top2, text="5. Bandingkan Pengupahan vs Struk Gaji", command=self.jalankan_banding_pengupahan)
        self.btn_banding_pengupahan.pack(side="left", padx=4)

        self._tombol = [self.btn_absensi, self.btn_gaji, self.btn_banding, self.btn_ekspor,
                        self.btn_pengupahan, self.btn_banding_pengupahan]

        mode_bar = ttk.LabelFrame(self, text="Mode Pemrosesan (pilih sebelum audit / banding)", padding=6)
        mode_bar.pack(fill="x", padx=8, pady=(0, 4))

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

        self.lbl_mode = ttk.Label(self, text=self._teks_status_mode(), foreground="#1a365d")
        self.lbl_mode.pack(fill="x", padx=10, pady=(0, 2))

        prog = ttk.Frame(self, padding=(8, 0, 8, 4))
        prog.pack(fill="x")
        ttk.Label(prog, text="Progres:").pack(side="left")
        self.progress = ttk.Progressbar(prog, orient="horizontal", mode="determinate",
                                        maximum=100, length=420)
        self.progress.pack(side="left", fill="x", expand=True, padx=8)
        self.lbl_persen = ttk.Label(prog, text="0%", width=6)
        self.lbl_persen.pack(side="left")
        self.lbl_progres_teks = ttk.Label(self, text="Siap.", foreground="#444")
        self.lbl_progres_teks.pack(fill="x", padx=10)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        frame1 = ttk.Frame(self.notebook)
        self.notebook.add(frame1, text="Anomali Absensi")
        self.tree_anomali = self._buat_tabel(frame1, ["NRP", "Nama", "Aturan", "Detail", "Tanggal"])

        frame2 = ttk.Frame(self.notebook)
        self.notebook.add(frame2, text="Banding Absensi vs Struk Gaji")
        self.tree_banding = self._buat_tabel(
            frame2, ["NRP", "Nama (Absensi)", "Nama (Struk Gaji)", "Status",
                     "Mode", "Sumber HKP", "Detail Perbedaan"])

        frame_gaji = ttk.Frame(self.notebook)
        self.notebook.add(frame_gaji, text="Struk Gaji")
        self.tree_gaji = self._buat_tabel(
            frame_gaji,
            ["NRP", "Nama", "Mode", "Sumber HKP", "HKP dokumen", "HKP dipakai",
             "HKP otomatis", "U.Pokok", "Upah Kotor", "U.Bersih", "Catatan HKP"])

        frame_pengupahan = ttk.Frame(self.notebook)
        self.notebook.add(frame_pengupahan, text="Laporan Pengupahan")
        self.tree_pengupahan = self._buat_tabel(
            frame_pengupahan,
            ["NRP", "Nama", "Bagian", "Bagian/Departemen", "Upah Kotor",
             "BPJS Kesehatan", "BPJS Tenaga Kerja", "Pot. Lain-lain", "Upah Bersih"])

        frame_banding_pengupahan = ttk.Frame(self.notebook)
        self.notebook.add(frame_banding_pengupahan, text="Banding Pengupahan")
        self.tree_banding_pengupahan = self._buat_tabel(
            frame_banding_pengupahan,
            ["NRP", "Nama (Pengupahan)", "Nama (Struk Gaji)", "Status", "Detail Perbedaan"])

        frame3 = ttk.Frame(self.notebook)
        self.notebook.add(frame3, text="Log Pemrosesan")
        self.txt_log = scrolledtext.ScrolledText(frame3, height=12, wrap="word", state="disabled")
        self.txt_log.pack(fill="both", expand=True)

        self.status_bar = ttk.Label(self, text="Siap. " + self._teks_status_mode(),
                                    relief="sunken", anchor="w")
        self.status_bar.pack(fill="x", side="bottom")

    def _buat_tabel(self, parent, kolom):
        tree = ttk.Treeview(parent, columns=kolom, show="headings")
        for k in kolom:
            tree.heading(k, text=k)
            tree.column(k, width=130, anchor="w")
        vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        tree.tag_configure("bad", background="#fce4e4")
        tree.tag_configure("ok", background="#e2f0d9")
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

    def ekspor(self):
        if not self.anomali_absensi and not self.hasil_banding and not self.data_gaji \
                and not self.hasil_banding_pengupahan and not self.data_pengupahan:
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

        def kerja(progress):
            return exporter.ekspor_hasil(
                path, anomali, banding, ring, data_gaji=gaji, mode=mode, progress=progress,
                hasil_banding_pengupahan=banding_pengupahan,
                ringkasan_banding_pengupahan=ring_pengupahan)

        def selesai(out):
            self._catat_log(f"Hasil diekspor ke {out} (mode={mode}).")
            messagebox.showinfo("Selesai", f"Hasil diekspor ke:\n{out}")

        def gagal(err):
            messagebox.showerror("Gagal ekspor", str(err))

        self._jalankan_latar(kerja, selesai, gagal, judul="Mengekspor Excel...")


if __name__ == "__main__":
    App().mainloop()
