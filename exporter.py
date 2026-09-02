"""
exporter.py
Ekspor hasil pemeriksaan absensi (anomali), hasil banding, dan rincian struk gaji
(sesuai Mode Sipil / Mode Normal) ke satu file Excel.
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

import rules

HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
HEADER_FONT = Font(bold=True)
BAD_FILL = PatternFill(start_color="FCE4E4", end_color="FCE4E4", fill_type="solid")
OK_FILL = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
MODE_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")


def _lapor(progress, persen, pesan=""):
    if progress:
        progress(int(max(0, min(100, persen))), pesan)


def _tulis_header(ws, kolom):
    for i, judul in enumerate(kolom, start=1):
        c = ws.cell(row=1, column=i, value=judul)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
    ws.freeze_panes = "A2"


def _lebarkan_kolom(ws, lebar):
    from openpyxl.utils import get_column_letter
    for i, w in enumerate(lebar, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _tulis_sheet_struk(wb, data_gaji, mode):
    ws = wb.create_sheet("Struk Gaji")
    mode = rules.normalisasi_mode(mode)
    label, aturan = rules.deskripsi_mode(mode)
    ws["A1"] = f"STRUK GAJI — {label}"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = f"Aturan HKP: {aturan}"
    ws.merge_cells("A2:F2")

    kolom = [
        "NRP", "Nama", "Departemen", "Bagian",
        "Mode", "Sumber HKP",
        "HKP dokumen struk", "HKP dipakai (audit)", "HKP hitung otomatis", "Batas group",
        "HKL", "LL", "IJIN", "DIRUMAHKAN",
        "U.Pokok", "U.Lembur", "U.Lembur Libur", "Upah Ijin", "Upah Dirumahkan",
        "U.Lain-lain", "Transport", "Tunjangan", "T.Jabatan", "T.Khusus",
        "Upah Kotor", "BPJS KT", "BPJS KS", "Pot.Lain", "U.Bersih",
        "Catatan HKP",
    ]
    for i, judul in enumerate(kolom, start=1):
        c = ws.cell(row=4, column=i, value=judul)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
    ws.freeze_panes = "A5"

    if data_gaji:
        for r, g in enumerate(data_gaji, start=5):
            nilai = [
                g.get("nrp"), g.get("nama"), g.get("departemen"), g.get("bagian"),
                g.get("mode_label") or g.get("mode"), g.get("sumber_hkp"),
                g.get("hkp_dokumen", g.get("hkp")), g.get("hkp_dipakai"),
                g.get("hkp_hitung_otomatis"), g.get("hkp_batas_group"),
                g.get("hkl"), g.get("ll"), g.get("ijin"), g.get("dirumahkan"),
                g.get("u_pokok"), g.get("u_lembur"), g.get("u_lembur_libur"),
                g.get("upah_ijin"), g.get("upah_dirumahkan"),
                g.get("u_lain_lain"), g.get("transport"), g.get("tunjangan"),
                g.get("t_jabatan"), g.get("t_khusus"),
                g.get("upah_kotor"), g.get("bpjs_kt"), g.get("bpjs_ks"),
                g.get("pot_lain"), g.get("u_bersih"),
                g.get("catatan_hkp"),
            ]
            for c, v in enumerate(nilai, start=1):
                cell = ws.cell(r, c, v)
                if c <= 10:
                    cell.fill = MODE_FILL
                cell.alignment = Alignment(wrap_text=True, vertical="top")

    _lebarkan_kolom(ws, [12, 26, 18, 16, 28, 36, 16, 16, 18, 12,
                         10, 10, 10, 14, 14, 12, 16, 12, 16,
                         12, 12, 12, 12, 12, 14, 12, 12, 12, 14, 70])
    return ws


def ekspor_hasil(path_output, anomali_absensi=None, hasil_banding=None, ringkasan_banding=None,
                 data_gaji=None, mode=None, progress=None,
                 hasil_banding_pengupahan=None, ringkasan_banding_pengupahan=None,
                 hasil_banding_bank=None, ringkasan_banding_bank=None):
    _lapor(progress, 2, "Menyiapkan workbook Excel...")
    wb = Workbook()
    mode = rules.normalisasi_mode(
        mode if mode is not None else (ringkasan_banding or {}).get("mode")
    )

    ws1 = wb.active
    ws1.title = "Anomali Absensi"
    _tulis_header(ws1, ["NRP", "Nama", "Aturan", "Detail", "Tanggal"])
    _lebarkan_kolom(ws1, [12, 28, 12, 60, 10])
    if anomali_absensi:
        n = len(anomali_absensi)
        for r, a in enumerate(anomali_absensi, start=2):
            ws1.cell(r, 1, a.nrp)
            ws1.cell(r, 2, a.nama)
            ws1.cell(r, 3, a.aturan)
            ws1.cell(r, 4, a.detail)
            ws1.cell(r, 5, a.tanggal)
            for c in range(1, 6):
                ws1.cell(r, c).fill = BAD_FILL
            if r % 20 == 0:
                _lapor(progress, 5 + int(15 * (r - 1) / n), f"Ekspor anomali {r - 1}/{n}")
    _lapor(progress, 20, "Sheet anomali selesai")

    ws2 = wb.create_sheet("Banding Absensi vs Gaji")
    _tulis_header(ws2, ["NRP", "Nama (Absensi)", "Nama (Struk Gaji)", "Status",
                        "Mode", "Sumber HKP", "Detail Perbedaan"])
    _lebarkan_kolom(ws2, [12, 26, 26, 16, 14, 36, 70])
    if hasil_banding:
        n = len(hasil_banding)
        for r, h in enumerate(hasil_banding, start=2):
            ws2.cell(r, 1, h.nrp)
            ws2.cell(r, 2, h.nama_absensi)
            ws2.cell(r, 3, h.nama_gaji)
            ws2.cell(r, 4, h.status)
            ws2.cell(r, 5, getattr(h, "mode", "") or "")
            ws2.cell(r, 6, getattr(h, "sumber_hkp", "") or "")
            ws2.cell(r, 7, "; ".join(h.detail) if h.detail else "")
            fill = OK_FILL if h.status == "COCOK" else BAD_FILL
            for c in range(1, 8):
                ws2.cell(r, c).fill = fill
                ws2.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
            if r % 20 == 0:
                _lapor(progress, 25 + int(20 * (r - 1) / n), f"Ekspor banding {r - 1}/{n}")
    _lapor(progress, 50, "Sheet banding selesai")

    if hasil_banding_pengupahan:
        ws2b = wb.create_sheet("Banding Pengupahan")
        _tulis_header(ws2b, ["NRP", "Nama (Laporan Pengupahan)", "Nama (Struk Gaji)",
                              "Status", "Detail Perbedaan"])
        _lebarkan_kolom(ws2b, [12, 30, 30, 16, 80])
        n = len(hasil_banding_pengupahan)
        for r, h in enumerate(hasil_banding_pengupahan, start=2):
            ws2b.cell(r, 1, h.nrp)
            ws2b.cell(r, 2, h.nama_pengupahan)
            ws2b.cell(r, 3, h.nama_gaji)
            ws2b.cell(r, 4, h.status)
            ws2b.cell(r, 5, "; ".join(h.detail) if h.detail else "")
            fill = OK_FILL if h.status == "COCOK" else BAD_FILL
            for c in range(1, 6):
                ws2b.cell(r, c).fill = fill
                ws2b.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
        _lapor(progress, 55, "Sheet banding pengupahan selesai")

    if hasil_banding_bank:
        ws2c = wb.create_sheet("Banding Transfer Bank")
        _tulis_header(ws2c, ["Nama (dinormalisasi)", "Jumlah Transfer Bank", "Upah Bersih Struk Gaji",
                              "Status", "Detail"])
        _lebarkan_kolom(ws2c, [30, 20, 20, 34, 70])
        n = len(hasil_banding_bank)
        for r, h in enumerate(hasil_banding_bank, start=2):
            ws2c.cell(r, 1, h.nama)
            ws2c.cell(r, 2, h.jumlah_bank)
            ws2c.cell(r, 3, h.jumlah_gaji)
            ws2c.cell(r, 4, h.status)
            ws2c.cell(r, 5, "; ".join(h.detail) if h.detail else "")
            fill = OK_FILL if h.status == "COCOK" else (
                MODE_FILL if h.status in ("TIDAK DITRANSFER (UPAH 0 - WAJAR)",) else BAD_FILL)
            for c in range(1, 6):
                ws2c.cell(r, c).fill = fill
                ws2c.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
        _lapor(progress, 58, "Sheet banding transfer bank selesai")

    ws3 = wb.create_sheet("Ringkasan")
    label, aturan = rules.deskripsi_mode(mode)
    ws3.cell(1, 1, "Ringkasan Hasil Banding").font = Font(bold=True, size=13)
    baris = [
        ("Mode pemrosesan", (ringkasan_banding or {}).get("mode") or mode),
        ("Label mode", label),
        ("Sumber / aturan HKP", (ringkasan_banding or {}).get("sumber_hkp") or aturan),
    ]
    if ringkasan_banding:
        baris += [
            ("Total karyawan dibandingkan", ringkasan_banding.get("total")),
            ("Cocok / sinkron", ringkasan_banding.get("cocok")),
            ("Tidak sinkron (perlu diperbaiki)", ringkasan_banding.get("tidak_sinkron")),
            ("Hanya ada di data absensi", ringkasan_banding.get("hanya_di_absensi")),
            ("Hanya ada di struk gaji", ringkasan_banding.get("hanya_di_struk_gaji")),
        ]
    for i, (judul, val) in enumerate(baris, start=3):
        ws3.cell(i, 1, judul)
        ws3.cell(i, 2, val)
        if i <= 5:
            ws3.cell(i, 1).fill = MODE_FILL
            ws3.cell(i, 2).fill = MODE_FILL

    if ringkasan_banding_pengupahan:
        baris_awal = len(baris) + 4
        ws3.cell(baris_awal, 1, "Ringkasan Banding Pengupahan").font = Font(bold=True, size=13)
        baris_pu = [
            ("Total karyawan dibandingkan (pengupahan)", ringkasan_banding_pengupahan.get("total")),
            ("Cocok / sinkron (pengupahan)", ringkasan_banding_pengupahan.get("cocok")),
            ("Tidak sinkron (pengupahan)", ringkasan_banding_pengupahan.get("tidak_sinkron")),
            ("Hanya ada di laporan pengupahan", ringkasan_banding_pengupahan.get("hanya_di_pengupahan")),
            ("Hanya ada di struk gaji (pengupahan)", ringkasan_banding_pengupahan.get("hanya_di_struk_gaji")),
        ]
        for j, (judul, val) in enumerate(baris_pu, start=baris_awal + 1):
            ws3.cell(j, 1, judul)
            ws3.cell(j, 2, val)
        baris_akhir_pu = baris_awal + len(baris_pu)
    else:
        baris_akhir_pu = len(baris) + 3

    if ringkasan_banding_bank:
        baris_awal_bank = baris_akhir_pu + 3
        ws3.cell(baris_awal_bank, 1, "Ringkasan Banding Transfer Bank").font = Font(bold=True, size=13)
        baris_bank = [
            ("Total nama dibandingkan (transfer bank)", ringkasan_banding_bank.get("total")),
            ("Cocok / sinkron (transfer bank)", ringkasan_banding_bank.get("cocok")),
            ("Tidak sinkron (transfer bank)", ringkasan_banding_bank.get("tidak_sinkron")),
            ("Tidak ditransfer - upah 0 (wajar)", ringkasan_banding_bank.get("tidak_ditransfer_wajar")),
            ("Hanya ada di transfer bank", ringkasan_banding_bank.get("hanya_di_bank")),
            ("Hanya ada di struk gaji (belum ditransfer)", ringkasan_banding_bank.get("hanya_di_struk_gaji")),
            ("Ambigu - nama kembar", ringkasan_banding_bank.get("ambigu")),
        ]
        for j, (judul, val) in enumerate(baris_bank, start=baris_awal_bank + 1):
            ws3.cell(j, 1, judul)
            ws3.cell(j, 2, val)

    _lebarkan_kolom(ws3, [36, 50])
    _lapor(progress, 60, "Sheet ringkasan selesai")

    if data_gaji:
        _tulis_sheet_struk(wb, data_gaji, mode)
        _lapor(progress, 90, "Sheet struk gaji selesai")

    _lapor(progress, 95, "Menyimpan file Excel...")
    wb.save(path_output)
    _lapor(progress, 100, f"Ekspor selesai: {path_output}")
    return path_output
