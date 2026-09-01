"""
gaji_parser.py
Parser untuk "LAPORAN UPAH KARYAWAN" (struk gaji) PDF.

Struktur tabel per baris karyawan (hasil extract_table pdfplumber):
[NO, "NRP\\nNAMA", HKL, LL, HKP, IJIN, DIRUMAHKAN, U.POKOK, U.LEMBUR,
 U.LEMBUR_LIBUR, UPAH_IJIN, UPAH_DIRUMAHKAN, "U.LAIN-LAIN\\nTRANSPORT",
 "TUNJANGAN\\nT.JABATAN\\nT.KHUSUS", UPAH_KOTOR, "BPJS_KT\\nBPJS_KS\\nPOT.LAIN",
 U.BERSIH, KETERANGAN, TTD]

Baris "TOTAL ..." adalah subtotal per bagian/departemen, bukan data karyawan - dilewati.

Konteks mode (Sipil vs Normal) dilekatkan pada setiap baris struk:
- Mode Sipil : HKP struk = nilai tercetak dokumen; tidak dihitung ulang.
- Mode Normal: HKP tercetak tetap disimpan; HKP audit memakai rumus otomatis
  dari data absensi (lihat terapkan_konteks_mode).
"""

import re
import pdfplumber

import rules


def _lapor(progress, persen, pesan=""):
    if progress:
        progress(int(max(0, min(100, persen))), pesan)


def _to_float(s):
    if s is None:
        return None
    s = s.strip()
    if s == "" or s == "-":
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    s = s.replace(".", "").replace(",", ".")
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None


def _to_int(s):
    v = _to_float(s)
    return int(v) if v is not None else None


def _split_multiline(cell, n):
    """Pecah cell multi-baris ('a\\nb\\nc') jadi list float, isi None kalau kurang."""
    if cell is None:
        parts = []
    else:
        parts = [p for p in cell.split("\n") if p.strip() != ""]
    vals = [_to_float(p) for p in parts]
    while len(vals) < n:
        vals.append(None)
    return vals[:n]


def parse_struk_gaji(pdf_path, mode=None, progress=None):
    """
    Mengembalikan list of dict, satu per karyawan.
    Kolom HKP dibaca apa adanya dari tabel struk (tidak dihitung ulang di parser).
    mode: rules.MODE_SIPIL atau rules.MODE_NORMAL — dilekatkan ke setiap baris.
    progress(persen, pesan) untuk UI thread-safe via queue.
    """
    mode = rules.normalisasi_mode(mode)
    mode_label, aturan = rules.deskripsi_mode(mode)
    karyawan_list = []
    departemen, bagian = None, None
    pending_labels = []

    _lapor(progress, 1, f"Membuka PDF struk gaji ({mode_label})...")
    with pdfplumber.open(pdf_path) as pdf:
        n_hal = len(pdf.pages) or 1
        for pi, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            deps = [d.strip() for d in re.findall(r"DEPARTEMENT\s+([A-Z0-9 .,/-]+)", text)]
            bags = [b.strip() for b in re.findall(r"BAGIAN\s+([A-Z0-9 .,/-]+)", text)]
            for i in range(max(len(deps), len(bags))):
                pending_labels.append((
                    deps[i] if i < len(deps) else None,
                    bags[i] if i < len(bags) else None,
                ))

            tables = page.extract_tables()
            for table in tables:
                if pending_labels:
                    d, b = pending_labels.pop(0)
                    departemen = d or departemen
                    bagian = b or bagian
                for row in table:
                    if not row or len(row) < 17:
                        continue
                    no_cell = (row[0] or "").strip()
                    if not no_cell.isdigit():
                        continue

                    nrp_nama = (row[1] or "").split("\n")
                    nrp = nrp_nama[0].strip() if nrp_nama else None
                    nama = " ".join(p.strip() for p in nrp_nama[1:]).strip()

                    u_lain, transport = _split_multiline(row[12], 2)
                    tunjangan, t_jabatan, t_khusus = _split_multiline(row[13], 3)
                    bpjs_kt, bpjs_ks, pot_lain = _split_multiline(row[15], 3)
                    hkp_dokumen = _to_float(row[4])

                    rec = {
                        "nrp": nrp,
                        "nama": nama,
                        "departemen": departemen,
                        "bagian": bagian,
                        "hkl": _to_float(row[2]),
                        "ll": _to_float(row[3]),
                        "hkp": hkp_dokumen,
                        "hkp_dokumen": hkp_dokumen,
                        "ijin": _to_int(row[5]),
                        "dirumahkan": _to_int(row[6]),
                        "u_pokok": _to_float(row[7]),
                        "u_lembur": _to_float(row[8]),
                        "u_lembur_libur": _to_float(row[9]),
                        "upah_ijin": _to_float(row[10]),
                        "upah_dirumahkan": _to_float(row[11]),
                        "u_lain_lain": u_lain,
                        "transport": transport,
                        "tunjangan": tunjangan,
                        "t_jabatan": t_jabatan,
                        "t_khusus": t_khusus,
                        "upah_kotor": _to_float(row[14]),
                        "bpjs_kt": bpjs_kt,
                        "bpjs_ks": bpjs_ks,
                        "pot_lain": pot_lain,
                        "u_bersih": _to_float(row[16]),
                        "mode": mode,
                        "mode_label": mode_label,
                        "aturan_mode": aturan,
                    }
                    karyawan_list.append(rec)

            _lapor(progress, 5 + int(80 * (pi + 1) / n_hal),
                   f"Parse struk gaji halaman {pi + 1}/{n_hal} ({len(karyawan_list)} karyawan)")

    terapkan_konteks_mode(karyawan_list, data_absensi=None, mode=mode, progress=None)
    _lapor(progress, 100, f"Selesai parse struk gaji ({len(karyawan_list)} karyawan, {mode_label})")
    return karyawan_list


def terapkan_konteks_mode(data_gaji, data_absensi=None, mode=None, progress=None):
    """
    Lengkapi setiap baris struk dengan HKP sesuai mode aktif.
    Mode Sipil : hkp_dipakai = HKP tercetak dokumen; hitung otomatis tidak dijalankan.
    Mode Normal: hkp_dipakai = hasil rules.hitung_hkp dari absensi (jika ada);
                 HKP tercetak tetap disimpan di hkp / hkp_dokumen.
    """
    if not data_gaji:
        return data_gaji
    mode = rules.normalisasi_mode(mode if mode is not None else data_gaji[0].get("mode"))
    mode_label, aturan = rules.deskripsi_mode(mode)
    by_absen = {k["id"]: k for k in (data_absensi or [])}
    n = len(data_gaji)

    for i, g in enumerate(data_gaji):
        g["mode"] = mode
        g["mode_label"] = mode_label
        g["aturan_mode"] = aturan
        g["hkp_dokumen"] = g.get("hkp_dokumen", g.get("hkp"))

        if mode == rules.MODE_SIPIL:
            g["hkp_hitung_otomatis"] = None
            g["hkp_batas_group"] = None
            g["hkp_dipakai"] = g.get("hkp_dokumen")
            g["sumber_hkp"] = "dokumen struk (manual, tidak dihitung ulang)"
            g["catatan_hkp"] = (
                "Mode Sipil: rincian gaji memakai HKP tercetak dokumen. "
                "Rumus HKP otomatis (batas group − izin) tidak dijalankan."
            )
        else:
            absen = by_absen.get(g.get("nrp"))
            if absen:
                hkp_hitung, batas, ada_tidak_steril = rules.hitung_hkp(absen)
                g["hkp_hitung_otomatis"] = hkp_hitung
                g["hkp_batas_group"] = batas
                g["hkp_dipakai"] = hkp_hitung
                g["sumber_hkp"] = "hitung otomatis absensi (batas group − izin)"
                if ada_tidak_steril:
                    g["catatan_hkp"] = (
                        f"Mode Normal: data izin belum steril ('?'). "
                        f"HKP tercetak struk = {g.get('hkp_dokumen')}; hitung otomatis ditunda."
                    )
                    g["hkp_dipakai"] = None
                else:
                    g["catatan_hkp"] = (
                        f"Mode Normal: HKP otomatis = {hkp_hitung} (batas group {batas}). "
                        f"HKP tercetak struk = {g.get('hkp_dokumen')}."
                    )
            else:
                g["hkp_hitung_otomatis"] = None
                g["hkp_batas_group"] = None
                g["hkp_dipakai"] = None
                g["sumber_hkp"] = "menunggu data absensi (rumus otomatis)"
                g["catatan_hkp"] = (
                    "Mode Normal: absensi belum cocok NRP ini. "
                    f"HKP tercetak struk = {g.get('hkp_dokumen')} disimpan, "
                    "belum diganti rumus otomatis."
                )

        if i % 5 == 0 or i + 1 == n:
            _lapor(progress, 100 * (i + 1) / n,
                   f"Konteks mode struk gaji {i + 1}/{n} ({mode_label})")
    return data_gaji


if __name__ == "__main__":
    import sys, json
    data = parse_struk_gaji(sys.argv[1])
    print(f"Total karyawan terbaca: {len(data)}")
    print(json.dumps(data[:3], indent=2, ensure_ascii=False))
