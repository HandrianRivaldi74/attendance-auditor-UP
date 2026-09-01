"""
pengupahan_parser.py
Parser untuk "LAPORAN PENGUPAHAN KARYAWAN" (ringkasan total upah per karyawan,
dikelompokkan per BAGIAN dengan subtotal TOTAL per bagian dan GRAND TOTAL di
akhir dokumen). Dipakai untuk audit-silang lima nilai finansial (UPAH KOTOR,
BPJS KESEHATAN, BPJS TENAGA KERJA, POT. LAIN-LAIN, UPAH BERSIH) terhadap
struk gaji hasil gaji_parser.py.

CATATAN STRUKTUR PDF: dokumen ini hanya punya garis tabel di baris HEADER
kolom (NRP/NAMA KARYAWAN/BAGIAN/...); baris data & baris TOTAL tidak
bergaris sama sekali. Akibatnya pdfplumber.extract_tables() cuma menangkap
baris header (terverifikasi lewat page.rects — garis cuma ada di y=header).
Parser ini karena itu merekonstruksi baris lewat posisi kata
(page.extract_words(), dikelompokkan per `top`) lalu memetakan tiap kata ke
kolom berdasarkan rentang x0 tetap (KOLOM_X), diambil dari batas kolom pada
garis tabel header — sudah dicek sama persis di kedua contoh dokumen
(Mode Sipil & Mode Normal): [55.6, 97.0, 246.1, 331.0, 388.4, 457.4, 536.5,
600.8, 673.3]. Kalau format dokumen berubah / kolom bergeser, sesuaikan
KOLOM_X di bawah.

Layout per bagian (dikonfirmasi dari kedua dokumen contoh):
    <NAMA BAGIAN>              <- baris judul, satu kata saja
    NRP NAMA KARYAWAN BAGIAN...<- header kolom (diulang di tiap bagian/halaman)
    <baris data karyawan> x N
    TOTAL <5 angka>             <- subtotal bagian ini
    <NAMA BAGIAN berikutnya>
    ...
     GRAND TOTAL <5 angka>      <- baris terakhir dokumen
"OSAMU TAKIMOTO" dkk. dengan bagian nol semua adalah wajar (mis. gaji 0),
bukan baris rusak.
"""

import re
import pdfplumber


def _lapor(progress, persen, pesan=""):
    if progress:
        progress(int(max(0, min(100, persen))), pesan)


def _to_float(s):
    if s is None:
        return None
    s = str(s).strip()
    if s in ("", "-"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    s = s.replace(".", "").replace(",", ".")
    if s in ("", "-"):
        return None
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None


# Batas kolom (x0, dalam satuan poin PDF) hasil pengukuran garis tabel header.
KOLOM_X = [
    (0,   97,   "nrp"),
    (97,  246,  "nama"),
    (246, 331,  "bagian"),
    (331, 388,  "upah_kotor"),
    (388, 457,  "bpjs_kesehatan"),
    (457, 537,  "bpjs_tenaga_kerja"),
    (537, 601,  "pot_lain"),
    (601, 9999, "upah_bersih"),
]

FIELD_ANGKA = ["upah_kotor", "bpjs_kesehatan", "bpjs_tenaga_kerja", "pot_lain", "upah_bersih"]


def _kolom_untuk_x(x0):
    for lo, hi, nama in KOLOM_X:
        if lo <= x0 < hi:
            return nama
    return None


def _kelompokkan_baris(words):
    """Kelompokkan word-level hasil extract_words() jadi baris berdasarkan posisi top."""
    baris = {}
    for w in words:
        key = round(w["top"], 1)
        baris.setdefault(key, []).append(w)
    return [sorted(baris[top], key=lambda w: w["x0"]) for top in sorted(baris.keys())]


def _rakit_kolom(baris_words):
    """Gabungkan kata-kata satu baris jadi dict per kolom (nama/bagian bisa multi-kata)."""
    kolom = {}
    for w in baris_words:
        nama_kolom = _kolom_untuk_x(w["x0"])
        if nama_kolom is None:
            continue
        kolom.setdefault(nama_kolom, []).append(w["text"])
    return {k: " ".join(v) for k, v in kolom.items()}


def _rakit_angka(kolom):
    return {f: _to_float(kolom.get(f)) for f in FIELD_ANGKA}


def parse_pengupahan(pdf_path, progress=None):
    """
    Parse "Laporan Pengupahan Karyawan".
    Return dict:
        {
            "karyawan": [{nrp, nama, bagian, bagian_section, upah_kotor,
                          bpjs_kesehatan, bpjs_tenaga_kerja, pot_lain,
                          upah_bersih}, ...],
            "subtotal_bagian": [{bagian, upah_kotor, ..., upah_bersih}, ...],
            "grand_total": {upah_kotor, ..., upah_bersih} atau None,
        }
    "bagian" = isi kolom BAGIAN pada baris karyawan itu sendiri (mis.
    "MAINTENANCE", "SIPIL UTILITY"). "bagian_section" = judul seksi tempat
    baris TOTAL-nya berada (mis. "ENGINEERING", "SIPIL") — dipakai untuk
    mencocokkan karyawan ke subtotal_bagian yang benar di validasi_internal(),
    karena satu seksi bisa berisi beberapa nilai BAGIAN yang berbeda-beda.
    """
    karyawan = []
    subtotal_bagian = []
    grand_total = None
    bagian_aktif = None

    _lapor(progress, 1, "Membuka PDF laporan pengupahan...")
    with pdfplumber.open(pdf_path) as pdf:
        n_hal = len(pdf.pages) or 1
        for pi, page in enumerate(pdf.pages):
            words = page.extract_words()
            for baris_words in _kelompokkan_baris(words):
                if not baris_words:
                    continue
                kata_pertama = baris_words[0]["text"].strip().upper()

                if kata_pertama in ("LAPORAN", "PERIODE"):
                    continue  # judul dokumen / periode, terulang di tiap halaman

                if kata_pertama == "NRP":
                    continue  # baris header kolom, terulang di tiap bagian

                if kata_pertama == "GRAND":
                    kolom = _rakit_kolom(baris_words)
                    grand_total = _rakit_angka(kolom)
                    continue

                if kata_pertama == "TOTAL":
                    kolom = _rakit_kolom(baris_words)
                    subtotal_bagian.append({"bagian": bagian_aktif, **_rakit_angka(kolom)})
                    continue

                # Baris judul bagian: satu kata/token saja, huruf besar, bukan angka
                if len(baris_words) == 1 and re.fullmatch(r"[A-Z.\-' ]+", kata_pertama):
                    bagian_aktif = kata_pertama
                    continue

                # Baris data karyawan: kolom pertama (NRP) harus berupa angka
                if not re.fullmatch(r"\d{2,10}", kata_pertama):
                    continue  # baris tak dikenal di luar pola, dilewati defensif

                kolom = _rakit_kolom(baris_words)
                karyawan.append({
                    "nrp": kolom.get("nrp"),
                    "nama": kolom.get("nama"),
                    "bagian": kolom.get("bagian"),
                    "bagian_section": bagian_aktif,
                    **_rakit_angka(kolom),
                })

            _lapor(progress, 5 + int(90 * (pi + 1) / n_hal),
                   f"Parse laporan pengupahan halaman {pi + 1}/{n_hal} ({len(karyawan)} karyawan)")

    _lapor(progress, 100, f"Selesai parse laporan pengupahan ({len(karyawan)} karyawan)")
    return {
        "karyawan": karyawan,
        "subtotal_bagian": subtotal_bagian,
        "grand_total": grand_total,
    }


def validasi_internal(hasil_parse, toleransi=1.0):
    """
    Cross-check internal dokumen (bukan terhadap struk gaji): jumlah baris
    karyawan per bagian harus sama dengan subtotal TOTAL bagian tsb, dan
    jumlah semua subtotal harus sama dengan GRAND TOTAL. Berguna untuk
    memastikan parser tidak melewatkan/salah baca baris sebelum data dipakai
    untuk banding ke struk gaji.
    Return: list of str (pesan masalah); list kosong berarti semua cocok.
    """
    masalah = []
    karyawan = hasil_parse.get("karyawan") or []
    subtotal_bagian = hasil_parse.get("subtotal_bagian") or []
    grand_total = hasil_parse.get("grand_total")

    per_bagian = {}
    for k in karyawan:
        per_bagian.setdefault(k.get("bagian_section"), []).append(k)

    for sub in subtotal_bagian:
        bagian = sub.get("bagian")
        anggota = per_bagian.get(bagian, [])
        for f in FIELD_ANGKA:
            jumlah = sum((a.get(f) or 0.0) for a in anggota)
            nilai_sub = sub.get(f) or 0.0
            if abs(jumlah - nilai_sub) > toleransi:
                masalah.append(
                    f"Subtotal bagian '{bagian}' kolom {f} tidak cocok: "
                    f"jumlah baris karyawan = {round(jumlah, 2)}, TOTAL tercetak = {nilai_sub}")

    if grand_total is not None:
        for f in FIELD_ANGKA:
            jumlah = sum((s.get(f) or 0.0) for s in subtotal_bagian)
            nilai_grand = grand_total.get(f) or 0.0
            if abs(jumlah - nilai_grand) > toleransi:
                masalah.append(
                    f"GRAND TOTAL kolom {f} tidak cocok: jumlah semua subtotal bagian = "
                    f"{round(jumlah, 2)}, GRAND TOTAL tercetak = {nilai_grand}")

    return masalah
