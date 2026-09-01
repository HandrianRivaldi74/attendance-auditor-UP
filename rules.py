"""
rules.py
Aturan bisnis pemeriksaan absensi & perhitungan HKP (Hari Kerja Pengurang / Normal-Otomatis).

Versi ini mengimplementasikan rumus HKP presisi yang diberikan user lewat
contoh angka konkret, menggantikan versi Kondisi 1/Kondisi 2 sebelumnya yang
terbukti salah (selisih besar vs HKP sistem untuk hampir semua karyawan,
termasuk yang datanya sudah benar 100% seperti Agus Dwi Purwanto / Agus
Kusnadi / Ahmad Jaelani). Rumus baru (tanpa cabang jabatan/Kondisi lagi):

    nilai_tunjangan = batas_group - dirumahkan - izin_masuk_hitungan - izin_tidak_dihitung
    HKP             = nilai_tunjangan - jumlah_hari_pecahan + jumlah_nilai_pecahan

  - dirumahkan            : jumlah hari kode 'DR'
  - izin_masuk_hitungan   : jumlah hari kode izin valid (S1, HR, CT, CD, CK,
                             CL, CM, CN, KK, LR)
  - izin_tidak_dihitung   : jumlah hari kode S2/CI (tercatat di kolom Izin
                             tapi dikecualikan dari rekap "IZIN" resmi —
                             lihat Bagian 3 / hitung_rekap_izin_dan_dirumahkan)
  - jumlah_hari_pecahan   : jumlah hari (bukan libur, bukan kode izin/DR/S2/CI/
                             '?') dengan faktor hari kerja != 1.00 — termasuk
                             hari berkode IK/PC/DT yang datang terlambat/pulang
                             cepat sehingga faktornya pecahan
                  contoh   : batas_group=21, dirumahkan=2, izin_masuk_hitungan=3,
                             izin_tidak_dihitung=0, 1 hari pecahan bernilai 0.88
                             -> nilai_tunjangan = 21-2-3-0 = 16
                             -> HKP = 16 - 1 + 0.88 = 15.88

Satu fungsi inti (_hitung_hkp_lengkap) dipakai bersama oleh hitung_hkp(),
hitung_hkp_resmi(), dan Aturan 3 di jalankan_semua_aturan_banyak() — supaya
kalkulasi HKP Mode Normal, audit banding ke struk gaji (compare.py), dan
pemeriksaan anomali absensi semuanya memakai SATU rumus yang sama.

Asumsi yang masih perlu diverifikasi user (didokumentasikan di kode):
- "izin_tidak_dihitung" dipetakan ke kode S2/CI (TIDAKDIBAYAR_HKP) — sesuai
  namanya, kode ini tercatat di kolom Izin tapi dikecualikan dari rekap
  "IZIN" resmi Bagian 3. Kalau maksud user beda, tolong dikoreksi.
- Cabang jabatan struktural (KABAG/STAFF/DIREKTUR/dst.) dari versi
  sebelumnya DIHAPUS karena tidak muncul lagi di rumus baru — rumus di atas
  sekarang berlaku sama untuk semua karyawan, terlepas dari jabatan.
- Kode izin '?' (data belum steril) tetap tidak dihitung ke komponen mana pun
  (bukan pecahan, bukan izin, bukan dirumahkan) — tetap ditandai Aturan 4.
- Hari libur (Jam Kerja = 00:00/kosong) tidak masuk hitungan sama sekali.
- TIDAKADAJADWAL (dihitung absensi_parser.py) TIDAK dipakai di rumus baru ini
  karena tidak disebutkan user — masih disimpan di record absen untuk
  keperluan lain, tapi tidak mengurangi HKP.
- Batas group (hari_group, 21/25 hari) masih memakai pemetaan sebelumnya
  (nama group mengandung "3" -> 25, mengandung "1"/"2" -> 21).
"""

import re
from dataclasses import dataclass

# ---------- Mode ----------
MODE_NORMAL = "NORMAL"
MODE_SIPIL = "SIPIL"

_MODE_SIPIL_ALIASES = {"SIPIL", "MODE_SIPIL", "CIVIL"}


def normalisasi_mode(mode):
    """Terima string apa pun (atau None) dan kembalikan MODE_NORMAL / MODE_SIPIL."""
    if mode is None:
        return MODE_NORMAL
    m = str(mode).strip().upper()
    if m in _MODE_SIPIL_ALIASES:
        return MODE_SIPIL
    return MODE_NORMAL


def deskripsi_mode(mode):
    """Kembalikan (label, aturan) untuk ditampilkan di UI / ekspor Excel."""
    mode = normalisasi_mode(mode)
    if mode == MODE_SIPIL:
        return (
            "Mode Sipil (HKP dari dokumen)",
            "HKP diambil apa adanya dari dokumen absensi; rumus otomatis "
            "Aturan Operasional Resmi tidak dijalankan.",
        )
    return (
        "Mode Normal (HKP otomatis — Aturan Operasional Resmi)",
        "Nilai tunjangan = batas group − dirumahkan − izin masuk hitungan − izin "
        "tidak dihitung. HKP = nilai tunjangan − jumlah hari pecahan (faktor hari "
        "kerja ≠ 1.00) + jumlah nilai pecahan (total faktor hari-hari itu). "
        "Kode IK/PC/DT masuk mekanisme pecahan, bukan pengurang langsung.",
    )


# ---------- Helper umum ----------
def _to_float(val):
    if val is None:
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _kode_izin(hari_item):
    return (hari_item.get("izin") or "").strip().upper()


def _is_libur(hari_item):
    jk_in = hari_item.get("jam_kerja_in")
    jk_out = hari_item.get("jam_kerja_out")
    return jk_in in (None, "00:00") and jk_out in (None, "00:00")


def _batas_group(group):
    """
    Kembalikan (hari_group, dikenali) berdasarkan nama group absensi.
    dikenali=False berarti group tidak cocok pola yang diketahui -> pakai
    default 21 (ASUMSI, tolong konfirmasi kalau ada group lain).
    """
    if not group:
        return 21, False
    g = str(group).upper()
    if re.search(r"3", g):
        return 25, True
    if re.search(r"[12]", g):
        return 21, True
    return 21, False


# ---------- Simbol Aturan Operasional Resmi ----------
# Bagian 1 - rumus HKP
IZIN_VALID_HKP = {"S1", "HR", "CT", "CD", "CK", "CL", "CM", "CN", "KK", "LR"}   # izin masuk hitungan
DIRUMAHKAN_KODE = {"DR"}
TIDAKDIBAYAR_HKP = {"S2", "CI"}          # "izin tidak dihitung" (lihat asumsi di atas)
KODE_PECAHAN_KHUSUS = {"IK", "PC", "DT"}  # kode yang lazim memicu faktor hari kerja pecahan
KODE_TIDAK_STERIL = {"?"}                 # data belum steril, tidak dihitung

# Bagian 3 - rekap Izin & Dirumahkan (baris 'Izin' dikecualikan simbol ini)
IZIN_REKAP_EXCLUDE = {"CI", "S2", "DT", "PC", "DR", "IK"}


# ---------- Bagian 1: rumus HKP (satu sumber kebenaran) ----------
def _hitung_hkp_lengkap(absen):
    """
    Implementasi rumus HKP presisi (lihat contoh angka di docstring modul):
        nilai_tunjangan = batas_group - dirumahkan - izin_masuk_hitungan - izin_tidak_dihitung
        HKP             = nilai_tunjangan - jumlah_hari_pecahan + jumlah_nilai_pecahan
    Return dict berisi hkp beserta seluruh komponen & metadata untuk audit.
    """
    hari = absen.get("hari") or []
    hari_group, _dikenali = _batas_group(absen.get("group"))

    izin_valid = dirumahkan = tidakdibayar = tidak_steril = 0
    hadir_penuh = 0
    jumlah_hari_pecahan = 0
    jumlah_nilai_pecahan = 0.0

    for h in hari:
        if _is_libur(h):
            continue  # hari libur tidak masuk hitungan sama sekali
        kode = _kode_izin(h)
        faktor = _to_float(h.get("hari_kerja_faktor"))

        if kode in DIRUMAHKAN_KODE:
            dirumahkan += 1
        elif kode in IZIN_VALID_HKP:
            izin_valid += 1
        elif kode in TIDAKDIBAYAR_HKP:
            tidakdibayar += 1
        elif kode in KODE_TIDAK_STERIL:
            tidak_steril += 1
        else:
            # blank/_ , IK/PC/DT, atau kode lain di luar daftar resmi:
            # dianggap hadir, tapi kalau faktor hari kerjanya bukan 1.00 penuh
            # (mis. datang terlambat/pulang cepat) dikoreksi lewat mekanisme
            # hari pecahan, bukan dihitung 1 hari penuh begitu saja.
            if faktor != 1.0:
                jumlah_hari_pecahan += 1
                jumlah_nilai_pecahan += faktor
            else:
                hadir_penuh += 1

    tidak_ada_jadwal = int(absen.get("tidak_ada_jadwal") or 0)

    nilai_tunjangan = hari_group - dirumahkan - izin_valid - tidakdibayar
    hkp = nilai_tunjangan - jumlah_hari_pecahan + jumlah_nilai_pecahan

    sumber = (
        f"nilai tunjangan = batas group ({hari_group}) - dirumahkan ({dirumahkan}) "
        f"- izin masuk hitungan ({izin_valid}) - izin tidak dihitung ({tidakdibayar}) "
        f"= {nilai_tunjangan}; HKP = {nilai_tunjangan} - hari pecahan ({jumlah_hari_pecahan}) "
        f"+ nilai pecahan ({round(jumlah_nilai_pecahan, 2)}) = {round(hkp, 2)}"
    )

    return {
        "hkp": round(hkp, 2),
        "hari_group": hari_group,
        "hadir_penuh": hadir_penuh,
        "izin_valid": izin_valid,
        "dirumahkan": dirumahkan,
        "tidakdibayar": tidakdibayar,
        "tidak_ada_jadwal": tidak_ada_jadwal,
        "tidak_steril": tidak_steril,
        "jumlah_hari_pecahan": jumlah_hari_pecahan,
        "jumlah_nilai_pecahan": round(jumlah_nilai_pecahan, 2),
        "nilai_tunjangan": nilai_tunjangan,
        "ada_tidak_steril": tidak_steril > 0,
        "sumber": sumber,
    }


def hitung_hkp(absen):
    """
    Dipakai gaji_parser.py (Mode Normal, HKP otomatis per karyawan).
    Return: (hkp_hitung: float, hari_group: int, ada_tidak_steril: bool)
    """
    d = _hitung_hkp_lengkap(absen)
    return d["hkp"], d["hari_group"], d["ada_tidak_steril"]


def hitung_hkp_resmi(absen):
    """
    Dipakai compare.py sebagai standar verifikasi HKP terhadap struk gaji
    (poin 5 Aturan Operasional Resmi — menggantikan penjumlahan pecahan
    'hari kerja fisik' yang kaku sebelumnya).
    Return: (hkp_resmi: float, sumber: str)
    """
    d = _hitung_hkp_lengkap(absen)
    return round(d["hkp"], 2), d["sumber"]


# ---------- Bagian 3: rekap Izin & Dirumahkan ----------
def hitung_rekap_izin_dan_dirumahkan(absen):
    """
    IZIN = total baris 'Izin' KECUALI simbol {CI, S2, DT, PC, DR, IK}.
    DIRUMAHKAN = khusus simbol {DR}.
    Kode '?' (belum steril) tidak dihitung (lihat Aturan 4).
    Fallback ke ijin_sistem / dirumahkan_sistem bila rincian harian kosong.
    Return: (izin_resmi: int, dirumahkan_resmi: int)
    """
    hari = absen.get("hari") or []
    if hari:
        izin = 0
        dirumahkan = 0
        for h in hari:
            kode = _kode_izin(h)
            if kode in ("", "_") or kode in KODE_TIDAK_STERIL:
                continue
            if kode in DIRUMAHKAN_KODE:
                dirumahkan += 1
                continue
            if kode in IZIN_REKAP_EXCLUDE:
                continue
            izin += 1
        return izin, dirumahkan
    return int(_to_float(absen.get("ijin_sistem"))), int(_to_float(absen.get("dirumahkan_sistem")))


# ---------- Bagian 4: poin tunjangan & transportasi ----------
def hitung_poin_tunjangan(absen):
    """
    Pengali TUNJ MASA KERJA / TUNJ JABATAN / TUNJ KHUSUS / UPAH TRANSPORT:
    jumlah hari dengan faktor hari kerja != 0. Nilai ini juga sudah dihitung
    langsung oleh absensi_parser.py sebagai field 'poin_tunjangan'; fungsi ini
    dipakai sebagai fallback/cross-check kalau field itu belum ada.
    """
    if "poin_tunjangan" in absen:
        return int(absen["poin_tunjangan"])
    hari = absen.get("hari") or []
    return sum(1 for h in hari if _to_float(h.get("hari_kerja_faktor")) != 0)


# ---------- Aturan pemeriksaan anomali absensi ----------
@dataclass
class AnomaliAbsensi:
    nrp: str
    nama: str
    aturan: str
    detail: str
    tanggal: str = "-"


def _lapor(progress, persen, pesan=""):
    if progress:
        progress(int(max(0, min(100, persen))), pesan)


def jalankan_semua_aturan_banyak(data_absensi, mode=None, progress=None):
    """
    Jalankan semua aturan pemeriksaan anomali untuk setiap karyawan.

    Aturan 1 - Kelengkapan Data: rincian harian atau nama tidak terbaca.
    Aturan 2 - Konsistensi Jam Kerja: ada jam kerja tapi faktor hari kerja 0
               (atau sebaliknya).
    Aturan 3 - Rumus HKP Resmi (Bagian 1): HKP hasil _hitung_hkp_lengkap()
               dibandingkan ke HKP sistem (tunj_masa_kerja). Kalau data
               belum steril (kode '?') tidak diperiksa di sini (ditangani
               Aturan 4). Kalau cocok, data otomatis dianggap VALID — tidak
               ditandai anomali sama sekali (sesuai poin 5 Aturan
               Operasional Resmi).
    Aturan 4 - Kode Izin Belum Steril: kode '?' pada rincian harian.
    """
    mode = normalisasi_mode(mode)
    hasil = []
    n = len(data_absensi) or 1

    for i, absen in enumerate(data_absensi):
        nrp = absen.get("id")
        nama = absen.get("nama")
        hari = absen.get("hari") or []

        # Aturan 1
        if not hari:
            hasil.append(AnomaliAbsensi(nrp, nama, "Aturan 1 - Kelengkapan Data",
                                         "Rincian harian tidak terbaca dari PDF"))
        if not nama:
            hasil.append(AnomaliAbsensi(nrp, nama, "Aturan 1 - Kelengkapan Data",
                                         "Nama karyawan kosong/tidak terbaca"))

        # Aturan 2 (hari dengan kode izin/dirumahkan/tidakdibayar/? sengaja
        # dikecualikan: faktor hari kerja = 0 pada hari itu memang wajar,
        # bukan inkonsistensi)
        for h in hari:
            kode = _kode_izin(h)
            if kode not in ("", "_"):
                continue
            jk_in = h.get("jam_kerja_in")
            jk_out = h.get("jam_kerja_out")
            faktor = _to_float(h.get("hari_kerja_faktor"))
            ada_jam = jk_in not in (None, "00:00", "") or jk_out not in (None, "00:00", "")
            if ada_jam and faktor == 0.0:
                hasil.append(AnomaliAbsensi(
                    nrp, nama, "Aturan 2 - Konsistensi Jam Kerja",
                    f"Ada jam kerja ({jk_in}-{jk_out}) tapi faktor hari kerja = 0 (tanpa kode izin)",
                    str(h.get("tanggal", "-"))))
            elif not ada_jam and faktor > 0.0:
                hasil.append(AnomaliAbsensi(
                    nrp, nama, "Aturan 2 - Konsistensi Jam Kerja",
                    f"Faktor hari kerja = {faktor} tapi tidak ada jam kerja tercatat",
                    str(h.get("tanggal", "-"))))

        # Aturan 3 (hanya Mode Normal, rumus resmi Bagian 1)
        if mode == MODE_NORMAL:
            d = _hitung_hkp_lengkap(absen)
            hkp_sistem = absen.get("tunj_masa_kerja")
            if not d["ada_tidak_steril"] and hkp_sistem is not None:
                if abs(d["hkp"] - _to_float(hkp_sistem)) > 0.01:
                    hasil.append(AnomaliAbsensi(
                        nrp, nama, "Aturan 3 - Rumus HKP Resmi",
                        f"HKP hitung resmi = {d['hkp']} ({d['sumber']}), "
                        f"HKP sistem = {hkp_sistem}"))
                # Jika cocok: tidak ditambahkan ke hasil sama sekali -> VALID.

        # Aturan 4
        for h in hari:
            if _kode_izin(h) in KODE_TIDAK_STERIL:
                hasil.append(AnomaliAbsensi(
                    nrp, nama, "Aturan 4 - Kode Izin Belum Steril",
                    "Kode izin '?' belum diselesaikan/steril di sistem",
                    str(h.get("tanggal", "-"))))

        if progress and (i % 3 == 0 or i + 1 == n):
            _lapor(progress, 100 * (i + 1) / n, f"Memeriksa anomali {i + 1}/{n}")

    return hasil
