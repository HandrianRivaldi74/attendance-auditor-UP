from dataclasses import dataclass
import re
import rules

@dataclass
class HasilBanding:
    nrp: str
    nama_absensi: str
    nama_gaji: str
    status: str
    detail: list
    mode: str = ""
    sumber_hkp: str = ""


def bandingkan(data_absensi, data_gaji, mode=rules.MODE_NORMAL, toleransi=0.01, progress=None):
    mode = rules.normalisasi_mode(mode)
    sumber = (
        "Dokumen Langsung (Mode Sipil — tanpa rumus batas group)"
        if mode == rules.MODE_SIPIL else
        "Kalkulasi Operasional Resmi (Mode Normal)"
    )
    
    by_absen = {k["id"]: k for k in data_absensi}
    by_gaji = {k["nrp"]: k for k in data_gaji}
    semua_nrp = sorted(set(by_absen) | set(by_gaji))
    
    hasil = []
    n = len(semua_nrp) or 1

    for i, nrp in enumerate(semua_nrp):
        absen = by_absen.get(nrp)
        gaji = by_gaji.get(nrp)

        if absen and not gaji:
            hasil.append(HasilBanding(nrp, absen["nama"], None, "HANYA DI ABSENSI", ["Tidak di struk gaji"], mode=mode, sumber_hkp=sumber))
        elif gaji and not absen:
            hasil.append(HasilBanding(nrp, None, gaji["nama"], "HANYA DI STRUK GAJI", ["Tidak di data absensi"], mode=mode, sumber_hkp=sumber))
        else:
            detail = []
            
            # 1. Cek Nama
            if absen["nama"].strip().upper() != (gaji["nama"] or "").strip().upper():
                detail.append(f"Nama beda: '{absen['nama']}' vs '{gaji['nama']}'")

            if mode == rules.MODE_SIPIL:
                # Mode Sipil: TIDAK memakai rumus Aturan Operasional Resmi (batas
                # group, izin_valid/dirumahkan/pecahan dst). Bandingkan langsung
                # nilai per atribut yang tercetak di dokumen absensi vs struk gaji.
                hkp_absensi_dok = absen.get("tunj_masa_kerja")
                hkp_gaji = gaji.get("hkp")
                if hkp_absensi_dok is not None and hkp_gaji is not None \
                        and abs(round(hkp_absensi_dok, 2) - round(hkp_gaji, 2)) > toleransi:
                    detail.append(
                        f"HKP tidak cocok (dokumen langsung): Absensi = {hkp_absensi_dok}, "
                        f"Struk Gaji = {hkp_gaji}")

                izin_absensi_dok = absen.get("ijin_sistem")
                if izin_absensi_dok is not None and gaji.get("ijin") is not None \
                        and int(izin_absensi_dok) != int(gaji.get("ijin")):
                    detail.append(
                        f"IJIN tidak sinkron (dokumen langsung): Absensi = {izin_absensi_dok}, "
                        f"Struk Gaji = {gaji.get('ijin')}")

                dirumahkan_absensi_dok = absen.get("dirumahkan_sistem")
                if dirumahkan_absensi_dok is not None and gaji.get("dirumahkan") is not None \
                        and int(dirumahkan_absensi_dok) != int(gaji.get("dirumahkan")):
                    detail.append(
                        f"DIRUMAHKAN tidak sinkron (dokumen langsung): Absensi = {dirumahkan_absensi_dok}, "
                        f"Struk Gaji = {gaji.get('dirumahkan')}")
            else:
                # Mode Normal: pakai rumus Aturan Operasional Resmi (Bagian 1 & 3)
                # sebagai standar tunggal verifikasi validitas data (Bagian 5).
                hkp_resmi, sumber_hkp_resmi = rules.hitung_hkp_resmi(absen)
                hkp_gaji = gaji.get("hkp")
                if hkp_gaji is not None and abs(round(hkp_resmi, 2) - round(hkp_gaji, 2)) > toleransi:
                    detail.append(
                        f"HKP tidak cocok: Hitung Resmi = {hkp_resmi} ({sumber_hkp_resmi}), "
                        f"Struk Gaji = {hkp_gaji}")

                izin_resmi, dirumahkan_resmi = rules.hitung_rekap_izin_dan_dirumahkan(absen)
                if gaji.get("ijin") is not None and int(izin_resmi) != int(gaji.get("ijin")):
                    detail.append(f"IJIN tidak sinkron: Hitung Resmi = {izin_resmi}, Struk Gaji = {gaji.get('ijin')}")

                if gaji.get("dirumahkan") is not None and int(dirumahkan_resmi) != int(gaji.get("dirumahkan")):
                    detail.append(f"DIRUMAHKAN tidak sinkron: Hitung Resmi = {dirumahkan_resmi}, Struk Gaji = {gaji.get('dirumahkan')}")

            # 4. Cek HKL (sama untuk kedua mode — nilai lembur sudah dibaca
            #    langsung dari kolom sistem, tidak melalui rumus batas group)
            lembur_absensi = absen.get("lembur_hkl")
            lembur_gaji = gaji.get("hkl")
            if lembur_absensi is not None and lembur_gaji is not None and abs(round(lembur_absensi, 2) - round(lembur_gaji, 2)) > toleransi:
                detail.append(f"HKL tidak sinkron: Absensi = {lembur_absensi}, Struk Gaji = {lembur_gaji}")

            status = "COCOK" if not detail else "TIDAK SINKRON"
            hasil.append(HasilBanding(nrp, absen["nama"], gaji["nama"], status, detail, mode=mode, sumber_hkp=sumber))

        if progress and (i % 3 == 0 or i + 1 == n):
            progress(int(100 * (i + 1) / n), f"Membandingkan {i + 1}/{n}")

    return hasil


def ringkasan(hasil_banding):
    return {
        "total": len(hasil_banding),
        "cocok": sum(1 for h in hasil_banding if h.status == "COCOK"),
        "tidak_sinkron": sum(1 for h in hasil_banding if h.status == "TIDAK SINKRON"),
        "hanya_di_absensi": sum(1 for h in hasil_banding if h.status == "HANYA DI ABSENSI"),
        "hanya_di_struk_gaji": sum(1 for h in hasil_banding if h.status == "HANYA DI STRUK GAJI"),
        "mode": hasil_banding[0].mode if hasil_banding else "",
        "sumber_hkp": hasil_banding[0].sumber_hkp if hasil_banding else "",
    }


# ---------- Banding Laporan Pengupahan (total upah) vs Struk Gaji ----------
# Kelima kolom finansial di "Laporan Pengupahan Karyawan" (pengupahan_parser.py)
# dicocokkan langsung ke kolom struk gaji hasil gaji_parser.py — bukan hasil
# rumus HKP apa pun (dokumen ini murni rekap upah, tidak ada kolom HKP/izin).
_PEMETAAN_KOLOM_PENGUPAHAN = {
    "upah_kotor": "upah_kotor",
    "bpjs_kesehatan": "bpjs_ks",
    "bpjs_tenaga_kerja": "bpjs_kt",
    "pot_lain": "pot_lain",
    "upah_bersih": "u_bersih",
}


@dataclass
class HasilBandingPengupahan:
    nrp: str
    nama_pengupahan: str
    nama_gaji: str
    status: str
    detail: list


def bandingkan_pengupahan(data_pengupahan, data_gaji, toleransi=1.0, progress=None):
    """
    Cocokkan "Laporan Pengupahan Karyawan" (pengupahan_parser.parse_pengupahan()
    -> pakai list di key "karyawan") terhadap struk gaji (gaji_parser.py) per NRP,
    murni banding nilai dokumen ke dokumen — tidak melibatkan rumus HKP/Aturan
    Operasional Resmi sama sekali.
    Return: list HasilBandingPengupahan.
    """
    by_pengupahan = {k["nrp"]: k for k in data_pengupahan}
    by_gaji = {k["nrp"]: k for k in data_gaji}
    semua_nrp = sorted(set(by_pengupahan) | set(by_gaji))

    hasil = []
    n = len(semua_nrp) or 1

    for i, nrp in enumerate(semua_nrp):
        pu = by_pengupahan.get(nrp)
        gj = by_gaji.get(nrp)

        if pu and not gj:
            hasil.append(HasilBandingPengupahan(nrp, pu["nama"], None, "HANYA DI LAPORAN PENGUPAHAN", ["Tidak ada di struk gaji"]))
        elif gj and not pu:
            hasil.append(HasilBandingPengupahan(nrp, None, gj["nama"], "HANYA DI STRUK GAJI", ["Tidak ada di laporan pengupahan"]))
        else:
            detail = []
            if (pu["nama"] or "").strip().upper() != (gj["nama"] or "").strip().upper():
                detail.append(f"Nama beda: '{pu['nama']}' vs '{gj['nama']}'")

            for kolom_pu, kolom_gj in _PEMETAAN_KOLOM_PENGUPAHAN.items():
                nilai_pu = pu.get(kolom_pu)
                nilai_gj = gj.get(kolom_gj)
                # Bandingkan BESARAN (abs), bukan nilai mentah: struk gaji
                # (gaji_parser.py) menyimpan potongan seperti BPJS sebagai
                # angka negatif (konvensi "(50.529)" di PDF -> -50529), sedang
                # Laporan Pengupahan mencetaknya sebagai angka positif murni.
                # Beda tanda ini bukan anomali data — cuma beda konvensi
                # penulisan antar dokumen.
                if nilai_pu is not None and nilai_gj is not None \
                        and abs(abs(round(nilai_pu, 2)) - abs(round(nilai_gj, 2))) > toleransi:
                    detail.append(
                        f"{kolom_pu} tidak cocok: Laporan Pengupahan = {nilai_pu}, "
                        f"Struk Gaji = {nilai_gj}")

            status = "COCOK" if not detail else "TIDAK SINKRON"
            hasil.append(HasilBandingPengupahan(nrp, pu["nama"], gj["nama"], status, detail))

        if progress and (i % 3 == 0 or i + 1 == n):
            progress(int(100 * (i + 1) / n), f"Membandingkan laporan pengupahan {i + 1}/{n}")

    return hasil


def ringkasan_pengupahan(hasil_banding):
    return {
        "total": len(hasil_banding),
        "cocok": sum(1 for h in hasil_banding if h.status == "COCOK"),
        "tidak_sinkron": sum(1 for h in hasil_banding if h.status == "TIDAK SINKRON"),
        "hanya_di_pengupahan": sum(1 for h in hasil_banding if h.status == "HANYA DI LAPORAN PENGUPAHAN"),
        "hanya_di_struk_gaji": sum(1 for h in hasil_banding if h.status == "HANYA DI STRUK GAJI"),
    }


# ---------- Banding Transfer Bank vs Struk Gaji ----------
# File transfer bank (bank_transfer_parser.py) TIDAK punya NRP terisi (kolom
# NIP selalu kosong di template bank), jadi pencocokan di sini WAJIB lewat
# NAMA (dinormalisasi), bukan NRP seperti fungsi banding lainnya.
def _normalisasi_nama(nama):
    if not nama:
        return ""
    return re.sub(r"\s+", " ", str(nama).strip().upper())


@dataclass
class HasilBandingTransferBank:
    nama: str
    jumlah_bank: float
    jumlah_gaji: float
    status: str
    detail: list


def bandingkan_transfer_bank(data_bank, data_gaji, toleransi=1.0, progress=None):
    """
    Cocokkan transaksi transfer bank (bank_transfer_parser.parse_transfer_bank())
    terhadap struk gaji (gaji_parser.py) lewat NAMA (bukan NRP — lihat catatan
    modul bank_transfer_parser.py). Amount transfer dibandingkan ke u_bersih
    (UPAH BERSIH) struk gaji.

    Penanganan kasus khusus:
    - Karyawan dengan UPAH BERSIH 0/None di struk gaji WAJAR tidak muncul di
      transfer bank (tidak ada yang ditransfer) -> status "TIDAK DITRANSFER
      (UPAH 0 - WAJAR)", bukan anomali.
    - Nama yang muncul lebih dari sekali di salah satu sisi (nama kembar/
      ambigu) TIDAK ditebak pasangannya secara otomatis -> status "AMBIGU
      (NAMA KEMBAR)" supaya diperiksa manual, karena mencocokkan gaji orang
      yang salah adalah risiko nyata.
    Return: list HasilBandingTransferBank.
    """
    by_bank = {}
    for b in data_bank:
        by_bank.setdefault(_normalisasi_nama(b.get("nama")), []).append(b)

    by_gaji = {}
    for g in data_gaji:
        by_gaji.setdefault(_normalisasi_nama(g.get("nama")), []).append(g)

    semua_nama = sorted(set(by_bank) | set(by_gaji))
    hasil = []
    n = len(semua_nama) or 1

    for i, nama_key in enumerate(semua_nama):
        entri_bank = by_bank.get(nama_key, [])
        entri_gaji = by_gaji.get(nama_key, [])

        if len(entri_bank) > 1 or len(entri_gaji) > 1:
            hasil.append(HasilBandingTransferBank(
                nama_key,
                sum((b.get("jumlah") or 0) for b in entri_bank) if entri_bank else None,
                sum((g.get("u_bersih") or 0) for g in entri_gaji) if entri_gaji else None,
                "AMBIGU (NAMA KEMBAR)",
                [f"{len(entri_bank)} transaksi bank vs {len(entri_gaji)} baris struk gaji "
                 f"dengan nama sama — tidak dicocokkan otomatis, periksa manual"]))
            continue

        if entri_bank and not entri_gaji:
            hasil.append(HasilBandingTransferBank(
                nama_key, entri_bank[0].get("jumlah"), None,
                "HANYA DI TRANSFER BANK",
                ["Tidak ada nama ini di struk gaji — periksa kemungkinan transfer ke penerima keliru"]))
            continue

        if entri_gaji and not entri_bank:
            u_bersih = entri_gaji[0].get("u_bersih")
            if u_bersih is None or abs(u_bersih) < toleransi:
                hasil.append(HasilBandingTransferBank(
                    nama_key, None, u_bersih, "TIDAK DITRANSFER (UPAH 0 - WAJAR)",
                    ["Upah bersih 0/kosong di struk gaji — wajar tidak ada transfer bank"]))
            else:
                hasil.append(HasilBandingTransferBank(
                    nama_key, None, u_bersih, "HANYA DI STRUK GAJI (BELUM ADA TRANSFER BANK)",
                    [f"Upah bersih struk gaji = {u_bersih}, tapi tidak ditemukan transaksi transfer bank"]))
            continue

        jumlah_bank = entri_bank[0].get("jumlah")
        jumlah_gaji = entri_gaji[0].get("u_bersih")
        detail = []
        if jumlah_bank is not None and jumlah_gaji is not None \
                and abs(round(jumlah_bank, 2) - round(jumlah_gaji, 2)) > toleransi:
            detail.append(f"Jumlah tidak cocok: Transfer Bank = {jumlah_bank}, Struk Gaji (Upah Bersih) = {jumlah_gaji}")
        status = "COCOK" if not detail else "TIDAK SINKRON"
        hasil.append(HasilBandingTransferBank(nama_key, jumlah_bank, jumlah_gaji, status, detail))

        if progress and (i % 5 == 0 or i + 1 == n):
            progress(int(100 * (i + 1) / n), f"Membandingkan transfer bank {i + 1}/{n}")

    return hasil


def ringkasan_transfer_bank(hasil_banding):
    return {
        "total": len(hasil_banding),
        "cocok": sum(1 for h in hasil_banding if h.status == "COCOK"),
        "tidak_sinkron": sum(1 for h in hasil_banding if h.status == "TIDAK SINKRON"),
        "tidak_ditransfer_wajar": sum(1 for h in hasil_banding if h.status == "TIDAK DITRANSFER (UPAH 0 - WAJAR)"),
        "hanya_di_bank": sum(1 for h in hasil_banding if h.status == "HANYA DI TRANSFER BANK"),
        "hanya_di_struk_gaji": sum(1 for h in hasil_banding if h.status == "HANYA DI STRUK GAJI (BELUM ADA TRANSFER BANK)"),
        "ambigu": sum(1 for h in hasil_banding if h.status == "AMBIGU (NAMA KEMBAR)"),
    }
