from dataclasses import dataclass
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
                if nilai_pu is not None and nilai_gj is not None \
                        and abs(round(nilai_pu, 2) - round(nilai_gj, 2)) > toleransi:
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
