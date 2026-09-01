import re
import pdfplumber

def _lapor(progress, persen, pesan=""):
    if progress:
        progress(int(max(0, min(100, persen))), pesan)

LABELS_RINGKASAN = ["HKP", "HKL", "LL", "IZIN", "DIRUMAHKAN"]

def _to_float(s):
    if s is None:
        return None
    s = str(s).strip().strip(",")
    if s in ("", "-"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").strip()
    s = s.replace(".", "").replace(",", ".")
    if s in ("", "-"):
        return None
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None

def _cluster_text(words, x_min, x_max, top_min, top_max):
    sel = [w for w in words if x_min <= w["x0"] <= x_max and top_min <= w["top"] <= top_max]
    sel.sort(key=lambda w: (round(w["top"], 1), w["x0"]))
    lines = {}
    for w in sel:
        key = round(w["top"], 0)
        lines.setdefault(key, []).append(w["text"])
    out = []
    for k in sorted(lines):
        out.append(" ".join(lines[k]))
    return " ".join(out).strip()

def _to_float_faktor(s):
    if s is None:
        return None
    s = str(s).strip()
    if s in ("", "-"):
        return None
    if s.startswith(","):
        s = "0" + s.replace(",", ".", 1)
    try:
        return float(s)
    except ValueError:
        return _to_float(s)

def _grid_ke_hari(tabel):
    """
    Ubah grid tabel per-tanggal jadi list dict per hari.
    Kolom yang benar-benar kosong (Jam Kerja DAN Mesin Absen sama-sama kosong)
    tidak dimasukkan ke `hari`, melainkan dihitung terpisah sebagai
    TIDAKADAJADWAL (Aturan Operasional Resmi Bagian 1) — dikembalikan sebagai
    nilai kedua supaya bisa dipakai rules.hitung_hkp().
    """
    hari = []
    tidak_ada_jadwal = 0
    if not tabel or len(tabel) < 4:
        return hari, tidak_ada_jadwal
    baris_jam, baris_mesin, baris_lembur_izin, baris_hk = tabel[0], tabel[1], tabel[2], tabel[3]
    n = len(baris_jam)
    for i in range(n):
        jam = (baris_jam[i] or "").strip()
        if jam == "" and (baris_mesin[i] or "").strip() == "":
            tidak_ada_jadwal += 1
            continue
        
        jam_parts = jam.split("\n")
        jk_in = jam_parts[0] if len(jam_parts) > 0 else None
        jk_out = jam_parts[1] if len(jam_parts) > 1 else None

        mesin = (baris_mesin[i] or "").strip().split("\n")
        m_in = mesin[0] if len(mesin) > 0 else None
        m_out = mesin[1] if len(mesin) > 1 else None

        li = (baris_lembur_izin[i] or "").strip().split("\n")
        lembur_val = _to_float(li[0]) if len(li) > 0 else None
        izin = li[1] if len(li) > 1 else "_"
        hk = _to_float_faktor((baris_hk[i] or "").strip())

        hari.append({
            "tanggal": i + 1,
            "jam_kerja_in": jk_in, "jam_kerja_out": jk_out,
            "mesin_in": m_in, "mesin_out": m_out,
            "izin": izin if izin else "_",
            "hari_kerja_faktor": hk if hk is not None else 0.0,
            "lembur_hkl_hari": lembur_val,
        })
    return hari, tidak_ada_jadwal


def _hitung_poin_tunjangan(hari_list):
    """
    Aturan Operasional Resmi Bagian 4 (Tunjangan & Transportasi):
    hari dengan faktor hari kerja != 0 dihitung 1, selain itu 0. Total poin ini
    dipakai sebagai pengali TUNJ MASA KERJA / TUNJ JABATAN / TUNJ KHUSUS /
    UPAH TRANSPORT.
    """
    return sum(1 for h in hari_list if (h.get("hari_kerja_faktor") or 0.0) != 0)

def _hitung_hkl_ll_resmi(hari_list, ll_sistem_default=0, hkl_sistem_default=0):
    total_hkl = 0.0
    total_ll = 0.0
    ada_rincian_lembur = False

    for h in hari_list:
        val = h.get("lembur_hkl_hari")
        if val is None:
            continue
        
        ada_rincian_lembur = True
        jk_in = h.get("jam_kerja_in")
        jk_out = h.get("jam_kerja_out")
        
        is_libur = (jk_in in (None, "00:00")) and (jk_out in (None, "00:00"))
        
        if is_libur:
            total_ll += val
        else:
            total_hkl += val

    if not ada_rincian_lembur:
        return hkl_sistem_default or 0, ll_sistem_default or 0
    return total_hkl, total_ll

def parse_absensi(pdf_path, progress=None):
    out = []
    _lapor(progress, 1, "Membuka PDF absensi...")
    with pdfplumber.open(pdf_path) as pdf:
        n_hal = len(pdf.pages) or 1
        for i, page in enumerate(pdf.pages):
            words = page.extract_words()
            tables = page.extract_tables()

            anchors = []
            for w in words:
                if 88 <= w["x0"] <= 108 and re.fullmatch(r"\d{3,7}", w["text"]):
                    anchors.append((w["top"], w["text"]))
            anchors.sort(key=lambda a: a[0])

            for idx, (T, nrp) in enumerate(anchors):
                nama = _cluster_text(words, 88, 180, T + 12, T + 49)
                group = _cluster_text(words, 18, 88, T + 12, T + 49)
                bagian = _cluster_text(words, 18, 88, T + 50, T + 62)
                posisi = _cluster_text(words, 88, 180, T + 50, T + 62)

                ringkasan = {}
                for w in words:
                    if 930 <= w["x0"] <= 996 and w["text"] in LABELS_RINGKASAN and (T - 32) <= w["top"] <= (T + 36):
                        label = w["text"]
                        nilai_words = [v for v in words if 1015 <= v["x0"] <= 1042 and abs(v["top"] - w["top"]) < 1.5]
                        if nilai_words:
                            ringkasan[label] = _to_float(nilai_words[0]["text"])

                tabel = tables[idx] if idx < len(tables) else None
                hari, tidak_ada_jadwal = _grid_ke_hari(tabel)
                
                hkl_calc, ll_calc = _hitung_hkl_ll_resmi(
                    hari, 
                    ll_sistem_default=ringkasan.get("LL"), 
                    hkl_sistem_default=ringkasan.get("HKL")
                )
                poin_tunjangan = _hitung_poin_tunjangan(hari)

                out.append({
                    "id": nrp,
                    "nama": nama,
                    "group": group,
                    "bagian": bagian,
                    "posisi": posisi,
                    "hari": hari,
                    "tidak_ada_jadwal": tidak_ada_jadwal,
                    "poin_tunjangan": poin_tunjangan,
                    "lembur_libur": ll_calc,
                    "lembur_hkl": hkl_calc,
                    "tunj_masa_kerja": ringkasan.get("HKP"),
                    "ijin_sistem": ringkasan.get("IZIN"),
                    "dirumahkan_sistem": ringkasan.get("DIRUMAHKAN"),
                })
            _lapor(progress, 5 + int(90 * (i + 1) / n_hal), f"Parse absensi halaman {i + 1}/{n_hal}")
    _lapor(progress, 100, f"Selesai parse absensi ({len(out)} karyawan)")
    return out
