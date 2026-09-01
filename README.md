# Pemeriksa Absensi & Struk Gaji - PT. URASE PRIMA

## Cara jalankan
```
pip install -r requirements.txt
python app.py
```

## Alur pemakaian
0. **Pilih Mode Pemrosesan** di UI (wajib sebelum audit):
   - **Mode Normal** — HKP dihitung otomatis (batas group dikurangi izin).
   - **Mode Sipil** — HKP diambil dari dokumen, tidak dihitung ulang.
1. **Buka PDF Absensi** -> aplikasi jalankan aturan sesuai mode + baca HKP/IJIN/DIRUMAHKAN yang dilaporkan dokumen, tampil di tab "Anomali Absensi".
2. **Buka PDF Struk Gaji** -> aplikasi parse data upah per karyawan.
3. **Bandingkan** -> cocokkan kedua dokumen per NRP dengan aturan mode aktif, tampil di tab "Banding Absensi vs Struk Gaji" dengan status:
   - `COCOK` - HKP, IJIN, DIRUMAHKAN, dan jam LEMBURAN sinkron antara absensi dan struk gaji
   - `TIDAK SINKRON` - ada atribut yang beda (detail perbedaan ditampilkan)
   - `HANYA DI ABSENSI` / `HANYA DI STRUK GAJI` - NRP hanya muncul di salah satu dokumen
4. **Ekspor ke Excel** -> simpan semua hasil (3 sheet: Anomali Absensi, Banding, Ringkasan).

Parser sama (`absensi_parser.py`, `gaji_parser.py`) dipakai untuk kedua jenis dokumen
(layout PDF sama). Mode dipilih di UI, **bukan** diinferensikan dari nama group.

## Rumus HKP — dua mode (rules.py / compare.py)
Mode dipilih user; logika Normal dan Sipil tidak saling memanggil.
- **Mode Normal (HKP otomatis):** batas group (21 untuk NS1/NS2/GROUP1/GROUP2, 25
  untuk NS3/GROUP3, dst.) dikurangi hari S2/CI, dikurangi hari DR, dikurangi hari
  izin lain. Banding ke struk gaji memakai hasil hitung ini.
- **Mode Sipil (HKP manual):** HKP **tidak** dihitung otomatis. Nilai HKP dokumen
  absensi dipakai apa adanya lalu dibandingkan ke kolom HKP struk gaji. Aturan 3
  (hitung ulang) tidak dijalankan.

## Status validasi (sudah diuji dengan SEMUA dokumen ASLI yang diberikan)

### Versi otomatis (report_struk_gaji_all.pdf + report_laporan_rekap_absensi_hkp_otomatis__2_.pdf, periode 26/06-25/07/2026)
- `gaji_parser.py` - cocok 100% (66 karyawan, semua subtotal per bagian pas dengan TOTAL tercetak)
- `absensi_parser.py` - 68 karyawan terbaca, HKP/IJIN/DIRUMAHKAN/LEMBURAN sistem cocok 100%
  dengan struk gaji KECUALI SUMARDI (NRP 2312233) - lihat temuan di bawah
- Banding dua dokumen: **65 cocok, 1 tidak sinkron (SUMARDI), 2 hanya di absensi**
  (2 Wakil Direktur berupah 0, wajar tidak ada di struk gaji)
- Aturan 1, 2, 4 - nol false-positive; Aturan 4 otomatis mendeteksi kode '?' milik SUMARDI
- Aturan 3 (rumus HKP Mode Normal) - masih ada ~20 selisih kecil dibanding HKP
  resmi sistem (kemungkinan kode DT/CT butuh penanganan pecahan hari yang belum tertangkap
  rumusnya). Di Mode Normal, banding ke struk gaji memakai HKP hitung otomatis.

### Versi SIPIL / tidak otomatis (report_struk_gaji.pdf + report_laporan_rekap_absensi.pdf, periode 01/07-15/07/2026)
- `gaji_parser.py` - cocok 100% (30 karyawan, total HKP/IJIN/DIRUMAHKAN/LEMBURAN/U.BERSIH
  pas persis dengan GRAND TOTAL tercetak)
- `absensi_parser.py` - 30 karyawan terbaca, semua field sistem terbaca dengan benar
- Banding dua dokumen: **30 cocok, 0 tidak sinkron** - sinkron sempurna
- Aturan 1, 2, 4 - **nol anomali** jika dijalankan di **Mode Sipil** (HKP dokumen
  tidak dihitung ulang; rumus batas 21/25 Mode Normal tidak dijalankan)

## Temuan nyata dari data yang diuji
**SUMARDI (NRP 2312233, versi otomatis)** - HKP di absensi = 24, di struk gaji = 23
(selisih karena kode izin '?' di tanggal 30 tidak ikut dikurangi saat sistem menghitung
HKP-nya). Jumlah IJIN juga tidak sinkron (absensi=1, struk gaji=2). Ini persis anomali
yang sama dengan yang pernah ditemukan sebelumnya - tandanya fitur banding dua dokumen
ini berhasil menangkap masalah nyata, bukan cuma bug parsing.

Contoh hasil lengkap: `hasil_pemeriksaan_contoh.xlsx` (versi otomatis) dan
`hasil_pemeriksaan_sipil.xlsx` (versi SIPIL).
