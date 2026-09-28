# CLAUDE.md — Aplikasi Perhitungan Fuzzy DEMATEL (Thesis Risiko Rantai Pasok Geopolitik)

> File ini adalah instruksi lengkap untuk Claude Code. Baca seluruhnya sebelum menulis kode.
> Bahasa antarmuka aplikasi: **Bahasa Indonesia**. Nama variabel/fungsi di kode: bahasa Inggris.

---

## 1. Tujuan

Buat aplikasi **satu file HTML** bernama `fuzzy_dematel.html` yang:

1. Dibuka cukup dengan double-click di browser (Chrome/Edge), tanpa install Python/Node untuk pengguna.
2. Pengguna meng-upload (drag & drop atau tombol pilih file) file Excel `Responden_final_thesis.xlsx`.
3. Aplikasi langsung menghitung **Fuzzy DEMATEL** lengkap dan menampilkan semua hasil antara.
4. Pengguna bisa mengunduh hasil sebagai **Excel (.xlsx)** dan diagram **Impact Relation Map (.png)**.

Pengguna adalah mahasiswa S1 (bukan programmer). Utamakan: mudah dipakai, pesan error jelas dalam Bahasa Indonesia, dan hasil yang bisa langsung dimasukkan ke Bab 4 skripsi.

---

## 2. Konteks penelitian

- Judul: *Geopolitical Supply Chain Risk in Indonesian Import-Dependent FMCG Firms*.
- Metode: Fuzzy DEMATEL untuk memetakan hubungan sebab-akibat antar **20 faktor risiko (R1–R20)**.
- Jumlah pakar/responden saat ini: **7** (tapi aplikasi harus menerima berapa pun jumlah sheet responden ≥ 1).

Daftar faktor (kode → nama). Nama di file data kadang punya spasi di akhir (mis. `"Trade Route Disruption "`), jadi **selalu `trim()`**.

| Kode | Nama faktor | Causal tier (hipotesis penelitian) |
|---|---|---|
| R1 | Geopolitical Conflict and Instability | Tier 1 — Trigger |
| R2 | Trade Route Disruption | Tier 2 — Transmission |
| R3 | Economic Sanctions and Export Bans | Tier 1 — Trigger |
| R4 | Import Dependency Exposure | Amplifying Moderator |
| R5 | Supplier Concentration in High-Risk Regions | Amplifying Moderator |
| R6 | Energy Price Shock from Conflict | Tier 2 — Transmission |
| R7 | Raw Material Price Volatility | Tier 2 — Transmission |
| R8 | Foreign Exchange Rate Instability | Tier 2 — Transmission |
| R9 | Port Congestion and Maritime Disruption | Tier 2 — Transmission |
| R10 | Cross-Border Regulatory Uncertainty | Tier 2 — Transmission |
| R11 | Lead Time Escalation | Tier 3 — Outcome |
| R12 | Supplier Financial Distress under Conflict | Tier 2 — Transmission |
| R13 | Logistics Capacity Shortage | Tier 2 — Transmission |
| R14 | Conflict-Driven Demand Shock | Tier 3 — Outcome |
| R15 | Information Asymmetry in Conflict Zones | Cross-tier Amplifier |
| R16 | Strategic Export Restrictions by Supplier Nations | Tier 1 — Trigger |
| R17 | Alternative Route Cost Escalation | Tier 2 — Transmission |
| R18 | Inventory Stockout from Geopolitical Supply Shock | Tier 3 — Outcome |
| R19 | Political Risk in Supplier Countries | Tier 1 — Trigger |
| R20 | Coordination Failure under Geopolitical Pressure | Tier 3 — Outcome |

Kolom "Causal tier" hanya ditampilkan sebagai pembanding di tabel hasil. **Jangan** dipakai dalam perhitungan.

---

## 3. Format file input (`Responden_final_thesis.xlsx`)

Sudah dicek langsung dari file aslinya:

- Satu sheet per responden, nama sheet: `responden 1`, `responden 2`, … `responden 7`. Proses **semua sheet** yang sel A3-nya berisi `R1` (abaikan sheet lain, misalnya sheet ringkasan kalau nanti ditambahkan).
- Ukuran tiap sheet: 22 baris × 22 kolom.
  - Baris 1 (index 0): sel C1:V1 = kode faktor `R1`…`R20`.
  - Baris 2 (index 1): B2 = `"Affected By →"`, C2:V2 = nama faktor.
  - Baris 3–22 (index 2–21): kolom A = kode faktor, kolom B = nama faktor, kolom C:V = skor.
- Skor berupa angka `0, 1, 2, 3, 4` (tersimpan sebagai float, mis. `2.0`). Tidak ada formula.
- Diagonal (faktor yang sama) diisi `0`.

**Temuan data yang harus ditangani:** di sheet `responden 7`, sel diagonal R11→R11 berisi `4` (seharusnya kosong/0). Aplikasi harus **memaksa diagonal = 0** dan menampilkan **peringatan kuning** yang menyebutkan sheet & sel mana yang dikoreksi. Jangan menghentikan proses karena ini.

### Validasi (hentikan proses + pesan error merah bila gagal)
- Tidak ada sheet responden yang valid.
- Kode baris/kolom tidak berurutan `R1`…`R20`, atau urutan kode berbeda antar sheet.
- Sel non-diagonal kosong, bukan angka, atau di luar {0,1,2,3,4}. Pesan error harus menyebut **nama sheet + alamat sel Excel** (mis. `responden 3, sel F12`).

Buat parser fleksibel terhadap jumlah faktor n (ambil dari baris 1), jangan hard-code 20, supaya tetap jalan kalau jumlah faktor berubah.

---

## 4. ⚠️ Arah matriks (WAJIB dibuat bisa dipilih)

Ada ambiguitas arah pengaruh di kuesioner:

- Sheet *Instruction* kuesioner berbunyi: *"rate the degree to which the **COLUMN** factor influences the **ROW** factor"* → sel (baris i, kolom j) = pengaruh **j terhadap i**.
- Tetapi header kuesioner `"Affected By →"` di atas kolom bisa dibaca sebaliknya: sel (baris i, kolom j) = pengaruh **i terhadap j** (konvensi standar DEMATEL, z_ij = pengaruh i → j).

Kedua arah memberi hasil cause/effect yang **berbeda total**. Karena itu:

1. Sediakan pilihan di panel Pengaturan:
   - `row_influences_column` → **default**. Matriks dipakai apa adanya (baris = pemberi pengaruh). Label: *"Baris memengaruhi kolom (konvensi standar DEMATEL, dipakai apa adanya)"*.
   - `column_influences_row` → matriks tiap responden **di-transpose** sebelum dihitung. Label: *"Kolom memengaruhi baris (sesuai teks instruksi kuesioner)"*.
2. Tampilkan kotak info di atas hasil yang menjelaskan arah yang sedang dipakai.
3. Jangan pilih arah secara diam-diam di tempat lain di kode. Semua perhitungan setelah langkah 1 selalu memakai konvensi z_ij = pengaruh i → j.

---

## 5. Metode perhitungan (ikuti PERSIS urutan ini)

Notasi: n = jumlah faktor, K = jumlah responden, TFN = (l, m, u).

### Langkah 1 — Baca & orientasikan
Untuk tiap responden k, ambil matriks skor A^k (n×n). Jika orientasi = `column_influences_row`, gunakan transpose-nya. Paksa diagonal = 0.

### Langkah 2 — Konversi ke Triangular Fuzzy Number (Tabel 2 thesis)
| Skor | Variabel linguistik | TFN (l, m, u) |
|---|---|---|
| 0 | No influence | (0, 0.1, 0.3) |
| 1 | Very low influence | (0.1, 0.3, 0.5) |
| 2 | Low influence | (0.3, 0.5, 0.7) |
| 3 | High influence | (0.5, 0.7, 0.9) |
| 4 | Very high influence | (0.7, 0.9, 1.0) |

**Sel diagonal selalu TFN (0, 0, 0)**, bukan (0, 0.1, 0.3). Simpan skala ini sebagai konstanta di satu tempat di kode.

### Langkah 3 — Agregasi (rata-rata aritmetika antar pakar)
z̃_ij = ( (1/K)·Σ l_ij^k , (1/K)·Σ m_ij^k , (1/K)·Σ u_ij^k )

### Langkah 4 — Normalisasi
s = max_i ( Σ_j u_ij )   ← maksimum jumlah baris dari komponen **u**
x̃_ij = z̃_ij / s  (bagi l, m, u masing-masing dengan s yang sama)

### Langkah 5 — Total relation matrix fuzzy
Hitung terpisah untuk tiap komponen:
T_l = X_l (I − X_l)⁻¹,  T_m = X_m (I − X_m)⁻¹,  T_u = X_u (I − X_u)⁻¹

Invers matriks: implementasikan **Gauss-Jordan dengan partial pivoting** di JavaScript murni (float64). Jika pivot < 1e-12, hentikan dan tampilkan error "matriks singular".

### Langkah 6 — Defuzzifikasi CFCS (Opricovic & Tzeng, 2003)
Pakai min dan max **global** dari seluruh matriks T:
- min_l = min_ij (T_l,ij), max_u = max_ij (T_u,ij), Δ = max_u − min_l
- xl = (l − min_l)/Δ, xm = (m − min_l)/Δ, xu = (u − min_l)/Δ
- xls = xm / (1 + xm − xl)
- xrs = xu / (1 + xu − xm)
- x = [ xls·(1 − xls) + xrs·xrs ] / (1 − xls + xrs)
- t_ij (crisp) = min_l + x·Δ

### Langkah 7 — Prominence & Relation
- R_i = Σ_j t_ij (jumlah baris; di beberapa buku disebut D)
- C_j = Σ_i t_ij (jumlah kolom)
- Prominence = R + C, Relation = R − C
- Klasifikasi: R − C > 0 → **Cause**, R − C < 0 → **Effect**, = 0 → **Netral** (pakai toleransi |R−C| < 1e-12).
- Ranking: urutkan berdasarkan R + C (descending) dan juga R − C (descending); tampilkan kedua ranking.

### Langkah 8 — Threshold & hubungan signifikan
Pilihan di Pengaturan:
- `mean_offdiag` (**default**): α = rata-rata t_ij untuk i ≠ j.
- `mean_all`: α = rata-rata seluruh n² elemen t_ij.
- `manual`: pengguna mengetik nilai α.

Hubungan signifikan = pasangan (i ≠ j) dengan t_ij > α. Tampilkan sebagai daftar "Ri → Rj (nilai)" diurutkan descending, plus matriks T dengan sel signifikan diberi warna.

---

## 6. Tampilan aplikasi

Satu halaman, bersih, responsif. Susunan:

1. **Header**: judul aplikasi + satu kalimat petunjuk.
2. **Area upload**: drag & drop + tombol "Pilih file". Setelah upload tampilkan ringkasan: nama file, jumlah responden terbaca, jumlah faktor, daftar peringatan (kuning) / error (merah).
3. **Panel Pengaturan**: arah matriks (§4), metode threshold (§5 langkah 8). Mengubah pengaturan langsung menghitung ulang tanpa upload ulang.
4. **Hasil** dalam tab:
   - `Ringkasan`: tabel R, C, R+C, R−C, klasifikasi, ranking, causal tier hipotesis; top 5 cause & top 5 effect; nilai s, min_l, max_u, α, jumlah hubungan signifikan.
   - `Impact Relation Map`: scatter R+C (x) vs R−C (y), label kode faktor di tiap titik, garis y = 0, garis putus-putus vertikal di rata-rata R+C, warna + bentuk marker berbeda untuk Cause/Effect (aman buta warna). Gambar dengan **SVG** atau `<canvas>` sendiri (tanpa library grafik). Tombol "Unduh PNG" resolusi tinggi (min. 2000 px lebar).
   - `Hubungan Signifikan`: daftar + matriks T berwarna.
   - `Matriks Antara`: dropdown untuk melihat matriks tiap responden (skor), agregasi fuzzy (l;m;u), ternormalisasi, T fuzzy (l, m, u), T crisp.
   - `Cek Manual`: pilih dua faktor (dropdown dari–ke), tampilkan jejak perhitungan satu sel dari skor tiap pakar → TFN → agregasi → normalisasi → T fuzzy → CFCS langkah demi langkah → crisp. Ini untuk verifikasi manual di Excel dan untuk lampiran skripsi.
5. **Tombol "Unduh hasil Excel"** → `hasil_fuzzy_dematel.xlsx` dengan sheet:
   `Ringkasan_RC`, `Total_Relation_Crisp`, `Significant_Relations`, `T_fuzzy_l`, `T_fuzzy_m`, `T_fuzzy_u`, `Normalized_l`, `Normalized_m`, `Normalized_u`, `Aggregated_l`, `Aggregated_m`, `Aggregated_u`, `Parameters` (orientasi, metode threshold, α, s, min_l, max_u, Δ, jumlah pakar, nama sheet yang dibaca, daftar peringatan), dan `Skala_TFN`.
   Semua angka disimpan sebagai **angka** (bukan teks), tampilkan 4 desimal.

Tampilan angka di layar: 4 desimal. Perhitungan internal: presisi penuh, jangan dibulatkan di tengah jalan.

---

## 7. Teknis

- Satu file `fuzzy_dematel.html`: HTML + CSS + JS di dalam file yang sama.
- Baca & tulis Excel memakai **SheetJS** (`xlsx.full.min.js`). Agar aplikasi tetap jalan **tanpa internet**, unduh file library-nya dan **tanam (inline)** isinya ke dalam `<script>` di HTML. Kalau inline tidak memungkinkan, pakai CDN `https://cdn.sheetjs.com/` dengan versi dipin, dan beri tahu pengguna bahwa butuh internet.
- Tidak ada library lain. Tidak ada server, tidak ada `fetch` ke luar, tidak ada localStorage wajib.
- Pisahkan kode menjadi dua blok:
  - `<script id="engine">` → fungsi murni tanpa DOM: `parseWorkbook(workbook)`, `toTFN`, `aggregate`, `normalize`, `invert`, `totalRelation`, `defuzzifyCFCS`, `computeRC`, `computeThreshold`, `runFuzzyDematel(experts, options)`. Tambahkan komentar singkat yang menyebut nomor langkah §5.
  - `<script id="ui">` → semua kode tampilan.
- Tangani error dengan `try/catch`, tampilkan ke pengguna, jangan hanya `console.error`.

---

## 8. Pengujian (WAJIB sebelum menyatakan selesai)

Nilai acuan di bawah dihitung dengan implementasi Python independen (NumPy) dari data asli `Responden_final_thesis.xlsx`, memakai metode §5 dan threshold `mean_offdiag`.

1. Buat `tests/test_engine.mjs` (Node.js) yang:
   - membaca `fuzzy_dematel.html`, mengambil isi `<script id="engine">`, dan menjalankannya di Node (pakai paket npm `xlsx` untuk membaca workbook);
   - menjalankan `runFuzzyDematel` pada `Responden_final_thesis.xlsx` untuk **kedua** orientasi;
   - membandingkan dengan nilai acuan di bawah, toleransi **1e-4** (untuk R, C, R+C, R−C toleransi 5e-4 karena acuan dibulatkan 4 desimal);
   - memastikan peringatan diagonal `responden 7` sel R11 terdeteksi;
   - menguji validasi: file dengan sel kosong / nilai 5 / teks harus menghasilkan error yang menyebut sheet dan sel.
2. Jalankan tes, perbaiki sampai semua lulus. Laporkan hasilnya ke pengguna.
3. Tes tambahan kecil: matriks 3×3 buatan tangan dengan hasil yang dicek manual.

### Nilai acuan (reference)
Kunci `Z_agg_R1_to_R2` dan `T_fuzzy_R1_to_R2` adalah (l, m, u) untuk sel pengaruh R1 → R2 setelah orientasi diterapkan.

```json
{
 "row_influences_column": {
  "scale_factor_s": 15.985714,
  "cfcs_min_l": 0.015107,
  "cfcs_max_u": 0.719461,
  "threshold_mean_offdiag": 0.209336,
  "threshold_mean_all": 0.20756,
  "n_significant_with_offdiag_threshold": 196,
  "Z_agg_R1_to_R2": [
   0.557143,
   0.757143,
   0.914286
  ],
  "T_fuzzy_R1_to_R2": [
   0.060836,
   0.145453,
   0.699541
  ],
  "T_crisp_R1_to_R2": 0.23669,
  "factors": {
   "R1": {
    "R": 4.5049,
    "C": 4.239,
    "R_plus_C": 8.7439,
    "R_minus_C": 0.266
   },
   "R2": {
    "R": 4.2778,
    "C": 4.258,
    "R_plus_C": 8.5359,
    "R_minus_C": 0.0198
   },
   "R3": {
    "R": 4.1835,
    "C": 4.242,
    "R_plus_C": 8.4255,
    "R_minus_C": -0.0585
   },
   "R4": {
    "R": 4.2736,
    "C": 4.0479,
    "R_plus_C": 8.3215,
    "R_minus_C": 0.2257
   },
   "R5": {
    "R": 4.1509,
    "C": 4.175,
    "R_plus_C": 8.3259,
    "R_minus_C": -0.0241
   },
   "R6": {
    "R": 4.0553,
    "C": 4.1026,
    "R_plus_C": 8.1579,
    "R_minus_C": -0.0473
   },
   "R7": {
    "R": 4.2301,
    "C": 4.3804,
    "R_plus_C": 8.6106,
    "R_minus_C": -0.1503
   },
   "R8": {
    "R": 3.9446,
    "C": 3.8443,
    "R_plus_C": 7.7889,
    "R_minus_C": 0.1003
   },
   "R9": {
    "R": 4.1143,
    "C": 4.1222,
    "R_plus_C": 8.2364,
    "R_minus_C": -0.0079
   },
   "R10": {
    "R": 4.2428,
    "C": 4.1732,
    "R_plus_C": 8.4161,
    "R_minus_C": 0.0696
   },
   "R11": {
    "R": 4.1248,
    "C": 4.2331,
    "R_plus_C": 8.3579,
    "R_minus_C": -0.1083
   },
   "R12": {
    "R": 4.2381,
    "C": 4.1808,
    "R_plus_C": 8.4188,
    "R_minus_C": 0.0573
   },
   "R13": {
    "R": 4.1102,
    "C": 4.1466,
    "R_plus_C": 8.2568,
    "R_minus_C": -0.0364
   },
   "R14": {
    "R": 3.7386,
    "C": 3.7729,
    "R_plus_C": 7.5115,
    "R_minus_C": -0.0342
   },
   "R15": {
    "R": 4.049,
    "C": 4.0521,
    "R_plus_C": 8.1011,
    "R_minus_C": -0.0031
   },
   "R16": {
    "R": 4.2423,
    "C": 4.2891,
    "R_plus_C": 8.5313,
    "R_minus_C": -0.0468
   },
   "R17": {
    "R": 3.9781,
    "C": 4.0752,
    "R_plus_C": 8.0533,
    "R_minus_C": -0.097
   },
   "R18": {
    "R": 4.1478,
    "C": 4.1644,
    "R_plus_C": 8.3122,
    "R_minus_C": -0.0166
   },
   "R19": {
    "R": 4.2332,
    "C": 4.2663,
    "R_plus_C": 8.4995,
    "R_minus_C": -0.0331
   },
   "R20": {
    "R": 4.1842,
    "C": 4.2591,
    "R_plus_C": 8.4432,
    "R_minus_C": -0.0749
   }
  }
 },
 "column_influences_row": {
  "scale_factor_s": 15.628571,
  "cfcs_min_l": 0.016133,
  "cfcs_max_u": 1.033155,
  "threshold_mean_offdiag": 0.266823,
  "threshold_mean_all": 0.264972,
  "n_significant_with_offdiag_threshold": 199,
  "Z_agg_R1_to_R2": [
   0.557143,
   0.757143,
   0.9
  ],
  "T_fuzzy_R1_to_R2": [
   0.061048,
   0.150155,
   0.948403
  ],
  "T_crisp_R1_to_R2": 0.286556,
  "factors": {
   "R1": {
    "R": 5.4041,
    "C": 5.7405,
    "R_plus_C": 11.1446,
    "R_minus_C": -0.3364
   },
   "R2": {
    "R": 5.4316,
    "C": 5.4557,
    "R_plus_C": 10.8874,
    "R_minus_C": -0.0241
   },
   "R3": {
    "R": 5.4112,
    "C": 5.3372,
    "R_plus_C": 10.7483,
    "R_minus_C": 0.074
   },
   "R4": {
    "R": 5.1709,
    "C": 5.4516,
    "R_plus_C": 10.6225,
    "R_minus_C": -0.2807
   },
   "R5": {
    "R": 5.3285,
    "C": 5.3009,
    "R_plus_C": 10.6294,
    "R_minus_C": 0.0276
   },
   "R6": {
    "R": 5.2381,
    "C": 5.1796,
    "R_plus_C": 10.4177,
    "R_minus_C": 0.0585
   },
   "R7": {
    "R": 5.5872,
    "C": 5.3962,
    "R_plus_C": 10.9834,
    "R_minus_C": 0.191
   },
   "R8": {
    "R": 4.9155,
    "C": 5.041,
    "R_plus_C": 9.9564,
    "R_minus_C": -0.1255
   },
   "R9": {
    "R": 5.2608,
    "C": 5.2508,
    "R_plus_C": 10.5117,
    "R_minus_C": 0.01
   },
   "R10": {
    "R": 5.3271,
    "C": 5.4127,
    "R_plus_C": 10.7399,
    "R_minus_C": -0.0856
   },
   "R11": {
    "R": 5.4015,
    "C": 5.2646,
    "R_plus_C": 10.6661,
    "R_minus_C": 0.1368
   },
   "R12": {
    "R": 5.3394,
    "C": 5.4069,
    "R_plus_C": 10.7463,
    "R_minus_C": -0.0675
   },
   "R13": {
    "R": 5.2931,
    "C": 5.2472,
    "R_plus_C": 10.5403,
    "R_minus_C": 0.0458
   },
   "R14": {
    "R": 4.8287,
    "C": 4.7842,
    "R_plus_C": 9.6129,
    "R_minus_C": 0.0445
   },
   "R15": {
    "R": 5.1777,
    "C": 5.1751,
    "R_plus_C": 10.3528,
    "R_minus_C": 0.0026
   },
   "R16": {
    "R": 5.47,
    "C": 5.4162,
    "R_plus_C": 10.8863,
    "R_minus_C": 0.0538
   },
   "R17": {
    "R": 5.2067,
    "C": 5.0835,
    "R_plus_C": 10.2903,
    "R_minus_C": 0.1232
   },
   "R18": {
    "R": 5.3192,
    "C": 5.2969,
    "R_plus_C": 10.616,
    "R_minus_C": 0.0223
   },
   "R19": {
    "R": 5.4422,
    "C": 5.4045,
    "R_plus_C": 10.8467,
    "R_minus_C": 0.0377
   },
   "R20": {
    "R": 5.4353,
    "C": 5.3433,
    "R_plus_C": 10.7786,
    "R_minus_C": 0.092
   }
  }
 }
}
```

---

## 9. Cara kerja yang diharapkan dari Claude Code

1. Buka dan periksa `Responden_final_thesis.xlsx` terlebih dahulu; pastikan strukturnya sesuai §3. Jika berbeda, **berhenti dan tanyakan** ke pengguna.
2. Jangan mengubah metode di §5 (skala TFN, normalisasi, CFCS, threshold) tanpa bertanya. Jika menemukan hal yang menurutmu keliru secara metodologis, jelaskan ke pengguna dan minta keputusan.
3. Bangun `fuzzy_dematel.html`, lalu tes sesuai §8.
4. Buka hasilnya dan cek tampilan (tidak ada teks terpotong, IRM terbaca, label tidak saling menumpuk parah).
5. Di akhir, beri pengguna ringkasan singkat dalam Bahasa Indonesia: cara membuka aplikasi, hasil tes, dan pengingat untuk mengonfirmasi arah matriks (§4) sesuai bab metodologi thesis.

## 10. Struktur folder akhir
```
./
├── CLAUDE.md                     (file ini)
├── Responden_final_thesis.xlsx   (data asli, jangan diubah)
├── fuzzy_dematel.html            (aplikasi, deliverable utama)
└── tests/
    └── test_engine.mjs
```
