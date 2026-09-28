# Fuzzy DEMATEL – Risiko Rantai Pasok Geopolitik

Aplikasi satu file (`fuzzy_dematel.html`) untuk menghitung Fuzzy DEMATEL dari file Excel jawaban responden.
Spesifikasi lengkap: [CLAUDE.md](CLAUDE.md).

## Cara pakai
1. Unduh `fuzzy_dematel.html`, lalu buka dengan double-click (Chrome/Edge). Tidak perlu internet atau instalasi.
2. Tarik & lepas `Responden_final_thesis.xlsx` (atau klik **Pilih file**).
3. Pilih arah matriks dan metode threshold di panel **Pengaturan**.
4. Klik **Unduh hasil Excel** dan **Unduh PNG** (tab Impact Relation Map).

## Tes (untuk pengembang)
```bash
npm install
npm test
```
Tes membutuhkan `Responden_final_thesis.xlsx` di folder ini (file data tidak disertakan di repo).

Pustaka SheetJS Community Edition 0.18.5 (Apache-2.0) ditanam di dalam HTML.
