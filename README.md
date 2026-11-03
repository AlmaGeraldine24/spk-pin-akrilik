# Dashboard Keputusan Produksi Pin Akrilik

Website statis (tanpa server) untuk Praktikum Terintegrasi II. Unggah data, lalu grafik beranimasi dan status **Aman / Sedang / Tidak aman** muncul otomatis.

## Isi
| File | Fungsi |
|---|---|
| `index.html`, `style.css`, `app.js` | Seluruh website |
| `contoh/*.csv` | Contoh format data yang bisa diunggah |

## Tiga analisis (sesuai modul)
1. **Produksi & Mutu (Modul I–II)**: Aman bila produksi ≥ 200 pcs dan cacat ≤ 4%. Sedang bila produksi ≥ BEP dengan angsuran (159 pcs) tetapi < 200, atau cacat 4–5%. Tidak aman bila produksi < BEP dengan angsuran atau cacat > 5%.
2. **Persediaan (Modul III)**: Aman bila stok > ROP 101 pin. Sedang bila stok antara safety stock 51 dan 101 pin (segera pesan ulang). Tidak aman bila stok < 51 pin (darurat, SOP-06B).
3. **Alternatif perbaikan (Modul IV, SAW)**: normalisasi (C1, C3, C4 cost; C2 benefit), bobot 30/30/20/20, nilai V dan peringkat. Aman bila V ≥ 0,75, Sedang bila V ≥ 0,60.

Keputusan keseluruhan = status terendah dari modul yang sudah diunggah.

> Batas cacat 4% dan skor SAW 0,75 / 0,60 adalah asumsi kelompok (tidak ada di modul). Semuanya bisa diubah di panel **Parameter**.

## Cara pakai di GitHub Pages
1. Buat repository baru, unggah semua file ini (termasuk folder `contoh`).
2. Buka **Settings → Pages**, pilih Branch `main` dan folder `/ (root)`, lalu Save.
3. Tunggu satu menit, buka `https://USERNAME.github.io/NAMA-REPO/`.

Situs butuh internet saat dibuka karena memuat Chart.js dan SheetJS dari CDN.
