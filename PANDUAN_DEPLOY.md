# Panduan Deploy ke Streamlit Community Cloud

## Isi proyek
| Berkas | Fungsi |
|---|---|
| `app.py` | Tampilan HP (menu bawah, kartu, kamera penuh), semua halaman |
| `omr_engine.py` | Mesin OMR OpenCV: 4 kotak sudut, QR, hitung arsiran vs kunci |
| `database_manager.py` | Database SQLite, impor/ekspor Excel, cadangan/pulihkan |
| `ljk_generator.py` | Membuat PDF LJK (1-100 soal) |
| `requirements.txt` | Pustaka (memakai `opencv-python-headless`) |
| `.streamlit/config.toml` | Tema & pengaturan server |
| `.gitignore` | Mencegah database & foto siswa ikut terunggah |

## Langkah deploy
1. Buat repositori GitHub baru (disarankan **Private**), unggah semua berkas di atas
   (termasuk folder `.streamlit`). Jangan unggah `ljk_data.db` atau folder `scans/`.
2. Buka https://share.streamlit.io, login dengan GitHub, klik **Create app**.
3. Pilih repositori, branch `main`, Main file path: `app.py`, lalu **Deploy**.
4. **Atur kata sandi** (sangat disarankan; tanpa ini siapa pun yang tahu link bisa membuka data siswa):
   Settings > Secrets, isi
   ```toml
   APP_PASSWORD = "kata-sandi-anda"
   ```
5. Buka link aplikasi di HP.

## Pasang di layar utama HP
- Android (Chrome): menu ⋮ > **Tambahkan ke layar utama** / **Instal aplikasi**.
- iPhone (Safari): tombol Bagikan > **Tambah ke Layar Utama**.

Ini mode "mirip aplikasi" (layar penuh, ikon sendiri). Mode offline penuh tidak bisa
di Streamlit Community Cloud, jadi koneksi internet tetap diperlukan.

## PENTING: data tidak permanen di Cloud
Disk Streamlit Cloud sementara. Saat aplikasi restart, di-deploy ulang, atau tertidur karena
lama tidak dipakai, database dan foto bisa HILANG. Biasakan:
- Setelah sesi scan: menu 🏠 Kuis > **Cadangan & pulihkan data** > **Unduh cadangan**.
- Sebelum mulai lagi: **Pulihkan data** dengan file cadangan itu.
- Ekspor nilai ke Excel dari menu 📊 Rekap setelah selesai.
Untuk data permanen otomatis, perlu database eksternal (di luar cakupan paket ini).

## Tips scan
- Kamera di browser (`st.camera_input`) beresolusi rendah dan biasanya kamera depan.
  Untuk akurasi terbaik pakai **📁 Unggah foto**: di HP tombolnya membuka kamera bawaan
  dengan resolusi penuh. Foto dengan sisi pendek di bawah ±720 px sering gagal terbaca.
- Cahaya merata, kertas rata, keempat kotak hitam sudut terlihat utuh.
- Sensitivitas arsiran bisa diubah di ⚙️ Opsi scan.

## Jalankan di komputer sendiri
Windows: klik `run_app.bat`. Mac/Linux: `./run_app.sh`. Atau: `pip install -r requirements.txt`
lalu `streamlit run app.py`.
